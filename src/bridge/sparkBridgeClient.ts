// SPDX-License-Identifier: Apache-2.0
//
// sparkBridgeClient.ts — task-pane-side implementation of SparkBridge.
//
// Forwards each SparkBridge method call across the Office dialog message channel
// as a BridgeRequest envelope and resolves on the matching BridgeResponse.
// BridgeEvents (progress, status) are handled asynchronously and update local
// state / invoke an optional onEvent callback.
//
// Transport is injected so SparkBridgeClient itself never imports Office.js and
// can be tested without it.  createDialogBridge() wraps the Office dialog API
// into the transport shape.

import type {
  SparkBridge,
  ConnectOptions,
  SparkResult,
  ColumnMeta,
  RuntimeStatus,
  BridgeMethod,
  BridgeRequest,
  BridgeResponse,
  BridgeEvent,
} from "../seam";
import { encodeMessage, decodeMessage } from "../seam";

// ---------------------------------------------------------------------------
// Transport abstraction
// ---------------------------------------------------------------------------

export interface BridgeTransport {
  /** Send a string message to the dialog. */
  send(msg: string): void;
  /**
   * Register a listener for messages arriving from the dialog.
   * Returns an unsubscribe function.
   */
  subscribe(listener: (msg: string) => void): () => void;
}

// ---------------------------------------------------------------------------
// SparkBridgeClient
// ---------------------------------------------------------------------------

export class SparkBridgeClient implements SparkBridge {
  private readonly _transport: BridgeTransport;
  private _nextId = 1;
  private _pending = new Map<
    number,
    { resolve: (v: unknown) => void; reject: (e: Error) => void }
  >();
  // Last-known status (updated by BridgeEvent "status" pushes).
  private _status: RuntimeStatus = {
    crossOriginIsolated: false,
    pyodideReady: false,
    connected: false,
  };
  private readonly _onEvent: ((evt: BridgeEvent) => void) | undefined;

  constructor(transport: BridgeTransport, opts?: { onEvent?: (evt: BridgeEvent) => void }) {
    this._transport = transport;
    this._onEvent = opts?.onEvent;
    transport.subscribe((raw) => this._handleMessage(raw));
  }

  // -------------------------------------------------------------------------
  // SparkBridge interface
  // -------------------------------------------------------------------------

  ensureReady(): Promise<void> {
    return this._call("ensureReady") as Promise<void>;
  }

  connect(uri: string, opts?: ConnectOptions): Promise<void> {
    return this._call("connect", uri, opts) as Promise<void>;
  }

  async runSQL(sql: string, rowCap: number): Promise<SparkResult> {
    return this._call("runSQL", sql, rowCap) as Promise<SparkResult>;
  }

  async schemaOf(sql: string): Promise<ColumnMeta[]> {
    return this._call("schemaOf", sql) as Promise<ColumnMeta[]>;
  }

  /** Returns the last-known status pushed by the dialog. Synchronous. */
  status(): RuntimeStatus {
    return { ...this._status };
  }

  cancel(): void {
    // Fire-and-forget: send the cancel request but do not wait for a response.
    const id = this._nextId++;
    const req: BridgeRequest = {
      kind: "req",
      id,
      method: "cancel",
      args: [],
    };
    try {
      this._transport.send(encodeMessage(req));
    } catch {
      /* best-effort */
    }
    // cancel() is void on the SparkBridge interface — no pending promise needed.
  }

  // -------------------------------------------------------------------------
  // Internal messaging
  // -------------------------------------------------------------------------

  private _call(method: BridgeMethod, ...args: unknown[]): Promise<unknown> {
    return new Promise<unknown>((resolve, reject) => {
      const id = this._nextId++;
      this._pending.set(id, { resolve, reject });
      const req: BridgeRequest = { kind: "req", id, method, args };
      try {
        this._transport.send(encodeMessage(req));
      } catch (err) {
        this._pending.delete(id);
        reject(err instanceof Error ? err : new Error(String(err)));
      }
    });
  }

  private _handleMessage(raw: string): void {
    let msg;
    try {
      msg = decodeMessage(raw);
    } catch {
      // Malformed message — ignore.
      return;
    }

    if (msg.kind === "res") {
      const res = msg as BridgeResponse;
      const pending = this._pending.get(res.id);
      if (!pending) return;
      this._pending.delete(res.id);
      if (res.ok) {
        pending.resolve(res.result);
      } else {
        const e = res.error;
        pending.reject(new Error(e ? `${e.name}: ${e.message}` : "Unknown bridge error"));
      }
      return;
    }

    if (msg.kind === "evt") {
      const evt = msg as BridgeEvent;
      if (evt.event === "status" && evt.payload) {
        // Merge the status update into our cached status.
        const update = evt.payload as Partial<RuntimeStatus>;
        this._status = { ...this._status, ...update };
      }
      this._onEvent?.(evt);
      return;
    }
    // "req" arriving on the client side would be unexpected — ignore.
  }

  /**
   * Reject all pending calls (e.g. when the dialog closes unexpectedly).
   */
  rejectAll(reason: string): void {
    for (const [id, pending] of this._pending) {
      pending.reject(new Error(reason));
      this._pending.delete(id);
    }
  }
}

// ---------------------------------------------------------------------------
// createDialogBridge — Office.js-backed factory
// ---------------------------------------------------------------------------

/**
 * Options for createDialogBridge.
 */
export interface DialogBridgeOptions {
  /** Width of the Office dialog as a percentage of screen width (default 60). */
  width?: number;
  /** Height of the Office dialog as a percentage of screen height (default 60). */
  height?: number;
  /** Callback for unsolicited BridgeEvents (progress, log, status). */
  onEvent?: (evt: BridgeEvent) => void;
}

/**
 * Open the COI host dialog and wire it to a SparkBridgeClient.
 *
 * Calls `Office.context.ui.displayDialogAsync`, then subscribes to
 * DialogMessageReceived and DialogEventReceived (dialog closed).
 *
 * @param dialogUrl  URL of the dialog page (must be hosted with COOP/COEP).
 * @param opts       Optional width/height/onEvent.
 */
export async function createDialogBridge(
  dialogUrl: string,
  opts: DialogBridgeOptions = {},
): Promise<SparkBridgeClient> {
  const { width = 60, height = 60, onEvent } = opts;

  return new Promise<SparkBridgeClient>((resolve, reject) => {
    Office.context.ui.displayDialogAsync(
      dialogUrl,
      { width, height, promptBeforeOpen: false },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          reject(new Error(`displayDialogAsync failed: ${result.error?.message ?? "unknown"}`));
          return;
        }

        const dialog = result.value;
        // Listeners registered by the client.
        const listeners: Array<(msg: string) => void> = [];

        const transport: BridgeTransport = {
          send(msg: string) {
            dialog.messageChild(msg);
          },
          subscribe(listener: (msg: string) => void) {
            listeners.push(listener);
            return () => {
              const idx = listeners.indexOf(listener);
              if (idx !== -1) listeners.splice(idx, 1);
            };
          },
        };

        const client = new SparkBridgeClient(transport, { onEvent });

        // Forward inbound dialog messages to all subscribed listeners.
        dialog.addEventHandler(
          Office.EventType.DialogMessageReceived,
          (args: { message: string; origin: string | undefined } | { error: number }) => {
            if ("message" in args) {
              for (const l of listeners) {
                l(args.message);
              }
            }
          },
        );

        // Handle dialog closed / navigated away.
        dialog.addEventHandler(
          Office.EventType.DialogEventReceived,
          (args: { message: string; origin: string | undefined } | { error: number }) => {
            if ("error" in args) {
              // 12006 = dialog closed by user, 12005 = navigation rejected, etc.
              const reason = `Dialog closed (code ${args.error})`;
              client.rejectAll(reason);
            }
          },
        );

        resolve(client);
      },
    );
  });
}
