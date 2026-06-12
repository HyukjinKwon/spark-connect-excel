// SPDX-License-Identifier: Apache-2.0
//
// runPython.ts — request/response RPC machinery for the pcw_run / pcw_result /
// pcw_run_error worker protocol (mirrors `installRunPython` in
// /tmp/pcw/pyspark_connect_web/jupyterlite/run_python_bridge.js).
//
// Lane C owns this file. Do not import it from src/taskpane/** (DECISIONS #7).

"use strict";

export interface PendingCall {
  resolve: (value: string) => void;
  reject: (reason: Error) => void;
}

/**
 * RpcHandle encapsulates the pending-promise map and the message listener
 * wired to a Worker. Call `send(code)` to dispatch a pcw_run request and
 * get a Promise<string> back. Call `rejectAll(reason)` to drain pending
 * calls on worker error or termination.
 */
export class RpcHandle {
  private _seq = 0;
  private _pending = new Map<number, PendingCall>();
  private _worker: Worker;
  private _onMessage: (ev: MessageEvent) => void;

  constructor(worker: Worker) {
    this._worker = worker;
    this._onMessage = (ev: MessageEvent) => {
      const msg = (ev.data ?? {}) as Record<string, unknown>;
      if (msg.type === "pcw_result" && typeof msg.id === "number") {
        const entry = this._pending.get(msg.id);
        if (entry) {
          this._pending.delete(msg.id);
          entry.resolve(String(msg.result ?? ""));
        }
      } else if (msg.type === "pcw_run_error" && typeof msg.id === "number") {
        const entry = this._pending.get(msg.id);
        if (entry) {
          this._pending.delete(msg.id);
          entry.reject(new Error(String(msg.message ?? "Unknown Python error")));
        }
      }
    };
    worker.addEventListener("message", this._onMessage);
  }

  /**
   * Dispatch `code` to the worker via pcw_run and resolve/reject when the
   * worker posts pcw_result / pcw_run_error with the matching id.
   */
  send(code: string): Promise<string> {
    const id = ++this._seq;
    return new Promise<string>((resolve, reject) => {
      this._pending.set(id, { resolve, reject });
      this._worker.postMessage({ type: "pcw_run", id, code });
    });
  }

  /**
   * Reject every in-flight call with `reason` and clear the map. Called when
   * the worker errors or is terminated so callers don't hang forever.
   */
  rejectAll(reason: Error): void {
    for (const entry of this._pending.values()) {
      entry.reject(reason);
    }
    this._pending.clear();
  }

  /** Detach the message listener from the worker. */
  detach(): void {
    this._worker.removeEventListener("message", this._onMessage);
  }
}

/**
 * Factory helper: create an RpcHandle wired to `worker`.
 * Optional convenience so callers don't import the class name.
 */
export function createRpc(worker: Worker): RpcHandle {
  return new RpcHandle(worker);
}
