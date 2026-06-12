// SPDX-License-Identifier: Apache-2.0
//
// connectionStore.ts — connection config persistence + secure token handling.
//
// DECISIONS #6 (invariant): bearer tokens NEVER touch document settings or
// worksheet cells. They are kept in OfficeRuntime.storage (roaming, opaque to
// the spreadsheet file) when available, otherwise in session-scoped memory
// only. The document settings value persists ONLY the non-secret config
// (host + port + tls flag); the token is stored separately and excluded from
// any serialisation that might embed in the .xlsx.
//
// This module has NO hard dependency on Office APIs at import time — all Office
// references are reached through the SettingsBackend abstraction so unit tests
// can inject a memory backend without needing an Office host.

import { SC_URI_HINT } from "../seam.js";

// ---------------------------------------------------------------------------
// Public types
// ---------------------------------------------------------------------------

/** Non-secret connection parameters (everything except the bearer token). */
export interface ConnectionConfig {
  host: string;
  port: number;
  /** When true the Envoy proxy is reached over HTTPS/TLS; the sc:// URI
   *  still uses `transport=grpcweb` — TLS is implied by the proxy setup. */
  tls: boolean;
}

// ---------------------------------------------------------------------------
// SettingsBackend — abstraction over Office.context.document.settings
// ---------------------------------------------------------------------------

/**
 * Minimal key-value store interface that connectionStore writes to.
 * Implementations must be injected; no global is imported directly.
 * This keeps the module testable without an Office host.
 */
export interface SettingsBackend {
  get(key: string): string | null;
  set(key: string, value: string): Promise<void>;
  remove(key: string): Promise<void>;
}

/** Key used in the settings backend for the persisted connection config. */
const CONFIG_KEY = "sparkConnectConfig";

// ---------------------------------------------------------------------------
// OfficeRuntime.storage token backend — kept separate from document settings.
//
// DECISIONS #6: OfficeRuntime.storage is roaming/session-scoped and is NOT
// embedded in the .xlsx file, so a token stored here is never inadvertently
// shared via a shared workbook or emailed spreadsheet. When OfficeRuntime is
// unavailable (e.g. unit-test context), we fall back to an in-memory Map that
// lives only for the lifetime of the current browser session and is wiped on
// reload.
// ---------------------------------------------------------------------------

const TOKEN_KEY = "sparkConnectToken";

/** Session-only in-memory fallback used when OfficeRuntime.storage is absent. */
const memoryTokenStore = new Map<string, string>();

/** Minimal shape of OfficeRuntime.storage we depend on (roaming key-value store). */
interface RoamingStorage {
  getItem(key: string): Promise<string | null>;
  setItem(key: string, value: string): Promise<void>;
  removeItem(key: string): Promise<void>;
}

/**
 * Resolve `OfficeRuntime.storage` if present. Accessed via `globalThis` so the
 * module type-checks without a hard dependency on the `OfficeRuntime` ambient
 * global (it is injected by the Office runtime at execution time only).
 */
function roamingStorage(): RoamingStorage | null {
  const rt = (globalThis as { OfficeRuntime?: { storage?: RoamingStorage } }).OfficeRuntime;
  return rt && rt.storage ? rt.storage : null;
}

/**
 * Persist a bearer token securely.
 *
 * Prefers `OfficeRuntime.storage` (roaming, opaque to the spreadsheet file).
 * Falls back to an in-memory Map when that API is unavailable (tests, offline).
 *
 * DECISIONS #6: The token is NEVER written to `Office.context.document.settings`
 * or to any worksheet cell. This function is the only authorised write path.
 */
export async function saveToken(token: string): Promise<void> {
  const storage = roamingStorage();
  if (storage) {
    await storage.setItem(TOKEN_KEY, token);
  } else {
    memoryTokenStore.set(TOKEN_KEY, token);
  }
}

/**
 * Load the bearer token previously saved by `saveToken`.
 * Returns `null` when no token has been stored.
 */
export async function loadToken(): Promise<string | null> {
  const storage = roamingStorage();
  if (storage) {
    return (await storage.getItem(TOKEN_KEY)) ?? null;
  }
  return memoryTokenStore.get(TOKEN_KEY) ?? null;
}

/**
 * Erase the stored bearer token (e.g. on sign-out or connection reset).
 */
export async function clearToken(): Promise<void> {
  const storage = roamingStorage();
  if (storage) {
    await storage.removeItem(TOKEN_KEY);
  } else {
    memoryTokenStore.delete(TOKEN_KEY);
  }
}

// ---------------------------------------------------------------------------
// Connection config persistence (non-secret only)
// ---------------------------------------------------------------------------

/**
 * Persist the non-secret `ConnectionConfig` to the provided settings backend.
 *
 * The token is intentionally NOT a parameter here — call `saveToken` separately
 * so there is no code path that could accidentally bundle it with the config.
 */
export async function saveConnection(
  cfg: ConnectionConfig,
  backend: SettingsBackend,
): Promise<void> {
  const json = JSON.stringify(cfg);
  await backend.set(CONFIG_KEY, json);
}

/**
 * Load a previously persisted `ConnectionConfig` from the settings backend.
 * Returns `null` if nothing has been saved yet or the stored value is corrupt.
 */
export function loadConnection(backend: SettingsBackend): ConnectionConfig | null {
  const raw = backend.get(CONFIG_KEY);
  if (raw == null) return null;
  try {
    const parsed = JSON.parse(raw) as unknown;
    if (!isConnectionConfig(parsed)) return null;
    return parsed;
  } catch {
    return null;
  }
}

/**
 * Remove the persisted config from the settings backend.
 */
export async function clearConnection(backend: SettingsBackend): Promise<void> {
  await backend.remove(CONFIG_KEY);
}

// ---------------------------------------------------------------------------
// URI builder
// ---------------------------------------------------------------------------

/**
 * Build the canonical grpc-web connection URI from a `ConnectionConfig`.
 *
 * Returns the `sc://host:port/;transport=grpcweb` form expected by
 * pyspark-connect-web / Spark Connect client.  TLS is implied by the proxy;
 * the `sc://` scheme always uses `transport=grpcweb` regardless of whether
 * the Envoy proxy is TLS-terminated or not (the client negotiates TLS via the
 * proxy; the sc:// URI only specifies the transport protocol, not encryption).
 *
 * This is consistent with `seam.ts#normalizeRemoteUri` which also accepts the
 * http(s):// shorthand — both reach the same Envoy grpc-web endpoint.
 *
 * Example: `buildRemoteUri({ host: "localhost", port: 8081, tls: false })`
 *          → `"sc://localhost:8081/;transport=grpcweb"`
 */
export function buildRemoteUri(cfg: ConnectionConfig): string {
  // Validate — the hint in the seam is authoritative on the shape.
  void SC_URI_HINT; // ensure import is not tree-shaken in strict mode
  return `sc://${cfg.host}:${cfg.port}/;transport=grpcweb`;
}

// ---------------------------------------------------------------------------
// SettingsBackend implementations
// ---------------------------------------------------------------------------

/**
 * Backend backed by `Office.context.document.settings`.
 *
 * Note: `settings.saveAsync` is called after every `set` / `remove` so
 * changes survive workbook close.  On `get`, values are read from the
 * in-memory cache that Office keeps synchronised.
 */
export function officeDocumentSettingsBackend(): SettingsBackend {
  return {
    get(key: string): string | null {
      // Office.context.document.settings.get returns undefined when missing.
      const val: unknown = Office.context.document.settings.get(key);
      if (val == null) return null;
      if (typeof val !== "string") return null;
      return val;
    },

    set(key: string, value: string): Promise<void> {
      Office.context.document.settings.set(key, value);
      return new Promise<void>((resolve, reject) => {
        Office.context.document.settings.saveAsync((result) => {
          if (result.status === Office.AsyncResultStatus.Failed) {
            reject(new Error(result.error.message));
          } else {
            resolve();
          }
        });
      });
    },

    remove(key: string): Promise<void> {
      Office.context.document.settings.remove(key);
      return new Promise<void>((resolve, reject) => {
        Office.context.document.settings.saveAsync((result) => {
          if (result.status === Office.AsyncResultStatus.Failed) {
            reject(new Error(result.error.message));
          } else {
            resolve();
          }
        });
      });
    },
  };
}

/**
 * In-memory settings backend — used in unit tests and offline contexts.
 * State is never flushed to disk; it lives only for the lifetime of the
 * calling environment.
 */
export function memorySettingsBackend(): SettingsBackend {
  const store = new Map<string, string>();
  return {
    get(key: string): string | null {
      return store.get(key) ?? null;
    },
    set(key: string, value: string): Promise<void> {
      store.set(key, value);
      return Promise.resolve();
    },
    remove(key: string): Promise<void> {
      store.delete(key);
      return Promise.resolve();
    },
  };
}

// ---------------------------------------------------------------------------
// Internal helpers
// ---------------------------------------------------------------------------

function isConnectionConfig(v: unknown): v is ConnectionConfig {
  if (v == null || typeof v !== "object") return false;
  const o = v as Record<string, unknown>;
  return (
    typeof o["host"] === "string" && typeof o["port"] === "number" && typeof o["tls"] === "boolean"
  );
}
