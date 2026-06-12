<!-- SPDX-License-Identifier: Apache-2.0 -->

# Lane D findings

## 1. Bearer-token / auth mechanism

**Mechanism chosen:** Append `token=<value>` as a URI parameter to the
`sc://...;transport=grpcweb` connection string.

```python
effective_uri = uri.rstrip(";") + ";" + f"token={token}"
# e.g.  sc://host:8081/;transport=grpcweb;token=eyJhb...
```

**Why this works:** PySpark's `DefaultChannelBuilder` reads the `token=` URI
param and maps it to the `authorization: Bearer <token>` channel-level metadata
entry.  The pyspark-connect-web patch forwards that metadata as a grpc-web HTTP
header on every request to the Envoy proxy (see `patch.py`, `patched_to_channel`,
`meta = dict(self.metadata())`).  Envoy gate-checks the header.

**Evidence:** The pyspark-connect-web integration tests use exactly this form
(`sc://localhost:{port}/;token={token}`) for the native Connect client
(`test_real_round_trip.py` line 86).  The web transport carries the same token
via the `WebChannel.params` map populated from `self.metadata()`.

**Assumption documented:** The `token=` URI param is not double-appended if the
caller already includes it (`if "token=" not in uri` guard).  If pyspark-connect-web
exposes a dedicated builder option in a future version, prefer that over URI embedding.

## 2. Python module loading into Pyodide

The Python source is bundled at build time as a raw string via Vite's `?raw` import:

```ts
import runtimeSource from "../../python/spark_excel_runtime.py?raw";
```

At runtime inside `SparkBridgeHost.ensureReady()`:

1. `RuntimeHost.boot()` boots Pyodide + installs the pcw wheel (Lane C).
2. A `runPython` snippet calls `pcw.install()` (idempotent).
3. A loader snippet base64-encodes the source in TS, decodes in Python,
   `compile()`/`exec()`s it into a `types.ModuleType("spark_excel_runtime")`,
   registers it in `sys.modules`, then imports `connect`, `run_sql`, `schema_of`
   into the Pyodide globals for subsequent `runPython` call snippets.

Idempotency guard: `if "spark_excel_runtime" not in sys.modules` ensures the
module body runs only once.

## 3. Files created

| File | Purpose |
|------|---------|
| `python/spark_excel_runtime.py` | `connect`, `run_sql`, `schema_of` + `_normalize_value`/`_result_to_dict` |
| `python/tests/__init__.py` | Empty package marker |
| `python/tests/test_spark_excel_runtime.py` | 33 pytest tests (no live Spark) |
| `src/bridge/marshal.ts` | `parseResult`, `parseSchema`, `parseConnectResult` |
| `src/bridge/sparkBridgeHost.ts` | `SparkBridgeHost implements SparkBridge` |
| `src/bridge/sparkBridgeClient.ts` | `SparkBridgeClient` + `createDialogBridge` |
| `docs/findings-lane-D.md` | This file |

## 4. Test results

```
python -m pytest python/tests -q
33 passed in 8.77s
```

## 5. Seam assumptions

- `RuntimeHost.runPython(src)` resolves with the string value of the **last expression** in the snippet (Pyodide convention).  Every call snippet ends with the module function call whose return value is the JSON string.
- `RuntimeHost.boot()` installs pyspark-connect-web via micropip.  We call `pcw.install()` separately (also idempotent) to ensure the monkey-patch is active before the runtime module is loaded.
- `self.crossOriginIsolated` is available in the dialog window (top-level window served with COOP/COEP, DECISIONS #1/#2).
- `cancel()` is best-effort: pyspark-connect-web does not expose a Python-level interrupt; a full cancel requires the SAB/Atomics path owned by Lane C.
