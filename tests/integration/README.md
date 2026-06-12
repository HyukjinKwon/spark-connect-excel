<!-- SPDX-License-Identifier: Apache-2.0 -->

# Integration e2e - Python runtime vs a real Spark Connect server

`test_runtime_e2e.py` drives the real result-shaping functions in
`python/spark_excel_runtime.py` (`run_sql`, `schema_of`) against a **real**
in-process Spark Connect server, asserting the `SparkResult` JSON is correct on
genuine Spark output (column types, rows, truncation, nulls, error envelope).
This complements the fake-based unit tests in `python/tests/`, which never touch
Spark.

No browser, no Envoy, no `pyspark-connect-web`: the test sets the session on the
module directly and bypasses `connect()` (which installs the browser grpc-web
patch). It starts Spark via `SparkSession.builder.remote("local[*]")`.

## Run it

```bash
pip install "pyspark[connect]==4.0.0" pandas pyarrow pytest   # needs a JVM (Java 17)
npm run test:py:integration
# or: python -m pytest tests/integration -q
```

Without pyspark (or a JVM) it **skips** cleanly, so it is safe in any
environment. CI runs it for real in the `integration` job (`.github/workflows/ci.yml`).

This is the automated counterpart to the browser/Excel matrix in
`tests/e2e/` (which needs a real Excel host and stays deferred).
