# SPDX-License-Identifier: Apache-2.0
"""spark_excel_runtime.py — Python module loaded into Pyodide inside the COI dialog.

Every public function returns a JSON string (DECISIONS #8 / API_CONTRACT §3):

  connect(uri, token)   -> '{"ok": true}' | '{"ok": false, "error": {...}}'
  run_sql(sql, row_cap) -> SparkResult JSON (success) | '{"ok": false, ...}' (failure)
  schema_of(sql)        -> '{"schema": [...]}' (success) | '{"ok": false, ...}' (failure)

Auth mechanism
--------------
Bearer tokens are embedded in the Spark Connect sc:// URI as the ``token=``
parameter:  sc://host:port/;transport=grpcweb;token=<bearer-value>

PySpark's DefaultChannelBuilder maps ``token=<value>`` → ``authorization: Bearer
<value>`` in the channel metadata, which the pyspark-connect-web patch forwards
as a grpc-web HTTP header to the Envoy proxy.  This is the documented mechanism;
see pyspark-connect-web docs/connection-patterns.md and the integration tests
(test_real_round_trip.py line 86).

Success vs. failure shapes
---------------------------
``run_sql`` and ``schema_of`` return either the plain SparkResult / schema object
(on success) or ``{"ok": false, "error": {"name": ..., "message": ...}}`` on
failure.  The TypeScript side detects failure by checking ``"ok" in parsed &&
parsed.ok === false``.
"""
from __future__ import annotations

import json
import math
import traceback
from typing import Any

# ---------------------------------------------------------------------------
# Module-global session holder (one session per runtime lifetime).
# ---------------------------------------------------------------------------
_spark: Any = None  # pyspark.sql.SparkSession, set by connect()


# ---------------------------------------------------------------------------
# Value normalisation helpers (unit-testable without Spark)
# ---------------------------------------------------------------------------

def _normalize_value(v: Any) -> Any:
    """Convert a single pandas cell value to a JSON-serialisable Python object.

    Rules (applied in order):
    - None / pd.NA / pd.NaT → None
    - numpy bool_ → Python bool
    - numpy integer scalars → Python int
    - numpy floating scalars → Python float (NaN/Inf → None)
    - pandas Timestamp / datetime.datetime → ISO-8601 string
    - datetime.date → ISO-8601 string (YYYY-MM-DD)
    - Everything else: str() as last resort if not already a JSON primitive.
    """
    # Guard: import pandas lazily so the module can be imported in environments
    # that have pyspark stubs but no pandas (tests override toPandas anyway).
    import datetime
    import decimal

    try:
        import pandas as pd
        import numpy as np
    except ImportError:  # pragma: no cover — Pyodide always has both
        pd = None  # type: ignore[assignment]
        np = None  # type: ignore[assignment]

    # None / pandas NA sentinels
    if v is None:
        return None
    if pd is not None:
        if v is pd.NA:
            return None
        if v is pd.NaT:
            return None
        # pandas Timestamp
        if isinstance(v, pd.Timestamp):
            if v is pd.NaT:
                return None
            return v.isoformat()

    if np is not None:
        # numpy bool must come before numpy integer (bool_ is a subclass of int)
        if isinstance(v, np.bool_):
            return bool(v)
        if isinstance(v, np.integer):
            return int(v)
        if isinstance(v, np.floating):
            if math.isnan(v) or math.isinf(v):
                return None
            return float(v)

    # Python float NaN / Inf
    if isinstance(v, float):
        if math.isnan(v) or math.isinf(v):
            return None
        return v

    # datetime.datetime before datetime.date (datetime is a subclass of date)
    if isinstance(v, datetime.datetime):
        return v.isoformat()
    if isinstance(v, datetime.date):
        return v.isoformat()

    # Python bool (after numpy so we don't double-handle)
    if isinstance(v, bool):
        return v

    # decimal.Decimal (Spark DECIMAL columns + SQL numeric literals like 1.5) ->
    # float, so Excel receives a number, not text. Excel/JS use float64 anyway;
    # very-high-precision decimals lose precision, which is inherent to a grid.
    if isinstance(v, decimal.Decimal):
        if v.is_nan():
            return None
        return float(v)

    # Python int / float already serialisable
    if isinstance(v, (int, float)):
        return v

    # Strings are fine
    if isinstance(v, str):
        return v

    # Last resort: stringify
    return str(v)


def _result_to_dict(df: Any, pdf: Any, row_cap: int) -> dict:
    """Shape a (Spark DataFrame schema + pandas DataFrame) into a SparkResult dict.

    Parameters
    ----------
    df:
        The Spark DataFrame (``df.schema`` is used for column types).
    pdf:
        The pandas DataFrame from ``df.limit(row_cap + 1).toPandas()``.
    row_cap:
        The maximum number of rows to return.

    Returns
    -------
    dict matching the SparkResult TypeScript interface.
    """
    truncated = len(pdf) > row_cap
    pdf = pdf.iloc[:row_cap]

    schema = [
        {"name": field.name, "type": field.dataType.simpleString()}
        for field in df.schema.fields
    ]

    rows = [
        [_normalize_value(pdf.iloc[r, c]) for c in range(len(schema))]
        for r in range(len(pdf))
    ]

    return {
        "schema": schema,
        "rows": rows,
        "rowCount": len(rows),
        "truncated": truncated,
    }


# ---------------------------------------------------------------------------
# Error helper
# ---------------------------------------------------------------------------

def _error_json(exc: BaseException) -> str:
    """Return a JSON error envelope string."""
    return json.dumps({
        "ok": False,
        "error": {
            "name": type(exc).__name__,
            "message": str(exc),
        },
    })


# ---------------------------------------------------------------------------
# Public API — called from Pyodide via runPython snippets
# ---------------------------------------------------------------------------

def connect(uri: str, token: str | None) -> str:
    """Build (or replace) the module-global SparkSession.

    Parameters
    ----------
    uri:
        Spark Connect URI in sc:// form, e.g.
        ``sc://host:8081/;transport=grpcweb``.
        Must NOT already include a ``token=`` param (we append it here if
        ``token`` is given).
    token:
        Optional bearer token.  Appended to the URI as ``;token=<value>`` so
        PySpark's DefaultChannelBuilder maps it to the ``Authorization: Bearer``
        grpc-web header forwarded by Envoy.

    Returns
    -------
    JSON string: ``{"ok": true}`` or ``{"ok": false, "error": {...}}``.
    """
    global _spark
    try:
        import pyspark_connect_web as pcw  # noqa: F401
        pcw.install()  # idempotent

        from pyspark.sql import SparkSession

        # Append the token as a URI param if provided. Spark Connect URI params
        # are ';'-separated (sc://host:port/;transport=grpcweb;token=<value>),
        # which DefaultChannelBuilder maps to the Authorization: Bearer header.
        # Skip if the caller already included a token= param.
        effective_uri = uri
        if token and "token=" not in uri:
            effective_uri = uri.rstrip(";") + f";token={token}"

        # Stop any existing session before creating a new one.
        if _spark is not None:
            try:
                _spark.stop()
            except Exception:
                pass

        _spark = SparkSession.builder.remote(effective_uri).getOrCreate()
        return json.dumps({"ok": True})
    except Exception as exc:  # noqa: BLE001
        return _error_json(exc)


def run_sql(sql: str, row_cap: int) -> str:
    """Execute SQL and return at most ``row_cap`` rows as a SparkResult JSON string.

    On success returns a JSON object shaped like the TypeScript ``SparkResult``
    interface (no ``"ok"`` wrapper).  On failure returns ``{"ok": false, ...}``.

    The TypeScript side detects failure by checking ``"ok" in parsed &&
    !parsed.ok`` after JSON.parse.

    Row-cap semantics: fetches ``row_cap + 1`` rows; if the extra row is present
    the result is marked ``truncated: true`` and that row is dropped before
    returning.
    """
    try:
        if _spark is None:
            raise RuntimeError("Not connected — call connect() first.")
        df = _spark.sql(sql)
        pdf = df.limit(row_cap + 1).toPandas()
        result = _result_to_dict(df, pdf, row_cap)
        return json.dumps(result)
    except Exception as exc:  # noqa: BLE001
        return _error_json(exc)


def schema_of(sql: str) -> str:
    """Resolve the schema of a SQL expression without fetching any data.

    Uses ``spark.sql(sql).schema`` (schema resolution only, no execute/collect).

    Returns ``{"schema": [{name, type}, ...]}`` on success,
    ``{"ok": false, "error": {...}}`` on failure.
    """
    try:
        if _spark is None:
            raise RuntimeError("Not connected — call connect() first.")
        df = _spark.sql(sql)
        schema = [
            {"name": field.name, "type": field.dataType.simpleString()}
            for field in df.schema.fields
        ]
        return json.dumps({"schema": schema})
    except Exception as exc:  # noqa: BLE001
        return _error_json(exc)
