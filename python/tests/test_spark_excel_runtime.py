# SPDX-License-Identifier: Apache-2.0
"""Unit tests for spark_excel_runtime.py.

Tests the pure-Python normalisation helpers and result-shaping logic using a
fake Spark/DataFrame stub.  No live Spark cluster, no grpcio, no Pyodide
required.

If pandas is not importable the entire module is skipped gracefully.
"""
from __future__ import annotations

import datetime
import json
import math
import sys
import types

import pytest

pandas = pytest.importorskip("pandas")
import pandas as pd  # noqa: E402 — after importorskip

# ---------------------------------------------------------------------------
# Bring in the module under test.  It lives two directories up from this file.
# ---------------------------------------------------------------------------
import importlib
import pathlib

_RUNTIME_DIR = pathlib.Path(__file__).parent.parent
if str(_RUNTIME_DIR) not in sys.path:
    sys.path.insert(0, str(_RUNTIME_DIR))

import spark_excel_runtime as _srt  # noqa: E402


# ---------------------------------------------------------------------------
# Fake Spark stubs
# ---------------------------------------------------------------------------

class FakeDataType:
    """Stub for a Spark DataType."""

    def __init__(self, simple: str) -> None:
        self._simple = simple

    def simpleString(self) -> str:  # noqa: N802
        return self._simple


class FakeField:
    """Stub for a Spark StructField."""

    def __init__(self, name: str, type_str: str) -> None:
        self.name = name
        self.dataType = FakeDataType(type_str)  # noqa: N815


class FakeSchema:
    """Stub for a Spark StructType."""

    def __init__(self, fields: list[FakeField]) -> None:
        self.fields = fields


class FakeDataFrame:
    """Stub for a Spark DataFrame; pre-loaded with a fixed pandas result."""

    def __init__(
        self,
        schema: FakeSchema,
        pdf: pd.DataFrame,
        *,
        limit_cap: int | None = None,
    ) -> None:
        self.schema = schema
        self._pdf = pdf
        self._limit_cap = limit_cap  # how many rows limit() returns

    def limit(self, n: int) -> "FakeDataFrame":  # noqa: N802
        cap = self._limit_cap if self._limit_cap is not None else n
        sliced = self._pdf.iloc[:min(n, cap)]
        return FakeDataFrame(self.schema, sliced)

    def toPandas(self) -> pd.DataFrame:  # noqa: N802
        return self._pdf.copy()


class FakeSpark:
    """Stub for a SparkSession."""

    def __init__(
        self,
        schema: FakeSchema,
        pdf: pd.DataFrame,
        *,
        limit_cap: int | None = None,
    ) -> None:
        self._df = FakeDataFrame(schema, pdf, limit_cap=limit_cap)

    def sql(self, _query: str) -> FakeDataFrame:
        return self._df


# ---------------------------------------------------------------------------
# Helper: build a schema from (name, type_str) pairs
# ---------------------------------------------------------------------------

def _schema(*cols: tuple[str, str]) -> FakeSchema:
    return FakeSchema([FakeField(name, t) for name, t in cols])


# ---------------------------------------------------------------------------
# _normalize_value tests
# ---------------------------------------------------------------------------

class TestNormalizeValue:
    def test_none_stays_none(self) -> None:
        assert _srt._normalize_value(None) is None

    def test_pandas_na_becomes_none(self) -> None:
        assert _srt._normalize_value(pd.NA) is None

    def test_pandas_nat_becomes_none(self) -> None:
        assert _srt._normalize_value(pd.NaT) is None

    def test_bool_stays_bool(self) -> None:
        assert _srt._normalize_value(True) is True
        assert _srt._normalize_value(False) is False

    def test_int_stays_int(self) -> None:
        result = _srt._normalize_value(42)
        assert result == 42
        assert isinstance(result, int)

    def test_float_stays_float(self) -> None:
        result = _srt._normalize_value(3.14)
        assert abs(result - 3.14) < 1e-10

    def test_float_nan_becomes_none(self) -> None:
        assert _srt._normalize_value(float("nan")) is None

    def test_float_inf_becomes_none(self) -> None:
        assert _srt._normalize_value(float("inf")) is None
        assert _srt._normalize_value(float("-inf")) is None

    def test_string_passthrough(self) -> None:
        assert _srt._normalize_value("hello") == "hello"

    def test_decimal_becomes_float(self) -> None:
        import decimal

        result = _srt._normalize_value(decimal.Decimal("1.5"))
        assert isinstance(result, float)
        assert result == 1.5

    def test_decimal_nan_becomes_none(self) -> None:
        import decimal

        assert _srt._normalize_value(decimal.Decimal("NaN")) is None

    def test_datetime_to_iso(self) -> None:
        dt = datetime.datetime(2024, 3, 15, 12, 0, 0)
        assert _srt._normalize_value(dt) == "2024-03-15T12:00:00"

    def test_date_to_iso(self) -> None:
        d = datetime.date(2024, 3, 15)
        assert _srt._normalize_value(d) == "2024-03-15"

    def test_pandas_timestamp_to_iso(self) -> None:
        ts = pd.Timestamp("2024-03-15 12:00:00")
        result = _srt._normalize_value(ts)
        assert result == "2024-03-15T12:00:00"

    def test_numpy_bool(self) -> None:
        import numpy as np
        result = _srt._normalize_value(np.bool_(True))
        assert result is True
        assert isinstance(result, bool)

    def test_numpy_int(self) -> None:
        import numpy as np
        result = _srt._normalize_value(np.int64(99))
        assert result == 99
        assert isinstance(result, int)

    def test_numpy_float(self) -> None:
        import numpy as np
        result = _srt._normalize_value(np.float64(1.5))
        assert abs(result - 1.5) < 1e-10

    def test_numpy_float_nan(self) -> None:
        import numpy as np
        assert _srt._normalize_value(np.float64("nan")) is None

    def test_arbitrary_object_stringified(self) -> None:
        class _Obj:
            def __str__(self) -> str:
                return "custom"
        assert _srt._normalize_value(_Obj()) == "custom"


# ---------------------------------------------------------------------------
# _result_to_dict tests
# ---------------------------------------------------------------------------

class TestResultToDict:
    def _make(
        self,
        cols: list[tuple[str, str]],
        data: list[list],
        row_cap: int,
        *,
        limit_cap: int | None = None,
    ) -> dict:
        """Create a fake DF and call _result_to_dict."""
        schema = _schema(*cols)
        col_names = [c[0] for c in cols]
        pdf_full = pd.DataFrame(data, columns=col_names)
        df_fake = FakeDataFrame(schema, pdf_full, limit_cap=limit_cap)
        # Simulate what run_sql does: df.limit(row_cap + 1).toPandas()
        limited = df_fake.limit(row_cap + 1)
        pdf = limited.toPandas()
        return _srt._result_to_dict(df_fake, pdf, row_cap)

    def test_basic_shape(self) -> None:
        result = self._make(
            [("id", "bigint"), ("name", "string")],
            [[1, "alice"], [2, "bob"]],
            row_cap=10,
        )
        assert result["schema"] == [
            {"name": "id", "type": "bigint"},
            {"name": "name", "type": "string"},
        ]
        assert result["rows"] == [[1, "alice"], [2, "bob"]]
        assert result["rowCount"] == 2
        assert result["truncated"] is False

    def test_truncation_flag_set_when_more_rows_than_cap(self) -> None:
        # 5 rows in underlying table, cap=3 → fetch 4, see 4 > 3, truncate
        data = [[i, f"r{i}"] for i in range(5)]
        result = self._make(
            [("id", "bigint"), ("label", "string")],
            data,
            row_cap=3,
        )
        assert result["truncated"] is True
        assert result["rowCount"] == 3
        assert len(result["rows"]) == 3

    def test_no_truncation_when_exactly_at_cap(self) -> None:
        data = [[i] for i in range(5)]
        result = self._make([("id", "bigint")], data, row_cap=5)
        assert result["truncated"] is False
        assert result["rowCount"] == 5

    def test_no_truncation_when_fewer_than_cap(self) -> None:
        data = [[i] for i in range(3)]
        result = self._make([("id", "bigint")], data, row_cap=10)
        assert result["truncated"] is False
        assert result["rowCount"] == 3

    def test_iso_date_conversion_in_rows(self) -> None:
        schema = _schema(("created", "date"))
        pdf = pd.DataFrame(
            {"created": [datetime.date(2024, 1, 1), datetime.date(2024, 6, 15)]}
        )
        df_fake = FakeDataFrame(schema, pdf)
        limited = df_fake.limit(11)
        result = _srt._result_to_dict(df_fake, limited.toPandas(), 10)
        assert result["rows"][0][0] == "2024-01-01"
        assert result["rows"][1][0] == "2024-06-15"

    def test_null_nan_handling_in_rows(self) -> None:
        import numpy as np
        schema = _schema(("val", "double"), ("label", "string"))
        pdf = pd.DataFrame(
            {
                "val": [1.0, float("nan"), None],
                "label": ["a", None, "c"],
            }
        )
        df_fake = FakeDataFrame(schema, pdf)
        limited = df_fake.limit(4)
        result = _srt._result_to_dict(df_fake, limited.toPandas(), 3)
        assert result["rows"][0] == [1.0, "a"]
        assert result["rows"][1][0] is None  # NaN -> None
        assert result["rows"][2][0] is None  # None -> None

    def test_type_mapping_in_schema(self) -> None:
        result = self._make(
            [
                ("a", "bigint"),
                ("b", "string"),
                ("c", "double"),
                ("d", "boolean"),
                ("e", "timestamp"),
            ],
            [[1, "x", 1.5, True, datetime.datetime(2024, 1, 1)]],
            row_cap=10,
        )
        types = [col["type"] for col in result["schema"]]
        assert types == ["bigint", "string", "double", "boolean", "timestamp"]

    def test_empty_result_set(self) -> None:
        result = self._make([("id", "bigint")], [], row_cap=10)
        assert result["rows"] == []
        assert result["rowCount"] == 0
        assert result["truncated"] is False


# ---------------------------------------------------------------------------
# Token redaction tests
# ---------------------------------------------------------------------------

class TestTokenRedaction:
    """Verify that bearer tokens never appear in error envelopes."""

    def test_redact_removes_token_value(self) -> None:
        result = _srt._redact("sc://host:443/;transport=grpcweb;token=supersecret123")
        assert "supersecret123" not in result
        assert ";token=***" in result

    def test_redact_leaves_text_without_token_unchanged(self) -> None:
        text = "some error without token info"
        assert _srt._redact(text) == text

    def test_redact_case_insensitive(self) -> None:
        result = _srt._redact("sc://host/;TOKEN=MySecret")
        assert "MySecret" not in result
        assert ";token=***" in result.lower()

    def test_error_json_does_not_leak_token(self) -> None:
        """A connect URI with a token embedded must not appear in the error JSON."""
        token_value = "very_secret_bearer_token_abc123"
        exc = RuntimeError(
            f"Connection failed: sc://host:443/;transport=grpcweb;token={token_value}"
        )
        s = _srt._error_json(exc)
        parsed = json.loads(s)
        assert parsed["ok"] is False
        # The raw token value must not appear anywhere in the JSON string
        assert token_value not in s
        # The error name and a redacted message should still be present
        assert parsed["error"]["name"] == "RuntimeError"
        assert ";token=***" in parsed["error"]["message"]

    def test_connect_error_does_not_leak_token(self) -> None:
        """Simulate a connect() failure and verify the token is redacted."""
        import sys
        import types

        # Build a minimal fake pyspark_connect_web stub that raises on getOrCreate
        pcw_stub = types.ModuleType("pyspark_connect_web")
        pcw_stub.install = lambda: None  # type: ignore[attr-defined]
        sys.modules.setdefault("pyspark_connect_web", pcw_stub)

        class _FakeBuilder:
            def remote(self, _uri: str) -> "_FakeBuilder":
                return self
            def getOrCreate(self) -> None:
                raise RuntimeError(
                    "UNAVAILABLE: no route to host sc://bad:443/;token=leaked_token_xyz"
                )

        class _FakeSession:
            builder = _FakeBuilder()

        pyspark_stub = types.ModuleType("pyspark")
        pyspark_sql_stub = types.ModuleType("pyspark.sql")
        pyspark_sql_stub.SparkSession = _FakeSession  # type: ignore[attr-defined]
        sys.modules["pyspark"] = pyspark_stub
        sys.modules["pyspark.sql"] = pyspark_sql_stub

        orig_spark = _srt._spark
        try:
            result_str = _srt.connect("sc://bad:443/;transport=grpcweb", "leaked_token_xyz")
            parsed = json.loads(result_str)
            assert parsed["ok"] is False
            assert "leaked_token_xyz" not in result_str
        finally:
            _srt._spark = orig_spark
            sys.modules.pop("pyspark", None)
            sys.modules.pop("pyspark.sql", None)


# ---------------------------------------------------------------------------
# Error JSON shape tests
# ---------------------------------------------------------------------------

class TestErrorShape:
    def test_error_json_shape(self) -> None:
        exc = ValueError("bad query")
        s = _srt._error_json(exc)
        parsed = json.loads(s)
        assert parsed["ok"] is False
        assert parsed["error"]["name"] == "ValueError"
        assert "bad query" in parsed["error"]["message"]

    def test_run_sql_without_connection_returns_error(self) -> None:
        # Temporarily clear the global spark session
        orig = _srt._spark
        _srt._spark = None
        try:
            result_str = _srt.run_sql("SELECT 1", 10)
            parsed = json.loads(result_str)
            assert parsed["ok"] is False
            assert "name" in parsed["error"]
            assert "message" in parsed["error"]
        finally:
            _srt._spark = orig

    def test_schema_of_without_connection_returns_error(self) -> None:
        orig = _srt._spark
        _srt._spark = None
        try:
            result_str = _srt.schema_of("SELECT 1")
            parsed = json.loads(result_str)
            assert parsed["ok"] is False
        finally:
            _srt._spark = orig


# ---------------------------------------------------------------------------
# run_sql with fake spark (end-to-end through the module function)
# ---------------------------------------------------------------------------

class TestRunSqlWithFakeSpark:
    def setup_method(self) -> None:
        self._orig_spark = _srt._spark

    def teardown_method(self) -> None:
        _srt._spark = self._orig_spark

    def _inject_fake(
        self,
        cols: list[tuple[str, str]],
        data: list[list],
        *,
        limit_cap: int | None = None,
    ) -> None:
        schema = _schema(*cols)
        col_names = [c[0] for c in cols]
        pdf = pd.DataFrame(data, columns=col_names)
        _srt._spark = FakeSpark(schema, pdf, limit_cap=limit_cap)

    def test_basic_query_returns_spark_result(self) -> None:
        self._inject_fake(
            [("id", "bigint"), ("val", "double")],
            [[1, 2.5], [2, 3.5]],
        )
        raw = _srt.run_sql("SELECT id, val FROM t", 100)
        parsed = json.loads(raw)
        # Success = no "ok" wrapper at the top level (just schema/rows/etc.)
        assert "ok" not in parsed or parsed.get("ok") is not False
        assert parsed["schema"] == [
            {"name": "id", "type": "bigint"},
            {"name": "val", "type": "double"},
        ]
        assert parsed["rows"] == [[1, 2.5], [2, 3.5]]
        assert parsed["truncated"] is False

    def test_truncation_in_run_sql(self) -> None:
        # 6 rows in the table, cap = 4
        self._inject_fake(
            [("n", "bigint")],
            [[i] for i in range(6)],
        )
        raw = _srt.run_sql("SELECT n FROM t", 4)
        parsed = json.loads(raw)
        assert parsed["truncated"] is True
        assert parsed["rowCount"] == 4

    def test_result_is_valid_json(self) -> None:
        self._inject_fake([("x", "string")], [["hello"]])
        raw = _srt.run_sql("SELECT x FROM t", 10)
        # Should not raise
        json.loads(raw)


# ---------------------------------------------------------------------------
# schema_of with fake spark
# ---------------------------------------------------------------------------

class TestSchemaOfWithFakeSpark:
    def setup_method(self) -> None:
        self._orig_spark = _srt._spark

    def teardown_method(self) -> None:
        _srt._spark = self._orig_spark

    def test_returns_schema_without_data(self) -> None:
        schema = _schema(("a", "bigint"), ("b", "string"), ("c", "timestamp"))
        pdf = pd.DataFrame({"a": [], "b": [], "c": []})
        _srt._spark = FakeSpark(schema, pdf)

        raw = _srt.schema_of("SELECT a, b, c FROM t")
        parsed = json.loads(raw)
        assert "schema" in parsed
        assert parsed["schema"] == [
            {"name": "a", "type": "bigint"},
            {"name": "b", "type": "string"},
            {"name": "c", "type": "timestamp"},
        ]

    def test_error_returns_ok_false(self) -> None:
        class _ErrorSpark:
            def sql(self, _q: str) -> None:
                raise RuntimeError("analysis error")

        _srt._spark = _ErrorSpark()
        raw = _srt.schema_of("SELECT bad_col FROM missing_table")
        parsed = json.loads(raw)
        assert parsed["ok"] is False
        assert parsed["error"]["name"] == "RuntimeError"
