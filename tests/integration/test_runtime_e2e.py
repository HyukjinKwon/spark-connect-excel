# SPDX-License-Identifier: Apache-2.0
"""
Integration e2e for the Python runtime contract against a REAL Spark Connect
server (no browser, no Envoy). Mirrors how pyspark-connect-web tests itself.

It starts an in-process Spark Connect server via `SparkSession.builder.remote(
"local[*]")` (Spark 4.x embedded Connect) and drives the real result-shaping
functions in `spark_excel_runtime` (`run_sql` / `schema_of`), asserting the
SparkResult JSON is correct on real Spark output (types, rows, truncation,
nulls, errors) -- the part the fake-based unit tests can't cover.

Skips gracefully when pyspark or a JVM is unavailable (e.g. local dev sandbox);
CI installs pyspark + Java and runs it for real. We set the session directly on
the module rather than calling connect(), because connect() installs the
grpc-web (browser) transport patch, which is not what a server-side gRPC test
wants.
"""
import json
import os

import pytest

pyspark = pytest.importorskip("pyspark")  # skip whole module if not installed

import spark_excel_runtime as srt  # noqa: E402

# In CI we WANT this to run for real: set CI_REQUIRE_SPARK=1 so a startup failure
# is a hard error (surfaces the problem) instead of a silent skip. Locally
# (unset) it skips gracefully when no JVM/Spark is available.
_REQUIRE = os.environ.get("CI_REQUIRE_SPARK") == "1"


@pytest.fixture(scope="module")
def spark():
    """A real local Spark Connect session, or skip if Spark/JVM can't start."""
    try:
        from pyspark.sql import SparkSession

        session = SparkSession.builder.remote("local[*]").getOrCreate()
    except Exception as exc:  # noqa: BLE001
        if _REQUIRE:
            raise
        pytest.skip(f"Spark Connect server unavailable: {exc}")
    # Inject directly; bypass connect() (which installs the browser grpc-web patch).
    srt._spark = session
    yield session
    try:
        session.stop()
    finally:
        srt._spark = None


def test_run_sql_basic(spark):
    res = json.loads(srt.run_sql("SELECT id, id * 2 AS dbl FROM range(5)", 100))
    assert [c["name"] for c in res["schema"]] == ["id", "dbl"]
    assert all(c["type"] == "bigint" for c in res["schema"])
    assert res["rowCount"] == 5
    assert res["truncated"] is False
    assert res["rows"][0] == [0, 0]
    assert res["rows"][4] == [4, 8]


def test_run_sql_truncation(spark):
    res = json.loads(srt.run_sql("SELECT id FROM range(10)", 3))
    assert res["rowCount"] == 3
    assert res["truncated"] is True
    assert len(res["rows"]) == 3


def test_run_sql_types_and_nulls(spark):
    sql = (
        "SELECT CAST('2020-01-02' AS DATE) AS d, "
        "CAST('2020-01-02 03:04:05' AS TIMESTAMP) AS ts, "
        "1.5 AS dbl, true AS b, CAST(NULL AS STRING) AS n"
    )
    res = json.loads(srt.run_sql(sql, 10))
    row = res["rows"][0]
    types = {c["name"]: c["type"] for c in res["schema"]}
    assert types["d"] == "date" and types["ts"].startswith("timestamp")
    assert types["b"] == "boolean"
    # date/timestamp normalized to ISO-8601 strings.
    assert str(row[0]).startswith("2020-01-02")
    assert str(row[1]).startswith("2020-01-02")
    assert row[2] == 1.5
    assert row[3] is True
    assert row[4] is None  # NaN/None -> null


def test_schema_of_does_not_fetch_rows(spark):
    res = json.loads(srt.schema_of("SELECT 1 AS a, 'x' AS b"))
    assert "schema" in res and "rows" not in res
    assert [c["name"] for c in res["schema"]] == ["a", "b"]


def test_run_sql_groupby_agg(spark):
    sql = (
        "SELECT k, SUM(v) AS total FROM "
        "VALUES ('a', 1), ('a', 2), ('b', 5) AS t(k, v) "
        "GROUP BY k ORDER BY k"
    )
    res = json.loads(srt.run_sql(sql, 100))
    assert res["rowCount"] == 2
    assert res["rows"] == [["a", 3], ["b", 5]]


def test_run_sql_error_envelope(spark):
    res = json.loads(srt.run_sql("SELECT * FROM no_such_table_xyz", 10))
    assert res.get("ok") is False
    assert "error" in res and res["error"]["message"]
