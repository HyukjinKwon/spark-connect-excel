# SPDX-License-Identifier: Apache-2.0
"""Make the canonical runtime module (python/spark_excel_runtime.py) importable."""
import sys
from pathlib import Path

_RUNTIME_DIR = Path(__file__).resolve().parents[2] / "python"
if str(_RUNTIME_DIR) not in sys.path:
    sys.path.insert(0, str(_RUNTIME_DIR))
