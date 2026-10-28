"""Run the consolidation on the bundled synthetic files without persisting any output.

Used by the Streamlit app; kept separate so it can be tested without Streamlit.
"""
from __future__ import annotations

import tempfile
from pathlib import Path

import pandas as pd

from consolidate import Result, consolidate

SAMPLE_DIR = Path(__file__).resolve().parent / "sample_data" / "input"
# Mirrors the issues injected by generate_sample_data.py.
SAMPLE_ISSUES = {
    "sales_2026-01.xlsx": "None (clean file)",
    "sales_2026-02.xlsx": "2 rows re-exported from January (duplicates)",
    "sales_2026-03.xlsx": "Customer IDs saved as numbers, leading zeros lost",
    "sales_2026-04.xlsx": "5 typing errors: bad date, text quantity, blank name, negative price, bad invoice ID",
    "sales_2026-05.xlsx": "Messy headers, text dates, 1 conflicting duplicate",
    "sales_2026-06.xlsx": "Quantity column missing",
    "sales_2026-07.xlsx": "None (clean file)",
}


def describe_inputs(input_dir: Path = SAMPLE_DIR) -> pd.DataFrame:
    rows = [{"File": p.name, "Rows": len(pd.read_excel(p)),
             "Deliberate problem": SAMPLE_ISSUES.get(p.name, "Unknown")}
            for p in sorted(Path(input_dir).glob("*.xlsx"))]
    return pd.DataFrame(rows)


def run_sample(input_dir: Path = SAMPLE_DIR) -> tuple[Result, bytes]:
    """Consolidate into a throwaway folder and return the result plus the report bytes.

    The temporary folder (and report file) is deleted before returning, so nothing
    is stored on disk between runs or shared between visitors.
    """
    with tempfile.TemporaryDirectory(prefix="xl-demo-") as tmp:
        result = consolidate(input_dir, tmp)
        report = result.report_path.read_bytes()
    return result, report
