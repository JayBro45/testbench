"""
smr_excel_report.py
===================

SMR Excel Report Generator (Legacy-Compatible)

This module generates the **Engineering / Acceptance Excel report**
for SMR testing using the legacy acceptance logic and color conventions.

Responsibilities
----------------
- Convert UI grid data into an Excel worksheet
- Invoke SMRAcceptanceEngine for PASS / FAIL evaluation
- Apply legacy conditional formatting:
    • Red    → FAIL
    • Amber  → PASS but ABNORMAL
    • Green  → PASS
- Append a merged summary block at the bottom of the sheet

Strict Guarantees
-----------------
- Acceptance logic is NOT implemented here
- Column order, names, and layout are preserved
- No Excel formatting changes beyond legacy behavior

Dependencies
------------
- smr_acceptance_engine.SMRAcceptanceEngine
"""

from typing import List, Dict, Any

import pandas as pd
from xlsxwriter.utility import xl_col_to_name

from smr_acceptance_engine import SMRAcceptanceEngine


# =============================================================================
# Legacy Color Codes (DO NOT MODIFY)
# =============================================================================

COLOR_PASS = "#90ee90"      # Green
COLOR_FAIL = "#FFCCCB"      # Red
COLOR_ABNORMAL = "#FFBF00"  # Amber


# =============================================================================
# Public API
# =============================================================================

def generate_smr_excel_report(
    grid_rows: List[Dict[str, Any]],
    output_path: str
) -> None:
    """
    Generate the SMR Engineering / Acceptance Excel report.

    Parameters
    ----------
    grid_rows : List[Dict[str, Any]]
        Final grid rows captured from the UI.

    output_path : str
        Full filesystem path where the Excel file will be written.
    """

    # -------------------------------------------------------------------------
    # Run Acceptance Engine
    # -------------------------------------------------------------------------
    # We must evaluate the RAW rows. Converting to DataFrame/numeric first
    # might alter data strings, breaking the engine's exclusion logic.
    engine = SMRAcceptanceEngine(grid_rows)
    result = engine.evaluate()

    # -------------------------------------------------------------------------
    # Convert Grid → DataFrame (For Excel Output)
    # -------------------------------------------------------------------------
    df = pd.DataFrame(grid_rows)

    # Enforce numeric conversion for the Excel file itself (legacy behavior)
    for col in df.columns:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    # -------------------------------------------------------------------------
    # Write Excel
    # -------------------------------------------------------------------------
    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
        df.to_excel(writer, sheet_name="Result", index=False)

        workbook = writer.book
        worksheet = writer.sheets["Result"]

        # ---------------------------------------------------------------------
        # Formats (legacy semantics)
        # ---------------------------------------------------------------------
        fmt_fail = workbook.add_format({"bg_color": COLOR_FAIL})
        fmt_abnormal = workbook.add_format({"bg_color": COLOR_ABNORMAL})
        # fmt_pass is defined but implicitly used by default or summary
        fmt_pass = workbook.add_format({"bg_color": COLOR_PASS}) 

        fmt_summary_pass = workbook.add_format({
            "bg_color": COLOR_PASS,
            "text_wrap": True,
            "bold": True
        })

        fmt_summary_fail = workbook.add_format({
            "bg_color": COLOR_FAIL,
            "text_wrap": True,
            "bold": True
        })

        # ---------------------------------------------------------------------
        # Apply FAIL (Red) Cells
        # ---------------------------------------------------------------------
        for col_name, row_ids in result.invalid_cells.items():
            if not row_ids:
                continue

            # Ensure column exists in DF (safety check)
            if col_name not in df.columns:
                continue

            col_idx = df.columns.get_loc(col_name)
            col_letter = xl_col_to_name(col_idx)

            for row_id in row_ids:
                # Applies conditional format to the entire column based on the value
                # of the specific failing cell (Legacy AVR Logic)
                worksheet.conditional_format(
                    f"{col_letter}1:{col_letter}{len(df) + 1}",
                    {
                        "type": "cell",
                        "criteria": "equal to",
                        "value": f'${col_letter}${row_id}',
                        "format": fmt_fail,
                    },
                )

        # ---------------------------------------------------------------------
        # Apply ABNORMAL (Amber) Cells
        # ---------------------------------------------------------------------
        # safely retrieve abnormal_cells in case SMR engine does not implement it yet
        abnormal_cells = getattr(result, "abnormal_cells", {})
        
        for col_name, row_ids in abnormal_cells.items():
            if not row_ids:
                continue

            if col_name not in df.columns:
                continue

            col_idx = df.columns.get_loc(col_name)
            col_letter = xl_col_to_name(col_idx)

            for row_id in row_ids:
                worksheet.conditional_format(
                    f"{col_letter}1:{col_letter}{len(df) + 1}",
                    {
                        "type": "cell",
                        "criteria": "equal to",
                        "value": f'${col_letter}${row_id}',
                        "format": fmt_abnormal,
                    },
                )

        # ---------------------------------------------------------------------
        # Summary Block (Merged, Legacy Behavior)
        # ---------------------------------------------------------------------
        start_row = len(df) + 2
        summary_lines = result.summary.count("\n") + 1

        summary_format = (
            fmt_summary_pass if result.passed else fmt_summary_fail
        )

        worksheet.merge_range(
            start_row,
            0,
            start_row + summary_lines,
            len(df.columns) - 1,
            result.summary,
            summary_format,
        )