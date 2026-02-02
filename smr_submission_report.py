"""
smr_submission_report.py
========================

SMR Submission Excel Report Generator

Generates the **official SMR test report** intended for
submission along with the physical device.

Key Characteristics
-------------------
- EXACT legacy layout (submission format)
- NO acceptance / PASS / FAIL logic
- NO conditional coloring
- Clean, neutral, printable Excel output
- Optimized for A4 printing

Layout Rules
------------
- Section A: Report metadata (title, serial, date)
- Section B: Two-row table header
- Section C: Data rows
"""

from datetime import datetime
from typing import List, Dict, Any

import pandas as pd


# =============================================================================
# Public API
# =============================================================================

def generate_smr_submission_excel(
    grid_rows: List[Dict[str, Any]],
    output_path: str
) -> None:
    """
    Generate the official SMR submission Excel report.

    Parameters
    ----------
    grid_rows : List[Dict[str, Any]]
        Rows captured from the UI grid.

    output_path : str
        Full filesystem path for the generated Excel file.
    """

    # -------------------------------------------------------------------------
    # Build DataFrame
    # -------------------------------------------------------------------------
    # Define the exact columns to include in the submission report
    columns_to_export = [
        "V (in)", "I (in)", "P (in)", "PF (in)",
        "Vthd % (in)", "Ithd % (in)",
        "V (out)", "I (out)", "P (out)",
        "Ripple (out)", "Efficiency"
    ]
    
    # Filter rows to just these columns, handling missing keys gracefully
    df = pd.DataFrame(grid_rows)
    # Ensure all columns exist (add NaN if missing) to prevent key errors
    for col in columns_to_export:
        if col not in df.columns:
            df[col] = ""
            
    df = df[columns_to_export]

    # -------------------------------------------------------------------------
    # Write Excel
    # -------------------------------------------------------------------------
    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
        workbook = writer.book
        worksheet = workbook.add_worksheet("SMR TEST REPORT")
        writer.sheets["SMR TEST REPORT"] = worksheet

        # ---------------------------------------------------------------------
        # Gridlines (Legacy: visible everywhere)
        # ---------------------------------------------------------------------
        worksheet.hide_gridlines(2)

        # ---------------------------------------------------------------------
        # Formats
        # ---------------------------------------------------------------------
        fmt_title = workbook.add_format({
            "bold": True,
            "align": "center",
            "valign": "vcenter",
            "border": 1,
            "font_size": 15
        })

        fmt_header = workbook.add_format({
            "bold": True,
            "align": "center",
            "valign": "vcenter",
            "border": 1,
            "font_size": 10,
            "bg_color": "FFFF00"
        })

        fmt_cell = workbook.add_format({
            "align": "center",
            "valign": "vcenter",
            "border": 1,
            "font_size": 11  
        })

        fmt_left_bold = workbook.add_format({
            "bold": True,
            "align": "left",
            "valign": "vcenter",
            "border": 1,
            "font_size": 10
        })

        fmt_right_bold = workbook.add_format({
            "bold": True,
            "align": "right", 
            "valign": "vcenter",
            "border": 1,      
            "font_size": 10
        })

        # ---------------------------------------------------------------------
        # Column Widths
        # ---------------------------------------------------------------------
        # Adjust widths slightly for SMR specific data lengths
        widths = [10, 10, 10, 8, 10, 10, 10, 10, 10, 12, 11]
        for col, w in enumerate(widths):
            worksheet.set_column(col, col, w)

        # ---------------------------------------------------------------------
        # SECTION A — Header
        # ---------------------------------------------------------------------
        # Serial No (A1:C1)
        worksheet.merge_range("A1:C1", "Serial No. :", fmt_left_bold)
        
        # Title (D1:H1) - Adjusted range for SMR column count
        # Total columns = 11 (Indices 0-10, A-K). 
        # Center title roughly in the middle (D-G or D-H)
        worksheet.merge_range("D1:H1", "SMR TEST REPORT", fmt_title)
        
        # Date Label (I1)
        worksheet.write("I1", "Date ", fmt_right_bold)
        
        # Date Value (J1:K1)
        worksheet.merge_range(
            "J1:K1",
            datetime.now().strftime("%d-%m-%Y"),
            fmt_cell
        )

        worksheet.set_row(0, 32)

        # ---------------------------------------------------------------------
        # SECTION B — Table Header
        # ---------------------------------------------------------------------
        headers = [
            "V (in)", "I (in)", "P (in)", "PF (in)",
            "Vthd (in)", "Ithd (in)",
            "V (out)", "I (out)", "P (out)",
            "Ripple", "Efficiency"
        ]

        units = [
            "V", "A", "kW", "pf",
            "%", "%",
            "V", "A", "kW",
            "V", "%"
        ]

        # Header row
        for col, h in enumerate(headers):
            worksheet.write(1, col, h, fmt_header)

        # Unit row
        for col, u in enumerate(units):
            worksheet.write(2, col, u, fmt_header)

        worksheet.set_row(1, 24)
        worksheet.set_row(2, 22)

        # ---------------------------------------------------------------------
        # SECTION C — Data Rows
        # ---------------------------------------------------------------------
        for r, row in enumerate(df.itertuples(index=False), start=3):
            for c, val in enumerate(row):
                worksheet.write(r, c, val, fmt_cell)

        # Larger row height for readability (apply to first 6 rows at least)
        for r in range(3, 9):
            worksheet.set_row(r, 22)

        # ---------------------------------------------------------------------
        # Print Settings (A4)
        # ---------------------------------------------------------------------
        worksheet.set_paper(9)  # A4
        worksheet.set_landscape()
        worksheet.center_horizontally()
        worksheet.center_vertically()
        worksheet.fit_to_pages(1, 1)
        worksheet.set_margins(left=0.7, right=0.7, top=0.5, bottom=2)
        # Print area covers all columns (A to K)
        worksheet.print_area(f"A1:K{3 + len(df)}")