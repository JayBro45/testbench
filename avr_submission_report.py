"""
avr_submission_report.py
=======================

AVR Submission Excel Report Generator

Generates the **official AVR test report** intended for
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
- Section C: Data rows (any number, usually 6 rows)
"""

import os
from datetime import datetime
from typing import List, Dict

import pandas as pd


# =============================================================================
# Public API
# =============================================================================

def generate_avr_submission_excel(
    grid_rows: List[Dict[str, float | str]],
    output_path: str,
    serial_number: str | None = None
) -> None:
    """
    Generate the official AVR submission Excel report.

    Parameters
    ----------
    grid_rows : List[Dict[str, float | str]]
        Rows captured from the UI grid.

    output_path : str
        Full filesystem path for the generated Excel file.

    serial_number : str | None
        Serial number to display in the report header. If None, derived from
        the output path (folder name used during export).

    Raises
    ------
    ValueError
        If grid_rows is empty.
    """

    if len(grid_rows) == 0:
        raise ValueError("Submission report requires at least 1 row")

    # Derive serial from export folder name if not provided
    if serial_number is None:
        serial_number = os.path.basename(os.path.dirname(output_path)) or ""

    # -------------------------------------------------------------------------
    # Build DataFrame (Rows 5–10)
    # -------------------------------------------------------------------------
    df = pd.DataFrame(grid_rows)[[
        "Frequency", "V (in)", "I (in)", "kW (in)",
        "V (out)", "I (out)", "kW (out)",
        "VTHD (out)", "Efficiency", "Load", "Line"
    ]]

    # -------------------------------------------------------------------------
    # Write Excel
    # -------------------------------------------------------------------------
    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
        workbook = writer.book
        worksheet = workbook.add_worksheet("AVR TEST REPORT")
        writer.sheets["AVR TEST REPORT"] = worksheet

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
        # Column Widths (Legacy-compatible, untouched)
        # ---------------------------------------------------------------------
        widths = [10, 9, 9, 9, 9, 9, 10, 10, 10, 9, 9]
        for col, w in enumerate(widths):
            worksheet.set_column(col, col, w)

        # ---------------------------------------------------------------------
        # SECTION A — Header
        # ---------------------------------------------------------------------
        serial_text = f"Serial No. : {serial_number}" if serial_number else "Serial No. :"
        worksheet.merge_range("A1:C1", serial_text, fmt_left_bold)
        worksheet.merge_range("D1:H1", "AVR TEST REPORT", fmt_title)
        worksheet.write("I1", "Date ", fmt_right_bold)
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
            "Frequency", "V (in)", "I (in)", "kW (in)",
            "V (out)", "I (out)", "kW (out)",
            "VTHD (out)", "Efficiency", "Regulation", ""
        ]

        units = [
            "Hz", "V", "A", "kW",
            "V", "A", "kW",
            "V", "%", "Load", "Line"
        ]

        # Header row
        for col, h in enumerate(headers):
            worksheet.write(1, col, h, fmt_header)

        # Merge Regulation header (J–K)
        worksheet.merge_range(1, 9, 1, 10, "Regulation", fmt_header)

        # Unit row
        for col, u in enumerate(units):
            worksheet.write(2, col, u, fmt_header)

        worksheet.set_row(1, 24)
        worksheet.set_row(2, 22)

        # ---------------------------------------------------------------------
        # SECTION C — Data Rows
        # ---------------------------------------------------------------------
        last_data_row = 3 + len(df) - 1
        for r, row in enumerate(df.itertuples(index=False), start=3):
            for c, val in enumerate(row):
                worksheet.write(r, c, val, fmt_cell)

        # Larger row height for readability
        for r in range(3, last_data_row + 1):
            worksheet.set_row(r, 22)

        # ---------------------------------------------------------------------
        # Print Settings (A4)
        # ---------------------------------------------------------------------
        worksheet.set_paper(9)  # A4
        worksheet.set_landscape()
        worksheet.center_horizontally()
        worksheet.center_vertically()
        worksheet.fit_to_pages(1,1)
        worksheet.set_margins(left=0.7, right=0.7, top=0.5, bottom=2)
        worksheet.print_area(f"A1:K{last_data_row + 1}")
