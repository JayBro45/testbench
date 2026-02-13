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

import os
from datetime import datetime
from typing import List, Dict, Any

import pandas as pd


# =============================================================================
# Public API
# =============================================================================

def generate_smr_submission_excel(
    grid_rows: List[Dict[str, Any]],
    output_path: str,
    serial_number: str | None = None
) -> None:
    """
    Generate the official SMR submission Excel report.

    Parameters
    ----------
    grid_rows : List[Dict[str, Any]]
        Rows captured from the UI grid.

    output_path : str
        Full filesystem path for the generated Excel file.

    serial_number : str | None
        Serial number to display in the report header. If None, derived from
        the output path (folder name used during export).
    """

    # Derive serial from export folder name if not provided
    if serial_number is None:
        serial_number = os.path.basename(os.path.dirname(output_path)) or ""

    # -------------------------------------------------------------------------
    # Build DataFrame
    # -------------------------------------------------------------------------
    # Define the exact columns to include in the submission report
    columns_to_export = [
        "V (in)", "I (in)", "P (in)", "PF (in)",
        "Vthd % (in)", "Ithd % (in)",
        "V (out)", "I (out)", "P (out)",
        "Ripple (out)", "Efficiency", "PSO (mV)"
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
            "font_size": 15,
            "font_name": "Times New Roman"
        })

        fmt_header = workbook.add_format({
            "bold": True,
            "align": "center",
            "valign": "vcenter",
            "border": 1,
            "font_size": 10,
            "bg_color": "FFFF00",
            "font_name": "Times New Roman"
        })

        fmt_cell = workbook.add_format({
            "align": "center",
            "valign": "vcenter",
            "border": 1,
            "font_size": 11,
            "font_name": "Times New Roman"
        })

        fmt_left_bold = workbook.add_format({
            "bold": True,
            "align": "left",
            "valign": "vcenter",
            "border": 1,
            "font_size": 10,
            "font_name": "Times New Roman"
        })

        fmt_right_bold = workbook.add_format({
            "bold": True,
            "align": "right", 
            "valign": "vcenter",
            "border": 1,      
            "font_size": 10,
            "font_name": "Times New Roman"
        })

        # ---------------------------------------------------------------------
        # Column Widths
        # ---------------------------------------------------------------------
        widths = [10] * 12
        for col, w in enumerate(widths):
            worksheet.set_column(col, col, w)

        # ---------------------------------------------------------------------
        # SECTION A — Header
        # ---------------------------------------------------------------------
        serial_text = f"Serial No. : {serial_number}" if serial_number else "Serial No. :"
        worksheet.merge_range("A1:C1", serial_text, fmt_left_bold)
        
        # Title (D1:H1) - Adjusted range for SMR column count
        # Total columns = 12 (Indices 0-11, A-L).
        # Center title roughly in the middle (D-H)
        worksheet.merge_range("D1:H1", "SMR TEST REPORT", fmt_title)
        
        # Date Label (I1)
        worksheet.write("I1", "Date ", fmt_right_bold)
        
        # Date Value (J1:L1)
        worksheet.merge_range(
            "J1:L1",
            datetime.now().strftime("%d-%m-%Y"),
            fmt_cell
        )

        worksheet.set_row(0, 32)

        # ---------------------------------------------------------------------
        # SECTION B — Table Header
        # ---------------------------------------------------------------------
        headers = [
            "Vac (in)", "Iac (in)", "Pac (in)", "PF (in)",
            "Vthd% (in)", "Ithd% (in)",
            "Vdc (out)", "Idc (out)", "Pdc (out)",
            "Ripple", "Efficiency", "PSO"
        ]

        units = [
            "V", "A", "W", "PF",
            "%", "%",
            "V", "A", "W",
            "mV", "%", "mV"
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
        total_data_rows = len(df)        
        for r, row in enumerate(df.itertuples(index=False), start=3):
            # Set the row height for the current row dynamically
            worksheet.set_row(r, 14)        
            for c, val in enumerate(row):
                worksheet.write(r, c, val, fmt_cell)

        # ---------------------------------------------------------------------
        # Print Settings (A4)
        # ---------------------------------------------------------------------
        worksheet.set_paper(9)  # A4
        worksheet.set_landscape()
        worksheet.center_horizontally()        
        worksheet.fit_to_pages(1, 1)
        
        worksheet.set_margins(left=0.7, right=0.7, top=1.1 , bottom=1.5)
        worksheet.print_area(f"A1:L{3 + total_data_rows}")