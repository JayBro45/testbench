"""
smr_excel_report.py
===================

Excel Report Generator for SMR Tests.
Generates a professional report with:
- Submission Sheet (Clean data)
- Result Sheet (Color-coded validation)
"""

import pandas as pd
from typing import List, Dict, Any
from smr_acceptance_engine import SMRAcceptanceEngine

def generate_smr_excel_report(rows: List[Dict[str, Any]], output_path: str) -> None:
    """
    Creates an Excel file with SMR test results.
    
    :param rows: Grid data.
    :param output_path: Destination file path.
    """
    
    # 1. Run Engine to get Pass/Fail context
    engine = SMRAcceptanceEngine(rows)
    result = engine.evaluate()
    
    # 2. Prepare Dataframe
    df = pd.DataFrame(rows)
    
    # Numeric conversion for Excel formatting
    numeric_cols = [
        'V (in)', 'I (in)', 'P (in)', 'PF (in)', 
        'Vthd % (in)', 'Ithd % (in)', 
        'V (out)', 'I (out)', 'P (out)', 
        'Ripple (out)', 'Efficiency'
    ]
    for c in numeric_cols:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors='coerce')

    # 3. Create Excel Writer
    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
        workbook = writer.book
        
        # --- FORMATS ---
        header_fmt = workbook.add_format({
            "bold": True, "bg_color": "#DDEBF7", "border": 1, "align": "center"
        })
        cell_fmt = workbook.add_format({"border": 1, "align": "center"})
        fail_fmt = workbook.add_format({"bg_color": "#FFC7CE", "font_color": "#9C0006", "border": 1})
        pass_fmt = workbook.add_format({"bg_color": "#C6EFCE", "font_color": "#006100", "border": 1})

        # --- SHEET 1: SUBMISSION (Clean) ---
        df.to_excel(writer, sheet_name="Submission", index=False)
        ws_sub = writer.sheets["Submission"]
        _apply_formatting(ws_sub, df, header_fmt, cell_fmt)

        # --- SHEET 2: RESULT (Validation) ---
        df.to_excel(writer, sheet_name="Result", index=False)
        ws_res = writer.sheets["Result"]
        _apply_formatting(ws_res, df, header_fmt, cell_fmt)
        
        # Apply Conditional Formatting based on Invalid Cells
        for col_name, row_ids in result.invalid_cells.items():
            if not row_ids: continue
            
            # Find column index
            if col_name not in df.columns: continue
            col_idx = df.columns.get_loc(col_name)
            
            # Highlight Cells
            for row_id_str in row_ids:
                # Excel Row = int(row_id) - 1 (0-based)
                # But our row_ids are 1-based header+data style (e.g., "2" is first data row)
                # Dataframe index is 0-based.
                # If row_id is "2", that is DF index 0.
                try:
                    r_idx = int(row_id_str) - 2 
                    if 0 <= r_idx < len(df):
                        ws_res.write(r_idx + 1, col_idx, df.iloc[r_idx][col_name], fail_fmt)
                except ValueError:
                    pass

        # Write Summary
        summary_row = len(df) + 2
        ws_res.write(summary_row, 0, "Evaluation Summary:", header_fmt)
        ws_res.merge_range(summary_row + 1, 0, summary_row + 10, 5, result.summary, cell_fmt)


def _apply_formatting(worksheet, df, header_fmt, cell_fmt):
    """Helper to apply basic grid styling."""
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num, value, header_fmt)
        worksheet.set_column(col_num, col_num, 15)