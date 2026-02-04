"""
strategies/avr_strategy.py
==========================

Strategy implementation for AVR (Automatic Voltage Regulator) testing.

This module acts as the bridge between the generic UI and the specific 
AVR acceptance logic and reporting tools.

Key Responsibilities
--------------------
- Defining the Grid Columns for AVR tests
- Mapping Meter AC parameters to UI labels
- Invoking the AVRAcceptanceEngine
- Triggering AVR-specific Excel reports
"""

import os
from typing import List, Dict, Tuple, Any
from .test_strategy import TestStrategy
from avr_acceptance_engine import AVRAcceptanceEngine
from avr_excel_report import generate_avr_excel_report
from avr_submission_report import generate_avr_submission_excel

class AVRStrategy(TestStrategy):
    """
    Concrete strategy for performing AVR tests.
    
    This class configures the test bench for Single Phase AC testing.
    It relies on AVRAcceptanceEngine for all validation rules.
    """

    @property
    def name(self) -> str:
        """Returns the display name for the UI selector."""
        return "AVR Test"

    @property
    def grid_headers(self) -> List[str]:
        """
        Defines the specific columns required for AVR testing.
        """
        return [
            "Frequency", "V (in)", "I (in)", "kW (in)",
            "V (out)", "I (out)", "kW (out)", "VTHD (out)",
            "Efficiency", "Load", "Line"
        ]

    @property
    def live_readings_map(self) -> Dict[str, Tuple[str, str]]:
        """
        Maps generic meter keys to AVR-specific labels.
        AVR output is AC, so we map 'vout' to 'V (out)'.
        """
        return {
            "vin": ("V (in)", "V"),
            "iin": ("I (in)", "A"),
            "kwin": ("1-Ph kW", "kW"),
            "frequency": ("Frequency", "Hz"),
            "vout": ("V (out)", "V"),
            "iout": ("I (out)", "A"),
            "kwout": ("1-Ph kW", "kW"),
            "vthd_out": ("V THD", "%")
        }
    
    def create_row_data(self, d: Dict[str, Any]) -> List[str]:
        """
        Formats data for the grid. 
        Safely handles None values by defaulting them to 0.0.
        """
        def safe_float(key, default=0.0):
            val = d.get(key)
            return float(val) if val is not None else default

        return [
            f"{safe_float('frequency'):.2f}",
            f"{safe_float('vin'):.1f}",
            f"{safe_float('iin'):.2f}",
            f"{safe_float('kwin'):.2f}",
            f"{safe_float('vout'):.1f}",
            f"{safe_float('iout'):.2f}",
            f"{abs(safe_float('kwout')):.2f}",
            f"{safe_float('vthd_out'):.1f}",
            f"{safe_float('efficiency'):.2f}",
            "--", # Placeholder: Calculated during export
            "--"  # Placeholder: Calculated during export
        ]

    def validate(self, rows: List[Dict[str, Any]]) -> Any:
        """
        Delegates validation to the AVRAcceptanceEngine.
        """
        engine = AVRAcceptanceEngine(rows)
        return engine.evaluate()

    def generate_reports(self, rows: List[Dict[str, Any]], output_dir: str, prefix: str) -> None:
        
        # --- POST-PROCESSING: Calculate Regulation ---
        # Logic: Find 'Rated' row (Index 2) and 'No Load' row (Index 5 - arbitrary example)
        # or iterate to find min/max Vin for Line Regulation.
        
        # Simple implementation based on finding the "Rated" baseline (230V out target)
        # This fills the "--" gaps so the Engine can actually validate them.
        
        processed_rows = []
        for r in rows:
            new_r = r.copy()
            try:
                # Basic Load Regulation Calc: abs( (Vout - 230) / 230 * 100 )
                # Real regulation compares V_no_load vs V_full_load, but standard AVR
                # often uses deviations from Nominal. 
                vout = float(r.get("V (out)", 0))
                vin = float(r.get("V (in)", 0))
                
                # Calculate if currently "--"
                if r.get("Load") == "--":
                    load_reg = abs((vout - 230.0) / 230.0 * 100)
                    new_r["Load"] = f"{load_reg:.2f}"
                
                if r.get("Line") == "--":
                    # Line reg is valid usually when Load is fixed and Vin varies.
                    line_reg = abs((vout - 230.0) / 230.0 * 100)
                    new_r["Line"] = f"{line_reg:.2f}"
            except (ValueError, TypeError):
                pass
            processed_rows.append(new_r)

        eng_path = os.path.join(output_dir, f"{prefix}_AVR_RESULT.xlsx")
        sub_path = os.path.join(output_dir, f"{prefix}_AVR_SUBMISSION.xlsx")
        
        # Pass PROCESSED rows to report generators
        generate_avr_excel_report(processed_rows, eng_path)
        generate_avr_submission_excel(processed_rows, sub_path)