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
from avr_acceptance_engine import AVRAcceptanceEngine, RATED_OUTPUT_VOLTAGE
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
    
    def create_row_data(self, d: Dict[str, Any], row_index: int = 0) -> List[str]:
        """
        Formats data for the grid. 
        Calculates Regulation based on Row Index (Legacy Logic).
        """
        def safe_float(key, default=0.0):
            val = d.get(key)
            return float(val) if val is not None else default

        # 1. Get Rated Voltage from Config (Don't hardcode 230!)
        rated_voltage = RATED_OUTPUT_VOLTAGE
        vout = safe_float('vout')

        # 2. Calculate Regulation
        load_val, line_val = "--", "--"
        
        reg_calc = abs((vout - rated_voltage) / rated_voltage * 100)
        
        # Legacy Logic: 
        # Rows 2, 3, 6 -> Load Regulation
        # Rows 4, 5    -> Line Regulation
        if row_index in (2, 3, 6):
            load_val = f"{reg_calc:.2f}"
        elif row_index in (4, 5):
            line_val = f"{reg_calc:.2f}"

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
            load_val, # Calculated Live
            line_val  # Calculated Live
        ]

    def validate(self, rows: List[Dict[str, Any]]) -> Any:
        """
        Delegates validation to the AVRAcceptanceEngine.
        """
        engine = AVRAcceptanceEngine(rows)
        return engine.evaluate()

    def generate_reports(self, rows: List[Dict[str, Any]], output_dir: str, prefix: str) -> None:
        """
        Generates AVR reports.
        """
        eng_path = os.path.join(output_dir, f"{prefix}_AVR_RESULT.xlsx")
        sub_path = os.path.join(output_dir, f"{prefix}_AVR_SUBMISSION.xlsx")

        generate_avr_excel_report(rows, eng_path)
        generate_avr_submission_excel(rows, sub_path)