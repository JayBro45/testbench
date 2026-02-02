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

    def validate(self, rows: List[Dict[str, Any]]) -> Any:
        """
        Delegates validation to the AVRAcceptanceEngine.
        """
        engine = AVRAcceptanceEngine(rows)
        return engine.evaluate()

    def generate_reports(self, rows: List[Dict[str, Any]], output_dir: str, prefix: str) -> None:
        """
        Generates the legacy AVR Excel reports.
        
        The rated output voltage is fetched directly from the Acceptance Engine class
        to ensure a single source of truth.
        """
        eng_path = os.path.join(output_dir, f"{prefix}_AVR_RESULT.xlsx")
        sub_path = os.path.join(output_dir, f"{prefix}_AVR_SUBMISSION.xlsx")
        
        # Use the hardcoded constant from the engine
        rated_volt = AVRAcceptanceEngine.RATED_OUTPUT_VOLTAGE
        
        generate_avr_excel_report(rows, eng_path, rated_volt)
        generate_avr_submission_excel(rows, sub_path)