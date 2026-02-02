"""
strategies/avr_strategy.py
==========================

Strategy implementation for AVR (Automatic Voltage Regulator) testing.

This module acts as the glue between the generic UI and the specific 
AVR acceptance logic and reporting tools. It defines the specific 
columns, hardware mappings, and validation rules for AVRs.
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
    
    Configuration:
    - Rated Output Voltage is HARDCODED to 230.0V (Industrial Standard)
    - Output is AC (Single Phase)
    """

    # HARDCODED CONSTANT - Decoupled from config.json for security
    RATED_OUTPUT_VOLTAGE = 230.0

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
        Delegates validation to the existing AVRAcceptanceEngine.
        """
        engine = AVRAcceptanceEngine(rows)
        return engine.evaluate()

    def generate_reports(self, rows: List[Dict[str, Any]], output_dir: str, prefix: str) -> None:
        """
        Generates the legacy AVR Excel reports.
        """
        eng_path = os.path.join(output_dir, f"{prefix}_AVR_RESULT.xlsx")
        sub_path = os.path.join(output_dir, f"{prefix}_AVR_SUBMISSION.xlsx")
        
        # Pass the hardcoded rated voltage explicitly to the report generator
        generate_avr_excel_report(rows, eng_path, self.RATED_OUTPUT_VOLTAGE)
        generate_avr_submission_excel(rows, sub_path)