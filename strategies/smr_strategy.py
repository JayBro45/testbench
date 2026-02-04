"""
strategies/smr_strategy.py
==========================

Strategy implementation for SMR (Switched Mode Rectifier) testing.

This module adapts the Test Bench for DC-output devices.
It maps the UI to DC measurements and uses the SMRAcceptanceEngine for validation.

Key Responsibilities
--------------------
- Defining Grid Columns for SMR (including Ripple and PF)
- Mapping Meter DC parameters to UI labels
- Invoking the SMRAcceptanceEngine
"""

import os
from typing import List, Dict, Tuple, Any
from .test_strategy import TestStrategy
from smr_acceptance_engine import SMRAcceptanceEngine
from smr_excel_report import generate_smr_excel_report
from smr_submission_report import generate_smr_submission_excel

class SMRStrategy(TestStrategy):
    """
    Concrete strategy for performing SMR tests.
    
    Supports:
    - SMR SMPS (110V DC)
    - SMR Telecom (48V DC)
    """

    @property
    def name(self) -> str:
        """Returns the display name for the UI selector."""
        return "SMR Test"

    @property
    def grid_headers(self) -> List[str]:
        """
        Specific column order requested for SMR reports.
        Includes Ripple and DC output columns.
        """
        return [
            "V (in)", "I (in)", "P (in)", "PF (in)", 
            "Vthd % (in)", "Ithd % (in)", 
            "V (out)", "I (out)", "P (out)", 
            "Ripple (out)", "Efficiency"
        ]

    @property
    def live_readings_map(self) -> Dict[str, Tuple[str, str]]:
        """
        Maps generic meter keys to SMR-specific labels.
        
        Key Differences from AVR:
        - vout -> V (out) DC
        - pf -> Power Factor (No Unit)
        - ripple -> Added
        """
        return {
            # --- Input Readings ---
            "vin": ("V (in)", "V"),
            "iin": ("I (in)", "A"),
            "pin": ("P (in)", "W"),
            "pf": ("PF", ""), 
            "vthd_in": ("V THD (in)", "%"),
            "ithd_in": ("I THD (in)", "%"),

            # --- Output Readings ---
            "vout": ("V (out) DC", "V"),
            "iout": ("I (out) DC", "A"),
            "pout": ("P (out)", "W"),
            "ripple": ("Ripple", "mV"),
            "efficiency": ("Efficiency", "%")
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
            f"{safe_float('vin'):.1f}",
            f"{safe_float('iin'):.2f}",
            f"{safe_float('pin'):.2f}",
            f"{safe_float('pf'):.2f}", 
            f"{safe_float('vthd_in'):.1f}", 
            f"{safe_float('ithd_in'):.1f}",
            f"{safe_float('vout'):.2f}", 
            f"{safe_float('iout'):.2f}",
            f"{abs(safe_float('pout')):.2f}",
            f"{safe_float('ripple'):.1f}",
            f"{safe_float('efficiency'):.2f}"
        ]

    def validate(self, rows: List[Dict[str, Any]]) -> Any:
        """
        Runs the SMR-specific acceptance engine.
        """
        engine = SMRAcceptanceEngine(rows)
        return engine.evaluate()

    def generate_reports(self, rows: List[Dict[str, Any]], output_dir: str, prefix: str) -> None:
        """
        Generates SMR-specific Excel reports.
        1. Result Report (Engineering/Validation)
        2. Submission Report (Clean)
        """
        res_path = os.path.join(output_dir, f"{prefix}_SMR_RESULT.xlsx")
        sub_path = os.path.join(output_dir, f"{prefix}_SMR_SUBMISSION.xlsx")
        
        generate_smr_excel_report(rows, res_path)
        generate_smr_submission_excel(rows, sub_path)