"""
smr_acceptance_engine.py
========================

Acceptance Engine for Switched Mode Rectifiers (SMR).

This module implements the RDSO-compliant validation logic for:
1. SMR SMPS (110V)
2. SMR Telecom RE (48V, Railway Electrified)
3. SMR Telecom Non-RE (48V, Non-Railway Electrified)

The module automatically detects the SMR type based on:
- Output Voltage (detects SMPS vs Telecom)
- Input Voltage at Row 1 (detects RE vs Non-RE)

All acceptance criteria are HARDCODED and cannot be modified via configuration.
"""

import pandas as pd
import logging
from typing import List, Dict, Tuple
from dataclasses import dataclass

logger = logging.getLogger(__name__)

@dataclass
class SMRResult:
    """
    Standardized result container for SMR tests.
    """
    passed: bool
    summary: str
    invalid_cells: Dict[str, Tuple[str, ...]]
    abnormal_cells: Dict[str, Tuple[str, ...]]


class SMRAcceptanceEngine:
    """
    Validation engine for SMR test data.
    """

    def __init__(self, grid_rows: List[Dict[str, float | str]]):
        """
        Initialize the engine with raw grid data.
        
        :param grid_rows: List of dictionaries from the UI grid.
        """
        self.df = pd.DataFrame(grid_rows)
        
        # Coerce numeric columns for analysis
        cols = [
            'V (in)', 'I (in)', 'V (out)', 'I (out)', 
            'Vthd % (in)', 'Ithd % (in)', 'PF (in)', 
            'Efficiency', 'Ripple (out)'
        ]
        for c in cols:
            if c in self.df.columns:
                self.df[c] = pd.to_numeric(self.df[c], errors='coerce')

        # Logic State
        self.module_type = "SMR_SMPS"
        self.invalid_dict = {}
        self.abnormal_dict = {
            "Vac_SMR_SMPS": {}, "Vac_SMR_Telecom_RE": {}, 
            "Vac_SMR_Telecom_Non-RE": {}, "Efficiency": []
        }
        
        # --- HARDCODED RDSO CRITERIA ---
        self.SMR_differentiator_volt_threshold = 100
        
        # Reference Values
        self.SMR_SMPS_criteria = {"rated_load_current": 20}
        self.telecom_SMR_criteria = {"rated_load_current": 25}

        # Voltage Limits (Float/Boost/Under/Over)
        self.smps_battery_voltages = {
            "LM": {"over": 138.16, "under": 101.09, "float": 118.25, "boost": 133.5},
            "VRLA": {"over": 131.01, "under": 98.34, "float": 123.8, "boost": 126.5}
        }
        self.telecom_voltages = {
            "LM": {"over": 56.0, "under": 44.4, "float": 52.8, "boost": 64.8},
            "VRLA": {"over": 56.0, "under": 44.4, "float": 54.0, "boost": 55.2}
        }

        # Universal Limits
        self.input_current_percent_thd_limit = 10.0
        self.ripple_P2P_limit = 300.0  # mV
        self.abnormal_efficiency_limit = 96.0
        
        # Specific Limits (SMPS)
        self.smps_input_volt_thd_limit = 8.0
        self.smps_pf_limit = 0.90
        self.smps_pf_limit_230V = 0.95
        self.smps_eff_limit_275V_20A = 90.0
        self.smps_eff_limit_general = 85.0

        # Specific Limits (Telecom)
        self.telecom_input_volt_thd_limit = 10.0
        self.telecom_eff_limit_IFL_nominal = 85.0
        self.telecom_eff_limit_general = 80.0
        self.telecom_lead_pf_nominal = 0.98
        self.telecom_lag_pf_nominal = 0.95
        self.telecom_pf_general = 0.90

    # -------------------------------------------------------------------------
    # Helper Methods
    # -------------------------------------------------------------------------
    
    def _get_row_ids(self, boolean_series) -> Tuple[str, ...]:
        """Converts boolean series True values to Excel-style row numbers (Index+2)."""
        return tuple(str(i + 2) for i in boolean_series.index[boolean_series])

    def _is_in_tolerance(self, val, target, tol):
        """Absolute tolerance check."""
        return (target - tol) <= val <= (target + tol)

    def _detect_type(self):
        """Auto-detects SMR type based on output voltage and input voltage range."""
        out_mean = self.df['V (out)'].mean()
        
        if out_mean > self.SMR_differentiator_volt_threshold:
             self.module_type = "SMR_SMPS"
        else:
            # Differentiate RE vs Non-RE based on Row 1 (Index 0) Vin
            # RE starts ~165V, Non-RE starts ~90V
            first_row_vin = self.df.iloc[0]['V (in)'] if not self.df.empty else 0
            
            if self._is_in_tolerance(first_row_vin, 165, 16.5): # 10% tol
                self.module_type = "SMR_Telecom_RE"
            elif self._is_in_tolerance(first_row_vin, 90, 9.0): # 10% tol
                self.module_type = "SMR_Telecom_Non-RE"
            else:
                self.module_type = "SMR_Telecom_RE" # Default fallback

        logger.info(f"SMR Analysis Type Detected: {self.module_type}")

    # -------------------------------------------------------------------------
    # Core Evaluation
    # -------------------------------------------------------------------------

    def evaluate(self) -> SMRResult:
        """Main execution method."""
        if self.df.empty:
             return SMRResult(False, "No Data", {}, {})

        self._detect_type()
        
        # 1. Check Power Factor
        self._check_power_factor()
        
        # 2. Check Efficiency
        self._check_efficiency()
        
        # 3. Check THD (Current & Voltage)
        self._check_thd()
        
        # 4. Check Output Voltage (Regulation)
        self._check_output_voltage()
        
        # 5. Check Ripple
        ripple_fail = self.df['Ripple (out)'].gt(self.ripple_P2P_limit)
        self.invalid_dict['Ripple (out)'] = self._get_row_ids(ripple_fail)

        # 6. Abnormal Checks
        self._check_abnormal()

        summary = self._generate_summary()
        passed = not any(self.invalid_dict.values())

        return SMRResult(
            passed=passed,
            summary=summary,
            invalid_cells=self.invalid_dict,
            abnormal_cells=self.abnormal_dict
        )

    # -------------------------------------------------------------------------
    # Check Implementations
    # -------------------------------------------------------------------------

    def _check_power_factor(self):
        col = 'PF (in)'
        if self.module_type == "SMR_SMPS":
            # SMPS Logic
            is_230 = self.df['V (in)'].map(lambda x: self._is_in_tolerance(x, 230, 5))
            
            fail_230 = is_230 & self.df[col].abs().lt(self.smps_pf_limit_230V)
            fail_other = (~is_230) & self.df[col].abs().lt(self.smps_pf_limit)
            
            self.invalid_dict[col] = self._get_row_ids(fail_230 | fail_other)
            
        else:
            # Telecom Logic
            is_nom = self.df['V (in)'].map(lambda x: self._is_in_tolerance(x, 230, 5))
            rated_i = self.telecom_SMR_criteria["rated_load_current"]
            is_high_load = self.df['I (out)'] >= (0.75 * rated_i)
            
            # Nominal Voltage & High Load
            nom_subset = is_nom & is_high_load
            lead_fail = nom_subset & (self.df[col] >= 0) & (self.df[col] < self.telecom_lead_pf_nominal)
            lag_fail = nom_subset & (self.df[col] < 0) & (self.df[col].abs() < self.telecom_lag_pf_nominal)
            
            # General Case
            gen_subset = ~(is_nom & is_high_load)
            gen_fail = gen_subset & (self.df[col].abs() < self.telecom_pf_general)
            
            self.invalid_dict[col] = self._get_row_ids(lead_fail | lag_fail | gen_fail)

    def _check_efficiency(self):
        col = 'Efficiency'
        if self.module_type == "SMR_SMPS":
            # SMPS Logic: Special check at 275V/20A
            is_275 = self.df['V (in)'].map(lambda x: self._is_in_tolerance(x, 275, 5))
            is_20A = self.df['I (out)'].map(lambda x: self._is_in_tolerance(x, 20, 1))
            
            special_case = is_275 & is_20A
            fail_special = special_case & (self.df[col] < self.smps_eff_limit_275V_20A)
            fail_general = (~special_case) & (self.df[col] < self.smps_eff_limit_general)
            
            self.invalid_dict[col] = self._get_row_ids(fail_special | fail_general)
        else:
            # Telecom Logic
            is_nom = self.df['V (in)'].map(lambda x: self._is_in_tolerance(x, 230, 5))
            rated_i = self.telecom_SMR_criteria["rated_load_current"]
            is_full = self.df['I (out)'].map(lambda x: self._is_in_tolerance(x, rated_i, 1))
            
            special_case = is_nom & is_full
            fail_special = special_case & (self.df[col] < self.telecom_eff_limit_IFL_nominal)
            fail_general = (~special_case) & (self.df[col] < self.telecom_eff_limit_general)
            
            self.invalid_dict[col] = self._get_row_ids(fail_special | fail_general)

    def _check_thd(self):
        # Current THD
        col_i = 'Ithd % (in)'
        rated_i = self.SMR_SMPS_criteria["rated_load_current"] if "SMPS" in self.module_type else self.telecom_SMR_criteria["rated_load_current"]
        
        # Check applies only from 50% to 100% load
        is_relevant_load = self.df['I (out)'] >= (0.5 * rated_i)
        fail_i = is_relevant_load & (self.df[col_i] >= self.input_current_percent_thd_limit)
        self.invalid_dict[col_i] = self._get_row_ids(fail_i)
        
        # Voltage THD
        col_v = 'Vthd % (in)'
        limit_v = self.smps_input_volt_thd_limit if "SMPS" in self.module_type else self.telecom_input_volt_thd_limit
        fail_v = self.df[col_v] >= limit_v
        self.invalid_dict[col_v] = self._get_row_ids(fail_v)

    def _check_output_voltage(self):
        col = 'V (out)'
        # Determines bounds based on module type
        if "SMPS" in self.module_type:
            # Assuming LM batteries for standard check (simplification based on prompt)
            bounds = self.smps_battery_voltages["LM"]
        else:
            bounds = self.telecom_voltages["LM"]
            
        under = self.df[col] < bounds["under"]
        over = self.df[col] > bounds["over"]
        
        self.invalid_dict[col] = self._get_row_ids(under | over)

    def _check_abnormal(self):
        # Efficiency > 96%
        abn_eff = self.df['Efficiency'] > self.abnormal_efficiency_limit
        self.abnormal_dict['Efficiency'] = self._get_row_ids(abn_eff)

    def _generate_summary(self) -> str:
        lines = [f"Mode: {self.module_type}"]
        for k, rows in self.invalid_dict.items():
            if rows: lines.append(f"Invalid {k} in rows: {', '.join(rows)}")
        for k, rows in self.abnormal_dict.items():
            if rows and isinstance(rows, tuple): # Simple list check
                 lines.append(f"Abnormal {k} in rows: {', '.join(rows)}")
        
        passed = not any(self.invalid_dict.values())
        lines.append(f"RESULT: {'PASS' if passed else 'FAIL'}")
        return "\n".join(lines)