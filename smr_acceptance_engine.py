"""
smr_acceptance_engine.py
========================

SMR Acceptance Engine (Legacy-Equivalent, RDSO Standard)

This module implements the **exact acceptance logic** for Switched Mode Rectifiers
based on RDSO specifications (RDSO/SPN/TL/23/99 Ver.4 and RDSO-SPN 165).

⚠️ CRITICAL GUARANTEE
--------------------
- Logic matches RDSO specifications strictly
- Thresholds are HARDCODED and not user-configurable
- Automatic Module Type detection is enforced
- Behavior prioritizes safety and standard compliance over flexibility
- No early exits are performed; all checks run regardless of failures

Scope & Assumptions
-------------------
- Supports SMR SMPS (110V) for IPS applications
- Supports SMR Telecom (48V) for Power Plants (RE & Non-RE variants)
- Grid data is assumed to be coerced to numeric types where possible
- Row numbering in results is Excel-style (1-based, header at row 1)

This engine performs **pass / fail classification** and identifies:
- Invalid (FAIL) cells
- Abnormal (PASS but flagged) cells
"""

import pandas as pd
import logging
from typing import List, Dict, Tuple, cast
from dataclasses import dataclass

logger = logging.getLogger(__name__)

@dataclass
class SMRResult:
    """
    Container for final SMR acceptance evaluation.

    Attributes
    ----------
    passed : bool
        Overall PASS / FAIL result.

    summary : str
        Human-readable multi-line summary.

    invalid_cells : Dict[str, Tuple[str, ...]]
        Mapping of column name → Excel-style row numbers
        that FAILED acceptance criteria.

    abnormal_cells : Dict[str, Tuple[str, ...]]
        Mapping of column name → Excel-style row numbers
        that are PASS but flagged as abnormal.
    """
    passed: bool
    summary: str
    invalid_cells: Dict[str, Tuple[str, ...]]
    abnormal_cells: Dict[str, Tuple[str, ...]]


class SMRAcceptanceEngine:
    """
    RDSO-compliant and Legacy-Equivalent SMR acceptance evaluation engine.

    This class encapsulates the validation rules for DC power supplies.
    It automatically detects the module type and applies the specific
    limits for Efficiency, Power Factor, and Regulation.

    Notes
    -----
    - All numeric comparisons intentionally mirror RDSO standards
    - Type detection is automatic based on voltage characteristics
    - All checks are executed regardless of prior failures
    """

    def __init__(self, grid_rows: List[Dict[str, float | str]]):
        """
        Initialize the SMR acceptance engine.

        Parameters
        ----------
        grid_rows : List[Dict[str, float | str]]
            Parsed grid rows from the UI/export layer.
        """
        self.df = pd.DataFrame(grid_rows)
        
        # Coerce numeric columns for vector analysis
        cols = [
            'V (in)', 'I (in)', 'V (out)', 'I (out)', 
            'Vthd % (in)', 'Ithd % (in)', 'PF (in)', 
            'Efficiency', 'Ripple (out)'
        ]
        for c in cols:
            if c in self.df.columns:
                self.df[c] = pd.to_numeric(self.df[c], errors='coerce')

        # State initialization
        self.module_type = "SMR_SMPS"
        self.invalid_dict = {}
        self.abnormal_dict = {
            "Vac_SMR_SMPS": {}, "Vac_SMR_Telecom_RE": {}, 
            "Vac_SMR_Telecom_Non-RE": {}, "Efficiency": []
        }
        
        # ---------------------------------------------------------------------
        # Hardcoded Legacy-Code Specifications
        # ---------------------------------------------------------------------
        
        # Type Detection Threshold
        self.SMR_differentiator_volt_threshold = 100.0
        
        # Rated Load Currents (A)
        self.RATED_CURRENT_SMPS = 20.0
        self.RATED_CURRENT_TELECOM = 25.0

        # Output Voltage Limits (V)
        self.VOLTAGE_LIMITS = {
            "SMPS": {"under": 101.09, "over": 138.16},  # Based on LM battery spec
            "TELECOM": {"under": 44.4, "over": 66.0}    # Based on LM battery spec
        }

        # Universal Limits
        self.LIMIT_ITHD = 10.0         # %
        self.LIMIT_RIPPLE = 300.0      # mV
        self.LIMIT_EFF_ABNORMAL = 96.0 # %
        
        # SMPS Specific Limits
        self.LIMIT_SMPS_VTHD = 8.0     # %
        self.LIMIT_SMPS_PF = 0.90      # General
        self.LIMIT_SMPS_PF_NOM = 0.95  # @ 230V
        self.LIMIT_SMPS_EFF_GEN = 85.0 # %
        self.LIMIT_SMPS_EFF_HIGH = 90.0# % (@ 275V/20A) TO DO: 230V/20A

        # Telecom Specific Limits
        self.LIMIT_TEL_VTHD = 10.0     # %
        self.LIMIT_TEL_EFF_GEN = 80.0  # %
        self.LIMIT_TEL_EFF_HIGH = 85.0 # % (@ Nom/Full Load)
        self.LIMIT_TEL_PF_LEAD = 0.98
        self.LIMIT_TEL_PF_LAG = 0.95
        self.LIMIT_TEL_PF_GEN = 0.90

    # -------------------------------------------------------------------------
    # Utility Helpers
    # -------------------------------------------------------------------------

    def _get_row_ids(self, boolean_series) -> Tuple[str, ...]:
        """
        Convert boolean Series to Excel-style row numbers (Index + 2).
        
        Logic
        -----
        Returns a tuple of strings representing row numbers where series is True.
        Row numbering starts at Excel row 2 (row 1 is the header).
        """
        return tuple(str(i + 2) for i in boolean_series.index[boolean_series])

    def _is_in_tolerance(self, val, target, tol):
        """Check if value is within absolute tolerance (target ± tol)."""
        return (target - tol) <= val <= (target + tol)

    def _detect_type(self):
        """
        Detect the SMR Module Type automatically.

        Logic
        -----
        1. **Output Voltage Check:**
           - Mean V(out) > 100V -> **SMPS** (110V System)
           - Mean V(out) <= 100V -> **Telecom** (48V System)
        
        2. **Input Voltage Check (Telecom Only):**
           - Row 1 V(in) ~ 165V (±10%) -> **RE** (Railway Electrified)
           - Row 1 V(in) ~ 90V  (±10%) -> **Non-RE**
           - Default Fallback -> **RE**
        
        Why
        ---
        Ensures the correct validation limits (RDSO/SPN/TL/23/99 vs RDSO-SPN 165) 
        are applied without requiring user configuration.
        """
        if self.df.empty:
            return # Default or exit

        out_mean_raw = self.df['V (out)'].mean()
        out_mean = cast(float, out_mean_raw)
        # Handle NaN case if column exists but is empty/null
        if pd.isna(out_mean):
            return 
            
        if out_mean > self.SMR_differentiator_volt_threshold:
             self.module_type = "SMR_SMPS"
        else:
            first_row_vin = self.df.iloc[0]['V (in)'] if not self.df.empty else 0
            
            if self._is_in_tolerance(first_row_vin, 165, 165*0.1): # 10% tol
                self.module_type = "SMR_Telecom_RE"
            elif self._is_in_tolerance(first_row_vin, 90, 90*0.1): # 10% tol
                self.module_type = "SMR_Telecom_Non-RE"
            else:
                self.module_type = "SMR_Telecom_RE" # Default fallback

        logger.info(f"SMR Analysis Type Detected: {self.module_type}")

    # -------------------------------------------------------------------------
    # Validation Logic
    # -------------------------------------------------------------------------

    def check_power_factor(self):
        """
        Validate Power Factor (PF).

        Rules (Math)
        ------------
        **1. SMR SMPS (110V):**
           - **Nominal Case:** If Vin is 230V (±5%) AND I_out is 100% of Rated (±1A):
             - FAIL if: abs(PF) < 0.95
           - **General Case:** Otherwise:
             - FAIL if: abs(PF) < 0.90

        **2. SMR Telecom (48V):**
           - **High Load Case:** If Vin is 230V (±5%) AND I_out >= 75% Rated:
             - Leading PF: FAIL if PF < 0.98
             - Lagging PF: FAIL if abs(PF) < 0.95
           - **General Case:** Otherwise:
             - FAIL if: abs(PF) < 0.90

        Why
        ---
        Ensures the SMR does not introduce excessive reactive load to the grid,
        complying with utility connection standards.
        """
        col = 'PF (in)'
        if self.module_type == "SMR_SMPS":
            # Nominal case: 230 V (±5%) and full-load output current (Rated ±1A)
            is_230 = self.df['V (in)'].map(
                lambda x: self._is_in_tolerance(x, 230, 230 * 0.05)  # 5% tolerance
            )
            is_full_load = self.df['I (out)'].map(
                lambda x: self._is_in_tolerance(x, self.RATED_CURRENT_SMPS, 1.0)  # 1A tolerance   #TO DO: CONFIRM WITH IQBAL
            )

            is_nominal_point = is_230 & is_full_load

            fail_230 = is_nominal_point & self.df[col].abs().lt(self.LIMIT_SMPS_PF_NOM)
            fail_other = (~is_nominal_point) & self.df[col].abs().lt(self.LIMIT_SMPS_PF)
            
            self.invalid_dict[col] = self._get_row_ids(fail_230 | fail_other)
        else:
            is_nom = self.df['V (in)'].map(lambda x: self._is_in_tolerance(x, 230, 230*0.05)) # 5% tolerance
            rated_i = self.RATED_CURRENT_TELECOM
            is_high_load = self.df['I (out)'] >= (0.75 * rated_i)
            
            nom_subset = is_nom & is_high_load
            lead_fail = nom_subset & (self.df[col] >= 0) & (self.df[col] < self.LIMIT_TEL_PF_LEAD)
            lag_fail = nom_subset & (self.df[col] < 0) & (self.df[col].abs() < self.LIMIT_TEL_PF_LAG)
            
            gen_subset = ~(is_nom & is_high_load)
            gen_fail = gen_subset & (self.df[col].abs() < self.LIMIT_TEL_PF_GEN)
            
            self.invalid_dict[col] = self._get_row_ids(lead_fail | lag_fail | gen_fail)

    def check_efficiency(self):
        """
        Validate Efficiency.

        Rules (Math)
        ------------
        **1. SMR SMPS (110V):**
           - **High Voltage/Current Case:** If Vin is 230V (±5%) AND I_out is 20A (±1A):
             - FAIL if: Efficiency < 90.0 %
           - **General Case:** Otherwise:
             - FAIL if: Efficiency < 85.0 %

        **2. SMR Telecom (48V):**
           - **Nominal/Full Load Case:** If Vin is 230V (±5%) AND I_out is Rated (±1A):
             - FAIL if: Efficiency < 85.0 %
           - **General Case:** Otherwise:
             - FAIL if: Efficiency < 80.0 %

        Why
        ---
        Ensures energy conversion losses are within acceptable limits defined
        by the RDSO specification for thermal management and cost efficiency.
        """
        col = 'Efficiency'
        if self.module_type == "SMR_SMPS":
            # High-efficiency check at nominal input and full load
            is_230 = self.df['V (in)'].map(
                lambda x: self._is_in_tolerance(x, 230, 230 * 0.05)
            )  # 5% tolerance
            is_20A = self.df['I (out)'].map(lambda x: self._is_in_tolerance(x, 20, 1))      # 1A tolerance
            
            special_case = is_230 & is_20A
            fail_special = special_case & (self.df[col] < self.LIMIT_SMPS_EFF_HIGH)
            fail_general = (~special_case) & (self.df[col] < self.LIMIT_SMPS_EFF_GEN)
            
            self.invalid_dict[col] = self._get_row_ids(fail_special | fail_general)
        else:
            is_nom = self.df['V (in)'].map(lambda x: self._is_in_tolerance(x, 230, 230*0.05)) # 5% tolerance
            rated_i = self.RATED_CURRENT_TELECOM
            is_full = self.df['I (out)'].map(lambda x: self._is_in_tolerance(x, rated_i, 1))  # 1A tolerance
            
            special_case = is_nom & is_full
            fail_special = special_case & (self.df[col] < self.LIMIT_TEL_EFF_HIGH)
            fail_general = (~special_case) & (self.df[col] < self.LIMIT_TEL_EFF_GEN)
            
            self.invalid_dict[col] = self._get_row_ids(fail_special | fail_general)

    def check_thd(self):
        """
        Validate Total Harmonic Distortion (THD) for Current and Voltage.

        Rules (Math)
        ------------
        **1. Current THD (Ithd):**
           - Context: Only applies when Load >= 50% of Rated Current.
           - Rule: FAIL if Ithd (in) >= 10.0 %

        **2. Voltage THD (Vthd):**
           - **SMR SMPS:** FAIL if Vthd (in) >= 8.0 %
           - **SMR Telecom:** FAIL if Vthd (in) >= 10.0 %                   

        Why
        ---
        Prevents harmonic pollution of the mains supply which can degrade 
        other equipment on the same grid.
        """
        # Current THD
        col_i = 'Ithd % (in)'
        rated_i = self.RATED_CURRENT_SMPS if "SMPS" in self.module_type else self.RATED_CURRENT_TELECOM
        
        is_relevant_load = self.df['I (out)'] >= (0.5 * rated_i)
        fail_i = is_relevant_load & (self.df[col_i] >= self.LIMIT_ITHD)
        self.invalid_dict[col_i] = self._get_row_ids(fail_i)
        
        # Voltage THD
        col_v = 'Vthd % (in)'
        limit_v = self.LIMIT_SMPS_VTHD if "SMPS" in self.module_type else self.LIMIT_TEL_VTHD
        fail_v = self.df[col_v] >= limit_v
        self.invalid_dict[col_v] = self._get_row_ids(fail_v)

    def check_output_voltage(self):
        """
        Validate Output DC Voltage.

        Rules (Math)
        ------------
        **1. SMR SMPS (110V System):**
           - FAIL if V_out < 101.09 V
           - FAIL if V_out > 138.16 V

        **2. SMR Telecom (48V System):**
           - FAIL if V_out < 44.4 V
           - FAIL if V_out > 66.0 V

        Why
        ---
        Ensures output voltage stays within the safe charging and discharging
        limits of the connected battery bank (Lead-Acid/VRLA).
        """
        col = 'V (out)'
        if "SMPS" in self.module_type:
            bounds = self.VOLTAGE_LIMITS["SMPS"]
        else:
            bounds = self.VOLTAGE_LIMITS["TELECOM"]
            
        under = self.df[col] < bounds["under"]
        over = self.df[col] > bounds["over"]
        
        self.invalid_dict[col] = self._get_row_ids(under | over)

    def evaluate(self) -> SMRResult:
        """
        Execute all acceptance checks and produce final result.

        Returns
        -------
        SMRResult
            Final evaluation result with detailed summary.
        """
        if self.df.empty:
             return SMRResult(False, "No Data", {}, {})

        self._detect_type()
        self.check_power_factor()
        self.check_efficiency()
        self.check_thd()
        self.check_output_voltage()
        
        # Ripple Check
        ripple_fail = self.df['Ripple (out)'].gt(self.LIMIT_RIPPLE)
        self.invalid_dict['Ripple (out)'] = self._get_row_ids(ripple_fail)

        # Abnormal Efficiency
        abn_eff = self.df['Efficiency'] > self.LIMIT_EFF_ABNORMAL
        self.abnormal_dict['Efficiency'] = self._get_row_ids(abn_eff)

        # Summary Generation
        lines = [f"Mode: {self.module_type}"]
        for k, rows in self.invalid_dict.items():
            if rows: lines.append(f"Invalid {k} in rows: {', '.join(rows)}")
        for k, rows in self.abnormal_dict.items():
             if rows and isinstance(rows, tuple):
                 lines.append(f"Abnormal {k} in rows: {', '.join(rows)}")
        
        passed = not any(self.invalid_dict.values())
        lines.append(f"RESULT: {'PASS' if passed else 'FAIL'}")

        return SMRResult(
            passed=passed,
            summary="\n".join(lines),
            invalid_cells=self.invalid_dict,
            abnormal_cells=self.abnormal_dict
        )