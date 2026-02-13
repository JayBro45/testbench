"""
avr_acceptance_engine.py
========================

AVR Acceptance Engine (Legacy-Equivalent)

This module implements the **exact acceptance logic** used in the legacy
`evaluate_AVR` system, rewritten in a structured and testable form.

⚠️ CRITICAL GUARANTEE
--------------------
- Logic is intentionally NOT optimized
- Thresholds are NOT configurable
- Order of checks is preserved
- Behavior matches legacy Excel-based evaluation

Scope & Assumptions
-------------------
- Applicable ONLY to Unidirectional AVR systems
- At least 3 test rows are required (row 3 used for rated power/current)
- Grid data is assumed to be validated and numeric where required
- Row numbering in results is Excel-style (1-based, header at row 1)

This engine performs **pass / fail classification** and identifies:
- Invalid (FAIL) cells
- Abnormal (PASS but flagged) cells
"""

from dataclasses import dataclass
from typing import List, Dict, Tuple


# =============================================================================
# Legacy Constants (DO NOT MODIFY)
# =============================================================================

RATED_OUTPUT_VOLTAGE = 230.0

# Output voltage THD limit (%)
UNIDIR_VTHD_LIMIT = 8.0

# Efficiency thresholds (%)
EFFICIENCY_MIN = 85.0
EFFICIENCY_ABNORMAL = 96.0

# Tolerances (%)
AC_INPUT_VOLT_TOL = 5       # Input voltage tolerance
AC_OUTPUT_VOLT_TOL = 1     # Output voltage tolerance
LOAD_CURRENT_TOL = 10      # Load current tolerance

# Minimum rows for acceptance evaluation (row 2 used for rated power/current)
MIN_ROWS = 3


# =============================================================================
# Result Container
# =============================================================================

@dataclass
class AVRResult:
    """
    Container for final AVR acceptance evaluation.

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


# =============================================================================
# Acceptance Engine
# =============================================================================

class AVRAcceptanceEngine:
    """
    Legacy-equivalent AVR acceptance evaluation engine.

    This class evaluates a 6-row AVR test grid and applies
    all legacy acceptance checks in the original order.

    Notes
    -----
    - All numeric comparisons intentionally mirror legacy behavior
    - No early exits are performed
    - All checks are executed regardless of prior failures
    """

    def __init__(self, grid_rows: List[Dict[str, float | str]]):
        """
        Initialize the acceptance engine.

        Parameters
        ----------
        grid_rows : List[Dict[str, float | str]]
            Parsed grid rows from the UI/export layer.

        Raises
        ------
        ValueError
            If the number of rows is less than 3.
        """
        if len(grid_rows) < MIN_ROWS:
            raise ValueError("AVR evaluation requires at least 3 rows")

        self.rows = grid_rows
        self.rated_voltage = RATED_OUTPUT_VOLTAGE

        # Collected failures and abnormal flags
        self.invalid: Dict[str, Tuple[str, ...]] = {}
        self.abnormal: Dict[str, Tuple[str, ...]] = {}

        # ---------------------------------------------------------------------
        # Derived Rating Calculations
        # ---------------------------------------------------------------------
        # Rated power is derived from the 3rd row (index 2).
        # MATH: kW(out) is converted to Watts (x 1000).
        self.rated_power = abs(float(self.rows[2]["kW (out)"])) * 1000
        
        # Rated Load Current Calculation:
        # MATH: I_rated = P_rated / V_nominal (230V)
        self.rated_load_current = round(
            self.rated_power / self.rated_voltage, 2
        )

    # -------------------------------------------------------------------------
    # Utility Helpers (Legacy-Compatible)
    # -------------------------------------------------------------------------

    def _is_in_tolerance(self, measured, expected, tol_percent):
        """
        Check if a value is within ± tolerance percent.

        Math
        ----
        Lower Limit = expected * (1 - tol_percent / 100)
        Upper Limit = expected * (1 + tol_percent / 100)
        Pass if: Lower Limit <= measured <= Upper Limit
        """
        tol = tol_percent / 100.0
        return expected * (1 - tol) <= measured <= expected * (1 + tol)

    def _row_ids(self, flags):
        """
        Convert boolean flags into Excel-style row numbers.

        Logic
        -----
        Returns a tuple of strings representing row numbers where 'flags' is True.
        Row numbering starts at Excel row 2 (row 1 is the header).
        Example: Index 0 -> Row "2", Index 2 -> Row "4".
        """
        return tuple(str(i + 2) for i, f in enumerate(flags) if f)

    # -------------------------------------------------------------------------
    # Individual Legacy Checks
    # -------------------------------------------------------------------------

    def check_output_vthd(self):
        """
        Validate output voltage Total Harmonic Distortion (THD).

        Rule (Math)
        -----------
        FAIL if: VTHD (out) >= 8.0 %
        
        Why
        ---
        Ensures the output waveform quality meets the unidirectional AVR standard.
        """
        flags = [
            float(r["VTHD (out)"]) >= UNIDIR_VTHD_LIMIT
            for r in self.rows
        ]
        self.invalid["VTHD (out)"] = self._row_ids(flags)

    def check_efficiency(self):
        """
        Validate system efficiency.

        Rules (Math)
        ------------
        1. **FAIL Condition:**
           - Context: Only applies when operating at **Full Load** (I_out within ±10% of rated).
           - Rule: Efficiency < 85.0 %
        
        2. **ABNORMAL Condition (Warning):**
           - Context: Applies to any row.
           - Rule: Efficiency > 96.0 % (Suspiciously high, likely measurement error).

        Why
        ---
        Ensures the AVR operates within expected thermal and electrical loss limits.
        """
        invalid_flags = []
        abnormal_flags = []

        for r in self.rows:
            i_out = float(r["I (out)"])
            eff = float(r["Efficiency"])

            # Check if this row represents a Full Load test
            is_full_load = self._is_in_tolerance(
                i_out, self.rated_load_current, LOAD_CURRENT_TOL
            )

            # Apply logic
            invalid_flags.append(is_full_load and eff < EFFICIENCY_MIN)
            abnormal_flags.append(eff > EFFICIENCY_ABNORMAL)

        self.invalid["Efficiency"] = self._row_ids(invalid_flags)
        self.abnormal["Efficiency"] = self._row_ids(abnormal_flags)

    def check_output_voltage(self):
        """
        Validate output voltage regulation under various conditions.

        Rules (Math)
        ------------
        1. **General Safety Limits (Global):**
           - FAIL if V_out < 220.8 V (230V - 4%)
           - FAIL if V_out > 239.2 V (230V + 4%)

        2. **Nominal Input Check:**
           - Context: When Input Voltage is Nominal (230V ± 5%).
           - FAIL if V_out is not within 230V ± 1%.

        3. **Full Load Check:**
           - Context: When Output Current is Rated (± 10%).
           - FAIL if V_out is not within 230V ± 1%.

        Why
        ---
        Verifies the AVR's primary function: maintaining a stable 230V output.
        """
        undervolt = []
        overvolt = []
        invalid_vin_230 = []
        invalid_full_load = []

        for r in self.rows:
            vout = float(r["V (out)"])
            vin = float(r["V (in)"])
            iout = float(r["I (out)"])

            # 1. General ±4% limits
            undervolt.append(vout < self.rated_voltage * 0.96)
            overvolt.append(vout > self.rated_voltage * 1.04)

            # 2. ±1% when VIN ≈ 230 V
            if self._is_in_tolerance(vin, 230, AC_INPUT_VOLT_TOL):
                invalid_vin_230.append(
                    not self._is_in_tolerance(
                        vout, self.rated_voltage, AC_OUTPUT_VOLT_TOL
                    )
                )
            else:
                invalid_vin_230.append(False)

            # 3. ±1% at full load
            if self._is_in_tolerance(
                iout, self.rated_load_current, LOAD_CURRENT_TOL
            ):
                invalid_full_load.append(
                    not self._is_in_tolerance(
                        vout, self.rated_voltage, AC_OUTPUT_VOLT_TOL
                    )
                )
            else:
                invalid_full_load.append(False)

        combined = set(
            self._row_ids(undervolt)
            + self._row_ids(overvolt)
            + self._row_ids(invalid_vin_230)
            + self._row_ids(invalid_full_load)
        )

        self.invalid["V (out)"] = tuple(sorted(combined, key=int))

    def check_no_load(self):
        """
        Validate No-Load Input Power and No-Load Input Current.

        Context
        -------
        This check applies ONLY when Output Current (I_out) is effectively 0.

        Calculated Values (Math)
        ------------------------
        1. **Rated Input Current (Hardcoded Override):**
           - Instead of calculating P/V, we explicitly fetch the 'I (in)' value
             from the **3rd test row** (Index 2). This represents the nominal
             current reference for this specific unit.
           - `rated_input_current = float(self.rows[2]["I (in)"])`

        Rules (Math)
        ------------
        1. **No-Load Power Check:**
           - FAIL if: kW (in) > (10% of Rated Power)
           - Formula: `kwin > (0.1 * rated_power_watts / 1000)`

        2. **No-Load Current Check:**
           - FAIL if: I (in) > (25% of Rated Input Current)
           - Formula: `iin > 0.25 * rated_input_current`

        Why
        ---
        Ensures the device does not consume excessive power or current when idling.
        """
        no_load_power = []
        no_load_current = []

        # --- HARDCODED REFERENCE ---
        # Fetch Rated Input Current directly from Row 3 (Index 2) "I (in)" column.
        # This is a specific requirement to use the measured nominal input current
        # as the baseline for no-load comparisons.
        rated_input_current = float(self.rows[2]["I (in)"])

        for r in self.rows:
            iout = float(r["I (out)"])
            kwin = float(r["kW (in)"])
            iin = float(r["I (in)"])

            # Identify "No Load" condition (Integer 0 check handles small noise)
            is_no_load = int(iout) == 0

            # Rule 1: Power Limit (> 10% rated)
            no_load_power.append(
                is_no_load and kwin > (0.1 * self.rated_power / 1000)
            )

            # Rule 2: Current Limit (> 25% rated)
            no_load_current.append(
                is_no_load and iin > 0.25 * rated_input_current
            )

        self.invalid["kW (in)"] = self._row_ids(no_load_power)
        self.invalid["I (out)"] = self._row_ids(no_load_current)

    def check_regulation(self):
        """
        Validate Load Regulation and Line Regulation percentages.

        Rules (Math)
        ------------
        1. **Load Regulation:**
           - Skipped if value is "--".
           - **Case A (Low Voltage):** If Vin ≈ 160V (±5%), limit is **4%**.
             - FAIL if: abs(Load Reg) > 4.0
           - **Case B (Standard):** Otherwise, limit is **1%**.
             - FAIL if: abs(Load Reg) > 1.0

        2. **Line Regulation:**
           - Skipped if value is "--".
           - Limit is always **1%**.
           - FAIL if: abs(Line Reg) > 1.0

        Why
        ---
        Ensures the AVR stabilizes output despite changes in Load or Input Line voltage.
        """
        load_flags = []
        line_flags = []

        for r in self.rows:
            vin = float(r["V (in)"])

            # 1. Load Regulation Check
            if r["Load"] != "--":
                load_reg = abs(float(r["Load"]))

                # Special relaxation at low voltage (160V)
                if self._is_in_tolerance(vin, 160, AC_INPUT_VOLT_TOL):
                    load_flags.append(load_reg > 4)
                else:
                    load_flags.append(load_reg > 1)
            else:
                load_flags.append(False)

            # 2. Line Regulation Check
            if r["Line"] != "--":
                line_flags.append(abs(float(r["Line"])) > 1)
            else:
                line_flags.append(False)

        self.invalid["Load"] = self._row_ids(load_flags)
        self.invalid["Line"] = self._row_ids(line_flags)

    # -------------------------------------------------------------------------
    # Master Evaluation (Legacy Order)
    # -------------------------------------------------------------------------

    def evaluate(self) -> AVRResult:
        """
        Execute all acceptance checks and produce final result.

        Returns
        -------
        AVRResult
            Final evaluation result with detailed summary.
        """
        self.check_output_vthd()
        self.check_efficiency()
        self.check_output_voltage()
        self.check_no_load()
        self.check_regulation()

        summary_lines = []

        for k, rows in self.invalid.items():
            if rows:
                summary_lines.append(
                    f"Invalid {k} in rows: {', '.join(rows)}"
                )

        for k, rows in self.abnormal.items():
            if rows:
                summary_lines.append(
                    f"Abnormal {k} in rows: {', '.join(rows)}"
                )

        passed = not any(self.invalid.values())
        summary_lines.append(f"RESULT: {'PASS' if passed else 'FAIL'}")
        summary_lines.append(
            "LEGEND: [RED: FAIL, YELLOW: PASS BUT ABNORMAL, GREEN: PASS]"
        )

        return AVRResult(
            passed=passed,
            summary="\n".join(summary_lines),
            invalid_cells=self.invalid,
            abnormal_cells=self.abnormal,
        )