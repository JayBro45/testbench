# AVR Acceptance Engine – Checks for Testing Engineers

This document lists **all checks** performed by the AVR Acceptance Engine (`avr_acceptance_engine.py`). The logic is **legacy-equivalent** (matches the original Excel-based evaluation). Thresholds are **fixed** and not user-configurable.

---

## Scope and assumptions

- **Applies to:** Unidirectional AVR systems only.
- **Test grid:** Exactly **6 rows** are required (rows 2–7 in Excel; row 1 = header).
- **Rated output voltage:** **230 V** (fixed).
- **Rated values (derived from data):**
  - **Rated power (W):** From **row 3** (Excel row 4): `|kW (out)| × 1000`.
  - **Rated load current (A):** `Rated power ÷ 230`, rounded to 2 decimals.
  - **Rated input current (A):** From **row 3** (Excel row 4): **I (in)**. Used only for the no-load current check.

All checks run in a fixed order; there are no early exits.

---

## 1. Output voltage THD – VTHD (out)

**Purpose:** Output waveform quality for unidirectional AVR.

| Check        | Fail if                    |
|-------------|----------------------------|
| VTHD (out)  | **VTHD (out) ≥ 8.0 %**     |

- Applied to **every row**.

---

## 2. Efficiency

**Purpose:** Thermal and loss limits; flag suspiciously high values.

| Condition | Rule | Result |
|-----------|------|--------|
| **Full load** (I (out) within **±10%** of rated load current) | Efficiency **< 85.0 %** | **FAIL** |
| Any row | Efficiency **> 96.0 %** | **ABNORMAL** (PASS but flagged, e.g. amber) |

- Full load is determined by: `I (out)` within ±10% of **rated load current** (derived from row 3 as above).

---

## 3. Output voltage – V (out)

**Purpose:** Stable 230 V output under different input and load conditions.

### 3.1 General safety limits (all rows)

| Check   | Fail if |
|---------|--------|
| Under   | **V (out) < 220.8 V** (230 V − 4%) |
| Over    | **V (out) > 239.2 V** (230 V + 4%) |

### 3.2 Nominal input check

| Condition | Fail if |
|-----------|--------|
| **V (in)** within **230 V ± 5%** (218.5–241.5 V) | **V (out)** not within **230 V ± 1%** (227.7–232.3 V) |

### 3.3 Full load check

| Condition | Fail if |
|-----------|--------|
| **I (out)** within **rated load current ± 10%** | **V (out)** not within **230 V ± 1%** (227.7–232.3 V) |

- A row can fail on any of the above; all failures are combined and reported under **V (out)**.

---

## 4. No-load checks

**Purpose:** Limit power and current when the AVR is idling (no load).

**Applies only when:** **I (out) = 0** (no-load row(s)).

### 4.1 No-load input power – kW (in)

| Condition     | Fail if |
|---------------|--------|
| No-load row   | **kW (in) > 10% of rated power** (in kW), i.e. `kW (in) > 0.1 × (rated_power_W / 1000)` |

- Rated power (W) is from **row 3**: `|kW (out)| × 1000`.

### 4.2 No-load input current

| Condition     | Fail if |
|---------------|--------|
| No-load row   | **I (in) > 25% of rated input current** |

- **Rated input current** is the **I (in)** value from **row 3** (Excel row 4).
- In the report, this failure is shown as invalid **I (out)** for that row (no-load current check).

---

## 5. Load and line regulation

**Purpose:** Output stability when load or input voltage changes.

- **Load** and **Line** columns are calculated in the UI (from regulation %). Rows where the value is **"--"** are **not** evaluated.

### 5.1 Load regulation – Load (%)

| Condition | Fail if |
|-----------|--------|
| **V (in)** within **160 V ± 5%** (152–168 V) | **\|Load\| > 4.0 %** |
| **V (in)** not in 160 V ± 5% | **\|Load\| > 1.0 %** |

### 5.2 Line regulation – Line (%)

| Condition | Fail if |
|-----------|--------|
| Any row (where Line ≠ "--") | **\|Line\| > 1.0 %** |

---

## Summary table – limits at a glance

| Check            | Limit / condition | Fail if |
|------------------|-------------------|--------|
| **VTHD (out)**   | —                 | ≥ 8.0 % |
| **Efficiency**   | At full load (I_out ±10% rated) | < 85.0 % |
| **Efficiency**   | Any row           | > 96.0 % → abnormal only |
| **V (out)**      | All rows          | < 220.8 V or > 239.2 V |
| **V (out)**      | When V(in) = 230 V ± 5% | V(out) not 230 V ± 1% |
| **V (out)**      | When I(out) = rated ± 10% | V(out) not 230 V ± 1% |
| **kW (in)**      | No-load only (I(out)=0) | > 10% of rated power (kW) |
| **I (in)** (no-load) | No-load only   | > 25% of row-3 I(in) (reported as I (out)) |
| **Load**         | V(in) ≈ 160 V ± 5% | \|Load\| > 4 % |
| **Load**         | Otherwise         | \|Load\| > 1 % |
| **Line**         | —                 | \|Line\| > 1 % |

---

## Execution order

1. **check_output_vthd** – VTHD (out) ≥ 8 %
2. **check_efficiency** – Efficiency < 85 % at full load; > 96 % → abnormal
3. **check_output_voltage** – V (out) limits and 230 V ± 1 % at nominal / full load
4. **check_no_load** – kW (in) and I (in) at no-load rows
5. **check_regulation** – Load % and Line %

**Overall result:** **PASS** only if there are **no invalid (FAIL) cells** in any check. Abnormal cells are PASS but flagged.

---

## Constants (hardcoded in engine)

| Constant              | Value | Use |
|-----------------------|-------|-----|
| RATED_OUTPUT_VOLTAGE  | 230.0 V | Nominal output |
| UNIDIR_VTHD_LIMIT     | 8.0 %  | VTHD (out) |
| EFFICIENCY_MIN        | 85.0 % | Efficiency fail at full load |
| EFFICIENCY_ABNORMAL   | 96.0 % | Efficiency abnormal flag |
| AC_INPUT_VOLT_TOL     | 5 %    | V (in) tolerance (e.g. 230 ± 5 %, 160 ± 5 %) |
| AC_OUTPUT_VOLT_TOL    | 1 %    | V (out) 230 V ± 1 % |
| LOAD_CURRENT_TOL      | 10 %   | Full load = rated I(out) ± 10 % |
| EXPECTED_ROWS          | 6      | Required grid rows |

All of the above are **not** configurable via UI or config file.
