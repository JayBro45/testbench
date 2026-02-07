# SMR Acceptance Engine – Checks for Testing Engineers

This document lists **all checks** performed by the SMR Acceptance Engine (`smr_acceptance_engine.py`). The logic follows **RDSO specifications** (RDSO/SPN/TL/23/99 Ver.4 and RDSO-SPN 165). Thresholds are **fixed** and not user-configurable.

---

## 1. Module type detection (automatic)

Before any limits are applied, the engine **detects** the SMR type from the data. No user input is required.

| Step | Condition | Resulting type |
|------|-----------|----------------|
| 1 | Mean of **V (out)** over all rows **> 100 V** | **SMR_SMPS** (110 V system, IPS) |
| 2 | Mean of **V (out)** ≤ 100 V → use **first row V (in)** | — |
| 2a | First row **V (in)** within **165 V ± 10%** (148.5–181.5 V) | **SMR_Telecom_RE** (48 V, Railway Electrified) |
| 2b | First row **V (in)** within **90 V ± 10%** (81–99 V) | **SMR_Telecom_Non-RE** (48 V, Non-RE) |
| 2c | Otherwise | **SMR_Telecom_RE** (default) |

- **Rated currents used in checks:** SMPS = **20 A**, Telecom = **25 A**.

---

## 2. Power factor – PF (in)

**Purpose:** Ensures the SMR does not impose excessive reactive load on the grid (utility connection standards).

### SMR SMPS (110 V)

| Operating point | Condition | Fail if |
|-----------------|-----------|--------|
| **Nominal** | V (in) = **230 V ± 5%** (218.5–241.5 V) **and** I (out) = **20 A ± 1 A** (19–21 A) | **\|PF\| < 0.95** |
| **General** | All other rows | **\|PF\| < 0.90** |

### SMR Telecom (48 V)

| Operating point | Condition | Fail if |
|-----------------|-----------|--------|
| **High load** | V (in) = **230 V ± 5%** **and** I (out) **≥ 75% of 25 A** (≥ 18.75 A) | **Leading PF:** PF < 0.98 **or** **Lagging PF:** \|PF\| < 0.95 |
| **General** | All other rows | **\|PF\| < 0.90** |

---

## 3. Efficiency

**Purpose:** Keeps conversion losses within RDSO limits (thermal and cost).

### SMR SMPS (110 V)

| Operating point | Condition | Fail if |
|-----------------|-----------|--------|
| **Nominal / full load** | V (in) = **230 V ± 5%** **and** I (out) = **20 A ± 1 A** | **Efficiency < 90.0 %** |
| **General** | All other rows | **Efficiency < 85.0 %** |

### SMR Telecom (48 V)

| Operating point | Condition | Fail if |
|-----------------|-----------|--------|
| **Nominal / full load** | V (in) = **230 V ± 5%** **and** I (out) = **25 A ± 1 A** (rated ± 1 A) | **Efficiency < 85.0 %** |
| **General** | All other rows | **Efficiency < 80.0 %** |

---

## 4. Total harmonic distortion (THD)

**Purpose:** Limits harmonic pollution of the mains (protection of other equipment on the same supply).

### Current THD – Ithd % (in)

| Applies when | Fail if |
|--------------|--------|
| **I (out) ≥ 50% of rated current** (SMPS: ≥ 10 A, Telecom: ≥ 12.5 A) | **Ithd % (in) ≥ 10.0 %** |
| Below that load | Not evaluated (no fail) |

### Voltage THD – Vthd % (in)

| Module type | Fail if |
|-------------|--------|
| **SMR SMPS** | **Vthd % (in) ≥ 8.0 %** |
| **SMR Telecom** (RE or Non-RE) | **Vthd % (in) ≥ 10.0 %** |

---

## 5. Output DC voltage – V (out)

**Purpose:** Keeps output within safe charging/discharging limits of the battery (Lead-Acid/VRLA).

### SMR SMPS (110 V system)

| Check | Limit |
|-------|--------|
| Under-voltage | **FAIL if V (out) < 101.09 V** |
| Over-voltage | **FAIL if V (out) > 138.16 V** |

### SMR Telecom (48 V system)

| Check | Limit |
|-------|--------|
| Under-voltage | **FAIL if V (out) < 44.4 V** |
| Over-voltage | **FAIL if V (out) > 66.0 V** |

---

## 6. Output ripple – Ripple (out)

**Purpose:** Limits DC ripple on the output.

| Applies to | Fail if |
|------------|--------|
| All SMR types | **Ripple (out) > 300 mV** |

---

## 7. Abnormal efficiency (flag only, not fail)

**Purpose:** Flags unusually high efficiency for review; does **not** cause FAIL.

| Applies to | Flagged as **abnormal** if |
|------------|-----------------------------|
| All SMR types | **Efficiency > 96.0 %** |

- These rows remain **PASS** but are highlighted (e.g. amber) in the report.

---

## Summary table – limits at a glance

| Check | SMR SMPS (110 V) | SMR Telecom (48 V) |
|-------|-------------------|--------------------|
| **PF (in)** | ≥ 0.95 @ 230 V & 20 A; else ≥ 0.90 | ≥ 0.98 lead / ≥ 0.95 lag @ 230 V & high load; else ≥ 0.90 |
| **Efficiency** | ≥ 90% @ 230 V & 20 A; else ≥ 85% | ≥ 85% @ 230 V & 25 A; else ≥ 80% |
| **Ithd % (in)** | ≥ 10% → FAIL (when I(out) ≥ 10 A) | ≥ 10% → FAIL (when I(out) ≥ 12.5 A) |
| **Vthd % (in)** | ≥ 8% → FAIL | ≥ 10% → FAIL |
| **V (out)** | 101.09 V – 138.16 V | 44.4 V – 66.0 V |
| **Ripple (out)** | > 300 mV → FAIL | > 300 mV → FAIL |
| **Efficiency abnormal** | > 96% → flagged | > 96% → flagged |

---

## Execution order and result

1. **Module type** is detected from V (out) and (for Telecom) first-row V (in).
2. **All of the above checks run** for every row (no early exit).
3. **Overall result:** **PASS** only if there are **no invalid (FAIL) cells** in any check.
4. **Report:** Failing rows are listed by column and Excel row number; abnormal efficiency rows are listed separately.

All thresholds are **hardcoded** in the engine and are **not** configurable via the UI or config file.
