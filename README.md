# MSR Simulation — Portfolio, Reconciliation & Validation
**Portfolio Period:** December 2025 – January 2026
**Portfolio Size:** 1,000 loans (Dec) → 1,188 loans (Jan)
**Total UPB:** ~$291M (Dec) → ~$356M (Jan)
**Last Updated:** February 22, 2026
**Prepared by:** Larkin O'Hern

> **SIMULATED DATA** — All loan information is synthetic and generated for testing purposes only.

---

## Overview

This project simulates a mortgage servicing rights (MSR) portfolio across two months and builds out the core operational workflows around it:

1. **Tape generation** — deterministic, reproducible 1,000-loan portfolio with realistic rate curves, amortization, DQ migration, CPR-based payoffs, and new add boarding
2. **Monthly reconciliation** — automated count and UPB bridge, DQ tracking, curtailment detection, and investor mix reporting
3. **Input validation** — a dirty subservicer tape with realistic injected errors, and a two-layer validator that catches them

---

## Scripts

### `build_msr_tape.py`
Generates the entire portfolio from scratch. Produces 4 Excel output files:
- 1,000-loan Dec 2025 tape and 1,188-loan Jan 2026 tape (with Portfolio Summary bridge)
- Separate recon workbooks for New Adds, Paid in Full, and Capitalizations

Key simulation parameters:
- Annual CPR 13% → SMM ~1.15% → 12 PIF loans/month
- 200 gross new adds per month; recent originations at market rates
- DQ: ~0.5% 30 DPD, ~0.25% 60 DPD, ~0.1% 90+ DPD
- DQ loans capitalize accrued interest; current loans amortize scheduled principal
- Net Servicing Fee: FNMA/FHLMC/Portfolio = 25bps fixed; GNMA = triangular 19–69bps (median 44bps)

```bash
python build_msr_tape.py
```

---

### `recon_automation.py`
Reads any two monthly tape files and automatically reconciles them. No manual Excel work required. Outputs a Markdown report and a structured Excel summary covering:

- Loan count bridge (Beginning + New Adds − PIF = Ending)
- UPB bridge (tied to within rounding)
- DQ bucket migration with loan-level status change detail
- New add and PIF loan detail
- Curtailment / large paydown detection
- Investor mix comparison

**Single combined tape (two sheets):**
```bash
python recon_automation.py MSR_Sample_Tape_Dec2025_Jan2026.xlsx
```

**Two separate monthly files:**
```bash
python recon_automation.py MSR_Tape_Dec2025.xlsx MSR_Tape_Jan2026.xlsx
```

**Specify sheets manually:**
```bash
python recon_automation.py tape.xlsx --sheet-m1 "Dec 2025" --sheet-m2 "Jan 2026"
```

---

### `build_msr_tape_errors.py`
Generates a realistic "subservicer submitted" dirty tape by reading the clean Jan 2026 tape and injecting 22 errors across two categories:

| Category | Error Type | Count |
|---|---|---:|
| Hard Stop | UPB with extra zero (×10) | 3 |
| Hard Stop | UPB = $0 (active loan) | 1 |
| Hard Stop | Loan disappeared (no PIF) | 3 |
| Hard Stop | Duplicate loan ID | 1 |
| Hard Stop | Rate as whole number (6.50 vs 0.065) | 2 |
| Hard Stop | UPB > original balance | 1 |
| Yellow Light | NSF as percent (0.25 vs 0.0025, FNMA) | 2 |
| Yellow Light | NSF as whole bps (44 vs 0.0044, GNMA) | 2 |
| Yellow Light | Status skip (Current → 90+ DPD) | 2 |
| Yellow Light | P&I inflated ~20% | 2 |
| Yellow Light | Next Due Date in the past | 2 |
| Yellow Light | Remaining term unchanged | 1 |

Outputs `MSR_Tape_Jan2026_SUBSERVICER.xlsx` with an "Error Log - Reference" sheet documenting every injected error.

```bash
python build_msr_tape_errors.py
```

---

### `validate_msr_tape.py`
Two-layer validator that compares a subservicer submission against the prior month clean tape.

**Layer 1 — Standalone field-level rules** (applied to every loan in the submission):
- UPB = 0 for an active loan → Hard Stop
- UPB > original balance → Hard Stop
- Rate > 1.0 (whole number) or < 0.5% → Hard Stop
- NSF > 1.0 (whole basis points) → Hard Stop
- Duplicate loan IDs → Hard Stop
- NSF between 0.05–1.0 (possible percent misformat) → Yellow Light
- NSF outside investor-expected range → Yellow Light
- Next Due Date in the past for a Current loan → Yellow Light

**Layer 2 — Cross-period checks vs prior month** (continuing loans only):
- Loan in prior tape, absent from submission, no PIF → Hard Stop (or cleared if confirmed in PIF report)
- Status skipped a bucket (e.g. Current → 90+ DPD) → Yellow Light
- P&I increased more than 10% month-over-month → Yellow Light
- Remaining term did not decrease by 1 → Yellow Light
- Rate changed between months → Yellow Light
- New add in submission not found in New Add recon report → Yellow Light

The validator auto-discovers `Recon_PaidInFull_*.xlsx` and `Recon_NewAdds_*.xlsx` in the same directory. Missing loans confirmed in the PIF report are cleared (informational only); only genuinely unexplained absences become hard stops.

Outputs `Validation_<submission>.xlsx` (Summary, Hard Stops, Yellow Lights, Missing Loans tabs) and a Markdown report. The Missing Loans tab is split into two sections: Unexplained (action required) and PIF-Explained (cleared).

**Default (auto-discovers all files in script directory):**
```bash
python validate_msr_tape.py
```

**Explicit file paths:**
```bash
python validate_msr_tape.py \
  --tape MSR_Sample_Tape_Dec2025_Jan2026.xlsx \
  --submission MSR_Tape_Jan2026_SUBSERVICER.xlsx \
  --pif-report Recon_PaidInFull_Jan2026.xlsx \
  --new-add-report Recon_NewAdds_Jan2026.xlsx
```

---

## Output Files

| File | Description |
|---|---|
| `MSR_Sample_Tape_Dec2025_Jan2026.xlsx` | Clean 2-month tape: Dec 2025 tab, Jan 2026 tab, Portfolio Summary |
| `Recon_NewAdds_Jan2026.xlsx` | 200 new add loans onboarded in January |
| `Recon_PaidInFull_Jan2026.xlsx` | 12 loans paid off in January |
| `Recon_Capitalization_Dec2025_Jan2026.xlsx` | DQ loans with capitalized interest/advances |
| `Recon_Report_Dec_2025_to_Jan_2026.md` | Auto-generated Markdown reconciliation report |
| `Recon_Summary_Dec_2025_to_Jan_2026.xlsx` | Auto-generated Excel reconciliation summary |
| `MSR_Tape_Jan2026_SUBSERVICER.xlsx` | Dirty subservicer submission tape (22 injected errors) |
| `Validation_Jan2026_SUBSERVICER.xlsx` | Validation report — Excel with 4 tabs |
| `Validation_Jan2026_SUBSERVICER.md` | Validation report — Markdown |
| `MSR_Tape_Jan2026_SUBSERVICER_LO_Jitter.xlsx` | Blind test submission — errors manually injected by Larkin O'Hern |
| `Validation_Jan2026_SUBSERVICER_LO_Jitter.xlsx` | Blind test validation report — Excel with 4 tabs |
| `Validation_Jan2026_SUBSERVICER_LO_Jitter.md` | Blind test validation report — Markdown |

---

## Tape Column Layout (16 columns)

| # | Column | Format | Notes |
|---|---|---|---|
| 1 | Loan ID | Text | MSR + 6-digit number |
| 2 | Loan Type | Text | Conventional, FHA, VA, USDA |
| 3 | Purpose | Text | Purchase, Refinance |
| 4 | Investor | Text | FNMA, FHLMC, GNMA, Portfolio |
| 5 | Orig Date | Date | MM/DD/YYYY |
| 6 | Original Bal ($) | Currency | Rounded to $5K buckets |
| 7 | Current UPB ($) | Currency | As of tape date |
| 8 | Rate | Percent | Annualized, 4 decimal places |
| 9 | Net Serv Fee | Percent | FNMA/FHLMC=0.0025; GNMA=0.0019–0.0069 |
| 10 | Rem Term | Integer | Months remaining |
| 11 | Maturity | Date | MM/DD/YYYY |
| 12 | P&I ($) | Currency | Monthly principal + interest payment |
| 13 | Escrow ($) | Currency | Monthly escrow payment |
| 14 | Total Pmt ($) | Currency | P&I + Escrow |
| 15 | Status | Text | Current, 30 DPD, 60 DPD, 90+ DPD |
| 16 | Next Due Date | Date | First unpaid installment date |

---

## Reconciliation Identity

```
Dec Portfolio Count + New Adds - Paid in Full = Jan Portfolio Count

Dec Portfolio UPB
  - Scheduled Amortization  (current loans only)
  - Curtailments
  + Capitalizations          (DQ loans)
  - PIF UPB Removed
  + New Adds UPB
= Jan Portfolio UPB          [ties to within rounding]
```

---

## Validation Run Against Dirty Tape

Running `validate_msr_tape.py` against `MSR_Tape_Jan2026_SUBSERVICER.xlsx` with PIF and New Add recon files:

```
Prior month:    1,000 loans
Submission:     1,186 loans (raw, incl. dups)
Missing loans:  15 total
  PIF-explained:  12  (cleared — confirmed in Recon_PaidInFull_Jan2026.xlsx)
  Unexplained:     3  <-- HARD STOP
HARD STOPS:     13  <-- ACTION REQUIRED
YELLOW LIGHTS:   9  <-- REVIEW REQUIRED
Clean loans:    1,166

Hard stop breakdown:
  [ 3]  Missing Loan (not in PIF report)   ← 3 injected (removed without PIF)
  [ 4]  UPB Exceeds Original Balance       ← 3 ×10 errors + 1 explicit
  [ 2]  NSF Expressed as Whole Basis Points
  [ 2]  Rate Expressed as Whole Number
  [ 1]  Duplicate Loan ID
  [ 1]  UPB = Zero (active loan)

Yellow light breakdown:
  [ 2]  NSF May Be Expressed as Percent
  [ 2]  Next Due Date in Past (Current Loan)
  [ 2]  P&I Inflated vs Prior Month
  [ 2]  Status Bucket Skip
  [ 1]  Remaining Term Did Not Decrease
```

The 12 legitimate PIFs are automatically cleared by cross-referencing `Recon_PaidInFull_Jan2026.xlsx`.
Only the 3 genuinely unexplained missing loans (removed without PIF) become hard stops.

---

## Blind Test — Validator Against Unknown Injections

`MSR_Tape_Jan2026_SUBSERVICER_LO_Jitter.xlsx` was manually constructed by Larkin O'Hern with undisclosed errors to test whether the validator catches issues it was not designed around in advance. The validator had no knowledge of what was injected.

```bash
python validate_msr_tape.py --submission MSR_Tape_Jan2026_SUBSERVICER_LO_Jitter.xlsx
```

```
Prior month:    1,000 loans
Submission:     1,187 loans (raw, incl. dups)
Missing loans:  15 total
  PIF-explained:  12  (cleared)
  Unexplained:     3  <-- HARD STOP
HARD STOPS:     13  <-- ACTION REQUIRED
YELLOW LIGHTS:  17  <-- REVIEW REQUIRED
Clean loans:    1,161
```

**Result: 8 for 8 — all manually injected errors caught.**

| Loan ID | Rule Triggered | Injection |
|---|---|---|
| MSR100005 | NSF May Be Expressed as Percent | NSF = 0.5500 (stacked on existing UPB error) |
| MSR100006 | NSF May Be Expressed as Percent | NSF = 0.2500 (stacked on existing UPB error) |
| MSR100015 | NSF May Be Expressed as Percent | NSF = 0.2500 on a clean loan |
| MSR100028 | Invalid Status Value | Status = `"60+ DPD"` (should be `"60 DPD"`) |
| MSR300198 | Unboarded Loan — not in New Add report | Phantom loan ID not in any recon file |
| MSR100017 | Status Bucket Skip | Current → 90+ DPD |
| MSR100024 | Status Bucket Skip | Current → 90+ DPD |
| MSR100025 | Status Bucket Skip | Current → 90+ DPD |

Notable catches: the `"60+ DPD"` typo (realistic field encoding error) and the phantom new add `MSR300198` (loan appearing in submission with no recon trail) were both flagged correctly with no prior knowledge of either injection.

---

## Color Coding (Excel Files)

| Color | Meaning |
|---|---|
| Dark navy header | Column headers / title rows |
| Dark green header | New Adds sections |
| Dark red header | PIF / DQ / Error sections |
| Amber/orange header | Capitalization / noteworthy sections |
| Yellow banner | Simulated data disclaimer row |
| Green status text | Current / performing loans |
| Red status text | Delinquent loans |
| Light red row | Hard stop flagged loans |
| Light yellow row | Yellow light flagged loans |
| Blue-green total rows | Subtotals and summary rows |

---

## Dependencies

```bash
pip install openpyxl
```

Python 3.8+ required. No other external dependencies.
