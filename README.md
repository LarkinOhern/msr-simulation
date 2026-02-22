# CoWorkProjects — MSR Servicing Artifacts
**Portfolio Period:** December 2025 – January 2026
**Portfolio Size:** 1,000 loans (Dec) → 1,188 loans (Jan)
**Total UPB:** ~$291M (Dec) → ~$357M (Jan)
**Last Updated:** February 21, 2026
**Prepared by:** Larkin O'Hern

---

## Folder Contents

### MSR Sample Tape
**File:** `MSR_Sample_Tape_Dec2025_Jan2026.xlsx`

Two-month MSR tape for a ~1,000-loan portfolio. Contains three tabs:

- **Dec 2025** — Full loan-level tape as of 12/31/2025 (1,000 loans)
- **Jan 2026** — Full loan-level tape as of 01/31/2026 (1,188 loans)
- **Portfolio Summary** — Reconciliation bridge with count and UPB tie-outs, DQ migration table, investor composition, and loan-ID-level verification that PIF loans are absent from January and New Add loans are present

Key portfolio parameters:
- Average UPB ~$350K
- Annual CPR: 13% (~1.15% SMM → 12 PIF per month)
- Growth: 200 gross new adds per month
- Delinquency: 30 DPD ~0.5%, 60 DPD ~0.25%, 90+ DPD ~0.1% (cumulative)
- DQ loans accrue interest / do not amortize; small bucket migrations month to month

---

### Recon — New Adds
**File:** `Recon_NewAdds_Jan2026.xlsx`

Details on 200 loans onboarded to the portfolio in January 2026. Includes transfer date, origination details, investor, rate, and full payment structure. Row totals reconcile to the +200 / +$70.6M UPB shown in the Portfolio Summary bridge.

---

### Recon — Paid in Full
**File:** `Recon_PaidInFull_Jan2026.xlsx`

Details on 12 loans that paid off in January 2026 (driven by 13% CPR). Includes payoff date, final UPB, payoff amount, interest due, fees, and payoff reason. Row totals reconcile to the -12 / -$3.6M UPB shown in the Portfolio Summary bridge.

---

### Recon — Capitalization
**File:** `Recon_Capitalization_Dec2025_Jan2026.xlsx`

Loans where amounts were capitalized onto the UPB during the period (deferred interest, escrow advance capitalizations on 60/90+ DPD loans). Includes prior UPB, capitalized amount, new UPB, authorization reference, and effective date.

---

### Recon Report (Auto-Generated)
**Files:** `Recon_Report_Dec_2025_to_Jan_2026.md` and `Recon_Summary_Dec_2025_to_Jan_2026.xlsx`

Output of the recon automation script (see below). Documents all changes between the Dec and Jan tapes: count bridge, UPB bridge, DQ migration, new add details, PIF details, curtailments, and investor mix comparison.

---

## Python Scripts

### `build_msr_tape.py`
Generates the entire portfolio from scratch: 1,000-loan Dec tape, CPR-based PIF selection, DQ assignments with month-to-month migration, scheduled amortization, curtailments, capitalizations, 200 new adds, and all four Excel output files. Re-run this to regenerate the sample data.

```bash
python build_msr_tape.py
```

---

### `recon_automation.py`
Reads any two monthly MSR tape files and automatically reconciles them — no manual Excel work required. Outputs a markdown report and a structured Excel summary covering:

- Loan count bridge (Beginning + New Adds - PIF = Ending, verified)
- UPB bridge (tied to within rounding)
- Delinquency bucket migration with loan-level status change detail
- New add and PIF loan detail
- Curtailment / large paydown detection
- Investor mix comparison

**Usage — single combined tape (two sheets):**
```bash
python recon_automation.py MSR_Sample_Tape_Dec2025_Jan2026.xlsx
```

**Usage — two separate monthly tape files:**
```bash
python recon_automation.py MSR_Tape_Dec2025.xlsx MSR_Tape_Jan2026.xlsx
```

**Usage — auto-discover tapes in a folder:**
```bash
python recon_automation.py --folder ./CoWorkProjects
```

**Usage — specify output directory:**
```bash
python recon_automation.py tape1.xlsx tape2.xlsx --output-dir ./reports
```

The script is sheet-name agnostic and will auto-detect the correct data tabs. To specify sheets manually:
```bash
python recon_automation.py tape.xlsx --sheet-m1 "Dec 2025" --sheet-m2 "Jan 2026"
```

---

## Reconciliation Logic

The core tie-out identity is:

```
Dec Portfolio Count + New Adds - Paid in Full = Jan Portfolio Count
Dec Portfolio UPB
  - Scheduled Amortization (current loans only)
  - Curtailments
  + Capitalizations (DQ loans)
  - PIF UPB Removed
  + New Adds UPB
= Jan Portfolio UPB  ✅
```

Delinquency is tracked at the loan level. DQ loans do not amortize; their UPB increases slightly via capitalization. DQ bucket migration is documented in the Portfolio Summary and in the automation report.

---

## Color Coding (Excel Files)

| Color | Meaning |
|---|---|
| Blue text | Hardcoded input values |
| Black text | Calculated formulas |
| Dark navy header | Column headers / title rows |
| Dark green header | New Adds sections |
| Dark red header | PIF / DQ sections |
| Green status text | Current / performing loans |
| Red status text | Delinquent loans |
| Blue-green total rows | Subtotals and summary rows |

---

## Notes

- All balances in USD; rates are annualized
- Sample data uses realistic but synthetic loan-level details
- CPR = 13% annual; SMM = 1 − (1 − 0.13)^(1/12) ≈ 1.15%
- Amortization follows standard mortgage formula
- DQ loans: no scheduled principal reduction; accrued interest may capitalize
