# MSR Tape Validation Report
**Prior Month:** Dec2025_Jan2026  |  **Submitted:** Jan2026 SUBSERVICER
**Generated:** February 21, 2026
**Status:** [FAIL]

> **SIMULATED DATA** — All loan information is synthetic and generated for testing purposes only.

---

## Summary

| Metric | Count |
|---|---:|
| Prior Month Loans | 1,000 |
| Loans in Submission (raw) | 1,186 |
| Duplicate Loan IDs | 1 |
| Missing Loans (no PIF) | 15 |
| New Adds | 200 |
| Unique Loans Evaluated | 1,185 |
| **HARD STOPS** | **25** |
| **YELLOW LIGHTS** | **9** |
| Loans Passing All Checks | 1,166 |

---

## Hard Stops

| Loan ID | Investor | Layer | Rule | Field | Submitted | Expected |
|---|---|---|---|---|---|---|
| MSR100002 | GNMA | Layer 1 | NSF Expressed as Whole Basis Points | Net Serv Fee | 46.0 | Decimal (e.g. 0.0025 for 25bps) |
| MSR100005 | GNMA | Layer 1 | UPB Exceeds Original Balance | Current UPB ($) | $1,577,419.00 | <= $325,000.00 (Orig Bal) |
| MSR100006 | FHLMC | Layer 1 | UPB Exceeds Original Balance | Current UPB ($) | $2,046,985.60 | <= $215,000.00 (Orig Bal) |
| MSR100007 | FHLMC | Layer 1 | UPB Exceeds Original Balance | Current UPB ($) | $3,125,296.60 | <= $325,000.00 (Orig Bal) |
| MSR100010 | GNMA | Layer 1 | NSF Expressed as Whole Basis Points | Net Serv Fee | 28.0 | Decimal (e.g. 0.0025 for 25bps) |
| MSR100051 | FNMA | Layer 1 | UPB = Zero (active loan) | Current UPB ($) | $0.00 | > $0 (not marked Paid in Full) |
| MSR100202 | GNMA | Layer 1 | Rate Expressed as Whole Number | Rate | 5.8300 | Decimal (e.g. 0.0650 for 6.50%) |
| MSR100203 | GNMA | Layer 1 | Rate Expressed as Whole Number | Rate | 4.5200 | Decimal (e.g. 0.0650 for 6.50%) |
| MSR100307 | Portfolio | Layer 1 | UPB Exceeds Original Balance | Current UPB ($) | $347,200.00 | <= $310,000.00 (Orig Bal) |
| MSR100152 | GNMA | Layer 1 | Duplicate Loan ID | Loan ID | Appears 2+ times | Each Loan ID appears exactly once |
| MSR100034 | FHLMC | Layer 2 | Missing Loan (not in PIF report) | — | Not present | Present (no PIF entry found for this loan ID) |
| MSR100100 | FNMA | Layer 2 | Missing Loan (not in PIF report) | — | Not present | Present (no PIF entry found for this loan ID) |
| MSR100102 | FHLMC | Layer 2 | Missing Loan (not in PIF report) | — | Not present | Present (no PIF entry found for this loan ID) |
| MSR100103 | FHLMC | Layer 2 | Missing Loan (not in PIF report) | — | Not present | Present (no PIF entry found for this loan ID) |
| MSR100104 | GNMA | Layer 2 | Missing Loan (not in PIF report) | — | Not present | Present (no PIF entry found for this loan ID) |
| MSR100241 | GNMA | Layer 2 | Missing Loan (not in PIF report) | — | Not present | Present (no PIF entry found for this loan ID) |
| MSR100250 | FNMA | Layer 2 | Missing Loan (not in PIF report) | — | Not present | Present (no PIF entry found for this loan ID) |
| MSR100252 | FNMA | Layer 2 | Missing Loan (not in PIF report) | — | Not present | Present (no PIF entry found for this loan ID) |
| MSR100253 | FHLMC | Layer 2 | Missing Loan (not in PIF report) | — | Not present | Present (no PIF entry found for this loan ID) |
| MSR100264 | GNMA | Layer 2 | Missing Loan (not in PIF report) | — | Not present | Present (no PIF entry found for this loan ID) |
| MSR100625 | FNMA | Layer 2 | Missing Loan (not in PIF report) | — | Not present | Present (no PIF entry found for this loan ID) |
| MSR100709 | FNMA | Layer 2 | Missing Loan (not in PIF report) | — | Not present | Present (no PIF entry found for this loan ID) |
| MSR100911 | FNMA | Layer 2 | Missing Loan (not in PIF report) | — | Not present | Present (no PIF entry found for this loan ID) |
| MSR100913 | FHLMC | Layer 2 | Missing Loan (not in PIF report) | — | Not present | Present (no PIF entry found for this loan ID) |
| MSR100952 | FHLMC | Layer 2 | Missing Loan (not in PIF report) | — | Not present | Present (no PIF entry found for this loan ID) |

---

## Yellow Lights

| Loan ID | Investor | Layer | Rule | Field | Submitted | Expected |
|---|---|---|---|---|---|---|
| MSR100000 | FNMA | Layer 1 | NSF May Be Expressed as Percent | Net Serv Fee | 0.2500 | ~0.0019–0.0069 (GNMA) or 0.0025 (FNMA/FHLMC) |
| MSR100003 | FNMA | Layer 1 | NSF May Be Expressed as Percent | Net Serv Fee | 0.2500 | ~0.0019–0.0069 (GNMA) or 0.0025 (FNMA/FHLMC) |
| MSR100011 | GNMA | Layer 1 | Next Due Date in Past (Current Loan) | Next Due Date | 2025-06-01 | >= 2026-01-31 for Current-status loans |
| MSR100012 | FHLMC | Layer 1 | Next Due Date in Past (Current Loan) | Next Due Date | 2025-06-01 | >= 2026-01-31 for Current-status loans |
| MSR100001 | FHLMC | Layer 2 | Status Bucket Skip | Status | Current -> 90+ DPD | Max 1-bucket change per month |
| MSR100008 | FHLMC | Layer 2 | P&I Inflated vs Prior Month | P&I ($) | $2,185.25 | ~$1,821.04 (unchanged from prior month) |
| MSR100009 | FNMA | Layer 2 | P&I Inflated vs Prior Month | P&I ($) | $1,718.66 | ~$1,432.22 (unchanged from prior month) |
| MSR100004 | FNMA | Layer 2 | Status Bucket Skip | Status | Current -> 90+ DPD | Max 1-bucket change per month |
| MSR100013 | FNMA | Layer 2 | Remaining Term Did Not Decrease | Rem Term | 112.0 | <= 111.0 (should not increase) |

---

## Missing Loans

| Loan ID | Investor | Prior UPB ($) |
|---|---|---:|
| MSR100034 | FHLMC | $424,666.91 |
| MSR100100 | FNMA | $291,530.86 |
| MSR100102 | FHLMC | $197,083.68 |
| MSR100103 | FHLMC | $181,657.15 |
| MSR100104 | GNMA | $148,285.60 |
| MSR100241 | GNMA | $265,375.94 |
| MSR100250 | FNMA | $106,147.08 |
| MSR100252 | FNMA | $365,421.00 |
| MSR100253 | FHLMC | $304,042.36 |
| MSR100264 | GNMA | $502,226.21 |
| MSR100625 | FNMA | $450,060.90 |
| MSR100709 | FNMA | $311,889.86 |
| MSR100911 | FNMA | $212,771.66 |
| MSR100913 | FHLMC | $374,742.05 |
| MSR100952 | FHLMC | $233,578.42 |

---

_Report generated by MSR Tape Validator — February 21, 2026_
