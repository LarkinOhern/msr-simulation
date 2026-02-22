# Try It Yourself — MSR Tape Validation Challenge

> **SIMULATED DATA** — All loan information is synthetic and generated for testing purposes only.

---

## The Challenge

A subservicer submitted a monthly MSR tape. Somewhere in 1,188 loans there are errors — some obvious, some subtle. Can your process catch them before the tape gets accepted?

This project lets you run the same validator used in this simulation against a pre-built dirty tape. You'll see exactly which loans fail, why, and which issues require action vs. further review.

**No mortgage experience required. Just curiosity.**

---

## What You'll Run

| Script | What It Does |
|---|---|
| `validate_msr_tape.py` | Validates a subservicer submission against the prior month tape |

**Input files (already in the repo):**
- `MSR_Sample_Tape_Dec2025_Jan2026.xlsx` — prior month clean tape (Dec 2025 tab)
- `MSR_Tape_Jan2026_SUBSERVICER.xlsx` — dirty submission (22 injected errors)
- `Recon_PaidInFull_Jan2026.xlsx` — PIF recon (clears 12 legitimate payoffs)
- `Recon_NewAdds_Jan2026.xlsx` — new add recon (confirms 200 new loans)

**One command:**
```bash
python validate_msr_tape.py
```

**Outputs:**
- `Validation_Jan2026_SUBSERVICER.xlsx` — 4-tab Excel report
- `Validation_Jan2026_SUBSERVICER.md` — Markdown summary

---

## How to Access — Three Options

### Option A: GitHub Codespaces (Recommended — No Install Required)

The easiest path for non-technical users. All you need is a free GitHub account.

1. Go to [github.com/LarkinOhern/msr-simulation](https://github.com/LarkinOhern/msr-simulation)
2. Click the green **Code** button → **Codespaces** tab → **Create codespace on main**
3. Wait ~60 seconds for the environment to load (browser-based VS Code opens)
4. In the terminal at the bottom, run:
   ```bash
   pip install openpyxl
   python validate_msr_tape.py
   ```
5. Right-click any output `.xlsx` file in the file explorer → **Download** to open in Excel

That's it. No Python installation. No package managers. No command line on your own machine.

---

### Option B: Local Python (If You Have Python Installed)

```bash
git clone https://github.com/LarkinOhern/msr-simulation.git
cd msr-simulation
pip install openpyxl
python validate_msr_tape.py
```

Requirements: Python 3.8+, one package (`openpyxl`).

---

### Option C: Download and Run (Windows, No Git)

1. Download the repo as a ZIP from GitHub (Code → Download ZIP)
2. Extract to any folder
3. Open Command Prompt in that folder
4. Run:
   ```
   pip install openpyxl
   python validate_msr_tape.py
   ```

If you don't have Python: [python.org/downloads](https://www.python.org/downloads/) — install with "Add to PATH" checked.

---

## What to Look For

When the validator runs, it prints a summary to the terminal:

```
HARD STOPS:     13  <-- ACTION REQUIRED
YELLOW LIGHTS:   9  <-- REVIEW REQUIRED
PIF-Explained:  12  (cleared)
Unexplained:     3  <-- HARD STOP
```

Open `Validation_Jan2026_SUBSERVICER.xlsx` and check each tab:

| Tab | What It Shows |
|---|---|
| Summary | Scorecard with all counts |
| Hard Stops | 13 loans requiring immediate correction |
| Yellow Lights | 9 loans requiring manual review |
| Missing Loans | 15 missing — 12 PIF-cleared, 3 unexplained hard stops |

**Can you explain why each hard stop fired before reading the Error Log sheet in the SUBSERVICER file?**

---

## The Blind Test (Advanced)

The repository also includes `MSR_Tape_Jan2026_SUBSERVICER_LO_Jitter.xlsx` — a second dirty tape with a different set of errors, manually constructed with no knowledge of what the validator would check.

Run it:
```bash
python validate_msr_tape.py --submission MSR_Tape_Jan2026_SUBSERVICER_LO_Jitter.xlsx
```

Result: **8 for 8** — every unknown injection was caught.

Interesting catches:
- A `"60+ DPD"` typo (should be `"60 DPD"`) — realistic field encoding error
- A phantom loan `MSR300198` with no recon trail — appeared in submission, never boarded
- NSF values stacked on top of already-flagged UPB errors

---

## Make It Harder

Once you've run the default test, try modifying the submission yourself:

1. Open `MSR_Tape_Jan2026_SUBSERVICER.xlsx`
2. Pick any loan, change a field (rate, UPB, status, NSF)
3. Save as a new filename
4. Run: `python validate_msr_tape.py --submission your_file.xlsx`
5. See if the validator catches it

Error types the validator is looking for are documented in `Content_SubservicerChecklist.xlsx` — use it as a guide for what to inject.

---

## Share Your Results

If you run this and find something interesting — an edge case, a false positive, a catch you didn't expect — that's the kind of real-world feedback that makes validation logic better.

The validator is not exhaustive. It catches the errors it was designed for plus some it wasn't. The blind test proved that. The question for a production system is: what's it missing?

---

*Built with Python + openpyxl. No external dependencies beyond standard library.*
*Simulated data only — all loan IDs, balances, and borrower details are synthetic.*
