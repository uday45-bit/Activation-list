# SBI POS Machine Activation Status

A Streamlit web app that checks which office POS machines are **ACTIVE** for a
given month, based on SBI POS QR and Card transaction data.

## Activation rule

QR (BharatQR) and Card figures are **combined** per office:

- **Total Count** = QR Count + Card Count
- **Total Amount** = QR Amount + Card Amount

An office is **ACTIVE** if **either** is true:

| Condition 1 (transactions) | Condition 2 (amount) |
|---|---|
| Total Count ≥ 10 | Total Amount > ₹5,000 |

Otherwise it is **INACTIVE**.

## Output

The screen shows exactly two boxes, then the result:

1. **🔄 Update Reference Data** box — shows how many offices/divisions are
   loaded from master data, and the POS supply/whitelisting totals, with
   optional one-off override uploads for each.
2. **📤 Upload Transaction File** box — pick the **report period**
   ("Current month (till date)" or "Previous month (complete)"), set the
   **division activation criteria (%)**, and upload the transaction file.

As soon as the transaction file is uploaded:

- A **color-coded division-wise summary image** appears, **sorted in
  descending order of activation %**, with a bold **Total** row at the
  bottom. Columns: `Office ID` (Division ID), `Office Name` (Division Name),
  `SBI POS Machines Issued`, `Active`, `Inactive`, `% Activation`. The
  heading is two lines: the report period, and the office-level activation
  criteria plus the division activation criteria you chose.
- Color coding is **3-tier**, driven by the criteria % you set:
  - **Green** — division activation % is at or above the criteria.
  - **Pink** — up to 10 percentage points below the criteria.
  - **Red** — more than 10 percentage points below the criteria.
- A single **Download Excel Report** button containing 3 sheets:
  1. **Division Summary** — the same color-coded, descending-sorted table
     shown on screen (2-line title in one wrapped cell, no gap before the
     header row).
  2. **Active Offices** — Division + Office Name + combined count/amount,
     for offices currently at/above the activation threshold.
  3. **POS Supply Overview** (division-wise) — Devices Supplied, No. of
     Offices (SPO/HPO/GPO + the 5 named Hyderabad Sorting Division
     exceptions), Not Supplied count, Not Whitelisted count, and **Excess**
     (Devices Supplied − No. of Offices).

Nothing else is shown on screen. If a column genuinely can't be
auto-detected, a small confirmation prompt appears inside the "Upload
Transaction File" box (not a separate box) — otherwise the output appears
immediately after upload, no extra clicks needed.

**Master data (`master_data.csv`) and POS supply data
(`pos_supply_status.xlsx`) both ship inside the repo** and are auto-loaded —
you don't upload either every session. Master data columns: `circle-office-id`,
`circle-name`, `region-office-id`, `region-office-name`, `division-office-id`,
`division-office-name`, `office-id`, `office-name`, `office-type-code` (only
the division and office columns are currently used).

### Machines vs. offices: the Hyderabad GPO case

A division's **"SBI POS Machines Issued"** figure comes from the POS supply
data, not the raw office count — the two aren't always equal. Hyderabad GPO
Division has just **1 office** but **7 machines supplied**. Active/Inactive
counts are the office-level activation ratio applied to the machines-issued
count, so if that one office is active, **all 7 machines** show as active
(and if it's inactive, all 7 show as inactive). This scales the same way for
any division where machines supplied and office count don't match exactly.

### Office types counted toward the report

Not every row in the master file counts as an office for this report. Only
these office types are **included**: `SPO`, `HPO`, `GPO` — plus five named
Hyderabad Sorting Division exceptions that do carry SBI POS machines despite
not being SPO/HPO/GPO: PBC Autonagar, SPC Counter, BPC Jeedimetla,
Secunderabad RSTMO, and Hyderabad Deccan RSTMO. This is set once as
`INCLUDED_OFFICE_TYPES` and `SPECIAL_INCLUDE_OFFICE_IDS` near the top of
`app.py` — update those if the policy on which offices count ever changes.

### POS supply & whitelisting data (`pos_supply_status.xlsx`)

This bundled workbook needs **3 sheets**, in this structure:

1. A monthly summary sheet (any name, e.g. "July 2026") with columns:
   Division, SBI POS Machines Supplied, Offices not whitelisted, Offices not
   supplied with SBI POS. A blank leading row above the header is handled
   automatically. The "All Divisions" total row (if present) is ignored —
   the app recomputes totals itself.
2. A "Whitelisting Pending" detail sheet: Division, Office ID, Office Name.
3. A "Not supplied offices" detail sheet: Division Name, Office Name.

The app reads sheet 1 positionally (since its name changes monthly) and
finds sheets 2–3 by keyword match, so you can keep just updating the same
workbook and re-uploading/replacing it each month without renaming anything.

## Office matching

**Matching is always done by Office ID** (the most reliable key) — the app
auto-detects which transaction-file column holds it by comparing values
against master data. Office Name and Division Name are used only for
display in the outputs, never for matching.

## Uploading the transaction file

Each session you only need to:

1. Pick the **report period** in the "Upload Transaction File" box —
   "Current month (till date)" (1st of this month to today) or "Previous
   month (complete)" (the full prior calendar month).
2. Upload the **transaction file** (`.csv`, `.xlsx`, or `.xls`).

The app auto-detects these four columns by header name: `SBIPOS-CARD (Cnt)`,
`SBIPOS-CARD (Amt)`, `SBIPOS BHARATQR (Cnt)`, `SBIPOS BHARATQR (Amt)` (a few
common naming variants are also recognized), plus the Office ID column as
described above.

### Updating reference data monthly

Master data and POS supply data both change about once a month. To update
either permanently:

```bash
# replace with the new version, keeping the same column names/sheet structure
git add master_data.csv pos_supply_status.xlsx
git commit -m "Update reference data for <month>"
git push
```

Streamlit Cloud redeploys automatically and the app picks up the new files.

If you need a quick one-off check without touching GitHub, use the
**"🔄 Update Reference Data"** box inside the app to upload a replacement
file for the current session only (it resets on reload — not a permanent fix).

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

Then open the local URL shown in the terminal (usually `http://localhost:8501`).

## Deploy for free on Streamlit Community Cloud (via GitHub)

**Step 1 — Push this project to GitHub**

```bash
cd sbi-pos-activation
git init
git add .
git commit -m "Initial commit: SBI POS activation status app"
git branch -M main
git remote add origin https://github.com/<your-username>/<your-repo>.git
git push -u origin main
```

(Create the empty repo first on github.com if you haven't — click **New repository**,
give it a name, don't add a README there since you already have one.)

**Step 2 — Deploy on Streamlit**

1. Go to **https://share.streamlit.io** and sign in with your GitHub account.
2. Click **"Create app"** → **"Deploy a public app from GitHub"**.
3. Select your repository, branch (`main`), and main file path (`app.py`).
4. Click **Deploy**. Streamlit installs `requirements.txt` automatically and
   gives you a public URL like `https://<your-app-name>.streamlit.app`.

Any time you `git push` new changes to the repo, the hosted app auto-updates.

## Project structure

```
sbi-pos-activation/
├── app.py                    # Streamlit app
├── master_data.csv           # Office/Division master data (update ~monthly)
├── pos_supply_status.xlsx    # POS supply/whitelisting reference (update ~monthly)
├── requirements.txt          # Python dependencies
└── README.md                 # This file
```

## Notes

- No data is stored on any server — files are processed in-memory for your
  session only.
- The bundled `master_data.csv`, after office-type filtering, currently
  resolves to **280 offices across 7 divisions**: Hyderabad City, Hyderabad
  GPO, Hyderabad South East, Medak, Sangareddy, Secunderabad, and Hyderabad
  Sorting Division.
- Offices in the transaction file that don't match any Office ID in master
  data are simply excluded from the results (they contribute nothing to
  totals). If your active/inactive counts look off, check the transaction
  file's Office ID column against master data for formatting mismatches.
- The activation rule (combined QR+Card totals, ≥10 transactions OR >₹5,000)
  is defined once as constants at the top of `app.py` — change
  `MIN_TRANSACTIONS` / `MIN_AMOUNT` there if the policy changes. The division
  color-coding criteria (%) is set live in the app each session, not hard-coded.
