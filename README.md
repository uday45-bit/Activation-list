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

1. **🔄 Update Master Data** box — shows how many offices/divisions are
   currently loaded, with an optional one-off override upload.
2. **📤 Upload Transaction File** box — pick the **report period**
   ("Current month (till date)" or "Previous month (complete)") and upload
   the transaction file.

As soon as the transaction file is uploaded:

- A **color-coded division-wise summary image** appears — green row if that
  division's activation % is ≥ 50%, red if < 50%, with a bold **Total** row.
  Columns match the standard consolidated report format: `Office ID`
  (Division ID), `Office Name` (Division Name), `SBI POS Machines Issued`,
  `Active`, `Inactive`, `% Activation`. Heading reads `SBI POS Machine
  Activation Status (Period: dd.mm.yyyy to dd.mm.yyyy)` using the period you
  picked.
- A single **Download Excel Report** button — same color-coded division
  summary (title and header sit right next to each other, no blank row),
  plus a second sheet listing only offices currently **INACTIVE** (Division
  Name + Office Name, with their combined transaction count/amount for
  context). There's no "unmapped offices" sheet in the export.

Nothing else is shown on screen. If a column genuinely can't be
auto-detected, a small confirmation prompt appears inside the "Upload
Transaction File" box (not a separate box) — otherwise the output appears
immediately after upload, no extra clicks needed.

**Master data (`master_data.csv`) ships inside the repo** and is auto-loaded
by the app — you don't upload it every session. It's expected to have these
columns: `circle-office-id`, `circle-name`, `region-office-id`,
`region-office-name`, `division-office-id`, `division-office-name`,
`office-id`, `office-name`, `office-type-code` (only the division and office
columns are currently used; circle/region are carried through for future use).

### Office types counted toward the report

Not every row in the master file counts as an office for this report. These
office types are **excluded**: `BPO` (Branch Post Office), `PDN`, `RDN`,
`SDO`, `PSD`, `MMS` — along with any row with a missing/invalid division
(e.g. region-level administrative records). Everything else (SPO, HPO, GPO,
IDC, BPC, TMO, and other departmental/service types) counts. This is set
once as `EXCLUDED_OFFICE_TYPES` near the top of `app.py` — update that list
if the policy on which office types count ever changes.

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

### Updating master data monthly

Master data changes about once a month. To update it permanently:

```bash
# replace master_data.csv with the new version, keeping the same column names
git add master_data.csv
git commit -m "Update master data for <month>"
git push
```

Streamlit Cloud redeploys automatically and the app picks up the new file.

If you need a quick one-off check without touching GitHub, use the
**"🔄 Update Master Data"** box inside the app to upload a replacement file
for the current session only (it resets on reload — not a permanent fix).

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
├── app.py              # Streamlit app
├── master_data.csv     # Office/Division master data (update ~monthly)
├── requirements.txt    # Python dependencies
└── README.md           # This file
```

## Notes

- No data is stored on any server — files are processed in-memory for your
  session only.
- The bundled `master_data.csv` currently resolves to **7 divisions**:
  Hyderabad City, Hyderabad GPO, Hyderabad South East, Medak, Sangareddy,
  Secunderabad, and Hyderabad Sorting Division.
- Offices in the transaction file that don't match any Office ID in master
  data are simply excluded from the results (they contribute nothing to
  totals). If your active/inactive counts look off, check the transaction
  file's Office ID column against master data for formatting mismatches.
- The activation rule (combined QR+Card totals, ≥10 transactions OR >₹5,000)
  is defined once as constants at the top of `app.py` — change
  `MIN_TRANSACTIONS` / `MIN_AMOUNT` there if the policy changes.
