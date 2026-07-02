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

**Master data (`master_data.csv`) ships inside the repo** and is auto-loaded
by the app — you don't upload it every session. It's expected to have these
columns: `division-office-name`, `office-id`, `office-name`, `office-type-code`.

Each session you only need to upload:

- **Transaction file** — Office-wise monthly transactions. The app
  auto-detects these four columns by header name: `SBIPOS-CARD (Cnt)`,
  `SBIPOS-CARD (Amt)`, `SBIPOS BHARATQR (Cnt)`, `SBIPOS BHARATQR (Amt)`
  (a few common naming variants are also recognized). It also auto-detects
  which column identifies the office (by matching values against the master
  data) and whether to match by office name or office ID. If anything can't
  be auto-detected confidently, a "Detected column mapping" panel opens
  automatically so you can pick it manually — otherwise results appear as
  soon as the file is uploaded, no extra clicks needed.

The transaction file can be `.csv`, `.xlsx`, or `.xls`.

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
**"🔄 Update master data"** panel inside the app to upload a replacement file
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
- Offices present in the transaction file but missing from the master file
  are flagged separately as "Unmapped Offices" so you can catch naming
  mismatches (e.g. extra spaces, abbreviations).
- The activation rule (combined QR+Card totals) is defined once as constants
  at the top of `app.py` — change `MIN_TRANSACTIONS` / `MIN_AMOUNT` there if
  the policy changes.
