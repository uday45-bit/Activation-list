# SBI POS Machine Activation Status

A Streamlit web app that checks which office POS machines are **ACTIVE** for a
given month, based on SBI POS QR and Card transaction data.

## Activation rule

QR and Card figures are **combined** per office before checking the conditions:

- **Total Count** = QR Count + Card Count
- **Total Amount** = QR Amount + Card Amount

An office is **ACTIVE** if **both** are true:

| Condition 1 (transactions) | Condition 2 (amount) |
|---|---|
| Total Count ≥ 10 | Total Amount ≥ ₹5,000 |

Otherwise it is **INACTIVE**.

You upload two files inside the app itself (nothing is hard-coded):

1. **Master file** — Office Name + Division Name (for mapping offices to divisions).
2. **Transaction file** — Office-wise monthly transactions, containing the four
   columns: `SBI POS QR(Count)`, `SBI POS QR(Amount)`, `SBI POS Card(Count)`,
   `SBI POS Card(Amount)` (column names in your file can be anything — you map
   them inside the app).

Both files can be `.csv`, `.xlsx`, or `.xls`.

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
