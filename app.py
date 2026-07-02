import os
import re
import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import plotly.express as px

st.set_page_config(page_title="SBI POS Activation Status", page_icon="🏧", layout="wide")

# ---------------------------------------------------------------------------
# Constants (activation rules)
# ---------------------------------------------------------------------------
MIN_TRANSACTIONS = 10
MIN_AMOUNT = 5000

# Master data ships inside the repo and is updated once a month by replacing
# this file (and pushing to GitHub) -- no need to upload it every session.
MASTER_DATA_FILENAME = "master_data.csv"
MASTER_DATA_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), MASTER_DATA_FILENAME)

MASTER_COLUMN_MAP = {
    "division-office-name": "Division Name",
    "office-name": "Office Name",
    "office-id": "Office ID",
    "office-type-code": "Office Type",
}

# Known transaction-file header variants -> canonical field. Headers are
# normalized (lowercased, hyphens -> spaces, whitespace collapsed) before
# comparison, so "SBIPOS-CARD (Cnt)" and "sbipos card (cnt)" both match.
TXN_COLUMN_ALIASES = {
    "Card Count": ["sbipos card (cnt)", "sbi pos card(count)", "sbi pos card (count)", "card count"],
    "Card Amount": ["sbipos card (amt)", "sbi pos card(amount)", "sbi pos card (amount)", "card amount"],
    "QR Count": ["sbipos bharatqr (cnt)", "sbi pos qr(count)", "sbi pos qr (count)", "bharatqr count", "qr count"],
    "QR Amount": ["sbipos bharatqr (amt)", "sbi pos qr(amount)", "sbi pos qr (amount)", "bharatqr amount", "qr amount"],
}

# ---------------------------------------------------------------------------
# Helper functions
# ---------------------------------------------------------------------------

def normalize_header(h):
    h = str(h).strip().lower().replace("-", " ")
    return re.sub(r"\s+", " ", h)


def read_any(file):
    """Read an uploaded csv/xlsx file into a dict of {sheet_name: DataFrame}."""
    name = file.name.lower()
    if name.endswith(".csv"):
        df = pd.read_csv(file)
        return {"Sheet1": df}
    xls = pd.ExcelFile(file, engine="openpyxl")
    return {sheet: xls.parse(sheet) for sheet in xls.sheet_names}


def clean_numeric(series):
    return pd.to_numeric(
        series.astype(str).str.replace(",", "", regex=False).str.strip(),
        errors="coerce",
    ).fillna(0)


def normalize_key(series):
    """Normalize a join key (name or id) so minor formatting differences
    (extra spaces, case, trailing .0) don't cause false 'unmapped' mismatches."""
    return (
        series.astype(str)
        .str.strip()
        .str.replace(r"\.0$", "", regex=True)
        .str.upper()
        .str.replace(r"\s+", " ", regex=True)
    )


@st.cache_data
def load_master_data(path):
    df = pd.read_csv(path)
    missing = [c for c in MASTER_COLUMN_MAP if c not in df.columns]
    if missing:
        raise ValueError(f"master_data.csv is missing expected column(s): {missing}")
    df = df.rename(columns=MASTER_COLUMN_MAP)[list(MASTER_COLUMN_MAP.values())]
    for c in ["Division Name", "Office Name", "Office Type"]:
        df[c] = df[c].astype(str).str.strip()
    df["Office ID"] = df["Office ID"].astype(str).str.strip()
    df = df.drop_duplicates(subset="Office Name")
    return df


def auto_detect_txn_columns(columns):
    """Match transaction-file headers to canonical fields by known aliases."""
    normalized = {normalize_header(c): c for c in columns}
    detected = {}
    for canonical, aliases in TXN_COLUMN_ALIASES.items():
        match = None
        for alias in aliases:
            if alias in normalized:
                match = normalized[alias]
                break
        detected[canonical] = match
    return detected


def auto_detect_identifier(txn_df, master_df, sample_rows=500):
    """Guess which transaction-file column identifies the office, and whether
    it should be matched against Office Name or Office ID, by checking which
    column's values overlap the most with master data (after normalizing)."""
    master_name_keys = set(normalize_key(master_df["Office Name"]))
    master_id_keys = set(normalize_key(master_df["Office ID"]))
    best = None  # (col, key_type, score)
    sample = txn_df.head(sample_rows)
    for col in txn_df.columns:
        vals = normalize_key(sample[col].astype(str))
        vals = vals[vals != ""]
        if len(vals) == 0:
            continue
        name_score = vals.isin(master_name_keys).mean()
        id_score = vals.isin(master_id_keys).mean()
        if best is None or name_score > best[2]:
            best = (col, "Office Name", name_score)
        if id_score > best[2]:
            best = (col, "Office ID", id_score)
    return best


def to_excel_bytes(sheets: dict):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        for sheet_name, df in sheets.items():
            df.to_excel(writer, index=False, sheet_name=sheet_name[:31])
    return output.getvalue()


# ---------------------------------------------------------------------------
# Header
# ---------------------------------------------------------------------------
st.title("🏧 SBI POS Machine Activation Status")
st.caption(
    "Master office/division data is bundled with this app. Just upload the "
    "monthly transaction file — columns and office matching are detected "
    "automatically."
)

with st.expander("ℹ️ Activation rule used in this app", expanded=False):
    st.markdown(
        f"""
QR (BharatQR) and Card figures are **combined** for each office:

- **Total Count** = QR Count + Card Count
- **Total Amount** = QR Amount + Card Amount

An office's POS machine is marked **ACTIVE** if **either** is true:

| Condition 1 (transactions) | Condition 2 (amount) |
|---|---|
| Total Count ≥ {MIN_TRANSACTIONS} | Total Amount > ₹{MIN_AMOUNT:,} |

**Overall status = ACTIVE if Total Count ≥ {MIN_TRANSACTIONS} OR Total Amount > ₹{MIN_AMOUNT:,}.**
Otherwise the office is marked **INACTIVE**.
"""
    )

# ---------------------------------------------------------------------------
# Master data (auto-loaded from the bundled file)
# ---------------------------------------------------------------------------
try:
    default_master = load_master_data(MASTER_DATA_PATH)
except FileNotFoundError:
    st.error(f"`{MASTER_DATA_FILENAME}` was not found next to app.py. Add it to the repo root and redeploy.")
    st.stop()
except ValueError as e:
    st.error(str(e))
    st.stop()

master_df = st.session_state.get("master_override_df", default_master)
using_override = "master_override_df" in st.session_state

st.success(
    f"✅ Master data loaded: **{len(master_df)} offices** across "
    f"**{master_df['Division Name'].nunique()} divisions**"
    + (" (session override active)" if using_override else f" — from `{MASTER_DATA_FILENAME}`")
)

with st.expander("🔄 Update master data", expanded=False):
    st.caption(
        "Master data is meant to be updated **once a month**. For a permanent "
        f"update: edit/replace `{MASTER_DATA_FILENAME}` in the GitHub repo (same "
        "column names: `division-office-name`, `office-id`, `office-name`, "
        "`office-type-code`) and push — the hosted app picks it up automatically."
    )
    st.caption("For a one-off override in this session only, upload a replacement file below:")
    override_upload = st.file_uploader("Upload replacement master file (csv/xlsx)", type=["csv", "xlsx", "xls"], key="master_upload")
    if override_upload is not None:
        try:
            raw = list(read_any(override_upload).values())[0]
            missing = [c for c in MASTER_COLUMN_MAP if c not in raw.columns]
            if missing:
                st.error(f"Uploaded file is missing expected column(s): {missing}")
            else:
                raw = raw.rename(columns=MASTER_COLUMN_MAP)[list(MASTER_COLUMN_MAP.values())]
                for c in ["Division Name", "Office Name", "Office Type"]:
                    raw[c] = raw[c].astype(str).str.strip()
                raw["Office ID"] = raw["Office ID"].astype(str).str.strip()
                raw = raw.drop_duplicates(subset="Office Name")
                st.session_state["master_override_df"] = raw
                st.success(f"Override loaded: {len(raw)} offices.")
                st.rerun()
        except Exception as e:
            st.error(f"Could not read file: {e}")
    if using_override and st.button("Revert to bundled master data"):
        del st.session_state["master_override_df"]
        st.rerun()

# ---------------------------------------------------------------------------
# Transaction file upload -> fully automatic processing
# ---------------------------------------------------------------------------
st.markdown("### 📤 Upload Transaction File")
txn_file = st.file_uploader("Office-wise monthly transaction file", type=["csv", "xlsx", "xls"], key="txn_upload")

if txn_file:
    txn_sheets = read_any(txn_file)
    t_sheet = (
        st.selectbox("Transaction file sheet", list(txn_sheets.keys()))
        if len(txn_sheets) > 1
        else list(txn_sheets.keys())[0]
    )
    txn_df = txn_sheets[t_sheet]

    detected_cols = auto_detect_txn_columns(txn_df.columns)
    id_guess = auto_detect_identifier(txn_df, master_df)

    missing_metric_cols = [k for k, v in detected_cols.items() if v is None]
    id_col_guess, id_key_guess, id_score = id_guess if id_guess else (None, "Office Name", 0)
    need_manual_id = id_guess is None or id_score < 0.5

    with st.expander("🧭 Detected column mapping (click to review or adjust)", expanded=bool(missing_metric_cols or need_manual_id)):
        st.caption("Auto-detected from your file's headers and values. Override anything below if it looks wrong.")
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("**Amount / count columns**")
            for canonical in ["QR Count", "QR Amount", "Card Count", "Card Amount"]:
                default_val = detected_cols[canonical]
                options = list(txn_df.columns)
                idx = options.index(default_val) if default_val in options else 0
                detected_cols[canonical] = st.selectbox(
                    canonical, options, index=idx, key=f"map_{canonical}"
                )
        with c2:
            st.markdown("**Office matching**")
            match_key = st.radio("Match using", ["Office Name", "Office ID"], index=0 if id_key_guess == "Office Name" else 1, horizontal=True)
            options = list(txn_df.columns)
            idx = options.index(id_col_guess) if id_col_guess in options else 0
            txn_key_col = st.selectbox(f"{match_key} column (transaction file)", options, index=idx)
            if id_guess:
                st.caption(f"Auto-match confidence: {id_score:.0%} of sampled values matched master data.")

    if not need_manual_id:
        match_key, txn_key_col = id_key_guess, id_col_guess

    # ---- run computation automatically ----
    master = master_df.copy()
    master_key_col = "Office Name" if match_key == "Office Name" else "Office ID"
    master["Merge Key"] = normalize_key(master[master_key_col])

    txn = txn_df[[txn_key_col] + [detected_cols[c] for c in ["QR Count", "QR Amount", "Card Count", "Card Amount"]]].copy()
    txn.columns = ["Merge Key Raw", "QR Count", "QR Amount", "Card Count", "Card Amount"]
    txn["Merge Key"] = normalize_key(txn["Merge Key Raw"])
    for c in ["QR Count", "QR Amount", "Card Count", "Card Amount"]:
        txn[c] = clean_numeric(txn[c])
    txn = txn.groupby("Merge Key", as_index=False).agg(
        {"Merge Key Raw": "first", "QR Count": "sum", "QR Amount": "sum", "Card Count": "sum", "Card Amount": "sum"}
    )

    merged = pd.merge(master, txn, on="Merge Key", how="left")
    unmapped_txn = txn[~txn["Merge Key"].isin(master["Merge Key"])]

    for c in ["QR Count", "QR Amount", "Card Count", "Card Amount"]:
        merged[c] = merged[c].fillna(0)

    merged["Total Count"] = merged["QR Count"] + merged["Card Count"]
    merged["Total Amount"] = merged["QR Amount"] + merged["Card Amount"]

    is_active = (merged["Total Count"] >= MIN_TRANSACTIONS) | (merged["Total Amount"] > MIN_AMOUNT)
    merged["Activation Status"] = np.where(is_active, "ACTIVE", "INACTIVE")
    merged = merged.drop(columns=["Merge Key"])
    unmapped_txn = unmapped_txn.drop(columns=["Merge Key"]).rename(columns={"Merge Key Raw": match_key})

    st.session_state["result"] = merged
    st.session_state["unmapped"] = unmapped_txn

# ---------------------------------------------------------------------------
# Results
# ---------------------------------------------------------------------------
if "result" in st.session_state:
    merged = st.session_state["result"]
    unmapped = st.session_state["unmapped"]

    st.markdown("---")
    st.markdown("## 📊 Results")

    total = len(merged)
    active = (merged["Activation Status"] == "ACTIVE").sum()
    inactive = total - active
    pct = round(active / total * 100, 1) if total else 0

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Total Offices", total)
    k2.metric("Active", active)
    k3.metric("Inactive", inactive)
    k4.metric("Activation %", f"{pct}%")

    st.markdown("### 🔍 Filter")
    f1, f2, f3 = st.columns(3)
    with f1:
        divisions = ["All"] + sorted(merged["Division Name"].dropna().unique().tolist())
        sel_div = st.selectbox("Division", divisions)
    with f2:
        office_types = ["All"] + sorted(merged["Office Type"].dropna().unique().tolist())
        sel_type = st.selectbox("Office Type", office_types)
    with f3:
        status_filter = st.selectbox("Status", ["All", "ACTIVE", "INACTIVE"])

    view = merged.copy()
    if sel_div != "All":
        view = view[view["Division Name"] == sel_div]
    if sel_type != "All":
        view = view[view["Office Type"] == sel_type]
    if status_filter != "All":
        view = view[view["Activation Status"] == status_filter]

    st.dataframe(view, use_container_width=True, hide_index=True)

    st.markdown("### 🏢 Division-wise Activation")
    div_summary = (
        merged.groupby("Division Name")["Activation Status"]
        .apply(lambda s: (s == "ACTIVE").sum())
        .reset_index(name="Active Count")
    )
    div_summary["Total Offices"] = merged.groupby("Division Name")["Office Name"].count().values
    div_summary["Inactive Count"] = div_summary["Total Offices"] - div_summary["Active Count"]
    div_summary["Activation %"] = round(div_summary["Active Count"] / div_summary["Total Offices"] * 100, 1)

    fig = px.bar(
        div_summary, x="Division Name", y=["Active Count", "Inactive Count"], barmode="stack",
        title="Active vs Inactive Offices by Division", color_discrete_sequence=["#2ecc71", "#e74c3c"],
    )
    st.plotly_chart(fig, use_container_width=True)

    fig2 = px.pie(
        merged, names="Activation Status", title="Overall Activation Split", color="Activation Status",
        color_discrete_map={"ACTIVE": "#2ecc71", "INACTIVE": "#e74c3c"},
    )
    st.plotly_chart(fig2, use_container_width=True)

    if len(unmapped) > 0:
        st.markdown("### ⚠️ Unmapped Offices")
        st.warning(
            f"{len(unmapped)} office(s) found in the transaction file with no matching "
            "entry in the master data. They are excluded from the results above — "
            "check for spelling/ID mismatches."
        )
        st.dataframe(unmapped, use_container_width=True, hide_index=True)

    st.markdown("### ⬇️ Download Report")
    export_dict = {"Activation Status": merged, "Division Summary": div_summary}
    if len(unmapped) > 0:
        export_dict["Unmapped Offices"] = unmapped
    excel_bytes = to_excel_bytes(export_dict)
    st.download_button(
        "Download Excel Report", data=excel_bytes, file_name="SBI_POS_Activation_Status.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.info("👆 Upload the transaction file above — results will appear automatically.")
