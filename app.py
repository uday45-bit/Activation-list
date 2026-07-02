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

# ---------------------------------------------------------------------------
# Helper functions
# ---------------------------------------------------------------------------

def read_any(file):
    """Read an uploaded csv/xlsx file into a dict of {sheet_name: DataFrame}."""
    name = file.name.lower()
    if name.endswith(".csv"):
        df = pd.read_csv(file)
        return {"Sheet1": df}
    xls = pd.ExcelFile(file, engine="openpyxl")
    return {sheet: xls.parse(sheet) for sheet in xls.sheet_names}


def clean_numeric(series):
    """Convert a messy numeric column (commas, blanks, text) into clean floats."""
    return pd.to_numeric(
        series.astype(str).str.replace(",", "", regex=False).str.strip(),
        errors="coerce",
    ).fillna(0)


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
    "Upload the Master Office/Division file and the monthly Office-wise "
    "Transaction file to compute POS activation status."
)

with st.expander("ℹ️ Activation rule used in this app", expanded=False):
    st.markdown(
        f"""
QR and Card figures are **combined** for each office before checking the
conditions:

- **Total Count** = QR Count + Card Count
- **Total Amount** = QR Amount + Card Amount

An office's POS machine is marked **ACTIVE** if **both** are true:

| Condition 1 (transactions) | Condition 2 (amount) |
|---|---|
| Total Count ≥ {MIN_TRANSACTIONS} | Total Amount ≥ ₹{MIN_AMOUNT:,} |

Otherwise the office is marked **INACTIVE**.
"""
    )

# ---------------------------------------------------------------------------
# File uploads
# ---------------------------------------------------------------------------
col1, col2 = st.columns(2)
with col1:
    master_file = st.file_uploader(
        "1️⃣ Upload Master File (Office + Division mapping)",
        type=["csv", "xlsx", "xls"],
    )
with col2:
    txn_file = st.file_uploader(
        "2️⃣ Upload Transaction File (Office-wise monthly transactions)",
        type=["csv", "xlsx", "xls"],
    )

if master_file and txn_file:
    master_sheets = read_any(master_file)
    txn_sheets = read_any(txn_file)

    m_sheet = (
        st.selectbox("Master file sheet", list(master_sheets.keys()))
        if len(master_sheets) > 1
        else list(master_sheets.keys())[0]
    )
    t_sheet = (
        st.selectbox("Transaction file sheet", list(txn_sheets.keys()))
        if len(txn_sheets) > 1
        else list(txn_sheets.keys())[0]
    )

    master_df = master_sheets[m_sheet]
    txn_df = txn_sheets[t_sheet]

    st.markdown("### 🔗 Column Mapping")
    st.caption("Tell the app which column in your files corresponds to each required field.")

    mcol1, mcol2 = st.columns(2)
    with mcol1:
        st.markdown("**Master file**")
        master_office_col = st.selectbox("Office Name column", master_df.columns, key="m_office")
        master_division_col = st.selectbox("Division Name column", master_df.columns, key="m_div")
    with mcol2:
        st.markdown("**Transaction file**")
        txn_office_col = st.selectbox("Office Name column", txn_df.columns, key="t_office")
        qr_count_col = st.selectbox("SBI POS QR (Count) column", txn_df.columns, key="qr_c")
        qr_amount_col = st.selectbox("SBI POS QR (Amount) column", txn_df.columns, key="qr_a")
        card_count_col = st.selectbox("SBI POS Card (Count) column", txn_df.columns, key="card_c")
        card_amount_col = st.selectbox("SBI POS Card (Amount) column", txn_df.columns, key="card_a")

    if st.button("🚀 Compute Activation Status", type="primary"):
        # ---- prep master ----
        master = master_df[[master_office_col, master_division_col]].copy()
        master.columns = ["Office Name", "Division Name"]
        master["Office Name"] = master["Office Name"].astype(str).str.strip()
        master = master.drop_duplicates(subset="Office Name")

        # ---- prep transactions ----
        txn = txn_df[
            [txn_office_col, qr_count_col, qr_amount_col, card_count_col, card_amount_col]
        ].copy()
        txn.columns = ["Office Name", "QR Count", "QR Amount", "Card Count", "Card Amount"]
        txn["Office Name"] = txn["Office Name"].astype(str).str.strip()
        for c in ["QR Count", "QR Amount", "Card Count", "Card Amount"]:
            txn[c] = clean_numeric(txn[c])
        # in case an office appears in multiple rows, aggregate first
        txn = txn.groupby("Office Name", as_index=False).sum(numeric_only=True)

        # ---- merge ----
        merged = pd.merge(master, txn, on="Office Name", how="left")
        unmapped_txn = txn[~txn["Office Name"].isin(master["Office Name"])]

        for c in ["QR Count", "QR Amount", "Card Count", "Card Amount"]:
            merged[c] = merged[c].fillna(0)

        merged["Total Count"] = merged["QR Count"] + merged["Card Count"]
        merged["Total Amount"] = merged["QR Amount"] + merged["Card Amount"]

        is_active = (merged["Total Count"] >= MIN_TRANSACTIONS) & (
            merged["Total Amount"] >= MIN_AMOUNT
        )
        merged["Activation Status"] = np.where(is_active, "ACTIVE", "INACTIVE")

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
    f1, f2 = st.columns(2)
    with f1:
        divisions = ["All"] + sorted(merged["Division Name"].dropna().unique().tolist())
        sel_div = st.selectbox("Division", divisions)
    with f2:
        status_filter = st.selectbox("Status", ["All", "ACTIVE", "INACTIVE"])

    view = merged.copy()
    if sel_div != "All":
        view = view[view["Division Name"] == sel_div]
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
    div_summary["Activation %"] = round(
        div_summary["Active Count"] / div_summary["Total Offices"] * 100, 1
    )

    fig = px.bar(
        div_summary,
        x="Division Name",
        y=["Active Count", "Inactive Count"],
        barmode="stack",
        title="Active vs Inactive Offices by Division",
        color_discrete_sequence=["#2ecc71", "#e74c3c"],
    )
    st.plotly_chart(fig, use_container_width=True)

    fig2 = px.pie(
        merged,
        names="Activation Status",
        title="Overall Activation Split",
        color="Activation Status",
        color_discrete_map={"ACTIVE": "#2ecc71", "INACTIVE": "#e74c3c"},
    )
    st.plotly_chart(fig2, use_container_width=True)

    if len(unmapped) > 0:
        st.markdown("### ⚠️ Unmapped Offices")
        st.warning(
            f"{len(unmapped)} office(s) found in the transaction file with no matching "
            "entry in the master file. They are excluded from the results above — "
            "check for spelling/naming mismatches."
        )
        st.dataframe(unmapped, use_container_width=True, hide_index=True)

    st.markdown("### ⬇️ Download Report")
    export_dict = {"Activation Status": merged, "Division Summary": div_summary}
    if len(unmapped) > 0:
        export_dict["Unmapped Offices"] = unmapped
    excel_bytes = to_excel_bytes(export_dict)
    st.download_button(
        "Download Excel Report",
        data=excel_bytes,
        file_name="SBI_POS_Activation_Status.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.info(
        "👆 Upload both files, map the columns, and click **Compute Activation Status** "
        "to see results."
    )
