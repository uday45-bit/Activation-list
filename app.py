import os
import re
from datetime import date, timedelta
from io import BytesIO

import streamlit as st
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

st.set_page_config(page_title="SBI POS Activation Status", page_icon="🏧", layout="wide")

# ---------------------------------------------------------------------------
# Constants (activation rules)
# ---------------------------------------------------------------------------
MIN_TRANSACTIONS = 10
MIN_AMOUNT = 5000

MASTER_DATA_FILENAME = "master_data.csv"
MASTER_DATA_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), MASTER_DATA_FILENAME)

MASTER_COLUMN_MAP = {
    "division-office-id": "Division ID",
    "division-office-name": "Division Name",
    "office-name": "Office Name",
    "office-id": "Office ID",
    "office-type-code": "Office Type",
}

# Office types excluded from the activation-report universe (e.g. Branch Post
# Offices and other non-departmental/administrative types).
EXCLUDED_OFFICE_TYPES = ["BPO", "PDN", "RDN", "SDO", "PSD", "MMS"]

# Known transaction-file header variants -> canonical field.
TXN_COLUMN_ALIASES = {
    "Card Count": ["sbipos card (cnt)", "sbi pos card(count)", "sbi pos card (count)", "card count"],
    "Card Amount": ["sbipos card (amt)", "sbi pos card(amount)", "sbi pos card (amount)", "card amount"],
    "QR Count": ["sbipos bharatqr (cnt)", "sbi pos qr(count)", "sbi pos qr (count)", "bharatqr count", "qr count"],
    "QR Amount": ["sbipos bharatqr (amt)", "sbi pos qr(amount)", "sbi pos qr (amount)", "bharatqr amount", "qr amount"],
}

GREEN_BG, GREEN_TEXT = "#d9f2df", "#1e7d34"
RED_BG, RED_TEXT = "#fbdada", "#b83232"
HEADER_BG = "#2c3e50"
TOTAL_BG = "#eceff1"

# ---------------------------------------------------------------------------
# Helper functions
# ---------------------------------------------------------------------------

def normalize_header(h):
    h = str(h).strip().lower().replace("-", " ")
    return re.sub(r"\s+", " ", h)


def read_any(file):
    name = file.name.lower()
    if name.endswith(".csv"):
        return {"Sheet1": pd.read_csv(file)}
    xls = pd.ExcelFile(file, engine="openpyxl")
    return {sheet: xls.parse(sheet) for sheet in xls.sheet_names}


def clean_numeric(series):
    return pd.to_numeric(
        series.astype(str).str.replace(",", "", regex=False).str.strip(),
        errors="coerce",
    ).fillna(0)


def normalize_key(series):
    return (
        series.astype(str)
        .str.strip()
        .str.replace(r"\.0$", "", regex=True)
        .str.upper()
        .str.replace(r"\s+", " ", regex=True)
    )


def clean_master_df(df):
    missing = [c for c in MASTER_COLUMN_MAP if c not in df.columns]
    if missing:
        raise ValueError(f"File is missing expected column(s): {missing}")
    df = df.rename(columns=MASTER_COLUMN_MAP)[list(MASTER_COLUMN_MAP.values())]
    for c in ["Division Name", "Office Name", "Office Type"]:
        df[c] = df[c].astype(str).str.strip()
    df["Office ID"] = df["Office ID"].astype(str).str.strip()
    df["Division ID"] = df["Division ID"].astype(str).str.strip()
    df = df[~df["Office Type"].isin(EXCLUDED_OFFICE_TYPES)]
    df = df[~df["Division Name"].str.lower().isin(["", "undefined", "nan"])]
    df = df.drop_duplicates(subset="Office ID")
    return df


@st.cache_data
def load_master_data(path):
    return clean_master_df(pd.read_csv(path))


def auto_detect_txn_columns(columns):
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


def auto_detect_office_id_column(txn_df, master_df, sample_rows=500):
    """Find which transaction-file column holds Office ID by checking value
    overlap against master data (matching is always done via Office ID)."""
    master_id_keys = set(normalize_key(master_df["Office ID"]))
    best = None  # (col, score)
    sample = txn_df.head(sample_rows)
    for col in txn_df.columns:
        vals = normalize_key(sample[col].astype(str))
        vals = vals[vals != ""]
        if len(vals) == 0:
            continue
        score = vals.isin(master_id_keys).mean()
        if best is None or score > best[1]:
            best = (col, score)
    return best


def build_division_summary(merged):
    div_summary = (
        merged.groupby("Division Name")["Activation Status"]
        .apply(lambda s: (s == "ACTIVE").sum())
        .reset_index(name="Active Count")
    )
    div_summary["Total Offices"] = merged.groupby("Division Name")["Office Name"].count().values
    div_summary["Division ID"] = merged.groupby("Division Name")["Division ID"].first().values
    div_summary["Inactive Count"] = div_summary["Total Offices"] - div_summary["Active Count"]
    div_summary["Activation %"] = round(div_summary["Active Count"] / div_summary["Total Offices"] * 100, 1)
    return div_summary.sort_values("Division Name").reset_index(drop=True)


def build_division_summary_image(div_summary, period_str):
    columns = ["Office ID", "Office Name", "SBI POS Machines Issued", "Active", "Inactive", "% Activation"]
    total_offices = int(div_summary["Total Offices"].sum())
    total_active = int(div_summary["Active Count"].sum())
    total_inactive = int(div_summary["Inactive Count"].sum())
    total_pct = round(total_active / total_offices * 100, 1) if total_offices else 0

    rows, cell_colors = [], []
    for _, r in div_summary.iterrows():
        rows.append([r["Division ID"], r["Division Name"], int(r["Total Offices"]), int(r["Active Count"]), int(r["Inactive Count"]), f'{r["Activation %"]}%'])
        bg = GREEN_BG if r["Activation %"] >= 50 else RED_BG
        cell_colors.append([bg] * 6)
    rows.append(["Total", "", total_offices, total_active, total_inactive, f"{total_pct}%"])
    cell_colors.append([TOTAL_BG] * 6)

    n_rows = len(rows) + 1  # + header row
    fig_h = 0.6 + 0.42 * n_rows
    fig, ax = plt.subplots(figsize=(11.5, fig_h))
    ax.axis("off")
    ax.set_title(f"SBI POS Machine Activation Status (Period: {period_str})", fontsize=13, fontweight="bold", loc="left", pad=16)

    table = ax.table(
        cellText=rows,
        colLabels=columns,
        cellColours=cell_colors,
        colColours=[HEADER_BG] * 6,
        cellLoc="center",
        colWidths=[0.12, 0.25, 0.24, 0.13, 0.13, 0.13],
        loc="center",
    )
    table.auto_set_font_size(False)
    table.set_fontsize(10.5)
    table.scale(1, 1.7)

    n_data_rows = len(rows)
    for (row, col), cell in table.get_celld().items():
        cell.set_edgecolor("#b0bec5")
        if row == 0:
            cell.get_text().set_color("white")
            cell.get_text().set_fontweight("bold")
        if col in (0, 1):
            cell.get_text().set_ha("left")
            cell.PAD = 0.04
        if row == n_data_rows:  # grand total row (last row, header counts as row 0)
            cell.get_text().set_fontweight("bold")

    plt.tight_layout()
    buf = BytesIO()
    fig.savefig(buf, format="png", dpi=200, bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return buf.getvalue()


def build_excel_report(merged, div_summary, period_str):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book

        title_fmt = workbook.add_format({"bold": True, "font_size": 14})
        header_fmt = workbook.add_format({
            "bold": True, "bg_color": HEADER_BG, "font_color": "white",
            "border": 1, "align": "center", "valign": "vcenter",
        })
        green_fmt = workbook.add_format({"bg_color": GREEN_BG, "border": 1, "align": "center"})
        green_fmt_left = workbook.add_format({"bg_color": GREEN_BG, "border": 1, "align": "left"})
        red_fmt = workbook.add_format({"bg_color": RED_BG, "border": 1, "align": "center"})
        red_fmt_left = workbook.add_format({"bg_color": RED_BG, "border": 1, "align": "left"})
        total_fmt = workbook.add_format({"bold": True, "bg_color": TOTAL_BG, "border": 1, "align": "center"})
        total_fmt_left = workbook.add_format({"bold": True, "bg_color": TOTAL_BG, "border": 1, "align": "left"})

        # ---- Sheet 1: Division Summary (title -> header immediately below, no gap) ----
        ws = workbook.add_worksheet("Division Summary")
        ws.merge_range(0, 0, 0, 5, f"SBI POS Machine Activation Status (Period: {period_str})", title_fmt)

        headers = ["Office ID", "Office Name", "SBI POS Machines Issued", "Active", "Inactive", "% Activation"]
        for col, h in enumerate(headers):
            ws.write(1, col, h, header_fmt)

        row_idx = 2
        for _, row in div_summary.iterrows():
            is_green = row["Activation %"] >= 50
            fmt, fmt_left = (green_fmt, green_fmt_left) if is_green else (red_fmt, red_fmt_left)
            ws.write(row_idx, 0, row["Division ID"], fmt_left)
            ws.write(row_idx, 1, row["Division Name"], fmt_left)
            ws.write(row_idx, 2, int(row["Total Offices"]), fmt)
            ws.write(row_idx, 3, int(row["Active Count"]), fmt)
            ws.write(row_idx, 4, int(row["Inactive Count"]), fmt)
            ws.write(row_idx, 5, row["Activation %"], fmt)
            row_idx += 1

        total_offices = int(div_summary["Total Offices"].sum())
        total_active = int(div_summary["Active Count"].sum())
        total_inactive = int(div_summary["Inactive Count"].sum())
        total_pct = round(total_active / total_offices * 100, 1) if total_offices else 0
        ws.write(row_idx, 0, "Total", total_fmt_left)
        ws.write(row_idx, 1, "", total_fmt_left)
        ws.write(row_idx, 2, total_offices, total_fmt)
        ws.write(row_idx, 3, total_active, total_fmt)
        ws.write(row_idx, 4, total_inactive, total_fmt)
        ws.write(row_idx, 5, total_pct, total_fmt)

        ws.set_column(0, 0, 14)
        ws.set_column(1, 1, 30)
        ws.set_column(2, 5, 16)

        # ---- Sheet 2: Inactive Offices only (Office Name + Division Name, not raw ID) ----
        inactive_df = (
            merged[merged["Activation Status"] == "INACTIVE"]
            [["Division Name", "Office Name", "Total Count", "Total Amount"]]
            .sort_values(["Division Name", "Office Name"])
            .reset_index(drop=True)
        )
        inactive_df.to_excel(writer, sheet_name="Inactive Offices", index=False)
        ws2 = writer.sheets["Inactive Offices"]
        ws2.set_column(0, 1, 30)
        ws2.set_column(2, 3, 15)

    return output.getvalue()


# ---------------------------------------------------------------------------
# UI: exactly two boxes (Update Master Data, Upload Transaction File),
# then the division-wise visualization, then a single Excel download.
# ---------------------------------------------------------------------------
st.title("🏧 SBI POS Machine Activation Status")

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

# ---- Box 1: Update Master Data ----
with st.container(border=True):
    st.markdown("#### 🔄 Update Master Data")
    st.caption(
        f"Currently loaded: **{len(master_df)} offices** across **{master_df['Division Name'].nunique()} divisions**"
        + (" (session override active)." if using_override else f", from `{MASTER_DATA_FILENAME}`.")
    )
    override_upload = st.file_uploader(
        "Upload replacement master file for a one-off override (this session only). "
        "For a permanent monthly update, replace master_data.csv in the GitHub repo and push.",
        type=["csv", "xlsx", "xls"], key="master_upload",
    )
    if override_upload is not None:
        try:
            raw = list(read_any(override_upload).values())[0]
            raw = clean_master_df(raw)
            st.session_state["master_override_df"] = raw
            st.success(f"Override loaded: {len(raw)} offices. Re-upload the transaction file below to apply it.")
        except ValueError as e:
            st.error(str(e))
        except Exception as e:
            st.error(f"Could not read file: {e}")
    if using_override and st.button("Revert to bundled master data"):
        del st.session_state["master_override_df"]
        st.rerun()

# ---- Box 2: Upload Transaction File ----
with st.container(border=True):
    st.markdown("#### 📤 Upload Transaction File")
    period_choice = st.radio(
        "Report period", ["Current month (till date)", "Previous month (complete)"], horizontal=True
    )
    txn_file = st.file_uploader("Office-wise monthly transaction file", type=["csv", "xlsx", "xls"])
    manual_map_placeholder = st.container()

today = date.today()
if period_choice == "Current month (till date)":
    period_start, period_end = today.replace(day=1), today
else:
    first_of_this_month = today.replace(day=1)
    period_end = first_of_this_month - timedelta(days=1)
    period_start = period_end.replace(day=1)
period_str = f"{period_start.strftime('%d.%m.%Y')} to {period_end.strftime('%d.%m.%Y')}"

if txn_file:
    txn_sheets = read_any(txn_file)
    t_sheet = list(txn_sheets.keys())[0]
    if len(txn_sheets) > 1:
        with manual_map_placeholder:
            t_sheet = st.selectbox("Transaction file sheet", list(txn_sheets.keys()))
    txn_df = txn_sheets[t_sheet]

    detected_cols = auto_detect_txn_columns(txn_df.columns)
    id_guess = auto_detect_office_id_column(txn_df, master_df)
    missing_metric_cols = [k for k, v in detected_cols.items() if v is None]
    id_col_guess, id_score = id_guess if id_guess else (None, 0)
    need_manual_id = id_guess is None or id_score < 0.5

    if missing_metric_cols or need_manual_id:
        with manual_map_placeholder:
            st.warning("Some columns could not be auto-detected — please confirm below.")
            for canonical in missing_metric_cols:
                detected_cols[canonical] = st.selectbox(f"{canonical} column", list(txn_df.columns), key=f"map_{canonical}")
            options = list(txn_df.columns)
            idx = options.index(id_col_guess) if id_col_guess in options else 0
            id_col_guess = st.selectbox("Office ID column (transaction file)", options, index=idx, key="map_office_id")

    # ---- compute (always matched via Office ID) ----
    master = master_df.copy()
    master["Merge Key"] = normalize_key(master["Office ID"])

    txn = txn_df[[id_col_guess] + [detected_cols[c] for c in ["QR Count", "QR Amount", "Card Count", "Card Amount"]]].copy()
    txn.columns = ["Office ID Raw", "QR Count", "QR Amount", "Card Count", "Card Amount"]
    txn["Merge Key"] = normalize_key(txn["Office ID Raw"])
    for c in ["QR Count", "QR Amount", "Card Count", "Card Amount"]:
        txn[c] = clean_numeric(txn[c])
    txn = txn.groupby("Merge Key", as_index=False).agg(
        {"QR Count": "sum", "QR Amount": "sum", "Card Count": "sum", "Card Amount": "sum"}
    )

    merged = pd.merge(master, txn, on="Merge Key", how="left")
    for c in ["QR Count", "QR Amount", "Card Count", "Card Amount"]:
        merged[c] = merged[c].fillna(0)
    merged["Total Count"] = merged["QR Count"] + merged["Card Count"]
    merged["Total Amount"] = merged["QR Amount"] + merged["Card Amount"]
    is_active = (merged["Total Count"] >= MIN_TRANSACTIONS) | (merged["Total Amount"] > MIN_AMOUNT)
    merged["Activation Status"] = np.where(is_active, "ACTIVE", "INACTIVE")
    merged = merged.drop(columns=["Merge Key"])

    div_summary = build_division_summary(merged)
    image_bytes = build_division_summary_image(div_summary, period_str)
    excel_bytes = build_excel_report(merged, div_summary, period_str)

    st.markdown("---")
    st.image(image_bytes, use_container_width=True)
    st.download_button(
        "⬇️ Download Excel Report", data=excel_bytes,
        file_name=f"SBI_POS_Activation_Status_{today.strftime('%d.%m.%Y')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
