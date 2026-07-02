import os
import re
from datetime import date
import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

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


def build_division_summary(merged):
    div_summary = (
        merged.groupby("Division Name")["Activation Status"]
        .apply(lambda s: (s == "ACTIVE").sum())
        .reset_index(name="Active Count")
    )
    div_summary["Total Offices"] = merged.groupby("Division Name")["Office Name"].count().values
    div_summary["Inactive Count"] = div_summary["Total Offices"] - div_summary["Active Count"]
    div_summary["Activation %"] = round(div_summary["Active Count"] / div_summary["Total Offices"] * 100, 1)
    return div_summary.sort_values("Division Name").reset_index(drop=True)


def render_division_summary_html(div_summary, period_str):
    GREEN_BG, GREEN_TEXT = "#d9f2df", "#1e7d34"
    RED_BG, RED_TEXT = "#fbdada", "#b83232"
    HEADER_BG = "#2c3e50"
    TOTAL_BG = "#eceff1"

    rows_html = ""
    for _, row in div_summary.iterrows():
        active_pct = row["Activation %"]
        bg, text_color = (GREEN_BG, GREEN_TEXT) if active_pct >= 50 else (RED_BG, RED_TEXT)
        rows_html += f"""
        <tr style="background-color:{bg};">
            <td style="padding:9px 14px; border:1px solid #cfd8dc; text-align:left;">{row['Division Name']}</td>
            <td style="padding:9px 14px; border:1px solid #cfd8dc; text-align:center;">{int(row['Total Offices'])}</td>
            <td style="padding:9px 14px; border:1px solid #cfd8dc; text-align:center;">{int(row['Active Count'])}</td>
            <td style="padding:9px 14px; border:1px solid #cfd8dc; text-align:center;">{int(row['Inactive Count'])}</td>
            <td style="padding:9px 14px; border:1px solid #cfd8dc; text-align:center; font-weight:700; color:{text_color};">{active_pct}%</td>
        </tr>"""

    total_offices = int(div_summary["Total Offices"].sum())
    total_active = int(div_summary["Active Count"].sum())
    total_inactive = int(div_summary["Inactive Count"].sum())
    total_pct = round(total_active / total_offices * 100, 1) if total_offices else 0
    total_row = f"""
        <tr style="background-color:{TOTAL_BG}; font-weight:700;">
            <td style="padding:9px 14px; border:1px solid #b0bec5; text-align:left;">All Divisions</td>
            <td style="padding:9px 14px; border:1px solid #b0bec5; text-align:center;">{total_offices}</td>
            <td style="padding:9px 14px; border:1px solid #b0bec5; text-align:center;">{total_active}</td>
            <td style="padding:9px 14px; border:1px solid #b0bec5; text-align:center;">{total_inactive}</td>
            <td style="padding:9px 14px; border:1px solid #b0bec5; text-align:center;">{total_pct}%</td>
        </tr>"""

    html = f"""
    <div style="font-family:Arial, Helvetica, sans-serif;">
        <div style="font-size:18px; font-weight:700; margin-bottom:2px;">SBI POS Machine Activation Status — Division-wise Summary</div>
        <div style="font-size:14px; color:#555; margin-bottom:12px;">Period: {period_str}</div>
        <table style="border-collapse:collapse; width:100%; table-layout:fixed; font-size:14px;">
            <colgroup>
                <col style="width:40%;"><col style="width:15%;"><col style="width:15%;"><col style="width:15%;"><col style="width:15%;">
            </colgroup>
            <thead>
                <tr style="background-color:{HEADER_BG}; color:#ffffff;">
                    <th style="padding:10px 14px; border:1px solid #cfd8dc; text-align:left;">Division</th>
                    <th style="padding:10px 14px; border:1px solid #cfd8dc;">Total Offices</th>
                    <th style="padding:10px 14px; border:1px solid #cfd8dc;">Active</th>
                    <th style="padding:10px 14px; border:1px solid #cfd8dc;">Inactive</th>
                    <th style="padding:10px 14px; border:1px solid #cfd8dc;">Activation %</th>
                </tr>
            </thead>
            <tbody>
                {rows_html}
                {total_row}
            </tbody>
        </table>
        <div style="font-size:12px; color:#777; margin-top:8px;">
            <span style="background-color:{GREEN_BG}; padding:2px 8px; border-radius:3px; color:{GREEN_TEXT}; font-weight:700;">■</span> ≥ 50% activation
            &nbsp;&nbsp;
            <span style="background-color:{RED_BG}; padding:2px 8px; border-radius:3px; color:{RED_TEXT}; font-weight:700;">■</span> &lt; 50% activation
        </div>
    </div>
    """
    return html


def build_excel_report(merged, div_summary, unmapped, period_str):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book

        title_fmt = workbook.add_format({"bold": True, "font_size": 14})
        subtitle_fmt = workbook.add_format({"italic": True, "font_size": 10, "font_color": "#555555"})
        header_fmt = workbook.add_format({
            "bold": True, "bg_color": "#2c3e50", "font_color": "white",
            "border": 1, "align": "center", "valign": "vcenter",
        })
        green_fmt = workbook.add_format({"bg_color": "#d9f2df", "border": 1, "align": "center"})
        green_fmt_left = workbook.add_format({"bg_color": "#d9f2df", "border": 1, "align": "left"})
        red_fmt = workbook.add_format({"bg_color": "#fbdada", "border": 1, "align": "center"})
        red_fmt_left = workbook.add_format({"bg_color": "#fbdada", "border": 1, "align": "left"})
        total_fmt = workbook.add_format({"bold": True, "bg_color": "#eceff1", "border": 1, "align": "center"})
        total_fmt_left = workbook.add_format({"bold": True, "bg_color": "#eceff1", "border": 1, "align": "left"})
        pct_fmt = workbook.add_format({"num_format": "0.0\"%\""})

        ws = workbook.add_worksheet("Division Summary")
        ws.merge_range(0, 0, 0, 4, "SBI POS Machine Activation Status – Division-wise Summary", title_fmt)
        ws.merge_range(1, 0, 1, 4, f"Period: {period_str}", subtitle_fmt)

        headers = ["Division", "Total Offices", "Active", "Inactive", "Activation %"]
        for col, h in enumerate(headers):
            ws.write(3, col, h, header_fmt)

        row_idx = 4
        for _, row in div_summary.iterrows():
            is_green = row["Activation %"] >= 50
            fmt, fmt_left = (green_fmt, green_fmt_left) if is_green else (red_fmt, red_fmt_left)
            ws.write(row_idx, 0, row["Division Name"], fmt_left)
            ws.write(row_idx, 1, int(row["Total Offices"]), fmt)
            ws.write(row_idx, 2, int(row["Active Count"]), fmt)
            ws.write(row_idx, 3, int(row["Inactive Count"]), fmt)
            ws.write(row_idx, 4, row["Activation %"], fmt)
            row_idx += 1

        total_offices = int(div_summary["Total Offices"].sum())
        total_active = int(div_summary["Active Count"].sum())
        total_inactive = int(div_summary["Inactive Count"].sum())
        total_pct = round(total_active / total_offices * 100, 1) if total_offices else 0
        ws.write(row_idx, 0, "All Divisions", total_fmt_left)
        ws.write(row_idx, 1, total_offices, total_fmt)
        ws.write(row_idx, 2, total_active, total_fmt)
        ws.write(row_idx, 3, total_inactive, total_fmt)
        ws.write(row_idx, 4, total_pct, total_fmt)

        ws.set_column(0, 0, 34)
        ws.set_column(1, 4, 15)

        merged.to_excel(writer, sheet_name="Full Detail", index=False)
        writer.sheets["Full Detail"].set_column(0, len(merged.columns) - 1, 18)

        if len(unmapped) > 0:
            unmapped.to_excel(writer, sheet_name="Unmapped Offices", index=False)
            writer.sheets["Unmapped Offices"].set_column(0, len(unmapped.columns) - 1, 18)

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
# Results — a single color-coded division-wise summary table
# ---------------------------------------------------------------------------
if "result" in st.session_state:
    merged = st.session_state["result"]
    unmapped = st.session_state["unmapped"]

    today = date.today()
    period_start = today.replace(day=1)
    period_str = f"{period_start.strftime('%d.%m.%Y')} to {today.strftime('%d.%m.%Y')}"

    div_summary = build_division_summary(merged)

    st.markdown("---")
    st.markdown(render_division_summary_html(div_summary, period_str), unsafe_allow_html=True)

    excel_bytes = build_excel_report(merged, div_summary, unmapped, period_str)
    st.download_button(
        "⬇️ Download Excel Report",
        data=excel_bytes,
        file_name=f"SBI_POS_Activation_Status_{today.strftime('%d.%m.%Y')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    with st.expander("🔎 Full office-level detail / unmapped offices (optional)", expanded=False):
        if len(unmapped) > 0:
            st.warning(
                f"{len(unmapped)} office(s) found in the transaction file with no matching "
                "entry in the master data — excluded from the summary above. Check for "
                "spelling/ID mismatches."
            )
            st.dataframe(unmapped, use_container_width=True, hide_index=True)
        st.dataframe(merged, use_container_width=True, hide_index=True)
else:
    st.info("👆 Upload the transaction file above — results will appear automatically.")
