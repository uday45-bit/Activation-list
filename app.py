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

POS_SUPPLY_FILENAME = "pos_supply_status.xlsx"
POS_SUPPLY_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), POS_SUPPLY_FILENAME)

MASTER_COLUMN_MAP = {
    "division-office-id": "Division ID",
    "division-office-name": "Division Name",
    "office-name": "Office Name",
    "office-id": "Office ID",
    "office-type-code": "Office Type",
}

# Office types included in the activation-report universe.
INCLUDED_OFFICE_TYPES = ["SPO", "HPO", "GPO"]

# Specific Hyderabad Sorting Division offices included by exception even
# though their office-type-code isn't SPO/HPO/GPO (they do carry SBI POS
# machines): PBC Autonagar, SPC Counter, BPC Jeedimetla, Secunderabad RSTMO,
# Hyderabad Deccan RSTMO.
SPECIAL_INCLUDE_OFFICE_IDS = ["30260015", "30260017", "30260013", "30680009", "30680008"]

# Column-header aliases for the bundled POS supply/whitelisting reference file.
POS_SUPPLY_SUMMARY_ALIASES = {
    "Division Name": ["division", "divisions", "division name"],
    "Machines Supplied": ["sbi pos machines supplied", "machines supplied", "pos machines supplied"],
    "Not Whitelisted Count": ["offices not whitelisted", "not whitelisted"],
    "Not Supplied Count": ["offices not supplied with sbi pos", "offices not supplied", "not supplied"],
}
WHITELIST_DETAIL_ALIASES = {
    "Division Name": ["divisions", "division", "division name"],
    "Office ID": ["office id"],
    "Office Name": ["office name"],
}
NOT_SUPPLIED_DETAIL_ALIASES = {
    "Division Name": ["division name", "divisions", "division"],
    "Office Name": ["office name"],
}

# Known transaction-file header variants -> canonical field.
TXN_COLUMN_ALIASES = {
    "Card Count": ["sbipos card (cnt)", "sbi pos card(count)", "sbi pos card (count)", "card count"],
    "Card Amount": ["sbipos card (amt)", "sbi pos card(amount)", "sbi pos card (amount)", "card amount"],
    "QR Count": ["sbipos bharatqr (cnt)", "sbi pos qr(count)", "sbi pos qr (count)", "bharatqr count", "qr count"],
    "QR Amount": ["sbipos bharatqr (amt)", "sbi pos qr(amount)", "sbi pos qr (amount)", "bharatqr amount", "qr amount"],
}

GREEN_BG, GREEN_TEXT = "#d9f2df", "#1e7d34"
PINK_BG, PINK_TEXT = "#fce4ec", "#ad1457"
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
    df = df[df["Office Type"].isin(INCLUDED_OFFICE_TYPES) | df["Office ID"].isin(SPECIAL_INCLUDE_OFFICE_IDS)]
    df = df[~df["Division Name"].str.lower().isin(["", "undefined", "nan"])]
    df = df.drop_duplicates(subset="Office ID")
    return df


@st.cache_data
def load_master_data(path):
    return clean_master_df(pd.read_csv(path))


def rename_by_aliases(df, aliases_map):
    """Rename df's columns to canonical names using header aliases (matched
    case/space-insensitively). Raises if a required canonical field is missing."""
    normalized = {normalize_header(c): c for c in df.columns}
    rename = {}
    for canonical, aliases in aliases_map.items():
        for alias in aliases:
            if alias in normalized:
                rename[normalized[alias]] = canonical
                break
    df = df.rename(columns=rename)
    missing = [c for c in aliases_map if c not in df.columns]
    if missing:
        raise ValueError(f"Could not find column(s) for: {missing}")
    return df[list(aliases_map.keys())]


def read_pos_summary_sheet(xls, sheet_name):
    """The monthly summary sheet sometimes has a blank leading row above the
    real header — detect and skip it if so."""
    raw = xls.parse(sheet_name, header=0)
    first_col = str(raw.columns[0])
    if first_col.startswith("Unnamed") or first_col.strip() == "":
        raw = xls.parse(sheet_name, header=1)
    return raw


@st.cache_data
def load_pos_supply_data(path):
    """Loads the bundled POS supply/whitelisting reference file: a division-
    wise summary (machines supplied, not-whitelisted/not-supplied counts) plus
    two detail sheets. Sheet 1 is read positionally (its name changes monthly,
    e.g. 'July 2026'); the detail sheets are matched by keyword, falling back
    to position 2 and 3."""
    xls = pd.ExcelFile(path, engine="openpyxl")
    sheets = xls.sheet_names
    if len(sheets) < 3:
        raise ValueError(f"{POS_SUPPLY_FILENAME} must have 3 sheets: monthly summary, whitelisting pending, not supplied offices.")

    summary_sheet = sheets[0]
    whitelist_sheet = next((s for s in sheets[1:] if "whitelist" in s.lower()), sheets[1])
    notsupplied_sheet = next((s for s in sheets[1:] if "not supplied" in s.lower() or "supply" in s.lower()), sheets[2])

    summary = rename_by_aliases(read_pos_summary_sheet(xls, summary_sheet), POS_SUPPLY_SUMMARY_ALIASES)
    summary["Division Name"] = summary["Division Name"].astype(str).str.strip()
    summary = summary[~summary["Division Name"].str.lower().isin(["", "nan", "all divisions", "total"])]
    for c in ["Machines Supplied", "Not Whitelisted Count", "Not Supplied Count"]:
        summary[c] = clean_numeric(summary[c]).astype(int)

    whitelist = rename_by_aliases(xls.parse(whitelist_sheet), WHITELIST_DETAIL_ALIASES)
    whitelist["Division Name"] = whitelist["Division Name"].astype(str).str.strip()
    whitelist["Office Name"] = whitelist["Office Name"].astype(str).str.strip()
    whitelist["Office ID"] = (
        whitelist["Office ID"].astype(str).str.replace(r"\.0$", "", regex=True).replace("nan", "").str.strip()
    )

    notsupplied = rename_by_aliases(xls.parse(notsupplied_sheet), NOT_SUPPLIED_DETAIL_ALIASES)
    notsupplied["Division Name"] = notsupplied["Division Name"].astype(str).str.strip()
    notsupplied["Office Name"] = notsupplied["Office Name"].astype(str).str.strip()

    return summary, whitelist, notsupplied


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


def build_division_summary(merged, supply_summary):
    """Division-wise summary expressed in terms of SBI POS *machines* (not
    raw office counts) — a division's 'SBI POS Machines Issued' comes from
    the POS supply data, since machines-per-office isn't always 1:1 (e.g.
    Hyderabad GPO: 1 office, 7 machines). Active/Inactive machine counts are
    the office-level active ratio applied to the machines-issued count, so
    if a division's one office is active, all its machines count as active."""
    active_offices = merged.groupby("Division Name")["Activation Status"].apply(lambda s: (s == "ACTIVE").sum())
    office_count = merged.groupby("Division Name")["Office Name"].count()
    division_id = merged.groupby("Division Name")["Division ID"].first()

    base = pd.DataFrame({
        "Division Name": office_count.index,
        "Division ID": division_id.values,
        "Active Offices": active_offices.values,
        "Office Count": office_count.values,
    })

    supply = supply_summary[["Division Name", "Machines Supplied"]]
    base = pd.merge(base, supply, on="Division Name", how="left")
    base["Machines Supplied"] = base["Machines Supplied"].fillna(base["Office Count"]).astype(int)

    active_ratio = base["Active Offices"] / base["Office Count"]
    base["Active Count"] = (active_ratio * base["Machines Supplied"]).round().astype(int)
    base["Inactive Count"] = base["Machines Supplied"] - base["Active Count"]
    base["Activation %"] = round(base["Active Count"] / base["Machines Supplied"] * 100, 1)
    base["Total Offices"] = base["Machines Supplied"]  # column kept for "SBI POS Machines Issued" display

    return base[["Division Name", "Division ID", "Total Offices", "Active Count", "Inactive Count", "Activation %"]]


def build_pos_supply_overview(master_df, supply_summary):
    """Division-wise: devices supplied, office count (SPO/HPO/GPO + the 5
    named exceptions), not-supplied count, not-whitelisted count, and the
    excess of devices supplied over office count."""
    office_count = master_df.groupby("Division Name").size().reset_index(name="No. of Offices")
    out = pd.merge(office_count, supply_summary, on="Division Name", how="outer")
    out["No. of Offices"] = out["No. of Offices"].fillna(0).astype(int)
    for c in ["Machines Supplied", "Not Whitelisted Count", "Not Supplied Count"]:
        out[c] = out[c].fillna(0).astype(int)
    out["Excess"] = out["Machines Supplied"] - out["No. of Offices"]
    out = out.rename(columns={
        "Machines Supplied": "Devices Supplied",
        "Not Supplied Count": "Not Supplied",
        "Not Whitelisted Count": "Not Whitelisted",
    })
    return out[["Division Name", "Devices Supplied", "No. of Offices", "Not Supplied", "Not Whitelisted", "Excess"]].sort_values("Division Name").reset_index(drop=True)


def classify_color(pct, criteria):
    """3-tier color classification for a division's activation %, relative
    to a user-chosen criteria: green if at/above criteria, pink if within 10
    points below, red if more than 10 points below."""
    if pct >= criteria:
        return GREEN_BG, GREEN_TEXT
    if pct >= criteria - 10:
        return PINK_BG, PINK_TEXT
    return RED_BG, RED_TEXT


def build_title_lines(period_str, criteria):
    line1 = f"SBI POS Machine Activation Status (Period: {period_str})"
    line2 = (
        f"[Monthly Criteria for activation: No. of Transactions >= {MIN_TRANSACTIONS} / "
        f"Transaction amt >= Rs.{MIN_AMOUNT}/- | Division activation criteria: {criteria}%]"
    )
    return line1, line2


def build_division_summary_image(div_summary, period_str, criteria):
    columns = ["Office ID", "Office Name", "SBI POS Machines Issued", "Active", "Inactive", "% Activation"]
    div_summary = div_summary.sort_values("Activation %", ascending=False).reset_index(drop=True)
    total_offices = int(div_summary["Total Offices"].sum())
    total_active = int(div_summary["Active Count"].sum())
    total_inactive = int(div_summary["Inactive Count"].sum())
    total_pct = round(total_active / total_offices * 100, 1) if total_offices else 0

    rows, cell_colors = [], []
    for _, r in div_summary.iterrows():
        rows.append([r["Division ID"], r["Division Name"], int(r["Total Offices"]), int(r["Active Count"]), int(r["Inactive Count"]), f'{r["Activation %"]}%'])
        bg, _ = classify_color(r["Activation %"], criteria)
        cell_colors.append([bg] * 6)
    rows.append(["Total", "", total_offices, total_active, total_inactive, f"{total_pct}%"])
    cell_colors.append([TOTAL_BG] * 6)

    n_rows = len(rows) + 1  # + header row
    fig_h = 0.9 + 0.42 * n_rows
    fig, ax = plt.subplots(figsize=(11.5, fig_h))
    ax.axis("off")
    line1, line2 = build_title_lines(period_str, criteria)
    ax.set_title(f"{line1}\n{line2}", fontsize=12, fontweight="bold", loc="left", pad=20)

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


def build_excel_report(merged, div_summary, supply_overview, period_str, criteria):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book

        title_fmt = workbook.add_format({"bold": True, "font_size": 13, "text_wrap": True, "valign": "vcenter"})
        header_fmt = workbook.add_format({
            "bold": True, "bg_color": HEADER_BG, "font_color": "white",
            "border": 1, "align": "center", "valign": "vcenter",
        })
        green_fmt = workbook.add_format({"bg_color": GREEN_BG, "border": 1, "align": "center"})
        green_fmt_left = workbook.add_format({"bg_color": GREEN_BG, "border": 1, "align": "left"})
        pink_fmt = workbook.add_format({"bg_color": PINK_BG, "border": 1, "align": "center"})
        pink_fmt_left = workbook.add_format({"bg_color": PINK_BG, "border": 1, "align": "left"})
        red_fmt = workbook.add_format({"bg_color": RED_BG, "border": 1, "align": "center"})
        red_fmt_left = workbook.add_format({"bg_color": RED_BG, "border": 1, "align": "left"})
        total_fmt = workbook.add_format({"bold": True, "bg_color": TOTAL_BG, "border": 1, "align": "center"})
        total_fmt_left = workbook.add_format({"bold": True, "bg_color": TOTAL_BG, "border": 1, "align": "left"})
        plain_fmt = workbook.add_format({"border": 1, "align": "center"})
        plain_fmt_left = workbook.add_format({"border": 1, "align": "left"})

        color_fmt_map = {
            GREEN_BG: (green_fmt, green_fmt_left),
            PINK_BG: (pink_fmt, pink_fmt_left),
            RED_BG: (red_fmt, red_fmt_left),
        }

        # ---- Sheet 1: Division Summary — descending %, 3-tier color, 2-line title in one cell ----
        div_summary = div_summary.sort_values("Activation %", ascending=False).reset_index(drop=True)
        ws = workbook.add_worksheet("Division Summary")
        line1, line2 = build_title_lines(period_str, criteria)
        ws.merge_range(0, 0, 0, 5, f"{line1}\n{line2}", title_fmt)
        ws.set_row(0, 34)

        headers = ["Office ID", "Office Name", "SBI POS Machines Issued", "Active", "Inactive", "% Activation"]
        for col, h in enumerate(headers):
            ws.write(1, col, h, header_fmt)

        row_idx = 2
        for _, row in div_summary.iterrows():
            bg, _ = classify_color(row["Activation %"], criteria)
            fmt, fmt_left = color_fmt_map[bg]
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

        # ---- Sheet 2: Active Offices (Office Name + Division Name, not raw ID) ----
        active_df = (
            merged[merged["Activation Status"] == "ACTIVE"]
            [["Division Name", "Office Name", "Total Count", "Total Amount"]]
            .sort_values(["Division Name", "Office Name"])
            .reset_index(drop=True)
        )
        active_df.to_excel(writer, sheet_name="Active Offices", index=False)
        ws2 = writer.sheets["Active Offices"]
        ws2.set_column(0, 1, 30)
        ws2.set_column(2, 3, 15)

        # ---- Sheet 3: POS Supply Overview (division-wise) ----
        ws3 = workbook.add_worksheet("POS Supply Overview")
        ws3.merge_range(0, 0, 0, 5, "SBI POS Machines — Supply, Whitelisting & Excess (Division-wise)", title_fmt)
        headers3 = ["Division", "Devices Supplied", "No. of Offices", "Not Supplied", "Not Whitelisted", "Excess"]
        for col, h in enumerate(headers3):
            ws3.write(1, col, h, header_fmt)
        row_idx = 2
        for _, row in supply_overview.iterrows():
            fmt, fmt_left = (red_fmt, red_fmt_left) if row["Excess"] > 0 else (plain_fmt, plain_fmt_left)
            ws3.write(row_idx, 0, row["Division Name"], fmt_left)
            ws3.write(row_idx, 1, int(row["Devices Supplied"]), fmt)
            ws3.write(row_idx, 2, int(row["No. of Offices"]), fmt)
            ws3.write(row_idx, 3, int(row["Not Supplied"]), fmt)
            ws3.write(row_idx, 4, int(row["Not Whitelisted"]), fmt)
            ws3.write(row_idx, 5, int(row["Excess"]), fmt)
            row_idx += 1
        ws3.write(row_idx, 0, "Total", total_fmt_left)
        ws3.write(row_idx, 1, int(supply_overview["Devices Supplied"].sum()), total_fmt)
        ws3.write(row_idx, 2, int(supply_overview["No. of Offices"].sum()), total_fmt)
        ws3.write(row_idx, 3, int(supply_overview["Not Supplied"].sum()), total_fmt)
        ws3.write(row_idx, 4, int(supply_overview["Not Whitelisted"].sum()), total_fmt)
        ws3.write(row_idx, 5, int(supply_overview["Excess"].sum()), total_fmt)
        ws3.set_column(0, 0, 32)
        ws3.set_column(1, 5, 16)

    return output.getvalue()


# ---------------------------------------------------------------------------
# UI: exactly two boxes (Update Reference Data, Upload Transaction File),
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

try:
    default_supply_summary, default_whitelist, default_notsupplied = load_pos_supply_data(POS_SUPPLY_PATH)
except FileNotFoundError:
    st.error(f"`{POS_SUPPLY_FILENAME}` was not found next to app.py. Add it to the repo root and redeploy.")
    st.stop()
except ValueError as e:
    st.error(str(e))
    st.stop()

master_df = st.session_state.get("master_override_df", default_master)
using_master_override = "master_override_df" in st.session_state
supply_summary_raw, whitelist_detail, notsupplied_detail = st.session_state.get(
    "supply_override", (default_supply_summary, default_whitelist, default_notsupplied)
)
using_supply_override = "supply_override" in st.session_state

# ---- Box 1: Update Reference Data (master offices + POS supply/whitelisting) ----
with st.container(border=True):
    st.markdown("#### 🔄 Update Reference Data")
    st.caption(
        f"Master data: **{len(master_df)} offices** across **{master_df['Division Name'].nunique()} divisions**"
        + (" (session override active)." if using_master_override else f", from `{MASTER_DATA_FILENAME}`.")
    )
    st.caption(
        f"POS supply data: **{int(supply_summary_raw['Machines Supplied'].sum())} machines supplied**, "
        f"**{len(whitelist_detail)} pending whitelisting**, **{len(notsupplied_detail)} offices not supplied**"
        + (" (session override active)." if using_supply_override else f", from `{POS_SUPPLY_FILENAME}`.")
    )

    override_upload = st.file_uploader(
        "Replace master office file", type=["csv", "xlsx", "xls"], key="master_upload",
    )
    if override_upload is not None:
        try:
            raw = list(read_any(override_upload).values())[0]
            raw = clean_master_df(raw)
            st.session_state["master_override_df"] = raw
            st.success(f"Master override loaded: {len(raw)} offices. Re-upload the transaction file below to apply it.")
        except ValueError as e:
            st.error(str(e))
        except Exception as e:
            st.error(f"Could not read file: {e}")
    if using_master_override and st.button("Revert to bundled master data"):
        del st.session_state["master_override_df"]
        st.rerun()

    supply_upload = st.file_uploader(
        "Replace POS supply & whitelisting file", type=["xlsx", "xls"], key="supply_upload",
    )
    if supply_upload is not None:
        try:
            tmp_path = f"/tmp/{supply_upload.name}"
            with open(tmp_path, "wb") as f:
                f.write(supply_upload.getbuffer())
            new_summary, new_whitelist, new_notsupplied = load_pos_supply_data(tmp_path)
            st.session_state["supply_override"] = (new_summary, new_whitelist, new_notsupplied)
            st.success("POS supply data override loaded. Re-upload the transaction file below to apply it.")
        except ValueError as e:
            st.error(str(e))
        except Exception as e:
            st.error(f"Could not read file: {e}")
    if using_supply_override and st.button("Revert to bundled POS supply data"):
        del st.session_state["supply_override"]
        st.rerun()

# ---- Box 2: Upload Transaction File ----
with st.container(border=True):
    st.markdown("#### 📤 Upload Transaction File")
    pcol, ccol = st.columns(2)
    with pcol:
        period_choice = st.radio(
            "Report period", ["Current month (till date)", "Previous month (complete)"], horizontal=True
        )
    with ccol:
        criteria = st.number_input(
            "Division activation criteria (%)", min_value=0, max_value=100, value=50, step=1,
            help="Divisions at/above this % are green, up to 10 points below are pink, more than 10 points below are red.",
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

    div_summary = build_division_summary(merged, supply_summary_raw)
    supply_overview = build_pos_supply_overview(master_df, supply_summary_raw)

    image_bytes = build_division_summary_image(div_summary, period_str, criteria)
    excel_bytes = build_excel_report(merged, div_summary, supply_overview, period_str, criteria)

    st.markdown("---")
    st.image(image_bytes, use_container_width=True)
    st.download_button(
        "⬇️ Download Excel Report", data=excel_bytes,
        file_name=f"SBI_POS_Activation_Status_{today.strftime('%d.%m.%Y')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
