import re

import pandas as pd
import streamlit as st
from datetime import date

import gspread
from google.oauth2.service_account import Credentials

# =========================================================================
# CONFIG
#
# This is the MARKETING MANAGER view app. It shares the same Google Sheet
# and the same "Customer List.xlsx" as the data-entry app (app.py) and the
# manager app. It has no Pond Details editor and no Harvest Details form.
#
# It still does NOT include the Zone-Wise Harvest breakdown, the Running
# List, or the Species-wise Pond Summary sections that the manager app
# has — it does now include Last Visit Date Report at the bottom, same as
# the manager app.
#
# This app is VIEW-ONLY for Technical Officers & Marketing Managers — it
# performs NO writes to either Google Sheet. Sales Details and All Harvest
# Details both have no delete/recycle-bin control; a row already flagged
# as removed/settled by the full manager app (Settle = 'Yes', or Harvest
# Status = 'H') stays hidden here too, but nothing on this page can set
# either flag itself. Everything here is read-only:
#   1) "📋 Enter Customer Details" — Customer / Farm selection (used only
#      to choose which farm's records to view)
#   2) "📊 All Saved Records" + "🟦 Pond Layout" — a live, read-only view
#      of that farm's data straight from the Google Sheet.
#   3) "🧾 Sales Details" — that farm's sales line items + feed order
#      status boxes.
#   4) "🌾 All Harvest Details" — every harvested row across all
#      customers/farms/ponds.
#   5) "🗓️ Last Visit Date Report" — every Customer + Farm's Status
#      (Running / FULL H), latest data-entry date and days since.
#   6) "🤖 Chat with Anji" — NEW. A read-only AI chatbot section at the very
#      bottom that answers questions (DOC range customers, last feed order
#      of a farm, who harvested this week, a farm's current situation, who
#      has not purchased a product such as Osmin) straight from the same
#      Google Sheets data. It performs NO writes either.
#
# Deploy this file as its own Streamlit app (its own URL/link) so the
# marketing manager gets a separate link from the data-entry app and the
# full manager app.
# =========================================================================
st.set_page_config(page_title="Shrimp FarmFlow - KMN (Marketing Manager)", layout="wide", page_icon="📊")

CUSTOMER_FILE = "Customer List.xlsx"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]

# Second Google Sheet — Sales Details. Separate spreadsheet from the
# water-quality WaterQualityData sheet above; the same service account
# must also be shared (as Viewer or Editor) on this sheet.
SALES_SHEET_ID = "1S3csAE-E_hN8vstuHR0KkeAN7yCVQTFe4AkEVlw4vQw"

# Must match app.py's COLUMN_ORDER exactly, since all three apps read/
# write the same Google Sheet. Includes the second Harvest slot (Harvest
# Date 2 / Harvest Type 2 / Harvest KG 2 / Harvest ABW 2), and the
# "Expect Harvest (KG)" / "Survival QTY" auto-calculated columns.
COLUMN_ORDER = [
    "Timestamp", "Customer", "Farm Name with Code", "Zone", "Area",
    "Pond Number", "Date", "Species Culture", "Cycle Type",
    "DOC", "Density", "Feed Per Day", "ABW",
    "Expect Harvest (KG)", "Survival QTY",
    "Issues", "Water Color", "Grade", "Remark", "Technician",
    "Harvest Date", "Harvest Type", "Harvest KG", "Harvest ABW",
    "Harvest Date 2", "Harvest Type 2", "Harvest KG 2", "Harvest ABW 2",
    "Deleted",
]

# Expected columns in the Sales Details Google Sheet. "Settle" holds the
# persisted checkbox state from the "Financial Status" tick below (blank
# or "Yes") — it's created automatically in the sheet the first time a row
# is ticked if it doesn't already exist there.
SALES_COLUMN_ORDER = [
    "Date", "Item No.", "Item Description", "Customer Code",
    "Customer Name", "Quantity", "Sales Amt", "Settle",
]

# =========================================================================
# STYLE (kept visually consistent with the other two apps)
# =========================================================================
st.markdown("""
<style>
ul[role="listbox"], div[role="listbox"] {
    width: max-content !important;
    min-width: 220px !important;
    max-width: 92vw !important;
}
[role="option"] {
    width: auto !important;
    white-space: normal !important;
    overflow: visible !important;
    text-overflow: unset !important;
    word-break: break-word !important;
}
[role="option"] * {
    white-space: normal !important;
    overflow: visible !important;
    text-overflow: unset !important;
}

@media (max-width: 700px) {
    div[data-testid="stHorizontalBlock"] {
        flex-direction: column !important;
    }
    div[data-testid="stHorizontalBlock"] > div[data-testid="column"] {
        width: 100% !important;
        min-width: 100% !important;
    }
}
</style>
""", unsafe_allow_html=True)

st.markdown("<h1 style='text-align: center;'>Shrimp FarmFlow - KMN</h1>",
            unsafe_allow_html=True)
st.subheader("KMN Aqua Services — Marketing Manager & Technician View")
st.markdown("---")

# =========================================================================
# GOOGLE SHEETS BACKEND. Entirely read-only — this app never writes to
# either Google Sheet. get_or_create_column() is kept only because
# get_worksheet()/get_sales_worksheet() below are shared helpers with the
# other apps in this family; nothing in this file calls it anymore.
# =========================================================================
def _gsheet_configured():
    return "gcp_service_account" in st.secrets and "gsheet" in st.secrets and "sheet_id" in st.secrets["gsheet"]

@st.cache_resource(show_spinner=False)
def get_worksheet():
    creds_dict = dict(st.secrets["gcp_service_account"])
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    client = gspread.authorize(creds)
    sheet_id = st.secrets["gsheet"]["sheet_id"]
    worksheet_name = st.secrets["gsheet"].get("worksheet_name", "WaterQualityData")
    sh = client.open_by_key(sheet_id)
    return sh.worksheet(worksheet_name)

@st.cache_resource(show_spinner=False)
def get_sales_worksheet():
    """Separate spreadsheet (Sales Details) — same service account creds,
    different spreadsheet key. Worksheet/tab name can be overridden via
    st.secrets["gsheet"]["sales_worksheet_name"] (defaults to the first
    sheet/tab in the spreadsheet if not set)."""
    creds_dict = dict(st.secrets["gcp_service_account"])
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    client = gspread.authorize(creds)
    sh = client.open_by_key(SALES_SHEET_ID)
    worksheet_name = st.secrets.get("gsheet", {}).get("sales_worksheet_name", "")
    if worksheet_name:
        return sh.worksheet(worksheet_name)
    return sh.sheet1

def get_or_create_column(ws, header_name):
    """Returns the 1-based column number of the given header in ws,
    creating that header (in the first empty column) if it isn't there
    yet. Used for 'Settle' (Sales Details sheet) and 'Harvest Status' /
    'Harvest Status 2' (main WaterQualityData sheet). Expands the sheet's
    column count first if the new header would land past the sheet's
    current grid width — writing past the grid is what raises gspread's
    APIError, since the underlying Sheets API rejects it."""
    headers = ws.row_values(1)
    if header_name in headers:
        return headers.index(header_name) + 1
    new_col_idx = len(headers) + 1
    if new_col_idx > ws.col_count:
        ws.add_cols(new_col_idx - ws.col_count)
    ws.update_cell(1, new_col_idx, header_name)
    return new_col_idx

def load_data():
    """Always reads fresh from the Google Sheet (no caching), so the
    marketing manager always sees the latest saved records — including
    anything just added by the data-entry app or edited directly in the
    Sheet. Soft-deleted rows (Deleted = Yes) are filtered out, same as the
    data-entry app.

    NOTE: this app never writes 'Harvest Status' / 'Harvest Status 2' —
    there is no delete/recycle-bin control on "All Harvest Details" here
    (view-only for Technical Officers & Marketing Managers). The two
    columns are still read (and still respected if a value was set by the
    full manager app elsewhere) so a harvest event hidden there is also
    hidden from this app's All Harvest Details table — but nothing here
    can set that flag itself."""
    ws = get_worksheet()
    records = ws.get_all_records()
    df = pd.DataFrame(records)
    for c in COLUMN_ORDER:
        if c not in df.columns:
            df[c] = ""
    if "Harvest Status" not in df.columns:
        df["Harvest Status"] = ""
    # "Harvest Status 2" mirrors "Harvest Status" above but flags the 2nd
    # harvest slot (Harvest Date 2 / Harvest Type 2) independently — needed
    # because All Harvest Details now shows the 1st and 2nd harvest slots
    # of the same saved record as two separate rows, and each needs its
    # own recycle-bin flag so removing one slot never touches the other.
    if "Harvest Status 2" not in df.columns:
        df["Harvest Status 2"] = ""
    # "Harvest Submitted Date" already exists as its own column in the
    # Sheet (outside COLUMN_ORDER, same as "Harvest Status") — kept here
    # so the All Harvest Details table below can offer it as a Sort by
    # option instead of it getting dropped like any other unlisted column.
    if "Harvest Submitted Date" not in df.columns:
        df["Harvest Submitted Date"] = ""
    # "WQ Special Cases" already exists as its own column in the Sheet
    # (outside COLUMN_ORDER, same pattern as above) — kept here so the
    # Pond Layout section below can show a sad-face icon/label on a pond
    # whose latest record has text in this column.
    if "WQ Special Cases" not in df.columns:
        df["WQ Special Cases"] = ""
    # "Notes" already exists as the LAST column in the Sheet (outside
    # COLUMN_ORDER, same pattern as above) — kept here so the Pond Layout
    # section below can show each pond's latest Notes above its box.
    if "Notes" not in df.columns:
        df["Notes"] = ""
    if len(df) > 0:
        df = df[COLUMN_ORDER + ["Harvest Status", "Harvest Status 2", "Harvest Submitted Date", "WQ Special Cases", "Notes"]]
    df = df.astype(str).replace("nan", "")
    if "Deleted" in df.columns:
        is_deleted = df["Deleted"].astype(str).str.strip().str.lower().isin(["yes", "true", "1"])
        df = df[~is_deleted].reset_index(drop=True)
    return df

def load_sales_data():
    """Always reads fresh from the Sales Details Google Sheet (no caching),
    same pattern as load_data() above. Also attaches each row's real sheet
    row number (_RowNumber) so a tick in the UI can be written back to the
    correct cell in the 'Settle' column."""
    ws = get_sales_worksheet()
    records = ws.get_all_records()
    df = pd.DataFrame(records)
    for c in SALES_COLUMN_ORDER:
        if c not in df.columns:
            df[c] = ""
    df["_RowNumber"] = range(2, len(df) + 2)
    if len(df) > 0:
        df = df[["_RowNumber"] + SALES_COLUMN_ORDER]
    return df

if not _gsheet_configured():
    st.error("❌ Google Sheets is not configured yet. This app needs the same "
             "`.streamlit/secrets.toml` (the `[gcp_service_account]` and `[gsheet]` sections) "
             "used by the data-entry app.")
    st.stop()

try:
    get_worksheet()
except Exception as e:
    st.error(f"❌ Could not connect to the Google Sheet. Check your secrets and sharing settings.\n\n{e}")
    st.stop()

# =========================================================================
# LOAD CUSTOMER LIST (for the Customer / Farm selectors)
# =========================================================================
@st.cache_data
def load_customer_data():
    return pd.read_excel(CUSTOMER_FILE)

try:
    customer_df = load_customer_data()
except Exception as e:
    st.error(f"❌ Could not load '{CUSTOMER_FILE}'. Make sure it's in the app folder. ({e})")
    st.stop()

REQUIRED_COLS = ["Customer Name", "Farm Name with Code", "Zone", "Area"]
missing_cols = [c for c in REQUIRED_COLS if c not in customer_df.columns]
if missing_cols:
    st.error(f"❌ 'Customer List.xlsx' is missing required column(s): {', '.join(missing_cols)}")
    st.stop()

for _col in REQUIRED_COLS:
    customer_df[_col] = customer_df[_col].apply(
        lambda v: "" if pd.isna(v) else (str(int(v)) if isinstance(v, float) and v.is_integer() else str(v))
    )

all_customers = sorted(customer_df["Customer Name"].replace("", pd.NA).dropna().unique().tolist())

# =========================================================================
# STEP 1: "Enter Customer Details" — Customer / Farm selection.
# For the marketing manager this is just a filter to choose whose records
# to view — there is no data entry after this, only the read-only tables
# below.
# =========================================================================
st.subheader("📋 Enter Customer Details")

col1, col2 = st.columns(2)
with col1:
    customer = st.selectbox("Customer Name *", all_customers, key="customer_select")

farm_options = sorted(
    customer_df.loc[customer_df["Customer Name"] == customer, "Farm Name with Code"]
    .dropna().unique().tolist()
)
if not farm_options:
    farm_options = ["-- No farms found for this customer --"]

with col2:
    farm = st.selectbox("Farm Name with Code *", farm_options, key=f"farm_select_{customer}")

farm_row_match = customer_df[
    (customer_df["Customer Name"] == customer) & (customer_df["Farm Name with Code"] == farm)
]
if len(farm_row_match) > 0 and "Marketing Manager" in customer_df.columns:
    mm = farm_row_match.iloc[0].get("Marketing Manager", "")
    if str(mm).strip():
        st.caption(f"Marketing Manager: {mm}")

# The Customer Code / Customer ID for the selected farm (used below to
# filter the Sales Details sheet). "Customer List.xlsx" may store this
# under any of a few likely column names — the first one that exists and
# has a non-blank value for this row wins.
_CUSTOMER_CODE_COLUMN_CANDIDATES = [
    "Customer Code", "Customer ID", "Customer Code with Code", "Code", "Cust Code",
]
selected_customer_code = ""
if len(farm_row_match) > 0:
    for _cand in _CUSTOMER_CODE_COLUMN_CANDIDATES:
        if _cand in customer_df.columns:
            _val = str(farm_row_match.iloc[0].get(_cand, "")).strip()
            if _val and _val.lower() != "nan":
                selected_customer_code = _val
                break

# =========================================================================
# ALL SAVED RECORDS — read-only, straight from the Google Sheet, for the
# selected Customer + Farm across every pond.
# =========================================================================
st.markdown("---")
st.markdown(f"#### 📊 All Saved Records — {farm}")

if st.button("🔄 Refresh"):
    st.rerun()

# Hidden by default — tick this to show the full "All Saved Records"
# table below. The totals / Pond Layout sections further down are
# unaffected and still show either way.
show_all_saved_records = st.checkbox(
    "Show All Saved Records table", value=False, key="show_all_saved_records"
)

df_farm_summary = load_data()
_farm_required = {"Customer", "Farm Name with Code"}
if len(df_farm_summary) > 0 and _farm_required.issubset(df_farm_summary.columns):
    df_farm_summary = df_farm_summary[
        (df_farm_summary["Customer"] == customer) & (df_farm_summary["Farm Name with Code"] == farm)
    ].copy()
else:
    df_farm_summary = pd.DataFrame(columns=COLUMN_ORDER)

# Initialized here (before the branch below) so both are always defined —
# even when this farm has no saved records yet — since the Sales Details
# section further down the page reads total_density_by_species to compute
# the NANAMI/EGO feed "Do not exceed" limits.
total_expect_harvest_kg = None
total_density_by_species = {}
# Same as total_density_by_species but INCLUDES Full H ponds. Used only
# for the NANAMI / EGO feed "Do not exceed" limits in Sales Details.
total_density_for_feed_limit_by_species = {}

if len(df_farm_summary) > 0:
    if "Date" in df_farm_summary.columns:
        df_farm_summary["_ParsedDate"] = pd.to_datetime(df_farm_summary["Date"], errors="coerce")

        # Latest saved record per pond — shared basis for both totals below.
        _latest_per_pond = None
        if "Pond Number" in df_farm_summary.columns:
            _latest_per_pond = (
                df_farm_summary.dropna(subset=["_ParsedDate"])
                .sort_values("_ParsedDate")
                .groupby("Pond Number", as_index=False)
                .last()
            )

        # Total Expect Harvest (KG) for the farm = each pond's MOST RECENT
        # saved record's "Expect Harvest (KG)" value, summed across every
        # pond on this farm (not every historical row, which would double
        # count a pond's earlier daily estimates). Same logic as app.py.
        if _latest_per_pond is not None and "Expect Harvest (KG)" in df_farm_summary.columns:
            _harvest_vals = pd.to_numeric(_latest_per_pond["Expect Harvest (KG)"], errors="coerce").dropna()
            if len(_harvest_vals) > 0:
                total_expect_harvest_kg = float(_harvest_vals.sum())

        # Total Density for the farm = each pond's MOST RECENT saved
        # record's "Density" value, summed across every pond on this farm
        # EXCEPT ponds whose latest record shows a Full Harvest (checking
        # Harvest Type 2 first, then Harvest Type — same "2nd slot wins"
        # rule used by the Pond Layout section below). Split out per
        # Species Culture (e.g. Vannamei ponds get their own total,
        # separate from Monodon ponds) instead of one combined number,
        # since a farm can be running more than one species at once.
        if _latest_per_pond is not None and "Density" in df_farm_summary.columns:
            def _is_pond_full_h_for_density(prow):
                _t = str(prow.get("Harvest Type 2", "")).strip() or str(prow.get("Harvest Type", "")).strip()
                return "full" in _t.lower()

            _not_full_h_mask = ~_latest_per_pond.apply(_is_pond_full_h_for_density, axis=1)
            _density_pool = _latest_per_pond.loc[_not_full_h_mask].copy()
            _density_pool["Density"] = pd.to_numeric(_density_pool["Density"], errors="coerce")
            _density_pool = _density_pool.dropna(subset=["Density"])
            if len(_density_pool) > 0:
                if "Species Culture" in _density_pool.columns:
                    _density_species_label = (
                        _density_pool["Species Culture"].astype(str).str.strip().replace("", "Unspecified")
                    )
                else:
                    _density_species_label = pd.Series("Unspecified", index=_density_pool.index)
                total_density_by_species = (
                    _density_pool.groupby(_density_species_label)["Density"].sum().to_dict()
                )

            # ---- Feed-limit density: same idea, but Full H ponds are NOT excluded.
            _feed_limit_pool = _latest_per_pond.copy()
            _feed_limit_pool["Density"] = pd.to_numeric(_feed_limit_pool["Density"], errors="coerce")
            _feed_limit_pool = _feed_limit_pool.dropna(subset=["Density"])
            if len(_feed_limit_pool) > 0:
                if "Species Culture" in _feed_limit_pool.columns:
                    _feed_limit_species_label = (
                        _feed_limit_pool["Species Culture"].astype(str).str.strip().replace("", "Unspecified")
                    )
                else:
                    _feed_limit_species_label = pd.Series("Unspecified", index=_feed_limit_pool.index)
                total_density_for_feed_limit_by_species = (
                    _feed_limit_pool.groupby(_feed_limit_species_label)["Density"].sum().to_dict()
                )

        sort_cols = [c for c in ["Pond Number"] if c in df_farm_summary.columns] + ["_ParsedDate"]
        df_farm_summary = df_farm_summary.sort_values(by=sort_cols).drop(columns=["_ParsedDate"])

    # "DOC Today" = this row's saved DOC + however many days have passed
    # between its Date and today (i.e. what the DOC would be right now).
    # It's a live, always-changing number rather than something actually
    # saved in the Sheet, so it's shown in red/bold to stand out. Rows
    # whose Cycle Type is "Soon to be" haven't actually started yet, so
    # DOC Today just stays 0 for them instead of counting elapsed days.
    # A row whose Harvest Type (checking the more recent Harvest Type 2
    # first, then Harvest Type) says "Full" instead STOPS ADVANCING at
    # that row's Full Harvest date — the pond was fully harvested there,
    # so DOC Today shouldn't keep counting up to today.
    def _compute_doc_today(row):
        if str(row.get("Cycle Type") or "").strip() == "Soon to be":
            return "0"
        parsed = pd.to_datetime(row.get("Date"), errors="coerce")
        if pd.isna(parsed):
            return ""
        try:
            doc_num = int(float(row.get("DOC")))
        except (TypeError, ValueError):
            return ""
        _t2 = str(row.get("Harvest Type 2", "")).strip().lower()
        _t1 = str(row.get("Harvest Type", "")).strip().lower()
        _full_harvest_date_str = ""
        if "full" in _t2:
            _full_harvest_date_str = str(row.get("Harvest Date 2", "")).strip()
        elif "full" in _t1:
            _full_harvest_date_str = str(row.get("Harvest Date", "")).strip()
        if _full_harvest_date_str:
            _full_harvest_date = pd.to_datetime(_full_harvest_date_str, errors="coerce")
            if pd.notna(_full_harvest_date):
                return str(doc_num + (_full_harvest_date - parsed).days)
        days_passed = (pd.Timestamp(date.today()) - parsed).days
        return str(doc_num + days_passed)

    df_farm_summary["DOC Today"] = df_farm_summary.apply(_compute_doc_today, axis=1)

    # "Survival Rate" = Survival QTY expressed as a percentage of the
    # pond's stocked Density for that row (Survival QTY / Density * 100).
    # Density here is treated as the stocking count for the pond (same
    # value shown as "Stocking Density" in the Pond Layout cards below),
    # so this reads as "what fraction of what was stocked is still
    # estimated to be alive". Blank/zero/non-numeric Density or Survival
    # QTY leaves the cell blank rather than showing a divide-by-zero or
    # bogus value.
    def _compute_survival_rate(row):
        density_val = pd.to_numeric(row.get("Density"), errors="coerce")
        survival_qty_val = pd.to_numeric(row.get("Survival QTY"), errors="coerce")
        if pd.isna(density_val) or density_val <= 0 or pd.isna(survival_qty_val):
            return ""
        return f"{(survival_qty_val / density_val) * 100:,.2f}%"

    df_farm_summary["Survival Rate"] = df_farm_summary.apply(_compute_survival_rate, axis=1)

    _farm_display_cols = ["Pond Number", "Date", "Species Culture", "Cycle Type", "DOC", "DOC Today", "Density",
                           "Feed Per Day", "ABW", "Expect Harvest (KG)", "Survival QTY", "Survival Rate",
                           "Issues", "Water Color", "Grade", "Remark", "Technician",
                           "Harvest Date", "Harvest Type", "Harvest KG", "Harvest ABW",
                           "Harvest Date 2", "Harvest Type 2", "Harvest KG 2", "Harvest ABW 2"]
    _farm_display_cols = [c for c in _farm_display_cols if c in df_farm_summary.columns]

    # st.dataframe has no way to color/bold an individual column's text, so
    # this one table is rendered as a plain HTML table instead ("All Harvest
    # Details" below keeps using st.dataframe as before).
    def _escape_html(v):
        return str(v).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

    def _render_highlighted_table(df, cols, highlight_col, highlight_style="color:red;font-weight:bold;"):
        header_html = "".join(
            f"<th style='padding:6px 10px;border-bottom:2px solid #ccc;text-align:left;white-space:nowrap;'>{_escape_html(c)}</th>"
            for c in cols
        )
        rows_html = ""
        for _, r in df.iterrows():
            cells_html = ""
            for c in cols:
                cell_style = "padding:6px 10px;border-bottom:1px solid #eee;white-space:nowrap;"
                if c == highlight_col:
                    cell_style += highlight_style
                cells_html += f"<td style='{cell_style}'>{_escape_html(r.get(c, ''))}</td>"
            rows_html += f"<tr>{cells_html}</tr>"
        return (
            "<div style='overflow-x:auto; width:100%;'>"
            "<table style='width:100%; border-collapse:collapse; font-size:0.9rem;'>"
            f"<thead><tr>{header_html}</tr></thead>"
            f"<tbody>{rows_html}</tbody>"
            "</table></div>"
        )

    if show_all_saved_records:
        st.markdown(
            _render_highlighted_table(df_farm_summary, _farm_display_cols, "DOC Today"),
            unsafe_allow_html=True,
        )
        st.caption(f"{len(df_farm_summary)} saved record(s) across all ponds for {farm}.")

    if total_density_by_species:
        for _density_species, _density_val in total_density_by_species.items():
            st.markdown(
                f"**🧮 {_density_species} Ponds Total Density — {farm}: {_density_val:,.2f}**"
            )
        st.caption(
            "(sum of each pond's latest Density, excluding ponds at Full Harvest, split by Species Culture)"
        )

    if total_expect_harvest_kg is not None:
        st.markdown(
            f"**🌾 Total Expect Harvest (KG) — {farm}: {total_expect_harvest_kg:,.2f} kg** "
            "(sum of each pond's latest Expect Harvest (KG) estimate)"
        )

    # =========================================================================
    # POND LAYOUT — one card per Pond Number (using that pond's most recent
    # saved record). This combines the interactive colored-box "Pond
    # Layout" with the extra detail fields shown on the printable "Farm
    # Overview Report" pond cards in the full manager app: alongside the
    # big DOC Today / Full H / Soon to be status and the WQ Special Cases
    # flag + Total Harvest KG, each card now also shows Stocking Density,
    # L.V.D (that pond's own saved Date), Feed/Day, ABW, and an
    # Expecting Harvest (KG) / Harvest Weight line — same fields, same
    # "2nd harvest slot wins" and combined-harvest-KG parsing rules used
    # by the full manager app's Pond Layout + Farm Overview Report.
    # =========================================================================
    if "Pond Number" in df_farm_summary.columns and "DOC Today" in df_farm_summary.columns:
        st.markdown("---")
        st.markdown(f"#### 🟦 Pond Layout — {farm}")

        _pond_latest = (
            df_farm_summary.assign(_PondSortDate=pd.to_datetime(df_farm_summary["Date"], errors="coerce"))
            .dropna(subset=["_PondSortDate"])
            .sort_values("_PondSortDate")
            .groupby("Pond Number", as_index=False)
            .last()
            .sort_values("Pond Number")
        )

        def _escape_html_pond(v):
            return str(v).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

        def _pond_harvest_type(prow):
            return str(prow.get("Harvest Type 2", "")).strip() or str(prow.get("Harvest Type", "")).strip()

        # A pond keeps showing Partial H if ANY of its saved records ever
        # had a Partial harvest — not just its most recent row (daily
        # entries often leave Harvest Type blank after the harvest day).
        # Full H intentionally does NOT carry forward this way — it only
        # reflects the pond's latest record, same as before.
        _partial_history_by_pond = (
            df_farm_summary.assign(
                _HasPartial=(
                    df_farm_summary.get("Harvest Type", pd.Series("", index=df_farm_summary.index))
                    .astype(str).str.lower().str.contains("partial")
                    | df_farm_summary.get("Harvest Type 2", pd.Series("", index=df_farm_summary.index))
                    .astype(str).str.lower().str.contains("partial")
                )
            )
            .groupby("Pond Number")["_HasPartial"]
            .any()
        )

        # Latest NOTES per pond (the Sheet's last column, "Notes") — the most
        # recent saved record for that pond that actually has text in Notes.
        # Shown above each pond's box in the loop below. A pond with no Notes
        # text on any record simply has no entry here and shows nothing.
        _latest_notes_by_pond = {}
        if "Notes" in df_farm_summary.columns:
            _notes_src = df_farm_summary.assign(
                _NotesSortDate=pd.to_datetime(df_farm_summary["Date"], errors="coerce")
            )
            _notes_src = _notes_src[
                _notes_src["Notes"].astype(str).str.strip().replace("nan", "") != ""
            ]
            if len(_notes_src) > 0:
                _notes_src = _notes_src.sort_values("_NotesSortDate", kind="mergesort", na_position="first")
                _latest_notes_by_pond = (
                    _notes_src.groupby("Pond Number")["Notes"].last().astype(str).str.strip().to_dict()
                )

        # A Harvest KG / Harvest KG 2 cell is normally a plain number for a
        # single pond. But when a Partial (or Full) Harvest is done across
        # several ponds at once (e.g. Ponds 1 and 2 harvested together), the
        # user enters the COMBINED total with the pond count in parentheses,
        # e.g. "2000 (2)" — meaning 2000 kg total split across 2 ponds, and
        # that same combined value is saved on EACH of those ponds' own
        # rows. So this returns each pond's PER-POND share (2000 / 2 = 1000)
        # instead of counting the full combined figure for every pond
        # involved. Plain numeric values (no parentheses) are returned
        # as-is; unparseable values return NaN. Same parser used by the
        # full manager app's Pond Layout / Farm Overview Report.
        def _parse_pond_harvest_kg(raw_value):
            _s = str(raw_value).strip()
            if not _s:
                return float("nan")
            _match = re.match(r"^([\d,]+(?:\.\d+)?)\s*\(\s*(\d+)\s*\)\s*$", _s)
            if _match:
                _total = pd.to_numeric(_match.group(1).replace(",", ""), errors="coerce")
                _count = pd.to_numeric(_match.group(2), errors="coerce")
                if pd.notna(_total) and pd.notna(_count) and _count > 0:
                    return _total / _count
                return float("nan")
            return pd.to_numeric(_s.replace(",", ""), errors="coerce")

        # Total Harvest KG per pond — summed across EVERY saved record for
        # that pond (not just its latest one), so a pond that had one or
        # more Partial H harvests and then later a Full H harvest gets both
        # added together. Each row can contribute from both harvest slots
        # (Harvest KG / Harvest KG 2) whenever that slot's own Harvest Type
        # is filled in — a slot with a KG value but no Type text is skipped,
        # same "Type must say something" guard used elsewhere in this file.
        def _harvest_kg_sum_row(row):
            _row_total = 0.0
            _t1 = str(row.get("Harvest Type", "")).strip()
            _kg1 = _parse_pond_harvest_kg(row.get("Harvest KG", ""))
            if _t1 and pd.notna(_kg1):
                _row_total += _kg1
            _t2 = str(row.get("Harvest Type 2", "")).strip()
            _kg2 = _parse_pond_harvest_kg(row.get("Harvest KG 2", ""))
            if _t2 and pd.notna(_kg2):
                _row_total += _kg2
            return _row_total

        _total_harvest_kg_by_pond = (
            df_farm_summary.assign(_HarvestKGRow=df_farm_summary.apply(_harvest_kg_sum_row, axis=1))
            .groupby("Pond Number")["_HarvestKGRow"]
            .sum()
        )

        def _pond_status(prow):
            _pond_no_status = prow.get("Pond Number", "")
            _h_type_lower = _pond_harvest_type(prow).lower()
            _has_partial_history = bool(_partial_history_by_pond.get(_pond_no_status, False))
            if "full" in _h_type_lower:
                return "Full H"
            elif "partial" in _h_type_lower or _has_partial_history:
                return "Partial H"
            elif str(prow.get("Cycle Type", "")).strip() == "Soon to be":
                return "Soon to be"
            else:
                return "Running"

        def _pond_box_color(prow):
            _status = _pond_status(prow)
            if _status == "Partial H":
                return "#fff3cd"  # yellow — Partial Harvest
            elif _status == "Full H":
                return "#d4edda"  # green — Full Harvest
            elif _status == "Soon to be":
                return "#e2e2e2"  # gray — cycle hasn't started yet
            else:
                return "#eaf4ff"  # default blue — no harvest yet

        def _species_letter(prow):
            _species = str(prow.get("Species Culture", "")).strip().lower()
            if "vannamei" in _species:
                return "V"
            elif "monodon" in _species:
                return "M"
            else:
                return ""

        _pond_boxes_html = ""
        for _, _prow in _pond_latest.iterrows():
            _pond_no = _escape_html_pond(_prow.get("Pond Number", ""))
            _box_color = _pond_box_color(_prow)
            _status_box = _pond_status(_prow)

            # Latest Notes for this pond (Sheet's last column) — shown ABOVE
            # the pond box. Empty when the pond has no Notes text.
            _notes_val = str(_latest_notes_by_pond.get(_prow.get("Pond Number", ""), "")).strip()
            _notes_html = (
                "<div style='width:210px;box-sizing:border-box;margin-bottom:4px;padding:4px 6px;"
                "background:#fffbe6;border:1px solid #e6c200;border-radius:6px;font-size:0.8rem;"
                "color:#333;text-align:left;line-height:1.3;word-break:break-word;'>"
                f"{_escape_html_pond(_notes_val)}</div>"
                if _notes_val and _notes_val.lower() != "nan" else ""
            )

            # Sad-face icon shown in the top-right corner of the box when
            # this pond's latest saved record has any text in the
            # "WQ Special Cases" column, plus the same text repeated as a
            # small label under the box (same pattern as the full manager
            # app's Pond Layout).
            _wq_special_val = str(_prow.get("WQ Special Cases", "")).strip()
            _wq_special_icon_html = (
                "<div style='position:absolute;top:2px;right:4px;font-size:1.3rem;line-height:1;' "
                "title='WQ Special Case'>🫨</div>"
                if _wq_special_val else ""
            )
            _wq_special_text_html = (
                f"<div style='font-size:0.85rem;color:#b45309;text-align:center;"
                f"max-width:190px;margin-top:2px;'>🫨 {_escape_html_pond(_wq_special_val)}</div>"
                if _wq_special_val else ""
            )

            # ---- Farm-Overview-style detail lines shown on every card:
            # Stocking Density, L.V.D (this pond's own saved Date),
            # Feed/Day, ABW.
            _density_val = pd.to_numeric(_prow.get("Density", ""), errors="coerce")
            _density_str = f"{_density_val:,.0f}" if pd.notna(_density_val) else "-"
            _lvd_str = _escape_html_pond(str(_prow.get("Date", "")).strip() or "-")
            _feed_day_str = _escape_html_pond(_prow.get("Feed Per Day", "") or "-")
            _abw_str = _escape_html_pond(_prow.get("ABW", "") or "-")
            _extra_details_html = (
                "<div style='font-size:0.85rem;color:#333;text-align:left;width:100%;"
                "padding:0 8px;margin-top:4px;line-height:1.4;'>"
                f"<div>Stocking Density - {_density_str}</div>"
                f"<div>L.V.D - {_lvd_str}</div>"
                f"<div>Feed/Day - {_feed_day_str} &nbsp;|&nbsp; ABW - {_abw_str}</div>"
                "</div>"
            )

            # Issues (from this pond's latest saved record) — shown at the
            # bottom inside the box as "Disease: <issue>", with the issue
            # text in red. Stays empty when the latest record has no Issues.
            _issues_val = str(_prow.get("Issues", "")).strip()
            _issues_html = (
                "<div style='margin-top:auto;width:100%;text-align:center;font-size:0.85rem;"
                "font-weight:bold;border-top:1px dashed #bbb;padding-top:3px;'>"
                "<span style='color:#333;'></span>"
                f"<span style='color:red;'>{_escape_html_pond(_issues_val)}</span></div>"
                if _issues_val and _issues_val.lower() != "nan" else ""
            )

            if _status_box == "Full H":
                # Full H ponds: show "Full H" + its Harvest Date, plus the
                # pond's Total Harvest KG (all harvests summed) instead of
                # DOC Today.
                _h_date = str(_prow.get("Harvest Date 2", "")).strip() or str(_prow.get("Harvest Date", "")).strip()
                _h_date = _escape_html_pond(_h_date or "-")
                _total_kg_box = _total_harvest_kg_by_pond.get(_prow.get("Pond Number", ""), 0)
                _total_kg_html = (
                    f"<div style='font-size:0.75rem;color:#333;'>Total: {_total_kg_box:,.2f} KG</div>"
                    if _total_kg_box else ""
                )
                _box_middle_html = (
                    "<div style='font-size:1.2rem;font-weight:bold;color:red;'>Full H</div>"
                    f"<div style='font-size:0.75rem;color:#333;'>Harvest Date - {_h_date}</div>"
                    f"{_total_kg_html}"
                )
            elif _status_box == "Soon to be":
                # Soon to be ponds: show "Soon to be" instead of DOC Today,
                # matching the gray box color already assigned by
                # _pond_box_color() for this status.
                _box_middle_html = (
                    "<div style='font-size:1.1rem;font-weight:bold;color:#555;'>Soon to be</div>"
                )
            else:
                # Running / Partial H ponds: keep showing the DOC Today
                # number, but the label under it now shows the pond's
                # Started Date (today's date minus DOC Today days) instead
                # of the literal "DOC Today" text. Partial H ponds
                # additionally show the pond's Total Harvest KG (all
                # harvests summed so far) — Running ponds have no harvest
                # yet, so this stays blank for them.
                if _status_box == "Partial H":
                    _total_kg_box = _total_harvest_kg_by_pond.get(_prow.get("Pond Number", ""), 0)
                    _total_kg_html = (
                        f"<div style='font-size:0.7rem;color:#333;'>Total: {_total_kg_box:,.2f} KG</div>"
                        if _total_kg_box else ""
                    )
                else:
                    _total_kg_html = ""

                _doc_today_raw = _prow.get("DOC Today", "")
                _doc_today_val = _escape_html_pond(_doc_today_raw or "-")
                try:
                    _started_date = (
                        pd.Timestamp(date.today()) - pd.Timedelta(days=int(float(_doc_today_raw)))
                    ).strftime("%Y-%m-%d")
                except (TypeError, ValueError):
                    _started_label = "Started on ---"
                    _started_date = None
                if _started_date is not None:
                    _started_label = f"Started on {_started_date}"
                _box_middle_html = (
                    f"<div style='font-size:1.4rem;font-weight:bold;color:red;'>{_doc_today_val}</div>"
                    f"<div style='font-size:0.7rem;color:#777;'>{_escape_html_pond(_started_label)}</div>"
                    f"{_total_kg_html}"
                )

            # Expecting Harvest (KG) / Harvest Weight line — same label
            # switch and "2nd slot wins" rule used by the full manager
            # app's printable Farm Overview Report pond cards. Uses the
            # same _parse_pond_harvest_kg() parser as the "Total: X KG"
            # line above — a plain pd.to_numeric() here would fail on a
            # combined-harvest value like "3500 (2)" (3500kg split across
            # 2 ponds) and silently show "-" even though a real number was
            # saved, which is why Harvest Weight could show "-" while
            # Total still showed the correct figure.
            if _status_box == "Full H":
                _t2_expect = str(_prow.get("Harvest Type 2", "")).strip().lower()
                _kg2_expect = _parse_pond_harvest_kg(_prow.get("Harvest KG 2", ""))
                _kg1_expect = _parse_pond_harvest_kg(_prow.get("Harvest KG", ""))
                _harvest_kg_val = _kg2_expect if ("full" in _t2_expect and pd.notna(_kg2_expect)) else _kg1_expect
                _expect_label = "Harvest Weight"
                _expect_val = f"{_harvest_kg_val:,.2f} KG" if pd.notna(_harvest_kg_val) else "-"
            elif _status_box == "Soon to be":
                _expect_label = "Expecting Harvest"
                _expect_val = "-"
            else:
                _expect_label = "Expecting Harvest"
                _expect_kg = pd.to_numeric(_prow.get("Expect Harvest (KG)", ""), errors="coerce")
                _expect_val = f"{_expect_kg:,.2f} KG" if pd.notna(_expect_kg) else "-"

            _expect_html = (
                "<div style='font-size:0.85rem;color:#333;text-align:center;width:100%;margin-top:4px;"
                "border-top:1px dashed #bbb;padding-top:3px;'>"
                f"<b>{_expect_label}:</b> {_escape_html_pond(_expect_val)}</div>"
            )

            # Survival Rate line — Survival QTY as a percentage of this
            # pond's Density (Survival QTY / Density * 100), placed right
            # after the Expecting Harvest / Harvest Weight line above.
            # Blank/zero/non-numeric Density or Survival QTY shows "-"
            # instead of a divide-by-zero or bogus value.
            _density_for_survival = pd.to_numeric(_prow.get("Density", ""), errors="coerce")
            _survival_qty_val = pd.to_numeric(_prow.get("Survival QTY", ""), errors="coerce")
            if pd.notna(_density_for_survival) and _density_for_survival > 0 and pd.notna(_survival_qty_val):
                _survival_rate_str = f"{(_survival_qty_val / _density_for_survival) * 100:,.2f}%"
            else:
                _survival_rate_str = "-"
            _survival_rate_html = (
                "<div style='font-size:0.85rem;color:#333;text-align:center;width:100%;margin-top:2px;'>"
                f"<b>Survival Rate:</b> {_escape_html_pond(_survival_rate_str)}</div>"
            )

            _species_label = _species_letter(_prow)
            _species_html = (
                f"<div style='font-size:0.75rem;font-weight:bold;color:#444;margin-top:2px;'>{_species_label}</div>"
                if _species_label else ""
            )

            _pond_boxes_html += (
                "<div style='display:flex;flex-direction:column;align-items:center;margin:6px;'>"
                f"{_notes_html}"
                f"<div style='position:relative;width:210px;min-height:175px;border:2px solid #333;"
                "border-radius:6px;display:flex;flex-direction:column;align-items:center;"
                f"justify-content:flex-start;padding:8px 0;background:{_box_color};'>"
                f"{_wq_special_icon_html}"
                f"<div style='font-size:0.8rem;color:#555;'>Pond {_pond_no}</div>"
                f"{_box_middle_html}"
                f"{_expect_html}"
                f"{_survival_rate_html}"
                f"{_extra_details_html}"
                f"{_issues_html}"
                "</div>"
                f"{_species_html}"
                f"{_wq_special_text_html}"
                "</div>"
            )

        st.markdown(
            "<div style='display:flex;gap:18px;justify-content:center;margin-bottom:8px;font-size:0.85rem;'>"
            "<div><span style='display:inline-block;width:14px;height:14px;background:#eaf4ff;"
            "border:1px solid #333;border-radius:3px;vertical-align:middle;margin-right:6px;'></span>Running</div>"
            "<div><span style='display:inline-block;width:14px;height:14px;background:#fff3cd;"
            "border:1px solid #333;border-radius:3px;vertical-align:middle;margin-right:6px;'></span>Partial H</div>"
            "<div><span style='display:inline-block;width:14px;height:14px;background:#d4edda;"
            "border:1px solid #333;border-radius:3px;vertical-align:middle;margin-right:6px;'></span>Full H</div>"
            "</div>",
            unsafe_allow_html=True,
        )

        st.markdown(
            f"<div style='display:flex;flex-wrap:wrap;justify-content:center;'>{_pond_boxes_html}</div>",
            unsafe_allow_html=True,
        )
else:
    st.info(f"No saved records yet for {farm}.")

# =========================================================================
# SALES DETAILS — read-only, straight from the Sales Details Google Sheet,
# filtered to the Customer Code belonging to the Customer + Farm selected
# above. VIEW-ONLY: there is no delete/recycle-bin control here — a date
# already marked Settle = 'Yes' in the Sheet (e.g. by the full manager
# app) stays excluded, but nothing on this page can set that flag itself.
# =========================================================================
st.markdown("---")
st.markdown(f"#### 🧾 Sales Details — {farm}")

try:
    df_sales = load_sales_data()
except Exception as e:
    df_sales = None
    st.error(f"❌ Could not connect to the Sales Details Google Sheet. Check sharing settings.\n\n{e}")

if df_sales is not None:
    if not selected_customer_code:
        st.info(
            "No Customer Code found for this farm in 'Customer List.xlsx', so Sales Details "
            "can't be filtered. Add a 'Customer Code' column to the customer list to enable this."
        )
    elif len(df_sales) == 0:
        st.info("No sales records found in the Sales Details sheet.")
    else:
        df_sales_farm = df_sales[
            df_sales["Customer Code"].astype(str).str.strip().str.lower()
            == selected_customer_code.strip().lower()
        ].copy()

        if len(df_sales_farm) == 0:
            st.info(f"No sales records found for Customer Code '{selected_customer_code}'.")
        else:
            df_sales_farm["Quantity"] = pd.to_numeric(df_sales_farm["Quantity"], errors="coerce").fillna(0)
            df_sales_farm["Sales Amt"] = pd.to_numeric(df_sales_farm["Sales Amt"], errors="coerce").fillna(0)
            df_sales_farm["Settle"] = df_sales_farm["Settle"].astype(str)

            # Same pivot as before: one row per Date, one column per Item
            # Description, cell value = summed Quantity for that date/item.
            pivot_sales = df_sales_farm.pivot_table(
                index="Date",
                columns="Item Description",
                values="Quantity",
                aggfunc="sum",
                fill_value=0,
            )
            pivot_sales = pivot_sales.sort_index()
            pivot_sales.index.name = "Date"

            # Dates already marked Settle = 'Yes' in the Sheet are treated
            # as permanently removed — they're excluded before the table
            # is even built, so a refresh doesn't bring them back.
            _settled_by_date = (
                df_sales_farm.groupby("Date")["Settle"]
                .apply(lambda s: s.str.strip().str.lower().eq("yes").all())
            )

            pivot_display = pivot_sales.reset_index()
            pivot_display = pivot_display[
                ~pivot_display["Date"].map(_settled_by_date).fillna(False)
            ].reset_index(drop=True)

            # VIEW-ONLY: no recycle-bin / delete control here — Technical
            # Officers & Marketing Managers can only view these sales line
            # items, never remove one. A date already marked Settle = 'Yes'
            # in the Sheet (e.g. by the full manager app) is still excluded
            # above, but nothing on this page can set that flag itself.
            st.dataframe(
                pivot_display,
                use_container_width=True,
                hide_index=True,
            )

            # Last Feed Order — the most recent date on which any FEED
            # item ("Item No." starting with "FEED", same prefix check
            # used for the Total Quantity/Total Sales Amt figures below)
            # was purchased for this farm, plus the total feed Quantity
            # bought on that date. Uses df_sales_farm (not yet filtered to
            # visible/kept dates) so a settled date's last feed order still
            # shows here even if that date's row is excluded from the
            # pivot table above.
            _last_feed_mask = (
                df_sales_farm["Item No."].astype(str).str.strip().str.upper().str.startswith("FEED")
            )
            df_last_feed = df_sales_farm[_last_feed_mask]
            if len(df_last_feed) > 0:
                _last_feed_date = df_last_feed["Date"].max()
                _last_feed_items = (
                    df_last_feed[df_last_feed["Date"] == _last_feed_date]
                    .groupby("Item Description")["Quantity"]
                    .sum()
                )
                _last_feed_order_str = ", ".join(
                    f"{_item} - {_qty:,.0f}" for _item, _qty in _last_feed_items.items()
                )
                st.markdown(
                    f"**Last Feed Order: {_last_feed_order_str}  |  Last Feed Purchased Date: {_last_feed_date}**"
                )
            else:
                st.markdown("**Last Feed Order: -  |  Last Feed Purchased Date: -**")

            kept_dates = pivot_display["Date"].dropna()

            df_sales_farm_visible = df_sales_farm[df_sales_farm["Date"].isin(kept_dates)]

            _caption = (
                f"{len(df_sales_farm_visible)} sales line item(s) across "
                f"{len(kept_dates)} date(s) for Customer Code '{selected_customer_code}'."
            )
            st.caption(_caption)

            # Total Quantity / Total Sales Amt reflect FEED items only —
            # rows whose "Item No." starts with "FEED" (same prefix check
            # used elsewhere in the family of apps), so non-feed line
            # items no longer inflate these two totals.
            _sales_totals_feed_mask = (
                df_sales_farm_visible["Item No."].astype(str).str.strip().str.upper().str.startswith("FEED")
            )
            df_sales_farm_visible_feed_only = df_sales_farm_visible[_sales_totals_feed_mask]

            total_qty = df_sales_farm_visible_feed_only["Quantity"].sum() if len(df_sales_farm_visible_feed_only) else 0
            total_amt = df_sales_farm_visible_feed_only["Sales Amt"].sum() if len(df_sales_farm_visible_feed_only) else 0
            st.markdown(
                f"**Total Feed Quantity: {total_qty:,.0f}  |  Total Feed Sales Amt: {total_amt:,.2f}**"
            )

            # =================================================================
            # FEED ORDER STATUS — two rows of rectangles, one per feed size,
            # in a fixed order: NANAMI FEED sizes, then EGO FEED sizes. Each
            # box shows that size's total purchased Quantity in the middle
            # and its latest purchase Date below. Boxes are colored one way
            # if any quantity has been purchased, another if not yet.
            #
            # NANAMI_LIMIT_FACTORS / EGO_LIMIT_FACTORS hold the "Do not
            # exceed" formula for each brand's sizes. Each size's limit =
            # factor * that farm's Total Density for the species the brand
            # is fed to — NANAMI uses Vannamei Ponds Total Density, EGO
            # uses Monodon Ponds Total Density. These feed-limit densities
            # come from total_density_for_feed_limit_by_species (computed in
            # "All Saved Records" above), which INCLUDES ponds at Full
            # Harvest, unlike the displayed "Ponds Total Density" lines.
            # =================================================================
            NANAMI_FEED_ORDER = [
                "NANAMI 1", "NANAMI 1S", "NANAMI 2S", "NANAMI 3S",
                "NANAMI 3M", "NANAMI 3L", "NANAMI 4",
            ]
            EGO_FEED_ORDER = [
                "EGO - 01", "EGO - 01S", "EGO - 02S", "EGO - 03S",
                "EGO - 03M", "EGO - 03L", "EGO - 04L",
            ]
            NANAMI_LIMIT_FACTORS = {
                "NANAMI 1": 50 / 100000,
                "NANAMI 1S": 150 / 100000,
                "NANAMI 2S": 150 / 100000,
                "NANAMI 3S": 150 / 100000,
                "NANAMI 3M": 750 / 100000,
                "NANAMI 3L": 1000 / 100000
            }
            EGO_LIMIT_FACTORS = {
                "EGO - 01": 50 / 100000,
                "EGO - 01S": 150 / 100000,
                "EGO - 02S": 150 / 100000,
                "EGO - 03S": 200 / 100000,
                "EGO - 03M": 500 / 100000,
                "EGO - 03L": 750 / 100000, "EGO - 04L": 1000 / 100000
            }

            # The Vannamei / Monodon entries from
            # total_density_for_feed_limit_by_species (computed in the
            # "All Saved Records" section above, for this same Customer +
            # Farm, INCLUDING Full H ponds). Matched case-insensitively
            # since the exact text comes from whatever the Species Culture
            # column contains (e.g. "Vannamei", "L. vannamei", "Monodon",
            # etc.).
            _vannamei_total_density = 0.0
            _monodon_total_density = 0.0
            for _density_species_key, _density_species_val in total_density_for_feed_limit_by_species.items():
                _density_key_lower = str(_density_species_key).lower()
                if "vannamei" in _density_key_lower:
                    _vannamei_total_density = _density_species_val
                elif "monodon" in _density_key_lower:
                    _monodon_total_density = _density_species_val

            def _escape_html_feed(v):
                return str(v).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

            def _render_feed_row(title, size_labels, limit_factors=None, limit_density=0.0):
                _desc_upper = df_sales_farm_visible["Item Description"].astype(str).str.strip().str.upper()
                boxes_html = ""
                for _label in size_labels:
                    _subset = df_sales_farm_visible[_desc_upper == _label.upper()]
                    _qty = _subset["Quantity"].sum() if len(_subset) else 0
                    _latest_date = _subset["Date"].max() if len(_subset) else "-"
                    _has_qty = _qty > 0

                    # "Do not exceed" limit for this size, if it has one —
                    # factor * limit_density (Vannamei Ponds Total Density
                    # for NANAMI, Monodon Ponds Total Density for EGO).
                    _limit_val = None
                    if limit_factors and _label in limit_factors:
                        _limit_val = limit_factors[_label] * limit_density

                    if _limit_val is not None and _qty > _limit_val:
                        _box_color = "#ff4d4d"  # red — exceeded its "Do not exceed" limit
                    elif _has_qty:
                        _box_color = "#d4edda"  # green if purchased, gray if not yet
                    else:
                        _box_color = "#e2e2e2"

                    _qty_label = f"{_qty:,.0f}" if _has_qty else "-"

                    # Limit caption sits BELOW the box, outside its border —
                    # only shown for sizes that have a limit_factors entry.
                    _limit_caption_html = ""
                    if _limit_val is not None:
                        _limit_caption_html = (
                            "<div style='font-size:0.65rem;color:#555;text-align:center;"
                            f"margin-top:3px;max-width:120px;'>Do not exceed: {_limit_val:,.2f}</div>"
                        )

                    boxes_html += (
                        "<div style='display:flex;flex-direction:column;align-items:center;margin:5px;'>"
                        "<div style='width:120px;height:80px;border:2px solid #333;border-radius:6px;"
                        "display:flex;flex-direction:column;align-items:center;justify-content:center;"
                        f"background:{_box_color};'>"
                        f"<div style='font-size:0.7rem;color:#555;'>{_escape_html_feed(_label)}</div>"
                        f"<div style='font-size:1.2rem;font-weight:bold;color:#111;'>{_qty_label}</div>"
                        f"<div style='font-size:0.65rem;color:#777;'>{_escape_html_feed(_latest_date)}</div>"
                        "</div>"
                        f"{_limit_caption_html}"
                        "</div>"
                    )
                st.markdown(f"**{title}**")
                st.markdown(
                    f"<div style='display:flex;flex-wrap:wrap;justify-content:center;'>{boxes_html}</div>",
                    unsafe_allow_html=True,
                )

            st.markdown("---")
            _render_feed_row("🐟 NANAMI FEED", NANAMI_FEED_ORDER, limit_factors=NANAMI_LIMIT_FACTORS,
                              limit_density=_vannamei_total_density)
            st.markdown("")
            _render_feed_row("🐟 EGO FEED", EGO_FEED_ORDER, limit_factors=EGO_LIMIT_FACTORS,
                              limit_density=_monodon_total_density)

# =========================================================================
# ALL HARVEST DETAILS — every row (across ALL customers/farms/ponds in the
# Google Sheet, not just the one selected above) that has a non-blank
# value in EITHER harvest slot: Harvest Date/Type (the first harvest) or
# Harvest Date 2/Type 2 (a second harvest for the same pond row).
#
# A sheet row that has BOTH slots filled in (a Partial harvest followed
# later by a Full harvest on the same pond record) is shown here as TWO
# SEPARATE ROWS — one per harvest event — so each can be reviewed and
# removed on its own via the recycle bin, without affecting the other
# harvest event on that same underlying sheet row. This split is purely a
# display-time thing: nothing is duplicated, merged, or deleted in the
# Google Sheet because of it.
#
# VIEW-ONLY: unlike the full manager app, this table has NO recycle-bin /
# delete control — Technical Officers & Marketing Managers can only view
# these records here, never remove one. A harvest event already flagged
# 'H' (Harvest Status / Harvest Status 2) by the full manager app is still
# hidden from this table (same read-only filter as before), but nothing on
# this page can set that flag itself.
#
# This is the last section in the Marketing Manager view — the Zone-Wise
# Harvest breakdown, Running List, and Species-wise Pond Summary sections
# from the full manager app are intentionally not included here.
# =========================================================================
st.markdown("---")
st.markdown("#### 🌾 All Harvest Details")

df_all_records = load_data()
_harvest_cols_needed = {"Harvest Date", "Harvest Type", "Harvest Date 2", "Harvest Type 2"}
if len(df_all_records) > 0 and _harvest_cols_needed.issubset(df_all_records.columns):
    _harvest_mask = (
        (df_all_records["Harvest Date"].astype(str).str.strip() != "")
        | (df_all_records["Harvest Type"].astype(str).str.strip() != "")
        | (df_all_records["Harvest Date 2"].astype(str).str.strip() != "")
        | (df_all_records["Harvest Type 2"].astype(str).str.strip() != "")
    )
    df_harvest_source = df_all_records[_harvest_mask].copy()
else:
    df_harvest_source = pd.DataFrame(columns=COLUMN_ORDER)

if len(df_harvest_source) > 0:
    # "DOC" and "Date" below are computed ONCE per underlying sheet row
    # (unchanged formulas — DOC checks whichever slot's Type says "Full",
    # 2nd slot first; Date uses the Timestamp's own date) and then carried
    # onto BOTH split rows for that sheet row, since they describe the
    # saved record as a whole rather than one harvest slot specifically.
    def _harvest_doc_display(row):
        try:
            _doc_num = int(float(row.get("DOC")))
        except (TypeError, ValueError):
            return row.get("DOC", "")
        _row_date = pd.to_datetime(row.get("Date"), errors="coerce")
        if pd.isna(_row_date):
            return row.get("DOC", "")
        _t2 = str(row.get("Harvest Type 2", "")).strip().lower()
        _t1 = str(row.get("Harvest Type", "")).strip().lower()
        _full_harvest_date_str = ""
        if "full" in _t2:
            _full_harvest_date_str = str(row.get("Harvest Date 2", "")).strip()
        elif "full" in _t1:
            _full_harvest_date_str = str(row.get("Harvest Date", "")).strip()
        if _full_harvest_date_str:
            _full_harvest_date = pd.to_datetime(_full_harvest_date_str, errors="coerce")
            if pd.notna(_full_harvest_date):
                return str(_doc_num + (_full_harvest_date - _row_date).days)
        return str(_doc_num + (pd.Timestamp(date.today()) - _row_date).days)

    df_harvest_source["DOC"] = df_harvest_source.apply(_harvest_doc_display, axis=1)

    if "Timestamp" in df_harvest_source.columns:
        _harvest_timestamp_date = pd.to_datetime(
            df_harvest_source["Timestamp"], errors="coerce"
        ).dt.strftime("%Y-%m-%d")
        df_harvest_source["Date"] = _harvest_timestamp_date.fillna(df_harvest_source["Date"])

    # ---- split each qualifying sheet row into one row per filled harvest
    # slot. Each split row carries a "_SlotKey" (Timestamp + which slot)
    # used only internally to track recycle-bin deletions below — it is
    # never shown in the table.
    _split_rows = []
    for _, _r in df_harvest_source.iterrows():
        _status1 = str(_r.get("Harvest Status", "")).strip().upper()
        _status2 = str(_r.get("Harvest Status 2", "")).strip().upper()
        _slot1_filled = (
            str(_r.get("Harvest Date", "")).strip() != "" or str(_r.get("Harvest Type", "")).strip() != ""
        )
        _slot2_filled = (
            str(_r.get("Harvest Date 2", "")).strip() != "" or str(_r.get("Harvest Type 2", "")).strip() != ""
        )
        if _slot1_filled and _status1 != "H":
            _row1 = _r.to_dict()
            _row1["_SlotKey"] = f"{_r.get('Timestamp', '')}__1"
            _split_rows.append(_row1)
        if _slot2_filled and _status2 != "H":
            _row2 = _r.to_dict()
            _row2["_SlotKey"] = f"{_r.get('Timestamp', '')}__2"
            _row2["Harvest Date"] = _r.get("Harvest Date 2", "")
            _row2["Harvest Type"] = _r.get("Harvest Type 2", "")
            _row2["Harvest KG"] = _r.get("Harvest KG 2", "")
            _row2["Harvest ABW"] = _r.get("Harvest ABW 2", "")
            _split_rows.append(_row2)

    df_harvest_all = (
        pd.DataFrame(_split_rows) if _split_rows
        else df_harvest_source.iloc[0:0].assign(_SlotKey=pd.Series(dtype=str))
    )
else:
    df_harvest_all = pd.DataFrame(columns=COLUMN_ORDER)

if len(df_harvest_all) > 0:
    if "Date" in df_harvest_all.columns:
        df_harvest_all["_ParsedDate"] = pd.to_datetime(df_harvest_all["Date"], errors="coerce")
        _harvest_sort_cols = [c for c in ["Customer", "Farm Name with Code", "Pond Number"]
                               if c in df_harvest_all.columns] + ["_ParsedDate"]
        df_harvest_all = df_harvest_all.sort_values(by=_harvest_sort_cols).drop(columns=["_ParsedDate"])

    # "Estimated Harvest Value" — a self-contained per-row estimate:
    #   price per KG = 1500 + 20 for every 1g of ABW above 10
    #   Estimated Harvest Value = price per KG * Harvest KG
    # e.g. Harvest KG 1200, ABW 13 -> (1500 + 20*(13-10)) * 1200.
    #
    # Harvest KG can be a combined figure across several ponds written as
    # "1500 (3)" (1500 total split across 3 ponds) — same pattern already
    # parsed elsewhere in this file (Pond Layout's _parse_pond_harvest_kg)
    # — so that's parsed into a per-pond share (1500 / 3) here too, via a
    # local copy of that same parser so this section stays self-contained.
    # Harvest ABW can likewise be a range like "9-11" — parsed as its
    # midpoint, (9+11)/2 = 10.
    def _harvest_value_parse_kg(raw_value):
        _s = str(raw_value).strip()
        if not _s:
            return float("nan")
        _match = re.match(r"^([\d,]+(?:\.\d+)?)\s*\(\s*(\d+)\s*\)\s*$", _s)
        if _match:
            _total = pd.to_numeric(_match.group(1).replace(",", ""), errors="coerce")
            _count = pd.to_numeric(_match.group(2), errors="coerce")
            if pd.notna(_total) and pd.notna(_count) and _count > 0:
                return _total / _count
            return float("nan")
        return pd.to_numeric(_s.replace(",", ""), errors="coerce")

    def _harvest_value_parse_abw(raw_value):
        _s = str(raw_value).strip()
        if not _s:
            return float("nan")
        _range_match = re.match(r"^(\d+(?:\.\d+)?)\s*-\s*(\d+(?:\.\d+)?)$", _s)
        if _range_match:
            _lo = pd.to_numeric(_range_match.group(1), errors="coerce")
            _hi = pd.to_numeric(_range_match.group(2), errors="coerce")
            if pd.notna(_lo) and pd.notna(_hi):
                return (_lo + _hi) / 2
            return float("nan")
        return pd.to_numeric(_s, errors="coerce")

    def _harvest_estimated_value(row):
        _kg_val = _harvest_value_parse_kg(row.get("Harvest KG", ""))
        _abw_val = _harvest_value_parse_abw(row.get("Harvest ABW", ""))
        if pd.isna(_kg_val) or pd.isna(_abw_val):
            return ""
        _price_per_kg = 1500 + 20 * (_abw_val - 10)
        return f"{_price_per_kg * _kg_val:,.2f}"

    df_harvest_all["Estimated Harvest Value"] = df_harvest_all.apply(_harvest_estimated_value, axis=1)

    # NOTE: "Cycle Type" intentionally left out of this table's display
    # columns — same as the full manager app. "Harvest Date 2" /
    # "Harvest Type 2" / "Harvest KG 2" / "Harvest ABW 2" are also left
    # out now — each harvest slot gets its own row above, so the single
    # "Harvest Date/Type/KG/ABW" columns already carry whichever slot
    # that row represents.
    _harvest_display_cols = ["Customer", "Farm Name with Code", "Pond Number", "Date", "DOC",
                              "Species Culture", "Harvest Date", "Harvest Type",
                              "Harvest KG", "Harvest ABW", "Estimated Harvest Value",
                              "Harvest Submitted Date", "Technician"]
    _harvest_display_cols = [c for c in _harvest_display_cols if c in df_harvest_all.columns]

    # NOTE: This app is a VIEW-ONLY app for Technical Officers & Marketing
    # Managers — unlike the full manager app, there is intentionally no
    # recycle-bin / delete affordance here. The table below is a plain
    # read-only st.dataframe; nothing on this page can flag or hide a
    # harvest record in the Google Sheet.
    _harvest_full_cols = ["_SlotKey", "Timestamp"] + _harvest_display_cols
    df_harvest_view_source = df_harvest_all[_harvest_full_cols].reset_index(drop=True)

    # Manual Sort by / Order controls (kept from before, purely for
    # viewing convenience). Defaults to "Harvest Submitted Date" /
    # Descending so the most recently submitted harvests show up first.
    _hsort_col1, _hsort_col2 = st.columns(2)
    with _hsort_col1:
        _default_sort_idx = (
            _harvest_display_cols.index("Harvest Submitted Date")
            if "Harvest Submitted Date" in _harvest_display_cols else 0
        )
        _harvest_sort_by = st.selectbox(
            "Sort by", options=_harvest_display_cols, index=_default_sort_idx, key="harvest_sort_by"
        )
    with _hsort_col2:
        _harvest_sort_order = st.selectbox(
            "Order", options=["Ascending", "Descending"], index=1, key="harvest_sort_order"
        )

    def _harvest_sort_key(series):
        numeric = pd.to_numeric(series, errors="coerce")
        if numeric.notna().sum() >= max(1, len(series) * 0.5):
            return numeric
        parsed_dt = pd.to_datetime(series, errors="coerce")
        if parsed_dt.notna().sum() >= max(1, len(series) * 0.5):
            return parsed_dt
        return series.astype(str).str.lower()

    df_harvest_view_source = (
        df_harvest_view_source.assign(_SortKey=_harvest_sort_key(df_harvest_view_source[_harvest_sort_by]))
        .sort_values(by="_SortKey", ascending=(_harvest_sort_order == "Ascending"), na_position="last")
        .drop(columns=["_SortKey"])
        .reset_index(drop=True)
    )

    st.dataframe(
        df_harvest_view_source[_harvest_display_cols],
        use_container_width=True,
        hide_index=True,
    )
    st.caption(f"{len(df_harvest_view_source)} harvest record(s) shown.")
else:
    st.info("No harvest details recorded yet.")

# =========================================================================
# LAST VISIT DATE REPORT. Ported as-is from the full manager app
# (manager.py). For every Customer + Farm, shows:
#   Customer Name | Farm Name with Code | Zone | Status |
#   Data Entered Latest Date | Due Date
#
# Status = "FULL H" when every pond that farm has ever had a saved record
# for is currently at Full Harvest (same per-pond Full Harvest check used
# throughout this file: Harvest Type 2 first, then Harvest Type, on that
# pond's own most recent saved record) — otherwise "Running".
#
# Data Entered Latest Date = the most recent Date across ALL of that
# farm's saved records (any pond, any row) — i.e. the last time anything
# was entered for this farm.
#
# Due Date = days elapsed between today and Data Entered Latest Date
# (today's date minus that date, in days).
#
# Only rows with Status = "Running" AND Due Date > 7 are shown, so this
# report only surfaces farms that are overdue for a visit — farms at
# FULL H, or Running farms visited within the last 7 days, are left out.
#
# Entirely read-only, self-contained (own local Zone lookup and Full
# Harvest check), so it works on its own regardless of the sections above.
# =========================================================================
st.markdown("---")
st.markdown("#### 🗓️ Last Visit Date Report")

df_all_for_last_visit = load_data()
_last_visit_required = {"Customer", "Farm Name with Code", "Pond Number", "Date",
                         "Harvest Type", "Harvest Type 2"}
if len(df_all_for_last_visit) > 0 and _last_visit_required.issubset(df_all_for_last_visit.columns):
    df_all_for_last_visit = df_all_for_last_visit.copy()
    df_all_for_last_visit["_ParsedDate"] = pd.to_datetime(df_all_for_last_visit["Date"], errors="coerce")

    # Latest saved record per Customer+Farm+Pond, same rule used elsewhere
    # in this file — used only to determine each pond's Full Harvest
    # status for the farm-level "Status" column below.
    _latest_per_pond_last_visit = (
        df_all_for_last_visit.dropna(subset=["_ParsedDate"])
        .sort_values("_ParsedDate")
        .groupby(["Customer", "Farm Name with Code", "Pond Number"], as_index=False)
        .last()
    )

    def _is_full_harvest_pond_last_visit(prow):
        _t = str(prow.get("Harvest Type 2", "")).strip() or str(prow.get("Harvest Type", "")).strip()
        return "full" in _t.lower()

    _latest_per_pond_last_visit["_IsFullH"] = _latest_per_pond_last_visit.apply(
        _is_full_harvest_pond_last_visit, axis=1
    )
    _farm_full_h_status_last_visit = (
        _latest_per_pond_last_visit.groupby(["Customer", "Farm Name with Code"])
        .agg(_TotalPonds=("Pond Number", "nunique"), _FullHPonds=("_IsFullH", "sum"))
        .reset_index()
    )
    _farm_full_h_status_last_visit["Status"] = _farm_full_h_status_last_visit.apply(
        lambda r: "FULL H" if r["_FullHPonds"] >= r["_TotalPonds"] and r["_TotalPonds"] > 0 else "Running",
        axis=1,
    )

    # Data Entered Latest Date — the most recent Date across ALL saved
    # records for that farm (any pond, any row), not just the latest
    # record per pond.
    _farm_latest_date = (
        df_all_for_last_visit.dropna(subset=["_ParsedDate"])
        .groupby(["Customer", "Farm Name with Code"])["_ParsedDate"]
        .max()
        .reset_index()
        .rename(columns={"_ParsedDate": "_LatestDateParsed"})
    )

    _farm_last_visit = _farm_full_h_status_last_visit[
        ["Customer", "Farm Name with Code", "Status"]
    ].merge(_farm_latest_date, on=["Customer", "Farm Name with Code"], how="left")

    _today_last_visit = pd.Timestamp(date.today())
    _farm_last_visit["Data Entered Latest Date"] = _farm_last_visit["_LatestDateParsed"].dt.strftime(
        "%Y-%m-%d"
    ).fillna("-")
    # "_DueDateNum" (numeric days-elapsed, kept only for filtering below)
    # is computed alongside the display "Due Date" string so a farm with
    # no parsed date at all (shown as "-") is treated as not overdue
    # rather than crashing the > 7 comparison.
    _farm_last_visit["_DueDateNum"] = _farm_last_visit["_LatestDateParsed"].apply(
        lambda d: (_today_last_visit - d).days if pd.notna(d) else None
    )
    _farm_last_visit["Due Date"] = _farm_last_visit["_DueDateNum"].apply(
        lambda n: str(n) if n is not None else "-"
    )
    _farm_last_visit = _farm_last_visit.drop(columns=["_LatestDateParsed"])

    # Only farms that are still Running AND overdue by more than 7 days
    # get shown in this report — FULL H farms and farms visited within
    # the last 7 days are filtered out here, before the Zone filter below.
    _farm_last_visit = _farm_last_visit[
        (_farm_last_visit["Status"] == "Running")
        & (_farm_last_visit["_DueDateNum"].notna())
        & (_farm_last_visit["_DueDateNum"] > 7)
    ].drop(columns=["_DueDateNum"])

    # Attach Zone (from Customer List.xlsx), same lookup pattern used
    # elsewhere in this file.
    _zone_lookup_last_visit = customer_df[["Customer Name", "Farm Name with Code", "Zone"]].drop_duplicates(
        subset=["Customer Name", "Farm Name with Code"]
    ).rename(columns={"Customer Name": "Customer"})
    _farm_last_visit = _farm_last_visit.merge(
        _zone_lookup_last_visit, on=["Customer", "Farm Name with Code"], how="left"
    )
    _farm_last_visit = _farm_last_visit.rename(columns={"Customer": "Customer Name"})

    _last_visit_display_cols = [
        "Customer Name", "Farm Name with Code", "Zone", "Status",
        "Data Entered Latest Date", "Due Date",
    ]

    _zones_last_visit = sorted(
        {str(z).strip() for z in _farm_last_visit["Zone"].tolist()
         if str(z).strip() and str(z).strip().lower() != "nan"}
    )
    if _zones_last_visit:
        _selected_zones_last_visit = st.multiselect(
            "Select Zone(s)     ", options=_zones_last_visit, default=_zones_last_visit,
            key="last_visit_zone_filter"
        )
        if not _selected_zones_last_visit:
            st.info("Select at least one zone above to display the last visit date report.")
        else:
            _filtered_last_visit = _farm_last_visit[
                _farm_last_visit["Zone"].astype(str).str.strip().isin(_selected_zones_last_visit)
            ].sort_values(by=["Customer Name", "Farm Name with Code"])
            st.dataframe(
                _filtered_last_visit[_last_visit_display_cols], use_container_width=True, hide_index=True
            )
            st.caption(f"{len(_filtered_last_visit)} farm(s) shown.")
    else:
        _farm_last_visit = _farm_last_visit.sort_values(by=["Customer Name", "Farm Name with Code"])
        st.dataframe(_farm_last_visit[_last_visit_display_cols], use_container_width=True, hide_index=True)
        st.caption("No Zone information found on the customer list — showing unfiltered.")
else:
    st.info("No records available yet to build the last visit date report.")

# =========================================================================
# 🤖 CHAT WITH ANJI — NEW SECTION. A read-only chatbot that answers common
# marketing questions straight from the same Google Sheets data and
# "Customer List.xlsx" used by every section above. It performs NO writes.
#
# Questions it understands (type them naturally, or use the quick buttons):
#   1) "Who are the customers in DOC range 30-60?"
#        -> Zone x V/M summary + list of Customer/Farm, species, pond count.
#           (Only Running / Partial H ponds are counted — Full H and
#           "Soon to be" ponds are left out. DOC = each pond's DOC Today.)
#   2) "What is the last order of <customer / farm>?"
#        -> Last feed order items + Last Feed Purchased Date.
#   3) "Who are the farmers harvested this week?"
#        -> Harvest events dated Monday-to-today of the current week.
#   4) "What is the situation of <customer / farm> right now?"
#        -> Pond-by-pond status, DOC Today, expected harvest, last visit.
#   5) "Who are the customers that did not purchase Osmin?"
#        -> Customer + Farm entries with no matching item in Sales Details.
#
# Nothing above this section was changed.
# =========================================================================
st.markdown("---")
st.markdown("#### 🤖 Chat with Anji")
st.caption(
    "Ask Anji about DOC ranges, last feed orders, this week's harvests, a farm's current "
    "situation, or who hasn't purchased a product (e.g. Osmin). Read-only — Anji never changes any data."
)

_ANJI_ZONE_ORDER = ["Chilaw", "Middle Zone", "Puttalam", "Batticaloa"]


def _anji_norm_tokens(text):
    return set(re.findall(r"[a-z0-9]+", str(text).lower()))


def _anji_customer_code_of(row):
    for _cand in _CUSTOMER_CODE_COLUMN_CANDIDATES:
        if _cand in customer_df.columns:
            _v = str(row.get(_cand, "")).strip()
            if _v and _v.lower() != "nan":
                return _v
    return ""


def _anji_species_letter(species):
    _s = str(species).strip().lower()
    if "vannamei" in _s:
        return "V"
    if "monodon" in _s:
        return "M"
    return "-"


def _anji_doc_today(row):
    """Same rule as 'DOC Today' used above (Soon to be = 0, stops counting
    at a Full Harvest date, otherwise saved DOC + days since the record's
    Date). Returns None if it can't be worked out."""
    if str(row.get("Cycle Type") or "").strip() == "Soon to be":
        return 0
    parsed = pd.to_datetime(row.get("Date"), errors="coerce")
    if pd.isna(parsed):
        return None
    try:
        doc_num = int(float(row.get("DOC")))
    except (TypeError, ValueError):
        return None
    _t2 = str(row.get("Harvest Type 2", "")).strip().lower()
    _t1 = str(row.get("Harvest Type", "")).strip().lower()
    _full_date_str = ""
    if "full" in _t2:
        _full_date_str = str(row.get("Harvest Date 2", "")).strip()
    elif "full" in _t1:
        _full_date_str = str(row.get("Harvest Date", "")).strip()
    if _full_date_str:
        _full_date = pd.to_datetime(_full_date_str, errors="coerce")
        if pd.notna(_full_date):
            return doc_num + (_full_date - parsed).days
    return doc_num + (pd.Timestamp(date.today()) - parsed).days


def _anji_latest_ponds():
    """Latest saved record per Customer + Farm + Pond, with Zone, status
    (Running / Partial H / Full H / Soon to be), species letter and DOC Today."""
    _df = load_data()
    if len(_df) == 0:
        return pd.DataFrame()
    _df = _df.copy()
    _df["_ParsedDate"] = pd.to_datetime(_df["Date"], errors="coerce")

    _partial_hist = (
        _df.assign(
            _HasPartial=(
                _df["Harvest Type"].astype(str).str.lower().str.contains("partial")
                | _df["Harvest Type 2"].astype(str).str.lower().str.contains("partial")
            )
        )
        .groupby(["Customer", "Farm Name with Code", "Pond Number"])["_HasPartial"]
        .any()
    )

    _latest = (
        _df.dropna(subset=["_ParsedDate"])
        .sort_values("_ParsedDate")
        .groupby(["Customer", "Farm Name with Code", "Pond Number"], as_index=False)
        .last()
    )

    def _status(r):
        _t = str(r.get("Harvest Type 2", "")).strip() or str(r.get("Harvest Type", "")).strip()
        if "full" in _t.lower():
            return "Full H"
        if "partial" in _t.lower() or bool(
            _partial_hist.get((r["Customer"], r["Farm Name with Code"], r["Pond Number"]), False)
        ):
            return "Partial H"
        if str(r.get("Cycle Type", "")).strip() == "Soon to be":
            return "Soon to be"
        return "Running"

    _latest["Status"] = _latest.apply(_status, axis=1)
    _latest["V/M"] = _latest["Species Culture"].apply(_anji_species_letter)
    _latest["DOC Today"] = _latest.apply(_anji_doc_today, axis=1)

    _zone_lookup = (
        customer_df[["Customer Name", "Farm Name with Code", "Zone"]]
        .drop_duplicates(subset=["Customer Name", "Farm Name with Code"])
        .rename(columns={"Customer Name": "Customer", "Zone": "_CustZone"})
    )
    _latest = _latest.merge(_zone_lookup, on=["Customer", "Farm Name with Code"], how="left")
    _latest["Zone"] = _latest["_CustZone"].fillna("").astype(str).str.strip()
    _blank_zone = _latest["Zone"].eq("") | _latest["Zone"].str.lower().eq("nan")
    _latest.loc[_blank_zone, "Zone"] = (
        _latest.loc[_blank_zone, "Farm Name with Code"].map(lambda _x: "")  # placeholder, replaced below
    )
    # Fall back to the sheet's own Zone column when the customer list has none.
    if "Zone_x" in _latest.columns or "Zone" in _df.columns:
        _sheet_zone = (
            _df.sort_values("_ParsedDate")
            .groupby(["Customer", "Farm Name with Code"])["Zone"].last()
            .astype(str).str.strip()
        )
        _latest.loc[_blank_zone, "Zone"] = [
            _sheet_zone.get((c, f), "")
            for c, f in zip(_latest.loc[_blank_zone, "Customer"], _latest.loc[_blank_zone, "Farm Name with Code"])
        ]
    _latest["Zone"] = _latest["Zone"].replace("", "Unknown")
    return _latest


def _anji_find_farm(question):
    """Best-matching Customer + Farm row from the customer list for the words
    in the question. Returns (customer, farm, customer_code) or None."""
    _qtok = _anji_norm_tokens(question)
    _skip = {"farm", "the", "of", "is", "what", "last", "order", "situation", "right", "now", "for"}
    _best, _best_score = None, 0.0
    _farms = customer_df.drop_duplicates(subset=["Customer Name", "Farm Name with Code"])
    for _, _r in _farms.iterrows():
        _toks = _anji_norm_tokens(f"{_r['Customer Name']} {_r['Farm Name with Code']}") - _skip
        if not _toks:
            continue
        _hit = len(_toks & (_qtok - _skip))
        if _hit == 0:
            continue
        _ratio = _hit / len(_toks)
        _score = _ratio * 10 + _hit
        if _ratio >= 0.5 and _score > _best_score:
            _best_score = _score
            _best = (_r["Customer Name"], _r["Farm Name with Code"], _anji_customer_code_of(_r))
    return _best


def _anji_answer_doc_range(question):
    _m = re.search(r"(\d+)\s*(?:-|–|to|and)\s*(\d+)", question.lower())
    if not _m:
        return {"text": "Please give me a DOC range, e.g. *Who are the customers in DOC range 30-60?*", "tables": []}
    _lo, _hi = sorted((int(_m.group(1)), int(_m.group(2))))
    _ponds = _anji_latest_ponds()
    if len(_ponds) == 0:
        return {"text": "There are no saved pond records yet.", "tables": []}
    _hit = _ponds[
        _ponds["Status"].isin(["Running", "Partial H"])
        & _ponds["DOC Today"].notna()
        & (_ponds["DOC Today"] >= _lo)
        & (_ponds["DOC Today"] <= _hi)
    ].copy()
    if len(_hit) == 0:
        return {"text": f"No running ponds are currently in the DOC range {_lo}-{_hi}.", "tables": []}

    _zones = _ANJI_ZONE_ORDER + sorted(
        z for z in _hit["Zone"].unique() if z.lower() not in [x.lower() for x in _ANJI_ZONE_ORDER]
    )
    _summary_rows = []
    for _z in _zones:
        _zsub = _hit[_hit["Zone"].str.lower() == _z.lower()]
        if len(_zsub) == 0 and _z not in _ANJI_ZONE_ORDER:
            continue
        _v = int((_zsub["V/M"] == "V").sum())
        _mm = int((_zsub["V/M"] == "M").sum())
        _summary_rows.append({"Zone": _z, "V": _v, "M": _mm, "Total": len(_zsub)})
    _summary = pd.DataFrame(_summary_rows)
    _total_row = {"Zone": "Total", "V": int(_summary["V"].sum()), "M": int(_summary["M"].sum()),
                  "Total": int(_summary["Total"].sum())}
    _summary = pd.concat([_summary, pd.DataFrame([_total_row])], ignore_index=True)

    _detail = (
        _hit.groupby(["Zone", "Customer", "Farm Name with Code", "V/M"])
        .agg(**{"No of Ponds": ("Pond Number", "nunique")})
        .reset_index()
    )
    _detail["Customer Name with Farm Name"] = _detail["Customer"] + " — " + _detail["Farm Name with Code"]
    _detail = _detail.sort_values(["Zone", "Customer Name with Farm Name"])[
        ["Zone", "Customer Name with Farm Name", "V/M", "No of Ponds"]
    ]
    _n_farms = _detail[["Zone", "Customer Name with Farm Name"]].drop_duplicates().shape[0]
    return {
        "text": (f"**{_n_farms} farm(s)** have running ponds in DOC range **{_lo}-{_hi}** "
                 f"({len(_hit)} pond(s) in total). DOC is each pond's DOC Today; Full H and "
                 "Soon to be ponds are excluded."),
        "tables": [("Pond count by Zone (V = Vannamei, M = Monodon)", _summary),
                   ("Customers / Farms in range", _detail)],
    }


def _anji_answer_last_order(question):
    _found = _anji_find_farm(question)
    if not _found:
        return {"text": "I couldn't find that customer/farm. Please type the customer and farm name, "
                        "e.g. *Last order of Ruwan Chinthaka Ruwan Farm 1*.", "tables": []}
    _cust, _farm, _code = _found
    if not _code:
        return {"text": f"**{_cust} — {_farm}** has no Customer Code in 'Customer List.xlsx', "
                        "so I can't look up its sales.", "tables": []}
    _sales = load_sales_data()
    if len(_sales) == 0:
        return {"text": "The Sales Details sheet has no records.", "tables": []}
    _s = _sales[_sales["Customer Code"].astype(str).str.strip().str.lower() == _code.strip().lower()].copy()
    _s["Quantity"] = pd.to_numeric(_s["Quantity"], errors="coerce").fillna(0)
    _s["_D"] = pd.to_datetime(_s["Date"], errors="coerce")
    _feed = _s[_s["Item No."].astype(str).str.strip().str.upper().str.startswith("FEED") & _s["_D"].notna()]
    if len(_feed) == 0:
        return {"text": f"**{_cust} — {_farm}** (code {_code}) has no feed purchases on record.", "tables": []}
    _last_d = _feed["_D"].max()
    _items = _feed[_feed["_D"] == _last_d].groupby("Item Description")["Quantity"].sum()
    _tbl = pd.DataFrame({"Last Order": [f"{i} - {q:,.0f}" for i, q in _items.items()],
                         "Last Feed Purchased Date": [_last_d.strftime("%Y-%m-%d")] * len(_items)})
    return {"text": f"Last feed order for **{_cust} — {_farm}** (code {_code}):", "tables": [("", _tbl)]}


def _anji_answer_harvested_this_week(question):
    _df = load_data()
    if len(_df) == 0:
        return {"text": "There are no saved records yet.", "tables": []}
    _today = pd.Timestamp(date.today())
    _week_start = _today - pd.Timedelta(days=_today.weekday())  # Monday
    _rows = []
    for _, _r in _df.iterrows():
        for _slot, (_dc, _tc, _kc, _ac, _sc) in enumerate(
            [("Harvest Date", "Harvest Type", "Harvest KG", "Harvest ABW", "Harvest Status"),
             ("Harvest Date 2", "Harvest Type 2", "Harvest KG 2", "Harvest ABW 2", "Harvest Status 2")], start=1
        ):
            _hd = pd.to_datetime(str(_r.get(_dc, "")).strip(), errors="coerce")
            if pd.isna(_hd) or not (_week_start <= _hd <= _today):
                continue
            if str(_r.get(_sc, "")).strip().upper() == "H":
                continue
            _rows.append({
                "Customer": _r.get("Customer", ""), "Farm Name with Code": _r.get("Farm Name with Code", ""),
                "Pond Number": _r.get("Pond Number", ""), "Harvest Date": _hd.strftime("%Y-%m-%d"),
                "Harvest Type": _r.get(_tc, ""), "Harvest KG": _r.get(_kc, ""), "Harvest ABW": _r.get(_ac, ""),
            })
    if not _rows:
        return {"text": f"No harvests recorded this week ({_week_start:%Y-%m-%d} to {_today:%Y-%m-%d}).", "tables": []}
    _out = pd.DataFrame(_rows).drop_duplicates().sort_values(["Harvest Date", "Customer", "Farm Name with Code"])
    _n_farms = _out[["Customer", "Farm Name with Code"]].drop_duplicates().shape[0]
    return {"text": f"**{_n_farms} farm(s)** harvested this week ({_week_start:%Y-%m-%d} to {_today:%Y-%m-%d}):",
            "tables": [("", _out.reset_index(drop=True))]}


def _anji_answer_situation(question):
    _found = _anji_find_farm(question)
    if not _found:
        return {"text": "I couldn't find that farm. Please type the customer and farm name, "
                        "e.g. *What is the situation of Jais ASR Farm?*", "tables": []}
    _cust, _farm, _code = _found
    _ponds = _anji_latest_ponds()
    if len(_ponds) == 0:
        return {"text": "There are no saved pond records yet.", "tables": []}
    _p = _ponds[(_ponds["Customer"] == _cust) & (_ponds["Farm Name with Code"] == _farm)].copy()
    if len(_p) == 0:
        return {"text": f"No saved pond records yet for **{_cust} — {_farm}**.", "tables": []}
    _p["_PN"] = pd.to_numeric(_p["Pond Number"], errors="coerce")
    _p = _p.sort_values(["_PN", "Pond Number"])
    _last_visit = _p["_ParsedDate"].max()
    _days = (pd.Timestamp(date.today()) - _last_visit).days if pd.notna(_last_visit) else None
    _counts = _p["Status"].value_counts().to_dict()
    _tbl = pd.DataFrame({
        "Pond": _p["Pond Number"], "V/M": _p["V/M"], "Status": _p["Status"],
        "DOC Today": _p["DOC Today"].apply(lambda v: "" if v is None or pd.isna(v) else int(v)),
        "ABW": _p["ABW"], "Feed/Day": _p["Feed Per Day"],
        "Expect Harvest (KG)": _p["Expect Harvest (KG)"], "Issues": _p["Issues"],
        "Last Visit": _p["_ParsedDate"].dt.strftime("%Y-%m-%d"),
    })
    _summary = ", ".join(f"{v} {k}" for k, v in _counts.items())
    _visit_txt = (f" Last data entry: **{_last_visit:%Y-%m-%d}** ({_days} day(s) ago)."
                  if _days is not None else "")
    _expect_total = pd.to_numeric(
        _p.loc[_p["Status"].isin(["Running", "Partial H"]), "Expect Harvest (KG)"], errors="coerce"
    ).sum()
    _expect_txt = f" Total expected harvest from active ponds: **{_expect_total:,.0f} kg**." if _expect_total else ""
    return {"text": f"**{_cust} — {_farm}** has {len(_p)} pond(s): {_summary}.{_visit_txt}{_expect_txt}",
            "tables": [("", _tbl.reset_index(drop=True))]}


def _anji_answer_not_purchased(question):
    _m = re.search(r"(?:not|n't|never)\s+purchas\w*\s+(.+?)\s*\??$", question.lower().strip())
    _product = "osmin"
    if _m:
        _cand = re.sub(r"\b(any|the|a)\b", "", _m.group(1)).strip(" ?.")
        if _cand:
            _product = _cand
    elif "osmin" not in question.lower():
        _after = re.search(r"purchas\w*\s+(.+?)\s*\??$", question.lower().strip())
        if _after:
            _product = re.sub(r"\b(any|the|a)\b", "", _after.group(1)).strip(" ?.") or _product
    _sales = load_sales_data()
    if len(_sales) == 0:
        return {"text": "The Sales Details sheet has no records.", "tables": []}
    _has_code_col = any(c in customer_df.columns for c in _CUSTOMER_CODE_COLUMN_CANDIDATES)
    if not _has_code_col:
        return {"text": "'Customer List.xlsx' has no Customer Code column, so I can't match customers to sales.",
                "tables": []}
    _desc = _sales["Item Description"].astype(str).str.lower()
    _buyers = set(
        _sales.loc[_desc.str.contains(re.escape(_product), na=False), "Customer Code"]
        .astype(str).str.strip().str.lower()
    )
    _farms = customer_df.drop_duplicates(subset=["Customer Name", "Farm Name with Code"]).copy()
    _farms["_Code"] = _farms.apply(_anji_customer_code_of, axis=1)
    _no_code = int((_farms["_Code"] == "").sum())
    _farms = _farms[_farms["_Code"] != ""]
    _missing = _farms[~_farms["_Code"].str.lower().isin(_buyers)].copy()
    if len(_missing) == 0:
        return {"text": f"Every customer in the list has purchased **{_product.title()}**.", "tables": []}
    _out = _missing.rename(columns={"Customer Name": "Customer Name"})[
        ["Zone", "Customer Name", "Farm Name with Code", "_Code"]
    ].rename(columns={"_Code": "Customer Code"}).sort_values(["Zone", "Customer Name", "Farm Name with Code"])
    _note = f" ({_no_code} farm(s) without a Customer Code were skipped.)" if _no_code else ""
    return {"text": (f"**{len(_out)} customer/farm(s)** in Sales have not purchased any "
                     f"**{_product.title()}** item.{_note}"),
            "tables": [("", _out.reset_index(drop=True))]}


def _anji_answer(question):
    _q = question.lower().strip()
    try:
        if re.search(r"(not|n't|never|haven|without)\s+purchas", _q) or "do not purchase" in _q:
            return _anji_answer_not_purchased(question)
        if "harvest" in _q and ("this week" in _q or "week" in _q):
            return _anji_answer_harvested_this_week(question)
        if "doc" in _q and re.search(r"\d+\s*(?:-|–|to|and)\s*\d+", _q):
            return _anji_answer_doc_range(question)
        if "last order" in _q or "last feed" in _q or "last purchase" in _q:
            return _anji_answer_last_order(question)
        if "situation" in _q or "status" in _q or "how is" in _q or "update" in _q:
            return _anji_answer_situation(question)
        # Fallback: if a farm name is mentioned, give its situation.
        if _anji_find_farm(question):
            return _anji_answer_situation(question)
    except Exception as _e:
        return {"text": f"Sorry, I ran into a problem while looking that up: {_e}", "tables": []}
    return {
        "text": ("I can help with:\n"
                 "- *Who are the customers in DOC range 30-60?*\n"
                 "- *What is the last order of <customer farm>?*\n"
                 "- *Who harvested this week?*\n"
                 "- *What is the situation of <customer farm> right now?*\n"
                 "- *Who has not purchased Osmin?*"),
        "tables": [],
    }


_ANJI_GREETING = {
    "role": "assistant",
    "content": "Hi, I'm Anji 👋 Ask me a question below or tap one of the quick questions.",
    "tables": [],
}
if "anji_messages" not in st.session_state:
    st.session_state["anji_messages"] = [_ANJI_GREETING]

_anji_quick = [
    ("DOC range 30-60", "Who are the customers in DOC range 30-60?"),
    ("Last order", "What is the last order of Ruwan Chinthaka Ruwan Farm 1?"),
    ("Harvested this week", "Who are the farmers harvested this week?"),
    ("Jais ASR Farm", "What is the situation of Jais ASR Farm right now?"),
    ("Not purchased Osmin", "Who are the customers that have not purchased Osmin - 1Kg Pck?"),
]
_anji_btn_cols = st.columns(len(_anji_quick))
_anji_clicked = None
for _i, (_label, _qtext) in enumerate(_anji_quick):
    with _anji_btn_cols[_i]:
        if st.button(_label, key=f"anji_quick_{_i}", use_container_width=True):
            _anji_clicked = _qtext

if st.button("🧹 Clear chat", key="anji_clear"):
    st.session_state["anji_messages"] = [_ANJI_GREETING]
    st.rerun()

_anji_typed = st.chat_input("Chat with Anji — ask a question...", key="anji_chat_input")
_anji_prompt = _anji_typed or _anji_clicked

if _anji_prompt:
    st.session_state["anji_messages"].append({"role": "user", "content": _anji_prompt, "tables": []})
    with st.spinner("Anji is checking the data..."):
        _anji_reply = _anji_answer(_anji_prompt)
    st.session_state["anji_messages"].append(
        {"role": "assistant", "content": _anji_reply["text"], "tables": _anji_reply["tables"]}
    )

for _msg in st.session_state["anji_messages"]:
    with st.chat_message(_msg["role"]):
        st.markdown(_msg["content"])
        for _title, _tbl in _msg.get("tables", []):
            if _title:
                st.markdown(f"**{_title}**")
            st.dataframe(_tbl, use_container_width=True, hide_index=True)

st.markdown("---")
st.markdown(
    "<p style='text-align: center; color: gray;'>KMN Aqua Services - Water Quality Monitoring System "
    "(Marketing Manager & Technical Officers View — read only)</p>",
    unsafe_allow_html=True,
)
