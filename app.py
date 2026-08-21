
import io
import hashlib
from datetime import datetime

import numpy as np
import pandas as pd
import streamlit as st

# ==========================================================
# PAGE CONFIG
# ==========================================================
st.set_page_config(
    page_title="Stock Aging & Refill Analytics",
    page_icon="📦",
    layout="wide",
)

st.markdown(
    """
    <style>
    .main { background-color: #f7f8fa; }
    .stMetric {
        background-color: #ffffff;
        padding: 14px;
        border-radius: 10px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.08);
    }
    </style>
    """,
    unsafe_allow_html=True,
)

MAX_PREVIEW_ROWS = 2000  # cap for on-screen tables so the browser never chokes
NONE_OPTION = "-- None / Not Available --"

# ---- Fixed business thresholds (not user-configurable) ----
# Performance (Selling Days) tag bands
PERF_BANDS = [(15, "Fast"), (30, "Good"), (60, "Moderate"), (90, "Slow")]
PERF_LAST_TAG = "Very Slow"


# ==========================================================
# GENERIC HELPERS
# ==========================================================
@st.cache_data(show_spinner=False)
def get_sheet_names(file_bytes: bytes, filename: str):
    if filename.lower().endswith(".csv"):
        return ["Sheet1"]
    buffer = io.BytesIO(file_bytes)
    try:
        xls = pd.ExcelFile(buffer, engine="calamine")
    except Exception:
        buffer.seek(0)
        xls = pd.ExcelFile(buffer, engine="openpyxl")
    return xls.sheet_names


@st.cache_data(show_spinner=False)
def read_data_file(file_bytes: bytes, filename: str, sheet_name: str = None) -> pd.DataFrame:
    """Fast, cached reader. Tries calamine (Rust engine) first for large xlsx files,
    falls back to openpyxl, and uses the C engine for csv."""
    buffer = io.BytesIO(file_bytes)
    if filename.lower().endswith(".csv"):
        df = pd.read_csv(buffer, low_memory=False)
    else:
        try:
            df = pd.read_excel(buffer, sheet_name=sheet_name, engine="calamine")
        except Exception:
            buffer.seek(0)
            df = pd.read_excel(buffer, sheet_name=sheet_name, engine="openpyxl")
    return df


def clean_text_series(series: pd.Series) -> pd.Series:
    if series is None:
        return pd.Series(dtype=str)
    return (
        series.astype(str)
        .str.strip()
        .str.replace(r"\s+", " ", regex=True)
        .str.upper()
        .replace({"NAN": "", "NONE": ""})
    )


def clean_barcode_series(series: pd.Series) -> pd.Series:
    """Barcodes / EAN codes should be compared as clean strings, not floats."""
    s = series.astype(str).str.strip()
    s = s.str.replace(r"\.0$", "", regex=True)  # strip trailing .0 from float-read numbers
    return s


def parse_smart_dates(series: pd.Series) -> pd.Series:
    """Handles mixed date formats in the same column:
    - Excel serial numbers (e.g. 45678)
    - ISO strings (e.g. 2026-01-02 14:07:27) -> must NOT use dayfirst
    - Indian-style strings (e.g. 30/05/2026) -> ambiguous, needs dayfirst=True
    Tries each strategy in order and only falls back for values still unparsed.
    """
    s_str = series.astype(str).str.strip()
    dates = pd.Series(pd.NaT, index=series.index)

    s_num = pd.to_numeric(s_str, errors="coerce")
    is_serial = s_num.notna() & (s_num > 20000) & (s_num < 90000)
    dates.loc[is_serial] = pd.to_datetime(s_num[is_serial], unit="D", origin="1899-12-30", errors="coerce")

    remaining = dates.isna()
    dates.loc[remaining] = pd.to_datetime(s_str[remaining], errors="coerce", dayfirst=False)

    remaining = dates.isna()
    dates.loc[remaining] = pd.to_datetime(s_str[remaining], errors="coerce", dayfirst=True)

    return dates


def auto_suggest(cols, keywords):
    for col in cols:
        col_lower = str(col).lower()
        for kw in keywords:
            if kw in col_lower:
                return col
    return None


def mapping_selectbox(label, cols, keywords, key, allow_none=False):
    options = ([NONE_OPTION] + list(cols)) if allow_none else list(cols)
    suggestion = auto_suggest(cols, keywords)
    if suggestion and suggestion in options:
        default_index = options.index(suggestion)
    else:
        default_index = 0
    return st.selectbox(label, options, index=default_index, key=key)


def performance_tag(days):
    if pd.isna(days):
        return "Unknown"
    days = max(float(days), 0)
    for limit, tag in PERF_BANDS:
        if days <= limit:
            return tag
    return PERF_LAST_TAG


def resolve_text(*series_list):
    """First non-empty value across several optional text columns (priority order)."""
    result = None
    for s in series_list:
        if s is None:
            continue
        s = s.replace("", np.nan)
        result = s if result is None else result.fillna(s)
    return result.fillna("") if result is not None else pd.Series(dtype=str)


def format_dates_for_display(df: pd.DataFrame, date_cols) -> pd.DataFrame:
    """Returns a COPY of df with the given datetime columns rendered as
    dd-mm-yyyy strings, for on-screen preview only (date only, no time-of-day).
    The original dataframe (used for the Excel export) is left untouched."""
    display_df = df.copy()
    for col in date_cols:
        if col in display_df.columns:
            display_df[col] = pd.to_datetime(display_df[col], errors="coerce").dt.strftime("%d-%m-%Y")
            display_df[col] = display_df[col].replace("NaT", "")
    return display_df


# ==========================================================
# SESSION STATE
# ==========================================================
if "results" not in st.session_state:
    st.session_state.results = None  # holds the cached output of the last successful "Generate Report" run

# ==========================================================
# SIDEBAR - SETTINGS
# ==========================================================
st.sidebar.title("⚙️ Settings")

st.sidebar.subheader("Report Date")
report_date = st.sidebar.date_input("Report As-Of Date (used only for Age Days)", datetime.today())
report_date_dt = pd.to_datetime(report_date)

st.sidebar.markdown("---")
st.sidebar.subheader("Refilling Age Cutoffs")
fresh_cutoff = st.sidebar.number_input(
    "Fresh Cutoff (days)", min_value=1, max_value=365, value=20, step=1,
    help="Age Days from 0 up to this value = Fresh",
)
rtv_cutoff = st.sidebar.number_input(
    "RTV / Decision Cutoff (days)", min_value=fresh_cutoff + 1, max_value=730, value=40, step=1,
    help="Age Days above this value triggers the RTV / Refill / Sold-No-Refill decision. "
         "Between Fresh Cutoff and this value = Monitor.",
)

st.sidebar.markdown("---")
st.sidebar.caption(
    f"**Performance tags:** 0-15 Fast | 16-30 Good | 31-60 Moderate | 61-90 Slow | 91+ Very Slow\n\n"
    f"**Refilling ages:** 0-{fresh_cutoff} Fresh | {fresh_cutoff}-{rtv_cutoff} Monitor | "
    f"{rtv_cutoff}+ decision tree (see Refilling sheet)"
)

st.sidebar.markdown("---")
if st.sidebar.button("🔄 Reset Application"):
    for key in list(st.session_state.keys()):
        del st.session_state[key]
    st.rerun()

# ==========================================================
# MAIN - FILE UPLOADS
# ==========================================================
st.title("📦 Stock Aging & Refill Analytics")
st.markdown(
    "Matches **Outward + Sales + Current Stock** (+ optional **WH Stock**) by Barcode & Store. "
    "Current Stock file is always the final truth for current stock (never Outward − Sales)."
)

col1, col2, col3, col4 = st.columns(4)
with col1:
    st.subheader("A. Outward File")
    f_out = st.file_uploader("Upload Outward File", type=["xlsx", "xls", "csv"], key="out_file")
with col2:
    st.subheader("B. Sales File")
    f_sal = st.file_uploader("Upload Sales File", type=["xlsx", "xls", "csv"], key="sal_file")
with col3:
    st.subheader("C. Current Stock File")
    f_stk = st.file_uploader("Upload Current Stock File", type=["xlsx", "xls", "csv"], key="stk_file")
with col4:
    st.subheader("D. WH Stock File (optional)")
    f_wh = st.file_uploader("Upload WH Stock File", type=["xlsx", "xls", "csv"], key="wh_file")


def load_with_sheet_picker(uploaded_file, label_prefix):
    if uploaded_file is None:
        return None, None
    fbytes = uploaded_file.getvalue()
    fname = uploaded_file.name
    sheets = get_sheet_names(fbytes, fname)
    if len(sheets) > 1:
        sheet = st.sidebar.selectbox(f"{label_prefix} - Sheet", sheets, key=f"{label_prefix}_sheet")
    else:
        sheet = sheets[0]
    df = read_data_file(fbytes, fname, sheet)
    return df, fname


with st.spinner("Reading files..."):
    out_df_raw, out_name = load_with_sheet_picker(f_out, "Outward")
    sal_df_raw, sal_name = load_with_sheet_picker(f_sal, "Sales")
    stk_df_raw, stk_name = load_with_sheet_picker(f_stk, "Stock")
    wh_df_raw, wh_name = load_with_sheet_picker(f_wh, "WH")

if out_df_raw is not None:
    with st.expander(f"📄 Outward Preview - {out_name} ({len(out_df_raw):,} rows)"):
        st.dataframe(out_df_raw.head(5), use_container_width=True)

if sal_df_raw is not None:
    with st.expander(f"📄 Sales Preview - {sal_name} ({len(sal_df_raw):,} rows)"):
        st.dataframe(sal_df_raw.head(5), use_container_width=True)

if stk_df_raw is not None:
    with st.expander(f"📄 Stock Preview - {stk_name} ({len(stk_df_raw):,} rows)"):
        st.dataframe(stk_df_raw.head(5), use_container_width=True)

if wh_df_raw is not None:
    with st.expander(f"📄 WH Stock Preview - {wh_name} ({len(wh_df_raw):,} rows)"):
        st.dataframe(wh_df_raw.head(5), use_container_width=True)

# ==========================================================
# COLUMN MAPPING
# ==========================================================
if out_df_raw is not None and sal_df_raw is not None and stk_df_raw is not None:
    st.markdown("---")
    st.subheader("⚙️ Column Mapping")

    out_cols = list(out_df_raw.columns)
    sal_cols = list(sal_df_raw.columns)
    stk_cols = list(stk_df_raw.columns)
    wh_cols = list(wh_df_raw.columns) if wh_df_raw is not None else []

    m1, m2, m3 = st.columns(3)

    with m1:
        st.markdown("#### Outward Columns")
        m_out_store = mapping_selectbox("Store", out_cols, ["code", "store", "site"], "m_out_store")
        m_out_barcode = mapping_selectbox("Barcode / EAN", out_cols, ["barcode", "ean", "upc"], "m_out_barcode")
        m_out_qty = mapping_selectbox("Outward Qty", out_cols, ["qty", "quantity"], "m_out_qty")
        m_out_date = mapping_selectbox("Outward Date", out_cols, ["outward on", "outward date", "dispatch"], "m_out_date")
        m_out_article = mapping_selectbox("Article (optional)", out_cols, ["article"], "m_out_article", allow_none=True)
        m_out_color = mapping_selectbox("Color (optional)", out_cols, ["color", "colour"], "m_out_color", allow_none=True)
        m_out_size = mapping_selectbox("Size (optional)", out_cols, ["size"], "m_out_size", allow_none=True)

    with m2:
        st.markdown("#### Sales Columns")
        m_sal_store = mapping_selectbox("Store", sal_cols, ["site", "store"], "m_sal_store")
        m_sal_barcode = mapping_selectbox("Barcode / EAN", sal_cols, ["ean", "upc", "barcode"], "m_sal_barcode")
        m_sal_qty = mapping_selectbox("Sales Qty", sal_cols, ["qty", "quantity"], "m_sal_qty")
        m_sal_date = mapping_selectbox("Sale Date", sal_cols, ["calendar_day", "sale date", "date"], "m_sal_date")
        m_sal_article = mapping_selectbox("Article (optional)", sal_cols, ["article"], "m_sal_article", allow_none=True)
        m_sal_color = mapping_selectbox("Color (optional)", sal_cols, ["color", "colour"], "m_sal_color", allow_none=True)
        m_sal_size = mapping_selectbox("Size (optional)", sal_cols, ["size"], "m_sal_size", allow_none=True)

    with m3:
        st.markdown("#### Stock Columns")
        m_stk_store = mapping_selectbox("Store", stk_cols, ["site", "store"], "m_stk_store")
        m_stk_barcode = mapping_selectbox("Barcode / EAN", stk_cols, ["ean", "upc", "barcode"], "m_stk_barcode")
        m_stk_qty = mapping_selectbox("Current Stock Qty", stk_cols, ["stock_qty", "stock", "balance"], "m_stk_qty")
        m_stk_article = mapping_selectbox("Article (optional)", stk_cols, ["article"], "m_stk_article", allow_none=True)

    m_wh_barcode = None
    m_wh_qty = None
    if wh_df_raw is not None:
        st.markdown("#### WH Stock Columns (matched by Barcode only, no Store)")
        m4a, m4b = st.columns(2)
        with m4a:
            m_wh_barcode = mapping_selectbox("Barcode / EAN", wh_cols, ["barcode", "ean", "upc"], "m_wh_barcode")
        with m4b:
            m_wh_qty = mapping_selectbox("WH Qty", wh_cols, ["qty", "quantity"], "m_wh_qty")

    st.markdown("---")
    gen_btn = st.button("🚀 Generate Report", type="primary")

    # ==========================================================
    # PROCESSING - runs ONLY when the button is clicked.
    # Result is cached in st.session_state.results so that a later
    # rerun (e.g. clicking the download button) does not rebuild anything.
    # ==========================================================
    if gen_btn:
        warnings = []

        with st.spinner("Cleaning, aggregating and building the 3 sheets..."):

            # -------- Outward: clean + aggregate once --------
            df_out = out_df_raw.copy()
            df_out["Store"] = clean_text_series(df_out[m_out_store])
            df_out["Barcode"] = clean_barcode_series(df_out[m_out_barcode])
            df_out["Qty"] = pd.to_numeric(df_out[m_out_qty], errors="coerce").fillna(0)
            df_out["Outward_Date"] = parse_smart_dates(df_out[m_out_date])
            df_out["Article"] = clean_text_series(df_out[m_out_article]) if m_out_article != NONE_OPTION else ""
            df_out["Color"] = clean_text_series(df_out[m_out_color]) if m_out_color != NONE_OPTION else ""
            df_out["Size"] = clean_text_series(df_out[m_out_size]) if m_out_size != NONE_OPTION else ""

            valid_out = (df_out["Barcode"] != "") & (df_out["Store"] != "") & df_out["Outward_Date"].notna()
            df_out_valid = df_out[valid_out].copy()
            if (~valid_out).sum() > 0:
                warnings.append(f"Outward: {(~valid_out).sum():,} rows skipped (missing Barcode/Store/Date).")

            out_agg = df_out_valid.groupby(["Barcode", "Store"], as_index=False).agg(
                Article_Out=("Article", "first"),
                Color_Out=("Color", "first"),
                Size_Out=("Size", "first"),
                Outward_Qty=("Qty", "sum"),
                First_Outward_Date=("Outward_Date", "min"),
                Last_Outward_Date=("Outward_Date", "max"),
            )
            unique_out_combos = df_out_valid[["Barcode", "Store"]].drop_duplicates().shape[0]
            if len(out_agg) > unique_out_combos:
                warnings.append("Outward aggregation produced MORE rows than unique Barcode+Store combos - check Store/Barcode mapping.")

            # -------- Sales: clean once, then drop sales that happened
            # BEFORE this Barcode+Store's first tracked Outward Date
            # (those belong to older, untracked stock outside this file's window) --------
            df_sal = sal_df_raw.copy()
            df_sal["Store"] = clean_text_series(df_sal[m_sal_store])
            df_sal["Barcode"] = clean_barcode_series(df_sal[m_sal_barcode])
            df_sal["Qty"] = pd.to_numeric(df_sal[m_sal_qty], errors="coerce").fillna(0)
            df_sal["Sale_Date"] = parse_smart_dates(df_sal[m_sal_date])
            df_sal["Article"] = clean_text_series(df_sal[m_sal_article]) if m_sal_article != NONE_OPTION else ""
            df_sal["Color"] = clean_text_series(df_sal[m_sal_color]) if m_sal_color != NONE_OPTION else ""
            df_sal["Size"] = clean_text_series(df_sal[m_sal_size]) if m_sal_size != NONE_OPTION else ""

            valid_sal = (df_sal["Barcode"] != "") & (df_sal["Store"] != "")
            df_sal_valid = df_sal[valid_sal].copy()
            if (~valid_sal).sum() > 0:
                warnings.append(f"Sales: {(~valid_sal).sum():,} rows skipped (missing Barcode/Store).")

            # attach each sale's Barcode+Store first outward date, then filter
            df_sal_valid = df_sal_valid.merge(
                out_agg[["Barcode", "Store", "First_Outward_Date"]], on=["Barcode", "Store"], how="left"
            )
            before_outward_mask = df_sal_valid["Sale_Date"] < df_sal_valid["First_Outward_Date"]
            no_outward_at_all_mask = df_sal_valid["First_Outward_Date"].isna()
            drop_sales_mask = before_outward_mask | no_outward_at_all_mask
            dropped_sales_count = int(drop_sales_mask.sum())
            if dropped_sales_count > 0:
                warnings.append(
                    f"Sales: {dropped_sales_count:,} rows ignored - sale happened before the Barcode+Store's "
                    f"first tracked Outward Date (or no Outward record exists at all for that Barcode+Store)."
                )
            df_sal_clean = df_sal_valid[~drop_sales_mask].copy()

            sales_agg = df_sal_clean.groupby(["Barcode", "Store"], as_index=False).agg(
                Article_Sal=("Article", "first"),
                Color_Sal=("Color", "first"),
                Size_Sal=("Size", "first"),
                Sales_Qty=("Qty", "sum"),
                First_Sale_Date=("Sale_Date", "min"),
                Last_Sale_Date=("Sale_Date", "max"),
            )
            unique_sal_combos = df_sal_clean[["Barcode", "Store"]].drop_duplicates().shape[0]
            if len(sales_agg) > unique_sal_combos:
                warnings.append("Sales aggregation produced MORE rows than unique Barcode+Store combos - check Store/Barcode mapping.")

            # -------- Current Stock: clean + aggregate once (source of truth) --------
            df_stk = stk_df_raw.copy()
            df_stk["Store"] = clean_text_series(df_stk[m_stk_store])
            df_stk["Barcode"] = clean_barcode_series(df_stk[m_stk_barcode])
            df_stk["Qty"] = pd.to_numeric(df_stk[m_stk_qty], errors="coerce").fillna(0)
            df_stk["Article"] = clean_text_series(df_stk[m_stk_article]) if m_stk_article != NONE_OPTION else ""

            valid_stk = (df_stk["Barcode"] != "") & (df_stk["Store"] != "")
            df_stk_valid = df_stk[valid_stk].copy()
            if (~valid_stk).sum() > 0:
                warnings.append(f"Stock: {(~valid_stk).sum():,} rows skipped (missing Barcode/Store).")

            stock_agg = df_stk_valid.groupby(["Barcode", "Store"], as_index=False).agg(
                Article_Stk=("Article", "first"),
                Current_Stock=("Qty", "sum"),
            )
            unique_stk_combos = df_stk_valid[["Barcode", "Store"]].drop_duplicates().shape[0]
            if len(stock_agg) > unique_stk_combos:
                warnings.append("Stock aggregation produced MORE rows than unique Barcode+Store combos - check Store/Barcode mapping.")

            # -------- WH Stock: clean + aggregate once (optional, Barcode only) --------
            wh_available = wh_df_raw is not None and m_wh_barcode is not None and m_wh_qty is not None
            if wh_available:
                df_wh = wh_df_raw.copy()
                df_wh["Barcode"] = clean_barcode_series(df_wh[m_wh_barcode])
                df_wh["Qty"] = pd.to_numeric(df_wh[m_wh_qty], errors="coerce").fillna(0)
                df_wh_valid = df_wh[df_wh["Barcode"] != ""].copy()
                wh_agg = df_wh_valid.groupby("Barcode", as_index=False).agg(WH_Qty=("Qty", "sum"))
            else:
                wh_agg = pd.DataFrame(columns=["Barcode", "WH_Qty"])

            # ======================================================
            # SHEET 1: PERFORMANCE  (only Barcode+Store present in
            # BOTH Outward and Sales - inner join, single merge)
            # ======================================================
            perf = out_agg.merge(sales_agg, on=["Barcode", "Store"], how="inner")
            perf["Article"] = resolve_text(perf["Article_Out"], perf["Article_Sal"])
            perf["Color"] = resolve_text(perf["Color_Out"], perf["Color_Sal"])
            perf["Size"] = resolve_text(perf["Size_Out"], perf["Size_Sal"])

            perf["Selling Days"] = (perf["First_Sale_Date"] - perf["First_Outward_Date"]).dt.days.clip(lower=0)
            perf["Performance Tag"] = perf["Selling Days"].apply(performance_tag)

            performance_sheet = perf.rename(columns={
                "First_Outward_Date": "Outward Date",
                "First_Sale_Date": "Sale Date",
                "Outward_Qty": "Outward Qty",
                "Sales_Qty": "Sold Qty",
            })[[
                "Barcode", "Store", "Article", "Color", "Size",
                "Outward Date", "Sale Date", "Outward Qty", "Sold Qty",
                "Selling Days", "Performance Tag",
            ]].sort_values(["Store", "Barcode"]).reset_index(drop=True)

            # ======================================================
            # SHEET 2: REFILLING  (base = everything currently in
            # Stock file OR ever Outward'd - one outer merge each)
            # ======================================================
            refill = out_agg.merge(sales_agg, on=["Barcode", "Store"], how="outer")
            refill = refill.merge(stock_agg, on=["Barcode", "Store"], how="outer")

            for c in ["Outward_Qty", "Sales_Qty", "Current_Stock"]:
                refill[c] = refill[c].fillna(0)

            # keep only rows where something actually exists (drop all-zero noise rows)
            refill = refill[(refill["Current_Stock"] > 0) | (refill["Outward_Qty"] > 0)].copy()
            has_outward = refill["Last_Outward_Date"].notna()

            refill["Article"] = resolve_text(
                refill.get("Article_Out"), refill.get("Article_Sal"), refill.get("Article_Stk")
            )
            refill["Color"] = resolve_text(refill.get("Color_Out"), refill.get("Color_Sal"))
            refill["Size"] = resolve_text(refill.get("Size_Out"), refill.get("Size_Sal"))

            # Age Days: real value where Outward exists, RTV Cutoff value otherwise
            refill["Age Days"] = np.where(
                has_outward,
                (report_date_dt - refill["Last_Outward_Date"]).dt.days.clip(lower=0),
                rtv_cutoff,
            )

            # Selling Days (how fast it sold after Outward) - shown as its own column,
            # and used to decide Refill vs Sold-No-Refill in the sold-out case.
            refill["Selling Days"] = (refill["First_Sale_Date"] - refill["First_Outward_Date"]).dt.days
            refill["Selling Days"] = refill["Selling Days"].clip(lower=0)

            refill["Sold %"] = np.where(
                refill["Outward_Qty"] > 0,
                (refill["Sales_Qty"] / refill["Outward_Qty"] * 100).clip(upper=100).round(1),
                np.nan,
            )
            refill["Remaining Stock %"] = np.where(
                refill["Outward_Qty"] > 0,
                (refill["Current_Stock"] / refill["Outward_Qty"] * 100).clip(upper=100).round(1),
                np.nan,
            )

            age = refill["Age Days"]
            outward_qty = refill["Outward_Qty"]
            sold_qty = refill["Sales_Qty"]
            cur_stock = refill["Current_Stock"]
            sell_days = refill["Selling Days"]

            is_aged = age > rtv_cutoff
            fully_sold_but_stock_left = is_aged & (sold_qty >= outward_qty) & (cur_stock > 0)
            fast_sold_out = is_aged & (sold_qty > 0) & (cur_stock == 0) & (sell_days < rtv_cutoff)

            conditions = [
                ~has_outward,                                                    # no outward record at all
                has_outward & (age <= fresh_cutoff),                             # Fresh
                has_outward & (age > fresh_cutoff) & (age <= rtv_cutoff),        # Monitor
                is_aged & (sold_qty == 0),                                       # never sold, old
                fully_sold_but_stock_left,                                       # untracked leftover stock
                is_aged & (sold_qty > 0) & (cur_stock > 0) & (sold_qty < outward_qty),  # sold, stock still there
                fast_sold_out,                                                   # sold out fast -> WH check
                is_aged & (sold_qty > 0) & (cur_stock == 0) & (sell_days >= rtv_cutoff),  # sold out slow
            ]
            choices = ["RTV/IST", "Fresh", "Monitor", "RTV/IST", "RTV/IST", "RTV/IST", "Refill", "Sold – No Refill"]
            refill["Action"] = np.select(conditions, choices, default="Monitor")

            # -------- WH check: only applied to rows currently marked "Refill" (fast sold-out) --------
            if wh_available:
                refill = refill.merge(wh_agg, on="Barcode", how="left")
                refill["WH_Qty"] = refill["WH_Qty"].fillna(0)
                refill_candidate = refill["Action"] == "Refill"
                not_in_wh = refill_candidate & (refill["WH_Qty"] <= 0)
                refill.loc[not_in_wh, "Action"] = "Not in WH"
                refill["WH Qty"] = np.where(refill_candidate, refill["WH_Qty"], np.nan)
            else:
                refill["WH Qty"] = np.nan

            refilling_sheet = refill.rename(columns={
                "Last_Outward_Date": "Last Outward Date",
                "Outward_Qty": "Outward Qty",
                "Sales_Qty": "Sales Qty",
                "Current_Stock": "Current Stock",
            })[[
                "Barcode", "Store", "Article", "Color", "Size",
                "Last Outward Date", "Outward Qty", "Sales Qty", "Current Stock",
                "Age Days", "Selling Days", "Sold %", "Remaining Stock %", "WH Qty", "Action",
            ]].sort_values(["Store", "Barcode"]).reset_index(drop=True)

            # ======================================================
            # SHEET 3: RTV_IST
            # Outward Qty > 0 AND Sales Qty = 0 AND Current Stock = 0
            # ======================================================
            rtv_base = out_agg.merge(sales_agg, on=["Barcode", "Store"], how="left")
            rtv_base = rtv_base.merge(stock_agg, on=["Barcode", "Store"], how="left")
            rtv_base["Sales_Qty"] = rtv_base["Sales_Qty"].fillna(0)
            rtv_base["Current_Stock"] = rtv_base["Current_Stock"].fillna(0)

            rtv_mask = (rtv_base["Outward_Qty"] > 0) & (rtv_base["Sales_Qty"] == 0) & (rtv_base["Current_Stock"] == 0)
            rtv = rtv_base[rtv_mask].copy()
            rtv["Article"] = resolve_text(rtv["Article_Out"])
            rtv["Color"] = resolve_text(rtv["Color_Out"])
            rtv["Size"] = resolve_text(rtv["Size_Out"])
            rtv["Status"] = "RTV / IST"

            rtv_ist_sheet = rtv.rename(columns={
                "First_Outward_Date": "Outward Date",
                "Outward_Qty": "Outward Qty",
                "Sales_Qty": "Sales Qty",
                "Current_Stock": "Current Stock",
            })[[
                "Barcode", "Store", "Article", "Color", "Size",
                "Outward Date", "Outward Qty", "Sales Qty", "Current Stock", "Status",
            ]].sort_values(["Store", "Barcode"]).reset_index(drop=True)

            refilling_sheet = refilling_sheet[
                ~refilling_sheet.set_index(["Barcode", "Store"]).index.isin(
                    rtv_ist_sheet.set_index(["Barcode", "Store"]).index
                )
            ].reset_index(drop=True)

            # ==========================================================
            # KPI totals (computed once, cached - not recomputed on rerun)
            # ==========================================================
            kpi_totals = {
                "total_outward": int(out_agg["Outward_Qty"].sum()),
                "total_sales": int(sales_agg["Sales_Qty"].sum()),
                "total_stock": int(stock_agg["Current_Stock"].sum()),
            }

            # ==========================================================
            # EXCEL EXPORT - ONLY 3 sheets, dates formatted dd-mm-yyyy
            # Built once here and cached as bytes (normal xlsxwriter mode -
            # NOT constant_memory, which corrupts already-flushed rows when
            # a column format like the date format is applied afterwards).
            # ==========================================================
            date_cols_by_sheet = {
                "Performance": ["Outward Date", "Sale Date"],
                "Refilling": ["Last Outward Date"],
                "RTV_IST": ["Outward Date"],
            }

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="xlsxwriter", datetime_format="dd-mm-yyyy", date_format="dd-mm-yyyy") as writer:
                sheets = {
                    "Performance": performance_sheet,
                    "Refilling": refilling_sheet,
                    "RTV_IST": rtv_ist_sheet,
                }
                workbook = writer.book
                date_fmt = workbook.add_format({"num_format": "dd-mm-yyyy"})

                for sheet_name, df in sheets.items():
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                    worksheet = writer.sheets[sheet_name]
                    for col_name in date_cols_by_sheet.get(sheet_name, []):
                        if col_name in df.columns:
                            col_idx = df.columns.get_loc(col_name)
                            worksheet.set_column(col_idx, col_idx, 14, date_fmt)
                    # sensible default width for the rest of the columns
                    for i, col_name in enumerate(df.columns):
                        if col_name not in date_cols_by_sheet.get(sheet_name, []):
                            width = max(12, min(28, int(df[col_name].astype(str).str.len().max() or 12) + 2))
                            worksheet.set_column(i, i, width)

            excel_bytes = output.getvalue()

        # -------- cache everything needed to render the page, so a later
        # rerun (e.g. clicking Download) does not repeat any of the work above --------
        st.session_state.results = {
            "performance_sheet": performance_sheet,
            "refilling_sheet": refilling_sheet,
            "rtv_ist_sheet": rtv_ist_sheet,
            "warnings": warnings,
            "kpi_totals": kpi_totals,
            "report_date_dt": report_date_dt,
            "excel_bytes": excel_bytes,
            "date_cols_by_sheet": date_cols_by_sheet,
        }

    # ==========================================================
    # RENDER - reads from the cached session_state.results, so this
    # section runs cheaply on every rerun (including the download click)
    # without repeating any cleaning/aggregation/Excel-building.
    # ==========================================================
    if st.session_state.results is not None:
        results = st.session_state.results
        performance_sheet = results["performance_sheet"]
        refilling_sheet = results["refilling_sheet"]
        rtv_ist_sheet = results["rtv_ist_sheet"]
        warnings = results["warnings"]
        kpi_totals = results["kpi_totals"]
        cached_report_date_dt = results["report_date_dt"]
        excel_bytes = results["excel_bytes"]
        date_cols_by_sheet = results["date_cols_by_sheet"]

        if warnings:
            with st.expander("⚠️ Data Validation Warnings", expanded=True):
                for w in warnings:
                    st.warning(w)

        st.markdown("---")
        st.subheader("🔍 Dashboard")

        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Total Outward Qty", f"{kpi_totals['total_outward']:,}")
        k2.metric("Total Sales Qty", f"{kpi_totals['total_sales']:,}")
        k3.metric("Total Current Stock", f"{kpi_totals['total_stock']:,}")
        k4.metric("Report Date", cached_report_date_dt.strftime("%d-%m-%Y"))

        k5, k6, k7, k8 = st.columns(4)
        k5.metric("Performance Rows", f"{len(performance_sheet):,}")
        k6.metric("Refill Recommended", f"{(refilling_sheet['Action'] == 'Refill').sum():,}")
        k7.metric("RTV/IST Lines", f"{len(rtv_ist_sheet):,}")
        k8.metric("Refilling Rows (total)", f"{len(refilling_sheet):,}")

        tab_a, tab_b, tab_c = st.tabs(["🚀 Performance", "🔁 Refilling", "↩️ RTV_IST"])

        with tab_a:
            st.subheader("Performance - Sold quantity, Outward Date → Sale Date")
            if len(performance_sheet) > MAX_PREVIEW_ROWS:
                st.caption(f"Showing first {MAX_PREVIEW_ROWS:,} of {len(performance_sheet):,} rows. Full data in Excel export.")
            display_perf = format_dates_for_display(
                performance_sheet.head(MAX_PREVIEW_ROWS), date_cols_by_sheet["Performance"]
            )
            st.dataframe(display_perf, use_container_width=True)

        with tab_b:
            st.subheader("Refilling - Current Stock lifecycle & Action")
            if len(refilling_sheet) > MAX_PREVIEW_ROWS:
                st.caption(f"Showing first {MAX_PREVIEW_ROWS:,} of {len(refilling_sheet):,} rows. Full data in Excel export.")
            display_refill = format_dates_for_display(
                refilling_sheet.head(MAX_PREVIEW_ROWS), date_cols_by_sheet["Refilling"]
            )
            st.dataframe(display_refill, use_container_width=True)

        with tab_c:
            st.subheader("RTV_IST - Outward exists but not found in Sales or Stock")
            if len(rtv_ist_sheet) > MAX_PREVIEW_ROWS:
                st.caption(f"Showing first {MAX_PREVIEW_ROWS:,} of {len(rtv_ist_sheet):,} rows. Full data in Excel export.")
            display_rtv = format_dates_for_display(
                rtv_ist_sheet.head(MAX_PREVIEW_ROWS), date_cols_by_sheet["RTV_IST"]
            )
            st.dataframe(display_rtv, use_container_width=True)

        st.markdown("---")
        st.subheader("📥 Export Report")
        st.download_button(
            label="📥 Download Report (Performance, Refilling, RTV_IST)",
            data=excel_bytes,
            file_name=f"Stock_Aging_Refill_Report_{datetime.today().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )

else:
    st.info("💡 Please upload Outward, Sales and Current Stock files to begin. WH Stock file is optional.")
