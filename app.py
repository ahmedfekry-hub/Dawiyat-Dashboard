
from __future__ import annotations

from io import BytesIO
from pathlib import Path
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(
    page_title="Dawiyat Project Intelligence Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

DEFAULT_FILE = "Dawiyat Master Sheet.xlsx"


# ---------------------- helpers ----------------------
def clean_text(value):
    if pd.isna(value):
        return np.nan
    text = str(value).strip()
    return np.nan if text in {"", "nan", "None", "null"} else text


def first_existing(df: pd.DataFrame, names: list[str]) -> str | None:
    norm = {str(c).strip().lower(): c for c in df.columns}
    for name in names:
        if name.strip().lower() in norm:
            return norm[name.strip().lower()]
    for c in df.columns:
        low = str(c).strip().lower()
        if any(name.strip().lower() in low for name in names):
            return c
    return None


def choose_mode(series: pd.Series):
    series = series.dropna()
    if series.empty:
        return np.nan
    mode = series.mode(dropna=True)
    return mode.iloc[0] if not mode.empty else series.iloc[0]


def ratio_pct(progress, scope):
    p = pd.to_numeric(progress, errors="coerce")
    s = pd.to_numeric(scope, errors="coerce")
    with np.errstate(divide="ignore", invalid="ignore"):
        out = np.where(s > 0, (p / s) * 100, np.nan)
    return pd.Series(out, index=p.index).clip(lower=0, upper=100)


def safe_numeric(series, default=0):
    if isinstance(series, (int, float)):
        return pd.Series([series])
    if series is None:
        return pd.Series(dtype="float64")
    return pd.to_numeric(series, errors="coerce").fillna(default)


def safe_series(df: pd.DataFrame, col: str, default=np.nan):
    if col in df.columns:
        return df[col]
    return pd.Series([default] * len(df), index=df.index)


def coalesce_into(df: pd.DataFrame, target: str, *fallbacks: str):
    if target not in df.columns:
        df[target] = np.nan
    result = df[target].copy()
    for fb in fallbacks:
        if fb in df.columns:
            result = result.combine_first(df[fb])
    df[target] = result
    return df[target]


def fmt_money(x):
    if pd.isna(x):
        return "-"
    x = float(x)
    if abs(x) >= 1_000_000:
        return f"SAR {x/1_000_000:,.2f}M"
    if abs(x) >= 1_000:
        return f"SAR {x/1_000:,.1f}K"
    return f"SAR {x:,.0f}"


def fmt_pct(x):
    return "-" if pd.isna(x) else f"{x:,.1f}%"


def add_card(title: str, value: str, subtitle: str = ""):
    st.markdown(
        f"""
        <div class="metric-card">
            <div class="metric-title">{title}</div>
            <div class="metric-value">{value}</div>
            <div class="metric-subtitle">{subtitle}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def to_excel_bytes(df: pd.DataFrame, sheet_name: str = "Export") -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name[:31])
    output.seek(0)
    return output.getvalue()


def to_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode("utf-8-sig")


def normalize_region(old_region, city):
    text = str(old_region).strip().upper() if pd.notna(old_region) else ""
    ctext = str(city).strip().upper() if pd.notna(city) else ""

    west_tokens = {"MAKKAH", "MECCA", "BAHA", "AL BAHA", "MADINAH", "MEDINA", "TABUK", "WESTERN"}
    south_tokens = {"JIZAN", "NAJRAN", "ABHA", "KHAMIS", "BISHA", "SOUTHERN"}
    east_tokens = {"EASTERN", "EASTERN REGION", "DAMMAM", "KHOBAR", "JUBAIL", "HASA"}
    north_tokens = {"NORTHERN", "AL JOUF", "ARAR", "HAIL", "QURAYAT", "TURAIF"}

    west_cities = {"JEDDAH", "TAIF", "AL BAHA", "BAHA", "TABUK", "MADINAH", "MEDINA", "MAKKAH"}
    south_cities = {"JIZAN", "ABHA", "NAJRAN", "KHAMIS MUSHAIT", "BISHA"}
    east_cities = {"DAMMAM", "KHOBAR", "JUBAIL", "AL AHSA", "HOFUF"}
    north_cities = {"ARAR", "SAKAKA", "JOUF", "AL JOUF", "HAIL", "QURAYAT", "TURAIF"}

    if text in west_tokens or ctext in west_cities:
        return "Western"
    if text in south_tokens or ctext in south_cities:
        return "Southern"
    if text in east_tokens or ctext in east_cities:
        return "Eastern"
    if text in north_tokens or ctext in north_cities:
        return "Northern"
    return "Not Classified"


def apply_theme_css(theme_mode: str):
    dark = theme_mode == "Dark"
    bg = "#07111f" if dark else "#F3F7FC"
    panel = "#091a31" if dark else "#FFFFFF"
    card = "#0f1f38" if dark else "#FFFFFF"
    text = "#ECF3FF" if dark else "#16253B"
    muted = "#94A9C7" if dark else "#667085"
    border = "rgba(148,163,184,0.16)" if dark else "rgba(15,23,42,0.08)"
    glow = "0 10px 30px rgba(59,130,246,0.18)" if dark else "0 10px 26px rgba(15,23,42,0.06)"
    st.markdown(
        f"""
        <style>
        .stApp {{
            background:
                radial-gradient(circle at top left, rgba(59,130,246,0.18), transparent 28%),
                radial-gradient(circle at top right, rgba(245,158,11,0.10), transparent 22%),
                {bg};
            color: {text};
        }}
        .block-container {{padding-top: 0.9rem; padding-bottom: 2rem; max-width: 95rem;}}
        section[data-testid="stSidebar"] {{
            background: linear-gradient(180deg, #11264A 0%, #081426 100%);
            border-right: 1px solid rgba(148,163,184,0.16);
        }}
        section[data-testid="stSidebar"] * {{ color: #ECF3FF !important; }}
        .top-banner {{
            background: linear-gradient(135deg, rgba(59,130,246,0.23), rgba(16,185,129,0.08), rgba(245,158,11,0.10));
            border: 1px solid {border};
            border-radius: 24px;
            padding: 22px 24px;
            margin-bottom: 14px;
            box-shadow: {glow};
        }}
        .metric-card {{
            background: linear-gradient(180deg, rgba(15,31,56,0.95), rgba(11,23,42,0.95));
            border: 1px solid {border};
            border-radius: 18px;
            padding: 16px 18px 14px 18px;
            min-height: 128px;
            box-shadow: {glow};
        }}
        .metric-title {{ color: #9BB1D0; font-size: 0.92rem; margin-bottom: 6px; }}
        .metric-value {{ color: #F8FBFF; font-size: 1.95rem; font-weight: 700; line-height: 1.15; margin-bottom: 8px; }}
        .metric-subtitle {{ color: #8FA8CA; font-size: 0.88rem; }}
        .section-card {{
            background: {card};
            border: 1px solid {border};
            border-radius: 22px;
            padding: 10px 14px 2px 14px;
            box-shadow: {glow};
        }}
        .guide-box {{
            background: {panel};
            border-left: 4px solid #3B82F6;
            border-radius: 14px;
            border: 1px solid {border};
            padding: 14px 16px;
            margin: 10px 0px;
        }}
        .guide-title {{font-weight:700; margin-bottom:6px;}}
        .subtle {{ color: {muted}; font-size: 0.95rem; }}
        .small-chip {{
            display:inline-block; padding: 6px 10px; border-radius: 999px;
            background: rgba(59,130,246,0.12); color: {text}; font-size: 0.85rem;
            border: 1px solid {border}; margin-right: 8px; margin-top: 4px;
        }}
        div[data-testid="stMetric"] {{
            background: {card};
            border: 1px solid {border};
            padding: 12px 14px;
            border-radius: 16px;
            box-shadow: {glow};
        }}
        .stDataFrame, div[data-testid="stTable"] {{
            border-radius: 16px; overflow: hidden; border: 1px solid {border};
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )


@st.cache_data(show_spinner=False)
def load_data(file_source=None):
    if file_source is None:
        file_source = Path(__file__).with_name(DEFAULT_FILE)
    xls = pd.ExcelFile(file_source)

    def find_sheet(candidates):
        normalized = {str(name).strip().lower(): name for name in xls.sheet_names}
        for c in candidates:
            if c.strip().lower() in normalized:
                return normalized[c.strip().lower()]
        for name in xls.sheet_names:
            low = str(name).strip().lower()
            if any(c.strip().lower() in low for c in candidates):
                return name
        return None

    main_name = find_sheet(["Dawaiyat Service Tool", "service tool"])
    district_name = find_sheet(["District"])
    penalties_name = find_sheet(["Penalties", "Penalty"])
    detail_name = find_sheet(["Workorder details", "Workorder detail"])

    if main_name is None:
        raise ValueError("Main sheet 'Dawaiyat Service Tool' was not found.")

    main = pd.read_excel(file_source, sheet_name=main_name)
    district = pd.read_excel(file_source, sheet_name=district_name) if district_name else pd.DataFrame()
    penalties = pd.read_excel(file_source, sheet_name=penalties_name) if penalties_name else pd.DataFrame()
    details = pd.read_excel(file_source, sheet_name=detail_name) if detail_name else pd.DataFrame()

    for df in [main, district, penalties, details]:
        if not df.empty:
            df.columns = [str(c).strip() for c in df.columns]
            for col in df.columns:
                if df[col].dtype == "object":
                    df[col] = df[col].map(clean_text)

    link_col = first_existing(main, ["Link Code"])
    workorder_col = first_existing(main, ["Work Order"])
    if link_col is None:
        raise ValueError("Column 'Link Code' not found in main sheet.")

    main = main[main[link_col].notna()].copy()
    main[link_col] = main[link_col].astype(str).str.strip()
    main.rename(columns={link_col: "Link Code"}, inplace=True)
    if workorder_col:
        main.rename(columns={workorder_col: "Work Order"}, inplace=True)

    # District mapping sheet
    data = main.copy()
    if not district.empty:
        d_link = first_existing(district, ["Link Code"])
        d_city = first_existing(district, ["City"])
        d_dist = first_existing(district, ["District"])
        d_region = first_existing(district, ["Region"])
        if d_link:
            district = district[district[d_link].notna()].copy()
            district[d_link] = district[d_link].astype(str).str.strip()
            agg = {}
            if d_city:
                agg["City"] = (d_city, choose_mode)
            if d_dist:
                agg["District"] = (d_dist, choose_mode)
            if d_region:
                agg["Region_sheet"] = (d_region, choose_mode)
            district_map = district.groupby(d_link, dropna=False).agg(**agg).reset_index().rename(columns={d_link: "Link Code"})
            data = data.merge(district_map, on="Link Code", how="left")

            district["p1"] = district[d_link].astype(str).str.split("-").str[0].str.upper()
            district["p3"] = district[d_link].astype(str).str.split("-").str[2].str.upper()
            city_prefix_map = district.groupby("p1")[d_city].agg(choose_mode).dropna().to_dict() if d_city else {}
            dist_prefix_map = district.groupby("p3")[d_dist].agg(choose_mode).dropna().to_dict() if d_dist else {}
        else:
            city_prefix_map, dist_prefix_map = {}, {}
    else:
        city_prefix_map, dist_prefix_map = {}, {}

    data["p1"] = data["Link Code"].astype(str).str.split("-").str[0].str.upper()
    data["p3"] = data["Link Code"].astype(str).str.split("-").str[2].str.upper()

    if "City" not in data.columns:
        data["City"] = np.nan
    if "District" not in data.columns:
        data["District"] = np.nan

    data["City"] = data["City"].fillna(data["p1"].map(city_prefix_map))
    data["District"] = data["District"].fillna(data["p3"].map(dist_prefix_map))

    data["City"] = data["City"].fillna(
        data["p1"].replace({
            "JED": "Jeddah",
            "TAI": "Taif",
            "TAB": "Tabuk",
            "BAH": "Al Baha",
            "SAM": "Jizan",
            "SHA": "Jizan",
            "SHU": "Jizan",
            "EDA": "Jizan",
            "DAI": "Jizan",
            "MOY": "Taif",
            "TUR": "Taif",
        })
    )
    data["District"] = data["District"].fillna(data["p3"].str.title())

    # Region logic: use district mapping region if available, else normalize to four macro-regions
    data["Region_raw_main"] = safe_series(data, "Region")
    data["Region"] = safe_series(data, "Region_sheet")
    data["Region"] = data["Region"].fillna(
        pd.Series(
            [normalize_region(r, c) for r, c in zip(data["Region_raw_main"], data["City"])],
            index=data.index
        )
    )

    date_candidates = ["Created", "Assigned at", "In Progress at", "Updated", "Closed at", "Targeted Completion", "Updated Target Date"]
    for col in date_candidates:
        if col in data.columns:
            data[col] = pd.to_datetime(data[col], errors="coerce")

    num_cols = [
        "Percentage of Completion","WO Cost","Cost","Updates","Trench Progress","Trench Scope",
        "MH/HH Progress","MH/HH Scope","Fiber Progress","Fiber Scope","ODBs Progress","ODBs Scope",
        "ODFs Progress","ODFs Scope","JCL Progress","JCL Scope","FAT Progress","FAT Scope",
        "PFAT Progress","PFAT Scope","SFAT Progress","SFAT Scope","Permits Progress","Permits Scope",
        "PIP rejection count","PAT rejection count","Approval rejection count",
        "As-Built Rejection Count","Handover Rejection Count"
    ]
    for col in num_cols:
        if col in data.columns:
            data[col] = pd.to_numeric(data[col], errors="coerce")

    snapshot_date = safe_series(data, "Updated").dropna().max()
    if pd.isna(snapshot_date):
        snapshot_date = pd.Timestamp.today().normalize()

    targeted = safe_series(data, "Targeted Completion")
    updated_target = safe_series(data, "Updated Target Date")
    data["effective_target"] = updated_target.combine_first(targeted)

    start_date = safe_series(data, "In Progress at").combine_first(safe_series(data, "Assigned at")).combine_first(safe_series(data, "Created"))
    data["start_date"] = pd.to_datetime(start_date, errors="coerce")

    elapsed = (snapshot_date - data["start_date"]).dt.days
    total = (data["effective_target"] - data["start_date"]).dt.days
    with np.errstate(divide="ignore", invalid="ignore"):
        data["planned_progress_pct"] = np.where(total > 0, np.clip((elapsed / total) * 100, 0, 100), np.nan)

    data["actual_progress_pct"] = pd.to_numeric(safe_series(data, "Percentage of Completion"), errors="coerce")
    data["actual_progress_capped"] = data["actual_progress_pct"].clip(lower=0, upper=100)
    data["lag_pp"] = data["planned_progress_pct"] - data["actual_progress_capped"]
    data["is_complete"] = data["actual_progress_capped"] >= 100
    data["is_overdue"] = (snapshot_date > data["effective_target"]) & (~data["is_complete"]) & data["effective_target"].notna()
    data["critical_lag"] = data["lag_pp"] >= 15

    updated_series = pd.to_datetime(safe_series(data, "Updated"), errors="coerce")
    data["days_since_update"] = (snapshot_date - updated_series).dt.days
    data["Updates"] = pd.to_numeric(safe_series(data, "Updates"), errors="coerce")
    data["needs_system_update"] = (data["Updates"] < 5) & (data["days_since_update"] > 5)

    elapsed_days = np.maximum((snapshot_date - data["start_date"]).dt.days, 1)
    actual_ratio = data["actual_progress_capped"] / 100.0
    est_total_duration = np.where(actual_ratio > 0, elapsed_days / actual_ratio, np.nan)
    data["forecast_completion_date"] = data["start_date"] + pd.to_timedelta(est_total_duration, unit="D")
    data["forecast_delay_days"] = (data["forecast_completion_date"] - data["effective_target"]).dt.days
    data["forecast_risk"] = np.select(
        [data["forecast_delay_days"] > 30, data["forecast_delay_days"] > 0, data["forecast_delay_days"] <= 0],
        ["High delay risk", "Moderate delay risk", "On forecast"],
        default="Insufficient data",
    )

    data["civil_completion_pct"] = ratio_pct(safe_series(data, "Trench Progress"), safe_series(data, "Trench Scope"))
    data["mhhh_completion_pct"] = ratio_pct(safe_series(data, "MH/HH Progress"), safe_series(data, "MH/HH Scope"))
    data["fiber_completion_pct"] = ratio_pct(safe_series(data, "Fiber Progress"), safe_series(data, "Fiber Scope"))
    data["permits_completion_pct"] = ratio_pct(safe_series(data, "Permits Progress"), safe_series(data, "Permits Scope"))

    # normalize dimensions
    for col in ["Year", "Work Order Status", "Type", "Class", "Project", "Subclass", "Stage", "Region", "City", "District"]:
        if col not in data.columns:
            data[col] = np.nan

    year_series = safe_series(data, "Year")
    target_year = pd.to_datetime(data["effective_target"], errors="coerce").dt.year.astype("Int64")
    data["Year"] = pd.to_numeric(year_series, errors="coerce").astype("Int64").fillna(target_year)

    for col in ["Work Order Status", "Type", "Class", "Project", "Subclass", "Stage", "Region", "City", "District"]:
        data[col] = data[col].fillna("Not Classified").astype(str).str.strip()

    # penalties
    if penalties.empty:
        penalties = pd.DataFrame(columns=["Link Code", "Deviation name", "Number of Deviations", "Penalties Amount", "Region", "City", "District"])
    else:
        p_link = first_existing(penalties, ["Cluster Name", "Link Code"])
        p_name = first_existing(penalties, ["Deviation name", "Penalty"])
        p_qty = first_existing(penalties, ["Number of Deviations", "Number", "Qty"])
        p_amt = first_existing(penalties, ["Penalties Amount", "Amount"])
        p_region = first_existing(penalties, ["Region"])
        p_city = first_existing(penalties, ["City"])
        p_dist = first_existing(penalties, ["District"])

        penalties["Link Code"] = penalties[p_link].astype(str).str.strip() if p_link else np.nan
        penalties["Number of Deviations"] = pd.to_numeric(penalties[p_qty], errors="coerce").fillna(0) if p_qty else 0
        penalties["Penalties Amount"] = pd.to_numeric(penalties[p_amt], errors="coerce").fillna(0) if p_amt else 0
        penalties["Deviation name"] = penalties[p_name] if p_name else "Deviation"

        if p_region and p_region != "Region":
            penalties.rename(columns={p_region: "Region"}, inplace=True)
        if p_city and p_city != "City":
            penalties.rename(columns={p_city: "City"}, inplace=True)
        if p_dist and p_dist != "District":
            penalties.rename(columns={p_dist: "District"}, inplace=True)

        pen_map = data[["Link Code", "Region", "City", "District"]].drop_duplicates("Link Code")
        penalties = penalties.merge(pen_map, on="Link Code", how="left", suffixes=("", "_main"))

        coalesce_into(penalties, "Region", "Region_main")
        coalesce_into(penalties, "City", "City_main")
        coalesce_into(penalties, "District", "District_main")

        penalties["Region"] = penalties["Region"].fillna(penalties["Region_main"] if "Region_main" in penalties.columns else np.nan)
        penalties["City"] = penalties["City"].fillna("Not Classified")
        penalties["District"] = penalties["District"].fillna("Not Classified")
        penalties["Region"] = pd.Series(
            [normalize_region(r, c) for r, c in zip(penalties["Region"], penalties["City"])],
            index=penalties.index
        ).fillna("Not Classified")

    pen_agg = penalties.groupby("Link Code", dropna=False).agg(
        penalty_rows=("Deviation name", "size"),
        penalty_qty=("Number of Deviations", "sum"),
        penalty_amount=("Penalties Amount", "sum"),
    ).reset_index()

    data = data.merge(pen_agg, on="Link Code", how="left")
    for c in ["penalty_rows", "penalty_qty", "penalty_amount"]:
        data[c] = pd.to_numeric(data[c], errors="coerce").fillna(0)

    if details.empty:
        details = data.copy()
    else:
        d_link = first_existing(details, ["Link Code"])
        if d_link and d_link != "Link Code":
            details.rename(columns={d_link: "Link Code"}, inplace=True)
        details["Link Code"] = details["Link Code"].astype(str).str.strip()
        enrich_cols = [c for c in [
            "Work Order", "Updates", "Updated", "effective_target", "City", "District", "Region",
            "Project", "Subclass", "Stage", "Type", "Class", "Work Order Status",
            "lag_pp", "planned_progress_pct", "actual_progress_capped", "forecast_risk",
            "penalty_qty", "penalty_amount", "days_since_update"
        ] if c in data.columns]
        details = details.merge(
            data[["Link Code"] + [c for c in enrich_cols if c != "Link Code"]].drop_duplicates("Link Code"),
            on="Link Code",
            how="left"
        )
        for col in details.columns:
            if details[col].dtype == "object":
                details[col] = details[col].map(clean_text)
        for col in ["Created", "Assigned at", "Updated", "Targeted Completion", "Updated Target Date", "Closed at", "effective_target"]:
            if col in details.columns:
                details[col] = pd.to_datetime(details[col], errors="coerce")

    warnings = []
    for c in ["Link Code", "Work Order", "Percentage of Completion", "Updated", "Updates", "Targeted Completion", "Cost", "Work Order Status"]:
        if c not in data.columns:
            warnings.append(c)

    return data, penalties, details, snapshot_date, warnings


# ---------------------- app start ----------------------
st.sidebar.title("Control Panel")
theme_mode = st.sidebar.radio("Theme", ["Dark", "Light"], horizontal=True, index=0)
apply_theme_css(theme_mode)

uploaded_file = st.sidebar.file_uploader("Upload refreshed Dawiyat workbook", type=["xlsx"])
source = uploaded_file if uploaded_file is not None else None

data, penalties, details, snapshot_date, data_warnings = load_data(source)

st.markdown(
    f"""
    <div class="top-banner">
        <h1 style="margin:0 0 6px 0;">Dawiyat Project Intelligence Dashboard</h1>
        <div class="subtle">
            Executive PMO dashboard for top management decision support.
            Snapshot date: <b>{pd.to_datetime(snapshot_date).strftime('%d %b %Y %H:%M')}</b>
        </div>
        <div style="margin-top:10px;">
            <span class="small-chip">Professional first-style executive layout</span>
            <span class="small-chip">Daily refresh from latest workbook</span>
            <span class="small-chip">Region filter normalized to Western / Southern / Eastern / Northern</span>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

if data_warnings:
    st.warning("Missing key columns detected in workbook: " + ", ".join(data_warnings))


def options_for(df, col):
    if col not in df.columns:
        return ["All"]
    vals = [v for v in pd.Series(df[col]).dropna().astype(str).unique().tolist() if v not in {"nan", "None"}]
    return ["All"] + sorted(vals)


base = data.copy()
region = st.sidebar.selectbox("Region", options_for(base, "Region"), index=0)
if region != "All":
    base = base[base["Region"].astype(str) == region]
city = st.sidebar.selectbox("City", options_for(base, "City"), index=0)
if city != "All":
    base = base[base["City"].astype(str) == city]
district = st.sidebar.selectbox("District", options_for(base, "District"), index=0)
if district != "All":
    base = base[base["District"].astype(str) == district]
project = st.sidebar.selectbox("Project", options_for(base, "Project"), index=0)
if project != "All":
    base = base[base["Project"].astype(str) == project]
stage = st.sidebar.selectbox("Stage", options_for(base, "Stage"), index=0)
if stage != "All":
    base = base[base["Stage"].astype(str) == stage]
year = st.sidebar.selectbox("Year", options_for(base.assign(Year=base["Year"].astype(str)), "Year"), index=0)
if year != "All":
    base = base[base["Year"].astype(str) == year]
status = st.sidebar.selectbox("Work Order Status", options_for(base, "Work Order Status"), index=0)
if status != "All":
    base = base[base["Work Order Status"].astype(str) == status]
wo_type = st.sidebar.selectbox("Type", options_for(base, "Type"), index=0)
if wo_type != "All":
    base = base[base["Type"].astype(str) == wo_type]
wo_class = st.sidebar.selectbox("Class", options_for(base, "Class"), index=0)
if wo_class != "All":
    base = base[base["Class"].astype(str) == wo_class]
subclass = st.sidebar.selectbox("Subclass", options_for(base, "Subclass"), index=0)
if subclass != "All":
    base = base[base["Subclass"].astype(str) == subclass]
link_code = st.sidebar.selectbox("Link Code", options_for(base, "Link Code"), index=0)

filtered = base.copy()
if link_code != "All":
    filtered = filtered[filtered["Link Code"].astype(str) == link_code]

if filtered.empty:
    st.warning("No records match the selected filters.")
    st.stop()

filtered_links = filtered["Link Code"].dropna().astype(str).unique().tolist()
filtered_details = details[details["Link Code"].astype(str).isin(filtered_links)].copy() if "Link Code" in details.columns else pd.DataFrame()
filtered_penalties = penalties[penalties["Link Code"].astype(str).isin(filtered_links)].copy() if "Link Code" in penalties.columns else pd.DataFrame()

# KPI calculations
total_wo = filtered["Work Order"].nunique() if "Work Order" in filtered.columns else len(filtered)
total_link = filtered["Link Code"].nunique()
avg_actual = filtered["actual_progress_capped"].mean()
avg_planned = filtered["planned_progress_pct"].mean()
avg_lag_pp = filtered["lag_pp"].mean()
lagged_pct = (filtered["lag_pp"] > 0).mean() * 100 if len(filtered) else np.nan
overdue_cnt = int(filtered["is_overdue"].sum())
critical_lag_cnt = int(filtered["critical_lag"].sum())
forecast_high_risk = int((filtered["forecast_risk"] == "High delay risk").sum())
penalty_rows = int(len(filtered_penalties))
penalty_qty = float(pd.to_numeric(filtered_penalties["Number of Deviations"], errors="coerce").fillna(0).sum()) if not filtered_penalties.empty else 0.0
penalty_amount = float(pd.to_numeric(filtered_penalties["Penalties Amount"], errors="coerce").fillna(0).sum()) if not filtered_penalties.empty else 0.0
avg_civil = filtered["civil_completion_pct"].mean()
avg_fiber = filtered["fiber_completion_pct"].mean()
rejection_cols = [c for c in ["PIP rejection count","PAT rejection count","Approval rejection count","As-Built Rejection Count","Handover Rejection Count"] if c in filtered.columns]
rejections_total = float(filtered[rejection_cols].fillna(0).sum().sum()) if rejection_cols else 0
update_needed = filtered[filtered["needs_system_update"]].copy()
in_progress_cnt = int(filtered["Work Order Status"].astype(str).str.contains("In Progress", case=False, na=False).sum())
cancelled_cnt = int(filtered["Work Order Status"].astype(str).str.contains("Cancel", case=False, na=False).sum())

cost_col = "Cost" if "Cost" in filtered.columns else ("WO Cost" if "WO Cost" in filtered.columns else None)
snapshot_month = pd.Timestamp(snapshot_date).to_period("M")
target_month_series = pd.to_datetime(safe_series(filtered, "Targeted Completion"), errors="coerce").dt.to_period("M")
updated_month_series = pd.to_datetime(safe_series(filtered, "Updated Target Date"), errors="coerce").dt.to_period("M")
if cost_col:
    current_month_target_cost = pd.to_numeric(filtered.loc[target_month_series == snapshot_month, cost_col], errors="coerce").fillna(0).sum()
    current_month_updated_target_cost = pd.to_numeric(filtered.loc[updated_month_series == snapshot_month, cost_col], errors="coerce").fillna(0).sum()
else:
    current_month_target_cost = 0
    current_month_updated_target_cost = 0


def chart_layout(fig, title=None, height=360):
    fig.update_layout(
        margin=dict(l=20, r=20, t=48, b=20),
        legend_title_text="",
        height=height,
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        title=title,
    )
    return fig


def monthly_progress_chart(df):
    temp = df.copy()
    temp["target_month"] = pd.to_datetime(temp["effective_target"], errors="coerce").dt.to_period("M").astype(str)
    grp = temp.groupby("target_month", dropna=False).agg(
        Planned=("planned_progress_pct", "mean"),
        Actual=("actual_progress_capped", "mean"),
    ).reset_index()
    grp = grp[grp["target_month"] != "NaT"].sort_values("target_month")
    if grp.empty:
        return None
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=grp["target_month"], y=grp["Planned"], mode="lines+markers", name="Planned"))
    fig.add_trace(go.Scatter(x=grp["target_month"], y=grp["Actual"], mode="lines+markers", name="Actual"))
    fig.update_yaxes(range=[0, 110], title="Progress %")
    fig.update_xaxes(title="Target Month")
    return chart_layout(fig, "Planned vs Actual Progress Trend")


def city_status_chart(df):
    grp = df.groupby(["City"], dropna=False).agg(
        Avg_Planned=("planned_progress_pct", "mean"),
        Avg_Actual=("actual_progress_capped", "mean"),
        Lag=("lag_pp", "mean"),
    ).reset_index().sort_values("Avg_Actual", ascending=False)
    if grp.empty:
        return None
    fig = go.Figure()
    fig.add_trace(go.Bar(x=grp["City"], y=grp["Avg_Planned"], name="Planned"))
    fig.add_trace(go.Bar(x=grp["City"], y=grp["Avg_Actual"], name="Actual"))
    fig.update_layout(barmode="group")
    fig.update_yaxes(title="Progress %")
    return chart_layout(fig, "City Performance")


def milestone_chart(df):
    rows = []
    for label, col in [("Civil", "civil_completion_pct"), ("MH/HH", "mhhh_completion_pct"), ("Fiber", "fiber_completion_pct"), ("Permits", "permits_completion_pct")]:
        if col in df.columns:
            rows.append({"Milestone": label, "Completion": df[col].mean()})
    grp = pd.DataFrame(rows)
    if grp.empty:
        return None
    fig = px.bar(grp, x="Milestone", y="Completion")
    fig.update_yaxes(range=[0, 110], title="Completion %")
    return chart_layout(fig, "Milestone Completion Summary")


def risk_chart(df):
    grp = df["forecast_risk"].value_counts().rename_axis("Risk").reset_index(name="Count")
    if grp.empty:
        return None
    fig = px.bar(grp, x="Risk", y="Count")
    return chart_layout(fig, "Forecast Risk Exposure")


def penalties_deviation_chart(df):
    if df.empty:
        return None
    grp = df.groupby("Deviation name", dropna=False).agg(
        Deviations=("Number of Deviations", "sum"),
        Deduction=("Penalties Amount", "sum"),
    ).reset_index().sort_values("Deviations", ascending=False).head(12)
    if grp.empty:
        return None
    fig = go.Figure()
    fig.add_trace(go.Bar(y=grp["Deviation name"], x=grp["Deviations"], name="Deviation Count", orientation="h"))
    fig.add_trace(go.Scatter(y=grp["Deviation name"], x=grp["Deduction"], name="Penalty Amount", mode="markers"))
    fig.update_layout(yaxis=dict(categoryorder="total ascending"))
    return chart_layout(fig, "Top Deviations vs Deduction Amount", height=420)


def district_lag_chart(df):
    grp = df.groupby("District", dropna=False).agg(Lag=("lag_pp", "mean")).reset_index().sort_values("Lag", ascending=False).head(12)
    if grp.empty:
        return None
    fig = px.bar(grp, x="Lag", y="District", orientation="h")
    fig.update_layout(yaxis=dict(categoryorder="total ascending"))
    return chart_layout(fig, "District Lag Ranking", height=420)


tabs = st.tabs(["Executive Overview", "PMO Summary", "Schedule & KPI", "Penalties & Quality", "Work Order Detail", "Dashboard Guide"])

with tabs[0]:
    row = st.columns(6)
    metrics = [
        ("Work Orders", f"{total_wo:,}", f"{total_link:,} link codes"),
        ("Avg Actual Progress", fmt_pct(avg_actual), "Overall completion"),
        ("Avg Planned Progress", fmt_pct(avg_planned), "Based on effective target dates"),
        ("Avg Lag", fmt_pct(avg_lag_pp), "Planned minus actual"),
        ("High Delay Risk", f"{forecast_high_risk:,}", "Forecast delay > 30 days"),
        ("Penalty Amount", fmt_money(penalty_amount), "Total deduction amount"),
    ]
    for col, metric in zip(row, metrics):
        with col:
            add_card(*metric)

    c1, c2 = st.columns([1.15, 0.85])
    with c1:
        st.markdown('<div class="section-card">', unsafe_allow_html=True)
        fig = monthly_progress_chart(filtered)
        if fig:
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Not enough target-date data for trend chart.")
        st.markdown('</div>', unsafe_allow_html=True)
    with c2:
        st.markdown('<div class="section-card">', unsafe_allow_html=True)
        fig = risk_chart(filtered)
        if fig:
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("No forecast risk data available.")
        st.markdown('</div>', unsafe_allow_html=True)

    c3, c4 = st.columns([1.05, 0.95])
    with c3:
        st.markdown('<div class="section-card">', unsafe_allow_html=True)
        fig = city_status_chart(filtered)
        if fig:
            st.plotly_chart(fig, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)
    with c4:
        st.markdown('<div class="section-card">', unsafe_allow_html=True)
        fig = milestone_chart(filtered)
        if fig:
            st.plotly_chart(fig, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)

with tabs[1]:
    row = st.columns(6)
    pmo_metrics = [
        ("Need System Update", f"{len(update_needed):,}", "Updates < 5 and no update for > 5 days"),
        ("Current Month Target Cost", fmt_money(current_month_target_cost), "From Targeted Completion"),
        ("Current Month Updated Target Cost", fmt_money(current_month_updated_target_cost), "From Updated Target Date"),
        ("Lagged Work Orders", fmt_pct(lagged_pct), "Share of work orders behind plan"),
        ("In Progress", f"{in_progress_cnt:,}", "From Work Order Status"),
        ("Cancelled", f"{cancelled_cnt:,}", "From Work Order Status"),
    ]
    for col, metric in zip(row, pmo_metrics):
        with col:
            add_card(*metric)

    left, right = st.columns([0.95, 1.05])
    with left:
        st.markdown('<div class="section-card">', unsafe_allow_html=True)
        fig = district_lag_chart(filtered)
        if fig:
            st.plotly_chart(fig, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)
    with right:
        st.markdown('<div class="section-card">', unsafe_allow_html=True)
        st.subheader("Link Codes Requiring System Update")
        if update_needed.empty:
            st.success("No link codes currently match the PMO update-follow-up rule.")
        else:
            cols = [c for c in ["Link Code", "Work Order", "Region", "City", "District", "Stage", "Updated", "Updates", "days_since_update", "actual_progress_capped", "planned_progress_pct", "lag_pp", "effective_target"] if c in update_needed.columns]
            pmo_table = update_needed[cols].copy().sort_values(["days_since_update", "lag_pp"], ascending=[False, False])
            st.dataframe(pmo_table, use_container_width=True, hide_index=True)
            st.download_button("Download PMO follow-up list (Excel)", data=to_excel_bytes(pmo_table, "PMO Follow Up"), file_name="pmo_follow_up_link_codes.xlsx")
        st.markdown('</div>', unsafe_allow_html=True)

with tabs[2]:
    row = st.columns(5)
    schedule_metrics = [
        ("Avg Civil Completion", fmt_pct(avg_civil), "Trench progress / scope"),
        ("Avg Fiber Completion", fmt_pct(avg_fiber), "Fiber progress / scope"),
        ("Overdue Work Orders", f"{overdue_cnt:,}", "Past effective target and not complete"),
        ("Critical Lag", f"{critical_lag_cnt:,}", "Lag >= 15%"),
        ("Total Rejections", f"{rejections_total:,.0f}", "All rejection counts combined"),
    ]
    for col, metric in zip(row, schedule_metrics):
        with col:
            add_card(*metric)

    left, right = st.columns(2)
    with left:
        st.markdown('<div class="section-card">', unsafe_allow_html=True)
        fig = monthly_progress_chart(filtered)
        if fig:
            st.plotly_chart(fig, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)
    with right:
        st.markdown('<div class="section-card">', unsafe_allow_html=True)
        fig = district_lag_chart(filtered)
        if fig:
            st.plotly_chart(fig, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)

with tabs[3]:
    row = st.columns(4)
    pen_metrics = [
        ("Deviation Rows", f"{penalty_rows:,}", "Rows / records in penalties after filter"),
        ("Number of Deviations", f"{penalty_qty:,.0f}", "Sum of deviation counts"),
        ("Penalty Deduction Amount", fmt_money(penalty_amount), "Sum of deducted amount only"),
        ("Deviation Records With No Deduction", f"{int((filtered_penalties['Penalties Amount'].fillna(0) <= 0).sum()) if not filtered_penalties.empty else 0:,}", "Not all deviations have penalties"),
    ]
    for col, metric in zip(row, pen_metrics):
        with col:
            add_card(*metric)

    st.markdown(
        """
        <div class="guide-box">
            <div class="guide-title">Penalties interpretation</div>
            <div class="subtle">
                Number of deviations and deduction amount are two different measures.
                The dashboard sums <b>Number of Deviations</b> as total observed deviations,
                while <b>Penalties Amount</b> sums only the values that will actually be deducted.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    left, right = st.columns([1.15, 0.85])
    with left:
        st.markdown('<div class="section-card">', unsafe_allow_html=True)
        fig = penalties_deviation_chart(filtered_penalties)
        if fig:
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("No penalties records available for the selected filters.")
        st.markdown('</div>', unsafe_allow_html=True)
    with right:
        st.markdown('<div class="section-card">', unsafe_allow_html=True)
        quality_df = filtered[["Link Code", "penalty_qty", "penalty_amount", "lag_pp", "forecast_risk"]].copy()
        quality_df = quality_df.sort_values(["penalty_amount", "penalty_qty", "lag_pp"], ascending=[False, False, False]).head(15)
        quality_df.rename(columns={"penalty_qty": "Deviation Count", "penalty_amount": "Deduction Amount", "lag_pp": "Lag %", "forecast_risk": "Forecast Risk"}, inplace=True)
        st.subheader("Top Quality / Penalty Exposure")
        st.dataframe(quality_df, use_container_width=True, hide_index=True)
        st.markdown('</div>', unsafe_allow_html=True)

with tabs[4]:
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.subheader("Work Order Detail")
    preferred_cols = [
        "Link Code", "Work Order", "Region", "City", "District", "Project", "Subclass", "Stage",
        "Type", "Class", "Work Order Status", "Percentage of Completion", "planned_progress_pct",
        "lag_pp", "Cost", "WO Cost", "Created", "Assigned at", "Updated", "Updates",
        "Targeted Completion", "Updated Target Date", "effective_target",
        "Trench Progress", "Trench Scope", "MH/HH Progress", "MH/HH Scope",
        "Fiber Progress", "Fiber Scope", "Permits Progress", "Permits Scope",
        "PIP rejection count", "PAT rejection count", "Approval rejection count",
        "As-Built Rejection Count", "Handover Rejection Count",
        "Parent", "Number", "Request", "SOR Reference Number", "Target Area",
        "Notes", "Scope of Work", "penalty_qty", "penalty_amount", "forecast_risk"
    ]
    visible_cols = [c for c in preferred_cols if c in filtered_details.columns]
    work_detail = filtered_details[visible_cols].copy() if visible_cols else filtered_details.copy()
    st.dataframe(work_detail, use_container_width=True, hide_index=True)
    b1, b2 = st.columns(2)
    with b1:
        st.download_button("Export Work Order Detail (Excel)", data=to_excel_bytes(work_detail, "Workorder Detail"), file_name="workorder_detail_export.xlsx")
    with b2:
        st.download_button("Export Work Order Detail (CSV)", data=to_csv_bytes(work_detail), file_name="workorder_detail_export.csv")
    st.markdown('</div>', unsafe_allow_html=True)

with tabs[5]:
    st.markdown(
        """
        <div class="guide-box">
            <div class="guide-title">Dashboard calculation guide</div>
            <div class="subtle">
                <b>Actual Progress</b> comes from Percentage of Completion and is capped at 100% for KPI averaging.<br>
                <b>Planned Progress</b> is derived from elapsed time between start date and effective target date.<br>
                <b>Effective Target Date</b> = Updated Target Date when available, otherwise Targeted Completion.<br>
                <b>Lag %</b> = Planned Progress - Actual Progress.<br>
                <b>Critical Lag</b> means lag is 15% or higher.<br>
                <b>Need System Update</b> means Updates &lt; 5 and days since last update &gt; 5 days.<br>
                <b>Current Month Target Cost</b> is the sum of Cost for work orders whose target date falls in the snapshot month.<br>
                <b>Penalty Deduction Amount</b> sums only deductible values, while Number of Deviations sums all recorded deviations.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    c1, c2 = st.columns(2)
    with c1:
        st.markdown(
            """
            <div class="guide-box">
                <div class="guide-title">Region logic</div>
                <div class="subtle">
                    Region is normalized for dashboard filtering to four macro groups:
                    Western, Southern, Eastern, and Northern.
                    District / City mapping is taken from the District sheet first, then from link code fallback.
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with c2:
        missing_city = int((filtered["City"] == "Not Classified").sum())
        missing_district = int((filtered["District"] == "Not Classified").sum())
        dq_score = 100 - (((missing_city + missing_district) / max(len(filtered) * 2, 1)) * 100)
        st.metric("Data Quality Score", f"{dq_score:,.1f}%")
        st.metric("Missing City", f"{missing_city:,}")
        st.metric("Missing District", f"{missing_district:,}")
