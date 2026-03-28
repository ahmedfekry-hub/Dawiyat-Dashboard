
from __future__ import annotations

from pathlib import Path
import re
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(
    page_title="Dawiyat Executive PMO Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

DEFAULT_FILE = "Dawiyat Master Sheet.xlsx"


def find_sheet_name(xls: pd.ExcelFile, candidates: list[str]) -> str:
    normalized = {str(name).strip().lower(): name for name in xls.sheet_names}
    for candidate in candidates:
        key = candidate.strip().lower()
        if key in normalized:
            return normalized[key]
    for name in xls.sheet_names:
        low = str(name).strip().lower()
        if any(candidate.strip().lower() in low for candidate in candidates):
            return name
    raise ValueError(f"Expected sheet not found. Tried: {candidates}")


def clean_text_value(value):
    if pd.isna(value):
        return np.nan
    text = str(value).strip()
    if text.lower() in {"", "nan", "none", "null"}:
        return np.nan
    return text


def canonical_key(value) -> str:
    if pd.isna(value):
        return ""
    text = str(value).upper().strip()
    text = re.sub(r"[\.\-_/]+", " ", text)
    text = re.sub(r"[^A-Z0-9 ]+", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


CITY_MAP = {
    "JEDDAH": "Jeddah",
    "JIZAN": "Jizan",
    "TABUK": "Tabuk",
    "TAIF": "Taif",
    "AL BAHA": "Al Baha",
    "ALBAHA": "Al Baha",
    "RIYADH": "Riyadh",
    "MADINAH": "Madinah",
    "MEDINAH": "Madinah",
    "MAKKAH": "Makkah",
    "MECCA": "Makkah",
}


def normalize_city(value):
    text = clean_text_value(value)
    if pd.isna(text):
        return np.nan
    key = canonical_key(text)
    return CITY_MAP.get(key, str(text).strip().title())


DISTRICT_NORMALIZATION_MAP = {
    "AL EDABI": "AL-EDABI",
    "ASH SHUQAYRI": "ASH-SHUQAYRI",
}


def normalize_district(value):
    text = clean_text_value(value)
    if pd.isna(text):
        return np.nan
    key = canonical_key(text)
    if key in DISTRICT_NORMALIZATION_MAP:
        return DISTRICT_NORMALIZATION_MAP[key]
    text = re.sub(r"[\._]+", " ", str(text))
    text = re.sub(r"\s+", " ", text).strip()
    return text.title()


def choose_mode(series: pd.Series):
    s = series.dropna()
    if s.empty:
        return np.nan
    modes = s.mode(dropna=True)
    return modes.iloc[0] if not modes.empty else s.iloc[0]


def first_existing(df: pd.DataFrame, candidates: list[str]) -> str | None:
    cmap = {str(c).strip().lower(): c for c in df.columns}
    for c in candidates:
        if c.strip().lower() in cmap:
            return cmap[c.strip().lower()]
    for c in df.columns:
        low = str(c).strip().lower()
        if any(x.strip().lower() in low for x in candidates):
            return c
    return None


@st.cache_data(show_spinner=False)
def load_data(file_source=None):
    if file_source is None:
        file_source = Path(__file__).with_name(DEFAULT_FILE)

    xls = pd.ExcelFile(file_source)
    main_name = find_sheet_name(xls, ["Dawaiyat Service Tool", "service tool"])
    district_name = find_sheet_name(xls, ["District"])
    penalties_name = find_sheet_name(xls, ["Penalties", "Penalty"])

    main = pd.read_excel(file_source, sheet_name=main_name)
    district = pd.read_excel(file_source, sheet_name=district_name)
    penalties = pd.read_excel(file_source, sheet_name=penalties_name)

    main.columns = [str(c).strip() for c in main.columns]
    district.columns = [str(c).strip() for c in district.columns]
    penalties.columns = [str(c).strip() for c in penalties.columns]

    for df in [main, district, penalties]:
        for col in df.columns:
            if df[col].dtype == "object":
                df[col] = df[col].map(clean_text_value)

    main = main[main["Link Code"].notna()].copy()
    main["Link Code"] = main["Link Code"].astype(str).str.strip()

    district_link = first_existing(district, ["Link Code"])
    district_city = first_existing(district, ["City"])
    district_name_col = first_existing(district, ["District"])

    if district_link:
        district = district[district[district_link].notna()].copy()
        district[district_link] = district[district_link].astype(str).str.strip()
        if district_city:
            district["City_clean"] = district[district_city].map(normalize_city)
        else:
            district["City_clean"] = np.nan
        if district_name_col:
            district["District_clean"] = district[district_name_col].map(normalize_district)
        else:
            district["District_clean"] = np.nan

        district_map = (
            district.groupby(district_link, dropna=False)
            .agg(City=("City_clean", choose_mode), District=("District_clean", choose_mode))
            .reset_index()
            .rename(columns={district_link: "Link Code"})
        )
    else:
        district_map = pd.DataFrame(columns=["Link Code", "City", "District"])

    data = main.merge(district_map, on="Link Code", how="left")

    text_cols = [
        "Region", "Project", "Subclass", "Stage", "Year", "Work Order Status",
        "Type", "Class", "City", "District", "Link Code"
    ]
    for col in text_cols:
        if col not in data.columns:
            data[col] = np.nan

    data["City"] = data["City"].map(normalize_city)
    data["District"] = data["District"].map(normalize_district)

    data["Link Code"] = data["Link Code"].astype(str).str.strip()
    data["Region"] = data["Region"].fillna("Not Classified")
    data["Project"] = data["Project"].fillna("Not Classified")
    data["Subclass"] = data["Subclass"].fillna("Not Classified")
    data["Stage"] = data["Stage"].fillna("Not Classified")
    data["Year"] = data["Year"].astype("string").replace({"<NA>": np.nan}).fillna("Not Classified")
    data["Work Order Status"] = data["Work Order Status"].fillna("Open / Not Classified")
    data["Type"] = data["Type"].fillna("Not Classified")
    data["Class"] = data["Class"].fillna("Not Classified")
    data["City"] = data["City"].fillna("Not Classified")
    data["District"] = data["District"].fillna("Not Classified")

    date_cols = ["Created", "Assigned at", "In Progress at", "Updated", "Closed at", "Targeted Completion", "Updated Target Date"]
    for col in date_cols:
        if col in data.columns:
            data[col] = pd.to_datetime(data[col], errors="coerce")

    numeric_cols = [
        "Percentage of Completion", "WO Cost", "Cost", "Trench Progress", "Trench Scope",
        "MH/HH Progress", "MH/HH Scope", "Fiber Progress", "Fiber Scope", "ODBs Progress",
        "ODBs Scope", "ODFs Progress", "ODFs Scope", "JCL Progress", "JCL Scope",
        "FAT Progress", "FAT Scope", "PFAT Progress", "PFAT Scope", "SFAT Progress",
        "SFAT Scope", "Permits Progress", "Permits Scope", "PIP rejection count",
        "PAT rejection count", "Approval rejection count", "As-Built Rejection Count",
        "Handover Rejection Count", "Number of Buildings", "Number of Households", "Parcels",
        "Updates"
    ]
    for col in numeric_cols:
        if col in data.columns:
            data[col] = pd.to_numeric(data[col], errors="coerce")

    snapshot_date = pd.to_datetime(data["Updated"], errors="coerce").max()
    if pd.isna(snapshot_date):
        snapshot_date = pd.Timestamp.today().normalize()

    data["effective_target"] = data["Updated Target Date"].combine_first(data["Targeted Completion"])
    data["start_date"] = data["In Progress at"].combine_first(data["Assigned at"]).combine_first(data["Created"])

    elapsed_days = (snapshot_date - data["start_date"]).dt.days
    total_days = (data["effective_target"] - data["start_date"]).dt.days
    with np.errstate(divide="ignore", invalid="ignore"):
        data["planned_progress_pct"] = np.where(total_days > 0, np.clip((elapsed_days / total_days) * 100, 0, 100), np.nan)

    data["actual_progress_pct"] = pd.to_numeric(data["Percentage of Completion"], errors="coerce").clip(lower=0, upper=100)
    data["schedule_variance_pp"] = data["actual_progress_pct"] - data["planned_progress_pct"]
    data["is_complete"] = data["actual_progress_pct"] >= 100
    data["is_overdue"] = data["effective_target"].notna() & (snapshot_date > data["effective_target"]) & (~data["is_complete"])
    data["critical_lag"] = data["schedule_variance_pp"] <= -15

    civil_scope = data["Trench Scope"].fillna(0) + data["MH/HH Scope"].fillna(0)
    civil_progress = data["Trench Progress"].fillna(0) + data["MH/HH Progress"].fillna(0)
    data["civil_completion_pct"] = np.where(civil_scope > 0, (civil_progress / civil_scope) * 100, np.nan)

    fiber_scope_total = (
        data["Fiber Scope"].fillna(0) + data["ODBs Scope"].fillna(0) + data["ODFs Scope"].fillna(0) +
        data["JCL Scope"].fillna(0) + data["FAT Scope"].fillna(0) + data["PFAT Scope"].fillna(0) +
        data["SFAT Scope"].fillna(0)
    )
    fiber_progress_total = (
        data["Fiber Progress"].fillna(0) + data["ODBs Progress"].fillna(0) + data["ODFs Progress"].fillna(0) +
        data["JCL Progress"].fillna(0) + data["FAT Progress"].fillna(0) + data["PFAT Progress"].fillna(0) +
        data["SFAT Progress"].fillna(0)
    )
    data["fiber_completion_pct"] = np.where(fiber_scope_total > 0, (fiber_progress_total / fiber_scope_total) * 100, np.nan)

    budget = data["WO Cost"].fillna(0)
    actual_cost = data["Cost"].fillna(0)
    data["EV"] = budget * (data["actual_progress_pct"].fillna(0) / 100)
    data["PV"] = budget * (data["planned_progress_pct"].fillna(0) / 100)
    data["AC"] = actual_cost
    data["SPI"] = np.where(data["PV"] > 0, data["EV"] / data["PV"], np.nan)
    data["CPI"] = np.where(data["AC"] > 0, data["EV"] / data["AC"], np.nan)

    actual_ratio = data["actual_progress_pct"] / 100
    elapsed_days_safe = np.maximum((snapshot_date - data["start_date"]).dt.days, 1)
    estimated_total_duration = np.where(actual_ratio > 0, elapsed_days_safe / actual_ratio, np.nan)
    data["forecast_completion_date"] = data["start_date"] + pd.to_timedelta(estimated_total_duration, unit="D")
    data["forecast_delay_days"] = (data["forecast_completion_date"] - data["effective_target"]).dt.days
    data["forecast_risk"] = np.select(
        [data["forecast_delay_days"] > 30, data["forecast_delay_days"] > 0, data["forecast_delay_days"] <= 0],
        ["High delay risk", "Moderate delay risk", "On forecast"],
        default="Insufficient data",
    )

    data["month"] = pd.to_datetime(data["Updated"], errors="coerce").dt.to_period("M").astype(str)

    # Penalties handling: supports the revised Penalties sheet structure
    penalties = penalties.copy()
    cluster_col = first_existing(penalties, ["Cluster Name", "Link Code", "Cluster"])
    penalties["Link Code"] = penalties[cluster_col].astype(str).str.strip() if cluster_col else np.nan

    qty_col = first_existing(penalties, ["Number of Deviations", "Number", "Qty", "Quantity"])
    penalties["penalty_qty"] = pd.to_numeric(penalties[qty_col], errors="coerce").fillna(0) if qty_col else 0

    amt_col = first_existing(penalties, ["Penalties Amount", "Penalty Amount", "Deduction"])
    penalties["penalty_amount"] = pd.to_numeric(penalties[amt_col], errors="coerce").fillna(0) if amt_col else 0

    reason_col = first_existing(penalties, ["Deviation name", "Penalty Reason", "Reason", "Impl # of Penalty"])
    if reason_col:
        penalties["penalty_reason"] = penalties[reason_col].fillna("Not Classified")
    else:
        penalties["penalty_reason"] = "Not Classified"

    assign_date_col = first_existing(penalties, ["Assigned Date", "Date"])
    penalties["Assigned Date"] = pd.to_datetime(penalties[assign_date_col], errors="coerce") if assign_date_col else pd.NaT

    region_col = first_existing(penalties, ["Region"])
    penalties["Region"] = penalties[region_col].fillna("Not Classified") if region_col else "Not Classified"

    penalties_geo = data[["Link Code", "City", "District", "Region"]].drop_duplicates()
    penalties = penalties.merge(penalties_geo, on="Link Code", how="left", suffixes=("", "_from_main"))
    penalties["City"] = penalties.get("City").fillna(penalties.get("City_from_main")).map(normalize_city).fillna("Not Classified")
    penalties["District"] = penalties.get("District").fillna(penalties.get("District_from_main")).map(normalize_district).fillna("Not Classified")
    penalties["Region"] = penalties["Region"].fillna(penalties.get("Region_from_main")).fillna("Not Classified")

    penalties["penalty_cases"] = np.where(penalties["Link Code"].notna(), 1, 0)

    penalties_agg = (
        penalties.groupby("Link Code", dropna=False)
        .agg(
            penalty_cases=("penalty_cases", "sum"),
            penalty_qty=("penalty_qty", "sum"),
            penalty_amount=("penalty_amount", "sum"),
        )
        .reset_index()
    )
    data = data.merge(penalties_agg, on="Link Code", how="left")
    for col in ["penalty_cases", "penalty_qty", "penalty_amount"]:
        data[col] = pd.to_numeric(data[col], errors="coerce").fillna(0)

    return data, penalties, snapshot_date


def fmt_pct(value, digits: int = 1):
    return "-" if pd.isna(value) else f"{value:,.{digits}f}%"


def fmt_ratio(value):
    return "-" if pd.isna(value) else f"{value:,.2f}"


def fmt_int(value):
    return "-" if pd.isna(value) else f"{int(round(value)):,.0f}"


def fmt_money(value):
    if pd.isna(value):
        return "-"
    value = float(value)
    if abs(value) >= 1_000_000:
        return f"${value/1_000_000:,.2f}M"
    if abs(value) >= 1_000:
        return f"${value/1_000:,.1f}K"
    return f"${value:,.0f}"


def plot_template(theme: str) -> str:
    return "plotly_dark" if theme == "Dark" else "plotly_white"


def gauge_color_steps(kind="progress"):
    if kind == "risk":
        return [
            {"range": [0, 40], "color": "rgba(34,197,94,0.90)"},
            {"range": [40, 70], "color": "rgba(245,158,11,0.90)"},
            {"range": [70, 100], "color": "rgba(239,68,68,0.90)"},
        ]
    return [
        {"range": [0, 50], "color": "rgba(239,68,68,0.90)"},
        {"range": [50, 80], "color": "rgba(245,158,11,0.90)"},
        {"range": [80, 100], "color": "rgba(34,197,94,0.90)"},
    ]


def render_metric_card(title: str, value: str, subtitle: str = "", glow: str = "blue"):
    accent = {
        "blue": "#38BDF8",
        "green": "#22C55E",
        "orange": "#F59E0B",
        "red": "#EF4444",
        "violet": "#8B5CF6",
    }.get(glow, "#38BDF8")
    st.markdown(
        f"""
        <div class="metric-tile">
            <div class="metric-kicker">{title}</div>
            <div class="metric-value" style="text-shadow: 0 0 22px {accent}66;">{value}</div>
            <div class="metric-subtitle">{subtitle}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def apply_theme_css(theme: str):
    dark = theme == "Dark"
    bg = "#07111f" if dark else "#eff4fb"
    card = "rgba(10,23,43,0.86)" if dark else "rgba(255,255,255,0.96)"
    text = "#ECF5FF" if dark else "#10233c"
    muted = "#94A3B8" if dark else "#667085"
    border = "rgba(148,163,184,0.16)" if dark else "rgba(15,23,42,0.08)"
    card_shadow = "0 10px 35px rgba(0,0,0,0.28)" if dark else "0 10px 35px rgba(15,23,42,0.08)"
    st.markdown(
        f"""
        <style>
            .stApp {{
                background:
                    radial-gradient(circle at 0% 0%, rgba(56,189,248,0.18), transparent 26%),
                    radial-gradient(circle at 100% 0%, rgba(249,115,22,0.16), transparent 26%),
                    radial-gradient(circle at 80% 80%, rgba(139,92,246,0.12), transparent 20%),
                    {bg};
                color: {text};
            }}
            .block-container {{ padding-top: 1.1rem; padding-bottom: 1.8rem; max-width: 1650px; }}
            [data-testid="stSidebar"] {{ background: linear-gradient(180deg, rgba(7,17,31,0.98), rgba(15,23,42,0.98)); }}
            [data-testid="stSidebar"] * {{ color: #E5EEF9; }}
            .hero-box {{
                background: linear-gradient(135deg, rgba(15,23,42,0.78), rgba(8,24,48,0.86));
                border: 1px solid rgba(148,163,184,0.16);
                border-radius: 28px; padding: 24px 28px; box-shadow: 0 18px 44px rgba(0,0,0,0.28); margin-bottom: 1rem;
            }}
            .hero-title {{ font-size: 2rem; font-weight: 800; letter-spacing: 0.03em; margin-bottom: 0.35rem; }}
            .hero-title span {{ color: #38BDF8; }}
            .hero-subtitle {{ color: #AFC2DB; font-size: 0.98rem; }}
            .metric-tile {{
                background: {card}; border: 1px solid {border}; border-radius: 22px; padding: 16px 18px; min-height: 124px;
                box-shadow: {card_shadow}; backdrop-filter: blur(10px);
            }}
            .metric-kicker {{ color: {muted}; font-size: 0.82rem; font-weight: 700; letter-spacing: 0.08em; text-transform: uppercase; }}
            .metric-value {{ font-size: 2rem; font-weight: 800; line-height: 1.15; margin-top: 0.35rem; }}
            .metric-subtitle {{ color: {muted}; font-size: 0.92rem; margin-top: 0.35rem; }}
            .panel-card {{
                background: {card}; border: 1px solid {border}; border-radius: 24px; padding: 8px 12px 10px 12px;
                box-shadow: {card_shadow}; backdrop-filter: blur(10px);
            }}
            .filter-note {{
                background: rgba(56,189,248,0.10); border: 1px solid rgba(56,189,248,0.25); border-radius: 16px;
                padding: 12px 14px; margin-top: 0.5rem; color: #DCEAF9; font-size: 0.92rem;
            }}
            .guide-box {{
                background: {card}; border: 1px solid {border}; border-left: 4px solid #38BDF8; border-radius: 18px;
                padding: 18px 20px; box-shadow: {card_shadow};
            }}
        </style>
        """,
        unsafe_allow_html=True,
    )


def cascade_multiselect(label: str, df: pd.DataFrame, column: str):
    options = sorted([x for x in df[column].dropna().astype(str).unique().tolist() if x and x != "nan"])
    return st.sidebar.multiselect(label, options)


st.sidebar.title("PMO Control Panel")
theme_mode = st.sidebar.radio("Theme", ["Dark", "Light"], horizontal=True, index=0)
apply_theme_css(theme_mode)

uploaded_file = st.sidebar.file_uploader("Upload refreshed Dawiyat workbook", type=["xlsx"])
source = uploaded_file if uploaded_file is not None else None
data, penalties, snapshot_date = load_data(source)

st.markdown(
    f"""
    <div class="hero-box">
        <div class="hero-title">DAWIYAT <span>EXECUTIVE PMO</span> DASHBOARD</div>
        <div class="hero-subtitle">
            Decision-support cockpit for PMO leadership, Operations Management, and top-management review.
            Snapshot date: <b>{snapshot_date.strftime('%d %b %Y')}</b>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

st.sidebar.markdown("### Dashboard Filters")
working = data.copy()
region_sel = cascade_multiselect("Region", working, "Region")
if region_sel:
    working = working[working["Region"].isin(region_sel)]

city_sel = cascade_multiselect("City", working, "City")
if city_sel:
    working = working[working["City"].isin(city_sel)]

district_sel = cascade_multiselect("District", working, "District")
if district_sel:
    working = working[working["District"].isin(district_sel)]

link_code_sel = cascade_multiselect("Link Code", working, "Link Code")
if link_code_sel:
    working = working[working["Link Code"].isin(link_code_sel)]

project_sel = cascade_multiselect("Project", working, "Project")
if project_sel:
    working = working[working["Project"].isin(project_sel)]

stage_sel = cascade_multiselect("Stage", working, "Stage")
if stage_sel:
    working = working[working["Stage"].isin(stage_sel)]

year_sel = cascade_multiselect("Year", working, "Year")
if year_sel:
    working = working[working["Year"].isin(year_sel)]

status_sel = cascade_multiselect("Work Order Status", working, "Work Order Status")
if status_sel:
    working = working[working["Work Order Status"].isin(status_sel)]

type_sel = cascade_multiselect("Type", working, "Type")
if type_sel:
    working = working[working["Type"].isin(type_sel)]

class_sel = cascade_multiselect("Class", working, "Class")
if class_sel:
    working = working[working["Class"].isin(class_sel)]

subclass_sel = cascade_multiselect("Subclass", working, "Subclass")
if subclass_sel:
    working = working[working["Subclass"].isin(subclass_sel)]

filtered = working.copy()

st.sidebar.markdown(
    """
    <div class="filter-note">
    Filters are cascading from the actual Dawiyat row relationships.<br>
    Example: after selecting <b>Jeddah</b>, only link codes and districts belonging to Jeddah remain available.
    </div>
    """,
    unsafe_allow_html=True,
)

if filtered.empty:
    st.warning("No records match the selected filters.")
    st.stop()

work_orders = filtered["Work Order"].nunique() if "Work Order" in filtered.columns else len(filtered)
link_codes = filtered["Link Code"].nunique()
avg_actual = filtered["actual_progress_pct"].mean()
avg_planned = filtered["planned_progress_pct"].mean()
avg_civil = filtered["civil_completion_pct"].mean()
avg_fiber = filtered["fiber_completion_pct"].mean()
on_track_pct = (filtered["schedule_variance_pp"] >= 0).mean() * 100
critical_lag_cnt = int(filtered["critical_lag"].sum())
overdue_cnt = int(filtered["is_overdue"].sum())
forecast_high = int((filtered["forecast_risk"] == "High delay risk").sum())
penalty_cases = int(filtered["penalty_cases"].sum())
penalty_amount = float(filtered["penalty_amount"].sum())
penalty_qty = float(filtered["penalty_qty"].sum())

bac = filtered["WO Cost"].fillna(0).sum()
ac = filtered["AC"].fillna(0).sum()
ev = filtered["EV"].fillna(0).sum()
pv = filtered["PV"].fillna(0).sum()
spi = ev / pv if pv > 0 else np.nan
cpi = ev / ac if ac > 0 else np.nan
eac = ac + np.maximum(bac - ev, 0) / cpi if pd.notna(cpi) and cpi > 0 else np.nan

m1, m2, m3, m4, m5, m6 = st.columns(6)
with m1:
    render_metric_card("Project Health", fmt_pct(on_track_pct, 0), "Work orders on or above plan", "green")
with m2:
    render_metric_card("Budget Performance", fmt_money(ac), f"BAC {fmt_money(bac)}", "green")
with m3:
    render_metric_card("SPI", fmt_ratio(spi), f"Planned {fmt_pct(avg_planned)}", "orange")
with m4:
    render_metric_card("CPI", fmt_ratio(cpi), "Cost efficiency based on EV / AC", "blue")
with m5:
    render_metric_card("EAC", fmt_money(eac), "Estimated at completion", "violet")
with m6:
    render_metric_card("Penalty Exposure", fmt_int(penalty_cases), f"Qty {fmt_int(penalty_qty)} | {fmt_money(penalty_amount)}", "red")

c1, c2, c3 = st.columns([1.2, 1.2, 1])
with c1:
    st.markdown('<div class="panel-card">', unsafe_allow_html=True)
    trend = (
        filtered.dropna(subset=["Updated"])
        .groupby("month", dropna=False)
        .agg(actual=("actual_progress_pct", "mean"), planned=("planned_progress_pct", "mean"))
        .reset_index()
    )
    if not trend.empty:
        fig = go.Figure()
        fig.add_trace(go.Scatter(x=trend["month"], y=trend["actual"], mode="lines+markers", name="Actual %"))
        fig.add_trace(go.Scatter(x=trend["month"], y=trend["planned"], mode="lines+markers", name="Planned %"))
        fig.add_hline(y=100, line_dash="dot", opacity=0.35)
        fig.update_layout(title="Progress Trend | Planned vs Actual", template=plot_template(theme_mode), height=390,
                          margin=dict(l=20, r=20, t=55, b=20), yaxis_title="Progress %", xaxis_title="")
        st.plotly_chart(fig, use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

with c2:
    st.markdown('<div class="panel-card">', unsafe_allow_html=True)
    city_perf = (
        filtered.groupby("City", dropna=False)
        .agg(actual=("actual_progress_pct", "mean"), planned=("planned_progress_pct", "mean"),
             wo=("Work Order", "count"), penalty_cases=("penalty_cases", "sum"))
        .reset_index()
    )
    city_perf["variance"] = city_perf["actual"] - city_perf["planned"]
    city_perf = city_perf.sort_values("actual", ascending=True)
    fig = px.bar(city_perf, x="actual", y="City", color="variance", orientation="h", text="actual",
                 title="City Performance Snapshot", template=plot_template(theme_mode),
                 hover_data=["planned", "wo", "penalty_cases"])
    fig.update_traces(texttemplate="%{text:.1f}%", textposition="outside")
    fig.update_layout(height=390, margin=dict(l=10, r=10, t=55, b=10), xaxis_title="Actual Progress %", yaxis_title="")
    st.plotly_chart(fig, use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

with c3:
    st.markdown('<div class="panel-card">', unsafe_allow_html=True)
    fig = go.Figure(go.Indicator(
        mode="gauge+number", value=0 if pd.isna(avg_actual) else avg_actual, number={"suffix": "%"},
        title={"text": "Overall Completion"},
        gauge={"axis": {"range": [0, 100]}, "bar": {"color": "rgba(250,204,21,0.90)"}, "steps": gauge_color_steps("progress")}
    ))
    fig.update_layout(template=plot_template(theme_mode), height=190, margin=dict(l=15, r=15, t=50, b=10))
    st.plotly_chart(fig, use_container_width=True)

    risk_score = min((forecast_high / max(len(filtered), 1)) * 100 + (critical_lag_cnt / max(len(filtered), 1)) * 100, 100)
    fig2 = go.Figure(go.Indicator(
        mode="gauge+number", value=risk_score, number={"suffix": "%"}, title={"text": "Forecast Risk Exposure"},
        gauge={"axis": {"range": [0, 100]}, "bar": {"color": "rgba(244,63,94,0.90)"}, "steps": gauge_color_steps("risk")}
    ))
    fig2.update_layout(template=plot_template(theme_mode), height=190, margin=dict(l=15, r=15, t=50, b=10))
    st.plotly_chart(fig2, use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

r1, r2, r3 = st.columns([1.2, 1, 1])
with r1:
    st.markdown('<div class="panel-card">', unsafe_allow_html=True)
    district_summary = (
        filtered.groupby(["City", "District"], dropna=False)
        .agg(actual=("actual_progress_pct", "mean"), planned=("planned_progress_pct", "mean"),
             civil=("civil_completion_pct", "mean"), fiber=("fiber_completion_pct", "mean"),
             spi=("SPI", "mean"), cpi=("CPI", "mean"), overdue=("is_overdue", "sum"))
        .reset_index()
    )
    district_summary["variance"] = district_summary["actual"] - district_summary["planned"]
    fig = px.bar(district_summary.sort_values("variance"), x="District", y="variance", color="City",
                 title="District Schedule Variance", template=plot_template(theme_mode),
                 hover_data=["actual", "planned", "civil", "fiber", "spi", "cpi", "overdue"])
    fig.add_hline(y=0, line_dash="dash", opacity=0.4)
    fig.update_layout(height=410, margin=dict(l=15, r=15, t=55, b=10), yaxis_title="Actual - Planned (pp)")
    st.plotly_chart(fig, use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

with r2:
    st.markdown('<div class="panel-card">', unsafe_allow_html=True)
    milestone_rows = []
    milestone_pairs = [
        ("Trench Progress", "Trench Progress"), ("MH/HH Progress", "MH/HH Progress"),
        ("Fiber Progress", "Fiber Progress"), ("ODBs Progress", "ODBs Progress"),
        ("ODFs Progress", "ODFs Progress"), ("JCL Progress", "JCL Progress"),
        ("FAT Progress", "FAT Progress"), ("PFAT Progress", "PFAT Progress"),
        ("SFAT Progress", "SFAT Progress"), ("Permits Progress", "Permits Progress"),
    ]
    for col, label in milestone_pairs:
        if col in filtered.columns:
            val = pd.to_numeric(filtered[col], errors="coerce").sum()
            if pd.notna(val):
                milestone_rows.append({"Milestone": label, "Total": val})
    milestone_df = pd.DataFrame(milestone_rows).sort_values("Total", ascending=True)
    if not milestone_df.empty:
        fig = px.bar(milestone_df, x="Total", y="Milestone", orientation="h", text="Total",
                     title="Milestone Delivered Quantity", template=plot_template(theme_mode))
        fig.update_layout(height=410, margin=dict(l=15, r=15, t=55, b=10), xaxis_title="Delivered Quantity", yaxis_title="")
        st.plotly_chart(fig, use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

with r3:
    st.markdown('<div class="panel-card">', unsafe_allow_html=True)
    summary_tiles = pd.DataFrame({"Status": ["On Track", "Overdue", "Critical Lag", "High Forecast Risk"],
                                  "Value": [on_track_pct, overdue_cnt, critical_lag_cnt, forecast_high]})
    fig = px.treemap(summary_tiles, path=["Status"], values="Value", color="Value",
                     title="Overall Status Mix", template=plot_template(theme_mode))
    fig.update_layout(height=205, margin=dict(l=10, r=10, t=45, b=10))
    st.plotly_chart(fig, use_container_width=True)

    stage_mix = filtered.groupby("Stage", dropna=False)["Work Order"].count().reset_index(name="count").sort_values("count", ascending=False)
    fig2 = px.pie(stage_mix, names="Stage", values="count", hole=0.58, title="Stage Distribution", template=plot_template(theme_mode))
    fig2.update_layout(height=205, margin=dict(l=10, r=10, t=45, b=10))
    st.plotly_chart(fig2, use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

st.markdown("### Detailed PMO Analysis")
tab1, tab2, tab3, tab4 = st.tabs(["Schedule & Cost", "Penalties & Quality", "Work Order Detail", "Dashboard Guide"])

with tab1:
    d1, d2 = st.columns([1.2, 1])
    with d1:
        top_delay = (
            filtered[filtered["forecast_delay_days"].notna()]
            .nlargest(15, "forecast_delay_days")
            [["Work Order", "Link Code", "City", "District", "Stage", "Subclass", "actual_progress_pct",
              "planned_progress_pct", "forecast_delay_days", "effective_target", "forecast_completion_date"]]
            .rename(columns={"actual_progress_pct": "Actual %", "planned_progress_pct": "Planned %",
                             "forecast_delay_days": "Forecast Delay Days", "effective_target": "Target Date",
                             "forecast_completion_date": "Forecast Finish"})
        )
        st.dataframe(top_delay, use_container_width=True, hide_index=True)
    with d2:
        forecast_profile = filtered.groupby("forecast_risk", dropna=False)["Work Order"].count().reset_index(name="count")
        fig = px.pie(forecast_profile, names="forecast_risk", values="count", hole=0.58,
                     title="Forecast Delay Risk Profile", template=plot_template(theme_mode))
        fig.update_layout(height=380, margin=dict(l=10, r=10, t=55, b=10))
        st.plotly_chart(fig, use_container_width=True)

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Work Orders", fmt_int(work_orders))
    k2.metric("Civil Completion", fmt_pct(avg_civil))
    k3.metric("Fiber Completion", fmt_pct(avg_fiber))
    k4.metric("Planned vs Actual Gap", f"{(avg_actual - avg_planned):+.1f} pp" if pd.notna(avg_actual) and pd.notna(avg_planned) else "-")

with tab2:
    filtered_pen = penalties.copy()
    if region_sel:
        filtered_pen = filtered_pen[filtered_pen["Region"].isin(region_sel)]
    if city_sel:
        filtered_pen = filtered_pen[filtered_pen["City"].isin(city_sel)]
    if district_sel:
        filtered_pen = filtered_pen[filtered_pen["District"].isin(district_sel)]

    p1, p2 = st.columns([1, 1])
    with p1:
        if not filtered_pen.empty:
            pen_city = filtered_pen.groupby("City", dropna=False).agg(qty=("penalty_qty", "sum"), amount=("penalty_amount", "sum")).reset_index().sort_values("qty", ascending=False)
            fig = px.bar(pen_city, x="City", y="qty", color="amount", title="Penalty Quantity by City",
                         template=plot_template(theme_mode), text="qty")
            fig.update_layout(height=380, margin=dict(l=10, r=10, t=55, b=10), yaxis_title="Penalty Qty")
            st.plotly_chart(fig, use_container_width=True)
    with p2:
        if not filtered_pen.empty:
            pen_type = filtered_pen.groupby("penalty_reason", dropna=False)["penalty_qty"].sum().reset_index().sort_values("penalty_qty", ascending=False).head(12)
            fig = px.bar(pen_type, x="penalty_qty", y="penalty_reason", orientation="h", title="Top Penalty Reasons",
                         template=plot_template(theme_mode), text="penalty_qty")
            fig.update_layout(height=380, margin=dict(l=10, r=10, t=55, b=10), xaxis_title="Penalty Qty", yaxis_title="")
            st.plotly_chart(fig, use_container_width=True)
    st.dataframe(filtered_pen, use_container_width=True, hide_index=True)

with tab3:
    detail = filtered.copy()
    detail["Target Date"] = pd.to_datetime(detail["effective_target"]).dt.date
    detail["Forecast Finish"] = pd.to_datetime(detail["forecast_completion_date"]).dt.date
    detail["Actual %"] = detail["actual_progress_pct"].round(1)
    detail["Planned %"] = detail["planned_progress_pct"].round(1)
    detail["Civil %"] = detail["civil_completion_pct"].round(1)
    detail["Fiber %"] = detail["fiber_completion_pct"].round(1)
    detail["SPI"] = detail["SPI"].round(2)
    detail["CPI"] = detail["CPI"].round(2)
    detail["Forecast Delay Days"] = detail["forecast_delay_days"].round(0)
    show_cols = [
        "Link Code", "Work Order", "Region", "City", "District", "Project", "Stage", "Subclass",
        "Work Order Status", "Type", "Class", "Target Date", "Forecast Finish", "Actual %", "Planned %",
        "Civil %", "Fiber %", "SPI", "CPI", "WO Cost", "Cost", "penalty_cases", "penalty_qty", "penalty_amount", "forecast_risk"
    ]
    show_cols = [c for c in show_cols if c in detail.columns]
    st.dataframe(detail[show_cols].sort_values(["Region", "City", "District", "Target Date"], na_position="last"),
                 use_container_width=True, hide_index=True)

with tab4:
    st.markdown(
        """
        <div class="guide-box">
        <b>Dashboard logic</b><br><br>
        • Contractor commitment is based on <i>Targeted Completion</i> and, when available, <i>Updated Target Date</i>.<br>
        • Actual overall progress is based on <i>Percentage of Completion</i>.<br>
        • Civil completion uses Trench + MH/HH progress against available civil scope.<br>
        • Fiber completion uses Fiber + ODB + ODF + JCL + FAT + PFAT + SFAT progress against available fiber scope.<br>
        • Planned progress is estimated from elapsed duration versus current effective target date.<br>
        • SPI / CPI are executive approximations from the available Dawiyat fields only.<br>
        • Cascading filters are based on actual row relationships, so selecting a city shows only its related districts.<br>
        • For Dawiyat, district names never reassign the city. Example: <i>Jeddah + Riyadh District</i> remains under Jeddah.<br>
        • The revised Penalties sheet is already linked by <i>Cluster Name / Link Code</i> and merged into the dashboard.<br>
        </div>
        """,
        unsafe_allow_html=True,
    )
