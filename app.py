
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
    return pd.to_numeric(series, errors="coerce").fillna(default)


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


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Workorder Detail")
    output.seek(0)
    return output.getvalue()


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


def apply_theme_css(theme_mode: str):
    dark = theme_mode == "Dark"
    bg = "#081426" if dark else "#F4F7FB"
    card = "#0f1f38" if dark else "#FFFFFF"
    text = "#EAF2FF" if dark else "#182230"
    muted = "#8BA2C8" if dark else "#667085"
    border = "rgba(148,163,184,0.16)" if dark else "rgba(15,23,42,0.08)"
    glow = "0 12px 32px rgba(59,130,246,0.18)" if dark else "0 10px 28px rgba(15,23,42,0.07)"
    st.markdown(
        f"""
        <style>
        .stApp {{
            background:
                radial-gradient(circle at top left, rgba(59,130,246,0.16), transparent 28%),
                radial-gradient(circle at top right, rgba(245,158,11,0.10), transparent 24%),
                {bg};
            color: {text};
        }}
        .block-container {{padding-top: 1.0rem; padding-bottom: 2rem;}}
        section[data-testid="stSidebar"] {{
            background: linear-gradient(180deg, rgba(15,31,56,0.95), rgba(8,20,38,0.98));
            border-right: 1px solid rgba(148,163,184,0.15);
        }}
        section[data-testid="stSidebar"] * {{ color: #EAF2FF !important; }}
        .top-banner {{
            background: linear-gradient(135deg, rgba(59,130,246,0.22), rgba(245,158,11,0.10));
            border: 1px solid {border}; border-radius: 24px; padding: 22px 24px; margin-bottom: 14px; box-shadow: {glow};
        }}
        .metric-card {{
            background: {card}; border: 1px solid {border}; border-radius: 18px; padding: 18px 18px 14px 18px;
            min-height: 132px; box-shadow: {glow};
        }}
        .metric-title {{ color: {muted}; font-size: 0.92rem; margin-bottom: 6px; }}
        .metric-value {{ color: {text}; font-size: 2rem; font-weight: 700; line-height: 1.15; margin-bottom: 8px; }}
        .metric-subtitle {{ color: {muted}; font-size: 0.9rem; }}
        .section-card {{ background: {card}; border: 1px solid {border}; border-radius: 22px; padding: 10px 14px 2px 14px; box-shadow: {glow}; }}
        .guide-box {{ background: {card}; border-left: 4px solid #3B82F6; border-radius: 14px; border: 1px solid {border}; padding: 14px 16px; margin: 10px 0px; }}
        .subtle {{ color: {muted}; font-size: 0.95rem; }}
        .small-chip {{ display:inline-block; padding: 6px 10px; border-radius: 999px; background: rgba(59,130,246,0.12); color: {text}; font-size: 0.85rem; border: 1px solid {border}; margin-right: 8px; margin-top: 4px; }}
        .stDataFrame, div[data-testid="stTable"] {{ border-radius: 16px; overflow: hidden; border: 1px solid {border}; }}
        </style>
        """,
        unsafe_allow_html=True,
    )


@st.cache_data(show_spinner=False)
def load_data(file_source=None):
    if file_source is None:
        file_source = Path(__file__).with_name(DEFAULT_FILE)
    xls = pd.ExcelFile(file_source)

    # load sheets
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

    # district mapping
    if not district.empty:
        d_link = first_existing(district, ["Link Code"])
        d_city = first_existing(district, ["City"])
        d_dist = first_existing(district, ["District"])
        if d_link:
            district = district[district[d_link].notna()].copy()
            district[d_link] = district[d_link].astype(str).str.strip()
            district_map = district.groupby(d_link, dropna=False).agg(
                City=(d_city, choose_mode) if d_city else (d_link, lambda s: np.nan),
                District=(d_dist, choose_mode) if d_dist else (d_link, lambda s: np.nan),
            ).reset_index().rename(columns={d_link: "Link Code"})
            data = main.merge(district_map, on="Link Code", how="left")

            district["p1"] = district[d_link].astype(str).str.split("-").str[0].str.upper()
            district["p3"] = district[d_link].astype(str).str.split("-").str[2].str.upper()
            city_prefix_map = district.groupby("p1")[d_city].agg(choose_mode).dropna().to_dict() if d_city else {}
            dist_prefix_map = district.groupby("p3")[d_dist].agg(choose_mode).dropna().to_dict() if d_dist else {}
        else:
            data = main.copy()
            city_prefix_map, dist_prefix_map = {}, {}
    else:
        data = main.copy()
        city_prefix_map, dist_prefix_map = {}, {}

    data["p1"] = data["Link Code"].astype(str).str.split("-").str[0].str.upper()
    data["p3"] = data["Link Code"].astype(str).str.split("-").str[2].str.upper()
    if "City" not in data.columns:
        data["City"] = np.nan
    if "District" not in data.columns:
        data["District"] = np.nan
    data["City"] = data["City"].fillna(data["p1"].map(city_prefix_map))
    data["District"] = data["District"].fillna(data["p3"].map(dist_prefix_map))
    # final fallback from link code segment to reduce unknowns
    data["City"] = data["City"].fillna(data["p1"].replace({"JED":"Jeddah","TAI":"Taif","TAB":"Tabuk","BAH":"Al Baha","SAM":"Jizan","SHA":"Jizan","SHU":"Jizan","EDA":"Jizan","DAI":"Jizan","MOY":"Taif","TUR":"Taif"}))
    data["District"] = data["District"].fillna(data["p3"].str.title())

    # dates and numeric columns
    date_candidates = ["Created", "Assigned at", "In Progress at", "Updated", "Closed at", "Targeted Completion", "Updated Target Date"]
    for col in date_candidates:
        if col in data.columns:
            data[col] = pd.to_datetime(data[col], errors="coerce")
    num_cols = ["Percentage of Completion","WO Cost","Cost","Updates","Trench Progress","Trench Scope","MH/HH Progress","MH/HH Scope","Fiber Progress","Fiber Scope","ODBs Progress","ODBs Scope","ODFs Progress","ODFs Scope","JCL Progress","JCL Scope","FAT Progress","FAT Scope","PFAT Progress","PFAT Scope","SFAT Progress","SFAT Scope","Permits Progress","Permits Scope","PIP rejection count","PAT rejection count","Approval rejection count","As-Built Rejection Count","Handover Rejection Count"]
    for col in num_cols:
        if col in data.columns:
            data[col] = pd.to_numeric(data[col], errors="coerce")

    snapshot_date = data["Updated"].dropna().max() if "Updated" in data.columns else pd.NaT
    if pd.isna(snapshot_date):
        snapshot_date = pd.Timestamp.today().normalize()

    targeted = data["Targeted Completion"] if "Targeted Completion" in data.columns else pd.Series(pd.NaT, index=data.index)
    updated_target = data["Updated Target Date"] if "Updated Target Date" in data.columns else pd.Series(pd.NaT, index=data.index)
    data["effective_target"] = updated_target.combine_first(targeted)
    start_date = data["In Progress at"] if "In Progress at" in data.columns else pd.Series(pd.NaT, index=data.index)
    assigned = data["Assigned at"] if "Assigned at" in data.columns else pd.Series(pd.NaT, index=data.index)
    created = data["Created"] if "Created" in data.columns else pd.Series(pd.NaT, index=data.index)
    data["start_date"] = start_date.combine_first(assigned).combine_first(created)

    elapsed = (snapshot_date - data["start_date"]).dt.days
    total = (data["effective_target"] - data["start_date"]).dt.days
    with np.errstate(divide="ignore", invalid="ignore"):
        data["planned_progress_pct"] = np.where(total > 0, np.clip((elapsed / total) * 100, 0, 100), np.nan)

    data["actual_progress_pct"] = pd.to_numeric(data.get("Percentage of Completion"), errors="coerce")
    data["actual_progress_capped"] = data["actual_progress_pct"].clip(lower=0, upper=100)
    data["lag_pp"] = data["planned_progress_pct"] - data["actual_progress_capped"]
    data["is_complete"] = data["actual_progress_capped"] >= 100
    data["is_overdue"] = (snapshot_date > data["effective_target"]) & (~data["is_complete"]) & data["effective_target"].notna()
    data["critical_lag"] = data["lag_pp"] >= 15
    data["days_since_update"] = (snapshot_date - data.get("Updated", pd.Series(pd.NaT, index=data.index))).dt.days
    data["needs_system_update"] = (safe_numeric(data.get("Updates", pd.Series(np.nan, index=data.index)), np.nan) < 5) & (data["days_since_update"] > 5)

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

    data["civil_completion_pct"] = ratio_pct(data.get("Trench Progress"), data.get("Trench Scope"))
    data["mhhh_completion_pct"] = ratio_pct(data.get("MH/HH Progress"), data.get("MH/HH Scope"))
    data["fiber_completion_pct"] = ratio_pct(data.get("Fiber Progress"), data.get("Fiber Scope"))
    data["permits_completion_pct"] = ratio_pct(data.get("Permits Progress"), data.get("Permits Scope"))

    for col in ["Year", "Work Order Status", "Type", "Class", "Project", "Subclass", "Stage", "Region", "City", "District"]:
        if col not in data.columns:
            data[col] = np.nan
    data["Year"] = data["Year"].fillna(pd.to_datetime(data["effective_target"], errors="coerce").dt.year.astype("Int64"))
    for col in ["Work Order Status", "Type", "Class", "Project", "Subclass", "Stage", "Region", "City", "District"]:
        data[col] = data[col].fillna("Not Classified")

    # penalties
    if penalties.empty:
        penalties = pd.DataFrame(columns=["Link Code", "Deviation name", "Number of Deviations", "Penalties Amount", "Region", "City"])
    else:
        p_link = first_existing(penalties, ["Cluster Name", "Link Code"])
        p_name = first_existing(penalties, ["Deviation name", "Penalty"])
        p_qty = first_existing(penalties, ["Number of Deviations", "Number", "Qty"])
        p_amt = first_existing(penalties, ["Penalties Amount", "Amount"])
        if p_link:
            penalties["Link Code"] = penalties[p_link].astype(str).str.strip()
        else:
            penalties["Link Code"] = np.nan
        penalties["Number of Deviations"] = safe_numeric(penalties[p_qty] if p_qty else 0)
        penalties["Penalties Amount"] = safe_numeric(penalties[p_amt] if p_amt else 0)
        if "Region" not in penalties.columns:
            penalties["Region"] = np.nan
        pen_map = data[["Link Code","City","District","Region"]].drop_duplicates()
        penalties = penalties.merge(pen_map, on="Link Code", how="left", suffixes=("", "_main"))
        # Safe fill for penalties dimensions even when original columns do not exist
        if "City" not in penalties.columns:
            penalties["City"] = pd.Series([pd.NA] * len(penalties), index=penalties.index, dtype="object")
        if "Region" not in penalties.columns:
            penalties["Region"] = pd.Series([pd.NA] * len(penalties), index=penalties.index, dtype="object")
        if "City_main" in penalties.columns:
            penalties["City"] = penalties["City"].combine_first(penalties["City_main"])
        if "Region_main" in penalties.columns:
            penalties["Region"] = penalties["Region"].combine_first(penalties["Region_main"])
        penalties["City"] = penalties["City"].fillna("Not Classified")
        penalties["Region"] = penalties["Region"].fillna("Not Classified")
        if p_name and p_name != "Deviation name":
            penalties.rename(columns={p_name: "Deviation name"}, inplace=True)

    pen_agg = penalties.groupby("Link Code", dropna=False).agg(
        penalty_cases=("Deviation name", "count") if "Deviation name" in penalties.columns else ("Number of Deviations", "size"),
        penalty_qty=("Number of Deviations", "sum"),
        penalty_amount=("Penalties Amount", "sum"),
    ).reset_index()
    data = data.merge(pen_agg, on="Link Code", how="left")
    for c in ["penalty_cases","penalty_qty","penalty_amount"]:
        data[c] = safe_numeric(data[c])

    # workorder details sheet
    if details.empty:
        details = data.copy()
    else:
        d_link = first_existing(details, ["Link Code"])
        if d_link and d_link != "Link Code":
            details.rename(columns={d_link: "Link Code"}, inplace=True)
        details["Link Code"] = details["Link Code"].astype(str).str.strip()
        # enrich with main columns not present in details
        enrich_cols = [c for c in ["Work Order", "Updates", "Updated", "effective_target", "City", "District", "Region", "Project", "Subclass", "Stage", "Type", "Class", "Work Order Status"] if c in data.columns]
        details = details.merge(data[["Link Code"] + [c for c in enrich_cols if c != "Link Code"]].drop_duplicates("Link Code"), on="Link Code", how="left")
        for col in details.columns:
            if details[col].dtype == "object":
                details[col] = details[col].map(clean_text)
        for col in ["Created","Assigned at","Updated","Targeted Completion","Updated Target Date","Closed at"]:
            if col in details.columns:
                details[col] = pd.to_datetime(details[col], errors="coerce")

    warnings = []
    for c in ["Link Code","Work Order","Percentage of Completion","Updated","Updates","Targeted Completion","Cost","Work Order Status"]:
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
            Executive PMO dashboard for top management decision support. Snapshot date: <b>{pd.to_datetime(snapshot_date).strftime('%d %b %Y %H:%M')}</b>
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
if region != "All": base = base[base["Region"].astype(str) == region]
city = st.sidebar.selectbox("City", options_for(base, "City"), index=0)
if city != "All": base = base[base["City"].astype(str) == city]
district = st.sidebar.selectbox("District", options_for(base, "District"), index=0)
if district != "All": base = base[base["District"].astype(str) == district]
project = st.sidebar.selectbox("Project", options_for(base, "Project"), index=0)
if project != "All": base = base[base["Project"].astype(str) == project]
stage = st.sidebar.selectbox("Stage", options_for(base, "Stage"), index=0)
if stage != "All": base = base[base["Stage"].astype(str) == stage]
year = st.sidebar.selectbox("Year", options_for(base.assign(Year=base["Year"].astype(str)), "Year"), index=0)
if year != "All": base = base[base["Year"].astype(str) == year]
status = st.sidebar.selectbox("Work Order Status", options_for(base, "Work Order Status"), index=0)
if status != "All": base = base[base["Work Order Status"].astype(str) == status]
wo_type = st.sidebar.selectbox("Type", options_for(base, "Type"), index=0)
if wo_type != "All": base = base[base["Type"].astype(str) == wo_type]
wo_class = st.sidebar.selectbox("Class", options_for(base, "Class"), index=0)
if wo_class != "All": base = base[base["Class"].astype(str) == wo_class]
subclass = st.sidebar.selectbox("Subclass", options_for(base, "Subclass"), index=0)
if subclass != "All": base = base[base["Subclass"].astype(str) == subclass]
link_code = st.sidebar.selectbox("Link Code", options_for(base, "Link Code"), index=0)
filtered = base.copy()
if link_code != "All": filtered = filtered[filtered["Link Code"].astype(str) == link_code]

if filtered.empty:
    st.warning("No records match the selected filters.")
    st.stop()

# filtered details and penalties
filtered_links = filtered["Link Code"].dropna().astype(str).unique().tolist()
filtered_details = details[details["Link Code"].astype(str).isin(filtered_links)].copy() if "Link Code" in details.columns else pd.DataFrame()
filtered_penalties = penalties[penalties["Link Code"].astype(str).isin(filtered_links)].copy() if "Link Code" in penalties.columns else pd.DataFrame()

# KPI calculations
total_wo = filtered["Work Order"].nunique() if "Work Order" in filtered.columns else len(filtered)
total_link = filtered["Link Code"].nunique()
avg_actual = filtered["actual_progress_capped"].mean()
avg_planned = filtered["planned_progress_pct"].mean()
avg_lag_pp = filtered["lag_pp"].mean()
lagged_pct = (filtered["lag_pp"] > 0).mean() * 100
overdue_cnt = int(filtered["is_overdue"].sum())
critical_lag_cnt = int(filtered["critical_lag"].sum())
forecast_high_risk = int((filtered["forecast_risk"] == "High delay risk").sum())
penalty_cases = int(filtered["penalty_cases"].sum())
penalty_qty = float(filtered["penalty_qty"].sum())
penalty_amount = float(filtered["penalty_amount"].sum())
avg_civil = filtered["civil_completion_pct"].mean()
avg_fiber = filtered["fiber_completion_pct"].mean()
rejection_cols = [c for c in ["PIP rejection count","PAT rejection count","Approval rejection count","As-Built Rejection Count","Handover Rejection Count"] if c in filtered.columns]
rejections_total = float(filtered[rejection_cols].fillna(0).sum().sum()) if rejection_cols else 0
update_needed = filtered[filtered["needs_system_update"]].copy()
in_progress_cnt = int(filtered["Work Order Status"].astype(str).str.contains("In Progress", case=False, na=False).sum())
cancelled_cnt = int(filtered["Work Order Status"].astype(str).str.contains("Cancel", case=False, na=False).sum())

snapshot_month = pd.Timestamp(snapshot_date).to_period("M")
current_month_target_cost = safe_numeric(filtered.loc[pd.to_datetime(filtered.get("Targeted Completion"), errors="coerce").dt.to_period("M") == snapshot_month, "Cost" if "Cost" in filtered.columns else "WO Cost"], 0).sum()
current_month_updated_target_cost = safe_numeric(filtered.loc[pd.to_datetime(filtered.get("Updated Target Date"), errors="coerce").dt.to_period("M") == snapshot_month, "Cost" if "Cost" in filtered.columns else "WO Cost"], 0).sum()

# charts helpers
def chart_layout(fig):
    fig.update_layout(margin=dict(l=20,r=20,t=40,b=20), legend_title_text="", height=360)
    return fig


def monthly_progress_chart(df):
    if "effective_target" not in df.columns or df["effective_target"].dropna().empty:
        return None
    temp = df.copy()
    temp["target_month"] = pd.to_datetime(temp["effective_target"], errors="coerce").dt.to_period("M").astype(str)
    grp = temp.groupby("target_month", dropna=False).agg(
        Planned=("planned_progress_pct", "mean"),
        Actual=("actual_progress_capped", "mean"),
    ).reset_index().sort_values("target_month")
    if grp.empty:
        return None
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=grp["target_month"], y=grp["Planned"], mode="lines+markers", name="Planned"))
    fig.add_trace(go.Scatter(x=grp["target_month"], y=grp["Actual"], mode="lines+markers", name="Actual"))
    fig.update_yaxes(range=[0, 110], title="Progress %")
    fig.update_xaxes(title="Target Month")
    fig.update_layout(title="Planned vs Actual Progress Trend")
    return chart_layout(fig)


def city_status_chart(df):
    if df.empty:
        return None
    grp = df.groupby(["City"], dropna=False).agg(
        Avg_Actual=("actual_progress_capped", "mean"),
        Avg_Planned=("planned_progress_pct", "mean"),
    ).reset_index().sort_values("Avg_Actual", ascending=False)
    fig = go.Figure()
    fig.add_trace(go.Bar(x=grp["City"], y=grp["Avg_Planned"], name="Planned"))
    fig.add_trace(go.Bar(x=grp["City"], y=grp["Avg_Actual"], name="Actual"))
    fig.update_layout(barmode="group", title="City Performance")
    fig.update_yaxes(title="Progress %")
    return chart_layout(fig)


def risk_chart(df):
    if df.empty:
        return None
    grp = df["forecast_risk"].value_counts().rename_axis("Risk").reset_index(name="Count")
    fig = px.bar(grp, x="Risk", y="Count", title="Forecast Risk Exposure")
    return chart_layout(fig)


def penalties_chart(df):
    if df.empty or "Deviation name" not in df.columns:
        return None
    grp = df.groupby("Deviation name", dropna=False)["Number of Deviations"].sum().reset_index().sort_values("Number of Deviations", ascending=False).head(10)
    if grp.empty:
        return None
    fig = px.bar(grp, x="Number of Deviations", y="Deviation name", orientation="h", title="Top Penalties / Deviations")
    fig.update_layout(yaxis=dict(categoryorder="total ascending"))
    return chart_layout(fig)


def milestone_chart(df):
    rows = []
    for label, col in [("Civil","civil_completion_pct"),("MH/HH","mhhh_completion_pct"),("Fiber","fiber_completion_pct"),("Permits","permits_completion_pct")]:
        if col in df.columns:
            rows.append({"Milestone": label, "Completion": df[col].mean()})
    grp = pd.DataFrame(rows)
    if grp.empty:
        return None
    fig = px.bar(grp, x="Milestone", y="Completion", title="Milestone Completion Summary")
    fig.update_yaxes(range=[0, 110], title="Completion %")
    return chart_layout(fig)


tabs = st.tabs(["Executive Overview", "PMO Summary", "Schedule & KPI", "Penalties & Quality", "Work Order Detail", "Dashboard Guide"])

with tabs[0]:
    r1 = st.columns(6)
    metrics = [
        ("Work Orders", f"{total_wo:,}", f"{total_link:,} link codes"),
        ("Avg Actual Progress", fmt_pct(avg_actual), "Overall completion"),
        ("Avg Planned Progress", fmt_pct(avg_planned), "Based on effective target dates"),
        ("Civil Completion", fmt_pct(avg_civil), "Trench scope delivered"),
        ("Fiber Completion", fmt_pct(avg_fiber), "Fiber scope delivered"),
        ("Penalty Cases", f"{penalty_cases:,}", f"Qty {penalty_qty:,.0f}"),
    ]
    for c, (t, v, s) in zip(r1, metrics):
        with c: add_card(t, v, s)

    r2 = st.columns(5)
    metrics2 = [
        ("Avg Lag", fmt_pct(avg_lag_pp), "Planned minus actual"),
        ("Lagged WO", fmt_pct(lagged_pct), "Share behind plan"),
        ("Overdue WO", f"{overdue_cnt:,}", "Past target date"),
        ("Critical Lag >15pp", f"{critical_lag_cnt:,}", "Immediate action"),
        ("High Forecast Risk", f"{forecast_high_risk:,}", "Projected major delay"),
    ]
    for c, (t, v, s) in zip(r2, metrics2):
        with c: add_card(t, v, s)

    c1, c2 = st.columns(2)
    with c1:
        fig = monthly_progress_chart(filtered)
        if fig is not None:
            st.plotly_chart(fig, use_container_width=True)
    with c2:
        fig = city_status_chart(filtered)
        if fig is not None:
            st.plotly_chart(fig, use_container_width=True)

    c3, c4 = st.columns(2)
    with c3:
        fig = milestone_chart(filtered)
        if fig is not None:
            st.plotly_chart(fig, use_container_width=True)
    with c4:
        fig = risk_chart(filtered)
        if fig is not None:
            st.plotly_chart(fig, use_container_width=True)

with tabs[1]:
    st.markdown('<div class="guide-box"><b>PMO follow-up summary</b><br>Records below highlight work orders that should be reviewed for system update follow-up and current month commercial exposure.</div>', unsafe_allow_html=True)

    s1 = st.columns(5)
    pmo_metrics = [
        ("Need System Update", f"{len(update_needed):,}", "Updates < 5 and last update > 5 days"),
        ("Current Month Target Cost", fmt_money(current_month_target_cost), "Based on Targeted Completion"),
        ("Current Month Updated Target Cost", fmt_money(current_month_updated_target_cost), "Based on Updated Target Date"),
        ("In Progress WO", f"{in_progress_cnt:,}", "From Work Order Status"),
        ("Cancelled WO", f"{cancelled_cnt:,}", "From Work Order Status"),
    ]
    for c, (t, v, s) in zip(s1, pmo_metrics):
        with c: add_card(t, v, s)

    pmo_cols = [c for c in ["Link Code","Work Order","City","District","Work Order Status","Updated","Updates","days_since_update","Targeted Completion","Updated Target Date","Cost","lag_pp","forecast_delay_days","Stage","Subclass","Notes"] if c in update_needed.columns]
    pmo_table = update_needed[pmo_cols].sort_values(["days_since_update","Updates"], ascending=[False, True]) if not update_needed.empty else pd.DataFrame(columns=pmo_cols)
    st.subheader("Link codes requiring system update review")
    st.dataframe(pmo_table, use_container_width=True, height=420)
    if not pmo_table.empty:
        st.download_button(
            "Download PMO summary (Excel)",
            data=to_excel_bytes(pmo_table),
            file_name="dawiyat_pmo_summary.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

with tabs[2]:
    c1, c2, c3, c4 = st.columns(4)
    with c1: add_card("Penalty Amount", fmt_money(penalty_amount), "Consultant deductions")
    with c2: add_card("Penalty Quantity", f"{penalty_qty:,.0f}", "Deviation count basis")
    with c3: add_card("Rejections Total", f"{rejections_total:,.0f}", "PIP / PAT / Handover")
    with c4: add_card("Permits Completion", fmt_pct(filtered["permits_completion_pct"].mean()), "Permit scope delivered")

    c5, c6 = st.columns(2)
    with c5:
        fig = monthly_progress_chart(filtered)
        if fig is not None:
            st.plotly_chart(fig, use_container_width=True)
    with c6:
        # lag distribution chart
        temp = filtered.copy()
        temp["Lag Band"] = np.select(
            [temp["lag_pp"] <= 0, temp["lag_pp"].between(0, 15, inclusive="right"), temp["lag_pp"] > 15],
            ["On / Ahead", "Minor Lag", "Critical Lag"],
            default="No Data",
        )
        grp = temp["Lag Band"].value_counts().rename_axis("Lag Band").reset_index(name="Count")
        fig = px.bar(grp, x="Lag Band", y="Count", title="Lag Distribution")
        st.plotly_chart(chart_layout(fig), use_container_width=True)

with tabs[3]:
    c1, c2 = st.columns(2)
    with c1:
        fig = penalties_chart(filtered_penalties)
        if fig is not None:
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("No penalties available for the selected filters.")
    with c2:
        dq_missing_city = int(filtered["City"].isna().sum())
        dq_missing_dist = int(filtered["District"].isna().sum())
        dq_missing_link = int(filtered["Link Code"].isna().sum())
        total_records = max(len(filtered), 1)
        dq_score = 100 - ((dq_missing_city + dq_missing_dist + dq_missing_link) / (3 * total_records) * 100)
        add_card("Data Quality Score", fmt_pct(dq_score), "City / District / Link Code completeness")
        st.markdown('<div class="guide-box">City and District are mapped first from the District sheet, then from Link Code pattern fallback to reduce unknown values.</div>', unsafe_allow_html=True)
        qcols = st.columns(3)
        with qcols[0]: add_card("Missing City", f"{dq_missing_city:,}")
        with qcols[1]: add_card("Missing District", f"{dq_missing_dist:,}")
        with qcols[2]: add_card("Missing Link Code", f"{dq_missing_link:,}")

    st.subheader("Penalty detail")
    penalty_cols = [c for c in ["Link Code","City","Region","Assigned Date","Inspection Consultant","Implementation Contractor","Deviation name","Number of Deviations","Penalties Amount"] if c in filtered_penalties.columns]
    st.dataframe(filtered_penalties[penalty_cols], use_container_width=True, height=360)

with tabs[4]:
    st.markdown('<div class="guide-box"><b>Work Order Detail</b><br>The table below uses the Workorder details sheet and is enriched with City, District, Work Order, Updates, Updated date, and current project dimensions from the master sheet. You can export the filtered result.</div>', unsafe_allow_html=True)
    preferred = [
        "Link Code","Work Order","City","District","Region","Project","Subclass","Stage","Type","Class","Work Order Status",
        "Percentage of Completion","Trench Progress","Trench Scope","MH/HH Progress","MH/HH Scope","Fiber Progress","Fiber Scope",
        "ODBs Progress","ODBs Scope","ODFs Progress","ODFs Scope","JCL Progress","JCL Scope","FAT Progress","FAT Scope",
        "PFAT Progress","PFAT Scope","SFAT Progress","SFAT Scope","WO Cost","Cost","Created","Assigned at","Updated",
        "Updates","Targeted Completion","Updated Target Date","effective_target","Closed at","Permit Review","Permit Completion Status",
        "Permits Progress","Permits Scope","PIP rejection count","PAT rejection count","Approval rejection count","As-Built Rejection Count",
        "Handover Rejection Count","Parent","Number","Request","SOR Reference Number","Target Area","Scope of Work","Notes"
    ]
    show_cols = [c for c in preferred if c in filtered_details.columns]
    detail_view = filtered_details[show_cols].copy() if show_cols else filtered_details.copy()
    st.dataframe(detail_view, use_container_width=True, height=520)
    d1, d2 = st.columns(2)
    with d1:
        st.download_button(
            "Download Workorder Detail (Excel)",
            data=to_excel_bytes(detail_view),
            file_name="dawiyat_workorder_detail.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    with d2:
        st.download_button(
            "Download Workorder Detail (CSV)",
            data=detail_view.to_csv(index=False).encode("utf-8-sig"),
            file_name="dawiyat_workorder_detail.csv",
            mime="text/csv",
        )

with tabs[5]:
    st.markdown(
        """
        <div class="guide-box"><b>Dashboard guide</b><br>
        <span class="small-chip">Actual Progress</span>
        Actual progress = <b>Percentage of Completion</b> from the main service tool sheet, capped at 100% for KPI comparison.<br><br>
        <span class="small-chip">Planned Progress</span>
        Planned progress = elapsed duration from start date to snapshot date divided by total duration from start date to <b>effective target</b>.<br><br>
        <span class="small-chip">Effective Target</span>
        Effective target = <b>Updated Target Date</b> if available, otherwise <b>Targeted Completion</b>.<br><br>
        <span class="small-chip">Lag %</span>
        Lag = planned progress minus actual progress. Positive value means the work order is behind plan.<br><br>
        <span class="small-chip">Critical Lag</span>
        Critical lag means lag is more than or equal to 15 percentage points.<br><br>
        <span class="small-chip">Civil Completion</span>
        Civil completion = Trench Progress / Trench Scope.<br><br>
        <span class="small-chip">Fiber Completion</span>
        Fiber completion = Fiber Progress / Fiber Scope.<br><br>
        <span class="small-chip">PMO Update Rule</span>
        A link code is flagged for PMO follow-up when <b>Updates &lt; 5</b> and <b>last Updated date exceeds 5 days</b> versus snapshot date.<br><br>
        <span class="small-chip">Current Month Target Cost</span>
        Sum of <b>Cost</b> for work orders whose <b>Targeted Completion</b> falls in the snapshot month.<br><br>
        <span class="small-chip">Current Month Updated Target Cost</span>
        Sum of <b>Cost</b> for work orders whose <b>Updated Target Date</b> falls in the snapshot month.<br><br>
        <span class="small-chip">Penalties</span>
        Penalties are aggregated by Link Code using the penalties sheet. Cases = count of deviation lines. Qty = sum of Number of Deviations. Amount = sum of Penalties Amount.
        </div>
        """,
        unsafe_allow_html=True,
    )
