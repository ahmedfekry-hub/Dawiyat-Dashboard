import io
import os
from datetime import datetime

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# =========================================================
# PAGE CONFIG
# =========================================================
st.set_page_config(
    page_title="Dawiyat Executive Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

DATA_FILE = "Dawiyat Master Sheet.xlsx"

# =========================================================
# STYLE
# =========================================================
st.markdown(
    """
    <style>
    :root {
        --bg-main: #0a1630;
        --bg-card: #12274a;
        --text-main: #ffffff;
        --text-soft: #cfe2ff;
    }

    .stApp {
        background: linear-gradient(180deg, #edf2f7 0%, #e9eef6 100%);
    }

    section[data-testid="stSidebar"] {
        background: linear-gradient(180deg, #071a3d 0%, #04122c 100%);
        border-right: 1px solid rgba(255,255,255,0.06);
    }

    section[data-testid="stSidebar"] * {
        color: #ffffff !important;
    }

    section[data-testid="stSidebar"] label,
    section[data-testid="stSidebar"] p,
    section[data-testid="stSidebar"] span,
    section[data-testid="stSidebar"] div {
        color: #ffffff !important;
        font-weight: 700 !important;
        font-size: 15px !important;
    }

    section[data-testid="stSidebar"] .stSelectbox label {
        color: #ffffff !important;
        font-size: 16px !important;
        font-weight: 800 !important;
    }

    div[data-baseweb="select"] > div {
        background: #ffffff !important;
        border-radius: 12px !important;
        min-height: 46px !important;
        border: 1px solid #dce6f5 !important;
        color: #111827 !important;
        font-weight: 800 !important;
    }

    div[data-baseweb="select"] span {
        color: #111827 !important;
        font-weight: 800 !important;
        opacity: 1 !important;
    }

    div[data-baseweb="select"] input {
        color: #111827 !important;
        font-weight: 800 !important;
    }

    .sidebar-title {
        color: #ffffff;
        font-size: 20px;
        font-weight: 900;
        margin-bottom: 14px;
    }

    .top-title {
        font-size: 40px;
        font-weight: 900;
        color: #1f2a44;
        margin-bottom: 0.1rem;
    }

    .top-subtitle {
        color: #5f6b82;
        font-size: 15px;
        margin-bottom: 1rem;
    }

    .card {
        background: linear-gradient(180deg, #132a4e 0%, #102341 100%);
        border-radius: 20px;
        padding: 18px 20px;
        min-height: 132px;
        box-shadow: 0 10px 24px rgba(7, 22, 44, 0.18);
        border: 1px solid rgba(255,255,255,0.06);
    }

    .card-title {
        color: #bcd7ff;
        font-size: 15px;
        font-weight: 700;
        margin-bottom: 10px;
    }

    .card-value {
        color: #ffffff;
        font-size: 24px;
        font-weight: 900;
        line-height: 1.15;
    }

    .card-note {
        color: #99b4de;
        font-size: 13px;
        margin-top: 10px;
    }

    .section-title {
        font-size: 28px;
        font-weight: 900;
        color: #1f2a44;
        margin-top: 10px;
        margin-bottom: 8px;
    }

    .guide-box {
        background: #ffffff;
        padding: 18px;
        border-radius: 16px;
        border: 1px solid #dde6f3;
        box-shadow: 0 4px 16px rgba(12, 25, 49, 0.06);
    }

    .footer-note {
        color: #6b7280;
        font-size: 12px;
        margin-top: 8px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# =========================================================
# HELPERS
# =========================================================
def clean_col_name(col):
    return str(col).strip()

def normalize_text(x):
    if pd.isna(x):
        return np.nan
    s = str(x).strip()
    if s == "" or s.lower() in ["none", "nan", "null"]:
        return np.nan
    return s

def to_datetime_series(s):
    return pd.to_datetime(s, errors="coerce")

def to_numeric_series(s):
    return pd.to_numeric(s, errors="coerce")

def pick_first_valid(*vals):
    for v in vals:
        if pd.notna(v):
            s = str(v).strip()
            if s and s.lower() not in ["none", "nan", "null"]:
                return s
    return np.nan

def fmt_int(x):
    if pd.isna(x):
        return "0"
    return str(int(round(float(x))))

def fmt_money(x):
    if pd.isna(x):
        return "SAR 0"
    return f"SAR {x:,.0f}"

def fmt_pct(x):
    if pd.isna(x):
        return "0.0%"
    return f"{x:.1f}%"

def current_ts():
    return datetime.now().strftime("%Y-%m-%d %H:%M")

def safe_download_excel(df, sheet_name):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return output.getvalue()

def empty_message(msg="No data available for selected filters"):
    st.info(msg)

def safe_month_label(period_value):
    try:
        return pd.Period(period_value, freq="M").strftime("%b%Y")
    except Exception:
        try:
            return pd.to_datetime(period_value).strftime("%b%Y")
        except Exception:
            return str(period_value)

def normalize_region(r):
    if pd.isna(r):
        return np.nan
    t = str(r).strip().lower()
    if "western" in t:
        return "Western"
    if "southern" in t:
        return "Southern"
    if "eastern" in t:
        return "Eastern"
    if "central" in t:
        return "Central"
    return str(r).strip()

def safe_group_sum(df, group_col, value_col):
    if df.empty or group_col not in df.columns or value_col not in df.columns:
        return pd.DataFrame(columns=[group_col, value_col])
    x = df[[group_col, value_col]].copy()
    x[value_col] = to_numeric_series(x[value_col]).fillna(0)
    return x.groupby(group_col, dropna=False)[value_col].sum().reset_index()

def build_monthly_cost(df):
    if df.empty or "Cost" not in df.columns:
        return pd.DataFrame(columns=["Month", "Targeted Completion Cost", "Updated Target Date Cost"])

    tmp = df.copy()
    tmp["Targeted Completion"] = to_datetime_series(tmp.get("Targeted Completion"))
    tmp["Updated Target Date"] = to_datetime_series(tmp.get("Updated Target Date"))
    tmp["Cost"] = to_numeric_series(tmp.get("Cost")).fillna(0)

    targeted = tmp.dropna(subset=["Targeted Completion"]).copy()
    updated = tmp.dropna(subset=["Updated Target Date"]).copy()

    if not targeted.empty:
        targeted["Month"] = targeted["Targeted Completion"].dt.to_period("M").astype(str)
        targeted_monthly = targeted.groupby("Month", dropna=False)["Cost"].sum().reset_index()
        targeted_monthly.rename(columns={"Cost": "Targeted Completion Cost"}, inplace=True)
    else:
        targeted_monthly = pd.DataFrame(columns=["Month", "Targeted Completion Cost"])

    if not updated.empty:
        updated["Month"] = updated["Updated Target Date"].dt.to_period("M").astype(str)
        updated_monthly = updated.groupby("Month", dropna=False)["Cost"].sum().reset_index()
        updated_monthly.rename(columns={"Cost": "Updated Target Date Cost"}, inplace=True)
    else:
        updated_monthly = pd.DataFrame(columns=["Month", "Updated Target Date Cost"])

    monthly_cost = pd.merge(targeted_monthly, updated_monthly, on="Month", how="outer")

    if monthly_cost.empty:
        return pd.DataFrame(columns=["Month", "Targeted Completion Cost", "Updated Target Date Cost"])

    monthly_cost["Targeted Completion Cost"] = to_numeric_series(monthly_cost["Targeted Completion Cost"]).fillna(0)
    monthly_cost["Updated Target Date Cost"] = to_numeric_series(monthly_cost["Updated Target Date Cost"]).fillna(0)

    monthly_cost["MonthSort"] = pd.PeriodIndex(monthly_cost["Month"], freq="M")
    monthly_cost = monthly_cost.sort_values("MonthSort")
    monthly_cost["MonthLabel"] = monthly_cost["MonthSort"].astype(str).map(lambda x: pd.Period(x, freq="M").strftime("%b%Y"))
    monthly_cost = monthly_cost[["MonthLabel", "Targeted Completion Cost", "Updated Target Date Cost"]]
    monthly_cost.rename(columns={"MonthLabel": "Month"}, inplace=True)
    return monthly_cost

# =========================================================
# LOAD DATA
# =========================================================
@st.cache_data(show_spinner=False)
def load_data():
    if not os.path.exists(DATA_FILE):
        raise FileNotFoundError(
            f"'{DATA_FILE}' not found. Keep it in the same GitHub folder as app.py"
        )

    service = pd.read_excel(DATA_FILE, sheet_name="Dawaiyat Service Tool")
    district = pd.read_excel(DATA_FILE, sheet_name="District ")
    penalties = pd.read_excel(DATA_FILE, sheet_name="Penalties")
    details = pd.read_excel(DATA_FILE, sheet_name="Workorder details")

    service.columns = [clean_col_name(c) for c in service.columns]
    district.columns = [clean_col_name(c) for c in district.columns]
    penalties.columns = [clean_col_name(c) for c in penalties.columns]
    details.columns = [clean_col_name(c) for c in details.columns]

    # Clean service
    data = service.copy()
    for c in data.columns:
        if data[c].dtype == object:
            data[c] = data[c].apply(normalize_text)

    numeric_cols = [
        "Percentage of Completion", "WO Cost", "Cost", "Updated", "Updates",
        "Trench Progress", "Trench Scope", "Fiber Progress", "Fiber Scope",
        "PIP rejection count", "PAT rejection count", "Approval rejection count",
        "As-Built Rejection Count", "Handover Rejection Count"
    ]
    for c in numeric_cols:
        if c in data.columns:
            data[c] = to_numeric_series(data[c])

    date_cols = ["Targeted Completion", "Updated Target Date", "Created", "Assigned at", "In Progress at", "Closed at"]
    for c in date_cols:
        if c in data.columns:
            data[c] = to_datetime_series(data[c])

    # Clean district sheet
    district_map = district.copy()
    for c in district_map.columns:
        if district_map[c].dtype == object:
            district_map[c] = district_map[c].apply(normalize_text)

    link_map = {}
    if "Link Code" in district_map.columns:
        for _, r in district_map.iterrows():
            lk = r.get("Link Code")
            if pd.notna(lk):
                link_map[str(lk).strip()] = {
                    "Region": r.get("Region"),
                    "City": r.get("City"),
                    "District": r.get("District"),
                }

    wo_map = {}
    if "Work Order" in district_map.columns:
        for _, r in district_map.iterrows():
            wo = r.get("Work Order")
            if pd.notna(wo):
                wo_map[str(wo).strip()] = {
                    "Region": r.get("Region"),
                    "City": r.get("City"),
                    "District": r.get("District"),
                }

    def enrich_main_row(row):
        lk = str(row.get("Link Code")).strip() if pd.notna(row.get("Link Code")) else None
        wo = str(row.get("Work Order")).strip() if pd.notna(row.get("Work Order")) else None
        a = link_map.get(lk, {})
        b = wo_map.get(wo, {})

        region = pick_first_valid(a.get("Region"), b.get("Region"), row.get("Region"))
        city = pick_first_valid(a.get("City"), b.get("City"))
        district_v = pick_first_valid(a.get("District"), b.get("District"))
        if pd.isna(district_v):
            district_v = "N/A"

        return pd.Series([normalize_region(region), city, district_v])

    data[["Region_final", "City_final", "District_final"]] = data.apply(enrich_main_row, axis=1)

    # Effective dates and metrics
    data["Effective Target Date"] = data["Updated Target Date"].combine_first(data["Targeted Completion"])

    today = pd.Timestamp.today().normalize()
    data["lag_days"] = np.where(
        data["Effective Target Date"].notna(),
        (today - data["Effective Target Date"]).dt.days,
        np.nan,
    )
    data["lag_days"] = np.where(pd.isna(data["lag_days"]), np.nan, np.where(data["lag_days"] < 0, 0, data["lag_days"]))

    completion = data["Percentage of Completion"].fillna(0)
    data["lag_pct"] = np.where(
        completion < 100,
        np.where(pd.Series(data["lag_days"]).fillna(0) > 0, 100 - completion, 0),
        0,
    )

    data["forecast_risk"] = np.select(
        [
            completion >= 100,
            pd.Series(data["lag_days"]).fillna(0) == 0,
            (pd.Series(data["lag_days"]).fillna(0) > 0) & (pd.Series(data["lag_days"]).fillna(0) <= 15),
            pd.Series(data["lag_days"]).fillna(0) > 15,
        ],
        [
            "Completed",
            "On forecast",
            "Moderate delay risk",
            "High delay risk",
        ],
        default="On forecast",
    )

    data["civil_completion_pct"] = np.where(
        data.get("Trench Scope", pd.Series([0]*len(data))).fillna(0) > 0,
        (data.get("Trench Progress", pd.Series([0]*len(data))).fillna(0) / data.get("Trench Scope", pd.Series([0]*len(data))).fillna(0)) * 100,
        0,
    )
    data["fiber_completion_pct"] = np.where(
        data.get("Fiber Scope", pd.Series([0]*len(data))).fillna(0) > 0,
        (data.get("Fiber Progress", pd.Series([0]*len(data))).fillna(0) / data.get("Fiber Scope", pd.Series([0]*len(data))).fillna(0)) * 100,
        0,
    )

    reject_cols = [c for c in [
        "PIP rejection count", "PAT rejection count", "Approval rejection count",
        "As-Built Rejection Count", "Handover Rejection Count"
    ] if c in data.columns]
    if reject_cols:
        data["total_rejections"] = data[reject_cols].fillna(0).sum(axis=1)
    else:
        data["total_rejections"] = 0

    # Clean penalties
    p = penalties.copy()
    for c in p.columns:
        if p[c].dtype == object:
            p[c] = p[c].apply(normalize_text)
    if "Number of Deviations" in p.columns:
        p["Number of Deviations"] = to_numeric_series(p["Number of Deviations"]).fillna(0)
    else:
        p["Number of Deviations"] = 0
    if "Penalties Amount" in p.columns:
        p["Penalties Amount"] = to_numeric_series(p["Penalties Amount"]).fillna(0)
    else:
        p["Penalties Amount"] = 0
    if "Region" in p.columns:
        p["Region"] = p["Region"].apply(normalize_region)

    # Clean details
    details_df = details.copy()
    for c in details_df.columns:
        if details_df[c].dtype == object:
            details_df[c] = details_df[c].apply(normalize_text)

    # Map details using both link code and work order
    def enrich_detail_row(row):
        lk = str(row.get("Link Code")).strip() if pd.notna(row.get("Link Code")) else None
        wo = str(row.get("Work Order")).strip() if pd.notna(row.get("Work Order")) else None
        a = link_map.get(lk, {})
        b = wo_map.get(wo, {})
        region = pick_first_valid(a.get("Region"), b.get("Region"))
        city = pick_first_valid(a.get("City"), b.get("City"))
        district_v = pick_first_valid(a.get("District"), b.get("District"))
        if pd.isna(district_v):
            district_v = "N/A"
        return pd.Series([normalize_region(region), city, district_v])

    if not details_df.empty:
        details_df[["Region_final", "City_final", "District_final"]] = details_df.apply(enrich_detail_row, axis=1)
    else:
        details_df["Region_final"] = []
        details_df["City_final"] = []
        details_df["District_final"] = []

    for c in ["Targeted Completion", "Updated Target Date", "Created", "Assigned at", "Closed at"]:
        if c in details_df.columns:
            details_df[c] = to_datetime_series(details_df[c])

    if "Link Code" in details_df.columns:
        details_df = details_df.merge(
            data[["Link Code", "forecast_risk", "lag_pct"]].drop_duplicates(subset=["Link Code"]),
            on="Link Code",
            how="left"
        )

    return data, p, details_df

# =========================================================
# LOAD
# =========================================================
try:
    data, penalties, details_df = load_data()
except Exception as e:
    st.error(str(e))
    st.stop()

# =========================================================
# SIDEBAR FILTERS
# =========================================================
st.sidebar.markdown('<div class="sidebar-title">Control Panel</div>', unsafe_allow_html=True)
st.sidebar.markdown(
    f"""
    <div style="color:#bcd7ff; font-size:13px; margin-bottom:16px;">
    Data source: <b>{DATA_FILE}</b><br>
    Last app refresh: <b>{current_ts()}</b>
    </div>
    """,
    unsafe_allow_html=True,
)

filter_df = data.copy()

region_options = ["All"] + sorted([x for x in filter_df["Region_final"].dropna().unique().tolist()])
selected_region = st.sidebar.selectbox("Region", region_options)

if selected_region != "All":
    filter_df = filter_df[filter_df["Region_final"] == selected_region]

city_options = ["All"] + sorted([x for x in filter_df["City_final"].dropna().unique().tolist()])
selected_city = st.sidebar.selectbox("City", city_options)

if selected_city != "All":
    filter_df = filter_df[filter_df["City_final"] == selected_city]

district_options = ["All"] + sorted([x for x in filter_df["District_final"].dropna().unique().tolist()])
selected_district = st.sidebar.selectbox("District", district_options)

if selected_district != "All":
    filter_df = filter_df[filter_df["District_final"] == selected_district]

project_options = ["All"]
if "Project" in filter_df.columns:
    project_options += sorted([x for x in filter_df["Project"].dropna().unique().tolist()])
selected_project = st.sidebar.selectbox("Project", project_options)

if selected_project != "All":
    filter_df = filter_df[filter_df["Project"] == selected_project]

subclass_options = ["All"]
if "Subclass" in filter_df.columns:
    subclass_options += sorted([x for x in filter_df["Subclass"].dropna().unique().tolist()])
selected_subclass = st.sidebar.selectbox("Subclass", subclass_options)

if selected_subclass != "All":
    filter_df = filter_df[filter_df["Subclass"] == selected_subclass]

wo_status_options = ["All"]
if "Work Order Status" in filter_df.columns:
    wo_status_options += sorted([x for x in filter_df["Work Order Status"].dropna().unique().tolist()])
selected_status = st.sidebar.selectbox("WorkOrderStatus", wo_status_options)

if selected_status != "All":
    filter_df = filter_df[filter_df["Work Order Status"] == selected_status]

link_code_options = ["All"] + sorted([x for x in filter_df["Link Code"].dropna().unique().tolist()])
selected_link = st.sidebar.selectbox("Link Code", link_code_options)

if selected_link != "All":
    filter_df = filter_df[filter_df["Link Code"] == selected_link]

filtered_links = filter_df["Link Code"].dropna().unique().tolist()

filtered_details = details_df.copy()
if "Link Code" in filtered_details.columns:
    filtered_details = filtered_details[filtered_details["Link Code"].isin(filtered_links)].copy()

filtered_penalties = penalties.copy()

# penalties filtering by link code if available
if selected_link != "All":
    if "Cluster Name" in filtered_penalties.columns:
        filtered_penalties = filtered_penalties[filtered_penalties["Cluster Name"] == selected_link].copy()
    elif "Link Code" in filtered_penalties.columns:
        filtered_penalties = filtered_penalties[filtered_penalties["Link Code"] == selected_link].copy()

if selected_region != "All" and "Region" in filtered_penalties.columns:
    filtered_penalties = filtered_penalties[filtered_penalties["Region"] == selected_region].copy()

# =========================================================
# TOP HEADER
# =========================================================
st.markdown('<div class="top-title">Dawiyat Executive Dashboard</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="top-subtitle">Top management dashboard for progress, schedule follow-up, penalties, PMO visibility, and work order details.</div>',
    unsafe_allow_html=True,
)

tabs = st.tabs([
    "Executive Overview",
    "PMO Summary",
    "Schedule & KPI",
    "Penalties & Quality",
    "Work Order Detail",
    "Dashboard Guide"
])

# =========================================================
# TAB 1: EXECUTIVE OVERVIEW
# =========================================================
with tabs[0]:
    if filter_df.empty:
        empty_message()
    else:
        total_links = filter_df["Link Code"].nunique()
        avg_progress = filter_df["Percentage of Completion"].fillna(0).mean()
        avg_civil = filter_df["civil_completion_pct"].fillna(0).mean()
        avg_fiber = filter_df["fiber_completion_pct"].fillna(0).mean()

        c1, c2, c3, c4 = st.columns(4)
        cards = [
            ("Total Link Codes", fmt_int(total_links), "Filtered link code count"),
            ("Avg Progress %", fmt_pct(avg_progress), "Overall progress average"),
            ("Avg Civil Completion", fmt_pct(avg_civil), "Trench progress / scope"),
            ("Avg Fiber Completion", fmt_pct(avg_fiber), "Fiber progress / scope"),
        ]
        for col, (title, value, note) in zip([c1, c2, c3, c4], cards):
            with col:
                st.markdown(
                    f"""
                    <div class="card">
                        <div class="card-title">{title}</div>
                        <div class="card-value">{value}</div>
                        <div class="card-note">{note}</div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

        c5, c6 = st.columns([1.2, 1])

        with c5:
            st.markdown('<div class="section-title">District Performance</div>', unsafe_allow_html=True)
            district_perf = (
                filter_df.groupby("District_final", dropna=False)["Percentage of Completion"]
                .mean()
                .reset_index()
                .sort_values("Percentage of Completion", ascending=False)
            )
            if district_perf.empty:
                empty_message("No district performance data for selected filters")
            else:
                fig = px.bar(
                    district_perf,
                    x="District_final",
                    y="Percentage of Completion",
                    text_auto=".1f",
                    color="Percentage of Completion",
                    color_continuous_scale="Blues",
                )
                fig.update_layout(
                    height=430,
                    paper_bgcolor="white",
                    plot_bgcolor="white",
                    margin=dict(l=10, r=10, t=20, b=60),
                    coloraxis_showscale=False,
                    xaxis_title="District",
                    yaxis_title="Avg Progress %",
                )
                st.plotly_chart(fig, use_container_width=True)

        with c6:
            st.markdown('<div class="section-title">Forecast Risk Mix</div>', unsafe_allow_html=True)
            risk_mix = (
                filter_df.groupby("forecast_risk")["Link Code"]
                .nunique()
                .reset_index(name="count")
            )
            if risk_mix.empty:
                empty_message("No forecast risk data for selected filters")
            else:
                fig2 = px.pie(
                    risk_mix,
                    names="forecast_risk",
                    values="count",
                    hole=0.55,
                    color_discrete_sequence=["#39d98a", "#2da8ff", "#ffb84d", "#ff6b6b"],
                )
                fig2.update_layout(
                    height=430,
                    paper_bgcolor="white",
                    margin=dict(l=10, r=10, t=20, b=10),
                    legend_title_text="Risk",
                )
                st.plotly_chart(fig2, use_container_width=True)

# =========================================================
# TAB 2: PMO SUMMARY
# =========================================================
with tabs[1]:
    st.markdown('<div class="section-title">PMO Summary</div>', unsafe_allow_html=True)

    if filter_df.empty:
        empty_message()
    else:
        pmo_links = (
            filter_df.groupby("Link Code", dropna=False)
            .agg({
                "Region_final": "first",
                "City_final": "first",
                "District_final": "first",
                "Percentage of Completion": "mean",
                "Updates": "max",
                "Updated": "max",
                "WO Cost": "sum",
                "Cost": "sum",
                "Targeted Completion": "min",
                "Updated Target Date": "min",
                "Effective Target Date": "min",
                "Work Order Status": lambda s: ", ".join(sorted(set([str(x) for x in s.dropna()]))),
                "lag_days": "max",
                "lag_pct": "max",
            })
            .reset_index()
        )

        # only meaningful records for update follow-up
        today = pd.Timestamp.today().normalize()
        two_months_later = today + pd.DateOffset(months=2)

        progress_nonzero = pmo_links["Percentage of Completion"].fillna(0) > 0
        update_rule = pmo_links["Updates"].fillna(0) < 5
        updated_rule = pmo_links["Updated"].fillna(0) > 5

        target_in_scope = (
            (
                pmo_links["Targeted Completion"].notna() &
                (pmo_links["Targeted Completion"] <= two_months_later)
            ) |
            (
                pmo_links["Updated Target Date"].notna() &
                (pmo_links["Updated Target Date"] <= two_months_later)
            )
        )

        followup_df = pmo_links[progress_nonzero & update_rule & updated_rule & target_in_scope].copy()

        pmo_links["best_flag"] = np.where(
            (pmo_links["Percentage of Completion"].fillna(0) >= 100) &
            (pmo_links["lag_days"].fillna(0) <= 0),
            1, 0
        )

        current_month = today.month
        current_year = today.year

        target_completion_cost = pmo_links[
            (pmo_links["Targeted Completion"].dt.month == current_month) &
            (pmo_links["Targeted Completion"].dt.year == current_year)
        ]["Cost"].fillna(0).sum()

        updated_target_cost = pmo_links[
            (pmo_links["Updated Target Date"].dt.month == current_month) &
            (pmo_links["Updated Target Date"].dt.year == current_year)
        ]["Cost"].fillna(0).sum()

        best_links = pmo_links[pmo_links["best_flag"] == 1].copy()
        best_links = best_links.sort_values("Percentage of Completion", ascending=False).head(10)

        cc1, cc2, cc3, cc4 = st.columns(4)
        metrics = [
            ("Need Update Follow-up", fmt_int(len(followup_df)), "Updates < 5, Updated > 5, progress > 0, target within 2 months"),
            ("Current Month Target Cost", fmt_money(target_completion_cost), "Based on Targeted Completion"),
            ("Updated Target Cost", fmt_money(updated_target_cost), "Based on Updated Target Date"),
            ("Best Completed Link Codes", fmt_int(len(best_links)), "Completed with no delay"),
        ]
        for col, (title, value, note) in zip([cc1, cc2, cc3, cc4], metrics):
            with col:
                st.markdown(
                    f"""
                    <div class="card">
                        <div class="card-title">{title}</div>
                        <div class="card-value">{value}</div>
                        <div class="card-note">{note}</div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

        a1, a2 = st.columns([1.1, 1])

        with a1:
            st.markdown('<div class="section-title">Link Codes Requiring Update</div>', unsafe_allow_html=True)
            if followup_df.empty:
                empty_message("No link codes require update under current filter")
            else:
                show_cols = ["Link Code", "Region_final", "City_final", "District_final", "Percentage of Completion", "Updates", "Updated", "Effective Target Date", "lag_pct"]
                followup_show = followup_df[show_cols].copy()
                followup_show.rename(columns={
                    "Region_final": "Region",
                    "City_final": "City",
                    "District_final": "District",
                    "Percentage of Completion": "Avg Progress %",
                    "Updated": "Last Updated (Days)",
                    "lag_pct": "Lag %",
                }, inplace=True)
                st.dataframe(followup_show, use_container_width=True, hide_index=True)

        with a2:
            st.markdown('<div class="section-title">Best Completed Link Codes</div>', unsafe_allow_html=True)
            if best_links.empty:
                empty_message("No fully completed no-delay link codes found in current filter")
            else:
                fig_best = px.bar(
                    best_links,
                    x="Link Code",
                    y="Percentage of Completion",
                    color="Percentage of Completion",
                    color_continuous_scale="Blues",
                    text_auto=".0f",
                )
                fig_best.update_layout(
                    height=360,
                    paper_bgcolor="white",
                    plot_bgcolor="white",
                    margin=dict(l=10, r=10, t=20, b=40),
                    coloraxis_showscale=False,
                    xaxis_title="Link Code",
                    yaxis_title="Completion %",
                )
                st.plotly_chart(fig_best, use_container_width=True)

        st.markdown('<div class="section-title">Monthly Cost Outlook</div>', unsafe_allow_html=True)
        monthly_cost = build_monthly_cost(pmo_links)
        if monthly_cost.empty:
            empty_message("No monthly cost data for selected filters")
        else:
            fig_cost = go.Figure()
            fig_cost.add_bar(
                x=monthly_cost["Month"],
                y=monthly_cost["Targeted Completion Cost"],
                name="Targeted Completion Cost"
            )
            fig_cost.add_scatter(
                x=monthly_cost["Month"],
                y=monthly_cost["Updated Target Date Cost"],
                mode="lines+markers",
                name="Updated Target Date Cost"
            )
            fig_cost.update_layout(
                height=420,
                paper_bgcolor="white",
                plot_bgcolor="white",
                margin=dict(l=10, r=10, t=20, b=40),
                xaxis_title="Month",
                yaxis_title="Cost",
            )
            st.plotly_chart(fig_cost, use_container_width=True)

# =========================================================
# TAB 3: SCHEDULE & KPI
# =========================================================
with tabs[2]:
    st.markdown('<div class="section-title">Schedule & KPI</div>', unsafe_allow_html=True)

    if filter_df.empty:
        empty_message()
    else:
        overdue_df = filter_df[
            (filter_df["Effective Target Date"].notna()) &
            (filter_df["Effective Target Date"] < pd.Timestamp.today().normalize()) &
            (filter_df["Percentage of Completion"].fillna(0) < 100)
        ].copy()

        critical_lag_count = int((filter_df["lag_pct"].fillna(0) >= 15).sum())
        total_rej = filter_df["total_rejections"].fillna(0).sum()

        k1, k2, k3, k4, k5 = st.columns(5)
        items = [
            ("Avg Civil Completion", fmt_pct(filter_df["civil_completion_pct"].fillna(0).mean()), "Trench progress / scope"),
            ("Avg Fiber Completion", fmt_pct(filter_df["fiber_completion_pct"].fillna(0).mean()), "Fiber progress / scope"),
            ("Overdue Work Orders", fmt_int(len(overdue_df)), "Past effective target and not complete"),
            ("Critical Lag", fmt_int(critical_lag_count), "Lag >= 15%"),
            ("Total Rejections", fmt_int(total_rej), "Sum of all rejection count columns"),
        ]
        for col, (title, value, note) in zip([k1, k2, k3, k4, k5], items):
            with col:
                st.markdown(
                    f"""
                    <div class="card">
                        <div class="card-title">{title}</div>
                        <div class="card-value">{value}</div>
                        <div class="card-note">{note}</div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

        st.markdown('<div class="section-title">Overdue Work Order Action List</div>', unsafe_allow_html=True)

        if overdue_df.empty:
            empty_message("No overdue work orders for selected filters")
        else:
            overdue_status_options = ["All"] + sorted(overdue_df["Work Order Status"].dropna().astype(str).unique().tolist())
            overdue_status_filter = st.selectbox("Filter overdue list by Work Order Status", overdue_status_options, key="od_status_filter")

            overdue_show = overdue_df.copy()
            if overdue_status_filter != "All":
                overdue_show = overdue_show[overdue_show["Work Order Status"] == overdue_status_filter]

            action_cols = [
                "Link Code", "Work Order", "Region_final", "City_final", "District_final",
                "Project", "Subclass", "Work Order Status", "Percentage of Completion",
                "Targeted Completion", "Updated Target Date", "Effective Target Date", "lag_days", "lag_pct"
            ]
            action_cols = [c for c in action_cols if c in overdue_show.columns]
            overdue_export = overdue_show[action_cols].copy()
            overdue_export.rename(columns={
                "Region_final": "Region",
                "City_final": "City",
                "District_final": "District",
                "Percentage of Completion": "Avg Progress %",
                "lag_days": "Lag Days",
                "lag_pct": "Lag %",
            }, inplace=True)

            st.dataframe(overdue_export, use_container_width=True, hide_index=True)

            cexp1, cexp2 = st.columns(2)
            with cexp1:
                st.download_button(
                    "Download Overdue List (Excel)",
                    data=safe_download_excel(overdue_export, "Overdue_WO"),
                    file_name="Dawiyat_Overdue_WorkOrders.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            with cexp2:
                st.download_button(
                    "Download Overdue List (CSV)",
                    data=overdue_export.to_csv(index=False).encode("utf-8-sig"),
                    file_name="Dawiyat_Overdue_WorkOrders.csv",
                    mime="text/csv",
                )

# =========================================================
# TAB 4: PENALTIES & QUALITY
# =========================================================
with tabs[3]:
    st.markdown('<div class="section-title">Penalties & Quality</div>', unsafe_allow_html=True)

    total_dev = filtered_penalties["Number of Deviations"].fillna(0).sum() if "Number of Deviations" in filtered_penalties.columns else 0
    total_penalty = filtered_penalties["Penalties Amount"].fillna(0).sum() if "Penalties Amount" in filtered_penalties.columns else 0

    pp1, pp2 = st.columns(2)
    metrics = [
        ("Number of Deviations", fmt_int(total_dev), "Sum of recorded deviations"),
        ("Penalty Deduction Amount", fmt_money(total_penalty), "Sum of deducted amount only"),
    ]
    for col, (title, value, note) in zip([pp1, pp2], metrics):
        with col:
            st.markdown(
                f"""
                <div class="card">
                    <div class="card-title">{title}</div>
                    <div class="card-value">{value}</div>
                    <div class="card-note">{note}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )

    s1, s2 = st.columns([1.1, 1])

    with s1:
        st.markdown('<div class="section-title">Top Deviations vs Penalty Amount</div>', unsafe_allow_html=True)
        if filtered_penalties.empty or "Deviation name" not in filtered_penalties.columns:
            empty_message("No penalties/deviation data for selected filters")
        else:
            top_dev = (
                filtered_penalties.groupby("Deviation name", dropna=False)
                .agg({
                    "Number of Deviations": "sum",
                    "Penalties Amount": "sum"
                })
                .reset_index()
                .sort_values(["Penalties Amount", "Number of Deviations"], ascending=False)
                .head(10)
            )

            if top_dev.empty:
                empty_message("No top deviations found for selected filters")
            else:
                fig_dev = go.Figure()
                fig_dev.add_bar(
                    y=top_dev["Deviation name"],
                    x=top_dev["Number of Deviations"],
                    orientation="h",
                    name="Number of Deviations"
                )
                fig_dev.add_scatter(
                    y=top_dev["Deviation name"],
                    x=top_dev["Penalties Amount"],
                    mode="markers",
                    name="Penalties Applied"
                )
                fig_dev.update_layout(
                    height=460,
                    paper_bgcolor="white",
                    plot_bgcolor="white",
                    margin=dict(l=10, r=10, t=20, b=20),
                    xaxis_title="Count / Amount",
                    yaxis_title="Deviation name",
                )
                st.plotly_chart(fig_dev, use_container_width=True)

    with s2:
        st.markdown('<div class="section-title">Top Quality / Penalty Exposure</div>', unsafe_allow_html=True)

        exposure = filter_df.groupby("Link Code", dropna=False).agg({"lag_pct": "max"}).reset_index()

        pen_link_col = None
        if "Cluster Name" in filtered_penalties.columns:
            pen_link_col = "Cluster Name"
        elif "Link Code" in filtered_penalties.columns:
            pen_link_col = "Link Code"

        if pen_link_col:
            pen_group = (
                filtered_penalties.groupby(pen_link_col, dropna=False)
                .agg({
                    "Number of Deviations": "sum",
                    "Penalties Amount": "sum"
                })
                .reset_index()
                .rename(columns={pen_link_col: "Link Code"})
            )
            exposure = exposure.merge(pen_group, on="Link Code", how="left")

        if "Number of Deviations" not in exposure.columns:
            exposure["Number of Deviations"] = 0
        if "Penalties Amount" not in exposure.columns:
            exposure["Penalties Amount"] = 0

        exposure["Number of Deviations"] = exposure["Number of Deviations"].fillna(0)
        exposure["Penalties Amount"] = exposure["Penalties Amount"].fillna(0)
        exposure.rename(columns={"lag_pct": "Lag %"}, inplace=True)

        exposure_show = exposure.sort_values(
            ["Penalties Amount", "Number of Deviations"], ascending=False
        ).head(15)

        if exposure_show.empty:
            empty_message("No exposure data for selected filters")
        else:
            st.dataframe(exposure_show, use_container_width=True, hide_index=True)

# =========================================================
# TAB 5: WORK ORDER DETAIL
# =========================================================
with tabs[4]:
    st.markdown('<div class="section-title">Work Order Detail</div>', unsafe_allow_html=True)

    if filtered_details.empty:
        empty_message("No work order details for selected filters")
    else:
        details_show = filtered_details.copy()

        pen_link_col = None
        if "Cluster Name" in filtered_penalties.columns:
            pen_link_col = "Cluster Name"
        elif "Link Code" in filtered_penalties.columns:
            pen_link_col = "Link Code"

        if pen_link_col:
            pen_link = (
                filtered_penalties.groupby(pen_link_col, dropna=False)
                .agg({
                    "Number of Deviations": "sum",
                    "Penalties Amount": "sum"
                })
                .reset_index()
                .rename(columns={pen_link_col: "Link Code"})
            )
            details_show = details_show.merge(pen_link, on="Link Code", how="left")

        if "Number of Deviations" not in details_show.columns:
            details_show["Number of Deviations"] = 0
        if "Penalties Amount" not in details_show.columns:
            details_show["Penalties Amount"] = 0

        details_show["Number of Deviations"] = details_show["Number of Deviations"].fillna(0)
        details_show["Penalties Amount"] = details_show["Penalties Amount"].fillna(0)

        desired_cols = [
            "Link Code", "Work Order", "Region_final", "City_final", "District_final", "Project", "Subclass", "Stage",
            "Work Order Status", "Percentage of Completion", "WO Cost", "Cost",
            "Targeted Completion", "Updated Target Date", "forecast_risk",
            "Number of Deviations", "Penalties Amount", "Parent", "Number", "Request",
            "SOR Reference Number", "Target Area", "Notes", "Scope of Work"
        ]
        desired_cols = [c for c in desired_cols if c in details_show.columns]
        details_show = details_show[desired_cols].copy()

        details_show.rename(columns={
            "Region_final": "Region",
            "City_final": "City",
            "District_final": "District",
            "Percentage of Completion": "Avg Progress %",
            "Penalties Amount": "Penalty Amount",
        }, inplace=True)

        st.dataframe(details_show, use_container_width=True, hide_index=True)

        d1, d2 = st.columns(2)
        with d1:
            st.download_button(
                "Download Work Order Detail (Excel)",
                data=safe_download_excel(details_show, "WorkOrder_Detail"),
                file_name="Dawiyat_WorkOrder_Detail.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        with d2:
            st.download_button(
                "Download Work Order Detail (CSV)",
                data=details_show.to_csv(index=False).encode("utf-8-sig"),
                file_name="Dawiyat_WorkOrder_Detail.csv",
                mime="text/csv",
            )

# =========================================================
# TAB 6: DASHBOARD GUIDE
# =========================================================
with tabs[5]:
    st.markdown('<div class="section-title">Dashboard Guide</div>', unsafe_allow_html=True)
    st.markdown(
        """
        <div class="guide-box">
        <h4>Key calculation logic</h4>
        <p><b>Avg Progress %</b>: average of <code>Percentage of Completion</code> after current filters.</p>
        <p><b>Avg Civil Completion</b>: <code>Trench Progress / Trench Scope</code>.</p>
        <p><b>Avg Fiber Completion</b>: <code>Fiber Progress / Fiber Scope</code>.</p>
        <p><b>Effective Target Date</b>: uses <code>Updated Target Date</code> when available, otherwise <code>Targeted Completion</code>.</p>
        <p><b>Overdue Work Orders</b>: effective target date already passed and completion is still below 100%.</p>
        <p><b>Critical Lag</b>: records with lag percentage at or above 15%.</p>
        <p><b>Total Rejections</b>: sum of PIP, PAT, Approval, As-Built, and Handover rejection count columns.</p>
        <p><b>Number of Deviations</b>: total recorded deviations from penalties sheet.</p>
        <p><b>Penalty Deduction Amount</b>: deducted amount only. Not every deviation has an applied penalty amount.</p>
        <p><b>Need Update Follow-up</b>: link code where progress &gt; 0, <code>Updates &lt; 5</code>, <code>Updated &gt; 5 days</code>, and target is within 2 months.</p>
        <p><b>Best Completed Link Codes</b>: completion reached 100% with no current delay.</p>
        <p><b>Monthly Cost Outlook</b>: safe aggregation by month for both <code>Targeted Completion</code> and <code>Updated Target Date</code>. Empty filtered selections do not crash the app.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

st.markdown('<div class="footer-note">Dawiyat Executive Dashboard | GitHub-hosted workbook mode | Production-ready safe filtering</div>', unsafe_allow_html=True)
