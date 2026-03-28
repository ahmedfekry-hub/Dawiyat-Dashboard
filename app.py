from __future__ import annotations

from pathlib import Path
import re
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


# ---------- Utilities ----------
def clean_text(value):
    if pd.isna(value):
        return np.nan
    text = str(value).strip()
    if text.lower() in {"", "nan", "none", "null"}:
        return np.nan
    return text


def canonical_key(value: str) -> str:
    text = str(value).upper().strip()
    text = re.sub(r"[\._/]+", " ", text)
    text = re.sub(r"[^A-Z0-9\- ]+", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


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


def first_existing(df: pd.DataFrame, names: list[str]) -> str | None:
    norm = {str(c).strip().lower(): c for c in df.columns}
    for name in names:
        key = name.strip().lower()
        if key in norm:
            return norm[key]
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


def ratio_pct(progress: pd.Series, scope: pd.Series) -> pd.Series:
    p = pd.to_numeric(progress, errors="coerce")
    s = pd.to_numeric(scope, errors="coerce")
    with np.errstate(divide="ignore", invalid="ignore"):
        result = np.where(s > 0, (p / s) * 100, np.nan)
    return pd.Series(result, index=progress.index).clip(lower=0, upper=100)


def fmt_money(x):
    if pd.isna(x):
        return "-"
    x = float(x)
    if abs(x) >= 1_000_000:
        return f"${x/1_000_000:,.2f}M"
    if abs(x) >= 1_000:
        return f"${x/1_000:,.1f}K"
    return f"${x:,.0f}"


def fmt_pct(x):
    return "-" if pd.isna(x) else f"{x:,.1f}%"


def color_template(theme_mode: str) -> str:
    return "plotly_dark" if theme_mode == "Dark" else "plotly_white"


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


# ---------- Theme ----------
def apply_theme_css(theme_mode: str):
    dark = theme_mode == "Dark"
    bg = "#081426" if dark else "#F5F7FB"
    card = "#0f1f38" if dark else "#FFFFFF"
    text = "#EAF2FF" if dark else "#182230"
    muted = "#8BA2C8" if dark else "#667085"
    border = "rgba(148,163,184,0.16)" if dark else "rgba(15,23,42,0.08)"
    accent = "#3B82F6"
    accent2 = "#F59E0B"
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
        .block-container {{
            padding-top: 1.2rem;
            padding-bottom: 2rem;
        }}
        section[data-testid="stSidebar"] {{
            background: linear-gradient(180deg, rgba(15,31,56,0.95), rgba(8,20,38,0.98));
            border-right: 1px solid rgba(148,163,184,0.15);
        }}
        section[data-testid="stSidebar"] * {{
            color: #EAF2FF !important;
        }}
        .top-banner {{
            background: linear-gradient(135deg, rgba(59,130,246,0.22), rgba(245,158,11,0.10));
            border: 1px solid {border};
            border-radius: 24px;
            padding: 22px 24px;
            margin-bottom: 14px;
            box-shadow: {glow};
        }}
        .metric-card {{
            background: {card};
            border: 1px solid {border};
            border-radius: 18px;
            padding: 18px 18px 14px 18px;
            min-height: 132px;
            box-shadow: {glow};
        }}
        .metric-title {{
            color: {muted};
            font-size: 0.92rem;
            margin-bottom: 6px;
        }}
        .metric-value {{
            color: {text};
            font-size: 2rem;
            font-weight: 700;
            line-height: 1.15;
            margin-bottom: 8px;
        }}
        .metric-subtitle {{
            color: {muted};
            font-size: 0.9rem;
        }}
        .section-card {{
            background: {card};
            border: 1px solid {border};
            border-radius: 22px;
            padding: 10px 14px 2px 14px;
            box-shadow: {glow};
        }}
        .guide-box {{
            background: {card};
            border-left: 4px solid {accent};
            border-radius: 14px;
            border: 1px solid {border};
            padding: 14px 16px;
            margin: 10px 0px;
        }}
        .subtle {{
            color: {muted};
            font-size: 0.95rem;
        }}
        .small-chip {{
            display: inline-block;
            padding: 6px 10px;
            border-radius: 999px;
            background: rgba(59,130,246,0.12);
            color: {text};
            font-size: 0.85rem;
            border: 1px solid {border};
            margin-right: 8px;
            margin-top: 4px;
        }}
        .stDataFrame, div[data-testid="stTable"] {{
            border-radius: 16px;
            overflow: hidden;
            border: 1px solid {border};
        }}
        .stTabs [data-baseweb="tab-list"] {{
            gap: 8px;
        }}
        .stTabs [data-baseweb="tab"] {{
            background: {card};
            border: 1px solid {border};
            border-radius: 12px 12px 0 0;
            padding: 10px 16px;
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )


# ---------- Data loading ----------
@st.cache_data(show_spinner=False)
def load_data(file_source=None):
    if file_source is None:
        file_source = Path(__file__).with_name(DEFAULT_FILE)

    xls = pd.ExcelFile(file_source)
    main_name = find_sheet_name(xls, ["Dawaiyat Service Tool", "service tool"])
    district_name = find_sheet_name(xls, ["District", "District "])
    penalties_name = find_sheet_name(xls, ["Penalties", "Penalty"])

    main = pd.read_excel(file_source, sheet_name=main_name)
    district = pd.read_excel(file_source, sheet_name=district_name)
    penalties = pd.read_excel(file_source, sheet_name=penalties_name)

    main.columns = [str(c).strip() for c in main.columns]
    district.columns = [str(c).strip() for c in district.columns]
    penalties.columns = [str(c).strip() for c in penalties.columns]

    main = main.loc[:, ~main.columns.isna()].copy()
    main = main[main[first_existing(main, ["Link Code"])].notna()].copy()

    for df in [main, district, penalties]:
        for col in df.columns:
            if df[col].dtype == "object":
                df[col] = df[col].map(clean_text)

    link_col = first_existing(main, ["Link Code"])
    district_link_col = first_existing(district, ["Link Code"])
    city_col_d = first_existing(district, ["City"])
    district_col_d = first_existing(district, ["District"])

    main[link_col] = main[link_col].astype(str).str.strip()
    district = district[district[district_link_col].notna()].copy()
    district[district_link_col] = district[district_link_col].astype(str).str.strip()

    district_map = (
        district.groupby(district_link_col, dropna=False)
        .agg(
            City=(city_col_d, choose_mode),
            District=(district_col_d, choose_mode),
        )
        .reset_index()
        .rename(columns={district_link_col: "Link Code"})
    )

    data = main.rename(columns={link_col: "Link Code"}).merge(district_map, on="Link Code", how="left")

    # Fallback city/district from link-code pattern for links missing in District sheet
    district["p1"] = district[district_link_col].astype(str).str.split("-").str[0].str.upper()
    district["p3"] = district[district_link_col].astype(str).str.split("-").str[2].str.upper()
    city_prefix_map = district.groupby("p1")[city_col_d].agg(choose_mode).dropna().to_dict()
    district_prefix_map = district.groupby("p3")[district_col_d].agg(choose_mode).dropna().to_dict()

    data["p1"] = data["Link Code"].astype(str).str.split("-").str[0].str.upper()
    data["p3"] = data["Link Code"].astype(str).str.split("-").str[2].str.upper()
    data["City"] = data["City"].fillna(data["p1"].map(city_prefix_map))
    data["District"] = data["District"].fillna(data["p3"].map(district_prefix_map))

    # Penalties
    cluster_col = first_existing(penalties, ["Cluster Name", "Link Code"])
    penalty_name_col = first_existing(penalties, ["Impl # of Penalty", "Penalty"])
    penalty_qty_col = first_existing(penalties, ["Number", "Qty"])
    penalty_amt_col = first_existing(penalties, ["Penalties Amount", "Amount"])
    penalty_city_col = first_existing(penalties, ["City"])
    penalty_region_col = first_existing(penalties, ["Region"])

    if cluster_col:
        penalties[cluster_col] = penalties[cluster_col].astype(str).str.strip()
        if penalty_qty_col:
            penalties[penalty_qty_col] = pd.to_numeric(penalties[penalty_qty_col], errors="coerce").fillna(0)
        else:
            penalties["_qty"] = 0
            penalty_qty_col = "_qty"
        if penalty_amt_col:
            penalties[penalty_amt_col] = pd.to_numeric(penalties[penalty_amt_col], errors="coerce").fillna(0)
        else:
            penalties["_amt"] = 0
            penalty_amt_col = "_amt"

        penalties_agg = (
            penalties.groupby(cluster_col)
            .agg(
                penalty_cases=(penalty_name_col, "count") if penalty_name_col else (penalty_qty_col, "size"),
                penalty_qty=(penalty_qty_col, "sum"),
                penalty_amount=(penalty_amt_col, "sum"),
            )
            .reset_index()
            .rename(columns={cluster_col: "Link Code"})
        )
        data = data.merge(penalties_agg, on="Link Code", how="left")
    else:
        data["penalty_cases"] = 0
        data["penalty_qty"] = 0
        data["penalty_amount"] = 0

    for col in ["penalty_cases", "penalty_qty", "penalty_amount"]:
        data[col] = pd.to_numeric(data[col], errors="coerce").fillna(0)

    # Dates
    for col in ["Created", "Assigned at", "In Progress at", "Updated", "Closed at", "Targeted Completion", "Updated Target Date"]:
        if col in data.columns:
            data[col] = pd.to_datetime(data[col], errors="coerce")
    if "Assigned Date" in penalties.columns:
        penalties["Assigned Date"] = pd.to_datetime(penalties["Assigned Date"], errors="coerce")

    # Numeric cleanup
    numeric_cols = [
        "Percentage of Completion", "WO Cost", "Cost", "Trench Progress", "Trench Scope", "MH/HH Progress", "MH/HH Scope",
        "Fiber Progress", "Fiber Scope", "ODBs Progress", "ODBs Scope", "ODFs Progress", "ODFs Scope", "JCL Progress",
        "JCL Scope", "FAT Progress", "FAT Scope", "PFAT Progress", "PFAT Scope", "SFAT Progress", "SFAT Scope",
        "Permits Progress", "Permits Scope", "PIP rejection count", "PAT rejection count", "Approval rejection count",
        "As-Built Rejection Count", "Handover Rejection Count"
    ]
    for col in numeric_cols:
        if col in data.columns:
            data[col] = pd.to_numeric(data[col], errors="coerce")

    snapshot_date = data["Updated"].dropna().max() if "Updated" in data.columns else pd.NaT
    if pd.isna(snapshot_date):
        snapshot_date = pd.Timestamp.today().normalize()

    data["effective_target"] = data.get("Updated Target Date").combine_first(data.get("Targeted Completion"))
    data["start_date"] = data.get("In Progress at").combine_first(data.get("Assigned at")).combine_first(data.get("Created"))

    elapsed = (snapshot_date - data["start_date"]).dt.days
    total = (data["effective_target"] - data["start_date"]).dt.days
    with np.errstate(divide="ignore", invalid="ignore"):
        data["planned_progress_pct"] = np.where(total > 0, np.clip((elapsed / total) * 100, 0, 100), np.nan)

    data["actual_progress_pct"] = pd.to_numeric(data.get("Percentage of Completion"), errors="coerce")
    data["actual_progress_capped"] = data["actual_progress_pct"].clip(lower=0, upper=100)
    data["schedule_variance_pp"] = data["actual_progress_capped"] - data["planned_progress_pct"]
    data["is_complete"] = data["actual_progress_capped"] >= 100
    data["is_overdue"] = (snapshot_date > data["effective_target"]) & (~data["is_complete"]) & data["effective_target"].notna()
    data["critical_lag"] = data["schedule_variance_pp"] <= -15

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

    # Milestone completion percentages
    data["civil_completion_pct"] = ratio_pct(data.get("Trench Progress"), data.get("Trench Scope"))
    data["mhhh_completion_pct"] = ratio_pct(data.get("MH/HH Progress"), data.get("MH/HH Scope"))
    data["fiber_completion_pct"] = ratio_pct(data.get("Fiber Progress"), data.get("Fiber Scope"))
    data["permits_completion_pct"] = ratio_pct(data.get("Permits Progress"), data.get("Permits Scope"))

    # Default labels
    for col in ["Year", "Work Order Status", "Type", "Class", "Project", "Subclass", "Stage", "Region", "City", "District"]:
        if col not in data.columns:
            data[col] = np.nan
    data["Year"] = data["Year"].fillna(pd.to_datetime(data["effective_target"], errors="coerce").dt.year.astype("Int64"))
    for col in ["Work Order Status", "Type", "Class", "Project", "Subclass", "Stage", "Region", "City", "District"]:
        data[col] = data[col].fillna("Not Classified")
    data["Month"] = pd.to_datetime(data.get("Updated"), errors="coerce").dt.to_period("M").astype(str)

    if penalty_region_col and "Region" not in penalties.columns:
        penalties["Region"] = penalties[penalty_region_col]
    if penalty_city_col and "City" not in penalties.columns:
        penalties["City"] = penalties[penalty_city_col]
    penalties["Region"] = penalties.get("Region", pd.Series(index=penalties.index)).fillna("Not Classified")
    penalties["City"] = penalties.get("City", pd.Series(index=penalties.index)).fillna("Not Classified")
    penalties["Link Code"] = penalties.get(cluster_col, pd.Series(index=penalties.index)).astype(str).replace("nan", np.nan)
    penalties["Number"] = pd.to_numeric(penalties.get(penalty_qty_col, 0), errors="coerce").fillna(0)
    penalties["Penalties Amount"] = pd.to_numeric(penalties.get(penalty_amt_col, 0), errors="coerce").fillna(0)

    return data, penalties, snapshot_date


# ---------- Sidebar and data ----------
st.sidebar.title("Control Panel")
theme_mode = st.sidebar.radio("Theme", ["Dark", "Light"], horizontal=True, index=0)
apply_theme_css(theme_mode)

uploaded_file = st.sidebar.file_uploader("Upload refreshed Dawiyat workbook", type=["xlsx"])
source = uploaded_file if uploaded_file is not None else None

data, penalties, snapshot_date = load_data(source)

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

# ---------- Cascading filters ----------
def options_for(df: pd.DataFrame, col: str):
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

# ---------- KPI calculations ----------
total_wo = filtered["Work Order"].nunique() if "Work Order" in filtered.columns else len(filtered)
total_link = filtered["Link Code"].nunique()
avg_actual = filtered["actual_progress_capped"].mean()
avg_planned = filtered["planned_progress_pct"].mean()
on_track_pct = (filtered["schedule_variance_pp"] >= 0).mean() * 100
overdue_cnt = int(filtered["is_overdue"].sum())
critical_lag_cnt = int(filtered["critical_lag"].sum())
forecast_high_risk = int((filtered["forecast_risk"] == "High delay risk").sum())
penalty_cases = int(filtered["penalty_cases"].sum())
penalty_qty = float(filtered["penalty_qty"].sum())
penalty_amount = float(filtered["penalty_amount"].sum())
avg_civil = filtered["civil_completion_pct"].mean()
avg_fiber = filtered["fiber_completion_pct"].mean()
rejections_total = filtered[[c for c in ["PIP rejection count", "PAT rejection count", "Approval rejection count", "As-Built Rejection Count", "Handover Rejection Count"] if c in filtered.columns]].fillna(0).sum().sum()

tabs = st.tabs(["Executive Overview", "Schedule & KPI", "Penalties & Quality", "Work Order Detail", "Dashboard Guide"])

with tabs[0]:
    r1 = st.columns(6)
    with r1[0]:
        add_card("Work Orders", f"{total_wo:,}", f"{total_link:,} link codes")
    with r1[1]:
        add_card("Avg Actual Progress", fmt_pct(avg_actual), "Overall completion")
    with r1[2]:
        add_card("Avg Planned Progress", fmt_pct(avg_planned), "Based on target dates")
    with r1[3]:
        add_card("Civil Completion", fmt_pct(avg_civil), "Trench scope delivered")
    with r1[4]:
        add_card("Fiber Completion", fmt_pct(avg_fiber), "Fiber scope delivered")
    with r1[5]:
        add_card("Penalty Cases", f"{penalty_cases:,}", f"Qty {penalty_qty:,.0f}")

    r2 = st.columns(4)
    with r2[0]:
        add_card("On Track", fmt_pct(on_track_pct), "Actual vs planned")
    with r2[1]:
        add_card("Overdue WO", f"{overdue_cnt:,}", "Past target date")
    with r2[2]:
        add_card("Critical Lag >15pp", f"{critical_lag_cnt:,}", "Immediate action")
    with r2[3]:
        add_card("High Forecast Risk", f"{forecast_high_risk:,}", "Delay exposure")

    c1, c2 = st.columns([1.3, 1])
    with c1:
        by_month = (
            filtered.dropna(subset=["Updated"])
            .groupby("Month", dropna=False)
            .agg(actual=("actual_progress_capped", "mean"), planned=("planned_progress_pct", "mean"))
            .reset_index()
        )
        by_month = by_month[by_month["Month"].notna() & (by_month["Month"] != "NaT")]
        if not by_month.empty:
            fig = go.Figure()
            fig.add_trace(go.Scatter(x=by_month["Month"], y=by_month["actual"], mode="lines+markers", name="Actual %"))
            fig.add_trace(go.Scatter(x=by_month["Month"], y=by_month["planned"], mode="lines+markers", name="Planned %"))
            fig.update_layout(
                title="Planned vs Actual Progress Trend",
                template=color_template(theme_mode),
                height=400,
                legend_orientation="h",
                margin=dict(l=20, r=20, t=50, b=20),
                yaxis_title="Progress %",
                xaxis_title="",
            )
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("No monthly update trend available for the current filter.")

    with c2:
        city_perf = (
            filtered.groupby(["City", "District"], dropna=False)
            .agg(
                work_orders=("Work Order", "nunique"),
                actual=("actual_progress_capped", "mean"),
                planned=("planned_progress_pct", "mean"),
                penalties=("penalty_cases", "sum"),
            )
            .reset_index()
            .sort_values(["actual", "penalties"], ascending=[False, True])
            .head(12)
        )
        fig2 = px.bar(
            city_perf,
            x="actual",
            y="District",
            color="City",
            orientation="h",
            text="actual",
            hover_data=["work_orders", "planned", "penalties"],
            template=color_template(theme_mode),
            title="Top District Performance",
        )
        fig2.update_traces(texttemplate="%{text:.1f}%")
        fig2.update_layout(height=400, margin=dict(l=10, r=10, t=50, b=10), yaxis_title="", xaxis_title="Actual Progress %")
        st.plotly_chart(fig2, use_container_width=True)

    c3, c4 = st.columns([1.05, 0.95])
    with c3:
        milestone = pd.DataFrame(
            {
                "Milestone": ["Civil / Trench", "MH/HH", "Fiber", "Permits"],
                "Completion": [
                    filtered["civil_completion_pct"].mean(),
                    filtered["mhhh_completion_pct"].mean(),
                    filtered["fiber_completion_pct"].mean(),
                    filtered["permits_completion_pct"].mean(),
                ],
            }
        ).dropna()
        fig3 = px.bar(
            milestone,
            x="Milestone",
            y="Completion",
            text="Completion",
            color="Milestone",
            template=color_template(theme_mode),
            title="Milestone Completion Summary",
        )
        fig3.update_traces(texttemplate="%{text:.1f}%")
        fig3.update_layout(height=360, margin=dict(l=10, r=10, t=50, b=10), showlegend=False, yaxis_title="Completion %")
        st.plotly_chart(fig3, use_container_width=True)

    with c4:
        risk_df = pd.DataFrame(
            [
                ["Schedule Delay", forecast_high_risk],
                ["Critical Lag", critical_lag_cnt],
                ["Overdue Orders", overdue_cnt],
                ["Penalty Cases", penalty_cases],
                ["Rejections", rejections_total],
            ],
            columns=["Risk", "Count"],
        )
        fig4 = px.treemap(
            risk_df,
            path=["Risk"],
            values="Count",
            color="Count",
            color_continuous_scale="RdYlGn_r",
            template=color_template(theme_mode),
            title="Risk Exposure Overview",
        )
        fig4.update_layout(height=360, margin=dict(l=10, r=10, t=50, b=10))
        st.plotly_chart(fig4, use_container_width=True)

with tabs[1]:
    c1, c2 = st.columns([1.2, 1])
    with c1:
        wo_sched = filtered[["Link Code", "Work Order", "City", "District", "actual_progress_capped", "planned_progress_pct", "schedule_variance_pp", "effective_target", "forecast_completion_date", "forecast_delay_days", "forecast_risk"]].copy()
        wo_sched = wo_sched.sort_values(["forecast_delay_days", "schedule_variance_pp"], ascending=[False, True]).head(20)
        wo_sched["Actual %"] = wo_sched["actual_progress_capped"].round(1)
        wo_sched["Planned %"] = wo_sched["planned_progress_pct"].round(1)
        fig = go.Figure()
        fig.add_trace(go.Bar(x=wo_sched["Link Code"], y=wo_sched["Planned %"], name="Planned %"))
        fig.add_trace(go.Bar(x=wo_sched["Link Code"], y=wo_sched["Actual %"], name="Actual %"))
        fig.update_layout(
            barmode="group",
            template=color_template(theme_mode),
            height=430,
            title="Top 20 Work Orders: Planned vs Actual",
            margin=dict(l=10, r=10, t=50, b=10),
            xaxis_title="Link Code",
            yaxis_title="Progress %",
        )
        st.plotly_chart(fig, use_container_width=True)

    with c2:
        forecast = (
            filtered.groupby("City", dropna=False)
            .agg(
                avg_delay=("forecast_delay_days", "mean"),
                high_risk=("forecast_risk", lambda s: (s == "High delay risk").sum()),
                overdue=("is_overdue", "sum"),
            )
            .reset_index()
            .sort_values("avg_delay", ascending=False)
        )
        figf = px.bar(
            forecast,
            x="City",
            y="avg_delay",
            color="high_risk",
            text="avg_delay",
            template=color_template(theme_mode),
            title="Forecast KPI Impact by City",
            color_continuous_scale="Turbo",
        )
        figf.update_traces(texttemplate="%{text:.0f}d")
        figf.update_layout(height=430, margin=dict(l=10, r=10, t=50, b=10), yaxis_title="Average Delay Days")
        st.plotly_chart(figf, use_container_width=True)

    district_kpi = (
        filtered.groupby(["City", "District"], dropna=False)
        .agg(
            work_orders=("Work Order", "nunique"),
            actual=("actual_progress_capped", "mean"),
            planned=("planned_progress_pct", "mean"),
            schedule_gap=("schedule_variance_pp", "mean"),
            penalties=("penalty_cases", "sum"),
            high_risk=("forecast_risk", lambda s: (s == "High delay risk").sum()),
        )
        .reset_index()
        .sort_values(["schedule_gap", "penalties"], ascending=[True, False])
    )
    figd = px.scatter(
        district_kpi,
        x="planned",
        y="actual",
        size="work_orders",
        color="schedule_gap",
        hover_name="District",
        hover_data=["City", "penalties", "high_risk"],
        template=color_template(theme_mode),
        title="District Planned vs Actual Performance",
        color_continuous_scale="RdYlGn",
    )
    figd.add_shape(type="line", x0=0, y0=0, x1=100, y1=100, line=dict(dash="dash"))
    figd.update_layout(height=420, margin=dict(l=10, r=10, t=50, b=10), xaxis_title="Planned %", yaxis_title="Actual %")
    st.plotly_chart(figd, use_container_width=True)

with tabs[2]:
    p1, p2 = st.columns([1, 1.1])
    with p1:
        if not penalties.empty:
            pen_city = (
                penalties.groupby(["Region", "City"], dropna=False)
                .agg(cases=("Link Code", "count"), qty=("Number", "sum"), amount=("Penalties Amount", "sum"))
                .reset_index()
                .sort_values("qty", ascending=False)
            )
            figp1 = px.bar(
                pen_city,
                x="City",
                y="qty",
                color="Region",
                text="cases",
                template=color_template(theme_mode),
                title="Penalty Quantity by City",
            )
            figp1.update_traces(texttemplate="Cases %{text}")
            figp1.update_layout(height=380, margin=dict(l=10, r=10, t=50, b=10), yaxis_title="Penalty Quantity")
            st.plotly_chart(figp1, use_container_width=True)
        else:
            st.info("No penalties sheet detected.")

    with p2:
        quality = pd.DataFrame(
            {
                "Metric": ["PIP", "PAT", "Approval", "As-Built", "Handover"],
                "Count": [
                    filtered.get("PIP rejection count", pd.Series(dtype=float)).fillna(0).sum(),
                    filtered.get("PAT rejection count", pd.Series(dtype=float)).fillna(0).sum(),
                    filtered.get("Approval rejection count", pd.Series(dtype=float)).fillna(0).sum(),
                    filtered.get("As-Built Rejection Count", pd.Series(dtype=float)).fillna(0).sum(),
                    filtered.get("Handover Rejection Count", pd.Series(dtype=float)).fillna(0).sum(),
                ],
            }
        )
        figq = px.bar(
            quality,
            x="Metric",
            y="Count",
            text="Count",
            color="Metric",
            template=color_template(theme_mode),
            title="Quality Rejections Summary",
        )
        figq.update_layout(height=380, showlegend=False, margin=dict(l=10, r=10, t=50, b=10))
        st.plotly_chart(figq, use_container_width=True)

    if not penalties.empty:
        pen_detail = penalties.copy()
        wanted = [c for c in ["Region", "City", "Link Code", "Impl # of Penalty", "Inspection Consultant", "Implementation Contractor", "Assigned Date", "Number", "Penalties Amount"] if c in pen_detail.columns]
        st.dataframe(pen_detail[wanted].sort_values("Assigned Date", ascending=False).head(200), use_container_width=True, height=320)

with tabs[3]:
    detail = filtered.copy()
    detail["Actual %"] = detail["actual_progress_capped"].round(1)
    detail["Planned %"] = detail["planned_progress_pct"].round(1)
    detail["Civil %"] = detail["civil_completion_pct"].round(1)
    detail["Fiber %"] = detail["fiber_completion_pct"].round(1)
    detail["Forecast Delay Days"] = detail["forecast_delay_days"].round(0)
    show_cols = [
        c for c in [
            "Region", "City", "District", "Link Code", "Work Order", "Project", "Subclass", "Stage",
            "Work Order Status", "Actual %", "Planned %", "Civil %", "Fiber %", "effective_target",
            "forecast_completion_date", "Forecast Delay Days", "forecast_risk", "penalty_cases", "penalty_qty",
            "WO Cost", "Cost"
        ] if c in detail.columns
    ]
    st.dataframe(detail[show_cols].sort_values(["Forecast Delay Days", "penalty_cases"], ascending=[False, False]), use_container_width=True, height=520)

with tabs[4]:
    st.markdown(
        """
        <div class="guide-box">
        <b>How to read this dashboard</b><br><br>
        • <b>Avg Actual Progress</b>: current overall work-order completion from the Dawiyat Service Tool.<br>
        • <b>Avg Planned Progress</b>: calculated from start date versus targeted or updated completion date.<br>
        • <b>Civil Completion</b>: trench progress divided by trench scope.<br>
        • <b>Fiber Completion</b>: fiber progress divided by fiber scope.<br>
        • <b>Critical Lag >15pp</b>: work orders where actual progress is more than 15 percentage points below plan.<br>
        • <b>High Forecast Risk</b>: forecast finish is more than 30 days later than the effective target date.<br>
        • <b>Penalty Cases</b>: count of consultant or quality penalties linked to the work order.<br>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown(
        f"""
        <span class="small-chip">Region: {region}</span>
        <span class="small-chip">City: {city}</span>
        <span class="small-chip">District: {district}</span>
        <span class="small-chip">Project: {project}</span>
        <span class="small-chip">Stage: {stage}</span>
        <span class="small-chip">Year: {year}</span>
        <span class="small-chip">Status: {status}</span>
        <span class="small-chip">Type: {wo_type}</span>
        <span class="small-chip">Class: {wo_class}</span>
        <span class="small-chip">Subclass: {subclass}</span>
        <span class="small-chip">Link Code: {link_code}</span>
        """,
        unsafe_allow_html=True,
    )
