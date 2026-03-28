
import math
from pathlib import Path
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st

st.set_page_config(
    page_title="Dawiyat Project Intelligence Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

DEFAULT_FILE = "Dawiyat Master Sheet.xlsx"

def find_sheet_name(xls, candidates):
    normalized = {str(name).strip().lower(): name for name in xls.sheet_names}
    for c in candidates:
        if c.strip().lower() in normalized:
            return normalized[c.strip().lower()]
    for name in xls.sheet_names:
        low = str(name).strip().lower()
        if any(c.strip().lower() in low for c in candidates):
            return name
    raise ValueError(f"Could not find any of the expected sheets: {candidates}")

@st.cache_data(show_spinner=False)
def load_data(file_bytes=None, file_name=None):
    if file_bytes is not None:
        xls = pd.ExcelFile(file_bytes)
        main_name = find_sheet_name(xls, ["Dawaiyat Service Tool", "service tool"])
        district_name = find_sheet_name(xls, ["District", "District "])
        penalties_name = find_sheet_name(xls, ["Penalties", "Penalty"])
        main = pd.read_excel(file_bytes, sheet_name=main_name)
        district = pd.read_excel(file_bytes, sheet_name=district_name)
        penalties = pd.read_excel(file_bytes, sheet_name=penalties_name)
    else:
        default_path = Path(__file__).with_name(DEFAULT_FILE)
        xls = pd.ExcelFile(default_path)
        main_name = find_sheet_name(xls, ["Dawaiyat Service Tool", "service tool"])
        district_name = find_sheet_name(xls, ["District", "District "])
        penalties_name = find_sheet_name(xls, ["Penalties", "Penalty"])
        main = pd.read_excel(default_path, sheet_name=main_name)
        district = pd.read_excel(default_path, sheet_name=district_name)
        penalties = pd.read_excel(default_path, sheet_name=penalties_name)

    main = main.loc[:, ~main.columns.isna()].copy()
    main = main[main["Link Code"].notna()].copy()

    district.columns = [str(c).strip() for c in district.columns]
    penalties.columns = [str(c).strip() for c in penalties.columns]
    main.columns = [str(c).strip() for c in main.columns]

    # Normalize text
    for df in [main, district, penalties]:
        for col in df.columns:
            if df[col].dtype == "object":
                df[col] = df[col].astype(str).replace("nan", np.nan).str.strip()

    # District mapping by Link Code
    district = district[district["Link Code"].notna()].copy()
    district["Link Code"] = district["Link Code"].astype(str).str.strip()
    district_map = (
        district.groupby("Link Code", dropna=False)
        .agg(
            City=("City", lambda s: s.dropna().mode().iat[0] if not s.dropna().empty else np.nan),
            District=("District", lambda s: s.dropna().mode().iat[0] if not s.dropna().empty else np.nan),
        )
        .reset_index()
    )

    main["Link Code"] = main["Link Code"].astype(str).str.strip()
    data = main.merge(district_map, on="Link Code", how="left")

    # Penalties aggregation
    if "Cluster Name" in penalties.columns:
        penalties["Cluster Name"] = penalties["Cluster Name"].astype(str).str.strip()
        penalties["Number"] = pd.to_numeric(penalties.get("Number"), errors="coerce").fillna(0)
        penalties["Penalties Amount"] = pd.to_numeric(penalties.get("Penalties Amount"), errors="coerce").fillna(0)
        penalties_agg = (
            penalties.groupby("Cluster Name")
            .agg(
                penalty_cases=("Impl # of Penalty", "count"),
                penalty_qty=("Number", "sum"),
                penalty_amount=("Penalties Amount", "sum"),
            )
            .reset_index()
            .rename(columns={"Cluster Name": "Link Code"})
        )
        data = data.merge(penalties_agg, on="Link Code", how="left")
    else:
        data["penalty_cases"] = 0
        data["penalty_qty"] = 0
        data["penalty_amount"] = 0

    for col in ["penalty_cases", "penalty_qty", "penalty_amount"]:
        data[col] = pd.to_numeric(data[col], errors="coerce").fillna(0)

    # Dates and numeric cleanup
    date_cols = ["Created", "Assigned at", "In Progress at", "Updated", "Closed at", "Targeted Completion", "Updated Target Date"]
    for col in date_cols:
        if col in data.columns:
            data[col] = pd.to_datetime(data[col], errors="coerce")

    num_cols = [
        "Percentage of Completion", "WO Cost", "Cost", "Trench Progress", "Trench Scope", "MH/HH Progress", "MH/HH Scope",
        "Fiber Progress", "Fiber Scope", "ODBs Progress", "ODBs Scope", "ODFs Progress", "ODFs Scope", "JCL Progress",
        "JCL Scope", "FAT Progress", "FAT Scope", "PFAT Progress", "PFAT Scope", "SFAT Progress", "SFAT Scope",
        "Permits Progress", "Permits Scope", "PIP rejection count", "PAT rejection count", "Approval rejection count",
        "As-Built Rejection Count", "Handover Rejection Count"
    ]
    for col in num_cols:
        if col in data.columns:
            data[col] = pd.to_numeric(data[col], errors="coerce")

    data["snapshot_date"] = data["Updated"].dropna().max()
    if pd.isna(data["snapshot_date"]).all() if isinstance(data["snapshot_date"], pd.Series) else pd.isna(data["snapshot_date"]):
        data["snapshot_date"] = pd.Timestamp.today().normalize()

    data["effective_target"] = data["Updated Target Date"].combine_first(data["Targeted Completion"])
    data["start_date"] = data["In Progress at"].combine_first(data["Assigned at"]).combine_first(data["Created"])
    snapshot_date = pd.to_datetime(data["snapshot_date"].iloc[0] if hasattr(data["snapshot_date"], "iloc") else data["snapshot_date"])

    elapsed = (snapshot_date - data["start_date"]).dt.days
    total = (data["effective_target"] - data["start_date"]).dt.days
    with np.errstate(divide="ignore", invalid="ignore"):
        planned_pct = np.where(
            total > 0,
            np.clip((elapsed / total) * 100, 0, 100),
            np.nan
        )
    data["planned_progress_pct"] = planned_pct
    data["actual_progress_pct"] = pd.to_numeric(data["Percentage of Completion"], errors="coerce")
    data["actual_progress_capped"] = data["actual_progress_pct"].clip(lower=0, upper=100)
    data["schedule_variance_pp"] = data["actual_progress_capped"] - data["planned_progress_pct"]
    data["is_complete"] = data["actual_progress_capped"] >= 100
    data["is_overdue"] = (snapshot_date > data["effective_target"]) & (~data["is_complete"]) & data["effective_target"].notna()
    data["critical_lag"] = data["schedule_variance_pp"] <= -15

    # Earned value approximations
    budget = pd.to_numeric(data["WO Cost"], errors="coerce").fillna(0)
    actual_cost = pd.to_numeric(data["Cost"], errors="coerce").fillna(0)
    data["EV"] = budget * (data["actual_progress_capped"].fillna(0) / 100.0)
    data["PV"] = budget * (data["planned_progress_pct"].fillna(0) / 100.0)
    data["AC"] = actual_cost
    data["SPI"] = np.where(data["PV"] > 0, data["EV"] / data["PV"], np.nan)
    data["CPI"] = np.where(data["AC"] > 0, data["EV"] / data["AC"], np.nan)
    data["cost_variance"] = data["EV"] - data["AC"]
    data["schedule_variance"] = data["EV"] - data["PV"]

    # Forecast finish estimate
    actual_ratio = data["actual_progress_capped"] / 100.0
    elapsed_days = np.maximum((snapshot_date - data["start_date"]).dt.days, 1)
    est_total_duration = np.where(actual_ratio > 0, elapsed_days / actual_ratio, np.nan)
    data["forecast_completion_date"] = data["start_date"] + pd.to_timedelta(est_total_duration, unit="D")
    data["forecast_delay_days"] = (data["forecast_completion_date"] - data["effective_target"]).dt.days
    data["forecast_risk"] = np.select(
        [
            data["forecast_delay_days"] > 30,
            data["forecast_delay_days"] > 0,
            data["forecast_delay_days"] <= 0,
        ],
        ["High delay risk", "Moderate delay risk", "On forecast"],
        default="Insufficient data"
    )

    # Helper group labels
    data["Month"] = pd.to_datetime(data["Updated"], errors="coerce").dt.to_period("M").astype(str)
    data["Year"] = data["Year"].fillna(pd.to_datetime(data["effective_target"], errors="coerce").dt.year.astype("Int64").astype(str))
    data["Work Order Status"] = data["Work Order Status"].fillna("Open / Not Classified")
    data["Type"] = data["Type"].fillna("Not Classified")
    data["Class"] = data["Class"].fillna("Not Classified")
    data["Project"] = data["Project"].fillna("Not Classified")
    data["Subclass"] = data["Subclass"].fillna("Not Classified")
    data["Stage"] = data["Stage"].fillna("Not Classified")
    data["Region"] = data["Region"].fillna("Not Classified")
    data["City"] = data["City"].fillna("Not Classified")
    data["District"] = data["District"].fillna("Not Classified")

    penalties["Assigned Date"] = pd.to_datetime(penalties.get("Assigned Date"), errors="coerce")
    penalties["Number"] = pd.to_numeric(penalties.get("Number"), errors="coerce").fillna(0)
    penalties["Penalties Amount"] = pd.to_numeric(penalties.get("Penalties Amount"), errors="coerce").fillna(0)

    return data, penalties, snapshot_date

def apply_theme_css(theme_mode: str):
    dark = theme_mode == "Dark"
    bg = "#081426" if dark else "#F5F7FB"
    card = "#0f1f38" if dark else "#FFFFFF"
    text = "#EAF2FF" if dark else "#182230"
    muted = "#8BA2C8" if dark else "#667085"
    border = "rgba(148,163,184,0.15)" if dark else "rgba(15,23,42,0.08)"
    accent = "#3B82F6"
    glow = "0 10px 30px rgba(59,130,246,0.15)" if dark else "0 8px 24px rgba(15,23,42,0.06)"
    st.markdown(f"""
    <style>
        .stApp {{
            background:
                radial-gradient(circle at top left, rgba(59,130,246,0.16), transparent 28%),
                radial-gradient(circle at top right, rgba(249,115,22,0.12), transparent 24%),
                {bg};
            color: {text};
        }}
        div[data-testid="stMetric"] {{
            background: {card};
            border: 1px solid {border};
            border-radius: 18px;
            padding: 14px 16px;
            box-shadow: {glow};
        }}
        .block-container {{
            padding-top: 1.4rem;
            padding-bottom: 2rem;
        }}
        .top-banner {{
            background: linear-gradient(135deg, rgba(59,130,246,0.18), rgba(249,115,22,0.12));
            border: 1px solid {border};
            border-radius: 24px;
            padding: 22px 24px;
            margin-bottom: 12px;
            box-shadow: {glow};
        }}
        .subtle {{
            color: {muted};
            font-size: 0.95rem;
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
        .stDataFrame, div[data-testid="stTable"] {{
            border-radius: 16px;
            overflow: hidden;
            border: 1px solid {border};
        }}
    </style>
    """, unsafe_allow_html=True)

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

def fmt_num(x):
    return "-" if pd.isna(x) else f"{x:,.0f}"

def color_template(theme_mode):
    return "plotly_dark" if theme_mode == "Dark" else "plotly_white"

def metric_delta(value, reference, kind="pct_points"):
    if pd.isna(value) or pd.isna(reference):
        return None
    diff = value - reference
    if kind == "ratio":
        return f"{diff:+.2f}"
    if kind == "money":
        return fmt_money(diff)
    return f"{diff:+.1f} pp"

st.sidebar.title("Control Panel")
theme_mode = st.sidebar.radio("Theme", ["Dark", "Light"], horizontal=True, index=0)
apply_theme_css(theme_mode)

uploaded_file = st.sidebar.file_uploader("Upload refreshed Dawiyat workbook", type=["xlsx"])
data, penalties, snapshot_date = load_data(
    file_bytes=uploaded_file if uploaded_file is not None else None,
    file_name=getattr(uploaded_file, "name", None),
)

st.markdown(f"""
<div class="top-banner">
    <h1 style="margin:0 0 6px 0;">Dawiyat Project Intelligence Dashboard</h1>
    <div class="subtle">
        Decision-support dashboard for PMO, Operations, and Performance Management.
        Snapshot date: <b>{snapshot_date.strftime("%d %b %Y %H:%M")}</b>
    </div>
</div>
""", unsafe_allow_html=True)

# Filters
def filter_widget(label, series):
    vals = sorted([v for v in pd.Series(series).dropna().astype(str).unique().tolist() if v != "nan"])
    return st.sidebar.multiselect(label, vals)

region_sel = filter_widget("Region", data["Region"])
city_sel = filter_widget("City", data["City"])
district_sel = filter_widget("District", data["District"])
stage_sel = filter_widget("Stage", data["Stage"])
year_sel = filter_widget("Year", data["Year"])
status_sel = filter_widget("Work Order Status", data["Work Order Status"])
type_sel = filter_widget("Type", data["Type"])
class_sel = filter_widget("Class", data["Class"])
project_sel = filter_widget("Project", data["Project"])
subclass_sel = filter_widget("Subclass", data["Subclass"])

filtered = data.copy()
for col, selection in [
    ("Region", region_sel), ("City", city_sel), ("District", district_sel), ("Stage", stage_sel),
    ("Year", year_sel), ("Work Order Status", status_sel), ("Type", type_sel), ("Class", class_sel),
    ("Project", project_sel), ("Subclass", subclass_sel)
]:
    if selection:
        filtered = filtered[filtered[col].astype(str).isin(selection)]

if filtered.empty:
    st.warning("No records match the selected filters.")
    st.stop()

# KPI calculations
total_wo = filtered["Work Order"].nunique()
total_link = filtered["Link Code"].nunique()
avg_actual = filtered["actual_progress_capped"].mean()
avg_planned = filtered["planned_progress_pct"].mean()
on_track_pct = (filtered["schedule_variance_pp"] >= 0).mean() * 100
overdue_cnt = int(filtered["is_overdue"].sum())
critical_lag_cnt = int(filtered["critical_lag"].sum())
forecast_high_risk = int((filtered["forecast_risk"] == "High delay risk").sum())
budget_total = filtered["WO Cost"].fillna(0).sum()
actual_cost_total = filtered["Cost"].fillna(0).sum()
ev_total = filtered["EV"].fillna(0).sum()
pv_total = filtered["PV"].fillna(0).sum()
spi_total = ev_total / pv_total if pv_total > 0 else np.nan
cpi_total = ev_total / actual_cost_total if actual_cost_total > 0 else np.nan
penalty_cases = int(filtered["penalty_cases"].sum())
penalty_qty = float(filtered["penalty_qty"].sum())
penalty_amount = float(filtered["penalty_amount"].sum())
avg_civil = filtered.loc[filtered["Subclass"].str.contains("Civil", case=False, na=False), "actual_progress_capped"].mean()
avg_fiber = filtered.loc[filtered["Subclass"].str.contains("Fiber", case=False, na=False), "actual_progress_capped"].mean()

tabs = st.tabs(["Executive Overview", "Schedule & KPI", "Penalties & Quality", "Work Order Detail", "Dashboard Guide"])

with tabs[0]:
    c1, c2, c3, c4, c5, c6 = st.columns(6)
    c1.metric("Work Orders", f"{total_wo:,}", f"{total_link:,} link codes")
    c2.metric("Avg Actual Progress", fmt_pct(avg_actual), metric_delta(avg_actual, avg_planned))
    c3.metric("Avg Planned Progress", fmt_pct(avg_planned))
    c4.metric("SPI", f"{spi_total:,.2f}" if pd.notna(spi_total) else "-", "Healthy" if pd.notna(spi_total) and spi_total >= 1 else "Below plan")
    c5.metric("CPI", f"{cpi_total:,.2f}" if pd.notna(cpi_total) else "-", "Healthy" if pd.notna(cpi_total) and cpi_total >= 1 else "Cost pressure")
    c6.metric("Penalty Cases", f"{penalty_cases:,}", f"Qty {penalty_qty:,.0f}")

    c7, c8, c9, c10 = st.columns(4)
    c7.metric("On Track", fmt_pct(on_track_pct))
    c8.metric("Overdue WO", f"{overdue_cnt:,}")
    c9.metric("Critical Lag >15pp", f"{critical_lag_cnt:,}")
    c10.metric("High Forecast Risk", f"{forecast_high_risk:,}")

    chart_col1, chart_col2 = st.columns([1.25, 1])
    with chart_col1:
        by_month = (
            filtered.dropna(subset=["Updated"])
            .assign(month=filtered["Updated"].dt.to_period("M").astype(str))
            .groupby("month")
            .agg(
                actual=("actual_progress_capped", "mean"),
                planned=("planned_progress_pct", "mean"),
                spi=("SPI", "mean"),
                cpi=("CPI", "mean"),
            ).reset_index()
        )
        if not by_month.empty:
            fig = go.Figure()
            fig.add_trace(go.Scatter(x=by_month["month"], y=by_month["actual"], mode="lines+markers", name="Actual %"))
            fig.add_trace(go.Scatter(x=by_month["month"], y=by_month["planned"], mode="lines+markers", name="Planned %"))
            fig.update_layout(
                title="Planned vs Actual Progress Trend",
                template=color_template(theme_mode),
                height=400,
                legend_orientation="h",
                margin=dict(l=20, r=20, t=50, b=20),
                yaxis_title="Progress %",
                xaxis_title=""
            )
            st.plotly_chart(fig, use_container_width=True)
    with chart_col2:
        city_perf = (
            filtered.groupby("City")
            .agg(
                actual=("actual_progress_capped", "mean"),
                planned=("planned_progress_pct", "mean"),
                overdue=("is_overdue", "sum"),
                penalties=("penalty_cases", "sum"),
            ).reset_index().sort_values("actual", ascending=True)
        )
        fig = px.bar(
            city_perf,
            x="actual",
            y="City",
            orientation="h",
            color="overdue",
            text="actual",
            title="City Performance Snapshot",
            template=color_template(theme_mode),
        )
        fig.update_traces(texttemplate="%{text:.1f}%", textposition="outside")
        fig.update_layout(height=400, margin=dict(l=10, r=10, t=50, b=10), xaxis_title="Avg Actual Progress %", yaxis_title="")
        st.plotly_chart(fig, use_container_width=True)

    chart_col3, chart_col4 = st.columns([1, 1])
    with chart_col3:
        status_df = pd.DataFrame({
            "Metric": ["Civil Completion", "Fiber Completion", "On Track", "Overdue Exposure"],
            "Value": [
                0 if pd.isna(avg_civil) else avg_civil,
                0 if pd.isna(avg_fiber) else avg_fiber,
                on_track_pct,
                min((overdue_cnt / max(total_wo, 1)) * 100, 100),
            ]
        })
        fig = px.bar_polar(status_df, r="Value", theta="Metric", template=color_template(theme_mode), title="Execution Health Mix")
        fig.update_layout(height=380, margin=dict(l=10, r=10, t=55, b=10))
        st.plotly_chart(fig, use_container_width=True)
    with chart_col4:
        stage_mix = filtered.groupby("Stage")["Work Order"].count().reset_index(name="count").sort_values("count", ascending=False)
        fig = px.treemap(stage_mix, path=["Stage"], values="count", color="count", title="Stage Distribution", template=color_template(theme_mode))
        fig.update_layout(height=380, margin=dict(l=10, r=10, t=50, b=10))
        st.plotly_chart(fig, use_container_width=True)

with tabs[1]:
    left, right = st.columns([1.15, 1])
    with left:
        district_summary = (
            filtered.groupby(["City", "District"], dropna=False)
            .agg(
                work_orders=("Work Order", "count"),
                actual=("actual_progress_capped", "mean"),
                planned=("planned_progress_pct", "mean"),
                spi=("SPI", "mean"),
                cpi=("CPI", "mean"),
                overdue=("is_overdue", "sum"),
                critical_lag=("critical_lag", "sum"),
                forecast_delay=("forecast_delay_days", "mean"),
            ).reset_index()
        )
        district_summary["variance_pp"] = district_summary["actual"] - district_summary["planned"]
        district_summary = district_summary.sort_values(["City", "variance_pp"])
        fig = px.bar(
            district_summary,
            x="District",
            y="variance_pp",
            color="City",
            title="District Schedule Variance (Actual - Planned)",
            template=color_template(theme_mode),
            hover_data=["actual", "planned", "spi", "cpi", "overdue", "critical_lag", "forecast_delay"]
        )
        fig.add_hline(y=0, line_dash="dash")
        fig.update_layout(height=440, margin=dict(l=10, r=10, t=55, b=10), yaxis_title="Variance pp")
        st.plotly_chart(fig, use_container_width=True)
    with right:
        risk = (
            filtered.groupby("forecast_risk")
            .agg(count=("Work Order", "count"))
            .reset_index()
            .sort_values("count", ascending=False)
        )
        fig = px.pie(risk, names="forecast_risk", values="count", hole=0.58, title="Forecast Delay Risk Profile", template=color_template(theme_mode))
        fig.update_layout(height=440, margin=dict(l=10, r=10, t=55, b=10))
        st.plotly_chart(fig, use_container_width=True)

    bottom1, bottom2 = st.columns([1, 1])
    with bottom1:
        top_delay = (
            filtered[filtered["forecast_delay_days"].notna()]
            .nlargest(15, "forecast_delay_days")[
                ["Work Order", "Link Code", "City", "District", "Subclass", "Stage", "actual_progress_capped", "planned_progress_pct", "forecast_delay_days", "effective_target", "forecast_completion_date"]
            ]
            .rename(columns={
                "actual_progress_capped": "Actual %",
                "planned_progress_pct": "Planned %",
                "forecast_delay_days": "Forecast Delay Days",
                "effective_target": "Target Date",
                "forecast_completion_date": "Forecast Finish",
            })
        )
        st.markdown("#### Top Forecast Delay Exposure")
        st.dataframe(top_delay, use_container_width=True, hide_index=True)
    with bottom2:
        milestone_cols = [
            ("Trench Progress", "Trench"),
            ("MH/HH Progress", "MH/HH"),
            ("Fiber Progress", "Fiber"),
            ("ODBs Progress", "ODBs"),
            ("ODFs Progress", "ODFs"),
            ("JCL Progress", "JCL"),
            ("FAT Progress", "FAT"),
            ("PFAT Progress", "PFAT"),
            ("SFAT Progress", "SFAT"),
            ("Permits Progress", "Permits"),
        ]
        milestone_data = []
        for col, label in milestone_cols:
            if col in filtered.columns:
                val = pd.to_numeric(filtered[col], errors="coerce").mean()
                if pd.notna(val):
                    milestone_data.append({"Milestone": label, "Avg Progress": val})
        if milestone_data:
            ms = pd.DataFrame(milestone_data).sort_values("Avg Progress", ascending=True)
            fig = px.bar(ms, x="Avg Progress", y="Milestone", orientation="h", title="Milestone Progress Tracker", template=color_template(theme_mode), text="Avg Progress")
            fig.update_traces(texttemplate="%{text:.1f}%")
            fig.update_layout(height=420, margin=dict(l=10, r=10, t=55, b=10), xaxis_title="Progress %", yaxis_title="")
            st.plotly_chart(fig, use_container_width=True)

with tabs[2]:
    p1, p2 = st.columns([1, 1])
    with p1:
        if not penalties.empty:
            pen_city = (
                penalties.groupby("City", dropna=False)
                .agg(cases=("Impl # of Penalty", "count"), qty=("Number", "sum"), amount=("Penalties Amount", "sum"))
                .reset_index()
                .sort_values("qty", ascending=False)
            )
            fig = px.bar(pen_city, x="City", y="qty", color="cases", title="Penalty Quantity by City", template=color_template(theme_mode), text="qty")
            fig.update_layout(height=400, margin=dict(l=10, r=10, t=55, b=10), yaxis_title="Penalty Qty")
            st.plotly_chart(fig, use_container_width=True)
    with p2:
        if not penalties.empty:
            pen_type = (
                penalties.groupby("Impl # of Penalty", dropna=False)
                .agg(qty=("Number", "sum"))
                .reset_index()
                .sort_values("qty", ascending=False)
                .head(10)
            )
            fig = px.bar(pen_type, x="qty", y="Impl # of Penalty", orientation="h", title="Top Penalty Reasons", template=color_template(theme_mode), text="qty")
            fig.update_layout(height=400, margin=dict(l=10, r=10, t=55, b=10), xaxis_title="Penalty Qty", yaxis_title="")
            st.plotly_chart(fig, use_container_width=True)

    filtered_pen = penalties.copy()
    # align penalties filters where possible
    if region_sel:
        filtered_pen = filtered_pen[filtered_pen["Region"].astype(str).isin(region_sel)]
    if city_sel:
        filtered_pen = filtered_pen[filtered_pen["City"].astype(str).isin(city_sel)]

    st.markdown("#### Penalty Register")
    st.dataframe(filtered_pen, use_container_width=True, hide_index=True)

with tabs[3]:
    detail = filtered.copy()
    detail["Target Date"] = pd.to_datetime(detail["effective_target"]).dt.date
    detail["Forecast Finish"] = pd.to_datetime(detail["forecast_completion_date"]).dt.date
    detail["Actual %"] = detail["actual_progress_capped"].round(1)
    detail["Planned %"] = detail["planned_progress_pct"].round(1)
    detail["SPI"] = detail["SPI"].round(2)
    detail["CPI"] = detail["CPI"].round(2)
    detail["Forecast Delay Days"] = detail["forecast_delay_days"].round(0)
    show_cols = [
        "Link Code", "Work Order", "Region", "City", "District", "Project", "Subclass", "Stage",
        "Work Order Status", "Type", "Class", "Target Date", "Forecast Finish",
        "Actual %", "Planned %", "SPI", "CPI", "WO Cost", "Cost", "penalty_cases", "penalty_qty", "forecast_risk"
    ]
    st.dataframe(detail[show_cols].sort_values(["Region", "City", "District", "Target Date"], na_position="last"), use_container_width=True, hide_index=True)

with tabs[4]:
    st.markdown("""
    <div class="guide-box">
    <b>How this dashboard should be used</b><br>
    • <b>Main purpose:</b> top-management view for decision making, recovery planning, and district / city performance comparison.<br>
    • <b>Actual progress:</b> taken from <i>Percentage of Completion</i> in the service tool.<br>
    • <b>Effective target date:</b> uses <i>Updated Target Date</i> first; if blank, falls back to <i>Targeted Completion</i>.<br>
    • <b>Planned progress:</b> estimated from elapsed time between work-order start date and effective target date.<br>
    • <b>SPI / CPI:</b> dashboard uses an earned-value style approximation from available sheet fields. This is useful for management direction, but it is not a replacement for a full cost-control system.<br>
    • <b>Critical lag:</b> work orders where actual progress is more than 15 percentage points behind planned progress.<br>
    • <b>Forecast delay:</b> estimated finish date based on current achievement speed versus remaining target date.<br>
    • <b>Penalty metrics:</b> shown from the Penalties tab and linked at link-code level where possible.
    </div>
    """, unsafe_allow_html=True)

    st.markdown("""
    #### Refresh Process
    1. Export the newest Dawiyat workbook with the same sheet structure.<br>
    2. Open this dashboard and upload the refreshed file from the sidebar.<br>
    3. All charts and KPIs will recalculate automatically.<br>
    4. Use filters to focus on Region, City, District, Stage, Year, Type, Class, Project, or Subclass.
    """, unsafe_allow_html=True)

    summary = pd.DataFrame([
        ["Total Work Orders", total_wo],
        ["Total Link Codes", total_link],
        ["Avg Actual Progress %", round(avg_actual, 1) if pd.notna(avg_actual) else np.nan],
        ["Avg Planned Progress %", round(avg_planned, 1) if pd.notna(avg_planned) else np.nan],
        ["SPI", round(spi_total, 2) if pd.notna(spi_total) else np.nan],
        ["CPI", round(cpi_total, 2) if pd.notna(cpi_total) else np.nan],
        ["Overdue Work Orders", overdue_cnt],
        ["Critical Lag >15pp", critical_lag_cnt],
        ["Penalty Cases", penalty_cases],
        ["Penalty Quantity", round(penalty_qty, 0)],
    ], columns=["Metric", "Value"])
    st.markdown("#### Current Filter Summary")
    st.table(summary)
