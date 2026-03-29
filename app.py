
from __future__ import annotations
from io import BytesIO
from pathlib import Path
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(page_title="Dawiyat Executive Dashboard", page_icon="📊", layout="wide", initial_sidebar_state="expanded")

DEFAULT_FILE = "Dawiyat Master Sheet.xlsx"

def clean_cols(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df

def find_sheet(xls: pd.ExcelFile, candidates: list[str]) -> str | None:
    normalized = {str(name).strip().lower(): name for name in xls.sheet_names}
    for c in candidates:
        if c.strip().lower() in normalized:
            return normalized[c.strip().lower()]
    for name in xls.sheet_names:
        low = str(name).strip().lower()
        if any(c.strip().lower() in low for c in candidates):
            return name
    return None

def first_col(df: pd.DataFrame, candidates: list[str]) -> str | None:
    norm = {str(c).strip().lower(): c for c in df.columns}
    for c in candidates:
        if c.strip().lower() in norm:
            return norm[c.strip().lower()]
    for c in df.columns:
        low = str(c).strip().lower()
        if any(x.strip().lower() in low for x in candidates):
            return c
    return None

def series_or_default(df: pd.DataFrame, col: str, default=np.nan):
    if col in df.columns:
        return df[col]
    return pd.Series([default] * len(df), index=df.index)

def to_num(s):
    return pd.to_numeric(s, errors="coerce")

def pct_ratio(progress, scope):
    p = to_num(progress)
    s = to_num(scope)
    out = np.where(s > 0, (p / s) * 100, np.nan)
    return pd.Series(out, index=p.index).clip(lower=0, upper=150)

def effective_date(df: pd.DataFrame):
    upd = pd.to_datetime(series_or_default(df, "Updated Target Date"), errors="coerce")
    tgt = pd.to_datetime(series_or_default(df, "Targeted Completion"), errors="coerce")
    return upd.combine_first(tgt)

def map_region(raw_region):
    txt = "" if pd.isna(raw_region) else str(raw_region).strip().upper()
    if txt in {"MAKKAH", "MECCA", "TABUK", "BAHA", "MADINAH", "MEDINA", "WESTERN"}:
        return "Western"
    if txt in {"JIZAN", "NAJRAN", "ABHA", "KHAMIS", "BISHA", "SOUTHERN"}:
        return "Southern"
    if txt in {"EASTERN REGION", "EASTERN", "DAMMAM", "KHOBAR", "JUBAIL", "HASA"}:
        return "Eastern"
    if txt in {"NORTHERN", "AL JOUF", "ARAR", "HAIL"}:
        return "Northern"
    return "Not Classified"

def risk_bucket(row):
    pct = row.get("Percentage of Completion", np.nan)
    eff = row.get("Effective Target Date", pd.NaT)
    if pd.isna(eff):
        return "No target date"
    days = (pd.Timestamp.today().normalize() - pd.Timestamp(eff).normalize()).days
    if pd.notna(pct) and pct >= 100:
        return "Completed"
    if days > 30:
        return "High delay risk"
    if days > 0:
        return "Moderate delay risk"
    if days >= -14:
        return "Watchlist"
    return "On forecast"

def fmt_money(x):
    if pd.isna(x):
        return "-"
    x = float(x)
    if abs(x) >= 1_000_000:
        return f"SAR {x/1_000_000:.2f}M"
    if abs(x) >= 1_000:
        return f"SAR {x/1_000:.1f}K"
    return f"SAR {x:.0f}"

def fmt_pct(x):
    return "-" if pd.isna(x) else f"{float(x):.1f}%"

def add_card(title, value, subtitle=""):
    st.markdown(f"""
    <div class="metric-card">
      <div class="metric-title">{title}</div>
      <div class="metric-value">{value}</div>
      <div class="metric-subtitle">{subtitle}</div>
    </div>
    """, unsafe_allow_html=True)

def to_excel_bytes(df, sheet_name="Export"):
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name[:31])
    bio.seek(0)
    return bio.getvalue()

def style(theme: str):
    dark = theme == "Dark"
    bg = "#07111f" if dark else "#F2F6FB"
    panel = "#0c1c33" if dark else "#FFFFFF"
    card_bg = "linear-gradient(180deg, rgba(14,31,57,0.98), rgba(9,20,39,0.98))" if dark else "linear-gradient(180deg, #1B2E4D, #142742)"
    txt = "#EAF2FF" if dark else "#16253B"
    muted = "#9CB0CD" if dark else "#8FA1BC"
    border = "rgba(148,163,184,0.18)" if dark else "rgba(15,23,42,0.08)"
    body_card = "#0D1D35" if dark else "#FFFFFF"
    st.markdown(f"""
    <style>
    .stApp {{
      background:
        radial-gradient(circle at top left, rgba(59,130,246,0.18), transparent 28%),
        radial-gradient(circle at top right, rgba(249,115,22,0.10), transparent 18%),
        {bg};
      color: {txt};
    }}
    .block-container {{padding-top: 1rem; max-width: 95rem;}}
    section[data-testid="stSidebar"] {{
      background: linear-gradient(180deg, #13284A 0%, #0A1628 100%);
      border-right: 1px solid rgba(148,163,184,0.16);
    }}
    section[data-testid="stSidebar"] * {{color:#ECF3FF !important;}}
    .hero {{
      background: linear-gradient(135deg, rgba(37,99,235,0.18), rgba(16,185,129,0.08), rgba(249,115,22,0.12));
      border: 1px solid {border};
      border-radius: 24px;
      padding: 20px 24px;
      margin-bottom: 10px;
      box-shadow: 0 10px 24px rgba(15,23,42,0.08);
    }}
    .metric-card {{
      background: {card_bg};
      border: 1px solid {border};
      border-radius: 20px;
      padding: 16px 18px;
      min-height: 130px;
      box-shadow: 0 12px 24px rgba(15,23,42,0.08);
    }}
    .metric-title {{font-size: .95rem; color: #A8BADA; margin-bottom: 8px;}}
    .metric-value {{font-size: 1.95rem; font-weight: 800; color: #FFFFFF; line-height: 1.1;}}
    .metric-subtitle {{font-size: .90rem; color: #A8BADA; margin-top: 8px;}}
    .panel {{
      background: {body_card};
      border: 1px solid {border};
      border-radius: 22px;
      padding: 14px 16px 8px 16px;
      box-shadow: 0 12px 24px rgba(15,23,42,0.06);
    }}
    .guide-box {{
      background: {panel};
      border: 1px solid {border};
      border-left: 4px solid #3B82F6;
      border-radius: 16px;
      padding: 12px 14px;
      margin-bottom: 10px;
    }}
    .small-note {{color:{muted}; font-size:0.92rem;}}
    div[data-testid="stMetric"] {{
      background: {body_card};
      border: 1px solid {border};
      border-radius: 16px;
      padding: 8px 12px;
    }}
    .stTabs [data-baseweb="tab-list"] {{gap: 12px;}}
    .stTabs [data-baseweb="tab"] {{
      height: 46px;
      white-space: pre-wrap;
      border-radius: 10px 10px 0 0;
      padding-left: 10px; padding-right: 10px;
    }}
    </style>
    """, unsafe_allow_html=True)

@st.cache_data(show_spinner=False)
def load_data(file_source):
    xls = pd.ExcelFile(file_source)
    main_name = find_sheet(xls, ["Dawaiyat Service Tool", "service tool"])
    if main_name is None:
        raise ValueError("Main sheet Dawaiyat Service Tool was not found.")
    main = clean_cols(pd.read_excel(file_source, sheet_name=main_name))

    district_name = find_sheet(xls, ["District"])
    penalties_name = find_sheet(xls, ["Penalties", "Penalty"])
    details_name = find_sheet(xls, ["Workorder details", "Workorder detail"])

    district = clean_cols(pd.read_excel(file_source, sheet_name=district_name)) if district_name else pd.DataFrame()
    penalties = clean_cols(pd.read_excel(file_source, sheet_name=penalties_name)) if penalties_name else pd.DataFrame()
    details = clean_cols(pd.read_excel(file_source, sheet_name=details_name)) if details_name else main.copy()

    # Use district sheet for city / district. Use mapped region from main raw region.
    warnings = []
    if not district.empty:
        d_link = first_col(district, ["Link Code"])
        d_city = first_col(district, ["City"])
        d_district = first_col(district, ["District"])
        if d_link is None:
            warnings.append("District sheet is missing Link Code.")
        else:
            mapping = district.rename(columns={
                d_link: "Link Code",
                d_city: "City" if d_city else "City",
                d_district: "District" if d_district else "District"
            })
            keep_cols = ["Link Code"] + [c for c in ["City", "District"] if c in mapping.columns]
            mapping = mapping[keep_cols].drop_duplicates()
            raw_region_map = main.groupby("Link Code", dropna=False)["Region"].first().map(map_region).reset_index(name="Macro Region")
            mapping = mapping.merge(raw_region_map, on="Link Code", how="left")
            main = main.drop(columns=[c for c in ["Macro Region", "City", "District"] if c in main.columns], errors="ignore").merge(mapping, on="Link Code", how="left")
            details = details.drop(columns=[c for c in ["Macro Region", "City", "District"] if c in details.columns], errors="ignore").merge(mapping, on="Link Code", how="left")
    else:
        main["Macro Region"] = main["Region"].map(map_region)
        details["Macro Region"] = details["Region"].map(map_region)
        warnings.append("District sheet not found; City and District may be incomplete.")

    for df in (main, details):
        df["Effective Target Date"] = effective_date(df)
        for dcol in ["Created", "Assigned at", "In Progress at", "Targeted Completion", "Updated Target Date", "Updated", "Closed at", "Effective Target Date"]:
            if dcol in df.columns:
                df[dcol] = pd.to_datetime(df[dcol], errors="coerce")
        for c in ["Percentage of Completion","WO Cost","Cost","Trench Progress","Trench Scope","Fiber Progress","Fiber Scope","Updates",
                  "PIP rejection count","PAT rejection count","Approval rejection count","As-Built Rejection Count","Handover Rejection Count"]:
            if c in df.columns:
                df[c] = to_num(df[c])
        df["Civil Completion %"] = pct_ratio(series_or_default(df, "Trench Progress"), series_or_default(df, "Trench Scope"))
        df["Fiber Completion %"] = pct_ratio(series_or_default(df, "Fiber Progress"), series_or_default(df, "Fiber Scope"))
        df["Lag %"] = np.where(df["Percentage of Completion"].fillna(0) < 100,
                               (100 - df["Percentage of Completion"].fillna(0)).clip(lower=0),
                               0)
        today = pd.Timestamp.today().normalize()
        df["Overdue"] = (df["Effective Target Date"].notna()) & (df["Effective Target Date"] < today) & (df["Percentage of Completion"].fillna(0) < 100)
        df["Forecast Risk"] = df.apply(risk_bucket, axis=1)
        updated_dates = pd.to_datetime(series_or_default(df, "Updated"), errors="coerce")
        update_counts = to_num(series_or_default(df, "Updates")).fillna(0)
        df["Update Follow-up"] = (update_counts < 5) & ((today - updated_dates).dt.days > 5) & (df["Percentage of Completion"].fillna(0) < 100)

    # penalties prep
    if not penalties.empty:
        penalties = penalties.rename(columns={first_col(penalties, ["Cluster Name", "Link Code", "Cluster"]): "Link Code"})
        dev_col = first_col(penalties, ["Deviation name", "Impl # of Penalty", "Penalty", "Deviation"])
        num_col = first_col(penalties, ["Number of Deviations", "Number"])
        amt_col = first_col(penalties, ["Penalties Amount", "Penalty Amount"])
        if dev_col and dev_col != "Deviation name":
            penalties = penalties.rename(columns={dev_col: "Deviation name"})
        if num_col and num_col != "Number of Deviations":
            penalties = penalties.rename(columns={num_col: "Number of Deviations"})
        if amt_col and amt_col != "Penalties Amount":
            penalties = penalties.rename(columns={amt_col: "Penalties Amount"})
        for c in ["Number of Deviations", "Penalties Amount"]:
            if c in penalties.columns:
                penalties[c] = to_num(penalties[c]).fillna(0)
            else:
                penalties[c] = 0
        if "Link Code" in penalties.columns:
            link_mapping = main[["Link Code", "Macro Region", "City", "District"]].drop_duplicates()
            penalties = penalties.merge(link_mapping, on="Link Code", how="left", suffixes=("","_map"))
            if "City_map" in penalties.columns:
                penalties["City"] = penalties["City"].combine_first(penalties["City_map"]) if "City" in penalties.columns else penalties["City_map"]
            if "District_map" in penalties.columns:
                penalties["District"] = penalties["District"].combine_first(penalties["District_map"]) if "District" in penalties.columns else penalties["District_map"]
            if "Macro Region_map" in penalties.columns:
                penalties["Macro Region"] = penalties["Macro Region"].combine_first(penalties["Macro Region_map"]) if "Macro Region" in penalties.columns else penalties["Macro Region_map"]
        penalties["Deviation name"] = series_or_default(penalties, "Deviation name", "").fillna("")
    else:
        penalties = pd.DataFrame(columns=["Link Code","Deviation name","Number of Deviations","Penalties Amount","Macro Region","City","District"])

    if not penalties.empty and "Link Code" in penalties.columns:
        p_group = penalties.groupby("Link Code", dropna=False).agg(
            **{
                "Number of Deviations": ("Number of Deviations", "sum"),
                "Penalty Amount": ("Penalties Amount", "sum"),
                "Main Deviation": ("Deviation name", lambda s: " | ".join(pd.Series(s).dropna().astype(str).drop_duplicates().head(3).tolist()))
            }
        ).reset_index()
    else:
        p_group = pd.DataFrame(columns=["Link Code","Number of Deviations","Penalty Amount","Main Deviation"])

    details = details.merge(p_group, on="Link Code", how="left")
    details["Number of Deviations"] = to_num(series_or_default(details, "Number of Deviations")).fillna(0).astype(int)
    details["Penalty Amount"] = to_num(series_or_default(details, "Penalty Amount")).fillna(0)

    snapshot_date = pd.to_datetime(series_or_default(main, "Updated"), errors="coerce").max()
    return main, penalties, details, snapshot_date, warnings

st.sidebar.markdown("## Control Panel")
theme_mode = st.sidebar.radio("Theme", ["Dark", "Light"], horizontal=True)
style(theme_mode)

uploaded = st.sidebar.file_uploader("Upload refreshed Dawiyat workbook", type=["xlsx"])
source = uploaded if uploaded is not None else Path(__file__).with_name(DEFAULT_FILE)
data, penalties, details, snapshot_date, data_warnings = load_data(source)

def options_from(series):
    vals = pd.Series(series).dropna().astype(str).str.strip()
    vals = vals[vals != ""]
    return sorted(vals.unique().tolist())

filtered = data.copy()
region_opts = ["All"] + options_from(filtered.get("Macro Region", pd.Series(dtype=str)))
region_sel = st.sidebar.selectbox("Region", region_opts)
if region_sel != "All":
    filtered = filtered[filtered["Macro Region"].astype(str) == region_sel]

city_opts = ["All"] + options_from(filtered.get("City", pd.Series(dtype=str)))
city_sel = st.sidebar.selectbox("City", city_opts)
if city_sel != "All":
    filtered = filtered[filtered["City"].astype(str) == city_sel]

district_opts = ["All"] + options_from(filtered.get("District", pd.Series(dtype=str)))
district_sel = st.sidebar.selectbox("District", district_opts)
if district_sel != "All":
    filtered = filtered[filtered["District"].astype(str) == district_sel]

project_opts = ["All"] + options_from(filtered.get("Project", pd.Series(dtype=str)))
project_sel = st.sidebar.selectbox("Project", project_opts)
if project_sel != "All":
    filtered = filtered[filtered["Project"].astype(str) == project_sel]

subclass_opts = ["All"] + options_from(filtered.get("Subclass", pd.Series(dtype=str)))
subclass_sel = st.sidebar.selectbox("Subclass", subclass_opts)
if subclass_sel != "All":
    filtered = filtered[filtered["Subclass"].astype(str) == subclass_sel]

status_opts = ["All"] + options_from(filtered.get("Work Order Status", pd.Series(dtype=str)))
status_sel = st.sidebar.selectbox("Work Order Status", status_opts)
if status_sel != "All":
    filtered = filtered[filtered["Work Order Status"].astype(str) == status_sel]

link_opts = ["All"] + options_from(filtered.get("Link Code", pd.Series(dtype=str)))
link_sel = st.sidebar.selectbox("Link Code", link_opts)
if link_sel != "All":
    filtered = filtered[filtered["Link Code"].astype(str) == link_sel]

filtered_details = details[details["Link Code"].isin(filtered["Link Code"].unique())].copy()
filtered_pen = penalties[penalties["Link Code"].isin(filtered["Link Code"].unique())].copy() if not penalties.empty else penalties.copy()

st.markdown(f"""
<div class="hero">
  <div style="font-size:1.0rem;color:#9CB0CD;">DAWIYAT PROJECT</div>
  <div style="font-size:2.0rem;font-weight:800;">Executive Project Intelligence Dashboard</div>
  <div class="small-note">Snapshot date: {snapshot_date.strftime("%d-%b-%Y %H:%M") if pd.notna(snapshot_date) else "Not available"} | Filters apply across overview, PMO, KPI, penalties, and work order detail pages.</div>
</div>
""", unsafe_allow_html=True)

if data_warnings:
    with st.expander("Data mapping warnings"):
        for w in data_warnings:
            st.warning(w)

tab_overview, tab_pmo, tab_kpi, tab_pen, tab_detail, tab_guide = st.tabs(
    ["Executive Overview", "PMO Summary", "Schedule & KPI", "Penalties & Quality", "Work Order Detail", "Dashboard Guide"]
)

today = pd.Timestamp.today().normalize()
total_links = filtered["Link Code"].nunique()
avg_progress = filtered.groupby("Link Code")["Percentage of Completion"].mean().mean() if not filtered.empty else 0
avg_civil = filtered["Civil Completion %"].mean()
avg_fiber = filtered["Fiber Completion %"].mean()
overdue_df = filtered[filtered["Overdue"]].copy()
critical_df = filtered[filtered["Lag %"] >= 15].copy()
cancelled_count = int((filtered["Work Order Status"].astype(str).str.upper() == "CANCELLED").sum()) if "Work Order Status" in filtered.columns else 0
in_progress_count = int((filtered["Work Order Status"].astype(str).str.upper() == "IN PROGRESS").sum()) if "Work Order Status" in filtered.columns else 0

with tab_overview:
    c1, c2, c3, c4, c5 = st.columns(5)
    with c1: add_card("Total Link Codes", f"{int(total_links)}", "Distinct link codes after filters")
    with c2: add_card("Avg Overall Progress", fmt_pct(avg_progress), "Mean % completion")
    with c3: add_card("Avg Civil Completion", fmt_pct(avg_civil), "Trench progress / scope")
    with c4: add_card("Avg Fiber Completion", fmt_pct(avg_fiber), "Fiber progress / scope")
    with c5: add_card("Overdue Work Orders", f"{len(overdue_df)}", "Past effective target and not complete")

    left, right = st.columns([1.3, 1])
    with left:
        st.markdown('<div class="panel">', unsafe_allow_html=True)
        city_perf = filtered.groupby(["City"], dropna=False)["Percentage of Completion"].mean().reset_index()
        city_perf["City"] = city_perf["City"].fillna("Unmapped")
        fig = px.bar(city_perf.sort_values("Percentage of Completion", ascending=False),
                     x="City", y="Percentage of Completion",
                     text_auto=".1f", title="City Average Progress")
        fig.update_layout(height=360, xaxis_title="", yaxis_title="Progress %")
        st.plotly_chart(fig, use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)

        st.markdown('<div class="panel">', unsafe_allow_html=True)
        stage_perf = filtered.groupby("Stage", dropna=False)["Link Code"].nunique().reset_index(name="Link Codes")
        if not stage_perf.empty:
            fig2 = px.bar(stage_perf.sort_values("Link Codes", ascending=False), x="Link Codes", y="Stage", orientation="h",
                          title="Stage Distribution by Link Code")
            fig2.update_layout(height=360, yaxis_title="", xaxis_title="Link Codes")
            st.plotly_chart(fig2, use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)

    with right:
        st.markdown('<div class="panel">', unsafe_allow_html=True)
        district_lag = filtered.groupby("District", dropna=False)["Lag %"].mean().reset_index().sort_values("Lag %", ascending=False).head(12)
        district_lag["District"] = district_lag["District"].fillna("Unmapped")
        fig3 = px.bar(district_lag, x="Lag %", y="District", orientation="h", title="District Lag Ranking", text_auto=".1f")
        fig3.update_layout(height=360, yaxis_title="", xaxis_title="Lag %")
        st.plotly_chart(fig3, use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)

        st.markdown('<div class="panel">', unsafe_allow_html=True)
        status_counts = filtered.groupby("Work Order Status", dropna=False)["Link Code"].count().reset_index(name="Count")
        status_counts["Work Order Status"] = status_counts["Work Order Status"].fillna("Unknown")
        fig4 = px.pie(status_counts, names="Work Order Status", values="Count", hole=0.55, title="Work Order Status Mix")
        fig4.update_layout(height=360)
        st.plotly_chart(fig4, use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)

with tab_pmo:
    flagged = filtered.groupby("Link Code", dropna=False).agg(
        Region=("Macro Region", lambda s: s.dropna().astype(str).mode().iloc[0] if not s.dropna().empty else np.nan),
        City=("City", lambda s: s.dropna().astype(str).mode().iloc[0] if not s.dropna().empty else np.nan),
        District=("District", lambda s: s.dropna().astype(str).mode().iloc[0] if not s.dropna().empty else np.nan),
        Avg_Progress=("Percentage of Completion", "mean"),
        Latest_Update=("Updated", "max"),
        Max_Updates=("Updates", "max"),
        Work_Orders=("Work Order", "nunique"),
        Effective_Target=("Effective Target Date", "max"),
        Total_Cost=("Cost", "sum")
    ).reset_index()
    flagged["Needs System Update"] = (flagged["Max_Updates"].fillna(0) < 5) & ((today - flagged["Latest_Update"]).dt.days > 5) & (flagged["Avg_Progress"].fillna(0) < 100)
    followup = flagged[flagged["Needs System Update"]].sort_values(["Latest_Update", "Avg_Progress"])

    best_link = flagged[(flagged["Avg_Progress"] >= 100) & ((flagged["Effective_Target"].isna()) | (flagged["Effective_Target"] >= today))]
    best_link = best_link.sort_values(["Avg_Progress", "Latest_Update"], ascending=[False, False]).head(15)

    cost_base = filtered.copy()
    cost_base["Original Target Month"] = pd.to_datetime(cost_base["Targeted Completion"], errors="coerce").dt.to_period("M").astype("string")
    cost_base["Effective Target Month"] = pd.to_datetime(cost_base["Effective Target Date"], errors="coerce").dt.to_period("M").astype("string")
    monthly_original = cost_base.groupby("Original Target Month", dropna=False)["Cost"].sum().reset_index(name="Original Target Cost")
    monthly_effective = cost_base.groupby("Effective Target Month", dropna=False)["Cost"].sum().reset_index(name="Updated / Final Target Cost")
    monthly = monthly_original.merge(monthly_effective, left_on="Original Target Month", right_on="Effective Target Month", how="outer")
    monthly["Month"] = monthly["Original Target Month"].combine_first(monthly["Effective Target Month"])
    monthly = monthly[monthly["Month"].notna()].sort_values("Month")

    a,b,c,d = st.columns(4)
    with a: add_card("Link Codes Need Update", f"{int(len(followup))}", "Updates < 5 and last update > 5 days")
    with b: add_card("Cancelled Work Orders", f"{cancelled_count}", "From Work Order Status")
    with c: add_card("In Progress Work Orders", f"{in_progress_count}", "From Work Order Status")
    with d: add_card("Best Completed Link Codes", f"{int(len(best_link))}", "Completed with no delay exposure")

    left, right = st.columns([1.2, 1])
    with left:
        st.markdown('<div class="panel">', unsafe_allow_html=True)
        st.subheader("PMO Follow-up Link Codes")
        st.caption("Top-management action list. Criteria: Updates < 5, last system update older than 5 days, and average progress below 100%.")
        view = followup[["Link Code","Region","City","District","Avg_Progress","Max_Updates","Latest_Update","Work_Orders"]].copy()
        view = view.rename(columns={"Avg_Progress":"Avg Progress %","Max_Updates":"Updates","Latest_Update":"Last Updated","Work_Orders":"Work Orders"})
        st.dataframe(view, use_container_width=True, hide_index=True)
        st.download_button("Export PMO follow-up (Excel)", data=to_excel_bytes(view, "PMO FollowUp"), file_name="pmo_followup_linkcodes.xlsx")
        st.markdown("</div>", unsafe_allow_html=True)

        st.markdown('<div class="panel">', unsafe_allow_html=True)
        st.subheader("Monthly Target Cost Analysis")
        st.caption("Original target month uses Targeted Completion. Final target month uses Updated Target Date when available; otherwise it keeps the original target date.")
        if not monthly.empty:
            fig = go.Figure()
            fig.add_trace(go.Bar(x=monthly["Month"], y=monthly["Original Target Cost"], name="Original Target Cost"))
            fig.add_trace(go.Scatter(x=monthly["Month"], y=monthly["Updated / Final Target Cost"], name="Final Target Cost", mode="lines+markers"))
            fig.update_layout(height=360, xaxis_title="Month", yaxis_title="Cost (SAR)")
            st.plotly_chart(fig, use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)

    with right:
        st.markdown('<div class="panel">', unsafe_allow_html=True)
        st.subheader("Best Link Codes Delivered On / Before Plan")
        st.caption("Completed link codes with no delay exposure based on effective target date.")
        show = best_link[["Link Code","Region","City","District","Avg_Progress","Latest_Update"]].copy()
        show = show.rename(columns={"Avg_Progress":"Avg Progress %","Latest_Update":"Last Updated"})
        st.dataframe(show, use_container_width=True, hide_index=True)
        st.markdown("</div>", unsafe_allow_html=True)

        st.markdown('<div class="panel">', unsafe_allow_html=True)
        city_link = flagged.groupby(["Region","City"], dropna=False).size().reset_index(name="Link Codes")
        city_link["City"] = city_link["City"].fillna("Unmapped")
        fig = px.bar(city_link, x="City", y="Link Codes", color="Region", title="Link Code Distribution by Region / City", barmode="group")
        fig.update_layout(height=360, xaxis_title="", yaxis_title="Link Codes")
        st.plotly_chart(fig, use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)

with tab_kpi:
    rej_cols = [c for c in ["PIP rejection count","PAT rejection count","Approval rejection count","As-Built Rejection Count","Handover Rejection Count"] if c in filtered.columns]
    total_rejections = int(filtered[rej_cols].fillna(0).sum().sum()) if rej_cols else 0

    a,b,c,d,e = st.columns(5)
    with a: add_card("Avg Civil Completion", fmt_pct(avg_civil), "Trench progress / scope")
    with b: add_card("Avg Fiber Completion", fmt_pct(avg_fiber), "Fiber progress / scope")
    with c: add_card("Overdue Work Orders", f"{len(overdue_df)}", "Past effective target and not complete")
    with d: add_card("Critical Lag", f"{len(critical_df)}", "Lag >= 15%")
    with e: add_card("Total Rejections", f"{total_rejections}", "Sum of PIP, PAT, Approval, As-Built, and Handover rejections")

    st.markdown('<div class="panel">', unsafe_allow_html=True)
    st.subheader("Overdue Work Orders - Action List")
    risk_options = sorted(overdue_df["Forecast Risk"].dropna().unique().tolist()) if not overdue_df.empty else []
    risk_filter = st.multiselect("Filter overdue list by forecast risk", options=risk_options, default=risk_options)
    overdue_view = overdue_df[overdue_df["Forecast Risk"].isin(risk_filter)] if risk_filter else overdue_df.iloc[0:0]
    overdue_view = overdue_view[[c for c in ["Link Code","Work Order","Macro Region","City","District","Subclass","Stage","Percentage of Completion","Targeted Completion","Updated Target Date","Effective Target Date","Updated","Updates","Lag %","Forecast Risk"] if c in overdue_view.columns]].copy()
    overdue_view = overdue_view.rename(columns={"Macro Region":"Region"})
    st.dataframe(overdue_view.sort_values(["Effective Target Date","Lag %"], ascending=[True, False]), use_container_width=True, hide_index=True)
    c1,c2 = st.columns(2)
    with c1:
        st.download_button("Export overdue work orders (Excel)", data=to_excel_bytes(overdue_view, "Overdue Work Orders"), file_name="overdue_workorders.xlsx")
    with c2:
        st.download_button("Export overdue work orders (CSV)", data=overdue_view.to_csv(index=False).encode("utf-8-sig"), file_name="overdue_workorders.csv")
    st.markdown("</div>", unsafe_allow_html=True)

    left, right = st.columns([1.1, 1])
    with left:
        st.markdown('<div class="panel">', unsafe_allow_html=True)
        lag_by_district = filtered.groupby("District", dropna=False)["Lag %"].mean().reset_index().sort_values("Lag %", ascending=False)
        lag_by_district["District"] = lag_by_district["District"].fillna("Unmapped")
        fig = px.bar(lag_by_district.head(15), x="Lag %", y="District", orientation="h", title="District Lag Ranking", text_auto=".1f")
        fig.update_layout(height=400)
        st.plotly_chart(fig, use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)
    with right:
        st.markdown('<div class="panel">', unsafe_allow_html=True)
        progress_trend = filtered.copy().dropna(subset=["Effective Target Date"])
        if not progress_trend.empty:
            monthly_prog = progress_trend.groupby(progress_trend["Effective Target Date"].dt.to_period("M").astype(str))["Percentage of Completion"].mean().reset_index()
            monthly_prog.columns = ["Month", "Avg Progress %"]
            fig = px.line(monthly_prog, x="Month", y="Avg Progress %", markers=True, title="Average Progress by Effective Target Month")
            fig.update_layout(height=400)
            st.plotly_chart(fig, use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)

with tab_pen:
    dev_total = int(filtered_pen["Number of Deviations"].sum()) if not filtered_pen.empty and "Number of Deviations" in filtered_pen.columns else 0
    pen_total = float(filtered_pen["Penalties Amount"].sum()) if not filtered_pen.empty and "Penalties Amount" in filtered_pen.columns else 0

    a,b = st.columns(2)
    with a: add_card("Number of Deviations", f"{dev_total}", "Sum of deviation counts")
    with b: add_card("Penalty Deduction Amount", fmt_money(pen_total), "Sum of deducted amount only")

    left, right = st.columns([1.2, 1])
    with left:
        st.markdown('<div class="panel">', unsafe_allow_html=True)
        st.subheader("Top Deviations vs Penalty Amount")
        if not filtered_pen.empty:
            dev_chart = filtered_pen.groupby("Deviation name", dropna=False).agg(
                Number_of_Deviations=("Number of Deviations","sum"),
                Penalty_Amount=("Penalties Amount","sum")
            ).reset_index().sort_values(["Number_of_Deviations","Penalty_Amount"], ascending=False).head(10)
            fig = go.Figure()
            fig.add_trace(go.Bar(y=dev_chart["Deviation name"], x=dev_chart["Number_of_Deviations"], orientation="h", name="Number of Deviations"))
            fig.add_trace(go.Scatter(y=dev_chart["Deviation name"], x=dev_chart["Penalty_Amount"], mode="markers", marker=dict(size=12), name="Penalty Amount"))
            fig.update_layout(height=420, yaxis_title="", xaxis_title="Count / SAR")
            st.plotly_chart(fig, use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)

        st.markdown('<div class="panel">', unsafe_allow_html=True)
        st.subheader("Deviation Analysis Table")
        if not filtered_pen.empty:
            dev_tbl = filtered_pen.groupby(["Deviation name"], dropna=False).agg(
                **{"Number of Deviations":("Number of Deviations","sum"),
                   "Penalty Amount":("Penalties Amount","sum")}
            ).reset_index().sort_values(["Number of Deviations","Penalty Amount"], ascending=False)
            st.dataframe(dev_tbl, use_container_width=True, hide_index=True)
        st.markdown("</div>", unsafe_allow_html=True)

    with right:
        st.markdown('<div class="panel">', unsafe_allow_html=True)
        st.subheader("Top Quality / Penalty Exposure")
        if not filtered_pen.empty:
            exposure = filtered_pen.groupby("Link Code", dropna=False).agg(
                **{"Number of Deviations":("Number of Deviations","sum"),
                   "Deduction Amount":("Penalties Amount","sum")}
            ).reset_index()
            lag_link = filtered.groupby("Link Code", dropna=False)["Lag %"].mean().reset_index()
            exposure = exposure.merge(lag_link, on="Link Code", how="left")
            exposure = exposure.sort_values(["Deduction Amount","Number of Deviations"], ascending=False)
            st.dataframe(exposure, use_container_width=True, hide_index=True)
        st.markdown("</div>", unsafe_allow_html=True)

with tab_detail:
    st.subheader("Work Order Detail")
    detail_cols = [c for c in [
        "Link Code","Work Order","Macro Region","City","District","Project","Subclass","Stage","Work Order Status","Type","Class",
        "Percentage of Completion","Civil Completion %","Fiber Completion %","Cost","Created","Assigned at","Targeted Completion","Updated Target Date",
        "Effective Target Date","Updated","Updates","Target Area","Notes","Scope of Work","Number of Deviations","Penalty Amount","Main Deviation","Forecast Risk"
    ] if c in filtered_details.columns]
    show = filtered_details[detail_cols].copy()
    show = show.rename(columns={"Macro Region":"Region","Penalty Amount":"Penalty Amount"})
    st.dataframe(show, use_container_width=True, hide_index=True)
    c1, c2 = st.columns(2)
    with c1:
        st.download_button("Export work order detail (Excel)", data=to_excel_bytes(show, "WorkOrder Detail"), file_name="workorder_detail.xlsx")
    with c2:
        st.download_button("Export work order detail (CSV)", data=show.to_csv(index=False).encode("utf-8-sig"), file_name="workorder_detail.csv")

with tab_guide:
    st.markdown('<div class="guide-box"><div style="font-weight:700;">Dashboard Guide</div><div class="small-note">This page explains how the main metrics are calculated in the dashboard.</div></div>', unsafe_allow_html=True)
    st.markdown('<div class="guide-box"><b>Region</b><br>Region is mapped to macro regions (Western, Southern, Eastern, Northern) using the raw Region value from the service-tool sheet. City and District come from the District sheet.</div>', unsafe_allow_html=True)
    st.markdown('<div class="guide-box"><b>Effective Target Date</b><br>Updated Target Date is used when available. Otherwise, Targeted Completion is used. This is the final date used for overdue and forecast-risk logic.</div>', unsafe_allow_html=True)
    st.markdown('<div class="guide-box"><b>Lag %</b><br>Lag % = max(0, 100 - Percentage of Completion). It shows the remaining completion gap for incomplete work orders.</div>', unsafe_allow_html=True)
    st.markdown('<div class="guide-box"><b>Overdue Work Orders</b><br>A work order is overdue when Effective Target Date is before today and Percentage of Completion is below 100%.</div>', unsafe_allow_html=True)
    st.markdown('<div class="guide-box"><b>Total Rejections</b><br>Total Rejections = PIP rejection count + PAT rejection count + Approval rejection count + As-Built Rejection Count + Handover Rejection Count.</div>', unsafe_allow_html=True)
    st.markdown('<div class="guide-box"><b>Number of Deviations</b><br>This is the sum of deviation counts from the Penalties tab. It is different from deducted penalty amount because not all deviations have penalties applied.</div>', unsafe_allow_html=True)
    st.markdown('<div class="guide-box"><b>Penalty Deduction Amount</b><br>Penalty Deduction Amount is the sum of deducted values only from the Penalties tab.</div>', unsafe_allow_html=True)
    st.markdown('<div class="guide-box"><b>PMO Follow-up Link Codes</b><br>A link code appears in the PMO follow-up list when Updates is below 5, the latest system update is older than 5 days, and progress is still below 100%.</div>', unsafe_allow_html=True)
