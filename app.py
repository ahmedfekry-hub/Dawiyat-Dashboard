from __future__ import annotations

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

# ---------- Helpers ----------
def clean_text(v):
    if pd.isna(v):
        return np.nan
    t = str(v).strip()
    if t.lower() in {"", "nan", "none", "null"}:
        return np.nan
    return t


def find_sheet_name(xls: pd.ExcelFile, candidates: list[str]) -> str:
    normalized = {str(n).strip().lower(): n for n in xls.sheet_names}
    for c in candidates:
        key = c.strip().lower()
        if key in normalized:
            return normalized[key]
    for n in xls.sheet_names:
        low = str(n).strip().lower()
        if any(c.strip().lower() in low for c in candidates):
            return n
    raise ValueError(f"Could not find sheet for {candidates}")


def first_existing(df: pd.DataFrame, names: list[str]) -> str | None:
    norm = {str(c).strip().lower(): c for c in df.columns}
    for n in names:
        key = n.strip().lower()
        if key in norm:
            return norm[key]
    for c in df.columns:
        low = str(c).strip().lower()
        if any(n.strip().lower() in low for n in names):
            return c
    return None


def choose_mode(s: pd.Series):
    s = s.dropna()
    if s.empty:
        return np.nan
    m = s.mode(dropna=True)
    return m.iloc[0] if not m.empty else s.iloc[0]


def ratio_pct(progress: pd.Series, scope: pd.Series) -> pd.Series:
    p = pd.to_numeric(progress, errors="coerce")
    s = pd.to_numeric(scope, errors="coerce")
    with np.errstate(divide="ignore", invalid="ignore"):
        out = np.where(s > 0, (p / s) * 100, np.nan)
    return pd.Series(out, index=progress.index).clip(lower=0, upper=100)


def fmt_pct(x):
    return "-" if pd.isna(x) else f"{x:,.1f}%"


def fmt_money(x):
    if pd.isna(x):
        return "-"
    x = float(x)
    if abs(x) >= 1_000_000:
        return f"${x/1_000_000:,.2f}M"
    if abs(x) >= 1_000:
        return f"${x/1_000:,.1f}K"
    return f"${x:,.0f}"


# ---------- Theme ----------
def apply_theme_css(theme_mode: str):
    dark = theme_mode == "Dark"
    bg = "#08111f" if dark else "#f5f7fb"
    card = "#0d1b2d" if dark else "#ffffff"
    card2 = "#10233b" if dark else "#ffffff"
    text = "#f3f7ff" if dark else "#1b2430"
    muted = "#9fb3d1" if dark else "#677489"
    border = "rgba(111,155,255,0.16)" if dark else "rgba(15,23,42,0.08)"
    glow = "0 10px 30px rgba(0,0,0,0.28), 0 0 24px rgba(45,130,255,0.12)" if dark else "0 10px 24px rgba(16,24,40,.08)"

    st.markdown(
        f"""
        <style>
        .stApp {{
            background:
                radial-gradient(circle at 12% 12%, rgba(32,145,249,.13), transparent 24%),
                radial-gradient(circle at 88% 10%, rgba(255,153,0,.08), transparent 22%),
                radial-gradient(circle at 65% 90%, rgba(43,215,175,.06), transparent 18%),
                {bg};
            color: {text};
        }}
        section[data-testid="stSidebar"] {{
            background: linear-gradient(180deg, #07111f, #0a1425) !important;
            border-right: 1px solid rgba(111,155,255,.14);
        }}
        section[data-testid="stSidebar"] * {{ color: #eaf2ff !important; }}
        .block-container {{ padding-top: 1.1rem; padding-bottom: 2rem; }}
        div[data-testid="stMetric"] {{
            background: linear-gradient(180deg, {card}, {card2});
            border: 1px solid {border};
            border-radius: 20px;
            padding: 14px 16px;
            box-shadow: {glow};
        }}
        div[data-testid="stMetricLabel"] {{ color: {muted}; }}
        div[data-testid="stMetricValue"] {{ color: {text}; }}
        .top-banner {{
            background: linear-gradient(135deg, rgba(20,44,82,.96), rgba(8,17,31,.96));
            border: 1px solid {border};
            border-radius: 24px;
            padding: 22px 24px;
            margin-bottom: 14px;
            box-shadow: {glow};
        }}
        .top-title {{ font-size: 2.0rem; font-weight: 800; letter-spacing: .02em; margin: 0; }}
        .top-title span {{ color: #46a3ff; }}
        .subtle {{ color: {muted}; font-size: .95rem; }}
        .section-card {{
            background: linear-gradient(180deg, {card}, {card2});
            border: 1px solid {border};
            border-radius: 22px;
            padding: 12px 14px 4px 14px;
            box-shadow: {glow};
            margin-bottom: 12px;
        }}
        .warn-box, .ok-box {{
            border-radius: 16px;
            padding: 14px 16px;
            border: 1px solid {border};
            margin: 8px 0 14px 0;
        }}
        .warn-box {{ background: rgba(255,153,0,.10); }}
        .ok-box {{ background: rgba(43,215,175,.10); }}
        div.stButton > button, .stDownloadButton > button {{
            border-radius: 12px; border: 1px solid {border};
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )


def template(theme_mode: str) -> str:
    return "plotly_dark" if theme_mode == "Dark" else "plotly_white"


# ---------- Load ----------
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

    for df in (main, district, penalties):
        for c in df.columns:
            if df[c].dtype == object:
                df[c] = df[c].map(clean_text)

    main = main[main["Link Code"].notna()].copy()
    main["Link Code"] = main["Link Code"].astype(str).str.strip()

    # District mapping
    d_link = first_existing(district, ["Link Code"])
    d_city = first_existing(district, ["City"])
    d_dist = first_existing(district, ["District"])
    if d_link is None:
        district_map = pd.DataFrame(columns=["Link Code", "City", "District"])
    else:
        district = district[district[d_link].notna()].copy()
        district[d_link] = district[d_link].astype(str).str.strip()
        district["City"] = district[d_city].map(clean_text) if d_city else np.nan
        district["District"] = district[d_dist].map(clean_text) if d_dist else np.nan
        district_map = (
            district.groupby(d_link, dropna=False)
            .agg(City=("City", choose_mode), District=("District", choose_mode))
            .reset_index()
            .rename(columns={d_link: "Link Code"})
        )

    data = main.merge(district_map, on="Link Code", how="left")

    # Heuristic fallback from Link Code when District sheet is incomplete
    parts = data["Link Code"].astype(str).str.split("-")
    data["link_prefix"] = parts.str[0].str.upper()
    data["link_district_code"] = parts.str[2].str.upper()

    city_by_prefix = (
        district_map.assign(link_prefix=district_map["Link Code"].astype(str).str.split("-").str[0].str.upper())
        .dropna(subset=["City"])
        .groupby("link_prefix")["City"].agg(choose_mode).to_dict()
    )
    city_by_prefix.update({
        "JED": "Jeddah", "TAI": "Taif", "TUR": "Taif", "MOY": "Taif",
        "SAM": "Jizan", "SHU": "Jizan", "EDA": "Jizan", "DAI": "Jizan", "SHA": "Jizan",
        "TAB": "Tabuk", "BAH": "Al Baha", "MAK": "Makkah", "JUB": "Jubail"
    })

    district_by_code = (
        district_map.assign(link_district_code=district_map["Link Code"].astype(str).str.split("-").str[2].str.upper())
        .dropna(subset=["District"])
        .groupby("link_district_code")["District"].agg(choose_mode).to_dict()
    )

    data["City"] = data["City"].fillna(data["link_prefix"].map(city_by_prefix))
    data["District"] = data["District"].fillna(data["link_district_code"].map(district_by_code))
    data["District"] = data["District"].fillna(data["link_district_code"].where(data["link_district_code"].notna() & (data["link_district_code"] != "NAN")))

    # Data quality warnings
    warnings = []
    if data["City"].isna().all():
        warnings.append("City mapping could not be derived from the District sheet.")
    if data["District"].isna().all():
        warnings.append("District mapping could not be derived from the District sheet.")

    # Penalties: map cluster name to link code
    p_cluster = first_existing(penalties, ["Cluster Name", "Cluster", "Link Code"])
    p_dev = first_existing(penalties, ["Number of Deviations", "Deviation Count", "Number"])
    p_amt = first_existing(penalties, ["Penalties Amount", "Penalty Amount", "Amount"])
    penalties["Link Code"] = penalties[p_cluster].astype(str).str.strip() if p_cluster else np.nan
    penalties["Number of Deviations"] = pd.to_numeric(penalties[p_dev], errors="coerce").fillna(0) if p_dev else 0
    penalties["Penalties Amount"] = pd.to_numeric(penalties[p_amt], errors="coerce").fillna(0) if p_amt else 0
    p_name = first_existing(penalties, ["Deviation name", "Deviation"])
    if p_name is None:
        penalties["Deviation name"] = "Penalty"
    penalties_agg = (
        penalties.groupby("Link Code", dropna=False)
        .agg(
            penalty_cases=("Deviation name", "count"),
            penalty_qty=("Number of Deviations", "sum"),
            penalty_amount=("Penalties Amount", "sum"),
        )
        .reset_index()
    )
    data = data.merge(penalties_agg, on="Link Code", how="left")
    for c in ["penalty_cases", "penalty_qty", "penalty_amount"]:
        data[c] = pd.to_numeric(data[c], errors="coerce").fillna(0)

    # Dates
    for c in ["Created", "Assigned at", "In Progress at", "Updated", "Closed at", "Targeted Completion", "Updated Target Date"]:
        if c in data.columns:
            data[c] = pd.to_datetime(data[c], errors="coerce")

    snapshot_date = pd.to_datetime(data.get("Updated"), errors="coerce").max()
    if pd.isna(snapshot_date):
        snapshot_date = pd.Timestamp.today().normalize()

    # Text defaults
    for c, default in {
        "Region": "Not Classified",
        "Project": "Not Classified",
        "Subclass": "Not Classified",
        "Stage": "Not Classified",
        "Work Order Status": "Open / Not Classified",
        "Type": "Not Classified",
        "Class": "Not Classified",
        "City": "Unknown",
        "District": "Unknown",
    }.items():
        if c not in data.columns:
            data[c] = default
        else:
            data[c] = data[c].fillna(default)

    if "Year" not in data.columns:
        data["Year"] = np.nan
    data["Year"] = data["Year"].astype("string").replace({"<NA>": np.nan}).fillna(data["Targeted Completion"].dt.year.astype("Int64").astype("string")).fillna("Not Classified")

    # Progress logic
    data["actual_progress_pct"] = pd.to_numeric(data.get("Percentage of Completion"), errors="coerce").clip(lower=0, upper=100)
    data["effective_target"] = data.get("Updated Target Date").combine_first(data.get("Targeted Completion"))
    data["start_date"] = data.get("In Progress at").combine_first(data.get("Assigned at")).combine_first(data.get("Created"))

    elapsed = (snapshot_date - data["start_date"]).dt.days
    total = (data["effective_target"] - data["start_date"]).dt.days
    with np.errstate(divide="ignore", invalid="ignore"):
        planned_pct = np.where(total > 0, np.clip((elapsed / total) * 100, 0, 100), np.nan)
    data["planned_progress_pct"] = planned_pct
    data["schedule_variance_pp"] = data["actual_progress_pct"] - data["planned_progress_pct"]
    data["is_complete"] = data["actual_progress_pct"] >= 100
    data["is_overdue"] = (snapshot_date > data["effective_target"]) & (~data["is_complete"]) & data["effective_target"].notna()
    data["critical_lag"] = data["schedule_variance_pp"] <= -15

    # EVM approximations
    data["WO Cost"] = pd.to_numeric(data.get("WO Cost"), errors="coerce").fillna(0)
    data["Cost"] = pd.to_numeric(data.get("Cost"), errors="coerce").fillna(0)
    data["EV"] = data["WO Cost"] * (data["actual_progress_pct"].fillna(0) / 100)
    data["PV"] = data["WO Cost"] * (data["planned_progress_pct"].fillna(0) / 100)
    data["AC"] = data["Cost"]
    data["SPI"] = np.where(data["PV"] > 0, data["EV"] / data["PV"], np.nan)
    data["CPI"] = np.where(data["AC"] > 0, data["EV"] / data["AC"], np.nan)
    data["EAC"] = np.where(data["CPI"] > 0, data["WO Cost"] / data["CPI"], np.nan)

    # Forecast completion
    actual_ratio = data["actual_progress_pct"] / 100
    elapsed_days = np.maximum((snapshot_date - data["start_date"]).dt.days, 1)
    est_total = np.where(actual_ratio > 0, elapsed_days / actual_ratio, np.nan)
    data["forecast_completion_date"] = data["start_date"] + pd.to_timedelta(est_total, unit="D")
    data["forecast_delay_days"] = (data["forecast_completion_date"] - data["effective_target"]).dt.days
    data["forecast_risk"] = np.select(
        [data["forecast_delay_days"] > 30, data["forecast_delay_days"] > 0, data["forecast_delay_days"] <= 0],
        ["High", "Medium", "Low"],
        default="Unknown",
    )

    # Milestone completion ratios
    data["civil_completion_pct"] = ratio_pct(data.get("Trench Progress", pd.Series(index=data.index)), data.get("Trench Scope", pd.Series(index=data.index)))
    civil2 = ratio_pct(data.get("MH/HH Progress", pd.Series(index=data.index)), data.get("MH/HH Scope", pd.Series(index=data.index)))
    civil3 = ratio_pct(data.get("Permits Progress", pd.Series(index=data.index)), data.get("Permits Scope", pd.Series(index=data.index)))
    data["civil_completion_pct"] = pd.concat([data["civil_completion_pct"], civil2, civil3], axis=1).mean(axis=1, skipna=True)
    data["fiber_completion_pct"] = ratio_pct(data.get("Fiber Progress", pd.Series(index=data.index)), data.get("Fiber Scope", pd.Series(index=data.index)))

    # Fallbacks by subclass if scope ratios missing
    civil_mask = data["Subclass"].astype(str).str.contains("civil", case=False, na=False)
    fiber_mask = data["Subclass"].astype(str).str.contains("fiber", case=False, na=False)
    data.loc[civil_mask & data["civil_completion_pct"].isna(), "civil_completion_pct"] = data.loc[civil_mask, "actual_progress_pct"]
    data.loc[fiber_mask & data["fiber_completion_pct"].isna(), "fiber_completion_pct"] = data.loc[fiber_mask, "actual_progress_pct"]

    # Penalties enrichment to city/district
    penalties = penalties.merge(district_map, on="Link Code", how="left")
    penalties = penalties.merge(data[["Link Code", "Region", "Project", "Subclass", "Stage", "Work Order Status"]].drop_duplicates("Link Code"), on="Link Code", how="left")
    penalties["City"] = penalties.get("City").fillna("Unknown")
    penalties["District"] = penalties.get("District").fillna("Unknown")

    # Data quality
    data_quality = {
        "missing_city": int(data["City"].isin(["Unknown", "Not Classified"]).sum()),
        "missing_district": int(data["District"].isin(["Unknown", "Not Classified"]).sum()),
        "missing_link": int(data["Link Code"].isna().sum()),
    }
    total_rows = max(len(data), 1)
    data_quality["score"] = round(100 - ((data_quality["missing_city"] + data_quality["missing_district"] + data_quality["missing_link"]) / (3 * total_rows) * 100), 1)

    return data, penalties, snapshot_date, warnings, data_quality


# ---------- UI ----------
st.sidebar.markdown("## Dawiyat PMO")
theme_mode = st.sidebar.radio("Theme", ["Dark", "Light"], horizontal=True, index=0)
apply_theme_css(theme_mode)
source = st.sidebar.file_uploader("Upload refreshed workbook", type=["xlsx"])

data, penalties, snapshot_date, warnings, dq = load_data(source if source is not None else None)

st.markdown(
    f"""
    <div class="top-banner">
        <p class="top-title">EXECUTIVE <span>PROJECT INTELLIGENCE</span> DASHBOARD</p>
        <div class="subtle">Dawiyat Project | PMO / Operations / Performance view | Snapshot: <b>{snapshot_date.strftime('%d %b %Y')}</b></div>
    </div>
    """,
    unsafe_allow_html=True,
)

if warnings:
    st.markdown("<div class='warn-box'>⚠️ " + " | ".join(warnings) + "</div>", unsafe_allow_html=True)

# Cascading filters
filtered = data.copy()

def choose_one(label, df, column):
    vals = sorted(v for v in df[column].dropna().astype(str).unique().tolist() if v != "nan")
    return st.sidebar.selectbox(label, ["All"] + vals)

region = choose_one("Region", filtered, "Region")
if region != "All":
    filtered = filtered[filtered["Region"] == region]
city = choose_one("City", filtered, "City")
if city != "All":
    filtered = filtered[filtered["City"] == city]
district = choose_one("District", filtered, "District")
if district != "All":
    filtered = filtered[filtered["District"] == district]
project = choose_one("Project", filtered, "Project")
if project != "All":
    filtered = filtered[filtered["Project"] == project]
stage = choose_one("Stage", filtered, "Stage")
if stage != "All":
    filtered = filtered[filtered["Stage"] == stage]
year = choose_one("Year", filtered, "Year")
if year != "All":
    filtered = filtered[filtered["Year"] == year]
status = choose_one("Work Order Status", filtered, "Work Order Status")
if status != "All":
    filtered = filtered[filtered["Work Order Status"] == status]
type_sel = choose_one("Type", filtered, "Type")
if type_sel != "All":
    filtered = filtered[filtered["Type"] == type_sel]
klass = choose_one("Class", filtered, "Class")
if klass != "All":
    filtered = filtered[filtered["Class"] == klass]
subclass = choose_one("Subclass", filtered, "Subclass")
if subclass != "All":
    filtered = filtered[filtered["Subclass"] == subclass]
link = choose_one("Link Code", filtered, "Link Code")
if link != "All":
    filtered = filtered[filtered["Link Code"] == link]

pen_filtered = penalties.copy()
for col, sel in [("Region", region), ("City", city), ("District", district), ("Project", project), ("Subclass", subclass), ("Stage", stage), ("Work Order Status", status), ("Link Code", link)]:
    if sel != "All" and col in pen_filtered.columns:
        pen_filtered = pen_filtered[pen_filtered[col].astype(str) == sel]

nav = st.sidebar.radio(
    "Navigation",
    ["Overview", "Performance", "Risks & Penalties", "Data Quality"],
    index=0,
)

if filtered.empty:
    st.warning("No records match the selected filters.")
    st.stop()

# KPIs
wo_count = filtered["Work Order"].nunique() if "Work Order" in filtered.columns else len(filtered)
link_count = filtered["Link Code"].nunique()
avg_actual = filtered["actual_progress_pct"].mean()
avg_planned = filtered["planned_progress_pct"].mean()
civil_pct = filtered["civil_completion_pct"].mean()
fiber_pct = filtered["fiber_completion_pct"].mean()
spi = filtered["EV"].sum() / filtered["PV"].sum() if filtered["PV"].sum() > 0 else np.nan
cpi = filtered["EV"].sum() / filtered["AC"].sum() if filtered["AC"].sum() > 0 else np.nan
eac = filtered["EAC"].mean()
on_track_pct = (filtered["schedule_variance_pp"] >= 0).mean() * 100
overdue_cnt = int(filtered["is_overdue"].sum())
critical_lag_cnt = int(filtered["critical_lag"].sum())
high_risk = int((filtered["forecast_risk"] == "High").sum())
penalty_cases = int(filtered["penalty_cases"].sum())
penalty_qty = float(filtered["penalty_qty"].sum())
penalty_amount = float(filtered["penalty_amount"].sum())

plot_bg = "rgba(0,0,0,0)"
grid = "rgba(160,185,220,.15)" if theme_mode == "Dark" else "rgba(15,23,42,.08)"
font_color = "#EAF2FF" if theme_mode == "Dark" else "#1B2430"
accent_blue = "#42A5FF"
accent_orange = "#FFB54A"
accent_green = "#31E7A8"
accent_red = "#FF6B6B"


def style_fig(fig, height=350):
    fig.update_layout(
        template=template(theme_mode),
        height=height,
        margin=dict(l=20, r=20, t=50, b=20),
        paper_bgcolor=plot_bg,
        plot_bgcolor=plot_bg,
        font=dict(color=font_color),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
    )
    fig.update_xaxes(showgrid=True, gridcolor=grid)
    fig.update_yaxes(showgrid=True, gridcolor=grid)
    return fig

if nav == "Overview":
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("PROJECT HEALTH", fmt_pct(on_track_pct), "On Track")
    c2.metric("BUDGET PERFORMANCE", fmt_money(filtered["WO Cost"].sum()), f"Actual {fmt_money(filtered['Cost'].sum())}")
    c3.metric("SCHEDULE PERFORMANCE", f"SPI {spi:.2f}" if pd.notna(spi) else "SPI -", f"Target {fmt_pct(avg_planned)}")
    c4.metric("CPI", f"{cpi:.2f}" if pd.notna(cpi) else "-", "+ healthy" if pd.notna(cpi) and cpi >= 1 else "cost pressure")
    c5.metric("EAC", fmt_money(eac), "Forecast")

    c6, c7, c8, c9 = st.columns(4)
    c6.metric("Total Link Codes", f"{link_count:,}")
    c7.metric("Avg Actual Progress", fmt_pct(avg_actual), f"Plan {fmt_pct(avg_planned)}")
    c8.metric("Civil Completion", fmt_pct(civil_pct))
    c9.metric("Fiber Completion", fmt_pct(fiber_pct))

    left, mid, right = st.columns([1.15, 1.15, 1])
    with left:
        monthly = (
            filtered.dropna(subset=["Updated"])
            .assign(Month=filtered["Updated"].dt.to_period("M").astype(str))
            .groupby("Month", as_index=False)
            .agg(CPI=("CPI", "mean"))
            .sort_values("Month")
        )
        if not monthly.empty:
            fig = go.Figure()
            fig.add_trace(go.Scatter(x=monthly["Month"], y=monthly["CPI"], mode="lines+markers", line=dict(color=accent_blue, width=3), name="CPI"))
            fig.add_hline(y=1, line_dash="dash", line_color=accent_green)
            fig.update_layout(title="CPI Trend", yaxis_title="CPI")
            st.plotly_chart(style_fig(fig, 330), use_container_width=True)
    with mid:
        monthly2 = (
            filtered.dropna(subset=["Updated"])
            .assign(Month=filtered["Updated"].dt.to_period("M").astype(str))
            .groupby("Month", as_index=False)
            .agg(SPI=("SPI", "mean"))
            .sort_values("Month")
        )
        if not monthly2.empty:
            fig = go.Figure()
            fig.add_trace(go.Scatter(x=monthly2["Month"], y=monthly2["SPI"], mode="lines+markers", line=dict(color=accent_orange, width=3), name="SPI"))
            fig.add_hline(y=1, line_dash="dash", line_color=accent_green)
            fig.update_layout(title="SPI Trend", yaxis_title="SPI")
            st.plotly_chart(style_fig(fig, 330), use_container_width=True)
    with right:
        city_perf = (
            filtered.groupby("City", as_index=False)
            .agg(Actual=("actual_progress_pct", "mean"), Planned=("planned_progress_pct", "mean"))
            .sort_values("Actual", ascending=False)
        )
        fig = go.Figure()
        fig.add_trace(go.Bar(x=city_perf["City"], y=city_perf["Actual"], name="Actual", marker_color=accent_blue))
        fig.add_trace(go.Scatter(x=city_perf["City"], y=city_perf["Planned"], mode="lines+markers", name="Forecast", line=dict(color=accent_orange, width=3)))
        fig.update_layout(title="Cashflow / Progress Performance", yaxis_title="%")
        st.plotly_chart(style_fig(fig, 330), use_container_width=True)

    left2, mid2, right2 = st.columns([1.1, 1.05, 1])
    with left2:
        summary = pd.DataFrame(
            {
                "Metric": ["Budget (BAC)", "Actual Cost", "Variance", "Planned Value (PV)", "Earned Value (EV)", "CPI", "SPI"],
                "Value": [
                    fmt_money(filtered["WO Cost"].sum()),
                    fmt_money(filtered["Cost"].sum()),
                    fmt_money(filtered["EV"].sum() - filtered["AC"].sum()),
                    fmt_money(filtered["PV"].sum()),
                    fmt_money(filtered["EV"].sum()),
                    "-" if pd.isna(cpi) else f"{cpi:.2f}",
                    "-" if pd.isna(spi) else f"{spi:.2f}",
                ],
            }
        )
        st.markdown("#### Cost & Schedule Summary")
        st.dataframe(summary, use_container_width=True, hide_index=True)
    with mid2:
        risk = pd.DataFrame(
            {
                "Risk": ["Financial", "Schedule", "Quality", "Penalties", "Forecast"],
                "Level": [
                    "High" if pd.notna(cpi) and cpi < 1 else "Low",
                    "High" if pd.notna(spi) and spi < 1 else "Low",
                    "Medium" if critical_lag_cnt > 0 else "Low",
                    "Medium" if penalty_cases > 0 else "Low",
                    "High" if high_risk > 0 else "Low",
                ],
                "Score": [85 if pd.notna(cpi) and cpi < 1 else 25, 85 if pd.notna(spi) and spi < 1 else 20, 60 if critical_lag_cnt > 0 else 20, 55 if penalty_cases > 0 else 20, 80 if high_risk > 0 else 20],
            }
        )
        fig = px.treemap(risk, path=["Level", "Risk"], values="Score", color="Score", color_continuous_scale=[[0,"#28d7ac"],[0.5,"#ffb54a"],[1,"#ff6b6b"]], title="Risk Exposure Heatmap")
        st.plotly_chart(style_fig(fig, 330), use_container_width=True)
    with right2:
        util = pd.DataFrame({"Resource": ["Civil", "Fiber", "Permits"], "Value": [civil_pct or 0, fiber_pct or 0, filtered["actual_progress_pct"].mean() or 0]})
        fig = px.pie(util, names="Resource", values="Value", hole=0.62, title="Resource Utilization")
        fig.update_traces(marker=dict(colors=[accent_blue, accent_orange, accent_green]))
        st.plotly_chart(style_fig(fig, 330), use_container_width=True)

    b1, b2, b3 = st.columns([1.1, 1.05, 1])
    with b1:
        ms = pd.DataFrame(
            {
                "Milestone": ["Trench", "MH/HH", "Fiber", "ODBs", "FAT"],
                "Completion": [
                    ratio_pct(filtered.get("Trench Progress", pd.Series(index=filtered.index)), filtered.get("Trench Scope", pd.Series(index=filtered.index))).mean(),
                    ratio_pct(filtered.get("MH/HH Progress", pd.Series(index=filtered.index)), filtered.get("MH/HH Scope", pd.Series(index=filtered.index))).mean(),
                    ratio_pct(filtered.get("Fiber Progress", pd.Series(index=filtered.index)), filtered.get("Fiber Scope", pd.Series(index=filtered.index))).mean(),
                    ratio_pct(filtered.get("ODBs Progress", pd.Series(index=filtered.index)), filtered.get("ODBs Scope", pd.Series(index=filtered.index))).mean(),
                    ratio_pct(filtered.get("FAT Progress", pd.Series(index=filtered.index)), filtered.get("FAT Scope", pd.Series(index=filtered.index))).mean(),
                ],
            }
        ).fillna(0)
        fig = px.bar(ms, x="Completion", y="Milestone", orientation="h", text="Completion", title="Project Milestones")
        fig.update_traces(marker_color=accent_orange, texttemplate="%{text:.1f}%")
        st.plotly_chart(style_fig(fig, 320), use_container_width=True)
    with b2:
        status_df = pd.DataFrame(
            {
                "Item": ["Time", "Cost", "Quality", "Risk"],
                "Status": [on_track_pct, (cpi * 100 if pd.notna(cpi) else 0), 100 - (penalty_cases / max(link_count, 1) * 100), 100 - (high_risk / max(link_count, 1) * 100)],
            }
        )
        fig = go.Figure(go.Indicator(mode="gauge+number", value=max(min(on_track_pct,100),0), title={"text":"Overall Status"}, gauge={"axis":{"range":[0,100]},"bar":{"color":accent_orange},"steps":[{"range":[0,50],"color":"#4b1f28"},{"range":[50,80],"color":"#5a4622"},{"range":[80,100],"color":"#174e43"}] }))
        st.plotly_chart(style_fig(fig, 320), use_container_width=True)
    with b3:
        days_remaining = int((filtered["effective_target"].max() - snapshot_date).days) if filtered["effective_target"].notna().any() else 0
        st.markdown("#### Overall Status")
        st.metric("Projected Completion", filtered["effective_target"].max().strftime("%b %Y") if filtered["effective_target"].notna().any() else "-", f"{days_remaining} days remaining")
        st.metric("Overdue Tasks", f"{overdue_cnt:,}")
        st.metric("High Risk Activities", f"{high_risk:,}")

elif nav == "Performance":
    left, right = st.columns([1.15, 1])
    with left:
        by_month = (
            filtered.dropna(subset=["effective_target"])
            .assign(Month=filtered["effective_target"].dt.to_period("M").astype(str))
            .groupby("Month", as_index=False)
            .agg(Actual=("actual_progress_pct", "mean"), Planned=("planned_progress_pct", "mean"))
            .sort_values("Month")
        )
        fig = go.Figure()
        fig.add_trace(go.Scatter(x=by_month["Month"], y=by_month["Actual"], mode="lines+markers", name="Actual", line=dict(color=accent_blue, width=3)))
        fig.add_trace(go.Scatter(x=by_month["Month"], y=by_month["Planned"], mode="lines+markers", name="Planned", line=dict(color=accent_orange, width=3)))
        fig.update_layout(title="Planned vs Actual Progress Trend", yaxis_title="Progress %")
        st.plotly_chart(style_fig(fig, 380), use_container_width=True)
    with right:
        delay = filtered[["Work Order", "Link Code", "City", "District", "forecast_delay_days"]].dropna().sort_values("forecast_delay_days", ascending=False).head(12)
        fig = px.bar(delay, x="Link Code", y="forecast_delay_days", color="City", title="Top Forecast Delay Exposure")
        st.plotly_chart(style_fig(fig, 380), use_container_width=True)

    d1, d2 = st.columns(2)
    with d1:
        district_sv = (
            filtered.groupby(["City", "District"], as_index=False)
            .agg(Actual=("actual_progress_pct", "mean"), Planned=("planned_progress_pct", "mean"), SPI=("SPI", "mean"), CPI=("CPI", "mean"))
        )
        district_sv["Variance"] = district_sv["Actual"] - district_sv["Planned"]
        fig = px.bar(district_sv, x="District", y="Variance", color="City", title="District Schedule Variance")
        fig.add_hline(y=0, line_dash="dash", line_color=accent_green)
        st.plotly_chart(style_fig(fig, 380), use_container_width=True)
    with d2:
        risk_df = filtered.groupby("forecast_risk", as_index=False).agg(Count=("Work Order", "count"))
        fig = px.pie(risk_df, names="forecast_risk", values="Count", hole=0.58, title="Forecast Delay Risk Profile")
        fig.update_traces(marker=dict(colors=[accent_red, accent_orange, accent_green, accent_blue]))
        st.plotly_chart(style_fig(fig, 380), use_container_width=True)

    st.markdown("#### Work Order Performance Detail")
    detail_cols = [c for c in ["Work Order", "Link Code", "City", "District", "Subclass", "Stage", "actual_progress_pct", "planned_progress_pct", "SPI", "CPI", "effective_target", "forecast_completion_date", "forecast_delay_days"] if c in filtered.columns]
    detail = filtered[detail_cols].copy().rename(columns={"actual_progress_pct":"Actual %", "planned_progress_pct":"Planned %", "effective_target":"Target Date", "forecast_completion_date":"Forecast Finish", "forecast_delay_days":"Forecast Delay Days"})
    st.dataframe(detail.sort_values("Forecast Delay Days", ascending=False), use_container_width=True, hide_index=True)

elif nav == "Risks & Penalties":
    p1, p2, p3, p4 = st.columns(4)
    p1.metric("Penalty Cases", f"{penalty_cases:,}")
    p2.metric("Deviation Qty", f"{penalty_qty:,.0f}")
    p3.metric("Penalty Amount", fmt_money(penalty_amount))
    p4.metric("Critical Lag >15pp", f"{critical_lag_cnt:,}")

    l1, l2 = st.columns([1.15, 1])
    with l1:
        if not pen_filtered.empty:
            city_pen = pen_filtered.groupby("City", as_index=False).agg(Deviations=("Number of Deviations", "sum"), Amount=("Penalties Amount", "sum"))
            fig = go.Figure()
            fig.add_trace(go.Bar(x=city_pen["City"], y=city_pen["Deviations"], name="Deviations", marker_color=accent_orange))
            fig.add_trace(go.Scatter(x=city_pen["City"], y=city_pen["Amount"], name="Amount", yaxis="y2", mode="lines+markers", line=dict(color=accent_blue, width=3)))
            fig.update_layout(title="Penalties by City", yaxis_title="Deviation Count", yaxis2=dict(title="Amount", overlaying="y", side="right"))
            st.plotly_chart(style_fig(fig, 380), use_container_width=True)
        else:
            st.info("No penalties match the selected filters.")
    with l2:
        if not pen_filtered.empty:
            top_dev = pen_filtered.groupby("Deviation name", as_index=False).agg(Count=("Number of Deviations", "sum")).sort_values("Count", ascending=False).head(10)
            fig = px.bar(top_dev, x="Count", y="Deviation name", orientation="h", title="Top Deviation Types")
            fig.update_traces(marker_color=accent_red)
            st.plotly_chart(style_fig(fig, 380), use_container_width=True)

    r1, r2 = st.columns([1.05, 1.15])
    with r1:
        risk_tbl = (
            filtered.groupby(["City", "District"], as_index=False)
            .agg(Overdue=("is_overdue", "sum"), CriticalLag=("critical_lag", "sum"), HighRisk=("forecast_risk", lambda s: (s == "High").sum()), PenaltyCases=("penalty_cases", "sum"))
            .sort_values(["HighRisk", "CriticalLag", "PenaltyCases"], ascending=False)
            .head(15)
        )
        st.markdown("#### Hotspot Table")
        st.dataframe(risk_tbl, use_container_width=True, hide_index=True)
    with r2:
        matrix = pd.DataFrame({
            "Category": ["Overdue", "Critical Lag", "High Forecast Risk", "Penalty Cases"],
            "Score": [overdue_cnt, critical_lag_cnt, high_risk, penalty_cases]
        })
        fig = px.treemap(matrix, path=["Category"], values="Score", color="Score", color_continuous_scale=[[0,"#28d7ac"],[0.5,"#ffb54a"],[1,"#ff6b6b"]], title="Risk Exposure Heatmap")
        st.plotly_chart(style_fig(fig, 380), use_container_width=True)

elif nav == "Data Quality":
    q1, q2, q3, q4 = st.columns(4)
    q1.metric("Missing City", dq["missing_city"])
    q2.metric("Missing District", dq["missing_district"])
    q3.metric("Missing Link Code", dq["missing_link"])
    q4.metric("Data Quality Score", f"{dq['score']}%")

    if dq["score"] >= 95:
        st.markdown("<div class='ok-box'>✅ Data quality is good. Mapping and required fields are largely complete.</div>", unsafe_allow_html=True)
    else:
        st.markdown("<div class='warn-box'>⚠️ Some core dimensions are missing. Review the rows below before presenting to management.</div>", unsafe_allow_html=True)

    issue_mask = (data["City"].isin(["Unknown", "Not Classified"])) | (data["District"].isin(["Unknown", "Not Classified"])) | (data["Link Code"].isna())
    issue_cols = [c for c in ["Link Code", "Work Order", "Region", "City", "District", "Project", "Subclass", "Stage", "Notes"] if c in data.columns]
    st.markdown("#### Rows needing mapping review")
    st.dataframe(data.loc[issue_mask, issue_cols].head(200), use_container_width=True, hide_index=True)

    st.markdown("#### Data Quality by City")
    city_dq = data.groupby("City", as_index=False).agg(Rows=("Link Code", "count"), MissingDistrict=("District", lambda s: s.isin(["Unknown", "Not Classified"]).sum()))
    city_dq["QualityScore"] = np.where(city_dq["Rows"] > 0, (1 - city_dq["MissingDistrict"] / city_dq["Rows"]) * 100, 100)
    fig = px.bar(city_dq.sort_values("QualityScore", ascending=False), x="City", y="QualityScore", text="QualityScore", title="Data Quality Score by City")
    fig.update_traces(marker_color=accent_blue, texttemplate="%{text:.1f}%")
    st.plotly_chart(style_fig(fig, 360), use_container_width=True)
