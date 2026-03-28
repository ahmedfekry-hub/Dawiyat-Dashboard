import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px

st.set_page_config(layout="wide")

# ==============================
# LOAD DATA (SAFE ENGINE)
# ==============================

@st.cache_data
def load_data(file):
    df = pd.read_excel(file, sheet_name=0)

    # Try penalties sheet
    try:
        penalties = pd.read_excel(file, sheet_name="Penalties")
    except:
        penalties = pd.DataFrame()

    return df, penalties


# ==============================
# SAFE COLUMN GETTER
# ==============================

def safe_col(df, col):
    return df[col] if col in df.columns else pd.Series([None]*len(df))


# ==============================
# DATA QUALITY CHECK
# ==============================

def data_quality(df):
    total = len(df)

    missing_city = safe_col(df, "City").isna().sum()
    missing_district = safe_col(df, "District").isna().sum()
    missing_link = safe_col(df, "Link Code").isna().sum()

    completeness = 100 - ((missing_city + missing_district + missing_link) / (3 * total) * 100)

    return {
        "missing_city": missing_city,
        "missing_district": missing_district,
        "missing_link": missing_link,
        "score": round(completeness, 1)
    }


# ==============================
# FILE UPLOAD
# ==============================

file = st.sidebar.file_uploader("Upload Excel", type=["xlsx"])

if file:
    df, penalties = load_data(file)
else:
    df, penalties = load_data("Dawiyat Master Sheet.xlsx")

# ==============================
# CLEAN DATA
# ==============================

df["City"] = safe_col(df, "City").fillna("Unknown")
df["District"] = safe_col(df, "District").fillna("Unknown")
df["Region"] = safe_col(df, "Region").fillna("Unknown")
df["Link Code"] = safe_col(df, "Link Code").fillna("Unknown")

# ==============================
# CASCADING FILTERS
# ==============================

region = st.sidebar.selectbox("Region", ["All"] + sorted(df["Region"].unique()))

filtered = df.copy()

if region != "All":
    filtered = filtered[filtered["Region"] == region]

city = st.sidebar.selectbox("City", ["All"] + sorted(filtered["City"].unique()))

if city != "All":
    filtered = filtered[filtered["City"] == city]

district = st.sidebar.selectbox("District", ["All"] + sorted(filtered["District"].unique()))

if district != "All":
    filtered = filtered[filtered["District"] == district]

link = st.sidebar.selectbox("Link Code", ["All"] + sorted(filtered["Link Code"].unique()))

if link != "All":
    filtered = filtered[filtered["Link Code"] == link]

# ==============================
# KPI CALCULATIONS
# ==============================

total_links = filtered["Link Code"].nunique()

progress = safe_col(filtered, "Percentage of Completion").mean()

civil = safe_col(filtered, "Trench Progress").mean()
fiber = safe_col(filtered, "Fiber Progress").mean()

# ==============================
# DATA QUALITY
# ==============================

dq = data_quality(filtered)

# ==============================
# SIDEBAR MENU
# ==============================

menu = st.sidebar.radio(
    "Navigation",
    ["Overview", "Performance", "Risks & Penalties", "Data Quality"]
)

# ==============================
# OVERVIEW
# ==============================

if menu == "Overview":

    st.title("Executive Project Dashboard")

    col1, col2, col3, col4 = st.columns(4)

    col1.metric("Total Link Codes", total_links)
    col2.metric("Avg Progress %", f"{progress:.1f}%")
    col3.metric("Civil Completion", f"{civil:.1f}%")
    col4.metric("Fiber Completion", f"{fiber:.1f}%")

    fig = px.bar(filtered, x="City", y="Percentage of Completion", color="City")
    st.plotly_chart(fig, use_container_width=True)

# ==============================
# PERFORMANCE
# ==============================

elif menu == "Performance":

    st.title("Performance Analysis")

    fig1 = px.line(filtered, x="Targeted Completion", y="Percentage of Completion")
    st.plotly_chart(fig1, use_container_width=True)

    fig2 = px.bar(filtered, x="District", y="Percentage of Completion", color="District")
    st.plotly_chart(fig2, use_container_width=True)

# ==============================
# RISKS & PENALTIES
# ==============================

elif menu == "Risks & Penalties":

    st.title("Penalties & Risk Overview")

    if not penalties.empty:
        penalties["City"] = safe_col(penalties, "City").fillna("Unknown")

        fig = px.bar(penalties, x="City", title="Penalties by City")
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.warning("No penalties data available")

# ==============================
# DATA QUALITY
# ==============================

elif menu == "Data Quality":

    st.title("Data Quality Dashboard")

    col1, col2, col3, col4 = st.columns(4)

    col1.metric("Missing City", dq["missing_city"])
    col2.metric("Missing District", dq["missing_district"])
    col3.metric("Missing Link Code", dq["missing_link"])
    col4.metric("Data Quality Score", f"{dq['score']}%")

    # Warning panel
    if dq["score"] < 90:
        st.error("⚠️ Data Quality Issue Detected")
    else:
        st.success("✅ Data Quality is Good")

    st.dataframe(filtered.head(50))
