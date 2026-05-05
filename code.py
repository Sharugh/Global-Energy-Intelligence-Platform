import streamlit as st
import requests
import pandas as pd
import re
from datetime import datetime
import matplotlib.pyplot as plt

# =========================
# CONFIG
# =========================
GDELT_API = "https://api.gdeltproject.org/api/v2/doc/doc"

COMPANIES = [
    "Amazon", "AWS", "Microsoft", "Google", "Meta",
    "Equinix", "Digital Realty", "QTS", "CyrusOne"
]

KEYWORDS_BASE = [
    "data center", "hyperscale", "AI infrastructure",
    "cloud", "server farm"
]

# =========================
# AUTO KEYWORD GENERATOR
# =========================
def generate_query():
    return (
        "(" + " OR ".join(KEYWORDS_BASE) + ") AND "
        "(investment OR expansion OR construction OR acquisition OR power)"
        " AND (USA OR 'United States' OR Texas OR Virginia OR California)"
    )

# =========================
# FETCH FROM GDELT
# =========================
def fetch_gdelt(query):
    params = {
        "query": query,
        "format": "json",
        "maxrecords": 100,
        "sort": "DateDesc"
    }

    res = requests.get(GDELT_API, params=params)

    if res.status_code != 200:
        st.error("GDELT fetch failed")
        return None

    data = res.json().get("articles", [])

    if not data:
        return None

    df = pd.DataFrame(data)[["title", "sourceCommonName", "seendate", "url"]]
    df.columns = ["Title", "Source", "Date", "URL"]

    return df

# =========================
# TAG COMPANIES
# =========================
def tag_companies(text):
    found = []
    for c in COMPANIES:
        if c.lower() in str(text).lower():
            found.append(c)
    return ", ".join(found) if found else "Other"

# =========================
# CLASSIFY TYPE
# =========================
def classify_article(title):
    title = title.lower()

    if "investment" in title or "funding" in title:
        return "Investment"
    elif "construction" in title or "build" in title:
        return "Construction"
    elif "power" in title or "energy" in title:
        return "Energy"
    elif "ai" in title:
        return "AI Infra"
    else:
        return "General"

# =========================
# EXTRACT CAPEX (SIMPLE)
# =========================
def extract_capex(text):
    match = re.search(r"\$[0-9]+(\.[0-9]+)?\s?(billion|million)", str(text).lower())
    return match.group(0) if match else "N/A"

# =========================
# MAIN APP
# =========================
def main():
    st.set_page_config(layout="wide")

    st.title("🏢 US Data Center Investment Tracker")

    st.sidebar.header("Controls")

    if st.sidebar.button("🔄 Auto Generate Query"):
        query = generate_query()
    else:
        query = st.sidebar.text_area(
            "Edit Query",
            value=generate_query()
        )

    if st.sidebar.button("🚀 Fetch Data"):
        with st.spinner("Fetching from GDELT..."):
            df = fetch_gdelt(query)

            if df is not None:

                # Tagging
                df["Company"] = df["Title"].apply(tag_companies)
                df["Category"] = df["Title"].apply(classify_article)
                df["Capex"] = df["Title"].apply(extract_capex)

                st.success(f"{len(df)} Articles Loaded")

                # =====================
                # METRICS
                # =====================
                col1, col2, col3 = st.columns(3)
                col1.metric("Articles", len(df))
                col2.metric("Companies", df["Company"].nunique())
                col3.metric("Categories", df["Category"].nunique())

                # =====================
                # CHARTS (Power BI Style)
                # =====================
                st.subheader("📊 Trends")

                # Category Chart
                cat_counts = df["Category"].value_counts()

                fig1, ax1 = plt.subplots()
                cat_counts.plot(kind="bar", ax=ax1)
                st.pyplot(fig1)

                # Company Chart
                comp_counts = df["Company"].value_counts().head(10)

                fig2, ax2 = plt.subplots()
                comp_counts.plot(kind="bar", ax=ax2)
                st.pyplot(fig2)

                # =====================
                # TABLE
                # =====================
                st.subheader("📄 Articles")

                st.dataframe(df, use_container_width=True)

                # =====================
                # DOWNLOAD
                # =====================
                csv = df.to_csv(index=False)

                st.download_button(
                    "📥 Download CSV",
                    csv,
                    "data_center_tracker.csv"
                )

            else:
                st.warning("No data found")


if __name__ == "__main__":
    main()
