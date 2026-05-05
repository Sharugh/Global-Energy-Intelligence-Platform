import streamlit as st
import requests
import pandas as pd
import re
from datetime import datetime, timedelta
from io import BytesIO

# =========================
# CONFIG
# =========================
NEWSAPI_KEY = "3087034a13564f75bfc769c0046e729c"
NEWSAPI_URL = "https://newsapi.org/v2/everything"

# 🔥 HIGH-PRECISION QUERY
BASE_QUERY = '"data center" OR "data centre" OR hyperscale OR colocation'

# 🔥 TRUSTED SOURCES
TRUSTED_SOURCES = [
    "datacenterdynamics.com",
    "datacenterknowledge.com",
    "reuters.com",
    "bloomberg.com"
]

US_STATES = [
    "All USA", "Texas", "Virginia", "California", "Arizona",
    "Ohio", "Georgia", "Illinois", "New York"
]

COMPANIES = [
    "Amazon", "AWS", "Microsoft", "Google", "Meta",
    "Equinix", "Digital Realty", "QTS", "CyrusOne"
]

# =========================
# RELEVANCE FILTER
# =========================
def is_relevant(article):
    text = f"{article.get('title','')} {article.get('description','')}".lower()
    url = article.get("url", "")

    keywords = ["data center", "data centre", "hyperscale", "colocation"]

    if not any(k in text for k in keywords):
        return False

    if not any(domain in url for domain in TRUSTED_SOURCES):
        return False

    return True

def relevance_score(text):
    text = text.lower()
    keywords = [
        "data center", "hyperscale", "colocation",
        "mw", "campus", "facility", "server"
    ]
    return sum(1 for k in keywords if k in text)

# =========================
# FETCH FUNCTION
# =========================
def fetch_news(query, start_date, end_date):
    params = {
        "q": query,
        "from": start_date,
        "to": end_date,
        "language": "en",
        "sortBy": "publishedAt",
        "pageSize": 100,
        "domains": ",".join(TRUSTED_SOURCES),
        "apiKey": NEWSAPI_KEY
    }

    try:
        res = requests.get(NEWSAPI_URL, params=params, timeout=10)

        if res.status_code != 200:
            st.error(f"API Error {res.status_code}")
            st.text(res.text[:300])
            return None

        data_json = res.json()
        articles = data_json.get("articles", [])

        # 🔥 STRICT FILTER
        articles = [a for a in articles if is_relevant(a)]

        if not articles:
            return None

        data = []

        for a in articles:
            text = f"{a.get('title','')} {a.get('description','')}"

            data.append({
                "Title": a.get("title"),
                "URL": a.get("url"),
                "Date": a.get("publishedAt"),
                "Description": a.get("description"),
                "Source": a.get("source", {}).get("name"),
                "Company": extract_company(text),
                "State": extract_location(text),
                "MW": extract_mw(text),
                "Acre": extract_acre(text),
                "Cost": extract_cost(text),
                "Project Type": classify_project(text),
                "Signal": extract_signal(text),
                "Score": relevance_score(text)
            })

        df = pd.DataFrame(data)

        # 🔥 Keep only strong matches
        df = df[df["Score"] >= 2]

        # Fix datetime
        df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
        df["Date"] = df["Date"].dt.tz_localize(None)

        df.drop_duplicates(subset="Title", inplace=True)
        df = df.sort_values(by="Date", ascending=False)

        return df

    except requests.exceptions.RequestException as e:
        st.error(f"Request failed: {e}")
        return None

# =========================
# EXTRACTION FUNCTIONS
# =========================
def extract_company(text):
    for c in COMPANIES:
        if c.lower() in text.lower():
            return c
    return "Unknown"

def extract_location(text):
    for s in US_STATES:
        if s.lower() in text.lower():
            return s
    return "USA"

def extract_mw(text):
    match = re.search(r"\d+\s?MW", text, re.IGNORECASE)
    return match.group(0) if match else "N/A"

def extract_acre(text):
    match = re.search(r"\d+\s?(acre|acres)", text, re.IGNORECASE)
    return match.group(0) if match else "N/A"

def extract_cost(text):
    match = re.search(r"\$[\d\.]+\s?(billion|million)", text, re.IGNORECASE)
    return match.group(0) if match else "N/A"

def classify_project(text):
    text = text.lower()
    if "investment" in text:
        return "Investment"
    elif "construction" in text or "build" in text:
        return "Construction"
    elif "expansion" in text:
        return "Expansion"
    else:
        return "General"

def extract_signal(text):
    keywords = ["approval", "permit", "announced", "planned"]
    for k in keywords:
        if k in text.lower():
            return k
    return "N/A"

# =========================
# EXCEL EXPORT
# =========================
def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

# =========================
# UI
# =========================
def main():
    st.set_page_config(layout="wide")
    st.title("🏢 US Data Center Investment Tracker")

    st.sidebar.header("🔍 Filters")
    state = st.sidebar.selectbox("Select State", US_STATES)

    tab1, tab2, tab3 = st.tabs(["Latest", "Past 10 Days", "Past 30 Days"])

    def run_tab(days, label):
        end_date = datetime.today()
        start_date = end_date - timedelta(days=days)

        if state != "All USA":
            query = f"{BASE_QUERY} {state}"
        else:
            query = f"{BASE_QUERY} USA"

        if st.button(f"Fetch {label}"):
            with st.spinner("Fetching high-quality data center news..."):
                df = fetch_news(
                    query,
                    start_date.strftime("%Y-%m-%d"),
                    end_date.strftime("%Y-%m-%d")
                )

                if df is not None and not df.empty:
                    st.success(f"{len(df)} High-Quality Articles Found")
                    st.dataframe(df, use_container_width=True)

                    excel = to_excel(df)

                    st.download_button(
                        "📥 Download Excel",
                        excel,
                        f"data_center_{label}.xlsx"
                    )
                else:
                    st.warning("No high-quality data center articles found")

    with tab1:
        run_tab(1, "latest")

    with tab2:
        run_tab(10, "10_days")

    with tab3:
        run_tab(30, "30_days")

if __name__ == "__main__":
    main()
