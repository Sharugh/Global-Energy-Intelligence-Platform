import streamlit as st
import requests
import pandas as pd
import re
from datetime import datetime, timedelta
from io import BytesIO

# =========================
# CONFIG
# =========================
API_KEY = "c3fd1ae889b305ab837a3a344d890e3f"

BASE_QUERY = "data center OR hyperscale OR AI infrastructure USA investment expansion"

US_STATES = [
    "All USA", "Texas", "Virginia", "California", "Arizona",
    "Ohio", "Georgia", "Illinois", "New York"
]

COMPANIES = [
    "Amazon", "AWS", "Microsoft", "Google", "Meta",
    "Equinix", "Digital Realty", "QTS", "CyrusOne"
]

# =========================
# FETCH FUNCTION
# =========================
def fetch_news(query, start_date):
    url = "https://gnews.io/api/v4/search"

    params = {
        "q": query,
        "lang": "en",
        "country": "us",
        "from": start_date,
        "max": 50,
        "apikey": API_KEY
    }

    try:
        res = requests.get(url, params=params, timeout=10)

        if res.status_code != 200:
            st.error(f"API Error {res.status_code}")
            st.text(res.text[:300])
            return None

        try:
            json_data = res.json()
        except:
            st.error("Invalid JSON response")
            return None

        articles = json_data.get("articles", [])

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
                "Company": extract_company(text),
                "State": extract_location(text),
                "MW": extract_mw(text),
                "Acre": extract_acre(text),
                "Cost": extract_cost(text),
                "Project Type": classify_project(text),
                "Signal": extract_signal(text)
            })

        df = pd.DataFrame(data)
        df.drop_duplicates(subset="Title", inplace=True)

        df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
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

    def run_tab(days):
        start_date = (datetime.today() - timedelta(days=days)).strftime("%Y-%m-%d")

        if state != "All USA":
            query = f"{BASE_QUERY} {state}"
        else:
            query = f"{BASE_QUERY} USA"

        if st.button(f"Fetch {days}-day data"):
            with st.spinner("Fetching news..."):
                df = fetch_news(query, start_date)

                if df is not None:
                    st.success(f"{len(df)} Articles Found")

                    st.dataframe(df, use_container_width=True)

                    excel = to_excel(df)

                    st.download_button(
                        "📥 Download Excel",
                        excel,
                        f"data_center_{days}days.xlsx"
                    )
                else:
                    st.warning("No data found")

    with tab1:
        run_tab(1)

    with tab2:
        run_tab(10)

    with tab3:
        run_tab(30)

if __name__ == "__main__":
    main()
