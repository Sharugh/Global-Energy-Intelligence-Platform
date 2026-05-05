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

BASE_QUERY = '''
("data center" OR hyperscale OR "AI infrastructure")
AND (investment OR expansion OR construction OR acquisition)
'''

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
    url = "https://gnews.io/"

    params = {
        "apikey": API_KEY,
        "q": query,
        "language": "en",
        "country": "us",
        "from_date": start_date,
        "size": 50
    }

    res = requests.get(url, params=params)

    if res.status_code != 200:
        st.error("API Error")
        return None

    articles = res.json().get("results", [])

    if not articles:
        return None

    data = []
    for a in articles:
        text = f"{a.get('title','')} {a.get('description','')}"

        data.append({
            "Title": a.get("title"),
            "URL": a.get("link"),
            "Date": a.get("pubDate"),
            "Description": a.get("description"),
            "Company": extract_company(text),
            "Location": extract_location(text),
            "MW": extract_mw(text),
            "Acre": extract_acre(text),
            "Cost": extract_cost(text),
            "Project Type": classify_project(text),
            "Signal": extract_signal(text)
        })

    df = pd.DataFrame(data)
    df.drop_duplicates(subset="Title", inplace=True)

    return df

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

    st.sidebar.header("Filters")

    state = st.sidebar.selectbox("Select State", US_STATES)

    # Time selection
    time_option = st.sidebar.radio(
        "Select Time Range",
        ["Latest", "Past 10 Days", "Past 30 Days"]
    )

    if time_option == "Latest":
        start_date = datetime.today() - timedelta(days=1)
    elif time_option == "Past 10 Days":
        start_date = datetime.today() - timedelta(days=10)
    else:
        start_date = datetime.today() - timedelta(days=30)

    # Build query
    if state != "All USA":
        query = BASE_QUERY + f" AND {state}"
    else:
        query = BASE_QUERY + " AND USA"

    # Tabs
    tab1, tab2, tab3 = st.tabs(["Latest", "Past 10 Days", "Past 30 Days"])

    if st.sidebar.button("🚀 Fetch Data"):
        with st.spinner("Fetching Data..."):
            df = fetch_news(query, start_date.strftime("%Y-%m-%d"))

            if df is not None:

                st.success(f"{len(df)} Articles Found")

                st.dataframe(df, use_container_width=True)

                excel = to_excel(df)

                st.download_button(
                    "📥 Download Excel",
                    excel,
                    "data_center_tracker.xlsx"
                )
            else:
                st.warning("No data found")

if __name__ == "__main__":
    main()
