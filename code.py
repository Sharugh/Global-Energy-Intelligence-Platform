import streamlit as st
import requests
import pandas as pd
import re
from datetime import datetime, timedelta
from io import BytesIO
from bs4 import BeautifulSoup

# =========================
# CONFIG
# =========================
NEWSAPI_KEY = "3087034a13564f75bfc769c0046e729c"
NEWSAPI_URL = "https://newsapi.org/v2/everything"

BASE_QUERY = '"data center" OR hyperscale OR colocation OR campus'

US_STATES = [
    "All USA", "Texas", "Virginia", "California", "Arizona",
    "Ohio", "Georgia", "Illinois", "New York"
]

COMPANIES = [
    "Amazon", "AWS", "Microsoft", "Google", "Meta",
    "Equinix", "Digital Realty", "QTS", "CyrusOne"
]

# =========================
# SCRAPER: DataCenterDynamics
# =========================
def scrape_dcd():
    url = "https://www.datacenterdynamics.com/en/news/"
    
    try:
        res = requests.get(url, timeout=10)
        soup = BeautifulSoup(res.text, "html.parser")

        articles = []
        cards = soup.find_all("a", class_="card__link")

        for c in cards[:20]:
            title = c.get_text(strip=True)
            link = "https://www.datacenterdynamics.com" + c.get("href")

            articles.append({
                "Title": title,
                "URL": link,
                "Date": datetime.now(),
                "Description": "From DataCenterDynamics",
                "Source": "DataCenterDynamics"
            })

        return pd.DataFrame(articles)

    except Exception as e:
        st.warning("DCD scraping failed")
        return pd.DataFrame()

# =========================
# NEWS API FETCH
# =========================
def fetch_news(query, start_date, end_date):
    params = {
        "q": query,
        "from": start_date,
        "to": end_date,
        "language": "en",
        "sortBy": "publishedAt",
        "pageSize": 100,
        "apiKey": NEWSAPI_KEY
    }

    try:
        res = requests.get(NEWSAPI_URL, params=params, timeout=10)
        articles = res.json().get("articles", [])

        data = []

        for a in articles:
            text = f"{a.get('title','')} {a.get('description','')}"

            data.append({
                "Title": a.get("title"),
                "URL": a.get("url"),
                "Date": a.get("publishedAt"),
                "Description": a.get("description"),
                "Source": a.get("source", {}).get("name")
            })

        df = pd.DataFrame(data)

        if not df.empty:
            df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
            df["Date"] = df["Date"].dt.tz_localize(None)

        return df

    except:
        return pd.DataFrame()

# =========================
# EXTRACTION
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

def extract_cost(text):
    match = re.search(r"\$[\d\.]+\s?(billion|million)", text, re.IGNORECASE)
    return match.group(0) if match else "N/A"

# =========================
# PROCESS DATA
# =========================
def process_df(df):
    if df.empty:
        return df

    df["Company"] = df["Title"].apply(extract_company)
    df["State"] = df["Title"].apply(extract_location)
    df["MW"] = df["Title"].apply(extract_mw)
    df["Cost"] = df["Title"].apply(extract_cost)

    df.drop_duplicates(subset="Title", inplace=True)
    return df

# =========================
# EXPORT
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
    st.title("🏢 US Data Center Intelligence Tracker")

    state = st.sidebar.selectbox("Select State", US_STATES)

    tab1, tab2, tab3 = st.tabs(["Latest", "Past 10 Days", "Past 30 Days"])

    def run_tab(days, label):
        end_date = datetime.today()
        start_date = end_date - timedelta(days=days)

        query = BASE_QUERY if state == "All USA" else f"{BASE_QUERY} {state}"

        if st.button(f"Fetch {label}"):
            with st.spinner("Fetching data..."):

                # API Data
                df_api = fetch_news(
                    query,
                    start_date.strftime("%Y-%m-%d"),
                    end_date.strftime("%Y-%m-%d")
                )

                # Scraped Data
                df_scrape = scrape_dcd()

                # Combine
                df = pd.concat([df_api, df_scrape], ignore_index=True)

                df = process_df(df)

                if not df.empty:
                    st.success(f"{len(df)} Articles Found")
                    st.dataframe(df, use_container_width=True)

                    excel = to_excel(df)

                    st.download_button(
                        "📥 Download Excel",
                        excel,
                        f"data_center_{label}.xlsx"
                    )
                else:
                    st.warning("No data found")

    with tab1:
        run_tab(1, "latest")

    with tab2:
        run_tab(10, "10_days")

    with tab3:
        run_tab(30, "30_days")

if __name__ == "__main__":
    main()
