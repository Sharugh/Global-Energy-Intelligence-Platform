import streamlit as st
import requests
import pandas as pd
import re
from datetime import datetime, timedelta
from io import BytesIO
from bs4 import BeautifulSoup

# ==========================================
# CONFIG
# ==========================================
NEWSAPI_KEY = "3087034a13564f75bfc769c0046e729c"
NEWSAPI_URL = "https://newsapi.org/v2/everything"

# STRICT DATA CENTER QUERY
BASE_QUERY = (
    '"data center" OR "data centre" OR '
    'hyperscale OR colocation OR "data center campus"'
)

US_STATES = [
    "All USA",
    "Texas",
    "Virginia",
    "California",
    "Arizona",
    "Ohio",
    "Georgia",
    "Illinois",
    "New York"
]

COMPANIES = [
    "Amazon",
    "AWS",
    "Microsoft",
    "Google",
    "Meta",
    "Equinix",
    "Digital Realty",
    "QTS",
    "CyrusOne",
    "Oracle",
    "Nvidia"
]

# ==========================================
# STRICT DATA CENTER FILTER
# ==========================================
def is_data_center_article(text):
    text = str(text).lower()

    # MUST HAVE
    strong_keywords = [
        "data center",
        "data centre",
        "hyperscale",
        "colocation"
    ]

    if not any(k in text for k in strong_keywords):
        return False

    # BOOST KEYWORDS
    boost_keywords = [
        "mw",
        "campus",
        "facility",
        "server",
        "expansion",
        "construction",
        "cloud",
        "ai infrastructure",
        "power",
        "substation"
    ]

    score = sum(1 for k in boost_keywords if k in text)

    # BLACKLIST
    blacklist = [
        "movie",
        "sports",
        "celebrity",
        "election",
        "war",
        "football",
        "basketball",
        "hollywood",
        "music",
        "netflix show"
    ]

    if any(b in text for b in blacklist):
        return False

    # REQUIRE AT LEAST ONE BOOST SIGNAL
    if score < 1:
        return False

    return True

# ==========================================
# SCRAPER: DATACENTERDYNAMICS
# ==========================================
def scrape_dcd():
    url = "https://www.datacenterdynamics.com/en/news/"

    try:
        headers = {
            "User-Agent": "Mozilla/5.0"
        }

        res = requests.get(url, headers=headers, timeout=15)

        soup = BeautifulSoup(res.text, "html.parser")

        articles = []

        cards = soup.find_all("a")

        for c in cards:

            title = c.get_text(strip=True)

            href = c.get("href")

            if not href:
                continue

            if len(title) < 20:
                continue

            full_text = title.lower()

            # STRICT FILTER
            if not is_data_center_article(full_text):
                continue

            # BUILD URL
            if href.startswith("/"):
                href = "https://www.datacenterdynamics.com" + href

            articles.append({
                "Title": title,
                "URL": href,
                "Date": datetime.now(),
                "Description": "Scraped from DataCenterDynamics",
                "Source": "DataCenterDynamics"
            })

        return pd.DataFrame(articles)

    except Exception as e:
        st.warning(f"DCD Scraper Failed: {e}")
        return pd.DataFrame()

# ==========================================
# NEWS API FETCH
# ==========================================
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
        res = requests.get(
            NEWSAPI_URL,
            params=params,
            timeout=15
        )

        if res.status_code != 200:
            st.error(f"API Error: {res.status_code}")
            return pd.DataFrame()

        json_data = res.json()

        articles = json_data.get("articles", [])

        data = []

        for a in articles:

            title = a.get("title", "")
            desc = a.get("description", "")

            text = f"{title} {desc}"

            # STRICT FILTER
            if not is_data_center_article(text):
                continue

            data.append({
                "Title": title,
                "URL": a.get("url"),
                "Date": a.get("publishedAt"),
                "Description": desc,
                "Source": a.get("source", {}).get("name")
            })

        df = pd.DataFrame(data)

        if not df.empty:
            df["Date"] = pd.to_datetime(
                df["Date"],
                errors="coerce"
            )

            # REMOVE TIMEZONE FOR EXCEL
            df["Date"] = df["Date"].dt.tz_localize(None)

        return df

    except Exception as e:
        st.error(f"Fetch Failed: {e}")
        return pd.DataFrame()

# ==========================================
# EXTRACTION FUNCTIONS
# ==========================================
def extract_company(text):
    text = str(text).lower()

    for c in COMPANIES:
        if c.lower() in text:
            return c

    return "Unknown"

def extract_location(text):
    text = str(text).lower()

    for s in US_STATES:
        if s.lower() in text:
            return s

    return "USA"

def extract_mw(text):
    match = re.search(r"\d+\s?MW", str(text), re.IGNORECASE)
    return match.group(0) if match else "N/A"

def extract_acre(text):
    match = re.search(
        r"\d+\s?(acre|acres)",
        str(text),
        re.IGNORECASE
    )
    return match.group(0) if match else "N/A"

def extract_cost(text):
    match = re.search(
        r"\$[\d\.]+\s?(billion|million)",
        str(text),
        re.IGNORECASE
    )
    return match.group(0) if match else "N/A"

def classify_project(text):

    text = str(text).lower()

    if "investment" in text:
        return "Investment"

    elif "construction" in text or "build" in text:
        return "Construction"

    elif "expansion" in text:
        return "Expansion"

    return "General"

def extract_signal(text):

    keywords = [
        "approval",
        "permit",
        "announced",
        "planned",
        "commissioned"
    ]

    text = str(text).lower()

    for k in keywords:
        if k in text:
            return k

    return "N/A"

# ==========================================
# PROCESS DATAFRAME
# ==========================================
def process_df(df):

    if df.empty:
        return df

    combined_text = (
        df["Title"].fillna("") + " " +
        df["Description"].fillna("")
    )

    df["Company"] = combined_text.apply(extract_company)
    df["State"] = combined_text.apply(extract_location)
    df["MW"] = combined_text.apply(extract_mw)
    df["Acre"] = combined_text.apply(extract_acre)
    df["Cost"] = combined_text.apply(extract_cost)
    df["Project Type"] = combined_text.apply(classify_project)
    df["Signal"] = combined_text.apply(extract_signal)

    df.drop_duplicates(subset="Title", inplace=True)

    df = df.sort_values(
        by="Date",
        ascending=False
    )

    return df

# ==========================================
# EXCEL EXPORT
# ==========================================
def to_excel(df):

    output = BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        df.to_excel(writer, index=False)

    return output.getvalue()

# ==========================================
# UI
# ==========================================
def main():

    st.set_page_config(
        page_title="US Data Center Intelligence",
        layout="wide"
    )

    st.title("🏢 US Data Center Intelligence Tracker")

    st.markdown(
        "Track US hyperscale, colocation, and AI infrastructure developments."
    )

    st.sidebar.header("🔍 Filters")

    state = st.sidebar.selectbox(
        "Select State",
        US_STATES
    )

    tab1, tab2, tab3 = st.tabs([
        "Latest",
        "Past 10 Days",
        "Past 30 Days"
    ])

    def run_tab(days, label):

        end_date = datetime.today()
        start_date = end_date - timedelta(days=days)

        if state != "All USA":
            query = f"{BASE_QUERY} {state}"
        else:
            query = BASE_QUERY

        if st.button(f"Fetch {label}"):

            with st.spinner("Fetching Data Center Intelligence..."):

                # NEWS API
                df_api = fetch_news(
                    query,
                    start_date.strftime("%Y-%m-%d"),
                    end_date.strftime("%Y-%m-%d")
                )

                # DCD SCRAPER
                df_dcd = scrape_dcd()

                # COMBINE
                df = pd.concat(
                    [df_api, df_dcd],
                    ignore_index=True
                )

                # FINAL PROCESSING
                df = process_df(df)

                if not df.empty:

                    st.success(
                        f"{len(df)} Data Center Articles Found"
                    )

                    st.dataframe(
                        df,
                        use_container_width=True
                    )

                    excel = to_excel(df)

                    st.download_button(
                        "📥 Download Excel",
                        excel,
                        f"data_center_{label}.xlsx"
                    )

                else:
                    st.warning(
                        "No relevant data center news found"
                    )

    with tab1:
        run_tab(1, "latest")

    with tab2:
        run_tab(10, "10_days")

    with tab3:
        run_tab(30, "30_days")

if __name__ == "__main__":
    main()
