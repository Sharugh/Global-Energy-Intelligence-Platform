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

# ==========================================
# US STATES
# ==========================================
US_STATES = [
    "Texas",
    "Virginia",
    "California",
    "Arizona",
    "Ohio",
    "Georgia",
    "Illinois",
    "New York",
    "Nevada",
    "Oregon",
    "Washington",
    "North Carolina",
    "Utah"
]

# ==========================================
# COMPANY LIST
# ==========================================
COMPANIES = [
    "Amazon",
    "AWS",
    "Microsoft",
    "Google",
    "Meta",
    "Oracle",
    "Nvidia",
    "OpenAI",
    "Equinix",
    "Digital Realty",
    "QTS",
    "CyrusOne",
    "Aligned",
    "CoreSite",
    "Stack Infrastructure",
    "Vantage",
    "DataBank",
    "Compass Datacenters",
    "EdgeConneX",
    "Switch",
    "Iron Mountain",
    "NTT",
    "Yondr",
    "Apple"
]

# ==========================================
# STRICT US DATA CENTER FILTER
# ==========================================
def is_us_data_center_article(text):

    text = str(text).lower()

    # MUST HAVE DATA CENTER TERMS
    dc_keywords = [
        "data center",
        "data centre",
        "hyperscale",
        "colocation",
        "server farm",
        "data center campus"
    ]

    if not any(k in text for k in dc_keywords):
        return False

    # MUST HAVE US LOCATION
    us_keywords = [
        "usa",
        "united states",
        "texas",
        "virginia",
        "california",
        "ohio",
        "arizona",
        "georgia",
        "illinois",
        "new york",
        "nevada",
        "oregon"
    ]

    if not any(k in text for k in us_keywords):
        return False

    # REMOVE JUNK
    blacklist = [
        "movie",
        "sports",
        "celebrity",
        "football",
        "basketball",
        "music",
        "war",
        "election",
        "politics"
    ]

    if any(b in text for b in blacklist):
        return False

    return True

# ==========================================
# EXTRACT COMPANY
# ==========================================
def extract_company(text):

    text = str(text).lower()

    found = []

    for c in COMPANIES:
        if c.lower() in text:
            found.append(c)

    if found:
        return ", ".join(found)

    return "Unknown"

# ==========================================
# EXTRACT STATE
# ==========================================
def extract_state(text):

    text = str(text).lower()

    for s in US_STATES:
        if s.lower() in text:
            return s

    return "USA"

# ==========================================
# EXTRACT LOCATION
# ==========================================
def extract_location(text):

    patterns = [
        r"([A-Z][a-z]+,\s?[A-Z]{2})",
        r"([A-Z][a-z]+\s[A-Z][a-z]+,\s?[A-Z]{2})"
    ]

    for p in patterns:
        match = re.search(p, str(text))
        if match:
            return match.group(0)

    return "N/A"

# ==========================================
# SCRAPER - DATACENTERDYNAMICS
# ==========================================
def scrape_dcd():

    url = "https://www.datacenterdynamics.com/en/news/"

    try:

        headers = {
            "User-Agent": "Mozilla/5.0"
        }

        res = requests.get(
            url,
            headers=headers,
            timeout=15
        )

        soup = BeautifulSoup(res.text, "html.parser")

        articles = []

        links = soup.find_all("a")

        for link in links:

            title = link.get_text(strip=True)

            href = link.get("href")

            if not href:
                continue

            if len(title) < 25:
                continue

            # STRICT FILTER
            if not is_us_data_center_article(title):
                continue

            if href.startswith("/"):
                href = "https://www.datacenterdynamics.com" + href

            articles.append({
                "Title": title,
                "URL": href,
                "Published Date": datetime.now().strftime("%Y-%m-%d"),
                "Description": "US Data Center article from DataCenterDynamics",
                "Source": "DataCenterDynamics",
                "Company": extract_company(title),
                "State": extract_state(title),
                "Location": extract_location(title)
            })

        return pd.DataFrame(articles)

    except Exception as e:

        st.warning(f"DCD scraping failed: {e}")

        return pd.DataFrame()

# ==========================================
# NEWS API FETCH
# ==========================================
def fetch_news(start_date, end_date):

    query = (
        '"data center" OR "data centre" OR '
        'hyperscale OR colocation'
    )

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
            return pd.DataFrame()

        data = res.json()

        articles = data.get("articles", [])

        rows = []

        for a in articles:

            title = a.get("title", "")
            desc = a.get("description", "")

            text = f"{title} {desc}"

            # STRICT FILTER
            if not is_us_data_center_article(text):
                continue

            date_value = a.get("publishedAt")

            try:
                date_value = pd.to_datetime(
                    date_value
                ).strftime("%Y-%m-%d")
            except:
                pass

            rows.append({
                "Title": title,
                "URL": a.get("url"),
                "Published Date": date_value,
                "Description": desc,
                "Source": a.get("source", {}).get("name"),
                "Company": extract_company(text),
                "State": extract_state(text),
                "Location": extract_location(text)
            })

        return pd.DataFrame(rows)

    except Exception as e:

        st.warning(f"NewsAPI failed: {e}")

        return pd.DataFrame()

# ==========================================
# EXPORT
# ==========================================
def to_excel(df):

    output = BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        df.to_excel(
            writer,
            index=False
        )

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
        """
        Track:
        - US hyperscale developments
        - AI infrastructure expansion
        - Colocation investments
        - Data center campus developments
        """
    )

    # ======================================
    # SIDEBAR
    # ======================================
    st.sidebar.header("Filters")

    time_filter = st.sidebar.radio(
        "Select Time Range",
        ["Latest", "Past 10 Days", "Past 30 Days"]
    )

    # ======================================
    # DATE LOGIC
    # ======================================
    end_date = datetime.today()

    if time_filter == "Latest":
        start_date = end_date - timedelta(days=1)

    elif time_filter == "Past 10 Days":
        start_date = end_date - timedelta(days=10)

    else:
        start_date = end_date - timedelta(days=30)

    # ======================================
    # FETCH BUTTON
    # ======================================
    if st.button("🚀 Fetch US Data Center News"):

        with st.spinner("Collecting US data center intelligence..."):

            # NEWS API
            df_api = fetch_news(
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

            # REMOVE DUPLICATES
            df.drop_duplicates(
                subset="Title",
                inplace=True
            )

            # SORT
            if not df.empty:
                df = df.sort_values(
                    by="Published Date",
                    ascending=False
                )

            # ==================================
            # OUTPUT
            # ==================================
            if not df.empty:

                st.success(
                    f"{len(df)} US Data Center Articles Found"
                )

                st.dataframe(
                    df,
                    use_container_width=True
                )

                excel = to_excel(df)

                st.download_button(
                    "📥 Download Excel",
                    excel,
                    "us_data_center_intelligence.xlsx"
                )

            else:

                st.warning(
                    "No US data center news found"
                )

# ==========================================
# MAIN
# ==========================================
if __name__ == "__main__":
    main()
