import streamlit as st
import requests
import pandas as pd
import re
from datetime import datetime, timedelta
from io import BytesIO
from bs4 import BeautifulSoup

# ==========================================================
# CONFIG
# ==========================================================
NEWSAPI_KEY = "3087034a13564f75bfc769c0046e729c"
NEWSAPI_URL = "https://newsapi.org/v2/everything"

# ==========================================================
# STRONG DATA CENTER KEYWORDS
# ==========================================================
KEYWORDS = [
    "data center",
    "data centre",
    "hyperscale",
    "colocation",
    "AI infrastructure",
    "AI data center",
    "data center campus",
    "server farm",
    "cloud campus",
    "digital infrastructure",
    "GPU cluster",
    "AI cluster",
    "compute campus",
    "cloud infrastructure",
    "edge data center",
    "data center construction",
    "data center investment",
    "data center expansion",
    "data center development",
    "substation",
    "power infrastructure",
    "switchyard",
    "gigawatt campus",
    "utility infrastructure",
    "server infrastructure",
    "rack density",
    "liquid cooling",
    "high-density compute",
    "AI compute",
    "hyperscale campus",
    "mega campus",
    "compute facility"
]

# ==========================================================
# US MARKET TERMS
# ==========================================================
US_TERMS = [
    "USA",
    "United States",
    "Texas",
    "Virginia",
    "California",
    "Arizona",
    "Ohio",
    "Georgia",
    "Illinois",
    "Nevada",
    "Oregon",
    "Washington",
    "Utah",
    "North Carolina",
    "New York"
]

# ==========================================================
# TRUSTED DATA CENTER SOURCES
# ==========================================================
TRUSTED_SOURCES = {
    "DataCenterDynamics": "https://www.datacenterdynamics.com/en/news/",
    "DataCenterKnowledge": "https://www.datacenterknowledge.com/",
    "DataCentreMagazine": "https://datacentremagazine.com/",
    "CapacityMedia": "https://www.capacitymedia.com/",
    "DCD": "https://www.datacenterdynamics.com/",
    "BizJournal": "https://www.bizjournals.com/",
    "Reuters": "https://www.reuters.com/",
    "Bloomberg": "https://www.bloomberg.com/",
    "CRN": "https://www.crn.com/",
    "TechTarget": "https://www.techtarget.com/searchdatacenter/",
    "TheRegister": "https://www.theregister.com/data_centre/",
    "NetworkWorld": "https://www.networkworld.com/",
    "DataCenterFrontier": "https://datacenterfrontier.com/"
}

# ==========================================================
# STRICT US DATA CENTER FILTER
# ==========================================================
def is_us_data_center_article(text):

    text = str(text).lower()

    # MUST CONTAIN DATA CENTER TERMS
    if not any(k.lower() in text for k in KEYWORDS):
        return False

    # MUST CONTAIN US TERMS
    if not any(u.lower() in text for u in US_TERMS):
        return False

    # REMOVE NOISE
    blacklist = [
        "sports",
        "movie",
        "celebrity",
        "football",
        "basketball",
        "music",
        "war",
        "politics",
        "election",
        "hollywood",
        "netflix",
        "streaming show",
        "actor"
    ]

    if any(b in text for b in blacklist):
        return False

    return True

# ==========================================================
# GENERATE QUERY
# ==========================================================
def build_query():

    keyword_query = " OR ".join(
        [f'"{k}"' for k in KEYWORDS]
    )

    us_query = " OR ".join(
        [f'"{u}"' for u in US_TERMS]
    )

    final_query = f"({keyword_query}) AND ({us_query})"

    return final_query

# ==========================================================
# NEWS API FETCH
# ==========================================================
def fetch_newsapi_articles(start_date, end_date):

    query = build_query()

    params = {
        "q": query,
        "from": start_date,
        "to": end_date,
        "language": "en",
        "sortBy": "publishedAt",
        "pageSize": 200,
        "searchIn": "title,description",
        "apiKey": NEWSAPI_KEY
    }

    try:

        res = requests.get(
            NEWSAPI_URL,
            params=params,
            timeout=20
        )

        if res.status_code != 200:
            return pd.DataFrame()

        data = res.json()

        articles = data.get("articles", [])

        rows = []

        for a in articles:

            title = a.get("title", "")
            desc = a.get("description", "")

            combined_text = f"{title} {desc}"

            if not is_us_data_center_article(combined_text):
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
                "Link": a.get("url"),
                "Description": desc,
                "Source": a.get("source", {}).get("name"),
                "Published Date": date_value
            })

        return pd.DataFrame(rows)

    except:
        return pd.DataFrame()

# ==========================================================
# GENERIC SCRAPER
# ==========================================================
def scrape_website(source_name, url):

    try:

        headers = {
            "User-Agent": "Mozilla/5.0"
        }

        res = requests.get(
            url,
            headers=headers,
            timeout=20
        )

        soup = BeautifulSoup(
            res.text,
            "html.parser"
        )

        articles = []

        links = soup.find_all("a")

        seen = set()

        for link in links:

            title = link.get_text(strip=True)

            href = link.get("href")

            if not href:
                continue

            if len(title) < 25:
                continue

            text = title.lower()

            if not is_us_data_center_article(text):
                continue

            if href in seen:
                continue

            seen.add(href)

            if href.startswith("/"):
                base = "/".join(url.split("/")[:3])
                href = base + href

            articles.append({
                "Title": title,
                "Link": href,
                "Description": f"US data center news from {source_name}",
                "Source": source_name,
                "Published Date": datetime.now().strftime("%Y-%m-%d")
            })

        return pd.DataFrame(articles)

    except:
        return pd.DataFrame()

# ==========================================================
# SCRAPE ALL SOURCES
# ==========================================================
def scrape_all_sources():

    all_dfs = []

    for source_name, url in TRUSTED_SOURCES.items():

        df = scrape_website(
            source_name,
            url
        )

        if not df.empty:
            all_dfs.append(df)

    if all_dfs:
        return pd.concat(
            all_dfs,
            ignore_index=True
        )

    return pd.DataFrame()

# ==========================================================
# FINAL CLEANING
# ==========================================================
def clean_dataframe(df):

    if df.empty:
        return df

    # REMOVE DUPLICATES
    df.drop_duplicates(
        subset="Title",
        inplace=True
    )

    # REMOVE EMPTY TITLES
    df = df[
        df["Title"].notna()
    ]

    # REMOVE SHORT TITLES
    df = df[
        df["Title"].str.len() > 20
    ]

    # SORT
    if "Published Date" in df.columns:

        try:
            df["Published Date"] = pd.to_datetime(
                df["Published Date"],
                errors="coerce"
            )

            df = df.sort_values(
                by="Published Date",
                ascending=False
            )

            df["Published Date"] = df[
                "Published Date"
            ].dt.strftime("%Y-%m-%d")

        except:
            pass

    return df

# ==========================================================
# EXPORT
# ==========================================================
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

# ==========================================================
# UI
# ==========================================================
def main():

    st.set_page_config(
        page_title="US Data Center Intelligence Tracker",
        layout="wide"
    )

    st.title(
        "🏢 US Data Center Intelligence Tracker"
    )

    st.markdown("""
    ### Track:
    - Hyperscale campuses
    - AI infrastructure
    - Data center developments
    - Colocation projects
    - Utility + power infrastructure
    - US compute campus expansions
    """)

    # ======================================================
    # SIDEBAR
    # ======================================================
    st.sidebar.header("Filters")

    time_filter = st.sidebar.radio(
        "Select Time Range",
        [
            "Latest",
            "Past 10 Days",
            "Past 30 Days"
        ]
    )

    # ======================================================
    # DATE LOGIC
    # ======================================================
    end_date = datetime.today()

    if time_filter == "Latest":
        start_date = end_date - timedelta(days=1)

    elif time_filter == "Past 10 Days":
        start_date = end_date - timedelta(days=10)

    else:
        start_date = end_date - timedelta(days=30)

    # ======================================================
    # FETCH
    # ======================================================
    if st.button("🚀 Fetch Maximum US Data Center News"):

        with st.spinner(
            "Collecting US Data Center Intelligence..."
        ):

            # NEWS API
            df_api = fetch_newsapi_articles(
                start_date.strftime("%Y-%m-%d"),
                end_date.strftime("%Y-%m-%d")
            )

            # SCRAPERS
            df_scraped = scrape_all_sources()

            # COMBINE
            final_df = pd.concat(
                [df_api, df_scraped],
                ignore_index=True
            )

            # CLEAN
            final_df = clean_dataframe(
                final_df
            )

            # ==================================================
            # OUTPUT
            # ==================================================
            if not final_df.empty:

                st.success(
                    f"{len(final_df)} US Data Center Articles Found"
                )

                st.dataframe(
                    final_df,
                    use_container_width=True
                )

                # DOWNLOAD
                excel = to_excel(final_df)

                st.download_button(
                    "📥 Download Excel",
                    excel,
                    "us_data_center_intelligence.xlsx"
                )

            else:

                st.warning(
                    "No US data center articles found"
                )

# ==========================================================
# MAIN
# ==========================================================
if __name__ == "__main__":
    main()
