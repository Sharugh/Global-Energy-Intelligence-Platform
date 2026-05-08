import streamlit as st
import requests
import pandas as pd
from datetime import datetime, timedelta
from io import BytesIO
from bs4 import BeautifulSoup

# ==========================================================
# CONFIG
# ==========================================================
NEWSAPI_KEY = "3087034a13564f75bfc769c0046e729c"
NEWSAPI_URL = "https://newsapi.org/v2/everything"

# ==========================================================
# DEVELOPMENT-FOCUSED SEARCH QUERIES
# ==========================================================
SEARCH_QUERIES = [

    '"data center" construction USA',
    '"data center" investment USA',
    '"data center" expansion USA',
    '"data center campus" USA',
    '"hyperscale campus" USA',
    '"data center" approved USA',
    '"data center" permit USA',
    '"data center" planning USA',
    '"data center" groundbreaking USA',
    '"data center" proposed USA',
    '"data center" zoning USA',
    '"data center" land acquisition USA',
    '"AI data center campus" USA',
    '"colocation facility" expansion USA',
    '"cloud campus" USA',
    '"digital infrastructure campus" USA',
    '"compute campus" USA',
    '"data center development" USA',
    '"server farm" construction USA',
    '"utility infrastructure" data center USA',
    '"substation" data center USA',
    '"gigawatt campus" USA',
    '"data center" utility approval USA',
    '"data center" power agreement USA',
    '"data center" site selection USA',
    '"data center" new facility USA',
    '"data center" campus expansion USA',
    '"data center" build USA',
    '"data center" construction permit USA',
    '"data center" infrastructure investment USA'
]

# ==========================================================
# TRUSTED SOURCES
# ==========================================================
SCRAPE_SOURCES = {

    "DataCenterDynamics":
        "https://www.datacenterdynamics.com/en/news/",

    "DataCenterKnowledge":
        "https://www.datacenterknowledge.com/",

    "DataCentreMagazine":
        "https://datacentremagazine.com/",

    "DataCenterFrontier":
        "https://datacenterfrontier.com/",

    "CapacityMedia":
        "https://www.capacitymedia.com/",

    "TechTarget":
        "https://www.techtarget.com/searchdatacenter/",

    "TheRegister":
        "https://www.theregister.com/data_centre/",

    "CRN":
        "https://www.crn.com/news/data-center",

    "NetworkWorld":
        "https://www.networkworld.com/",

    "BizJournal":
        "https://www.bizjournals.com/",

    "Reuters":
        "https://www.reuters.com/",

    "Bloomberg":
        "https://www.bloomberg.com/"
}

# ==========================================================
# DEVELOPMENT FILTER
# ==========================================================
def is_relevant_article(text):

    text = str(text).lower()

    # MUST HAVE DATA CENTER TERMS
    dc_terms = [

        "data center",
        "data centre",
        "hyperscale",
        "colocation",
        "server farm",
        "compute campus",
        "cloud campus",
        "ai infrastructure"
    ]

    if not any(k in text for k in dc_terms):
        return False

    # MUST HAVE DEVELOPMENT SIGNAL
    development_terms = [

        "investment",
        "construction",
        "expansion",
        "approved",
        "approval",
        "permit",
        "planned",
        "planning",
        "campus",
        "groundbreaking",
        "development",
        "proposal",
        "proposed",
        "zoning",
        "site plan",
        "land acquisition",
        "build",
        "built",
        "new facility",
        "substation",
        "power infrastructure",
        "opening",
        "launches",
        "announces",
        "acquires land",
        "utility agreement",
        "power agreement",
        "utility approval",
        "infrastructure expansion",
        "site selection"
    ]

    if not any(d in text for d in development_terms):
        return False

    # REMOVE GENERIC TECH NEWS
    blacklist = [

        "gpu launch",
        "earnings",
        "financial results",
        "quarterly results",
        "software update",
        "cpu",
        "processor",
        "gaming",
        "sports",
        "movie",
        "celebrity",
        "stock market",
        "share price",
        "advertisement",
        "cloud revenue",
        "phone launch",
        "laptop",
        "smartphone"
    ]

    if any(b in text for b in blacklist):
        return False

    return True

# ==========================================================
# FETCH FROM NEWSAPI
# ==========================================================
def fetch_newsapi_articles(start_date, end_date):

    all_articles = []

    for query in SEARCH_QUERIES:

        for page in range(1, 6):

            params = {

                "q": query,
                "from": start_date,
                "to": end_date,
                "language": "en",
                "sortBy": "publishedAt",
                "pageSize": 100,
                "page": page,
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
                    continue

                data = res.json()

                articles = data.get("articles", [])

                if not articles:
                    break

                for a in articles:

                    title = a.get("title", "")
                    desc = a.get("description", "")

                    combined_text = f"{title} {desc}"

                    if not is_relevant_article(combined_text):
                        continue

                    date_value = a.get("publishedAt")

                    try:

                        date_value = pd.to_datetime(
                            date_value
                        ).strftime("%Y-%m-%d")

                    except:
                        pass

                    all_articles.append({

                        "Title": title,
                        "Link": a.get("url"),
                        "Source": a.get(
                            "source",
                            {}
                        ).get("name"),
                        "Published Date": date_value
                    })

            except:
                continue

    return pd.DataFrame(all_articles)

# ==========================================================
# SCRAPER
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

            if not is_relevant_article(title):
                continue

            if href in seen:
                continue

            seen.add(href)

            if href.startswith("/"):

                base = "/".join(
                    url.split("/")[:3]
                )

                href = base + href

            articles.append({

                "Title": title,
                "Link": href,
                "Source": source_name,
                "Published Date":
                    datetime.now().strftime("%Y-%m-%d")
            })

        return pd.DataFrame(articles)

    except:
        return pd.DataFrame()

# ==========================================================
# SCRAPE ALL SOURCES
# ==========================================================
def scrape_all_sources():

    dfs = []

    for source_name, url in SCRAPE_SOURCES.items():

        df = scrape_website(
            source_name,
            url
        )

        if not df.empty:
            dfs.append(df)

    if dfs:

        return pd.concat(
            dfs,
            ignore_index=True
        )

    return pd.DataFrame()

# ==========================================================
# CLEAN DATAFRAME
# ==========================================================
def clean_dataframe(df):

    if df.empty:
        return df

    # REMOVE DUPLICATES
    df.drop_duplicates(
        subset="Title",
        inplace=True
    )

    # REMOVE VERY SHORT TITLES
    df = df[
        df["Title"].str.len() > 20
    ]

    # SORT BY DATE
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
# EXPORT TO EXCEL
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
# STREAMLIT UI
# ==========================================================
def main():

    st.set_page_config(
        page_title=
            "US Data Center Development Tracker",
        layout="wide"
    )

    st.title(
        "🏗️ US Data Center Development Tracker"
    )

    st.markdown("""
    ### Tracks:
    - Data center campus developments
    - Hyperscale construction
    - AI infrastructure investments
    - Colocation expansions
    - Planning approvals
    - Land acquisitions
    - Groundbreaking announcements
    - Utility + power infrastructure
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
    # FETCH BUTTON
    # ======================================================
    if st.button(
        "🚀 Fetch Development Articles"
    ):

        with st.spinner(
            "Collecting US Data Center Development Intelligence..."
        ):

            # ==============================================
            # NEWS API
            # ==============================================
            df_news = fetch_newsapi_articles(

                start_date.strftime("%Y-%m-%d"),
                end_date.strftime("%Y-%m-%d")
            )

            # ==============================================
            # SCRAPED DATA
            # ==============================================
            df_scraped = scrape_all_sources()

            # ==============================================
            # COMBINE
            # ==============================================
            final_df = pd.concat(

                [df_news, df_scraped],
                ignore_index=True
            )

            # ==============================================
            # CLEAN
            # ==============================================
            final_df = clean_dataframe(
                final_df
            )

            # ==============================================
            # OUTPUT
            # ==============================================
            if not final_df.empty:

                st.success(
                    f"{len(final_df)} Development Articles Found"
                )

                st.dataframe(
                    final_df,
                    use_container_width=True
                )

                # ==========================================
                # DOWNLOAD
                # ==========================================
                excel = to_excel(
                    final_df
                )

                st.download_button(

                    "📥 Download Excel",

                    excel,

                    "us_data_center_development_tracker.xlsx"
                )

            else:

                st.warning(
                    "No development-related articles found"
                )

# ==========================================================
# MAIN
# ==========================================================
if __name__ == "__main__":
    main()
