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
# MULTIPLE QUERIES
# ==========================================================
SEARCH_QUERIES = [

    '"data center" USA',
    '"hyperscale" USA',
    '"colocation" USA',
    '"AI infrastructure" USA',
    '"data center campus" USA',
    '"server farm" USA',
    '"cloud campus" USA',
    '"digital infrastructure" USA',
    '"compute campus" USA',
    '"AI data center" USA',
    '"GPU cluster" USA',
    '"data center expansion" USA',
    '"data center development" USA',
    '"data center construction" USA',
    '"data center investment" USA',
    '"cloud infrastructure" USA'
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
        "https://www.bizjournals.com/"
}

# ==========================================================
# DATA CENTER FILTER
# ==========================================================
def is_data_center_article(text):

    text = str(text).lower()

    keywords = [
        "data center",
        "data centre",
        "hyperscale",
        "colocation",
        "server farm",
        "campus",
        "cloud infrastructure",
        "compute",
        "gpu cluster",
        "ai infrastructure"
    ]

    return any(k in text for k in keywords)

# ==========================================================
# NEWSAPI FETCH
# ==========================================================
def fetch_newsapi(start_date, end_date):

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
                "searchIn": "title",
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

                    if not is_data_center_article(title):
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
                            "source", {}
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

            if len(title) < 20:
                continue

            if not is_data_center_article(title):
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
# SCRAPE ALL
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
# CLEAN
# ==========================================================
def clean_dataframe(df):

    if df.empty:
        return df

    df.drop_duplicates(
        subset="Title",
        inplace=True
    )

    df = df[
        df["Title"].str.len() > 20
    ]

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
        page_title=
            "US Data Center Intelligence",
        layout="wide"
    )

    st.title(
        "🏢 US Data Center Intelligence Tracker"
    )

    st.markdown("""
    Track:
    - Hyperscale developments
    - AI infrastructure
    - Data center campuses
    - Colocation expansions
    - Compute infrastructure
    """)

    st.sidebar.header("Filters")

    time_filter = st.sidebar.radio(
        "Select Time Range",
        [
            "Latest",
            "Past 10 Days",
            "Past 30 Days"
        ]
    )

    end_date = datetime.today()

    if time_filter == "Latest":
        start_date = end_date - timedelta(days=1)

    elif time_filter == "Past 10 Days":
        start_date = end_date - timedelta(days=10)

    else:
        start_date = end_date - timedelta(days=30)

    if st.button(
        "🚀 Fetch Maximum Articles"
    ):

        with st.spinner(
            "Collecting Data Center Intelligence..."
        ):

            # NEWSAPI
            df_news = fetch_newsapi(
                start_date.strftime("%Y-%m-%d"),
                end_date.strftime("%Y-%m-%d")
            )

            # SCRAPERS
            df_scraped = scrape_all_sources()

            # COMBINE
            final_df = pd.concat(
                [df_news, df_scraped],
                ignore_index=True
            )

            # CLEAN
            final_df = clean_dataframe(
                final_df
            )

            if not final_df.empty:

                st.success(
                    f"{len(final_df)} Articles Found"
                )

                st.dataframe(
                    final_df,
                    use_container_width=True
                )

                excel = to_excel(
                    final_df
                )

                st.download_button(
                    "📥 Download Excel",
                    excel,
                    "us_data_center_articles.xlsx"
                )

            else:

                st.warning(
                    "No articles found"
                )

# ==========================================================
# MAIN
# ==========================================================
if __name__ == "__main__":
    main()
