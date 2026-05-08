import streamlit as st
import requests
import pandas as pd
from bs4 import BeautifulSoup
from datetime import datetime, timedelta
from io import BytesIO
import time

# ==========================================================
# PAGE CONFIG
# ==========================================================
st.set_page_config(
    page_title="US Data Center Development Tracker",
    layout="wide"
)

# ==========================================================
# CONFIG
# ==========================================================
BASE_URL = "https://www.datacenterdynamics.com"

DCD_CONSTRUCTION_URL = (
    "https://www.datacenterdynamics.com/en/news/"
    "?term=the-data-center-construction-channel"
)

# ==========================================================
# US FILTER KEYWORDS
# ==========================================================
US_KEYWORDS = [

    "usa",
    "united states",

    "texas",
    "virginia",
    "california",
    "arizona",
    "ohio",
    "georgia",
    "illinois",
    "new york",
    "oregon",
    "washington",
    "utah",
    "north carolina",
    "nevada",
    "florida",
    "tennessee",
    "indiana",
    "louisiana",
    "new mexico",
    "mississippi",
    "iowa",
    "wisconsin",
    "pennsylvania",
    "new jersey",
    "massachusetts",
    "maryland",
    "south carolina",
    "alabama",
    "kentucky"
]

# ==========================================================
# DEVELOPMENT SIGNALS
# ==========================================================
DEVELOPMENT_KEYWORDS = [

    "construction",
    "investment",
    "expansion",
    "campus",
    "approved",
    "approval",
    "permit",
    "planning",
    "planned",
    "groundbreaking",
    "development",
    "proposal",
    "proposed",
    "land acquisition",
    "site selection",
    "build",
    "built",
    "new facility",
    "substation",
    "utility infrastructure",
    "power agreement",
    "utility approval",
    "infrastructure",
    "announces",
    "acquires land",
    "construction starts",
    "opening",
    "zoning",
    "site plan"
]

# ==========================================================
# DATA CENTER TERMS
# ==========================================================
DC_TERMS = [

    "data center",
    "data centre",
    "hyperscale",
    "colocation",
    "ai infrastructure",
    "cloud campus",
    "compute campus",
    "server farm"
]

# ==========================================================
# FILTER FUNCTION
# ==========================================================
def is_relevant_article(text):

    text = str(text).lower()

    # MUST HAVE DATA CENTER TERM
    if not any(k in text for k in DC_TERMS):
        return False

    # MUST HAVE DEVELOPMENT TERM
    if not any(k in text for k in DEVELOPMENT_KEYWORDS):
        return False

    # MUST HAVE US TERM
    if not any(k in text for k in US_KEYWORDS):
        return False

    # REMOVE IRRELEVANT NEWS
    blacklist = [

        "earnings",
        "financial results",
        "sports",
        "movie",
        "gaming",
        "celebrity",
        "processor",
        "gpu launch",
        "software update",
        "smartphone",
        "laptop",
        "cpu"
    ]

    if any(k in text for k in blacklist):
        return False

    return True

# ==========================================================
# EXTRACT ARTICLE DATE
# ==========================================================
def extract_date(date_text):

    try:

        parsed_date = pd.to_datetime(
            date_text,
            errors="coerce"
        )

        if pd.isna(parsed_date):
            return None

        return parsed_date

    except:
        return None

# ==========================================================
# SCRAPE DCD MULTIPLE PAGES
# ==========================================================
def scrape_dcd(max_pages=15, days_limit=30):

    all_articles = []

    cutoff_date = datetime.now() - timedelta(days=days_limit)

    headers = {
        "User-Agent": "Mozilla/5.0"
    }

    progress_bar = st.progress(0)

    for page in range(1, max_pages + 1):

        try:

            progress_bar.progress(
                page / max_pages
            )

            page_url = (
                f"{DCD_CONSTRUCTION_URL}&page={page}"
            )

            response = requests.get(

                page_url,
                headers=headers,
                timeout=20
            )

            if response.status_code != 200:
                continue

            soup = BeautifulSoup(
                response.text,
                "html.parser"
            )

            articles = soup.find_all("article")

            if not articles:
                continue

            for article in articles:

                try:

                    # ==================================
                    # TITLE
                    # ==================================
                    title_tag = article.find("h3")

                    if not title_tag:
                        continue

                    title = title_tag.get_text(
                        strip=True
                    )

                    # ==================================
                    # LINK
                    # ==================================
                    link_tag = article.find("a")

                    if not link_tag:
                        continue

                    href = link_tag.get("href")

                    if not href:
                        continue

                    if href.startswith("/"):

                        full_link = (
                            BASE_URL + href
                        )

                    else:

                        full_link = href

                    # ==================================
                    # DATE
                    # ==================================
                    time_tag = article.find("time")

                    if time_tag:

                        date_text = (
                            time_tag.get_text(
                                strip=True
                            )
                        )

                        parsed_date = extract_date(
                            date_text
                        )

                    else:

                        parsed_date = datetime.now()

                    # ==================================
                    # DATE FILTER
                    # ==================================
                    if parsed_date:

                        if parsed_date < cutoff_date:
                            continue

                    # ==================================
                    # DESCRIPTION
                    # ==================================
                    desc_tag = article.find("p")

                    description = ""

                    if desc_tag:
                        description = desc_tag.get_text(
                            strip=True
                        )

                    # ==================================
                    # COMBINED TEXT
                    # ==================================
                    combined_text = (
                        f"{title} {description}"
                    )

                    # ==================================
                    # FILTER
                    # ==================================
                    if not is_relevant_article(
                        combined_text
                    ):
                        continue

                    # ==================================
                    # ADD ARTICLE
                    # ==================================
                    all_articles.append({

                        "Title": title,

                        "Link": full_link,

                        "Source":
                            "DataCenterDynamics",

                        "Published Date":
                            parsed_date.strftime(
                                "%Y-%m-%d"
                            )
                    })

                except:
                    continue

            time.sleep(1)

        except:
            continue

    progress_bar.empty()

    return pd.DataFrame(all_articles)

# ==========================================================
# CLEAN DATAFRAME
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
# UI
# ==========================================================
def main():

    st.title(
        "🏗️ US Data Center Development Tracker"
    )

    st.markdown("""
    ### Tracks ONLY:
    - US data center construction
    - Hyperscale campus developments
    - AI infrastructure projects
    - Planning approvals
    - Investments
    - Groundbreaking announcements
    - Utility + power infrastructure
    - Campus expansions
    """)

    # ======================================================
    # SIDEBAR
    # ======================================================
    st.sidebar.header("Filters")

    time_range = st.sidebar.radio(

        "Select Time Range",

        [
            "Latest",
            "Past 10 Days",
            "Past 30 Days"
        ]
    )

    if time_range == "Latest":
        days_limit = 1

    elif time_range == "Past 10 Days":
        days_limit = 10

    else:
        days_limit = 30

    max_pages = st.sidebar.slider(

        "Pages to Scan",

        min_value=5,
        max_value=50,
        value=20
    )

    # ======================================================
    # FETCH BUTTON
    # ======================================================
    if st.button(
        "🚀 Fetch US Development Articles"
    ):

        with st.spinner(
            "Scanning DataCenterDynamics..."
        ):

            df = scrape_dcd(

                max_pages=max_pages,
                days_limit=days_limit
            )

            df = clean_dataframe(df)

            # ==============================================
            # OUTPUT
            # ==============================================
            if not df.empty:

                st.success(
                    f"{len(df)} US Development Articles Found"
                )

                st.dataframe(
                    df,
                    use_container_width=True
                )

                # ==========================================
                # DOWNLOAD
                # ==========================================
                excel = to_excel(df)

                st.download_button(

                    "📥 Download Excel",

                    excel,

                    "us_data_center_developments.xlsx"
                )

            else:

                st.warning(
                    "No US development-related articles found"
                )

# ==========================================================
# MAIN
# ==========================================================
if __name__ == "__main__":
    main()
