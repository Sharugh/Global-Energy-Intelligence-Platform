import streamlit as st
import requests
import pandas as pd
from bs4 import BeautifulSoup
from datetime import datetime, timedelta
from io import BytesIO
import time
import re

# ==========================================================
# PAGE CONFIG
# ==========================================================
st.set_page_config(
    page_title="US Data Center Development Tracker",
    layout="wide"
)

# ==========================================================
# BASE CONFIG
# ==========================================================
BASE_URL = "https://www.datacenterdynamics.com"

SEARCH_URL = (
    "https://www.datacenterdynamics.com/en/news/"
    "?term=the-data-center-construction-channel&page={}"
)

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 "
        "(Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 "
        "(KHTML, like Gecko) "
        "Chrome/120.0 Safari/537.36"
    )
}

# ==========================================================
# US STATES
# ==========================================================
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
    "North Carolina",
    "Washington",
    "Oregon",
    "Utah",
    "Florida",
    "Indiana",
    "Louisiana",
    "Tennessee",
    "Pennsylvania",
    "New Jersey",
    "Maryland",
    "Massachusetts",
    "South Carolina",
    "Alabama",
    "Kentucky",
    "Wisconsin",
    "Mississippi",
    "New Mexico",
    "Iowa"
]

# ==========================================================
# DEVELOPMENT KEYWORDS
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
DATA_CENTER_TERMS = [

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
# RELEVANT ARTICLE CHECK
# ==========================================================
def is_relevant_article(text, selected_states):

    text = str(text).lower()

    # MUST CONTAIN DATA CENTER TERM
    if not any(
        term.lower() in text
        for term in DATA_CENTER_TERMS
    ):
        return False

    # MUST CONTAIN DEVELOPMENT TERM
    if not any(
        term.lower() in text
        for term in DEVELOPMENT_KEYWORDS
    ):
        return False

    # USA FILTER
    usa_terms = [
        "usa",
        "united states"
    ]

    if selected_states:

        if not any(
            state.lower() in text
            for state in selected_states
        ):

            # fallback USA
            if not any(
                u in text for u in usa_terms
            ):
                return False

    return True

# ==========================================================
# EXTRACT DATE
# ==========================================================
def parse_date(date_text):

    try:

        parsed = pd.to_datetime(
            date_text,
            errors="coerce"
        )

        if pd.isna(parsed):
            return None

        return parsed

    except:
        return None

# ==========================================================
# SCRAPE SINGLE PAGE
# ==========================================================
def scrape_page(page_num):

    try:

        url = SEARCH_URL.format(page_num)

        response = requests.get(
            url,
            headers=HEADERS,
            timeout=30
        )

        if response.status_code != 200:
            return []

        soup = BeautifulSoup(
            response.text,
            "html.parser"
        )

        articles = []

        # FIND ALL ARTICLE BLOCKS
        article_blocks = soup.find_all(
            "article"
        )

        for block in article_blocks:

            try:

                # TITLE
                title_tag = block.find(
                    ["h2", "h3"]
                )

                if not title_tag:
                    continue

                title = title_tag.get_text(
                    strip=True
                )

                # LINK
                link_tag = block.find("a")

                if not link_tag:
                    continue

                href = link_tag.get("href")

                if not href:
                    continue

                if href.startswith("/"):

                    href = BASE_URL + href

                # DATE
                time_tag = block.find("time")

                if time_tag:

                    raw_date = (
                        time_tag.get_text(
                            strip=True
                        )
                    )

                    parsed_date = parse_date(
                        raw_date
                    )

                else:

                    parsed_date = None

                # DESCRIPTION
                desc_tag = block.find("p")

                description = ""

                if desc_tag:

                    description = (
                        desc_tag.get_text(
                            strip=True
                        )
                    )

                articles.append({

                    "title": title,
                    "link": href,
                    "description": description,
                    "date": parsed_date
                })

            except:
                continue

        return articles

    except:
        return []

# ==========================================================
# SCRAPE MULTIPLE PAGES
# ==========================================================
def scrape_dcd_articles(

    selected_states,
    start_date,
    end_date,
    max_pages=50
):

    collected = []

    progress_bar = st.progress(0)

    for page in range(1, max_pages + 1):

        progress_bar.progress(
            page / max_pages
        )

        page_articles = scrape_page(page)

        if not page_articles:
            continue

        for art in page_articles:

            try:

                title = art["title"]
                desc = art["description"]

                combined_text = (
                    f"{title} {desc}"
                )

                # RELEVANCE FILTER
                if not is_relevant_article(
                    combined_text,
                    selected_states
                ):
                    continue

                # DATE FILTER
                article_date = art["date"]

                if article_date:

                    article_date = (
                        article_date.replace(
                            tzinfo=None
                        )
                    )

                    if article_date < start_date:
                        continue

                    if article_date > end_date:
                        continue

                    formatted_date = (
                        article_date.strftime(
                            "%Y-%m-%d"
                        )
                    )

                else:

                    formatted_date = ""

                collected.append({

                    "Title": title,
                    "Link": art["link"],
                    "Source":
                        "DataCenterDynamics",
                    "Published Date":
                        formatted_date
                })

            except:
                continue

        time.sleep(1)

    progress_bar.empty()

    df = pd.DataFrame(collected)

    return df

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

    # REMOVE EMPTY TITLES
    df = df[
        df["Title"].notna()
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

    # REMOVE TIMEZONE ISSUES
    for col in df.columns:

        if pd.api.types.is_datetime64_any_dtype(
            df[col]
        ):

            df[col] = df[col].dt.tz_localize(
                None
            )

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
    - Data center construction
    - Campus developments
    - Hyperscale investments
    - Planning approvals
    - Groundbreaking
    - Land acquisitions
    - Utility infrastructure
    - Expansion projects
    """)

    # ======================================================
    # SIDEBAR
    # ======================================================
    st.sidebar.header("Filters")

    # TIME FILTER
    time_filter = st.sidebar.radio(

        "Select Time Range",

        [
            "Latest",
            "Past 10 Days",
            "Past 30 Days",
            "Custom Date"
        ]
    )

    today = datetime.now()

    if time_filter == "Latest":

        start_date = today - timedelta(days=1)
        end_date = today

    elif time_filter == "Past 10 Days":

        start_date = today - timedelta(days=10)
        end_date = today

    elif time_filter == "Past 30 Days":

        start_date = today - timedelta(days=30)
        end_date = today

    else:

        start_date = st.sidebar.date_input(
            "Start Date",
            today - timedelta(days=30)
        )

        end_date = st.sidebar.date_input(
            "End Date",
            today
        )

        start_date = datetime.combine(
            start_date,
            datetime.min.time()
        )

        end_date = datetime.combine(
            end_date,
            datetime.max.time()
        )

    # STATE FILTER
    selected_states = st.sidebar.multiselect(

        "Select USA States",

        options=US_STATES,

        default=[]
    )

    # PAGE SCAN
    max_pages = st.sidebar.slider(

        "Pages to Scan",

        min_value=5,
        max_value=100,
        value=30
    )

    # ======================================================
    # FETCH BUTTON
    # ======================================================
    if st.button(
        "🚀 Fetch Development Articles"
    ):

        with st.spinner(
            "Scanning DataCenterDynamics..."
        ):

            df = scrape_dcd_articles(

                selected_states=
                    selected_states,

                start_date=
                    start_date,

                end_date=
                    end_date,

                max_pages=
                    max_pages
            )

            df = clean_dataframe(df)

            # ==================================================
            # OUTPUT
            # ==================================================
            if not df.empty:

                st.success(
                    f"{len(df)} Articles Found"
                )

                st.dataframe(
                    df,
                    use_container_width=True
                )

                # DOWNLOAD
                excel = to_excel(df)

                st.download_button(

                    "📥 Download Excel",

                    excel,

                    "us_data_center_developments.xlsx"
                )

            else:

                st.warning(
                    "No matching US development articles found"
                )

# ==========================================================
# MAIN
# ==========================================================
if __name__ == "__main__":
    main()
