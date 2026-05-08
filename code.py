import streamlit as st
import pandas as pd
import feedparser
from datetime import datetime, timedelta
from io import BytesIO

# ==========================================================
# PAGE CONFIG
# ==========================================================
st.set_page_config(
    page_title="US Data Center Development Tracker",
    layout="wide"
)

# ==========================================================
# RSS FEEDS
# ==========================================================
RSS_FEEDS = {

    "DataCenterDynamics":
    "https://www.datacenterdynamics.com/en/rss/",

    "DataCenterKnowledge":
    "https://www.datacenterknowledge.com/rss.xml",

    "DataCentreMagazine":
    "https://datacentremagazine.com/rss",

    "DataCenterFrontier":
    "https://datacenterfrontier.com/feed/",

    "CapacityMedia":
    "https://www.capacitymedia.com/rss.xml"
}

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
    "server farm",
    "digital infrastructure"
]

# ==========================================================
# DEVELOPMENT TERMS
# ==========================================================
DEVELOPMENT_TERMS = [

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
# USA STATES
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
# RELEVANT ARTICLE FILTER
# ==========================================================
def is_relevant_article(text, selected_states):

    text = str(text).lower()

    # MUST HAVE DATA CENTER TERM
    has_dc_term = any(
        term.lower() in text
        for term in DATA_CENTER_TERMS
    )

    if not has_dc_term:
        return False

    # MUST HAVE DEVELOPMENT TERM
    has_dev_term = any(
        term.lower() in text
        for term in DEVELOPMENT_TERMS
    )

    if not has_dev_term:
        return False

    # USA FILTER
    usa_terms = [
        "usa",
        "united states",
        "u.s."
    ]

    if selected_states:

        state_match = any(
            state.lower() in text
            for state in selected_states
        )

        usa_match = any(
            term in text
            for term in usa_terms
        )

        if not state_match and not usa_match:
            return False

    return True

# ==========================================================
# DATE PARSER
# ==========================================================
def parse_date(entry):

    try:

        if hasattr(entry, "published_parsed"):

            parsed = datetime(*entry.published_parsed[:6])

            return parsed

    except:
        pass

    return None

# ==========================================================
# FETCH RSS ARTICLES
# ==========================================================
def fetch_rss_articles(

    selected_states,
    start_date,
    end_date
):

    all_articles = []

    progress = st.progress(0)

    total_feeds = len(RSS_FEEDS)

    current = 0

    for source, url in RSS_FEEDS.items():

        current += 1

        progress.progress(current / total_feeds)

        try:

            feed = feedparser.parse(url)

            entries = feed.entries

            for entry in entries:

                try:

                    title = entry.get(
                        "title",
                        ""
                    )

                    summary = entry.get(
                        "summary",
                        ""
                    )

                    link = entry.get(
                        "link",
                        ""
                    )

                    combined_text = (
                        f"{title} {summary}"
                    )

                    # FILTER ARTICLES
                    if not is_relevant_article(
                        combined_text,
                        selected_states
                    ):
                        continue

                    # DATE
                    article_date = parse_date(
                        entry
                    )

                    if article_date:

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

                    all_articles.append({

                        "Title": title,

                        "Link": link,

                        "Source": source,

                        "Published Date":
                            formatted_date
                    })

                except:
                    continue

        except:
            continue

    progress.empty()

    return pd.DataFrame(all_articles)

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
    ### Tracks:
    - Data center construction
    - Hyperscale developments
    - AI infrastructure campuses
    - Planning approvals
    - Groundbreaking projects
    - Utility infrastructure
    - Land acquisitions
    - Expansion announcements
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

        custom_start = st.sidebar.date_input(
            "Start Date",
            today - timedelta(days=30)
        )

        custom_end = st.sidebar.date_input(
            "End Date",
            today
        )

        start_date = datetime.combine(
            custom_start,
            datetime.min.time()
        )

        end_date = datetime.combine(
            custom_end,
            datetime.max.time()
        )

    # ======================================================
    # STATE FILTER
    # ======================================================
    selected_states = st.sidebar.multiselect(

        "Select USA States",

        options=US_STATES,

        default=[]
    )

    # ======================================================
    # FETCH BUTTON
    # ======================================================
    if st.button(
        "🚀 Fetch Development Articles"
    ):

        with st.spinner(
            "Collecting US Data Center Development News..."
        ):

            df = fetch_rss_articles(

                selected_states=
                    selected_states,

                start_date=
                    start_date,

                end_date=
                    end_date
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
                    "No matching development articles found"
                )

# ==========================================================
# MAIN
# ==========================================================
if __name__ == "__main__":
    main()
