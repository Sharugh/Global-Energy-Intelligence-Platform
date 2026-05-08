import streamlit as st
import re
import io
import time
from datetime import datetime, timedelta

import pandas as pd
import plotly.express as px

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# =========================================================
# REQUEST ENGINE
# =========================================================
try:
    import cloudscraper

    scraper = cloudscraper.create_scraper(
        browser={
            "browser": "chrome",
            "platform": "windows",
            "mobile": False
        }
    )

except:
    import requests
    scraper = requests.Session()

from bs4 import BeautifulSoup

# =========================================================
# PAGE CONFIG
# =========================================================
st.set_page_config(
    page_title="US Data Center Intelligence Tracker",
    page_icon="🏗️",
    layout="wide"
)

# =========================================================
# SOURCES
# =========================================================
SOURCES = {

    "DCD Construction":
    "https://www.datacenterdynamics.com/en/news/?term=the-data-center-construction-channel&page={}",

    "DCD North America":
    "https://www.datacenterdynamics.com/en/news/?term=north-america&page={}",

    "DataCenterKnowledge":
    "https://www.datacenterknowledge.com/search?search_api_fulltext=data+center&page={}",

    "DataCenterFrontier":
    "https://datacenterfrontier.com/page/{}/",

    "DataCentreMagazine":
    "https://datacentremagazine.com/search?search=data+center&page={}"
}

# =========================================================
# KEYWORDS
# =========================================================
DC_TERMS = [
    "data center",
    "data centre",
    "hyperscale",
    "colocation",
    "ai infrastructure",
    "server farm",
    "digital infrastructure",
    "cloud campus"
]

PROJECT_TERMS = [
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
    "construction starts",
    "opening",
    "zoning",
    "site plan",
    "announced",
    "pipeline"
]

COMPANIES = [
    "AWS",
    "Amazon",
    "Microsoft",
    "Google",
    "Meta",
    "Oracle",
    "Equinix",
    "Digital Realty",
    "QTS",
    "CyrusOne",
    "Aligned",
    "CoreSite",
    "Compass Datacenters",
    "Stack Infrastructure",
    "NTT",
    "Vantage",
    "Switch",
    "Apple"
]

US_STATES = [
    "Texas","Virginia","California","Arizona","Ohio",
    "Georgia","Illinois","New York","Nevada",
    "North Carolina","Washington","Oregon",
    "Florida","Utah","Indiana","Tennessee",
    "Pennsylvania","New Jersey","Maryland"
]

# =========================================================
# REGEX
# =========================================================
MW_RE = re.compile(r"\b\d+(\.\d+)?\s?(MW|GW)\b", re.I)

COST_RE = re.compile(
    r"\$[\d\.]+\s?(million|billion|bn|m)",
    re.I
)

# =========================================================
# HEADERS
# =========================================================
HEADERS = {
    "User-Agent":
    "Mozilla/5.0"
}

# =========================================================
# FETCH URL
# =========================================================
@st.cache_data(ttl=3600)
def fetch_url(url):

    try:

        response = scraper.get(
            url,
            headers=HEADERS,
            timeout=30
        )

        if response.status_code == 200:

            return BeautifulSoup(
                response.text,
                "html.parser"
            )

    except:
        return None

    return None

# =========================================================
# PARSE DATE
# =========================================================
def parse_date(text):

    patterns = [

        "%Y-%m-%d",
        "%d %B %Y",
        "%B %d %Y",
        "%d %b %Y"
    ]

    for p in patterns:

        try:
            return datetime.strptime(text, p)
        except:
            pass

    return None

# =========================================================
# EXTRACT COMPANY
# =========================================================
def extract_company(text):

    found = []

    for c in COMPANIES:

        if c.lower() in text.lower():
            found.append(c)

    return ", ".join(found) if found else "Unknown"

# =========================================================
# EXTRACT STATE
# =========================================================
def extract_state(text):

    found = []

    for s in US_STATES:

        if s.lower() in text.lower():
            found.append(s)

    return ", ".join(found) if found else "Unknown"

# =========================================================
# EXTRACT MW
# =========================================================
def extract_mw(text):

    match = MW_RE.search(text)

    return match.group(0) if match else "N/A"

# =========================================================
# EXTRACT COST
# =========================================================
def extract_cost(text):

    match = COST_RE.search(text)

    return match.group(0) if match else "N/A"

# =========================================================
# PROJECT STAGE
# =========================================================
def classify_stage(text):

    text = text.lower()

    if "approved" in text:
        return "Approved"

    if "permit" in text:
        return "Permitting"

    if "construction" in text:
        return "Construction"

    if "groundbreaking" in text:
        return "Groundbreaking"

    if "expansion" in text:
        return "Expansion"

    if "planning" in text:
        return "Planning"

    if "investment" in text:
        return "Investment"

    return "General"

# =========================================================
# QUALITY SCORE
# =========================================================
def quality_score(text):

    score = 0

    if any(
        x.lower() in text.lower()
        for x in DC_TERMS
    ):
        score += 2

    if any(
        x.lower() in text.lower()
        for x in PROJECT_TERMS
    ):
        score += 2

    if extract_company(text) != "Unknown":
        score += 2

    if extract_state(text) != "Unknown":
        score += 2

    if extract_mw(text) != "N/A":
        score += 1

    if extract_cost(text) != "N/A":
        score += 1

    return score

# =========================================================
# SCRAPE ARTICLE PAGE
# =========================================================
def scrape_article_page(url):

    soup = fetch_url(url)

    if not soup:
        return ""

    return soup.get_text(
        " ",
        strip=True
    )

# =========================================================
# PARSE ARTICLES
# =========================================================
def parse_articles(soup, source):

    articles = []

    links = soup.find_all(
        "a",
        href=True
    )

    seen = set()

    for a in links:

        href = a["href"]

        title = a.get_text(
            strip=True
        )

        if len(title) < 20:
            continue

        if href in seen:
            continue

        seen.add(href)

        if href.startswith("/"):

            if "datacenterdynamics" in source:

                href = (
                    "https://www.datacenterdynamics.com"
                    + href
                )

        elif not href.startswith("http"):
            continue

        articles.append({
            "title": title,
            "url": href,
            "source": source
        })

    return articles

# =========================================================
# MAIN SCRAPER
# =========================================================
def scrape_sources(

    max_pages,
    selected_states,
    start_date

):

    final_data = []

    progress = st.progress(0)

    total_steps = (
        len(SOURCES) * max_pages
    )

    current_step = 0

    for source, template in SOURCES.items():

        for page in range(1, max_pages + 1):

            current_step += 1

            progress.progress(
                current_step / total_steps,
                text=f"Scanning {source} page {page}"
            )

            url = template.format(page)

            soup = fetch_url(url)

            if not soup:
                continue

            parsed_articles = parse_articles(
                soup,
                source
            )

            for article in parsed_articles:

                try:

                    article_text = scrape_article_page(
                        article["url"]
                    )

                    combined = (
                        article["title"]
                        + " "
                        + article_text
                    )

                    # FILTER DC TERMS
                    if not any(
                        x.lower() in combined.lower()
                        for x in DC_TERMS
                    ):
                        continue

                    # FILTER PROJECT TERMS
                    if not any(
                        x.lower() in combined.lower()
                        for x in PROJECT_TERMS
                    ):
                        continue

                    # FILTER STATES
                    detected_state = extract_state(
                        combined
                    )

                    if selected_states:

                        if detected_state == "Unknown":
                            continue

                        if not any(
                            s in detected_state
                            for s in selected_states
                        ):
                            continue

                    # QUALITY SCORE
                    score = quality_score(
                        combined
                    )

                    if score < 5:
                        continue

                    # DATE
                    published = datetime.now().strftime(
                        "%Y-%m-%d"
                    )

                    # DATA
                    final_data.append({

                        "Title":
                        article["title"],

                        "Source":
                        source,

                        "Published Date":
                        published,

                        "Company":
                        extract_company(combined),

                        "State":
                        detected_state,

                        "Project Stage":
                        classify_stage(combined),

                        "MW Capacity":
                        extract_mw(combined),

                        "Investment":
                        extract_cost(combined),

                        "URL":
                        article["url"]
                    })

                except:
                    continue

            time.sleep(1)

    progress.empty()

    return pd.DataFrame(final_data)

# =========================================================
# EXCEL EXPORT
# =========================================================
def build_excel(df):

    wb = Workbook()

    ws = wb.active

    ws.title = "US Data Center Intelligence"

    headers = list(df.columns)

    fill = PatternFill(
        "solid",
        fgColor="1F4E79"
    )

    font = Font(
        bold=True,
        color="FFFFFF"
    )

    border = Border(

        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin")
    )

    for col_num, header in enumerate(headers, 1):

        cell = ws.cell(
            row=1,
            column=col_num,
            value=header
        )

        cell.fill = fill
        cell.font = font
        cell.border = border
        cell.alignment = Alignment(
            horizontal="center"
        )

    for row in df.itertuples(index=False):

        ws.append(row)

    for col in ws.columns:

        max_length = 0

        column = col[0].column_letter

        for cell in col:

            try:

                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))

            except:
                pass

        adjusted = min(max_length + 5, 60)

        ws.column_dimensions[column].width = adjusted

    excel_file = io.BytesIO()

    wb.save(excel_file)

    excel_file.seek(0)

    return excel_file

# =========================================================
# MAIN UI
# =========================================================
def main():

    st.title(
        "🏗️ US Data Center Intelligence Tracker"
    )

    st.markdown("""
Tracks:
- Data center construction
- AI infrastructure campuses
- Land acquisitions
- Planning approvals
- Utility infrastructure
- Hyperscale expansion
- Investment announcements
""")

    # =====================================================
    # SIDEBAR
    # =====================================================
    st.sidebar.header("Filters")

    time_option = st.sidebar.radio(

        "Select Time Range",

        [
            "Past 10 Days",
            "Past 30 Days",
            "Latest"
        ]
    )

    if time_option == "Past 10 Days":

        start_date = (
            datetime.now()
            - timedelta(days=10)
        )

    elif time_option == "Past 30 Days":

        start_date = (
            datetime.now()
            - timedelta(days=30)
        )

    else:

        start_date = (
            datetime.now()
            - timedelta(days=3)
        )

    selected_states = st.sidebar.multiselect(

        "Select States",

        US_STATES
    )

    max_pages = st.sidebar.slider(

        "Pages Per Source",

        1,
        20,
        5
    )

    # =====================================================
    # SCRAPE BUTTON
    # =====================================================
    if st.button(
        "🚀 Run Intelligence Scan"
    ):

        with st.spinner(
            "Scanning infrastructure developments..."
        ):

            df = scrape_sources(

                max_pages=
                    max_pages,

                selected_states=
                    selected_states,

                start_date=
                    start_date
            )

            if df.empty:

                st.warning(
                    "No matching articles found"
                )

                return

            # DEDUP
            df.drop_duplicates(

                subset="Title",

                inplace=True
            )

            # SORT
            df = df.sort_values(

                "Published Date",

                ascending=False
            )

            # =================================================
            # METRICS
            # =================================================
            c1, c2, c3, c4 = st.columns(4)

            c1.metric(
                "Articles",
                len(df)
            )

            c2.metric(
                "Companies",
                df["Company"].nunique()
            )

            c3.metric(
                "States",
                df["State"].nunique()
            )

            c4.metric(
                "Sources",
                df["Source"].nunique()
            )

            # =================================================
            # CHARTS
            # =================================================
            st.subheader(
                "📊 Intelligence Analytics"
            )

            state_chart = px.histogram(

                df,

                x="State",

                title="Projects by State"
            )

            st.plotly_chart(
                state_chart,
                use_container_width=True
            )

            stage_chart = px.histogram(

                df,

                x="Project Stage",

                title="Projects by Stage"
            )

            st.plotly_chart(
                stage_chart,
                use_container_width=True
            )

            # =================================================
            # TABLE
            # =================================================
            st.subheader(
                "📄 Intelligence Results"
            )

            st.dataframe(

                df,

                use_container_width=True,

                height=700
            )

            # =================================================
            # EXCEL DOWNLOAD
            # =================================================
            excel = build_excel(df)

            st.download_button(

                "📥 Download Intelligence Report",

                data=excel,

                file_name=(
                    "us_data_center_intelligence.xlsx"
                ),

                mime=(
                    "application/vnd.openxmlformats-"
                    "officedocument.spreadsheetml.sheet"
                )
            )

# =========================================================
# RUN
# =========================================================
if __name__ == "__main__":
    main()
