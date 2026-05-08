import streamlit as st
import re
import io
import time
from datetime import datetime, timedelta

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# =========================================================
# OPTIONAL CLOUDSCRAPER (Cloudflare bypass)
# =========================================================
try:
    import cloudscraper
    _SCRAPER = cloudscraper.create_scraper(
        browser={
            "browser": "chrome",
            "platform": "windows",
            "mobile": False
        }
    )
    _USE_CLOUDSCRAPER = True

except ImportError:
    import requests
    _SCRAPER = requests.Session()
    _USE_CLOUDSCRAPER = False

from bs4 import BeautifulSoup

# =========================================================
# CONFIG
# =========================================================
BASE_URL = "https://www.datacenterdynamics.com"

SEARCH_TERMS = [
    "the-data-center-construction-channel",
    "north-america"
]

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/125.0 Safari/537.36"
    )
}

# =========================================================
# ALL US STATES
# =========================================================
US_STATES = [
    "All USA",
    "Alabama","Alaska","Arizona","Arkansas","California","Colorado",
    "Connecticut","Delaware","Florida","Georgia","Hawaii","Idaho",
    "Illinois","Indiana","Iowa","Kansas","Kentucky","Louisiana",
    "Maine","Maryland","Massachusetts","Michigan","Minnesota",
    "Mississippi","Missouri","Montana","Nebraska","Nevada",
    "New Hampshire","New Jersey","New Mexico","New York",
    "North Carolina","North Dakota","Ohio","Oklahoma","Oregon",
    "Pennsylvania","Rhode Island","South Carolina","South Dakota",
    "Tennessee","Texas","Utah","Vermont","Virginia","Washington",
    "West Virginia","Wisconsin","Wyoming"
]

# =========================================================
# DISCLOSURE / CONSTRUCTION KEYWORDS
# =========================================================
PROJECT_KEYWORDS = [
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
    "site plan",
    "disclosed"
]

# =========================================================
# DATE PARSER
# =========================================================
MONTHS = {
    "jan": 1, "feb": 2, "mar": 3,
    "apr": 4, "may": 5, "jun": 6,
    "jul": 7, "aug": 8, "sep": 9,
    "oct": 10, "nov": 11, "dec": 12
}

def parse_date(raw):
    raw = raw.strip()

    # Example: 2026-05-08
    try:
        return datetime.strptime(raw, "%Y-%m-%d")
    except:
        pass

    # Example: 8 May 2026
    m = re.search(r"(\d{1,2})\s+(\w+)\s+(\d{4})", raw, re.I)

    if m:
        day = int(m.group(1))
        mon = m.group(2).lower()[:3]
        year = int(m.group(3))

        if mon in MONTHS:
            return datetime(year, MONTHS[mon], day)

    return None

# =========================================================
# FETCH HTML
# =========================================================
def fetch_page(url):

    try:

        if _USE_CLOUDSCRAPER:
            response = _SCRAPER.get(url, timeout=25)

        else:
            response = _SCRAPER.get(
                url,
                headers=HEADERS,
                timeout=25
            )

        response.raise_for_status()

        return BeautifulSoup(response.text, "html.parser")

    except Exception as e:
        st.warning(f"Error fetching: {url}")
        st.warning(str(e))
        return None

# =========================================================
# ARTICLE PARSER
# =========================================================
def extract_articles(soup):

    articles = []
    seen = set()

    links = soup.find_all(
        "a",
        href=re.compile(r"^/en/news/")
    )

    for link in links:

        href = link.get("href")

        if not href:
            continue

        if href in seen:
            continue

        seen.add(href)

        full_url = BASE_URL + href

        headline = link.get_text(strip=True)

        if len(headline) < 20:
            continue

        headline_lower = headline.lower()

        # Must contain disclosure / construction signals
        if not any(k in headline_lower for k in PROJECT_KEYWORDS):
            continue

        # Extract nearby text for date
        parent_text = link.parent.get_text(" ", strip=True)

        date_match = re.search(
            r"(\d{1,2}\s+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{4})",
            parent_text,
            re.I
        )

        if date_match:
            date_obj = parse_date(date_match.group(1))
        else:
            date_obj = None

        articles.append({
            "Headline": headline,
            "Date": date_obj.strftime("%Y-%m-%d") if date_obj else "Unknown",
            "_date": date_obj,
            "URL": full_url
        })

    return articles

# =========================================================
# USA FILTER
# =========================================================
def article_matches_state(headline, state):

    if state == "All USA":
        return True

    return state.lower() in headline.lower()

# =========================================================
# SCRAPER
# =========================================================
def scrape_articles(
    cutoff_date,
    max_pages,
    selected_state,
    progress_bar
):

    collected = []

    for page in range(1, max_pages + 1):

        progress_bar.progress(
            page / max_pages,
            text=f"Scraping page {page}/{max_pages}"
        )

        url = (
            f"{BASE_URL}/en/news/"
            f"?term=the-data-center-construction-channel"
            f"&term=north-america"
            f"&page={page}"
        )

        soup = fetch_page(url)

        if not soup:
            continue

        page_articles = extract_articles(soup)

        if not page_articles:
            continue

        for article in page_articles:

            # Filter by state
            if not article_matches_state(
                article["Headline"],
                selected_state
            ):
                continue

            # Filter by date
            if article["_date"]:

                if article["_date"] < cutoff_date:
                    continue

            collected.append(article)

        time.sleep(1)

    return collected

# =========================================================
# EXCEL EXPORT
# =========================================================
def build_excel(df):

    wb = Workbook()
    ws = wb.active

    ws.title = "US Data Center Disclosures"

    headers = [
        "#",
        "Headline",
        "Published Date",
        "URL"
    ]

    widths = [5, 90, 18, 70]

    header_fill = PatternFill(
        "solid",
        fgColor="1F4E79"
    )

    header_font = Font(
        bold=True,
        color="FFFFFF",
        size=11
    )

    thin = Side(style="thin", color="CCCCCC")

    border = Border(
        left=thin,
        right=thin,
        top=thin,
        bottom=thin
    )

    # HEADER
    for i, (h, w) in enumerate(zip(headers, widths), 1):

        cell = ws.cell(
            row=1,
            column=i,
            value=h
        )

        cell.fill = header_fill
        cell.font = header_font
        cell.border = border

        cell.alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

        ws.column_dimensions[
            get_column_letter(i)
        ].width = w

    # DATA
    for row_idx, row in enumerate(
        df.itertuples(index=False),
        start=2
    ):

        values = [
            row_idx - 1,
            row.Headline,
            row.Date,
            row.URL
        ]

        for col_idx, val in enumerate(values, 1):

            c = ws.cell(
                row=row_idx,
                column=col_idx,
                value=val
            )

            c.border = border

            if col_idx == 4:
                c.hyperlink = val
                c.style = "Hyperlink"

    ws.freeze_panes = "B2"

    buffer = io.BytesIO()

    wb.save(buffer)

    buffer.seek(0)

    return buffer.read()

# =========================================================
# APP
# =========================================================
def main():

    st.set_page_config(
        page_title="US Data Center Intelligence",
        page_icon="🏗️",
        layout="wide"
    )

    st.title("🏗️ US Data Center Disclosure Tracker")

    st.markdown("""
Tracks:
- Data Center Construction
- Investments
- Campus Developments
- Groundbreaking
- Permits
- Approvals
- Expansion Announcements
- Land Acquisition
- Site Selection
- Planning Activity
""")

    if not _USE_CLOUDSCRAPER:

        st.warning("""
Install cloudscraper for better scraping:

pip install cloudscraper
""")

    # =====================================================
    # SIDEBAR
    # =====================================================
    with st.sidebar:

        st.header("⚙️ Filters")

        # STATE FILTER
        selected_state = st.selectbox(
            "Select USA State",
            US_STATES
        )

        # DATE FILTER
        date_option = st.radio(
            "Date Filter",
            [
                "Latest Available",
                "Past 10 Days",
                "Past 30 Days",
                "Custom Date Range"
            ]
        )

        today = datetime.today()

        if date_option == "Latest Available":
            cutoff_date = datetime(2020, 1, 1)

        elif date_option == "Past 10 Days":
            cutoff_date = today - timedelta(days=10)

        elif date_option == "Past 30 Days":
            cutoff_date = today - timedelta(days=30)

        else:

            custom_start = st.date_input(
                "Start Date",
                today - timedelta(days=30)
            )

            custom_end = st.date_input(
                "End Date",
                today
            )

            cutoff_date = datetime.combine(
                custom_start,
                datetime.min.time()
            )

        max_pages = st.slider(
            "Pages to Scrape",
            1,
            50,
            15
        )

        scrape_btn = st.button(
            "🚀 Start Scraping",
            use_container_width=True
        )

    # =====================================================
    # MAIN
    # =====================================================
    if not scrape_btn:

        st.info("Select filters and click Start Scraping")

        return

    progress = st.progress(
        0,
        text="Starting scraper..."
    )

    articles = scrape_articles(
        cutoff_date,
        max_pages,
        selected_state,
        progress
    )

    progress.empty()

    if not articles:

        st.error("No articles found")

        return

    df = pd.DataFrame(articles)

    # REMOVE DUPLICATES
    df.drop_duplicates(
        subset=["Headline"],
        inplace=True
    )

    # DATE FILTER
    if date_option == "Custom Date Range":

        end_dt = datetime.combine(
            custom_end,
            datetime.max.time()
        )

        df["_date"] = pd.to_datetime(
            df["_date"],
            errors="coerce"
        )

        df = df[
            (df["_date"] >= cutoff_date) &
            (df["_date"] <= end_dt)
        ]

    # SORT
    df = df.sort_values(
        "_date",
        ascending=False
    )

    # KEEP COLUMNS
    df = df[
        [
            "Headline",
            "Date",
            "URL"
        ]
    ]

    # =====================================================
    # METRICS
    # =====================================================
    c1, c2, c3 = st.columns(3)

    c1.metric(
        "Articles",
        len(df)
    )

    c2.metric(
        "Pages Scraped",
        max_pages
    )

    c3.metric(
        "State Filter",
        selected_state
    )

    # =====================================================
    # TABLE
    # =====================================================
    st.dataframe(
        df,
        use_container_width=True,
        height=700,
        column_config={
            "URL": st.column_config.LinkColumn(
                "Article Link",
                display_text="🔗 Open Article"
            )
        }
    )

    # =====================================================
    # DOWNLOAD
    # =====================================================
    excel_data = build_excel(df)

    timestamp = datetime.now().strftime(
        "%Y%m%d_%H%M"
    )

    st.download_button(
        label="📥 Download Excel",
        data=excel_data,
        file_name=f"US_Data_Center_Disclosures_{timestamp}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

# =========================================================
# RUN
# =========================================================
if __name__ == "__main__":
    main()
