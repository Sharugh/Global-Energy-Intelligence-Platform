import streamlit as st
import requests
from bs4 import BeautifulSoup
import pandas as pd
from datetime import datetime, timedelta
import re
import io
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ─── US States for filtering ───────────────────────────────────────────────
US_STATES = [
    "Alabama", "Alaska", "Arizona", "Arkansas", "California", "Colorado",
    "Connecticut", "Delaware", "Florida", "Georgia", "Hawaii", "Idaho",
    "Illinois", "Indiana", "Iowa", "Kansas", "Kentucky", "Louisiana",
    "Maine", "Maryland", "Massachusetts", "Michigan", "Minnesota",
    "Mississippi", "Missouri", "Montana", "Nebraska", "Nevada",
    "New Hampshire", "New Jersey", "New Mexico", "New York",
    "North Carolina", "North Dakota", "Ohio", "Oklahoma", "Oregon",
    "Pennsylvania", "Rhode Island", "South Carolina", "South Dakota",
    "Tennessee", "Texas", "Utah", "Vermont", "Virginia", "Washington",
    "West Virginia", "Wisconsin", "Wyoming",
    # Common abbreviations
    "AL", "AK", "AZ", "AR", "CA", "CO", "CT", "DE", "FL", "GA",
    "HI", "ID", "IL", "IN", "IA", "KS", "KY", "LA", "ME", "MD",
    "MA", "MI", "MN", "MS", "MO", "MT", "NE", "NV", "NH", "NJ",
    "NM", "NY", "NC", "ND", "OH", "OK", "OR", "PA", "RI", "SC",
    "SD", "TN", "TX", "UT", "VT", "VA", "WA", "WV", "WI", "WY",
    # US cities often seen in data center news
    "Silicon Valley", "Northern Virginia", "NoVA", "Loudoun County",
    "Dallas", "Chicago", "Phoenix", "Atlanta", "Seattle", "Denver",
    "San Jose", "San Francisco", "Los Angeles", "Houston", "Miami",
    "New York City", "NYC", "Boston", "Portland", "Las Vegas",
    "Reno", "Quincy", "Ashburn", "Sterling", "Manassas",
    "United States", "U.S.", "US "
]

US_PATTERN = re.compile(
    r'\b(' + '|'.join(re.escape(s) for s in US_STATES) + r')\b',
    re.IGNORECASE
)

BASE_URL = "https://www.datacenterdynamics.com"
CHANNEL_URL = f"{BASE_URL}/en/news/?term=the-data-center-construction-channel"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "en-US,en;q=0.9",
}


# ─── Scraping helpers ───────────────────────────────────────────────────────

def parse_date(date_str: str) -> datetime | None:
    """Try multiple date formats and return a datetime or None."""
    date_str = date_str.strip()
    for fmt in ("%B %d, %Y", "%d %B %Y", "%Y-%m-%d", "%b %d, %Y"):
        try:
            return datetime.strptime(date_str, fmt)
        except ValueError:
            continue
    return None


def fetch_page(url: str) -> BeautifulSoup | None:
    try:
        r = requests.get(url, headers=HEADERS, timeout=15)
        r.raise_for_status()
        return BeautifulSoup(r.text, "html.parser")
    except Exception as e:
        st.warning(f"Failed to fetch {url}: {e}")
        return None


def scrape_articles(cutoff: datetime, max_pages: int = 10) -> list[dict]:
    """Scrape DCD construction channel up to cutoff date."""
    articles = []
    page = 1

    while page <= max_pages:
        url = CHANNEL_URL if page == 1 else f"{CHANNEL_URL}&page={page}"
        soup = fetch_page(url)
        if not soup:
            break

        cards = soup.select("article, .article-card, [data-component='article-card']")

        # Fallback: broader selector
        if not cards:
            cards = soup.select("a[href*='/en/news/']")

        if not cards:
            break

        found_any = False
        stop_early = False

        for card in cards:
            # ── headline ──
            headline_tag = card.select_one("h2, h3, h4, .headline, .title")
            if not headline_tag:
                if card.name == "a":
                    headline_tag = card
                else:
                    continue
            headline = headline_tag.get_text(strip=True)
            if not headline:
                continue

            # ── url ──
            link_tag = card.select_one("a") or (card if card.name == "a" else None)
            if not link_tag:
                continue
            href = link_tag.get("href", "")
            if not href.startswith("http"):
                href = BASE_URL + href

            # ── date ──
            date_tag = card.select_one("time, .date, [class*='date'], [class*='time']")
            article_date = None
            if date_tag:
                raw = date_tag.get("datetime") or date_tag.get_text(strip=True)
                article_date = parse_date(raw)

            if article_date and article_date < cutoff:
                stop_early = True
                break

            found_any = True
            articles.append({
                "Headline": headline,
                "Date": article_date.strftime("%Y-%m-%d") if article_date else "Unknown",
                "URL": href,
                "_date_obj": article_date,
            })

        if stop_early or not found_any:
            break
        page += 1

    return articles


def is_us_related(headline: str) -> bool:
    return bool(US_PATTERN.search(headline))


def filter_articles(articles: list[dict], days: int | None) -> list[dict]:
    """Filter by date window and US-only headlines."""
    now = datetime.now()
    result = []
    for a in articles:
        d = a.get("_date_obj")
        if days is not None and d:
            if (now - d).days > days:
                continue
        if not is_us_related(a["Headline"]):
            continue
        result.append(a)
    return result


# ─── Excel export ────────────────────────────────────────────────────────────

def build_excel(df: pd.DataFrame) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "DCD US Construction News"

    # Header style
    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(bold=True, color="FFFFFF", name="Arial", size=11)
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin = Side(border_style="thin", color="AAAAAA")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    cols = ["Headline", "Date", "URL"]
    col_widths = [70, 15, 60]

    for ci, (col, width) in enumerate(zip(cols, col_widths), start=1):
        cell = ws.cell(row=1, column=ci, value=col)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        cell.border = border
        ws.column_dimensions[get_column_letter(ci)].width = width

    ws.row_dimensions[1].height = 28

    # Alternating row colours
    fill_even = PatternFill("solid", fgColor="DCE6F1")
    fill_odd  = PatternFill("solid", fgColor="FFFFFF")
    link_font = Font(name="Arial", size=10, color="0563C1", underline="single")
    normal_font = Font(name="Arial", size=10)

    for ri, row in enumerate(df.itertuples(index=False), start=2):
        fill = fill_even if ri % 2 == 0 else fill_odd
        for ci, col in enumerate(cols, start=1):
            val = getattr(row, col)
            cell = ws.cell(row=ri, column=ci, value=val)
            cell.fill = fill
            cell.border = border
            cell.alignment = Alignment(vertical="center", wrap_text=(ci == 1))
            if ci == 3:  # URL column — make it a hyperlink
                cell.hyperlink = val
                cell.font = link_font
            else:
                cell.font = normal_font

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:C{len(df)+1}"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


# ─── Streamlit UI ────────────────────────────────────────────────────────────

def main():
    st.set_page_config(
        page_title="DCD US Construction News Scraper",
        page_icon="🏗️",
        layout="wide",
    )

    st.title("🏗️ Data Center Dynamics — US Construction News")
    st.caption(
        "Scrapes the [DCD Construction Channel]("
        "https://www.datacenterdynamics.com/en/news/?term=the-data-center-construction-channel"
        ") and filters for **American state** mentions."
    )

    # ── Sidebar controls ────────────────────────────────────────────────────
    with st.sidebar:
        st.header("⚙️ Filters")
        time_option = st.radio(
            "Date range",
            ["Latest (no limit)", "Past 30 days", "Past 10 days"],
            index=0,
        )
        days_map = {
            "Latest (no limit)": None,
            "Past 30 days": 30,
            "Past 10 days": 10,
        }
        selected_days = days_map[time_option]

        max_pages = st.slider("Max pages to scrape", 1, 20, 5)

        scrape_btn = st.button("🔍 Scrape Now", use_container_width=True, type="primary")

    # ── Main area ────────────────────────────────────────────────────────────
    if scrape_btn:
        cutoff = (
            datetime.min
            if selected_days is None
            else datetime.now() - timedelta(days=selected_days)
        )

        with st.spinner("Fetching articles from datacenterdynamics.com …"):
            raw = scrape_articles(cutoff, max_pages=max_pages)

        with st.spinner("Filtering for US-related headlines …"):
            filtered = filter_articles(raw, days=selected_days)

        st.success(
            f"Found **{len(raw)}** total articles → **{len(filtered)}** US-related"
        )

        if not filtered:
            st.info(
                "No US-related articles found for the selected range. "
                "Try increasing the date range or page limit."
            )
            return

        # Clean display DF
        df = pd.DataFrame(filtered)[["Headline", "Date", "URL"]]
        df = df.sort_values("Date", ascending=False).reset_index(drop=True)

        st.dataframe(
            df,
            use_container_width=True,
            height=500,
            column_config={
                "URL": st.column_config.LinkColumn("URL"),
            },
        )

        # Excel download
        excel_bytes = build_excel(df)
        ts = datetime.now().strftime("%Y%m%d_%H%M")
        label = time_option.replace(" ", "_").replace("(", "").replace(")", "")
        filename = f"DCD_US_Construction_{label}_{ts}.xlsx"

        st.download_button(
            label="📥 Download Excel",
            data=excel_bytes,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

        with st.expander("📊 Quick stats"):
            col1, col2, col3 = st.columns(3)
            col1.metric("Total articles", len(raw))
            col2.metric("US-related", len(filtered))
            col3.metric("Pages scraped", min(max_pages, (len(raw) // 10) + 1))

    else:
        st.info("👈 Configure filters in the sidebar and click **Scrape Now**.")


if __name__ == "__main__":
    main()
