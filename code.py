import streamlit as st
import re
import io
import time
from datetime import datetime, timedelta

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── optional cloudscraper (bypasses Cloudflare) ────────────────────────────
try:
    import cloudscraper
    _SCRAPER = cloudscraper.create_scraper()
    _USE_CLOUDSCRAPER = True
except ImportError:
    import requests
    _SCRAPER = requests.Session()
    _USE_CLOUDSCRAPER = False

from bs4 import BeautifulSoup

# ─── US geography keywords ─────────────────────────────────────────────────
_US_WORDS = [
    "Alabama","Alaska","Arizona","Arkansas","California","Colorado","Connecticut",
    "Delaware","Florida","Georgia","Hawaii","Idaho","Illinois","Indiana","Iowa",
    "Kansas","Kentucky","Louisiana","Maine","Maryland","Massachusetts","Michigan",
    "Minnesota","Mississippi","Missouri","Montana","Nebraska","Nevada",
    "New Hampshire","New Jersey","New Mexico","New York","North Carolina",
    "North Dakota","Ohio","Oklahoma","Oregon","Pennsylvania","Rhode Island",
    "South Carolina","South Dakota","Tennessee","Texas","Utah","Vermont",
    "Virginia","Washington","West Virginia","Wisconsin","Wyoming",
    "United States",
    # Data center hubs
    "Silicon Valley","Northern Virginia","NoVA","Loudoun","Ashburn","Sterling",
    "Manassas","Dallas","Chicago","Phoenix","Atlanta","Seattle","Denver",
    "San Jose","San Francisco","Los Angeles","Houston","Miami","Boston",
    "Portland","Las Vegas","Reno","Quincy","Salt Lake",
    "New York City","NYC","San Antonio","Austin","Columbus","Kansas City",
    "Nashville","Charlotte","Raleigh","Richmond","Sacramento","Boise",
    "Indianapolis","Baltimore","San Diego","Oakland","Pittsburgh","Newark",
    "Memphis","Louisville","Detroit","Minneapolis","Cleveland","Cincinnati",
    "Tampa","Orlando","Jacksonville","Spokane","Tacoma","Bellevue","Mesa",
    "Tucson","Chandler","Scottsdale","Henderson","El Paso","Fort Worth",
    "Garland","Lubbock","Amarillo","Midland","Killeen","Waco",
    "Perry County","Loudoun County","Prince William","Fairfax",
    "Guadalupe County","Tyrone","Sweetwater","Beacon Point",
]
_ABBREVS = [
    "AL","AK","AZ","AR","CA","CO","CT","DE","FL","GA","HI","ID","IL","IN",
    "IA","KS","KY","LA","ME","MD","MA","MI","MN","MS","MO","MT","NE","NV",
    "NH","NJ","NM","NY","NC","ND","OH","OK","OR","PA","RI","SC","SD","TN",
    "TX","UT","VT","VA","WA","WV","WI","WY","DC",
]
_parts = [r"\b" + re.escape(w) + r"\b" for w in _US_WORDS]
_parts += [r"\b" + a + r"\b" for a in _ABBREVS]
_parts += [r"\bU\.S\.\b", r"\bUS\b"]
US_RE = re.compile("|".join(_parts), re.IGNORECASE)

BASE_URL   = "https://www.datacenterdynamics.com"
CHAN_TERM  = "the-data-center-construction-channel"
NA_TERM    = "north-america"
MONTHS = {"jan":1,"feb":2,"mar":3,"apr":4,"may":5,"jun":6,
          "jul":7,"aug":8,"sep":9,"oct":10,"nov":11,"dec":12}

_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
    "Referer": "https://www.datacenterdynamics.com/",
}


def parse_date(raw: str):
    raw = raw.strip()
    try:
        return datetime.strptime(raw, "%Y-%m-%d")
    except ValueError:
        pass
    for pat in (r"(\d{1,2})\s+(\w+)\s+(\d{4})", r"(\w+)\s+(\d{1,2}),?\s+(\d{4})"):
        m = re.match(pat, raw)
        if m:
            g = m.groups()
            try:
                if g[0].isdigit():
                    day, mon, yr = int(g[0]), g[1].lower()[:3], int(g[2])
                else:
                    mon, day, yr = g[0].lower()[:3], int(g[1]), int(g[2])
                if mon in MONTHS:
                    return datetime(yr, MONTHS[mon], day)
            except Exception:
                pass
    return None


def fetch(url: str):
    """Fetch with cloudscraper (if available) or requests, retrying once."""
    for attempt in range(2):
        try:
            if _USE_CLOUDSCRAPER:
                resp = _SCRAPER.get(url, timeout=20)
            else:
                resp = _SCRAPER.get(url, headers=_HEADERS, timeout=20)
            resp.raise_for_status()
            return BeautifulSoup(resp.text, "html.parser")
        except Exception as e:
            if attempt == 0:
                time.sleep(2)
            else:
                st.warning(f"⚠️ Could not fetch page: {e}")
    return None


def parse_articles(soup: BeautifulSoup) -> list:
    articles = []
    seen = set()
    for a in soup.find_all("a", href=re.compile(r"^/en/news/[^?#]+/$")):
        href = a["href"]
        if href in seen:
            continue
        seen.add(href)
        h_tag   = a.find(["h1", "h2", "h3", "h4"])
        headline = h_tag.get_text(strip=True) if h_tag else a.get_text(strip=True)
        if not headline or len(headline) < 10:
            continue
        date_obj = None
        node = a.parent
        for _ in range(8):
            if node is None:
                break
            text = node.get_text(" ", strip=True)
            m = re.search(
                r"\b(\d{1,2})\s+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+(\d{4})\b",
                text, re.I
            )
            if m:
                date_obj = parse_date(m.group(0))
                break
            node = node.parent
        articles.append({
            "Headline": headline,
            "Date":     date_obj.strftime("%Y-%m-%d") if date_obj else "Unknown",
            "URL":      BASE_URL + href,
            "_date_obj": date_obj,
        })
    return articles


def scrape(cutoff, max_pages: int, use_na: bool, pbar) -> list:
    all_articles = []
    # warm-up: visit homepage to get session cookies
    fetch(BASE_URL + "/en/")
    for page in range(1, max_pages + 1):
        params = f"?term={CHAN_TERM}"
        if use_na:
            params += f"&term={NA_TERM}"
        if page > 1:
            params += f"&page={page}"
        url = f"{BASE_URL}/en/news/{params}"
        pbar.progress(page / max_pages, text=f"Fetching page {page}/{max_pages}…")
        soup = fetch(url)
        if not soup:
            break
        page_arts = parse_articles(soup)
        if not page_arts:
            break
        stop = False
        for art in page_arts:
            d = art["_date_obj"]
            if d and d < cutoff:
                stop = True
                break
            all_articles.append(art)
        if stop:
            break
        time.sleep(0.6)
    return all_articles


def is_us(text: str) -> bool:
    return bool(US_RE.search(text))


def build_excel(df: pd.DataFrame) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "DCD US Construction News"
    thin = Side(border_style="thin", color="CCCCCC")
    brd  = Border(left=thin, right=thin, top=thin, bottom=thin)
    cols, widths = ["#", "Headline", "Date", "URL"], [5, 72, 14, 65]
    h_fill = PatternFill("solid", fgColor="1F4E79")
    h_font = Font(bold=True, color="FFFFFF", name="Calibri", size=11)
    for ci, (col, w) in enumerate(zip(cols, widths), 1):
        c = ws.cell(row=1, column=ci, value=col)
        c.font, c.fill = h_font, h_fill
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = brd
        ws.column_dimensions[get_column_letter(ci)].width = w
    ws.row_dimensions[1].height = 26
    ev = PatternFill("solid", fgColor="DCE6F1")
    od = PatternFill("solid", fgColor="FFFFFF")
    nf = Font(name="Calibri", size=10)
    lf = Font(name="Calibri", size=10, color="0563C1", underline="single")
    for ri, row in enumerate(df.itertuples(index=False), start=2):
        fill = ev if ri % 2 == 0 else od
        for ci, val in enumerate([ri-1, row.Headline, row.Date, row.URL], 1):
            c = ws.cell(row=ri, column=ci, value=val)
            c.fill, c.border = fill, brd
            if ci == 4:
                c.hyperlink = val; c.font = lf
                c.alignment = Alignment(horizontal="center", vertical="center")
            elif ci == 2:
                c.font = nf
                c.alignment = Alignment(vertical="center", wrap_text=True)
            else:
                c.font = nf
                c.alignment = Alignment(horizontal="center", vertical="center")
    ws.freeze_panes = "B2"
    ws.auto_filter.ref = f"A1:D{len(df)+1}"
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


# ─── App ───────────────────────────────────────────────────────────────────
def main():
    st.set_page_config(page_title="DCD US Construction News", page_icon="🏗️", layout="wide")
    st.title("🏗️ Data Center Dynamics — US Construction News Scraper")
    st.caption(
        "Scrapes the [DCD Construction Channel]"
        "(https://www.datacenterdynamics.com/en/news/?term=the-data-center-construction-channel)"
        " and filters for **US-based** articles only."
    )
    if not _USE_CLOUDSCRAPER:
        st.warning(
            "💡 **Tip:** Install `cloudscraper` for better Cloudflare bypass: "
            "`pip install cloudscraper`",
            icon="⚠️",
        )

    with st.sidebar:
        st.header("⚙️ Settings")
        time_opt = st.radio(
            "Date range",
            ["Latest (all available)", "Past 30 days", "Past 10 days"],
            index=0,
        )
        days_map = {"Latest (all available)": None, "Past 30 days": 30, "Past 10 days": 10}
        sel_days = days_map[time_opt]
        max_pages = st.slider("Pages to scrape", 1, 30, 5,
                               help="~28 articles per page. 5 pages ≈ 140 articles.")
        use_na = st.checkbox(
            "Pre-filter: North America region",
            value=True,
            help="Adds &term=north-america to the DCD URL for better relevance.",
        )
        go = st.button("🔍 Scrape Now", use_container_width=True, type="primary")

    if not go:
        st.info("👈 Set filters in the sidebar and click **Scrape Now**.")
        st.markdown("""
**How it works**
1. Visits the DCD Construction Channel (server-rendered HTML)
2. Parses article headlines, dates, and URLs from anchor tags
3. Optionally pre-limits results to the North America tag
4. Post-filters using 50 US states + abbreviations + ~80 major data center cities/counties
5. Exports a formatted `.xlsx` with auto-filter, frozen headers, and clickable URL links

**Requirements** — add to `requirements.txt`:
```
streamlit>=1.35.0
requests>=2.31.0
cloudscraper>=1.2.71
beautifulsoup4>=4.12.0
lxml>=5.2.0
pandas>=2.2.0
openpyxl>=3.1.2
```
        """)
        return

    cutoff = datetime.min if sel_days is None else datetime.now() - timedelta(days=sel_days)
    pbar = st.progress(0, text="Starting…")
    raw = scrape(cutoff, max_pages, use_na, pbar)
    pbar.empty()

    if not raw:
        st.error(
            "No articles were fetched.\n\n"
            "**Most likely cause:** Cloudflare is blocking automated requests from this host.\n\n"
            "**Fix:** Make sure `cloudscraper` is installed (`pip install cloudscraper`) "
            "and re-deploy. If the issue persists, try running locally."
        )
        return

    us_arts = [a for a in raw if is_us(a["Headline"])]
    c1, c2, c3 = st.columns(3)
    c1.metric("Pages scraped", max_pages)
    c2.metric("Total articles", len(raw))
    c3.metric("US-related articles", len(us_arts))

    if not us_arts:
        st.warning("No US-related articles matched. Showing raw results for inspection:")
        with st.expander("Raw articles"):
            for a in raw[:15]:
                st.write(f"**{a['Headline']}** — {a['Date']}")
        return

    df = (
        pd.DataFrame(us_arts)[["Headline", "Date", "URL"]]
        .sort_values("Date", ascending=False)
        .reset_index(drop=True)
    )

    st.dataframe(
        df, use_container_width=True, height=520,
        column_config={"URL": st.column_config.LinkColumn("URL", display_text="🔗 Open")},
    )

    ts    = datetime.now().strftime("%Y%m%d_%H%M")
    label = re.sub(r"[^a-zA-Z0-9]", "_", time_opt)
    st.download_button(
        label="📥 Download Excel (.xlsx)",
        data=build_excel(df),
        file_name=f"DCD_US_Construction_{label}_{ts}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        type="primary",
    )


if __name__ == "__main__":
    main()
