import streamlit as st
import re
import io
import time
from datetime import datetime, timedelta
from collections import Counter

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── optional cloudscraper ──────────────────────────────────────────────────
try:
    import cloudscraper
    _SCRAPER = cloudscraper.create_scraper()
    _USE_CLOUDSCRAPER = True
except ImportError:
    import requests
    _SCRAPER = requests.Session()
    _USE_CLOUDSCRAPER = False

from bs4 import BeautifulSoup

# ─── CSS / Theme injection ─────────────────────────────────────────────────
CUSTOM_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Syne:wght@400;600;700;800&family=DM+Mono:wght@400;500&family=Inter:wght@300;400;500&display=swap');

/* ── Global reset ── */
html, body, [class*="css"] {
    font-family: 'Inter', sans-serif;
}
.stApp {
    background: #080c14;
    color: #e8edf5;
}

/* ── Sidebar ── */
[data-testid="stSidebar"] {
    background: #0d1220 !important;
    border-right: 1px solid #1e2a40;
}
[data-testid="stSidebar"] * { color: #c8d4e8 !important; }
[data-testid="stSidebar"] .stButton button {
    background: linear-gradient(135deg, #0057ff, #00c6ff) !important;
    color: #fff !important;
    border: none !important;
    border-radius: 8px !important;
    font-family: 'Syne', sans-serif !important;
    font-weight: 700 !important;
    font-size: 0.95rem !important;
    letter-spacing: 0.04em !important;
    padding: 0.65rem 1rem !important;
    transition: opacity .2s;
}
[data-testid="stSidebar"] .stButton button:hover { opacity: .85; }
[data-testid="stSidebar"] hr { border-color: #1e2a40 !important; }

/* ── Top header banner ── */
.dcd-banner {
    background: linear-gradient(135deg, #0a1628 0%, #0d2347 50%, #091830 100%);
    border: 1px solid #1a3260;
    border-radius: 16px;
    padding: 2rem 2.5rem;
    margin-bottom: 1.8rem;
    position: relative;
    overflow: hidden;
}
.dcd-banner::before {
    content: '';
    position: absolute;
    top: -60px; right: -60px;
    width: 280px; height: 280px;
    background: radial-gradient(circle, rgba(0,87,255,0.18) 0%, transparent 70%);
    border-radius: 50%;
}
.dcd-banner::after {
    content: '';
    position: absolute;
    bottom: -40px; left: 30%;
    width: 200px; height: 200px;
    background: radial-gradient(circle, rgba(0,198,255,0.10) 0%, transparent 70%);
    border-radius: 50%;
}
.banner-eyebrow {
    font-family: 'DM Mono', monospace;
    font-size: 0.72rem;
    letter-spacing: 0.18em;
    color: #00c6ff;
    text-transform: uppercase;
    margin-bottom: 0.5rem;
}
.banner-title {
    font-family: 'Syne', sans-serif;
    font-size: 2rem;
    font-weight: 800;
    color: #ffffff;
    line-height: 1.15;
    margin-bottom: 0.4rem;
}
.banner-title span { color: #00c6ff; }
.banner-sub {
    font-size: 0.88rem;
    color: #7a90b8;
    font-weight: 300;
}
.banner-ts {
    font-family: 'DM Mono', monospace;
    font-size: 0.72rem;
    color: #3a5280;
    margin-top: 0.8rem;
    letter-spacing: 0.05em;
}

/* ── KPI cards ── */
.kpi-row { display: flex; gap: 1rem; margin-bottom: 1.8rem; flex-wrap: wrap; }
.kpi-card {
    flex: 1; min-width: 160px;
    background: #0d1628;
    border: 1px solid #1a2b48;
    border-radius: 12px;
    padding: 1.2rem 1.5rem;
    position: relative;
    overflow: hidden;
    transition: border-color .2s, transform .2s;
}
.kpi-card:hover { border-color: #0057ff; transform: translateY(-2px); }
.kpi-card::before {
    content: '';
    position: absolute;
    top: 0; left: 0; right: 0;
    height: 2px;
}
.kpi-card.blue::before  { background: linear-gradient(90deg, #0057ff, #00c6ff); }
.kpi-card.cyan::before  { background: linear-gradient(90deg, #00c6ff, #00ffe7); }
.kpi-card.green::before { background: linear-gradient(90deg, #00e676, #00c853); }
.kpi-card.gold::before  { background: linear-gradient(90deg, #ffab00, #ff6d00); }
.kpi-label {
    font-family: 'DM Mono', monospace;
    font-size: 0.68rem;
    letter-spacing: 0.14em;
    text-transform: uppercase;
    color: #4a6490;
    margin-bottom: 0.5rem;
}
.kpi-value {
    font-family: 'Syne', sans-serif;
    font-size: 2rem;
    font-weight: 800;
    color: #ffffff;
    line-height: 1;
}
.kpi-delta {
    font-size: 0.75rem;
    color: #4a6490;
    margin-top: 0.3rem;
}

/* ── Section headers ── */
.section-header {
    font-family: 'Syne', sans-serif;
    font-size: 1rem;
    font-weight: 700;
    color: #c8d4e8;
    letter-spacing: 0.06em;
    text-transform: uppercase;
    border-left: 3px solid #0057ff;
    padding-left: 0.75rem;
    margin: 1.8rem 0 1rem 0;
}

/* ── Article cards ── */
.card-grid { display: flex; flex-direction: column; gap: 0.65rem; margin-bottom: 1.5rem; }
.article-card {
    background: #0d1628;
    border: 1px solid #1a2b48;
    border-radius: 10px;
    padding: 1rem 1.25rem;
    display: flex;
    justify-content: space-between;
    align-items: flex-start;
    gap: 1rem;
    transition: border-color .18s, background .18s;
    text-decoration: none;
}
.article-card:hover {
    border-color: #0057ff;
    background: #101e38;
}
.card-headline {
    font-size: 0.9rem;
    font-weight: 500;
    color: #d8e4f8;
    line-height: 1.45;
    flex: 1;
}
.card-headline a {
    color: #d8e4f8 !important;
    text-decoration: none;
}
.card-headline a:hover { color: #00c6ff !important; }
.card-meta {
    display: flex;
    flex-direction: column;
    align-items: flex-end;
    gap: 0.35rem;
    white-space: nowrap;
    flex-shrink: 0;
}
.card-date {
    font-family: 'DM Mono', monospace;
    font-size: 0.72rem;
    color: #3a5280;
}
.card-tag {
    font-family: 'DM Mono', monospace;
    font-size: 0.65rem;
    background: rgba(0,87,255,0.15);
    color: #00c6ff;
    border: 1px solid rgba(0,198,255,0.2);
    border-radius: 4px;
    padding: 2px 7px;
}
.card-link-btn {
    font-size: 0.72rem;
    color: #0057ff !important;
    text-decoration: none;
    font-family: 'DM Mono', monospace;
}
.card-link-btn:hover { color: #00c6ff !important; }

/* ── Insight pills ── */
.insight-row { display: flex; flex-wrap: wrap; gap: 0.5rem; margin-bottom: 1.5rem; }
.insight-pill {
    background: #0d1628;
    border: 1px solid #1a2b48;
    border-radius: 20px;
    padding: 0.4rem 1rem;
    font-size: 0.8rem;
    color: #7a90b8;
    display: flex; align-items: center; gap: 0.4rem;
}
.insight-pill b { color: #ffffff; }
.insight-pill .dot {
    width: 7px; height: 7px;
    border-radius: 50%;
    background: #0057ff;
    flex-shrink: 0;
}

/* ── Search bar ── */
.stTextInput input {
    background: #0d1628 !important;
    border: 1px solid #1a2b48 !important;
    border-radius: 8px !important;
    color: #d8e4f8 !important;
    font-family: 'Inter', sans-serif !important;
}
.stTextInput input:focus {
    border-color: #0057ff !important;
    box-shadow: 0 0 0 2px rgba(0,87,255,0.2) !important;
}

/* ── Multiselect ── */
.stMultiSelect [data-baseweb="select"] {
    background: #0d1628 !important;
    border-color: #1a2b48 !important;
    border-radius: 8px !important;
}

/* ── Selectbox ── */
.stSelectbox [data-baseweb="select"] > div {
    background: #0d1628 !important;
    border-color: #1a2b48 !important;
}

/* ── Download button ── */
.stDownloadButton button {
    background: linear-gradient(135deg, #00390f, #006622) !important;
    color: #00e676 !important;
    border: 1px solid #00c853 !important;
    border-radius: 8px !important;
    font-family: 'Syne', sans-serif !important;
    font-weight: 700 !important;
    letter-spacing: 0.04em !important;
    transition: opacity .2s !important;
}
.stDownloadButton button:hover { opacity: .85 !important; }

/* ── Tabs ── */
.stTabs [data-baseweb="tab-list"] {
    background: #0d1628 !important;
    border-radius: 10px !important;
    padding: 4px !important;
    gap: 4px !important;
    border: 1px solid #1a2b48;
}
.stTabs [data-baseweb="tab"] {
    background: transparent !important;
    color: #4a6490 !important;
    border-radius: 7px !important;
    font-family: 'Syne', sans-serif !important;
    font-weight: 600 !important;
    font-size: 0.82rem !important;
    letter-spacing: 0.04em !important;
    padding: 0.5rem 1.2rem !important;
}
.stTabs [aria-selected="true"] {
    background: #0057ff !important;
    color: #ffffff !important;
}
.stTabs [data-baseweb="tab-panel"] { padding-top: 1.2rem !important; }

/* ── Expander ── */
.streamlit-expanderHeader {
    background: #0d1628 !important;
    border: 1px solid #1a2b48 !important;
    border-radius: 8px !important;
    color: #c8d4e8 !important;
    font-family: 'Syne', sans-serif !important;
}
.streamlit-expanderContent {
    background: #080c14 !important;
    border: 1px solid #1a2b48 !important;
    border-top: none !important;
}

/* ── Divider ── */
hr { border-color: #1a2b48 !important; }

/* ── Hide Streamlit branding ── */
#MainMenu, footer, header { visibility: hidden; }
</style>
"""

# ─── US geography ──────────────────────────────────────────────────────────
US_STATES_FULL = [
    "Alabama","Alaska","Arizona","Arkansas","California","Colorado","Connecticut",
    "Delaware","Florida","Georgia","Hawaii","Idaho","Illinois","Indiana","Iowa",
    "Kansas","Kentucky","Louisiana","Maine","Maryland","Massachusetts","Michigan",
    "Minnesota","Mississippi","Missouri","Montana","Nebraska","Nevada",
    "New Hampshire","New Jersey","New Mexico","New York","North Carolina",
    "North Dakota","Ohio","Oklahoma","Oregon","Pennsylvania","Rhode Island",
    "South Carolina","South Dakota","Tennessee","Texas","Utah","Vermont",
    "Virginia","Washington","West Virginia","Wisconsin","Wyoming",
]
_US_WORDS = US_STATES_FULL + [
    "United States","Silicon Valley","Northern Virginia","NoVA","Loudoun","Ashburn",
    "Sterling","Manassas","Dallas","Chicago","Phoenix","Atlanta","Seattle","Denver",
    "San Jose","San Francisco","Los Angeles","Houston","Miami","Boston","Portland",
    "Las Vegas","Reno","Quincy","Salt Lake","New York City","NYC","San Antonio",
    "Austin","Columbus","Kansas City","Nashville","Charlotte","Raleigh","Richmond",
    "Sacramento","Boise","Indianapolis","Baltimore","San Diego","Oakland","Pittsburgh",
    "Newark","Memphis","Louisville","Detroit","Minneapolis","Cleveland","Cincinnati",
    "Tampa","Orlando","Jacksonville","Spokane","Tacoma","Bellevue","Mesa","Tucson",
    "Chandler","Scottsdale","Henderson","El Paso","Fort Worth","Garland","Lubbock",
    "Amarillo","Midland","Killeen","Waco","Perry County","Loudoun County",
    "Prince William","Fairfax","Guadalupe County","Tyrone","Sweetwater","Beacon Point",
]
_ABBREVS = [
    "AL","AK","AZ","AR","CA","CO","CT","DE","FL","GA","HI","ID","IL","IN",
    "IA","KS","KY","LA","ME","MD","MA","MI","MN","MS","MO","MT","NE","NV",
    "NH","NJ","NM","NY","NC","ND","OH","OK","OR","PA","RI","SC","SD","TN",
    "TX","UT","VT","VA","WA","WV","WI","WY","DC",
]
_parts  = [r"\b" + re.escape(w) + r"\b" for w in _US_WORDS]
_parts += [r"\b" + a + r"\b" for a in _ABBREVS]
_parts += [r"\bU\.S\.\b", r"\bUS\b"]
US_RE = re.compile("|".join(_parts), re.IGNORECASE)

# map state → abbreviation for tagging
STATE_TO_ABBR = {s: _ABBREVS[i] for i, s in enumerate(US_STATES_FULL)}

# Keywords for topic tagging
TOPIC_KEYWORDS = {
    "Hyperscale": ["hyperscale","microsoft","google","amazon","aws","meta","apple","oracle"],
    "Colocation": ["colo","colocation","equinix","digital realty","ironmountain","coresite"],
    "AI / GPU":   ["ai","gpu","nvidia","inference","llm","generative","artificial intelligence"],
    "Power":      ["mw","megawatt","gigawatt","power","energy","grid","nuclear","solar"],
    "Investment": ["invest","fund","reit","billion","million","acquire","acquisition","ipo","lease"],
    "Permits":    ["permit","zoning","moratorium","approved","approval","planning","ordinance"],
    "Construction":["broke ground","groundbreaking","opens","opens","topping","construction","campus","build"],
}

BASE_URL  = "https://www.datacenterdynamics.com"
CHAN_TERM = "the-data-center-construction-channel"
NA_TERM   = "north-america"
MONTHS    = {"jan":1,"feb":2,"mar":3,"apr":4,"may":5,"jun":6,
             "jul":7,"aug":8,"sep":9,"oct":10,"nov":11,"dec":12}

_HEADERS = {
    "User-Agent":(
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept":"text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language":"en-US,en;q=0.9",
    "Referer":"https://www.datacenterdynamics.com/",
}


# ─── Scraping (unchanged) ──────────────────────────────────────────────────
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
        h_tag    = a.find(["h1","h2","h3","h4"])
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
            "Headline":  headline,
            "Date":      date_obj.strftime("%Y-%m-%d") if date_obj else "Unknown",
            "URL":       BASE_URL + href,
            "_date_obj": date_obj,
        })
    return articles

def scrape(cutoff, max_pages: int, use_na: bool, pbar) -> list:
    all_articles = []
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


# ─── Enrichment ───────────────────────────────────────────────────────────
def detect_state(headline: str) -> str:
    for state in US_STATES_FULL:
        if re.search(r"\b" + re.escape(state) + r"\b", headline, re.I):
            return state
    # abbreviations — only if clearly a state context
    abbr_map = dict(zip(_ABBREVS, US_STATES_FULL))
    for abbr in _ABBREVS:
        if re.search(r"\b" + abbr + r"\b", headline):
            return abbr_map.get(abbr, abbr)
    return "Other US"

def detect_topic(headline: str) -> str:
    hl = headline.lower()
    for topic, kws in TOPIC_KEYWORDS.items():
        if any(k in hl for k in kws):
            return topic
    return "General"

def detect_mw(headline: str) -> str:
    m = re.search(r"(\d[\d,]*(?:\.\d+)?)\s*(GW|MW|megawatt|gigawatt)", headline, re.I)
    if m:
        val = m.group(1).replace(",","")
        unit = m.group(2).upper()
        return f"{val} {unit}"
    return ""

def enrich(articles: list) -> list:
    for a in articles:
        hl = a["Headline"]
        a["State"]    = detect_state(hl)
        a["Topic"]    = detect_topic(hl)
        a["Capacity"] = detect_mw(hl)
    return articles


# ─── Excel export ──────────────────────────────────────────────────────────
def build_excel(df: pd.DataFrame) -> bytes:
    wb = Workbook()

    # ── Sheet 1: Articles ──
    ws = wb.active
    ws.title = "US Construction Articles"
    thin = Side(border_style="thin", color="CCCCCC")
    brd  = Border(left=thin, right=thin, top=thin, bottom=thin)
    cols   = ["#", "Headline", "Date", "State", "Topic", "Capacity", "URL"]
    widths = [4,   62,          13,     18,       14,      12,         55]
    h_fill = PatternFill("solid", fgColor="1F4E79")
    h_font = Font(bold=True, color="FFFFFF", name="Calibri", size=11)
    for ci, (col, w) in enumerate(zip(cols, widths), 1):
        c = ws.cell(row=1, column=ci, value=col)
        c.font = h_font; c.fill = h_fill
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = brd
        ws.column_dimensions[get_column_letter(ci)].width = w
    ws.row_dimensions[1].height = 26
    ev = PatternFill("solid", fgColor="DCE6F1")
    od = PatternFill("solid", fgColor="FFFFFF")
    nf = Font(name="Calibri", size=10)
    lf = Font(name="Calibri", size=10, color="0563C1", underline="single")
    topic_colors = {
        "Hyperscale":"4472C4","Colocation":"70AD47","AI / GPU":"FF0000",
        "Power":"FFC000","Investment":"7030A0","Permits":"FF7C00","Construction":"00B0F0",
    }
    for ri, row in enumerate(df.itertuples(index=False), start=2):
        fill = ev if ri % 2 == 0 else od
        vals = [ri-1, row.Headline, row.Date, row.State, row.Topic, row.Capacity, row.URL]
        for ci, val in enumerate(vals, 1):
            c = ws.cell(row=ri, column=ci, value=str(val) if val else "")
            c.fill = fill; c.border = brd
            if ci == 7:
                c.hyperlink = val; c.font = lf
                c.alignment = Alignment(horizontal="center", vertical="center")
            elif ci == 5:   # Topic — colored
                c.font = Font(name="Calibri", size=10, bold=True,
                              color=topic_colors.get(str(val), "000000"))
                c.alignment = Alignment(horizontal="center", vertical="center")
            elif ci == 2:
                c.font = nf
                c.alignment = Alignment(vertical="center", wrap_text=True)
            else:
                c.font = nf
                c.alignment = Alignment(horizontal="center", vertical="center")
    ws.freeze_panes = "B2"
    ws.auto_filter.ref = f"A1:G{len(df)+1}"

    # ── Sheet 2: State Summary ──
    ws2 = wb.create_sheet("State Breakdown")
    state_counts = df["State"].value_counts().reset_index()
    state_counts.columns = ["State", "Article Count"]
    ws2.cell(1, 1, "State").font = h_font;        ws2.cell(1, 1).fill = h_fill
    ws2.cell(1, 2, "Article Count").font = h_font; ws2.cell(1, 2).fill = h_fill
    for ci in [1, 2]:
        ws2.cell(1, ci).alignment = Alignment(horizontal="center", vertical="center")
        ws2.cell(1, ci).border = brd
    ws2.column_dimensions["A"].width = 22
    ws2.column_dimensions["B"].width = 16
    for ri, (_, row) in enumerate(state_counts.iterrows(), start=2):
        for ci, val in enumerate([row["State"], row["Article Count"]], 1):
            c = ws2.cell(ri, ci, val)
            c.font = nf
            c.fill = ev if ri % 2 == 0 else od
            c.border = brd
            c.alignment = Alignment(horizontal="center" if ci == 2 else "left", vertical="center")

    # ── Sheet 3: Topic Summary ──
    ws3 = wb.create_sheet("Topic Breakdown")
    topic_counts = df["Topic"].value_counts().reset_index()
    topic_counts.columns = ["Topic", "Article Count"]
    ws3.cell(1, 1, "Topic").font = h_font;         ws3.cell(1, 1).fill = h_fill
    ws3.cell(1, 2, "Article Count").font = h_font;  ws3.cell(1, 2).fill = h_fill
    for ci in [1, 2]:
        ws3.cell(1, ci).alignment = Alignment(horizontal="center", vertical="center")
        ws3.cell(1, ci).border = brd
    ws3.column_dimensions["A"].width = 18
    ws3.column_dimensions["B"].width = 16
    for ri, (_, row) in enumerate(topic_counts.iterrows(), start=2):
        for ci, val in enumerate([row["Topic"], row["Article Count"]], 1):
            c = ws3.cell(ri, ci, val)
            c.font = nf
            c.fill = ev if ri % 2 == 0 else od
            c.border = brd
            c.alignment = Alignment(horizontal="center" if ci == 2 else "left", vertical="center")

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


# ─── Render helpers ────────────────────────────────────────────────────────
def kpi_card(label: str, value, color: str = "blue", delta: str = ""):
    delta_html = f'<div class="kpi-delta">{delta}</div>' if delta else ""
    return f"""
    <div class="kpi-card {color}">
        <div class="kpi-label">{label}</div>
        <div class="kpi-value">{value}</div>
        {delta_html}
    </div>"""

def article_card_html(headline: str, date: str, url: str, topic: str, capacity: str) -> str:
    cap_html = f'<span class="card-tag">⚡ {capacity}</span>' if capacity else ""
    return f"""
    <div class="article-card">
        <div class="card-headline">
            <a href="{url}" target="_blank">{headline}</a>
        </div>
        <div class="card-meta">
            <span class="card-date">📅 {date}</span>
            <span class="card-tag">{topic}</span>
            {cap_html}
            <a class="card-link-btn" href="{url}" target="_blank">↗ open</a>
        </div>
    </div>"""


# ─── Main App ──────────────────────────────────────────────────────────────
def main():
    st.set_page_config(
        page_title="DCD Intel | US Data Center Construction",
        page_icon="🏗️",
        layout="wide",
        initial_sidebar_state="expanded",
    )
    st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

    # ── Sidebar ────────────────────────────────────────────────────────────
    with st.sidebar:
        st.markdown("""
        <div style="padding:1rem 0 0.5rem;text-align:center;">
            <div style="font-family:'DM Mono',monospace;font-size:.65rem;letter-spacing:.2em;color:#3a5280;text-transform:uppercase;margin-bottom:.3rem;">Intelligence Platform</div>
            <div style="font-family:'Syne',sans-serif;font-size:1.2rem;font-weight:800;color:#fff;">DCD Intel</div>
            <div style="font-family:'DM Mono',monospace;font-size:.65rem;color:#3a5280;margin-top:.2rem;">US Construction Monitor</div>
        </div>
        """, unsafe_allow_html=True)
        st.divider()
        st.markdown("**📅 Date Range**")
        time_opt = st.radio(
            "", ["Latest (all available)", "Past 30 days", "Past 10 days"], index=0,
            label_visibility="collapsed",
        )
        days_map = {"Latest (all available)": None, "Past 30 days": 30, "Past 10 days": 10}
        sel_days = days_map[time_opt]
        st.markdown("**📄 Scrape Depth**")
        max_pages = st.slider("", 1, 30, 5, label_visibility="collapsed",
                               help="~28 articles per page")
        st.markdown(f"<div style='font-size:.75rem;color:#3a5280;margin-top:-.5rem;margin-bottom:.8rem;'>≈ {max_pages*28} articles to scan</div>", unsafe_allow_html=True)
        use_na = st.checkbox("🌎 Pre-filter: North America", value=True)
        st.divider()

        # Post-filters (shown after scrape)
        if "df_full" in st.session_state and st.session_state.df_full is not None:
            df_full = st.session_state.df_full
            st.markdown("**🔍 Refine Results**")
            all_topics  = sorted(df_full["Topic"].unique().tolist())
            all_states  = sorted(df_full["State"].unique().tolist())
            sel_topics  = st.multiselect("Topics", all_topics, default=all_topics)
            sel_states  = st.multiselect("States", all_states, default=all_states)
            keyword     = st.text_input("Keyword search", placeholder="e.g. Microsoft, 100MW…")
            st.session_state.filters = {
                "topics": sel_topics, "states": sel_states, "keyword": keyword
            }
        st.divider()
        go = st.button("🔍  Run Intelligence Scan", use_container_width=True, type="primary")

    # ── Banner ─────────────────────────────────────────────────────────────
    now_str = datetime.now().strftime("%A, %d %B %Y  ·  %H:%M UTC")
    st.markdown(f"""
    <div class="dcd-banner">
        <div class="banner-eyebrow">● Live Intelligence Feed</div>
        <div class="banner-title">US Data Center <span>Construction</span> Monitor</div>
        <div class="banner-sub">Real-time scrape of Data Center Dynamics · Filtered for American states & metros</div>
        <div class="banner-ts">🕐 {now_str}</div>
    </div>
    """, unsafe_allow_html=True)

    # ── Welcome state ──────────────────────────────────────────────────────
    if "df_full" not in st.session_state:
        st.session_state.df_full = None

    if not go and st.session_state.df_full is None:
        c1, c2, c3 = st.columns(3)
        c1.markdown("""
        <div class="kpi-card blue" style="opacity:.55">
            <div class="kpi-label">Articles indexed</div>
            <div class="kpi-value">—</div>
            <div class="kpi-delta">Run scan to populate</div>
        </div>""", unsafe_allow_html=True)
        c2.markdown("""
        <div class="kpi-card cyan" style="opacity:.55">
            <div class="kpi-label">US-related</div>
            <div class="kpi-value">—</div>
        </div>""", unsafe_allow_html=True)
        c3.markdown("""
        <div class="kpi-card green" style="opacity:.55">
            <div class="kpi-label">States covered</div>
            <div class="kpi-value">—</div>
        </div>""", unsafe_allow_html=True)
        st.markdown('<div class="section-header">How it works</div>', unsafe_allow_html=True)
        st.markdown("""
        <div style="display:flex;gap:1rem;flex-wrap:wrap;margin-bottom:1.5rem;">
            <div class="kpi-card" style="flex:1;min-width:200px;">
                <div style="font-size:1.5rem;margin-bottom:.5rem;">🕸️</div>
                <div style="font-family:'Syne',sans-serif;font-weight:700;color:#c8d4e8;margin-bottom:.3rem;">Smart Scrape</div>
                <div style="font-size:.82rem;color:#4a6490;">Fetches DCD's Construction Channel using Cloudflare bypass — server-rendered HTML, no browser needed.</div>
            </div>
            <div class="kpi-card" style="flex:1;min-width:200px;">
                <div style="font-size:1.5rem;margin-bottom:.5rem;">🗺️</div>
                <div style="font-family:'Syne',sans-serif;font-weight:700;color:#c8d4e8;margin-bottom:.3rem;">US State Filter</div>
                <div style="font-size:.82rem;color:#4a6490;">50 states + abbreviations + 80 major data center cities & counties — only American articles pass.</div>
            </div>
            <div class="kpi-card" style="flex:1;min-width:200px;">
                <div style="font-size:1.5rem;margin-bottom:.5rem;">🏷️</div>
                <div style="font-family:'Syne',sans-serif;font-weight:700;color:#c8d4e8;margin-bottom:.3rem;">Auto-Tagging</div>
                <div style="font-size:.82rem;color:#4a6490;">Articles auto-tagged by topic (Hyperscale, AI/GPU, Power, Investment…) and capacity (MW/GW).</div>
            </div>
            <div class="kpi-card" style="flex:1;min-width:200px;">
                <div style="font-size:1.5rem;margin-bottom:.5rem;">📊</div>
                <div style="font-family:'Syne',sans-serif;font-weight:700;color:#c8d4e8;margin-bottom:.3rem;">Rich Export</div>
                <div style="font-size:.82rem;color:#4a6490;">3-sheet Excel: articles with clickable links + State Breakdown + Topic Breakdown, all formatted.</div>
            </div>
        </div>
        """, unsafe_allow_html=True)
        return

    # ── Scrape ─────────────────────────────────────────────────────────────
    if go:
        st.session_state.filters = {"topics": [], "states": [], "keyword": ""}
        cutoff = datetime.min if sel_days is None else datetime.now() - timedelta(days=sel_days)
        pbar = st.progress(0, text="Initialising scan…")
        raw = scrape(cutoff, max_pages, use_na, pbar)
        pbar.empty()

        if not raw:
            st.error("No articles fetched. Check network / Cloudflare status.")
            return

        us_arts = [a for a in raw if is_us(a["Headline"])]
        us_arts = enrich(us_arts)

        df_full = (
            pd.DataFrame(us_arts)[["Headline","Date","URL","State","Topic","Capacity"]]
            .sort_values("Date", ascending=False)
            .reset_index(drop=True)
        )
        st.session_state.df_full   = df_full
        st.session_state.raw_count = len(raw)
        st.session_state.scan_time = datetime.now().strftime("%H:%M, %d %b %Y")
        st.rerun()

    # ── Dashboard ──────────────────────────────────────────────────────────
    df_full = st.session_state.df_full
    if df_full is None or df_full.empty:
        st.warning("No US articles found. Try increasing page count or disabling NA pre-filter.")
        return

    # Apply sidebar filters
    filters = st.session_state.get("filters", {})
    df = df_full.copy()
    if filters.get("topics"):
        df = df[df["Topic"].isin(filters["topics"])]
    if filters.get("states"):
        df = df[df["State"].isin(filters["states"])]
    if filters.get("keyword"):
        kw = filters["keyword"].lower()
        df = df[df["Headline"].str.lower().str.contains(kw, na=False)]
    df = df.reset_index(drop=True)

    # ── KPIs ──
    states_covered = df["State"].nunique()
    capacity_arts  = df[df["Capacity"] != ""]["Capacity"].count()
    scan_ts        = st.session_state.get("scan_time", "—")

    kpi_html = f"""
    <div class="kpi-row">
        {kpi_card("Articles Scanned", st.session_state.raw_count, "blue", "from DCD Construction Channel")}
        {kpi_card("US-Related", len(df_full), "cyan", "matched US state / city filter")}
        {kpi_card("Filtered View", len(df), "green", "after current filters")}
        {kpi_card("States Covered", states_covered, "gold", "distinct US states in results")}
        {kpi_card("Capacity Mentions", capacity_arts, "blue", "articles with MW/GW data")}
    </div>"""
    st.markdown(kpi_html, unsafe_allow_html=True)

    # ── Insight pills ──
    top_state  = df["State"].value_counts().idxmax() if not df.empty else "—"
    top_topic  = df["Topic"].value_counts().idxmax() if not df.empty else "—"
    latest_dt  = df["Date"].max() if not df.empty else "—"
    pills = f"""
    <div class="insight-row">
        <div class="insight-pill"><span class="dot"></span>Last scan: <b>{scan_ts}</b></div>
        <div class="insight-pill"><span class="dot"></span>Top state: <b>{top_state}</b></div>
        <div class="insight-pill"><span class="dot"></span>Top topic: <b>{top_topic}</b></div>
        <div class="insight-pill"><span class="dot"></span>Most recent article: <b>{latest_dt}</b></div>
        <div class="insight-pill"><span class="dot"></span>Source: <b>DataCenterDynamics.com</b></div>
    </div>"""
    st.markdown(pills, unsafe_allow_html=True)

    # ── Tabs ───────────────────────────────────────────────────────────────
    tab1, tab2, tab3, tab4 = st.tabs(["📰 Articles", "📊 Analytics", "🗺️ State View", "⬇️ Export"])

    # ── Tab 1: Articles ──
    with tab1:
        st.markdown('<div class="section-header">Latest US Construction Articles</div>', unsafe_allow_html=True)
        if df.empty:
            st.info("No articles match your current filters.")
        else:
            # Render each card individually — avoids Streamlit truncating large HTML blobs
            for _, row in df.iterrows():
                st.markdown(
                    article_card_html(
                        row["Headline"], row["Date"], row["URL"],
                        row["Topic"], row.get("Capacity", "")
                    ),
                    unsafe_allow_html=True,
                )

    # ── Tab 2: Analytics ──
    with tab2:
        import plotly.graph_objects as go
        import plotly.express as px

        _DARK_BG    = "#080c14"
        _PAPER_BG   = "#0d1628"
        _GRID_COL   = "#1a2b48"
        _TEXT_COL   = "#7a90b8"
        _TITLE_COL  = "#c8d4e8"
        _FONT       = "Inter, sans-serif"

        def _dark_layout(fig, title="", height=320):
            fig.update_layout(
                title=dict(text=title, font=dict(family=_FONT, size=13, color=_TITLE_COL), x=0.01),
                paper_bgcolor=_PAPER_BG,
                plot_bgcolor=_DARK_BG,
                font=dict(family=_FONT, color=_TEXT_COL),
                height=height,
                margin=dict(l=16, r=16, t=40, b=16),
                xaxis=dict(gridcolor=_GRID_COL, linecolor=_GRID_COL, tickfont=dict(size=11)),
                yaxis=dict(gridcolor=_GRID_COL, linecolor=_GRID_COL, tickfont=dict(size=11)),
                showlegend=False,
            )
            return fig

        TOPIC_PALETTE = {
            "Hyperscale":  "#0057ff",
            "Colocation":  "#00c6ff",
            "AI / GPU":    "#ff3b6b",
            "Power":       "#ffab00",
            "Investment":  "#a855f7",
            "Permits":     "#ff6d00",
            "Construction":"#00ffe7",
            "General":     "#3a5280",
        }

        # ── 1. Topic breakdown horizontal bar ──
        st.markdown('<div class="section-header">Articles by Topic</div>', unsafe_allow_html=True)
        topic_counts = df["Topic"].value_counts().reset_index()
        topic_counts.columns = ["Topic", "Count"]
        topic_counts = topic_counts.sort_values("Count")
        fig_topic = go.Figure(go.Bar(
            x=topic_counts["Count"],
            y=topic_counts["Topic"],
            orientation="h",
            marker=dict(
                color=[TOPIC_PALETTE.get(t, "#3a5280") for t in topic_counts["Topic"]],
                line=dict(width=0),
            ),
            text=topic_counts["Count"],
            textposition="outside",
            textfont=dict(color=_TITLE_COL, size=12),
            hovertemplate="<b>%{y}</b><br>Articles: %{x}<extra></extra>",
        ))
        _dark_layout(fig_topic, height=300)
        fig_topic.update_layout(xaxis_title="", yaxis_title="")
        st.plotly_chart(fig_topic, use_container_width=True, config={"displayModeBar": False})

        # ── 2. Top states bar chart ──
        st.markdown('<div class="section-header">Top 15 States by Article Volume</div>', unsafe_allow_html=True)
        state_counts = df["State"].value_counts().head(15).reset_index()
        state_counts.columns = ["State", "Count"]
        state_counts = state_counts.sort_values("Count")

        # gradient blue → cyan by rank
        n = len(state_counts)
        bar_colors = [f"rgba({int(0 + (0-0)*i/max(n-1,1))}, {int(87 + (198-87)*i/max(n-1,1))}, {int(255 + (255-255)*i/max(n-1,1))}, 0.85)" for i in range(n)]

        fig_state = go.Figure(go.Bar(
            x=state_counts["Count"],
            y=state_counts["State"],
            orientation="h",
            marker=dict(color=bar_colors, line=dict(width=0)),
            text=state_counts["Count"],
            textposition="outside",
            textfont=dict(color=_TITLE_COL, size=11),
            hovertemplate="<b>%{y}</b><br>Articles: %{x}<extra></extra>",
        ))
        _dark_layout(fig_state, height=420)
        st.plotly_chart(fig_state, use_container_width=True, config={"displayModeBar": False})

        # ── 3. Articles over time ──
        st.markdown('<div class="section-header">Publication Volume Over Time</div>', unsafe_allow_html=True)
        df_time = df[df["Date"] != "Unknown"].copy()
        if not df_time.empty:
            df_time["Date_dt"] = pd.to_datetime(df_time["Date"])
            daily = df_time.groupby(df_time["Date_dt"].dt.date).size().reset_index()
            daily.columns = ["Date", "Articles"]
            fig_time = go.Figure()
            fig_time.add_trace(go.Scatter(
                x=daily["Date"], y=daily["Articles"],
                mode="lines+markers",
                line=dict(color="#00c6ff", width=2.5),
                marker=dict(color="#0057ff", size=6, line=dict(color="#00c6ff", width=1.5)),
                fill="tozeroy",
                fillcolor="rgba(0,87,255,0.08)",
                hovertemplate="<b>%{x}</b><br>Articles: %{y}<extra></extra>",
            ))
            _dark_layout(fig_time, height=260)
            fig_time.update_layout(xaxis_title="", yaxis_title="Articles")
            st.plotly_chart(fig_time, use_container_width=True, config={"displayModeBar": False})

        # ── 4. Topic share donut ──
        st.markdown('<div class="section-header">Topic Share</div>', unsafe_allow_html=True)
        topic_pie = df["Topic"].value_counts().reset_index()
        topic_pie.columns = ["Topic", "Count"]
        fig_pie = go.Figure(go.Pie(
            labels=topic_pie["Topic"],
            values=topic_pie["Count"],
            hole=0.55,
            marker=dict(
                colors=[TOPIC_PALETTE.get(t, "#3a5280") for t in topic_pie["Topic"]],
                line=dict(color=_DARK_BG, width=2),
            ),
            textinfo="label+percent",
            textfont=dict(color=_TITLE_COL, size=12),
            hovertemplate="<b>%{label}</b><br>%{value} articles (%{percent})<extra></extra>",
        ))
        fig_pie.update_layout(
            paper_bgcolor=_PAPER_BG,
            plot_bgcolor=_DARK_BG,
            font=dict(family=_FONT, color=_TEXT_COL),
            height=380,
            margin=dict(l=16, r=16, t=30, b=16),
            showlegend=True,
            legend=dict(
                bgcolor="rgba(0,0,0,0)",
                font=dict(color=_TITLE_COL, size=11),
                orientation="v",
            ),
            annotations=[dict(
                text=f"<b>{len(df)}</b><br>articles",
                x=0.5, y=0.5,
                font=dict(size=16, color=_TITLE_COL, family=_FONT),
                showarrow=False,
            )],
        )
        st.plotly_chart(fig_pie, use_container_width=True, config={"displayModeBar": False})

        # ── 5. Capacity table ──
        st.markdown('<div class="section-header">⚡ Capacity Mentions</div>', unsafe_allow_html=True)
        cap_df = df[df["Capacity"] != ""][["Headline", "Capacity", "Date", "State", "Topic"]].head(20)
        if not cap_df.empty:
            st.dataframe(
                cap_df, use_container_width=True, hide_index=True,
                column_config={"Capacity": st.column_config.TextColumn("⚡ Capacity")},
            )
        else:
            st.info("No capacity mentions found in current filter.")

    # ── Tab 3: State View ──
    with tab3:
        st.markdown('<div class="section-header">State Breakdown</div>', unsafe_allow_html=True)
        state_df = df["State"].value_counts().reset_index()
        state_df.columns = ["State","Articles"]
        col1, col2 = st.columns([1, 2])
        with col1:
            st.dataframe(state_df, use_container_width=True, hide_index=True, height=420)
        with col2:
            sel_state = st.selectbox("Drill into a state", state_df["State"].tolist())
            state_articles = df[df["State"] == sel_state]
            st.markdown(f"**{len(state_articles)} articles** for **{sel_state}**")
            for _, row in state_articles.iterrows():
                st.markdown(
                    f"<div class='article-card'>"
                    f"<div class='card-headline'><a href='{row['URL']}' target='_blank'>{row['Headline']}</a></div>"
                    f"<div class='card-meta'><span class='card-date'>📅 {row['Date']}</span>"
                    f"<span class='card-tag'>{row['Topic']}</span></div></div>",
                    unsafe_allow_html=True
                )

    # ── Tab 4: Export ──
    with tab4:
        st.markdown('<div class="section-header">Export Data</div>', unsafe_allow_html=True)

        col_e1, col_e2 = st.columns(2)
        with col_e1:
            st.markdown("""
            <div class="kpi-card blue">
                <div style="font-size:1.8rem;margin-bottom:.5rem;">📊</div>
                <div style="font-family:'Syne',sans-serif;font-weight:700;color:#fff;font-size:1rem;margin-bottom:.4rem;">Excel Report (.xlsx)</div>
                <div style="font-size:.82rem;color:#4a6490;margin-bottom:.8rem;">
                    3 sheets: Articles (with clickable URLs) · State Breakdown · Topic Breakdown.<br>
                    Colour-coded topics, auto-filter, frozen header row.
                </div>
            </div>
            """, unsafe_allow_html=True)
            ts    = datetime.now().strftime("%Y%m%d_%H%M")
            label = re.sub(r"[^a-zA-Z0-9]", "_", time_opt)
            st.download_button(
                "📥 Download Excel Report",
                data=build_excel(df),
                file_name=f"DCD_US_Construction_{label}_{ts}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        with col_e2:
            st.markdown("""
            <div class="kpi-card cyan">
                <div style="font-size:1.8rem;margin-bottom:.5rem;">📄</div>
                <div style="font-family:'Syne',sans-serif;font-weight:700;color:#fff;font-size:1rem;margin-bottom:.4rem;">CSV Export</div>
                <div style="font-size:.82rem;color:#4a6490;margin-bottom:.8rem;">
                    Flat CSV of the current filtered view — ready for Excel, Python, or BI tools.
                </div>
            </div>
            """, unsafe_allow_html=True)
            csv_bytes = df.to_csv(index=False).encode()
            st.download_button(
                "📥 Download CSV",
                data=csv_bytes,
                file_name=f"DCD_US_Construction_{label}_{ts}.csv",
                mime="text/csv",
                use_container_width=True,
            )

        st.markdown('<div class="section-header">Preview (filtered)</div>', unsafe_allow_html=True)
        st.dataframe(
            df[["Headline","Date","State","Topic","Capacity","URL"]],
            use_container_width=True, height=400,
            column_config={"URL": st.column_config.LinkColumn("URL", display_text="🔗 Open")},
            hide_index=True,
        )


if __name__ == "__main__":
    main()

