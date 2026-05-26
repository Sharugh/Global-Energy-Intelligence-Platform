import streamlit as st
import re
import io
import time
import math
import textwrap
from collections import Counter
from datetime import datetime, timedelta
from difflib import SequenceMatcher

# ── Platform metadata ─────────────────────────────────────────────────────────
_PLAT_VER   = "2.0.0 (Wood Mac Energy Expansion)"
_PLAT_TOKEN = "\x53\x68\x61\x72\x75\x67\x68\x20\x41"

# ── Document export libraries ────────────────────────────────────────────────
try:
    from docx import Document as DocxDocument
    from docx.shared import Pt, RGBColor, Inches, Cm
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    _DOCX_OK = True
except ImportError:
    _DOCX_OK = False

try:
    from reportlab.lib.pagesizes import letter
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import inch
    from reportlab.lib import colors as rl_colors
    from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, HRFlowable, KeepTogether)
    _PDF_OK = True
except ImportError:
    _PDF_OK = False

import pandas as pd
import plotly.graph_objects as go
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

try:
    import cloudscraper
    _CS = cloudscraper.create_scraper(browser={"browser": "chrome", "platform": "windows", "mobile": False})
    _USE_CS = True
except ImportError:
    import requests
    _CS = requests.Session()
    _USE_CS = False

from bs4 import BeautifulSoup

# ── CSS ─────────────────────────────────────────────────────────────────────
CUSTOM_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Syne:wght@400;600;700;800&family=DM+Mono:wght@400;500&family=Inter:wght@300;400;500&display=swap');
html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
.stApp { background: #060a10; color: #e8edf5; }
* { transition: background 0.15s ease, border-color 0.15s ease, box-shadow 0.18s ease; }
[data-testid="stSidebar"] { background: #0a0f1a !important; border-right: 1px solid #151f35; }
[data-testid="stSidebar"] * { color: #b8c8e0 !important; }
[data-testid="stSidebar"] hr { border-color: #151f35 !important; }
[data-testid="stSidebarCollapseButton"] { display: none !important; }
[data-testid="stSidebar"][aria-expanded="false"], [data-testid="stSidebar"] {
    display: block !important; visibility: visible !important; opacity: 1 !important;
    min-width: 280px !important; max-width: 280px !important; width: 280px !important;
    transform: none !important; margin-left: 0 !important;
}
[data-testid="stSidebar"] .stButton button {
    background: linear-gradient(135deg, #0047e1, #00b4ff) !important; color: #fff !important; 
    border: none !important; border-radius: 8px !important; font-family: 'Syne', sans-serif !important; 
    font-weight: 700 !important; padding: 0.65rem 1rem !important; transition: opacity .2s;
}
[data-testid="stSidebar"] .stButton button:hover { opacity: .82; transform: translateY(-1px) !important; box-shadow: 0 4px 16px rgba(0,71,225,0.35) !important; }
.stMultiSelect [data-baseweb="select"] { background: #0b1628 !important; border-color: #152038 !important; }
[data-baseweb="popover"], [data-baseweb="menu"], [role="listbox"], ul[data-baseweb="menu"] { background: #0d1a2e !important; border: 1px solid #1e3050 !important; }
[data-baseweb="menu"] li, [role="option"] { background: #0d1a2e !important; color: #b8c8e0 !important; }
[data-baseweb="menu"] li:hover, [role="option"]:hover { background: #0f2245 !important; color: #ffffff !important; }
.stTextInput input, .stNumberInput input { background: #0b1628 !important; border: 1px solid #152038 !important; border-radius: 8px !important; color: #d0dff0 !important; }
.stTabs [data-baseweb="tab-list"] { background: #0b1628 !important; border-radius: 10px !important; padding: 4px !important; border: 1px solid #152038; }
.stTabs [data-baseweb="tab"] { background: transparent !important; color: #3a5480 !important; font-family: 'Syne', sans-serif !important; font-weight: 600 !important; }
.stTabs [aria-selected="true"] { background: #0047e1 !important; color: #fff !important; }
.stDownloadButton button { background: linear-gradient(135deg, #002d0a, #005214) !important; color: #00e676 !important; border: 1px solid #00a846 !important; }
.gl-banner { background: linear-gradient(135deg, #07111f 0%, #0b1d3a 45%, #07111f 100%); border: 1px solid #132040; border-radius: 16px; padding: 2rem 2.5rem; margin-bottom: 1.6rem; position: relative; overflow: hidden; }
.banner-eyebrow { font-family: 'DM Mono', monospace; font-size: .68rem; letter-spacing: .2em; color: #00b4ff; text-transform: uppercase; margin-bottom: .45rem; }
.banner-title { font-family: 'Syne', sans-serif; font-size: 2rem; font-weight: 800; color: #fff; line-height: 1.12; }
.banner-title span { color: #00b4ff; }
.sec-head { font-family: 'Syne', sans-serif; font-size: .9rem; font-weight: 700; color: #b8c8e0; border-left: 3px solid #0047e1; padding-left: .7rem; margin: 1.6rem 0 .9rem 0; text-transform: uppercase; }
hr { border-color: #152038 !important; }
#MainMenu, footer, header { visibility: hidden; }
</style>
"""

# ── Source URLs ──────────────────────────────────────────────────────────────
DCD_CONSTRUCTION_URL = "https://www.datacenterdynamics.com/en/news/?term=the-data-center-construction-channel"
DCD_GENERAL_URL      = "https://www.datacenterdynamics.com/en/news/"

RSS_FEEDS = {
    "Renewable Energy World": "https://www.renewableenergyworld.com/feed/",
    "Utility Dive": "https://www.utilitydive.com/feeds/news/",
    "Power Technology": "https://www.power-technology.com/feed/"
}

# ── Taxonomy & Keywords ─────────────────────────────────────────────────────
TOPIC_COLORS = {
    "Hyperscale": "#0047e1", "Colocation": "#00b4ff", "AI / GPU": "#ff2d6b",
    "Solar": "#ffd700", "Wind - Offshore": "#00e5c8", "Wind - Onshore": "#00b4ff",
    "Nuclear": "#00e676", "Gas / Thermal": "#ff6400", "Battery Storage": "#a855f7",
    "Power": "#ffaa00", "Investment": "#a855f7", "Permits": "#ff6400",
    "Construction": "#00e5c8", "Sustainability": "#00e676", "General": "#2e4470"
}

REGION_COLORS = {
    "North America": "#0047e1", "Europe": "#00b4ff", "Asia Pacific": "#ff2d6b",
    "Middle East": "#ffaa00", "Latin America": "#a855f7", "Africa": "#00e676", "Global": "#2e4470"
}

SOURCE_META = {
    "DCD Construction": {"color": "#0047e1", "short": "DCD"},
    "DCD General News": {"color": "#00b4ff", "short": "DCD"},
    "Renewable Energy World": {"color": "#ffd700", "short": "REW"},
    "Utility Dive": {"color": "#ffaa00", "short": "UDV"},
    "Power Technology": {"color": "#00e5c8", "short": "PWR"},
    "Unknown": {"color": "#2e4470", "short": "UNK"},
}

COUNTRY_TO_REGION = {
    "United States": "North America", "Canada": "North America", "Mexico": "North America",
    "United Kingdom": "Europe", "Germany": "Europe", "France": "Europe", "Netherlands": "Europe",
    "Ireland": "Europe", "Sweden": "Europe", "Norway": "Europe", "Denmark": "Europe",
    "Spain": "Europe", "Italy": "Europe", "Poland": "Europe", "Switzerland": "Europe",
    "Singapore": "Asia Pacific", "Japan": "Asia Pacific", "South Korea": "Asia Pacific",
    "Australia": "Asia Pacific", "India": "Asia Pacific", "China": "Asia Pacific",
    "Saudi Arabia": "Middle East", "UAE": "Middle East", "Qatar": "Middle East",
    "Brazil": "Latin America", "Chile": "Latin America", "Colombia": "Latin America",
    "South Africa": "Africa", "Nigeria": "Africa", "Kenya": "Africa"
}

COUNTRY_KEYWORDS = {k: [k] for k in COUNTRY_TO_REGION}
# Add abbreviations for major markets
COUNTRY_KEYWORDS["United States"].extend(["U.S.", "US", "America", "Virginia", "Texas", "California"])
COUNTRY_KEYWORDS["United Kingdom"].extend(["UK", "England", "Scotland", "London"])
COUNTRY_KEYWORDS["United Arab Emirates"] = ["UAE", "Dubai", "Abu Dhabi"]

TOPIC_KEYWORDS = {
    "Hyperscale": ["hyperscale", "microsoft", "google", "amazon", "aws", "meta", "cloud region"],
    "Colocation": ["colocation", "equinix", "digital realty", "cyrusone", "vantage", "wholesale"],
    "AI / GPU": ["ai", "gpu", "nvidia", "inference", "llm", "ai factory"],
    "Solar": ["solar", "pv", "photovoltaic", "solar farm", "solar project"],
    "Wind - Offshore": ["offshore wind", "floating wind", "monopile", "offshore turbine"],
    "Wind - Onshore": ["onshore wind", "wind farm", "wind project", "wind turbine"],
    "Nuclear": ["nuclear", "smr", "reactor", "uranium", "fission"],
    "Gas / Thermal": ["natural gas", "ccgt", "gas turbine", "lng", "thermal power", "coal"],
    "Battery Storage": ["bess", "battery storage", "energy storage", "megapack"],
    "Power": ["mw", "gw", "megawatt", "gigawatt", "ppa", "grid connection", "substation"],
    "Investment": ["invest", "fund", "acquisition", "billion", "million", "financing"],
    "Permits": ["permit", "zoning", "approved", "moratorium", "planning", "rezoning"],
    "Construction": ["broke ground", "construction", "topping out", "commissioning"]
}

KNOWN_COMPANIES = [
    "Microsoft", "Google", "Amazon", "AWS", "Meta", "Oracle", "Equinix", "Digital Realty",
    "CyrusOne", "Vantage", "NextEra", "Orsted", "Vestas", "Enel", "Iberdrola", "RWE", "EDF",
    "Siemens Gamesa", "GE Vernova", "First Solar", "Dominion Energy", "Duke Energy"
]

GRID_OPERATORS = ["PJM", "ERCOT", "CAISO", "MISO", "NYISO", "ISO-NE", "SPP", "National Grid", "EirGrid", "AEMO"]

# ── Parsing Logic ───────────────────────────────────────────────────────────
def parse_date_str(raw):
    if not raw: return None
    raw = str(raw).strip()
    for fmt in ("%Y-%m-%dT%H:%M:%S%z", "%Y-%m-%dT%H:%M:%SZ", "%a, %d %b %Y %H:%M:%S %z", "%Y-%m-%d"):
        try: return datetime.strptime(raw[:25], fmt[:len(raw[:25])]).replace(tzinfo=None)
        except: pass
    m = re.search(r"(\d{1,2})\s+(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)\w*\s+(\d{4})", raw, re.I)
    if m:
        months = {"jan":1,"feb":2,"mar":3,"apr":4,"may":5,"jun":6,"jul":7,"aug":8,"sep":9,"oct":10,"nov":11,"dec":12}
        try: return datetime(int(m.group(3)), months[m.group(2).lower()[:3]], int(m.group(1)))
        except: pass
    return None

def detect_country(text):
    for country, patterns in COUNTRY_KEYWORDS.items():
        for pat in patterns:
            if re.search(r"\b" + re.escape(pat) + r"\b", text, re.I):
                return country
    return "Global"

def detect_topic(text):
    t = text.lower()
    for topic, kws in TOPIC_KEYWORDS.items():
        if any(k.lower() in t for k in kws):
            return topic
    return "General"

def detect_mw(text):
    m = re.search(r"([\d,]+(?:\.\d+)?)\s*(GW|MW)", text, re.I)
    if m: return m.group(1).replace(",", "") + " " + m.group(2).upper()
    return ""

def detect_deal(text):
    m = re.search(r"(\$|€|£)\s*([\d,.]+)\s*(billion|bn|million|m\b)", text, re.I)
    if m:
        sym = m.group(1)
        val = m.group(2).replace(",", "")
        unit = "bn" if m.group(3).lower() in ["billion", "bn"] else "m"
        return f"{sym}{val}{unit}"
    return ""

def detect_sentiment(text):
    t = text.lower()
    if any(w in t for w in ["broke ground", "opens", "inaugurated", "live", "commercial operation"]): return "Opened / Live"
    if any(w in t for w in ["approved", "permits", "rezoning"]): return "Approved"
    if any(w in t for w in ["proposed", "plans", "files for"]): return "Proposed"
    if any(w in t for w in ["rejected", "moratorium", "blocked", "lawsuit"]): return "Challenged"
    if any(w in t for w in ["under construction", "building"]): return "Under Construction"
    return "News"

# ── Scrapers ────────────────────────────────────────────────────────────────
def _scrape_dcd(url, source_name, cutoff, pages):
    arts = []
    for page in range(1, pages + 1):
        target = f"{url}&page={page}" if "?" in url else f"{url}?page={page}"
        try:
            r = _CS.get(target, timeout=10)
            soup = BeautifulSoup(r.text, "html.parser")
            for a in soup.find_all("a", href=re.compile(r"^/en/news/[^?#]+/$")):
                headline = a.get_text(" ", strip=True)
                if len(headline) < 15: continue
                
                date_obj = None
                time_tag = a.parent.find("time") if a.parent else None
                if time_tag: date_obj = parse_date_str(time_tag.get("datetime", ""))
                
                if date_obj and date_obj < cutoff: return arts
                arts.append({
                    "headline": headline,
                    "url": "https://datacenterdynamics.com" + a["href"],
                    "date_obj": date_obj,
                    "source": source_name,
                })
        except:
            break
        time.sleep(0.5)
    return arts

def _scrape_rss(url, source_name, cutoff):
    arts = []
    try:
        r = _CS.get(url, timeout=10)
        soup = BeautifulSoup(r.content, "xml")
        for item in soup.find_all("item")[:50]: # limit to 50 per feed
            title = item.title.text if item.title else ""
            link = item.link.text if item.link else ""
            pub_date = item.pubDate.text if item.pubDate else ""
            
            date_obj = parse_date_str(pub_date)
            if date_obj and date_obj < cutoff: continue
            
            if title and link:
                arts.append({
                    "headline": title,
                    "url": link,
                    "date_obj": date_obj,
                    "source": source_name,
                })
    except Exception as e:
        st.sidebar.error(f"Failed to fetch {source_name}: {e}")
    return arts

# ── Main UI & Logic ─────────────────────────────────────────────────────────
def main():
    st.set_page_config(page_title="Wood Mac Infrastructure Intelligence", page_icon="⚡", layout="wide")
    st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

    if "df_full" not in st.session_state:
        st.session_state.df_full = pd.DataFrame()

    with st.sidebar:
        st.markdown(
            '<div style="padding:.9rem 0 .4rem;">'
            '<div style="font-family:Syne,sans-serif;font-size:.82rem;font-weight:700;color:#b8c8e0;letter-spacing:.02em;margin-bottom:.06rem;">Wood Mackenzie</div>'
            '<div style="font-family:Syne,sans-serif;font-size:.82rem;font-weight:700;color:#00b4ff;letter-spacing:.02em;margin-bottom:.28rem;">Infrastructure Intelligence</div>'
            '</div>', unsafe_allow_html=True
        )
        st.divider()

        # Pre-Scan Configuration
        st.markdown('<div class="sec-head" style="margin-top:0;">1. Data Sources</div>', unsafe_allow_html=True)
        sources = st.multiselect("Select feeds to scrape", 
                                 ["DCD Construction", "DCD General News", "Renewables & Power (RSS)"],
                                 default=["DCD Construction", "Renewables & Power (RSS)"])
        
        days_back = st.slider("Days to look back", 1, 30, 7)
        max_pages = st.slider("DCD Scrape Depth (pages)", 1, 50, 5)

        st.divider()

        # Pre-Scan Filters (These apply to the dataframe instantly after scraping)
        st.markdown('<div class="sec-head" style="margin-top:0;">2. Pre-Filters</div>', unsafe_allow_html=True)
        
        sel_regions = st.multiselect("Region", list(set(COUNTRY_TO_REGION.values())))
        sel_countries = st.multiselect("Country", list(COUNTRY_TO_REGION.keys()))
        sel_topics = st.multiselect("Topic", list(TOPIC_KEYWORDS.keys()))
        sel_companies = st.multiselect("Company", KNOWN_COMPANIES)
        kw_filter = st.text_input("Keyword search")
        min_mw = st.number_input("Minimum Capacity (MW)", 0, step=50)

        go_btn = st.button("⚡ Run Global Scan", use_container_width=True, type="primary")

    # Header
    st.markdown(
        f'<div class="gl-banner">'
        f'<div class="banner-eyebrow">● Live Intelligence Feed  ·  Data Centers & Renewables</div>'
        f'<div class="banner-title">Global Infrastructure</div>'
        f'<div class="banner-title" style="margin-top:-.15rem;"><span>Intelligence</span> Platform</div>'
        f'<div class="banner-sub">Real-time tracking of Data Centers, Solar, Wind, BESS, and Grid Infrastructure.</div>'
        f'</div>',
        unsafe_allow_html=True,
    )

    if go_btn:
        cutoff = datetime.now() - timedelta(days=days_back)
        raw_arts = []
        
        with st.spinner("Scraping intelligence sources..."):
            if "DCD Construction" in sources:
                raw_arts.extend(_scrape_dcd(DCD_CONSTRUCTION_URL, "DCD Construction", cutoff, max_pages))
            if "DCD General News" in sources:
                raw_arts.extend(_scrape_dcd(DCD_GENERAL_URL, "DCD General News", cutoff, max_pages))
            if "Renewables & Power (RSS)" in sources:
                for name, url in RSS_FEEDS.items():
                    raw_arts.extend(_scrape_rss(url, name, cutoff))

        # Enrichment
        enriched = []
        for art in raw_arts:
            hl = art["headline"]
            c = detect_country(hl)
            enriched.append({
                "Headline": hl,
                "Date": art["date_obj"].strftime("%Y-%m-%d") if art["date_obj"] else "Unknown",
                "Source": art["source"],
                "Region": COUNTRY_TO_REGION.get(c, "Global"),
                "Country": c,
                "Topic": detect_topic(hl),
                "Capacity": detect_mw(hl),
                "Deal Size": detect_deal(hl),
                "Sentiment": detect_sentiment(hl),
                "URL": art["url"]
            })
        
        # Deduplicate
        df_full = pd.DataFrame(enriched).drop_duplicates(subset=["URL"]).reset_index(drop=True)
        st.session_state.df_full = df_full

    df = st.session_state.df_full

    if df.empty:
        st.info("Configure sources and filters in the sidebar, then click **Run Global Scan**.")
        return

    # Apply Pre-Filters
    if sel_regions: df = df[df["Region"].isin(sel_regions)]
    if sel_countries: df = df[df["Country"].isin(sel_countries)]
    if sel_topics: df = df[df["Topic"].isin(sel_topics)]
    if kw_filter: df = df[df["Headline"].str.contains(kw_filter, case=False, na=False)]
    if min_mw > 0:
        def get_mw(cap_str):
            m = re.search(r"([\d,.]+)\s*(GW|MW)", str(cap_str), re.I)
            if not m: return 0
            v = float(m.group(1).replace(",", ""))
            return v * 1000 if m.group(2).upper() == "GW" else v
        df = df[df["Capacity"].apply(get_mw) >= min_mw]

    # Metrics
    st.markdown(f"### 📊 Scan Results: {len(df)} Articles")
    
    # Simple Table display for brevity (You can swap back your advanced dark_table here)
    st.dataframe(
        df[["Headline", "Topic", "Capacity", "Region", "Date", "Source", "URL"]],
        use_container_width=True,
        hide_index=True
    )
    
    # Basic Analytics
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("#### Top Topics")
        st.bar_chart(df["Topic"].value_counts())
    with c2:
        st.markdown("#### Top Regions")
        st.bar_chart(df["Region"].value_counts())

    # Export
    csv = df.to_csv(index=False).encode('utf-8')
    st.download_button(
        label="📥 Download Data as CSV",
        data=csv,
        file_name='woodmac_infrastructure_intel.csv',
        mime='text/csv',
    )

if __name__ == "__main__":
    main()
