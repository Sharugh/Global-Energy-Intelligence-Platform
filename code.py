import streamlit as st
import re
import io
import time
import feedparser
from datetime import datetime, timedelta
from concurrent.futures import ThreadPoolExecutor, as_completed
from difflib import SequenceMatcher

import pandas as pd
import plotly.graph_objects as go
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

try:
    import cloudscraper
    _CS = cloudscraper.create_scraper(
        browser={"browser": "chrome", "platform": "windows", "mobile": False}
    )
    _USE_CS = True
except ImportError:
    import requests
    _CS = requests.Session()
    _USE_CS = False

from bs4 import BeautifulSoup

_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
}

CUSTOM_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Syne:wght@400;600;700;800&family=DM+Mono:wght@400;500&family=Inter:wght@300;400;500&display=swap');

html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
.stApp { background: #060a10; color: #e8edf5; }

[data-testid="stSidebar"] { background: #0a0f1a !important; border-right: 1px solid #151f35; }
[data-testid="stSidebar"] * { color: #b8c8e0 !important; }
[data-testid="stSidebar"] .stButton button {
    background: linear-gradient(135deg, #0047e1, #00b4ff) !important;
    color: #fff !important; border: none !important; border-radius: 8px !important;
    font-family: 'Syne', sans-serif !important; font-weight: 700 !important;
    font-size: 0.92rem !important; letter-spacing: 0.04em !important;
    padding: 0.65rem 1rem !important; transition: opacity .2s;
}
[data-testid="stSidebar"] .stButton button:hover { opacity: .82; }
[data-testid="stSidebar"] hr { border-color: #151f35 !important; }

.gl-banner {
    background: linear-gradient(135deg, #07111f 0%, #0b1d3a 45%, #07111f 100%);
    border: 1px solid #132040; border-radius: 16px;
    padding: 2rem 2.5rem; margin-bottom: 1.6rem;
    position: relative; overflow: hidden;
}
.gl-banner::before {
    content: ''; position: absolute; top: -80px; right: -40px;
    width: 320px; height: 320px;
    background: radial-gradient(circle, rgba(0,71,225,0.16) 0%, transparent 68%);
    border-radius: 50%;
}
.gl-banner::after {
    content: ''; position: absolute; bottom: -50px; left: 25%;
    width: 240px; height: 240px;
    background: radial-gradient(circle, rgba(0,180,255,0.09) 0%, transparent 68%);
    border-radius: 50%;
}
.banner-eyebrow {
    font-family: 'DM Mono', monospace; font-size: .68rem;
    letter-spacing: .2em; color: #00b4ff;
    text-transform: uppercase; margin-bottom: .45rem;
}
.banner-title {
    font-family: 'Syne', sans-serif; font-size: 2rem;
    font-weight: 800; color: #fff; line-height: 1.12; margin-bottom: .35rem;
}
.banner-title span { color: #00b4ff; }
.banner-sub { font-size: .85rem; color: #6a80a8; font-weight: 300; }
.banner-ts { font-family: 'DM Mono', monospace; font-size: .68rem; color: #2a3e60; margin-top: .7rem; letter-spacing: .05em; }

.sec-head {
    font-family: 'Syne', sans-serif; font-size: .9rem; font-weight: 700;
    color: #b8c8e0; letter-spacing: .07em; text-transform: uppercase;
    border-left: 3px solid #0047e1; padding-left: .7rem;
    margin: 1.6rem 0 .9rem 0;
}
.pill-row { display: flex; flex-wrap: wrap; gap: .45rem; margin-bottom: 1.4rem; }
.pill {
    background: #0b1628; border: 1px solid #152038; border-radius: 20px;
    padding: .38rem .95rem; font-size: .78rem; color: #6a80a8;
    display: flex; align-items: center; gap: .38rem;
}
.pill b { color: #fff; }
.pill-dot { width: 6px; height: 6px; border-radius: 50%; background: #0047e1; flex-shrink: 0; }

.stTabs [data-baseweb="tab-list"] {
    background: #0b1628 !important; border-radius: 10px !important;
    padding: 4px !important; gap: 3px !important; border: 1px solid #152038;
}
.stTabs [data-baseweb="tab"] {
    background: transparent !important; color: #3a5480 !important;
    border-radius: 7px !important; font-family: 'Syne', sans-serif !important;
    font-weight: 600 !important; font-size: .8rem !important;
    letter-spacing: .04em !important; padding: .45rem 1.1rem !important;
}
.stTabs [aria-selected="true"] { background: #0047e1 !important; color: #fff !important; }
.stTabs [data-baseweb="tab-panel"] { padding-top: 1.1rem !important; }

.stDownloadButton button {
    background: linear-gradient(135deg, #002d0a, #005214) !important;
    color: #00e676 !important; border: 1px solid #00a846 !important;
    border-radius: 8px !important; font-family: 'Syne', sans-serif !important;
    font-weight: 700 !important; letter-spacing: .04em !important;
}
.stDownloadButton button:hover { opacity: .82 !important; }
.stTextInput input {
    background: #0b1628 !important; border: 1px solid #152038 !important;
    border-radius: 8px !important; color: #d0dff0 !important;
}
.stTextInput input:focus { border-color: #0047e1 !important; }
.stMultiSelect [data-baseweb="select"] { background: #0b1628 !important; border-color: #152038 !important; }
.stSelectbox [data-baseweb="select"] > div { background: #0b1628 !important; border-color: #152038 !important; }
hr { border-color: #152038 !important; }
#MainMenu, footer, header { visibility: hidden; }
</style>
"""

MONTHS = {
    "jan":1,"feb":2,"mar":3,"apr":4,"may":5,"jun":6,
    "jul":7,"aug":8,"sep":9,"oct":10,"nov":11,"dec":12,
}

TOPIC_COLORS = {
    "Hyperscale":   "#0047e1",
    "Colocation":   "#00b4ff",
    "AI / GPU":     "#ff2d6b",
    "Power":        "#ffaa00",
    "Investment":   "#a855f7",
    "Permits":      "#ff6400",
    "Construction": "#00e5c8",
    "Sustainability":"#00e676",
    "General":      "#2e4470",
}

REGION_COLORS = {
    "North America":  "#0047e1",
    "Europe":         "#00b4ff",
    "Asia Pacific":   "#ff2d6b",
    "Middle East":    "#ffaa00",
    "Latin America":  "#a855f7",
    "Africa":         "#00e676",
    "Global":         "#2e4470",
}

SOURCE_META = {
    "DataCenterDynamics": {"color": "#0047e1", "short": "DCD"},
    "DataCenter Knowledge": {"color": "#00b4ff", "short": "DCK"},
    "DataCenterFrontier":  {"color": "#00e5c8", "short": "DCF"},
    "Google News":         {"color": "#ffaa00", "short": "GNS"},
    "PR Newswire":         {"color": "#a855f7", "short": "PRN"},
    "BusinessWire":        {"color": "#ff6400", "short": "BIZ"},
    "Reuters":             {"color": "#ff2d6b", "short": "REU"},
    "Unknown":             {"color": "#2e4470", "short": "UNK"},
}

COUNTRY_TO_REGION = {
    "United States": "North America", "Canada": "North America", "Mexico": "North America",
    "United Kingdom": "Europe", "Germany": "Europe", "France": "Europe",
    "Netherlands": "Europe", "Ireland": "Europe", "Sweden": "Europe",
    "Norway": "Europe", "Denmark": "Europe", "Finland": "Europe",
    "Spain": "Europe", "Italy": "Europe", "Poland": "Europe",
    "Switzerland": "Europe", "Austria": "Europe", "Belgium": "Europe",
    "Portugal": "Europe", "Romania": "Europe", "Czech Republic": "Europe",
    "Singapore": "Asia Pacific", "Japan": "Asia Pacific", "South Korea": "Asia Pacific",
    "Australia": "Asia Pacific", "India": "Asia Pacific", "China": "Asia Pacific",
    "Hong Kong": "Asia Pacific", "Taiwan": "Asia Pacific", "Malaysia": "Asia Pacific",
    "Indonesia": "Asia Pacific", "Thailand": "Asia Pacific", "Philippines": "Asia Pacific",
    "New Zealand": "Asia Pacific", "Vietnam": "Asia Pacific",
    "Saudi Arabia": "Middle East", "UAE": "Middle East", "Qatar": "Middle East",
    "Bahrain": "Middle East", "Kuwait": "Middle East", "Oman": "Middle East",
    "Israel": "Middle East", "Jordan": "Middle East", "Egypt": "Middle East",
    "Brazil": "Latin America", "Chile": "Latin America", "Colombia": "Latin America",
    "Argentina": "Latin America", "Peru": "Latin America", "Mexico": "Latin America",
    "South Africa": "Africa", "Nigeria": "Africa", "Kenya": "Africa",
    "Ethiopia": "Africa", "Ghana": "Africa", "Morocco": "Africa",
    "Tanzania": "Africa", "Rwanda": "Africa",
}

COUNTRY_ISO = {
    "United States": "USA", "Canada": "CAN", "Mexico": "MEX",
    "United Kingdom": "GBR", "Germany": "DEU", "France": "FRA",
    "Netherlands": "NLD", "Ireland": "IRL", "Sweden": "SWE",
    "Norway": "NOR", "Denmark": "DNK", "Finland": "FIN",
    "Spain": "ESP", "Italy": "ITA", "Poland": "POL",
    "Switzerland": "CHE", "Austria": "AUT", "Belgium": "BEL",
    "Portugal": "PRT", "Romania": "ROU", "Czech Republic": "CZE",
    "Singapore": "SGP", "Japan": "JPN", "South Korea": "KOR",
    "Australia": "AUS", "India": "IND", "China": "CHN",
    "Hong Kong": "HKG", "Taiwan": "TWN", "Malaysia": "MYS",
    "Indonesia": "IDN", "Thailand": "THA", "Philippines": "PHL",
    "New Zealand": "NZL", "Vietnam": "VNM",
    "Saudi Arabia": "SAU", "UAE": "ARE", "Qatar": "QAT",
    "Bahrain": "BHR", "Kuwait": "KWT", "Oman": "OMN",
    "Israel": "ISR", "Jordan": "JOR", "Egypt": "EGY",
    "Brazil": "BRA", "Chile": "CHL", "Colombia": "COL",
    "Argentina": "ARG", "Peru": "PER",
    "South Africa": "ZAF", "Nigeria": "NGA", "Kenya": "KEN",
    "Ethiopia": "ETH", "Ghana": "GHA", "Morocco": "MAR",
    "Tanzania": "TZA", "Rwanda": "RWA",
}

COUNTRY_KEYWORDS = {k: [k] for k in COUNTRY_TO_REGION}
COUNTRY_KEYWORDS.update({
    "United States": ["United States", "U.S.", r"\bUS\b", "America", "American",
                      "Virginia", "Texas", "California", "Georgia", "Ohio",
                      "Indiana", "Nevada", "Arizona", "Illinois", "Pennsylvania",
                      "North Carolina", "Florida", "Oregon", "Washington",
                      "Colorado", "Utah", "Wyoming", "Montana", "Idaho",
                      "Minnesota", "Wisconsin", "Michigan", "Iowa", "Kansas",
                      "Missouri", "Oklahoma", "Arkansas", "Louisiana",
                      "Mississippi", "Alabama", "Tennessee", "Kentucky",
                      "Maryland", "New Jersey", "New York", "Massachusetts",
                      "Connecticut", "Maine", "New Mexico", "Alaska", "Hawaii",
                      "Dallas", "Austin", "Chicago", "Phoenix", "Atlanta",
                      "Seattle", "Denver", "Reno", "Ashburn", "Portland",
                      "Las Vegas", "San Jose", "San Francisco", "Los Angeles",
                      "Houston", "Miami", "Boston", "Columbus", "Nashville",
                      "Charlotte", "Northern Virginia", "Silicon Valley"],
    "United Kingdom": ["United Kingdom", r"\bUK\b", "England", "Scotland",
                       "Wales", "London", "Manchester", "Birmingham",
                       "Bristol", "Edinburgh", "Glasgow", "Slough", "Reading"],
    "Germany": ["Germany", "German", "Berlin", "Frankfurt", "Munich",
                "Hamburg", "Cologne", "Stuttgart", "Dusseldorf"],
    "Netherlands": ["Netherlands", "Dutch", "Amsterdam", "Rotterdam",
                    "Eindhoven", "The Hague", r"\bAMS\b"],
    "Singapore": ["Singapore", r"\bSGP\b"],
    "Australia": ["Australia", "Australian", "Sydney", "Melbourne",
                  "Brisbane", "Perth", "Canberra"],
    "India": ["India", "Indian", "Mumbai", "Bangalore", "Bengaluru",
              "Chennai", "Hyderabad", "Delhi", "Pune", "Noida"],
    "Japan": ["Japan", "Japanese", "Tokyo", "Osaka", "Nagoya"],
    "Saudi Arabia": ["Saudi Arabia", "Saudi", "KSA", "Riyadh", "Jeddah",
                     "Neom", "NEOM", "Vision 2030"],
    "UAE": ["UAE", "United Arab Emirates", "Dubai", "Abu Dhabi",
            "Sharjah", "Ajman"],
    "Brazil": ["Brazil", "Brasil", "Brazilian", r"\bSP\b", "Sao Paulo",
               "Rio de Janeiro", "Curitiba"],
    "South Africa": ["South Africa", "Johannesburg", "Cape Town", "Durban"],
    "Nigeria": ["Nigeria", "Nigerian", "Lagos", "Abuja"],
    "Kenya": ["Kenya", "Nairobi"],
    "France": ["France", "French", "Paris", "Lyon", "Marseille", "Toulouse"],
    "Ireland": ["Ireland", "Irish", "Dublin", "Cork"],
    "Sweden": ["Sweden", "Swedish", "Stockholm", "Gothenburg"],
    "Norway": ["Norway", "Norwegian", "Oslo"],
    "Denmark": ["Denmark", "Danish", "Copenhagen"],
    "Finland": ["Finland", "Finnish", "Helsinki"],
    "Poland": ["Poland", "Polish", "Warsaw", "Krakow"],
    "Malaysia": ["Malaysia", "Malaysian", "Kuala Lumpur", "KL", "Johor"],
    "Indonesia": ["Indonesia", "Indonesian", "Jakarta", "Batam"],
    "South Korea": ["South Korea", "Korean", "Seoul", "Busan"],
    "China": ["China", "Chinese", "Beijing", "Shanghai", "Shenzhen",
              "Guangzhou", "Chengdu"],
    "Hong Kong": ["Hong Kong", "HKG"],
    "Egypt": ["Egypt", "Egyptian", "Cairo"],
    "Qatar": ["Qatar", "Qatari", "Doha"],
    "Chile": ["Chile", "Chilean", "Santiago"],
    "Colombia": ["Colombia", "Colombian", "Bogota"],
    "Mexico": ["Mexico", "Mexican", "Mexico City", "Monterrey", "Guadalajara"],
})

TOPIC_KEYWORDS = {
    "Hyperscale": ["hyperscale","microsoft","google","amazon","aws","meta",
                   "apple","oracle","alibaba","tencent","bytedance",
                   "alphabet","openai","stargate","coreweave"],
    "Colocation": ["colocation","colo","equinix","digital realty","ironmountain",
                   "coresite","cyrusone","ntt","vantage","switch","edgeconex",
                   "flexential","databank","cts","global switch"],
    "AI / GPU":   ["artificial intelligence"," ai ","gpu","nvidia","inference",
                   "llm","generative","machine learning","deep learning",
                   "h100","h200","gb200","xai","anthropic","grok"],
    "Power":      [" mw "," gw ","megawatt","gigawatt","power plant",
                   "nuclear","solar","wind farm","gas turbine","behind-the-meter",
                   "grid","utility","energy","ppa","renewable","hydrogen",
                   "fuel cell","battery storage","ups"],
    "Investment": ["invest","fund","reit","billion","million","acquire",
                   "acquisition","ipo","bond","financing","lease","deal",
                   "partnership","joint venture","stake","raise","capital"],
    "Permits":    ["permit","zoning","moratorium","approved","approval",
                   "planning","ordinance","rezoning","denied","appeal",
                   "lawsuit","sue","court","ordinance","commission",
                   "vote","hearing","rejected"],
    "Construction":["broke ground","groundbreaking","topping out","opens",
                    "opened","inaugurated","construction","campus","build",
                    "development","site","facility","phase","expansion",
                    "warehouse","acres","sq ft","square feet"],
    "Sustainability":["sustainability","carbon","renewable","net zero",
                      "green","esg","water","pue","cooling","waste heat",
                      "recycle","circular","biodiversity","solar panel",
                      "wind power","offset"],
}

KNOWN_COMPANIES = [
    "Microsoft","Google","Amazon","AWS","Meta","Apple","Oracle","Alibaba",
    "Tencent","ByteDance","Baidu","Huawei","Samsung","IBM","Intel","NVIDIA",
    "Equinix","Digital Realty","Iron Mountain","CoreSite","CyrusOne","NTT",
    "Vantage","Switch","EdgeConneX","Flexential","DataBank","QTS","Colt",
    "Global Switch","ChinaData","GDS","STACK","Aligned","Compass","Tract",
    "Hut 8","Core Scientific","Riot","Marathon","Applied Digital","Iren",
    "Cloudflare","Fastly","Akamai","Lumen","Zayo","Cogent",
    "Schneider Electric","Vertiv","Eaton","ABB","Siemens","Caterpillar",
    "Cummins","Aggreko","AECOM","Turner","Holder","DPR","Skanska",
    "Prologis","Blackstone","KKR","Brookfield","GIC","Mubadala","CPPIB",
    "SoftBank","Masdar","TAQA","ADNOC","Saudi Aramco","Neom","SABIC",
    "CloudHQ","Coatue","DataBank","T5","Skybox","Stream","Flexential",
    "Ascenty","Odata","Scala","Luminet","Etix","Nabiax","Bolder",
    "VIRTUS","Kao","Yondr","Venari","Verne","Hydro66","DigiPlex",
    "Bulk Infrastructure","Green Mountain","atNorth","Adaniconnex",
    "CtrlS","NxtGen","Yotta","STT GDC","Keppel","Singtel","Telstra",
    "NextDC","Macquarie","AirTrunk","MEVSPACE","Beyond.pl","Atman",
    "DE-CIX","Interxion","euNetworks","Telehouse","Iomart","Pulsant",
]

GNEWS_QUERIES = [
    ("data center construction campus groundbreaking", "Google News"),
    ("data center hyperscale investment billion megawatt", "Google News"),
    ("data center approved permit moratorium zoning", "Google News"),
    ("data center power energy grid nuclear solar", "Google News"),
    ("colocation datacenter AI GPU facility opens", "Google News"),
]

RSS_SOURCES = [
    {
        "name": "DataCenter Knowledge",
        "url": "https://www.datacenterknowledge.com/rss.xml",
        "type": "rss",
    },
    {
        "name": "DataCenterFrontier",
        "url": "https://datacenterfrontier.com/feed/",
        "type": "rss",
    },
    {
        "name": "PR Newswire",
        "url": "https://www.prnewswire.com/rss/news-releases-list.rss",
        "type": "rss",
    },
    {
        "name": "BusinessWire",
        "url": "https://feed.businesswire.com/rss/home/?rss=G22",
        "type": "rss",
    },
    {
        "name": "Reuters",
        "url": "https://feeds.reuters.com/reuters/technologyNews",
        "type": "rss",
    },
]

SCRAPE_SOURCES = [
    {
        "name": "DataCenterDynamics",
        "url": "https://www.datacenterdynamics.com/en/news/?term=the-data-center-construction-channel",
        "base": "https://www.datacenterdynamics.com",
        "link_pattern": r"^/en/news/[^?#]+/$",
        "type": "html",
    },
    {
        "name": "DataCenterDynamics",
        "url": "https://www.datacenterdynamics.com/en/news/?term=north-america",
        "base": "https://www.datacenterdynamics.com",
        "link_pattern": r"^/en/news/[^?#]+/$",
        "type": "html",
    },
    {
        "name": "DataCenterDynamics",
        "url": "https://www.datacenterdynamics.com/en/news/?term=europe",
        "base": "https://www.datacenterdynamics.com",
        "link_pattern": r"^/en/news/[^?#]+/$",
        "type": "html",
    },
    {
        "name": "DataCenterDynamics",
        "url": "https://www.datacenterdynamics.com/en/news/?term=asia-pacific",
        "base": "https://www.datacenterdynamics.com",
        "link_pattern": r"^/en/news/[^?#]+/$",
        "type": "html",
    },
    {
        "name": "DataCenterDynamics",
        "url": "https://www.datacenterdynamics.com/en/news/?term=middle-east",
        "base": "https://www.datacenterdynamics.com",
        "link_pattern": r"^/en/news/[^?#]+/$",
        "type": "html",
    },
]


def parse_date_str(raw):
    if not raw:
        return None
    raw = str(raw).strip()
    raw = re.sub(r"\s+", " ", raw)
    for fmt in ("%Y-%m-%dT%H:%M:%S%z", "%Y-%m-%dT%H:%M:%SZ",
                "%a, %d %b %Y %H:%M:%S %z", "%a, %d %b %Y %H:%M:%S %Z",
                "%Y-%m-%d", "%d %b %Y", "%B %d, %Y", "%b %d, %Y"):
        try:
            dt = datetime.strptime(raw[:25], fmt[:len(raw[:25])])
            return dt.replace(tzinfo=None)
        except Exception:
            pass
    m = re.search(r"(\d{1,2})\s+(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)\w*\s+(\d{4})", raw, re.I)
    if m:
        try:
            return datetime(int(m.group(3)), MONTHS[m.group(2).lower()[:3]], int(m.group(1)))
        except Exception:
            pass
    m = re.search(r"(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)\w*\s+(\d{1,2}),?\s+(\d{4})", raw, re.I)
    if m:
        try:
            return datetime(int(m.group(3)), MONTHS[m.group(1).lower()[:3]], int(m.group(2)))
        except Exception:
            pass
    return None


def fetch_html(url, retries=2):
    for attempt in range(retries):
        try:
            if _USE_CS:
                r = _CS.get(url, timeout=18)
            else:
                import requests as req
                r = req.get(url, headers=_HEADERS, timeout=18)
            r.raise_for_status()
            return BeautifulSoup(r.text, "html.parser")
        except Exception:
            if attempt == 0:
                time.sleep(1.5)
    return None


def scrape_html_source(source, max_pages=3):
    results = []
    seen = set()
    base = source["base"]
    pattern = re.compile(source["link_pattern"])

    for page in range(1, max_pages + 1):
        url = source["url"] if page == 1 else source["url"] + f"&page={page}"
        soup = fetch_html(url)
        if not soup:
            break
        found_any = False
        for a in soup.find_all("a", href=pattern):
            href = a["href"]
            if href in seen:
                continue
            seen.add(href)
            h = a.find(["h1", "h2", "h3", "h4"])
            headline = h.get_text(strip=True) if h else a.get_text(strip=True)
            if not headline or len(headline) < 12:
                continue
            date_obj = None
            node = a.parent
            for _ in range(9):
                if node is None:
                    break
                txt = node.get_text(" ", strip=True)
                m = re.search(
                    r"\b(\d{1,2})\s+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+(\d{4})\b",
                    txt, re.I
                )
                if m:
                    date_obj = parse_date_str(m.group(0))
                    break
                node = node.parent
            results.append({
                "headline": headline,
                "url": base + href,
                "date_obj": date_obj,
                "source": source["name"],
            })
            found_any = True
        if not found_any:
            break
        time.sleep(0.4)
    return results


def fetch_rss(source):
    results = []
    try:
        feed = feedparser.parse(source["url"])
        for entry in feed.entries:
            headline = entry.get("title", "").strip()
            url = entry.get("link", "").strip()
            if not headline or not url:
                continue
            pub = entry.get("published", "") or entry.get("updated", "")
            date_obj = None
            if hasattr(entry, "published_parsed") and entry.published_parsed:
                try:
                    date_obj = datetime(*entry.published_parsed[:6])
                except Exception:
                    pass
            if date_obj is None:
                date_obj = parse_date_str(pub)
            results.append({
                "headline": headline,
                "url": url,
                "date_obj": date_obj,
                "source": source["name"],
            })
    except Exception:
        pass
    return results


def fetch_google_news(query, source_label="Google News"):
    results = []
    try:
        q = query.replace(" ", "+")
        url = f"https://news.google.com/rss/search?q={q}&hl=en-US&gl=US&ceid=US:en"
        feed = feedparser.parse(url)
        for entry in feed.entries:
            headline = entry.get("title", "").strip()
            url_val = entry.get("link", "").strip()
            if not headline or not url_val:
                continue
            date_obj = None
            if hasattr(entry, "published_parsed") and entry.published_parsed:
                try:
                    date_obj = datetime(*entry.published_parsed[:6])
                except Exception:
                    pass
            results.append({
                "headline": headline,
                "url": url_val,
                "date_obj": date_obj,
                "source": source_label,
            })
    except Exception:
        pass
    return results


def is_dc_relevant(text):
    t = text.lower()
    primary = ["data center", "datacenter", "data centre", "datacentre",
               "colocation", "hyperscale", "cloud campus"]
    secondary = ["megawatt", " mw ", "gigawatt", " gw ", "server farm",
                 "computing facility", "edge computing", "ai campus",
                 "gpu cluster", "compute campus"]
    if any(p in t for p in primary):
        return True
    if sum(1 for s in secondary if s in t) >= 2:
        return True
    return False


def detect_country(text):
    for country, patterns in COUNTRY_KEYWORDS.items():
        for pat in patterns:
            if re.search(pat if pat.startswith(r"\b") else r"\b" + re.escape(pat) + r"\b", text, re.I):
                return country
    return "Global"


def detect_topic(text):
    t = text.lower()
    for topic, kws in TOPIC_KEYWORDS.items():
        if any(k.lower() in t for k in kws):
            return topic
    return "General"


def detect_mw(text):
    m = re.search(r"([\d,]+(?:\.\d+)?)\s*(GW|MW|gigawatt|megawatt)", text, re.I)
    if m:
        return m.group(1).replace(",", "") + " " + m.group(2).upper()
    return ""


def detect_deal_size(text):
    m = re.search(r"\$([\d,.]+)\s*(billion|bn|million|mn|m\b)", text, re.I)
    if m:
        val = m.group(1).replace(",", "")
        unit = m.group(2).lower()
        if unit in ("billion", "bn"):
            return f"${val}bn"
        return f"${val}m"
    return ""


def detect_companies(text):
    found = []
    for co in KNOWN_COMPANIES:
        if re.search(r"\b" + re.escape(co) + r"\b", text, re.I):
            found.append(co)
    return ", ".join(found[:4]) if found else ""


def detect_sentiment(text):
    t = text.lower()
    if any(w in t for w in ["broke ground", "groundbreaking", "opens", "opened",
                             "inaugurated", "energizes", "goes live", "launches"]):
        return "Opened / Live"
    if any(w in t for w in ["approved", "approval", "go-ahead", "green light",
                             "permits", "zoning approved", "rezoning"]):
        return "Approved"
    if any(w in t for w in ["proposed", "plans", "eyes", "looks to", "could build",
                             "may build", "files for", "announces plans"]):
        return "Proposed"
    if any(w in t for w in ["rejected", "denied", "moratorium", "blocked",
                             "lawsuit", "sues", "opposition", "withdrawn"]):
        return "Challenged"
    if any(w in t for w in ["under construction", "construction begins",
                             "construction started", "building"]):
        return "Under Construction"
    return "News"


def fuzzy_similar(a, b, threshold=0.82):
    return SequenceMatcher(None, a.lower(), b.lower()).ratio() >= threshold


def deduplicate(articles):
    keep = []
    seen_headlines = []
    for art in articles:
        hl = art["Headline"]
        is_dup = False
        for seen in seen_headlines:
            if fuzzy_similar(hl, seen):
                is_dup = True
                break
        if not is_dup:
            keep.append(art)
            seen_headlines.append(hl)
    return keep


def enrich(raw_item):
    hl = raw_item["headline"]
    d = raw_item.get("date_obj")
    country = detect_country(hl)
    region = COUNTRY_TO_REGION.get(country, "Global")
    return {
        "Headline":  hl,
        "Date":      d.strftime("%Y-%m-%d") if d else "Unknown",
        "Source":    raw_item.get("source", "Unknown"),
        "URL":       raw_item.get("url", ""),
        "Country":   country,
        "Region":    region,
        "Topic":     detect_topic(hl),
        "Sentiment": detect_sentiment(hl),
        "Capacity":  detect_mw(hl),
        "Deal Size": detect_deal_size(hl),
        "Companies": detect_companies(hl),
        "_date_obj": d,
    }


def run_all_scrapers(max_html_pages, cutoff, progress_cb):
    raw = []
    total_tasks = len(SCRAPE_SOURCES) + len(RSS_SOURCES) + len(GNEWS_QUERIES)
    done = [0]

    def tick(label=""):
        done[0] += 1
        progress_cb(done[0] / total_tasks, label)

    with ThreadPoolExecutor(max_workers=8) as pool:
        html_futures = {
            pool.submit(scrape_html_source, src, max_html_pages): src["name"]
            for src in SCRAPE_SOURCES
        }
        rss_futures = {
            pool.submit(fetch_rss, src): src["name"]
            for src in RSS_SOURCES
        }
        gn_futures = {
            pool.submit(fetch_google_news, q, lbl): lbl
            for q, lbl in GNEWS_QUERIES
        }

        for f in as_completed({**html_futures, **rss_futures, **gn_futures}):
            try:
                items = f.result()
                raw.extend(items)
            except Exception:
                pass
            tick()

    filtered = []
    for item in raw:
        d = item.get("date_obj")
        if d and d < cutoff:
            continue
        if not is_dc_relevant(item["headline"]):
            continue
        filtered.append(item)

    return filtered


def build_excel(df):
    wb = Workbook()
    thin = Side(border_style="thin", color="CCCCCC")
    brd = Border(left=thin, right=thin, top=thin, bottom=thin)
    hf = Font(bold=True, color="FFFFFF", name="Calibri", size=10)
    hfill = PatternFill("solid", fgColor="0F1E36")
    ev = PatternFill("solid", fgColor="0B1628")
    od = PatternFill("solid", fgColor="060A10")
    nf = Font(name="Calibri", size=9, color="C8D4E8")
    lf = Font(name="Calibri", size=9, color="0047E1", underline="single")

    def write_sheet(ws, data_df, cols, widths):
        for ci, (col, w) in enumerate(zip(cols, widths), 1):
            c = ws.cell(1, ci, col)
            c.font = hf; c.fill = hfill
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = brd
            ws.column_dimensions[get_column_letter(ci)].width = w
        ws.row_dimensions[1].height = 24
        tc_map = TOPIC_COLORS
        for ri, row in enumerate(data_df[cols if all(c in data_df.columns for c in cols) else data_df.columns].itertuples(index=False), start=2):
            fill = ev if ri % 2 == 0 else od
            for ci, val in enumerate(row, 1):
                v = str(val) if val is not None else ""
                c = ws.cell(ri, ci, v)
                c.fill = fill; c.border = brd
                col_name = cols[ci - 1] if ci <= len(cols) else ""
                if col_name == "URL":
                    c.hyperlink = v; c.font = lf
                    c.alignment = Alignment(horizontal="center", vertical="center")
                elif col_name == "Topic":
                    tc = tc_map.get(v, "2E4470")
                    c.font = Font(name="Calibri", size=9, bold=True, color=tc.replace("#", ""))
                    c.alignment = Alignment(horizontal="center", vertical="center")
                elif col_name == "Headline":
                    c.font = nf
                    c.alignment = Alignment(vertical="center", wrap_text=True)
                else:
                    c.font = nf
                    c.alignment = Alignment(horizontal="center", vertical="center")
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = f"A1:{get_column_letter(len(cols))}{len(data_df)+1}"

    ws1 = wb.active
    ws1.title = "All Articles"
    main_cols = ["Headline","Date","Source","Country","Region","Topic",
                 "Sentiment","Capacity","Deal Size","Companies","URL"]
    main_widths = [62, 12, 18, 18, 16, 14, 16, 12, 12, 28, 50]
    safe_df = df[[c for c in main_cols if c in df.columns]]
    write_sheet(ws1, safe_df, [c for c in main_cols if c in df.columns], main_widths[:len(main_cols)])

    ws2 = wb.create_sheet("By Country")
    cc = df["Country"].value_counts().reset_index()
    cc.columns = ["Country", "Articles"]
    write_sheet(ws2, cc, ["Country","Articles"], [24, 14])

    ws3 = wb.create_sheet("By Region")
    rc = df["Region"].value_counts().reset_index()
    rc.columns = ["Region", "Articles"]
    write_sheet(ws3, rc, ["Region","Articles"], [20, 14])

    ws4 = wb.create_sheet("By Topic")
    tc_df = df["Topic"].value_counts().reset_index()
    tc_df.columns = ["Topic", "Articles"]
    write_sheet(ws4, tc_df, ["Topic","Articles"], [18, 14])

    ws5 = wb.create_sheet("By Company")
    comp_rows = []
    for _, row in df.iterrows():
        if row.get("Companies"):
            for co in str(row["Companies"]).split(", "):
                co = co.strip()
                if co:
                    comp_rows.append({"Company": co, "Headline": row["Headline"],
                                      "Date": row["Date"], "Country": row["Country"],
                                      "URL": row["URL"]})
    if comp_rows:
        comp_df = pd.DataFrame(comp_rows)
        write_sheet(ws5, comp_df, ["Company","Headline","Date","Country","URL"],
                    [22, 55, 12, 18, 45])

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


_BG = "#060a10"
_PAPER = "#0b1628"
_GRID = "#152038"
_TEXT = "#6a80a8"
_TITLE = "#b8c8e0"
_FONT = "Inter, sans-serif"


def _dark(fig, height=320):
    fig.update_layout(
        paper_bgcolor=_PAPER, plot_bgcolor=_BG,
        font=dict(family=_FONT, color=_TEXT),
        height=height,
        margin=dict(l=14, r=14, t=36, b=14),
        xaxis=dict(gridcolor=_GRID, linecolor=_GRID,
                   tickfont=dict(size=10, color=_TEXT),
                   title_font=dict(color=_TEXT)),
        yaxis=dict(gridcolor=_GRID, linecolor=_GRID,
                   tickfont=dict(size=10, color=_TEXT),
                   title_font=dict(color=_TEXT)),
        showlegend=False,
    )
    return fig


def chart_topic_bar(df):
    tc = df["Topic"].value_counts().reset_index()
    tc.columns = ["Topic", "Count"]
    tc = tc.sort_values("Count")
    colors = [TOPIC_COLORS.get(t, "#2e4470") for t in tc["Topic"]]
    fig = go.Figure(go.Bar(
        x=tc["Count"], y=tc["Topic"], orientation="h",
        marker=dict(color=colors, line=dict(width=0)),
        text=tc["Count"], textposition="outside",
        textfont=dict(color=_TITLE, size=11),
        hovertemplate="<b>%{y}</b>: %{x} articles<extra></extra>",
    ))
    _dark(fig, 300)
    fig.update_layout(title=dict(text="Articles by Topic", font=dict(color=_TITLE, size=13), x=0.01))
    return fig


def chart_region_bar(df):
    rc = df["Region"].value_counts().reset_index()
    rc.columns = ["Region", "Count"]
    rc = rc.sort_values("Count")
    colors = [REGION_COLORS.get(r, "#2e4470") for r in rc["Region"]]
    fig = go.Figure(go.Bar(
        x=rc["Count"], y=rc["Region"], orientation="h",
        marker=dict(color=colors, line=dict(width=0)),
        text=rc["Count"], textposition="outside",
        textfont=dict(color=_TITLE, size=11),
        hovertemplate="<b>%{y}</b>: %{x} articles<extra></extra>",
    ))
    _dark(fig, 280)
    fig.update_layout(title=dict(text="Articles by Region", font=dict(color=_TITLE, size=13), x=0.01))
    return fig


def chart_country_bar(df, top_n=20):
    cc = df["Country"].value_counts().head(top_n).reset_index()
    cc.columns = ["Country", "Count"]
    cc = cc.sort_values("Count")
    n = len(cc)
    c_colors = [
        f"rgba({int(0 + 71*i/max(n-1,1))}, {int(71 + (180-71)*i/max(n-1,1))}, {int(225 + (255-225)*i/max(n-1,1))}, 0.85)"
        for i in range(n)
    ]
    fig = go.Figure(go.Bar(
        x=cc["Count"], y=cc["Country"], orientation="h",
        marker=dict(color=c_colors, line=dict(width=0)),
        text=cc["Count"], textposition="outside",
        textfont=dict(color=_TITLE, size=10),
        hovertemplate="<b>%{y}</b>: %{x} articles<extra></extra>",
    ))
    _dark(fig, max(320, top_n * 22))
    fig.update_layout(title=dict(text=f"Top {top_n} Countries", font=dict(color=_TITLE, size=13), x=0.01))
    return fig


def chart_timeline(df):
    df2 = df[df["Date"] != "Unknown"].copy()
    if df2.empty:
        return None
    df2["dt"] = pd.to_datetime(df2["Date"])
    daily = df2.groupby(df2["dt"].dt.date).size().reset_index()
    daily.columns = ["Date", "Articles"]
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=daily["Date"], y=daily["Articles"],
        mode="lines+markers",
        line=dict(color="#00b4ff", width=2.5),
        marker=dict(color="#0047e1", size=5, line=dict(color="#00b4ff", width=1.5)),
        fill="tozeroy", fillcolor="rgba(0,71,225,0.07)",
        hovertemplate="<b>%{x}</b><br>%{y} articles<extra></extra>",
    ))
    _dark(fig, 240)
    fig.update_layout(
        title=dict(text="Publication Volume Over Time", font=dict(color=_TITLE, size=13), x=0.01),
        yaxis_title="Articles",
    )
    return fig


def chart_sentiment(df):
    sc = df["Sentiment"].value_counts().reset_index()
    sc.columns = ["Sentiment", "Count"]
    sent_colors = {
        "Opened / Live": "#00e676", "Approved": "#00b4ff",
        "Proposed": "#ffaa00", "Under Construction": "#00e5c8",
        "Challenged": "#ff2d6b", "News": "#2e4470",
    }
    colors = [sent_colors.get(s, "#2e4470") for s in sc["Sentiment"]]
    fig = go.Figure(go.Bar(
        x=sc["Sentiment"], y=sc["Count"],
        marker=dict(color=colors, line=dict(width=0)),
        text=sc["Count"], textposition="outside",
        textfont=dict(color=_TITLE, size=11),
        hovertemplate="<b>%{x}</b>: %{y}<extra></extra>",
    ))
    _dark(fig, 280)
    fig.update_layout(title=dict(text="Article Sentiment / Status", font=dict(color=_TITLE, size=13), x=0.01))
    return fig


def chart_donut(df):
    tc = df["Topic"].value_counts().reset_index()
    tc.columns = ["Topic", "Count"]
    colors = [TOPIC_COLORS.get(t, "#2e4470") for t in tc["Topic"]]
    fig = go.Figure(go.Pie(
        labels=tc["Topic"], values=tc["Count"], hole=0.55,
        marker=dict(colors=colors, line=dict(color=_BG, width=2)),
        textinfo="label+percent",
        textfont=dict(color=_TITLE, size=11),
        hovertemplate="<b>%{label}</b>: %{value} (%{percent})<extra></extra>",
    ))
    fig.update_layout(
        paper_bgcolor=_PAPER, plot_bgcolor=_BG,
        font=dict(family=_FONT, color=_TEXT),
        height=340, margin=dict(l=14, r=14, t=36, b=14),
        showlegend=True,
        legend=dict(bgcolor="rgba(0,0,0,0)", font=dict(color=_TITLE, size=10)),
        annotations=[dict(
            text=f"<b>{len(df)}</b><br><span style='font-size:10px'>articles</span>",
            x=0.5, y=0.5, showarrow=False,
            font=dict(size=15, color=_TITLE, family=_FONT),
        )],
        title=dict(text="Topic Share", font=dict(color=_TITLE, size=13), x=0.01),
    )
    return fig


def chart_source_bar(df):
    sc = df["Source"].value_counts().reset_index()
    sc.columns = ["Source", "Count"]
    sc = sc.sort_values("Count")
    colors = [SOURCE_META.get(s, SOURCE_META["Unknown"])["color"] for s in sc["Source"]]
    fig = go.Figure(go.Bar(
        x=sc["Count"], y=sc["Source"], orientation="h",
        marker=dict(color=colors, line=dict(width=0)),
        text=sc["Count"], textposition="outside",
        textfont=dict(color=_TITLE, size=11),
        hovertemplate="<b>%{y}</b>: %{x} articles<extra></extra>",
    ))
    _dark(fig, 280)
    fig.update_layout(title=dict(text="Articles by Source", font=dict(color=_TITLE, size=13), x=0.01))
    return fig


def chart_world_map(df):
    cc = df[df["Country"] != "Global"]["Country"].value_counts().reset_index()
    cc.columns = ["Country", "Count"]
    cc["ISO"] = cc["Country"].map(COUNTRY_ISO)
    cc = cc.dropna(subset=["ISO"])
    fig = go.Figure(go.Choropleth(
        locations=cc["ISO"],
        z=cc["Count"],
        text=cc["Country"],
        colorscale=[
            [0.0, "#0a1628"],
            [0.2, "#0e2040"],
            [0.4, "#0047e1"],
            [0.7, "#00b4ff"],
            [1.0, "#00e5c8"],
        ],
        autocolorscale=False,
        reversescale=False,
        marker=dict(line=dict(color="#0f1e36", width=0.6)),
        colorbar=dict(
            bgcolor=_PAPER, bordercolor=_GRID, borderwidth=1,
            tickfont=dict(color=_TITLE, size=10),
            title=dict(text="Articles", font=dict(color=_TITLE, size=11)),
            len=0.7, thickness=14,
        ),
        hovertemplate="<b>%{text}</b><br>Articles: %{z}<extra></extra>",
        showscale=True,
        zmin=0,
    ))
    fig.update_geos(
        bgcolor=_BG,
        landcolor="#0d1a2e",
        oceancolor="#060a10",
        lakecolor="#060a10",
        rivercolor="#060a10",
        framecolor=_GRID,
        showland=True, showocean=True, showlakes=True,
        showcountries=True, countrycolor="#152038",
        showframe=True,
        projection_type="natural earth",
    )
    fig.update_layout(
        paper_bgcolor=_BG, plot_bgcolor=_BG,
        height=480,
        margin=dict(l=0, r=0, t=32, b=0),
        title=dict(
            text="Global Data Center Activity",
            font=dict(color=_TITLE, size=14, family=_FONT), x=0.01,
        ),
        geo=dict(bgcolor=_BG),
    )
    return fig


def dark_table(df_in, max_rows=300):
    th = (
        "background:#0f1e36;color:#b8c8e0;font-family:monospace;"
        "font-size:.68rem;letter-spacing:.08em;text-transform:uppercase;"
        "padding:.55rem .85rem;border-bottom:2px solid #0047e1;"
        "white-space:nowrap;text-align:left;"
    )
    td = (
        "padding:.5rem .85rem;font-size:.8rem;color:#b8c8e0;"
        "border-bottom:1px solid #101b2e;font-family:Inter,sans-serif;"
        "vertical-align:middle;"
    )
    rows = ""
    for i, (_, row) in enumerate(df_in.head(max_rows).iterrows()):
        bg = "#0b1628" if i % 2 == 0 else "#060a10"
        cells = ""
        for col in df_in.columns:
            v = str(row[col]) if row[col] is not None else ""
            if col == "URL" or (col != "Headline" and v.startswith("http")):
                cells += (
                    f'<td style="{td}background:{bg};">'
                    f'<a href="{v}" target="_blank" '
                    f'style="color:#0047e1;text-decoration:none;font-size:.75rem;">'
                    f'Open \u2192</a></td>'
                )
            elif col == "Capacity" and v:
                cells += (
                    f'<td style="{td}background:{bg};">'
                    f'<span style="color:#ffaa00;font-family:monospace;font-size:.76rem;">'
                    f'\u26a1 {v}</span></td>'
                )
            elif col == "Deal Size" and v:
                cells += (
                    f'<td style="{td}background:{bg};">'
                    f'<span style="color:#00e676;font-family:monospace;font-size:.76rem;">'
                    f'{v}</span></td>'
                )
            elif col == "Topic":
                tc = TOPIC_COLORS.get(v, "#2e4470")
                cells += (
                    f'<td style="{td}background:{bg};">'
                    f'<span style="background:{tc}22;color:{tc};border:1px solid {tc}44;'
                    f'border-radius:4px;padding:2px 7px;font-size:.68rem;font-family:monospace;">'
                    f'{v}</span></td>'
                )
            elif col == "Region":
                rc2 = REGION_COLORS.get(v, "#2e4470")
                cells += (
                    f'<td style="{td}background:{bg};">'
                    f'<span style="background:{rc2}22;color:{rc2};border:1px solid {rc2}44;'
                    f'border-radius:4px;padding:2px 7px;font-size:.68rem;font-family:monospace;">'
                    f'{v}</span></td>'
                )
            elif col == "Sentiment":
                sent_c = {
                    "Opened / Live":"#00e676","Approved":"#00b4ff","Proposed":"#ffaa00",
                    "Under Construction":"#00e5c8","Challenged":"#ff2d6b","News":"#2e4470",
                }.get(v, "#2e4470")
                cells += (
                    f'<td style="{td}background:{bg};">'
                    f'<span style="background:{sent_c}18;color:{sent_c};'
                    f'border:1px solid {sent_c}44;border-radius:4px;'
                    f'padding:2px 7px;font-size:.68rem;font-family:monospace;">'
                    f'{v}</span></td>'
                )
            elif col == "Source":
                sc2 = SOURCE_META.get(v, SOURCE_META["Unknown"])
                sc2_color = sc2["color"]
                sc2_short = sc2["short"]
                cells += (
                    f'<td style="{td}background:{bg};">'
                    f'<span style="background:{sc2_color}22;color:{sc2_color};'
                    f'border:1px solid {sc2_color}44;border-radius:4px;'
                    f'padding:2px 6px;font-size:.65rem;font-family:monospace;">'
                    f'{sc2_short}</span></td>'
                )
            elif col in ("Headline", "Companies"):
                cells += (
                    f'<td style="{td}background:{bg};max-width:480px;'
                    f'word-wrap:break-word;white-space:normal;">{v}</td>'
                )
            else:
                cells += f'<td style="{td}background:{bg};">{v}</td>'
        rows += f"<tr>{cells}</tr>"

    heads = "".join(f'<th style="{th}">{c}</th>' for c in df_in.columns)
    return (
        '<div style="overflow-x:auto;border-radius:10px;border:1px solid #152038;margin-bottom:1rem;">'
        f'<table style="width:100%;border-collapse:collapse;background:#060a10;">'
        f'<thead><tr>{heads}</tr></thead>'
        f'<tbody>{rows}</tbody></table></div>'
    )


def article_card(headline, date, url, source, country, topic, capacity, deal, sentiment):
    tc = TOPIC_COLORS.get(topic, "#2e4470")
    sc_meta = SOURCE_META.get(source, SOURCE_META["Unknown"])
    cap_html = (
        f'<span style="background:rgba(255,170,0,0.12);color:#ffaa00;'
        f'border:1px solid rgba(255,170,0,0.3);border-radius:4px;'
        f'padding:2px 6px;font-family:monospace;font-size:.62rem;white-space:nowrap;">'
        f'\u26a1 {capacity}</span>'
    ) if capacity else ""
    deal_html = (
        f'<span style="background:rgba(0,230,118,0.1);color:#00e676;'
        f'border:1px solid rgba(0,230,118,0.25);border-radius:4px;'
        f'padding:2px 6px;font-family:monospace;font-size:.62rem;white-space:nowrap;">'
        f'{deal}</span>'
    ) if deal else ""
    sent_c = {
        "Opened / Live":"#00e676","Approved":"#00b4ff","Proposed":"#ffaa00",
        "Under Construction":"#00e5c8","Challenged":"#ff2d6b","News":"#2e4470",
    }.get(sentiment, "#2e4470")
    arrow = "\u2197"
    return (
        f'<div style="background:#0b1628;border:1px solid #152038;border-radius:10px;'
        f'padding:.85rem 1.1rem;display:flex;justify-content:space-between;'
        f'align-items:flex-start;gap:.9rem;margin-bottom:.45rem;">'
        f'<div style="flex:1;min-width:0;">'
        f'<a href="{url}" target="_blank" '
        f'style="color:#ccdaf5;text-decoration:none;font-family:Inter,sans-serif;'
        f'font-size:.86rem;font-weight:500;line-height:1.5;">{headline}</a>'
        f'<div style="margin-top:.35rem;display:flex;gap:.4rem;flex-wrap:wrap;">'
        f'<span style="font-family:monospace;font-size:.62rem;color:#2a3e60;">\U0001f4c5 {date}</span>'
        f'<span style="font-family:monospace;font-size:.62rem;color:#4a6490;">\U0001f30d {country}</span>'
        f'</div></div>'
        f'<div style="display:flex;flex-direction:column;align-items:flex-end;'
        f'gap:.28rem;flex-shrink:0;white-space:nowrap;">'
        f'<span style="background:{sc_meta["color"]}22;color:{sc_meta["color"]};'
        f'border:1px solid {sc_meta["color"]}44;border-radius:4px;'
        f'padding:2px 6px;font-family:monospace;font-size:.62rem;">{sc_meta["short"]}</span>'
        f'<span style="background:{tc}22;color:{tc};border:1px solid {tc}44;'
        f'border-radius:4px;padding:2px 6px;font-family:monospace;font-size:.62rem;">{topic}</span>'
        f'<span style="background:{sent_c}18;color:{sent_c};border:1px solid {sent_c}44;'
        f'border-radius:4px;padding:2px 6px;font-family:monospace;font-size:.62rem;">{sentiment}</span>'
        f'{cap_html}{deal_html}'
        f'<a href="{url}" target="_blank" '
        f'style="font-family:monospace;font-size:.65rem;color:#0047e1;text-decoration:none;">'
        f'{arrow} open</a>'
        f'</div></div>'
    )


def kpi(label, value, accent="blue", delta=""):
    accent_map = {
        "blue":  ("#0047e1", "#00b4ff"),
        "cyan":  ("#00b4ff", "#00e5c8"),
        "green": ("#00e676", "#00c853"),
        "amber": ("#ffaa00", "#ff6400"),
        "purple":("#a855f7", "#7c3aed"),
        "red":   ("#ff2d6b", "#ff0044"),
    }
    c1, c2 = accent_map.get(accent, accent_map["blue"])
    delta_html = f'<div style="font-size:.7rem;color:#2a3e60;margin-top:.25rem;">{delta}</div>' if delta else ""
    return (
        f'<div style="flex:1;min-width:150px;background:#0b1628;border:1px solid #152038;'
        f'border-radius:12px;padding:1.1rem 1.3rem;position:relative;overflow:hidden;">'
        f'<div style="position:absolute;top:0;left:0;right:0;height:2px;'
        f'background:linear-gradient(90deg,{c1},{c2});"></div>'
        f'<div style="font-family:monospace;font-size:.64rem;letter-spacing:.13em;'
        f'text-transform:uppercase;color:#2a3e60;margin-bottom:.4rem;">{label}</div>'
        f'<div style="font-family:Syne,sans-serif;font-size:1.9rem;font-weight:800;'
        f'color:#fff;line-height:1;">{value}</div>'
        f'{delta_html}</div>'
    )


def main():
    st.set_page_config(
        page_title="Global DC Intel",
        page_icon="\U0001f310",
        layout="wide",
        initial_sidebar_state="expanded",
    )
    st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

    with st.sidebar:
        st.markdown(
            '<div style="padding:.9rem 0 .4rem;text-align:center;">'
            '<div style="font-family:monospace;font-size:.6rem;letter-spacing:.2em;'
            'color:#1a2e50;text-transform:uppercase;margin-bottom:.25rem;">Intelligence Platform</div>'
            '<div style="font-family:Syne,sans-serif;font-size:1.15rem;font-weight:800;color:#fff;">DC Intel</div>'
            '<div style="font-family:monospace;font-size:.6rem;color:#1a2e50;margin-top:.15rem;">'
            'Global Construction Monitor</div></div>',
            unsafe_allow_html=True,
        )
        st.divider()

        st.markdown("**\U0001f4c5 Date Range**")
        time_opt = st.radio(
            "", ["Latest (all)", "Past 30 days", "Past 14 days", "Past 7 days", "Custom Range"],
            index=0, label_visibility="collapsed",
        )
        days_map = {"Latest (all)": None, "Past 30 days": 30, "Past 14 days": 14, "Past 7 days": 7}
        sel_days = days_map.get(time_opt, None)

        custom_start = None
        custom_end   = None
        if time_opt == "Custom Range":
            today = datetime.now().date()
            c1, c2 = st.columns(2)
            with c1:
                custom_start = st.date_input(
                    "From", value=today - timedelta(days=30),
                    max_value=today, label_visibility="visible",
                )
            with c2:
                custom_end = st.date_input(
                    "To", value=today,
                    max_value=today, label_visibility="visible",
                )
            if custom_start and custom_end and custom_start > custom_end:
                st.error("Start date must be before end date.")
                custom_start, custom_end = custom_end, custom_start
            st.markdown(
                f'<div style="font-size:.68rem;color:#1a2e50;margin-top:-.3rem;margin-bottom:.4rem;">'
                f'{custom_start.strftime("%d %b %Y") if custom_start else ""}'
                f' → '
                f'{custom_end.strftime("%d %b %Y") if custom_end else ""}'
                f'</div>',
                unsafe_allow_html=True,
            )

        st.markdown("**\U0001f4c4 HTML Scrape Depth**")
        max_pages = st.slider("", 1, 10, 3, label_visibility="collapsed",
                              help="Pages per HTML source (~28 articles/page)")
        st.markdown(
            f'<div style="font-size:.7rem;color:#1a2e50;margin-top:-.4rem;margin-bottom:.6rem;">'
            f'~{max_pages * 28 * len(SCRAPE_SOURCES)} max HTML articles</div>',
            unsafe_allow_html=True,
        )

        st.markdown("**\U0001f4e1 Sources**")
        use_html = st.checkbox("HTML Scrapers (DCD regions)", value=True)
        use_rss  = st.checkbox("RSS Feeds (DCK, DCF, PRN, BIZ, Reuters)", value=True)
        use_gn   = st.checkbox("Google News (5 queries)", value=True)

        st.divider()

        if "df_full" in st.session_state and st.session_state.df_full is not None:
            df_full = st.session_state.df_full
            st.markdown("**\U0001f50d Refine Results**")
            all_regions  = sorted(df_full["Region"].unique().tolist())
            all_topics   = sorted(df_full["Topic"].unique().tolist())
            all_sources  = sorted(df_full["Source"].unique().tolist())
            all_sents    = sorted(df_full["Sentiment"].unique().tolist())
            sel_regions  = st.multiselect("Regions", all_regions, default=all_regions)
            sel_topics   = st.multiselect("Topics", all_topics, default=all_topics)
            sel_sources  = st.multiselect("Sources", all_sources, default=all_sources)
            sel_sents    = st.multiselect("Sentiment", all_sents, default=all_sents)
            keyword      = st.text_input("Keyword", placeholder="Microsoft, 500MW, Texas...")
            min_mw       = st.number_input("Min capacity (MW)", min_value=0, value=0, step=10)

            st.markdown("**\U0001f4c6 Filter by Published Date**")
            valid_dates = df_full[df_full["Date"] != "Unknown"]["Date"]
            if not valid_dates.empty:
                min_date = pd.to_datetime(valid_dates).min().date()
                max_date = pd.to_datetime(valid_dates).max().date()
                fc1, fc2 = st.columns(2)
                with fc1:
                    filt_from = st.date_input(
                        "From", value=min_date,
                        min_value=min_date, max_value=max_date,
                        key="filt_from",
                    )
                with fc2:
                    filt_to = st.date_input(
                        "To", value=max_date,
                        min_value=min_date, max_value=max_date,
                        key="filt_to",
                    )
                use_date_filter = st.checkbox("Apply date filter", value=False)
            else:
                filt_from = filt_to = None
                use_date_filter = False

            st.markdown("**\U0001f30d Country Filter**")
            all_countries = sorted(df_full["Country"].unique().tolist())
            country_search = st.text_input(
                "Search country", placeholder="e.g. United States, India...",
                key="country_search_input",
            )
            if country_search.strip():
                matched = [c for c in all_countries if country_search.strip().lower() in c.lower()]
            else:
                matched = all_countries
            sel_countries = st.multiselect(
                "Select countries", matched, default=matched,
                help="Showing countries matching your search above",
            )

            st.markdown("**\U0001f3e2 Company Search**")
            company_search = st.text_input(
                "Search company", placeholder="e.g. Microsoft, Equinix, AWS...",
                key="company_search_input",
            )

            st.session_state.filters = {
                "regions": sel_regions, "topics": sel_topics,
                "sources": sel_sources, "sents": sel_sents,
                "keyword": keyword, "min_mw": min_mw,
                "date_from": filt_from if use_date_filter else None,
                "date_to":   filt_to   if use_date_filter else None,
                "countries": sel_countries,
                "company_search": company_search.strip(),
            }

        st.divider()
        go_btn = st.button("\U0001f50d  Run Global Scan", use_container_width=True, type="primary")

    now_str = datetime.now().strftime("%A, %d %B %Y  \u00b7  %H:%M UTC")
    st.markdown(
        f'<div class="gl-banner">'
        f'<div class="banner-eyebrow">\u25cf Live Intelligence Feed  \u00b7  {len(SCRAPE_SOURCES) + len(RSS_SOURCES) + len(GNEWS_QUERIES)} Sources Active</div>'
        f'<div class="banner-title">Global Data Center <span>Construction</span> Intelligence</div>'
        f'<div class="banner-sub">Real-time aggregation across trade press, RSS feeds & Google News \u00b7 '
        f'Auto-tagged by region, topic, company & capacity \u00b7 '
        f'Deduplicated across all sources</div>'
        f'<div class="banner-ts">\U0001f550 {now_str}</div>'
        f'</div>',
        unsafe_allow_html=True,
    )

    if "df_full" not in st.session_state:
        st.session_state.df_full = None

    if not go_btn and st.session_state.df_full is None:
        st.markdown(
            '<div style="display:flex;gap:.9rem;flex-wrap:wrap;margin-bottom:1.4rem;">',
            unsafe_allow_html=True,
        )
        features = [
            ("\U0001f578\ufe0f", "Multi-Source Scraping",
             "Hits DCD (5 regions), DataCenter Knowledge, DataCenterFrontier, "
             "PR Newswire, BusinessWire, Reuters & 5 Google News queries simultaneously."),
            ("\U0001f30d", "Global Coverage",
             "Covers 45+ countries across all continents with country/region auto-detection "
             "using 500+ geographic keywords and city names."),
            ("\U0001f9e0", "Smart Enrichment",
             "Every article auto-tagged: Topic, Sentiment (Approved/Proposed/Opened/Challenged), "
             "Capacity (MW/GW), Deal Size ($bn/$m), and up to 4 company names."),
            ("\u26a1", "Deduplication",
             "Fuzzy title matching (82% similarity threshold) collapses the same story "
             "appearing across multiple sources into a single clean record."),
            ("\U0001f5fa\ufe0f", "World Map",
             "Choropleth map showing article volume by country — instantly see "
             "where global data center activity is hottest."),
            ("\U0001f4ca", "5-Sheet Excel",
             "Export: All Articles + By Country + By Region + By Topic + By Company "
             "\u2014 all colour-coded with clickable hyperlinks."),
        ]
        row_html = ""
        for icon, title, desc in features:
            row_html += (
                f'<div style="flex:1;min-width:200px;background:#0b1628;border:1px solid #152038;'
                f'border-radius:10px;padding:1rem 1.15rem;">'
                f'<div style="font-size:1.4rem;margin-bottom:.4rem;">{icon}</div>'
                f'<div style="font-family:Syne,sans-serif;font-weight:700;color:#b8c8e0;'
                f'font-size:.9rem;margin-bottom:.3rem;">{title}</div>'
                f'<div style="font-size:.78rem;color:#3a5480;line-height:1.5;">{desc}</div>'
                f'</div>'
            )
        st.markdown(row_html + '</div>', unsafe_allow_html=True)
        return

    if go_btn:
        st.session_state.filters = {
            "regions": [], "topics": [], "sources": [],
            "sents": [], "keyword": "", "min_mw": 0,
        }
        if time_opt == "Custom Range" and custom_start and custom_end:
            cutoff    = datetime.combine(custom_start, datetime.min.time())
            cutoff_end = datetime.combine(custom_end,   datetime.max.time())
        elif sel_days is None:
            cutoff     = datetime.min
            cutoff_end = datetime.max
        else:
            cutoff     = datetime.now() - timedelta(days=sel_days)
            cutoff_end = datetime.max
        st.session_state.cutoff_end = cutoff_end

        pbar = st.progress(0.0, text="Initialising global scan...")

        active_html  = SCRAPE_SOURCES if use_html else []
        active_rss   = RSS_SOURCES    if use_rss  else []
        active_gn    = GNEWS_QUERIES  if use_gn   else []
        total_tasks  = len(active_html) + len(active_rss) + len(active_gn)
        done_count   = [0]

        def progress_cb(frac, label=""):
            pbar.progress(min(frac, 1.0), text=f"Scanning sources... {label}")

        raw = []
        with ThreadPoolExecutor(max_workers=10) as pool:
            futures = {}
            for src in active_html:
                futures[pool.submit(scrape_html_source, src, max_pages)] = src["name"]
            for src in active_rss:
                futures[pool.submit(fetch_rss, src)] = src["name"]
            for q, lbl in active_gn:
                futures[pool.submit(fetch_google_news, q, lbl)] = lbl

            for f in as_completed(futures):
                try:
                    items = f.result()
                    raw.extend(items)
                except Exception:
                    pass
                done_count[0] += 1
                pbar.progress(
                    done_count[0] / max(total_tasks, 1),
                    text=f"Fetched {futures[f]}... ({done_count[0]}/{total_tasks})",
                )

        pbar.progress(1.0, text="Enriching and deduplicating...")

        cutoff_end_val = st.session_state.get("cutoff_end", datetime.max)
        filtered = []
        for item in raw:
            d = item.get("date_obj")
            if d and d < cutoff:
                continue
            if d and d > cutoff_end_val:
                continue
            if not is_dc_relevant(item["headline"]):
                continue
            filtered.append(item)

        enriched = [enrich(i) for i in filtered]
        deduped  = deduplicate(enriched)

        df_full = (
            pd.DataFrame(deduped)
            .drop(columns=["_date_obj"], errors="ignore")
            .sort_values("Date", ascending=False)
            .reset_index(drop=True)
        )
        st.session_state.df_full   = df_full
        st.session_state.raw_count = len(raw)
        st.session_state.scan_time = datetime.now().strftime("%H:%M, %d %b %Y")
        pbar.empty()
        st.rerun()

    df_full = st.session_state.df_full
    if df_full is None or df_full.empty:
        st.warning("No articles found. Try expanding the date range or enabling more sources.")
        return

    filters = st.session_state.get("filters", {})
    df = df_full.copy()
    if filters.get("regions"):
        df = df[df["Region"].isin(filters["regions"])]
    if filters.get("topics"):
        df = df[df["Topic"].isin(filters["topics"])]
    if filters.get("sources"):
        df = df[df["Source"].isin(filters["sources"])]
    if filters.get("sents"):
        df = df[df["Sentiment"].isin(filters["sents"])]
    if filters.get("date_from") and filters.get("date_to"):
        df_dates = df[df["Date"] != "Unknown"].copy()
        df_dates["_dt"] = pd.to_datetime(df_dates["Date"], errors="coerce")
        mask = (df_dates["_dt"] >= pd.Timestamp(filters["date_from"])) &                (df_dates["_dt"] <= pd.Timestamp(filters["date_to"]))
        df_dates = df_dates[mask].drop(columns=["_dt"])
        df_unk = df[df["Date"] == "Unknown"]
        df = pd.concat([df_dates, df_unk], ignore_index=True)
    if filters.get("keyword"):
        kw = filters["keyword"].lower()
        df = df[df["Headline"].str.lower().str.contains(kw, na=False)]
    if filters.get("min_mw", 0) > 0:
        def extract_mw_val(cap):
            if not cap:
                return 0
            m = re.search(r"([\d.]+)\s*(GW|MW)", str(cap), re.I)
            if not m:
                return 0
            v = float(m.group(1))
            return v * 1000 if m.group(2).upper() == "GW" else v
        df = df[df["Capacity"].apply(extract_mw_val) >= filters["min_mw"]]
    if filters.get("countries"):
        df = df[df["Country"].isin(filters["countries"])]
    if filters.get("company_search"):
        cs = filters["company_search"].lower()
        df = df[
            df["Companies"].str.lower().str.contains(cs, na=False) |
            df["Headline"].str.lower().str.contains(cs, na=False)
        ]
    df = df.reset_index(drop=True)

    scan_ts = st.session_state.get("scan_time", "\u2014")
    top_country = df["Country"].value_counts().idxmax() if not df.empty else "\u2014"
    top_topic   = df["Topic"].value_counts().idxmax()   if not df.empty else "\u2014"
    cap_count   = int((df["Capacity"] != "").sum())
    deal_count  = int((df["Deal Size"] != "").sum())

    kpi_html = (
        '<div style="display:flex;gap:.8rem;margin-bottom:1.4rem;flex-wrap:wrap;">'
        + kpi("Sources Polled", len(SCRAPE_SOURCES) + len(RSS_SOURCES) + len(GNEWS_QUERIES), "blue", "HTML + RSS + Google News")
        + kpi("Raw Articles", st.session_state.raw_count, "cyan", "before dedup & filter")
        + kpi("Unique Articles", len(df_full), "green", "after deduplication")
        + kpi("Filtered View", len(df), "amber", "current filters applied")
        + kpi("Capacity Mentions", cap_count, "purple", "articles with MW/GW data")
        + kpi("Deal Mentions", deal_count, "red", "articles with $bn/$m data")
        + '</div>'
    )
    st.markdown(kpi_html, unsafe_allow_html=True)

    top_region = df["Region"].value_counts().idxmax() if not df.empty else "\u2014"
    latest_dt  = df["Date"].max() if not df.empty else "\u2014"
    pills_html = (
        '<div class="pill-row">'
        f'<div class="pill"><span class="pill-dot"></span>Scan: <b>{scan_ts}</b></div>'
        f'<div class="pill"><span class="pill-dot"></span>Top Region: <b>{top_region}</b></div>'
        f'<div class="pill"><span class="pill-dot"></span>Top Country: <b>{top_country}</b></div>'
        f'<div class="pill"><span class="pill-dot"></span>Top Topic: <b>{top_topic}</b></div>'
        f'<div class="pill"><span class="pill-dot"></span>Latest Article: <b>{latest_dt}</b></div>'
        f'<div class="pill"><span class="pill-dot"></span>Countries: <b>{df["Country"].nunique()}</b></div>'
        '</div>'
    )
    st.markdown(pills_html, unsafe_allow_html=True)

    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
        "\U0001f4f0 Feed",
        "\U0001f5fa\ufe0f World Map",
        "\U0001f4ca Analytics",
        "\U0001f3e2 By Company",
        "\U0001f9e0 Market Intel",
        "\u2b07\ufe0f Export",
    ])

    with tab1:
        st.markdown('<div class="sec-head">Global Intelligence Feed</div>', unsafe_allow_html=True)
        if df.empty:
            st.info("No articles match the current filters.")
        else:
            for _, row in df.iterrows():
                st.markdown(
                    article_card(
                        row["Headline"], row["Date"], row["URL"],
                        row["Source"], row["Country"], row["Topic"],
                        row.get("Capacity", ""), row.get("Deal Size", ""),
                        row.get("Sentiment", "News"),
                    ),
                    unsafe_allow_html=True,
                )

    with tab2:
        st.markdown('<div class="sec-head">Global Activity Map</div>', unsafe_allow_html=True)
        st.plotly_chart(chart_world_map(df), use_container_width=True, config={"displayModeBar": False})
        st.markdown('<div class="sec-head">Country Breakdown</div>', unsafe_allow_html=True)
        cc_df = df[df["Country"] != "Global"]["Country"].value_counts().reset_index()
        cc_df.columns = ["Country", "Articles"]
        cc_df["Region"] = cc_df["Country"].map(COUNTRY_TO_REGION).fillna("Global")
        st.markdown(dark_table(cc_df), unsafe_allow_html=True)

    with tab3:
        st.markdown('<div class="sec-head">Topic Distribution</div>', unsafe_allow_html=True)
        st.plotly_chart(chart_topic_bar(df), use_container_width=True, config={"displayModeBar": False})

        st.markdown('<div class="sec-head">Regional Distribution</div>', unsafe_allow_html=True)
        st.plotly_chart(chart_region_bar(df), use_container_width=True, config={"displayModeBar": False})

        st.markdown('<div class="sec-head">Top Countries</div>', unsafe_allow_html=True)
        st.plotly_chart(chart_country_bar(df), use_container_width=True, config={"displayModeBar": False})

        st.markdown('<div class="sec-head">Publication Volume Over Time</div>', unsafe_allow_html=True)
        tl = chart_timeline(df)
        if tl:
            st.plotly_chart(tl, use_container_width=True, config={"displayModeBar": False})

        st.markdown('<div class="sec-head">Sentiment / Project Status</div>', unsafe_allow_html=True)
        st.plotly_chart(chart_sentiment(df), use_container_width=True, config={"displayModeBar": False})

        st.markdown('<div class="sec-head">Topic Share</div>', unsafe_allow_html=True)
        st.plotly_chart(chart_donut(df), use_container_width=True, config={"displayModeBar": False})

        st.markdown('<div class="sec-head">Articles by Source</div>', unsafe_allow_html=True)
        st.plotly_chart(chart_source_bar(df), use_container_width=True, config={"displayModeBar": False})

        st.markdown('<div class="sec-head">Capacity Pipeline</div>', unsafe_allow_html=True)
        cap_df = df[df["Capacity"] != ""][["Headline", "Capacity", "Deal Size", "Country", "Topic", "Date"]].head(25)
        if not cap_df.empty:
            st.markdown(dark_table(cap_df), unsafe_allow_html=True)
        else:
            st.info("No capacity mentions in current filtered view.")

    with tab4:
        st.markdown('<div class="sec-head">Company Activity</div>', unsafe_allow_html=True)
        comp_rows = []
        for _, row in df.iterrows():
            if row.get("Companies"):
                for co in str(row["Companies"]).split(", "):
                    co = co.strip()
                    if co:
                        comp_rows.append(co)
        if comp_rows:
            from collections import Counter
            co_counts = Counter(comp_rows)
            co_df = pd.DataFrame(co_counts.most_common(30), columns=["Company", "Articles"])
            co_df_sorted = co_df.sort_values("Articles")
            n = len(co_df_sorted)
            co_colors = [
                f"rgba({int(0+71*i/max(n-1,1))},{int(71+(180-71)*i/max(n-1,1))},{int(225+(255-225)*i/max(n-1,1))},0.85)"
                for i in range(n)
            ]
            fig_co = go.Figure(go.Bar(
                x=co_df_sorted["Articles"], y=co_df_sorted["Company"],
                orientation="h",
                marker=dict(color=co_colors, line=dict(width=0)),
                text=co_df_sorted["Articles"], textposition="outside",
                textfont=dict(color=_TITLE, size=10),
                hovertemplate="<b>%{y}</b>: %{x} mentions<extra></extra>",
            ))
            _dark(fig_co, max(300, n * 22))
            fig_co.update_layout(
                title=dict(text="Top 30 Companies by Mentions", font=dict(color=_TITLE, size=13), x=0.01)
            )
            st.plotly_chart(fig_co, use_container_width=True, config={"displayModeBar": False})

            st.markdown('<div class="sec-head">Drill Into a Company</div>', unsafe_allow_html=True)
            sel_co = st.selectbox("Select company", co_df["Company"].tolist())
            co_articles = df[df["Companies"].str.contains(sel_co, na=False, case=False)]
            st.markdown(
                f'<div style="font-family:Inter,sans-serif;font-size:.82rem;color:#3a5480;margin-bottom:.7rem;">'
                f'<b style="color:#fff">{len(co_articles)}</b> articles mentioning '
                f'<b style="color:#00b4ff">{sel_co}</b></div>',
                unsafe_allow_html=True,
            )
            for _, row in co_articles.iterrows():
                st.markdown(
                    article_card(
                        row["Headline"], row["Date"], row["URL"],
                        row["Source"], row["Country"], row["Topic"],
                        row.get("Capacity", ""), row.get("Deal Size", ""),
                        row.get("Sentiment", "News"),
                    ),
                    unsafe_allow_html=True,
                )
        else:
            st.info("No company mentions detected in the current filtered view.")

    with tab5:
        st.markdown('<div class="sec-head">\U0001f9e0 Market Intelligence Summary</div>', unsafe_allow_html=True)

        filter_summary_parts = []
        if filters.get("countries") and len(filters["countries"]) < len(df_full["Country"].unique()):
            filter_summary_parts.append("Countries: " + ", ".join(filters["countries"][:5]) + ("..." if len(filters["countries"]) > 5 else ""))
        if filters.get("company_search"):
            filter_summary_parts.append(f"Company: {filters['company_search']}")
        if filters.get("topics") and len(filters["topics"]) < len(df_full["Topic"].unique()):
            filter_summary_parts.append("Topics: " + ", ".join(filters["topics"]))
        if filters.get("date_from") and filters.get("date_to"):
            filter_summary_parts.append(f"Date: {filters['date_from']} to {filters['date_to']}")
        if filters.get("keyword"):
            filter_summary_parts.append(f"Keyword: {filters['keyword']}")
        scan_date_range = f"{df['Date'].min()} to {df['Date'].max()}" if not df.empty and df["Date"].min() != "Unknown" else "all dates"

        sel_desc = " | ".join(filter_summary_parts) if filter_summary_parts else "All results (no specific filters applied)"

        ctx_html = (
            f'<div style="background:#0b1628;border:1px solid #152038;border-radius:10px;'
            f'padding:.9rem 1.2rem;margin-bottom:1rem;font-size:.8rem;color:#6a80a8;">'
            f'<b style="color:#b8c8e0;">Current selection:</b> {sel_desc}<br>'
            f'<b style="color:#b8c8e0;">Articles in view:</b> {len(df)} &nbsp;&nbsp;'
            f'<b style="color:#b8c8e0;">Date range:</b> {scan_date_range}'
            f'</div>'
        )
        st.markdown(ctx_html, unsafe_allow_html=True)

        if df.empty:
            st.info("No articles in the current filtered view. Adjust filters to generate a summary.")
        else:
            col_gen1, col_gen2 = st.columns([3, 1])
            with col_gen1:
                st.markdown(
                    '<div style="font-size:.82rem;color:#3a5480;line-height:1.6;">'  
                    'The AI will analyse all filtered articles and generate a structured market intelligence '  
                    'briefing covering key themes, major players, capacity pipeline, regulatory developments, '  
                    'and forward-looking signals.</div>',
                    unsafe_allow_html=True,
                )
            with col_gen2:
                gen_btn = st.button("\U0001f9e0 Generate Summary", use_container_width=True, type="primary")

            if gen_btn or st.session_state.get("intel_summary"):
                if gen_btn:
                    headlines_block = ""
                    for _, row in df.iterrows():
                        headlines_block += (
                            f"- [{row['Date']}] [{row['Country']}] [{row['Topic']}] "
                            f"[{row.get('Sentiment', '')}] {row['Headline']}"
                        )
                        if row.get("Capacity"):
                            headlines_block += f" | Capacity: {row['Capacity']}"
                        if row.get("Deal Size"):
                            headlines_block += f" | Deal: {row['Deal Size']}"
                        if row.get("Companies"):
                            headlines_block += f" | Companies: {row['Companies']}"
                        headlines_block += "\n"

                    prompt = f"""You are a senior data center market intelligence analyst at Wood Mackenzie. 
You have been given {len(df)} news articles about data center construction and investment activity.

Selection context: {sel_desc}
Date range covered: {scan_date_range}

Here are all the articles:
{headlines_block}

Write a structured market intelligence briefing with the following sections. Be specific, cite company names, locations, capacity figures, and deal sizes from the articles. Be analytical, not just descriptive — identify patterns, trends, and what it means for the market.

## Executive Summary
3-4 sentence high-level overview of what is happening in this market during this period.

## Key Themes & Trends
The 4-6 most significant patterns emerging from these articles. What is driving activity? What is changing?

## Major Projects & Deals
The most significant individual projects, investments, and deals. Include company, location, capacity or deal size where available.

## Regulatory & Permitting Landscape
Any moratoriums, approvals, rejections, legal challenges, or policy developments visible in the data.

## Power & Infrastructure
Key observations about power sourcing, capacity announcements, utility deals, and energy infrastructure.

## Company Activity Highlights
Which companies are most active? What strategies are visible? Any notable market entries or exits?

## Market Outlook
Based on the activity in these articles, what signals exist about near-term market direction? What should an analyst watch?

Write in a professional, analytical tone. Be concise but specific. Use bullet points within sections."""

                    with st.spinner("Generating market intelligence briefing..."):
                        try:
                            import requests as _req
                            resp = _req.post(
                                "https://api.anthropic.com/v1/messages",
                                headers={"Content-Type": "application/json"},
                                json={
                                    "model": "claude-sonnet-4-20250514",
                                    "max_tokens": 2000,
                                    "messages": [{"role": "user", "content": prompt}],
                                },
                                timeout=60,
                            )
                            resp.raise_for_status()
                            data = resp.json()
                            summary_text = "".join(
                                block.get("text", "") for block in data.get("content", [])
                                if block.get("type") == "text"
                            )
                            st.session_state.intel_summary = summary_text
                            st.session_state.intel_context = sel_desc
                        except Exception as e:
                            st.error(f"Could not generate summary: {e}")
                            st.session_state.intel_summary = None

                if st.session_state.get("intel_summary"):
                    context_label = st.session_state.get("intel_context", "")
                    brief_hdr = (
                        f'<div style="background:#0b1628;border:1px solid #0047e1;border-radius:10px;'
                        f'padding:1.2rem 1.5rem;margin-top:.8rem;">'
                        f'<div style="font-family:monospace;font-size:.65rem;letter-spacing:.12em;'
                        f'color:#0047e1;text-transform:uppercase;margin-bottom:.7rem;">'
                        f'🧠 AI Market Intelligence Briefing  ·  {context_label}</div>'
                    )
                    st.markdown(brief_hdr, unsafe_allow_html=True)
                    st.markdown(st.session_state.intel_summary)
                    st.markdown('</div>', unsafe_allow_html=True)

                    ts_intel = datetime.now().strftime("%Y%m%d_%H%M")
                    st.download_button(
                        "\U0001f4e5 Download Briefing (.txt)",
                        data=st.session_state.intel_summary.encode(),
                        file_name=f"DC_Intel_Briefing_{ts_intel}.txt",
                        mime="text/plain",
                        use_container_width=False,
                    )

    with tab6:
        st.markdown('<div class="sec-head">Export Data</div>', unsafe_allow_html=True)
        col_a, col_b = st.columns(2)
        ts    = datetime.now().strftime("%Y%m%d_%H%M")
        label = re.sub(r"[^a-zA-Z0-9]", "_", time_opt)

        with col_a:
            st.markdown(
                '<div style="background:#0b1628;border:1px solid #152038;border-radius:10px;'
                'padding:1.1rem 1.2rem;margin-bottom:.8rem;">'
                '<div style="font-size:1.5rem;margin-bottom:.4rem;">\U0001f4ca</div>'
                '<div style="font-family:Syne,sans-serif;font-weight:700;color:#b8c8e0;'
                'font-size:.95rem;margin-bottom:.3rem;">Excel Report (.xlsx)</div>'
                '<div style="font-size:.78rem;color:#2a3e60;line-height:1.5;">'
                '5 sheets: All Articles \u00b7 By Country \u00b7 By Region \u00b7 By Topic \u00b7 By Company<br>'
                'Colour-coded badges, auto-filter, frozen headers, clickable URLs.</div></div>',
                unsafe_allow_html=True,
            )
            st.download_button(
                "\U0001f4e5 Download Excel Report",
                data=build_excel(df),
                file_name=f"GlobalDCIntel_{label}_{ts}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        with col_b:
            st.markdown(
                '<div style="background:#0b1628;border:1px solid #152038;border-radius:10px;'
                'padding:1.1rem 1.2rem;margin-bottom:.8rem;">'
                '<div style="font-size:1.5rem;margin-bottom:.4rem;">\U0001f4c4</div>'
                '<div style="font-family:Syne,sans-serif;font-weight:700;color:#b8c8e0;'
                'font-size:.95rem;margin-bottom:.3rem;">CSV Export</div>'
                '<div style="font-size:.78rem;color:#2a3e60;line-height:1.5;">'
                'Flat CSV of the filtered view.<br>'
                'Ready for Excel, Python, PowerBI, or Tableau.</div></div>',
                unsafe_allow_html=True,
            )
            st.download_button(
                "\U0001f4e5 Download CSV",
                data=df.to_csv(index=False).encode(),
                file_name=f"GlobalDCIntel_{label}_{ts}.csv",
                mime="text/csv",
                use_container_width=True,
            )

        st.markdown('<div class="sec-head">Full Article Preview</div>', unsafe_allow_html=True)
        display_cols = ["Headline","Date","Source","Country","Region",
                        "Topic","Sentiment","Capacity","Deal Size","Companies","URL"]
        st.markdown(
            dark_table(df[[c for c in display_cols if c in df.columns]]),
            unsafe_allow_html=True,
        )


if __name__ == "__main__":
    main()
