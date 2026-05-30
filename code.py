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
# _PLAT_VER   : internal build version string
# _PLAT_TOKEN : encoded authorship/license identifier — do not modify
_PLAT_VER   = "1.0.0"
_PLAT_TOKEN = "\x53\x68\x61\x72\x75\x67\x68\x20\x41"   # platform license key
# Copyright (c) Sharugh A. All rights reserved. Unauthorised redistribution
# or removal of authorship metadata is prohibited under applicable IP law.

# ── Word (.docx) export ──────────────────────────────────────────────────────
try:
    from docx import Document as DocxDocument
    from docx.shared import Pt, RGBColor, Inches, Cm
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    _DOCX_OK = True
except ImportError:
    _DOCX_OK = False

# ── PDF export ───────────────────────────────────────────────────────────────
try:
    from reportlab.lib.pagesizes import letter
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import inch
    from reportlab.lib import colors as rl_colors
    from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer,
                                    HRFlowable, KeepTogether)
    from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_JUSTIFY
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
@import url('https://fonts.googleapis.com/css2?family=Syne:wght@400;600;700;800&family=DM+Mono:wght@400;500&family=Inter:wght@300;400;500;600&display=swap');

html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
.stApp { background: #05080f; color: #e2eaf8; }

/* ── Scrollbar ──────────────────────────────────────────────────────────────── */
::-webkit-scrollbar { width: 4px; height: 4px; }
::-webkit-scrollbar-track { background: #05080f; }
::-webkit-scrollbar-thumb { background: #1a2d4a; border-radius: 4px; }
::-webkit-scrollbar-thumb:hover { background: #0047e1; }

/* ── Global transitions (exclude Plotly to avoid lag) ─────────────────────── */
*:not(.plotly-graph-div):not(.plotly-graph-div *) {
    transition: background 0.15s ease, border-color 0.15s ease,
                box-shadow 0.18s ease, opacity 0.15s ease, transform 0.15s ease;
}

/* ── Sidebar shell ─────────────────────────────────────────────────────────── */
[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #070d1c 0%, #060a16 100%) !important;
    border-right: 1px solid #0f1e35 !important;
}
[data-testid="stSidebar"] * { color: #b0c4de !important; }
[data-testid="stSidebar"] hr { border-color: #0f1e35 !important; }

/* ── Sidebar: always expanded ──────────────────────────────────────────────── */
[data-testid="stSidebarCollapseButton"] { display: none !important; }
[data-testid="stSidebar"][aria-expanded="false"],
[data-testid="stSidebar"] {
    display: block !important; visibility: visible !important;
    opacity: 1 !important; pointer-events: auto !important;
    min-width: 252px !important; max-width: 252px !important;
    width: 252px !important; transform: none !important;
    margin-left: 0 !important; left: 0 !important;
    position: relative !important; overflow: visible !important;
}
[data-testid="stSidebarCollapsedControl"],
div[data-testid="collapsedControl"] {
    display: flex !important; visibility: visible !important;
    opacity: 1 !important; pointer-events: auto !important;
    background: #0047e1 !important; border-radius: 0 10px 10px 0 !important;
    border: 1px solid #0060ff !important; border-left: none !important;
    width: 2rem !important; z-index: 99999 !important;
}
[data-testid="stSidebarCollapsedControl"] button,
div[data-testid="collapsedControl"] button { color:#fff!important; background:transparent!important; }
[data-testid="stSidebarCollapsedControl"] svg,
div[data-testid="collapsedControl"] svg { fill:#fff!important; stroke:#fff!important; }

/* ── Sidebar buttons ───────────────────────────────────────────────────────── */
[data-testid="stSidebar"] .stButton button {
    background: linear-gradient(135deg, #0044d4, #0090e0) !important;
    color: #fff !important; border: none !important; border-radius: 9px !important;
    font-family: 'Syne', sans-serif !important; font-weight: 700 !important;
    font-size: 0.88rem !important; letter-spacing: 0.05em !important;
    padding: 0.6rem 1rem !important;
    box-shadow: 0 2px 12px rgba(0,71,225,0.25) !important;
}
[data-testid="stSidebar"] .stButton button:hover {
    opacity: .88 !important;
    box-shadow: 0 4px 20px rgba(0,71,225,0.4) !important;
    transform: translateY(-1px) !important;
}

/* ── Sidebar multiselect ───────────────────────────────────────────────────── */
[data-testid="stSidebar"] [data-baseweb="select"] { background: #080f1e !important; }
[data-testid="stSidebar"] [data-baseweb="select"] > div:first-child {
    background: #080f1e !important; border: 1px solid #172540 !important;
    border-radius: 8px !important;
}
[data-testid="stSidebar"] [data-baseweb="select"] > div:first-child:focus-within {
    border-color: #0047e1 !important;
    box-shadow: 0 0 0 2px rgba(0,71,225,0.2) !important;
}
[data-testid="stSidebar"] [data-baseweb="select"] span { color: #2e4870 !important; }
[data-testid="stSidebar"] [data-baseweb="tag"] {
    background: #0e2040 !important; border: 1px solid #0047e1 !important;
    border-radius: 5px !important;
}
[data-testid="stSidebar"] [data-baseweb="tag"] span { color: #70b0ff !important; }
[data-testid="stSidebar"] [data-baseweb="tag"] [role="presentation"] svg { fill: #2e4870 !important; }
[data-testid="stSidebar"] [data-baseweb="select"] input {
    background: transparent !important; color: #c0d0e8 !important;
    caret-color: #0047e1 !important;
}

/* ── Dropdown menus ────────────────────────────────────────────────────────── */
[data-baseweb="popover"], [data-baseweb="menu"],
[role="listbox"], ul[data-baseweb="menu"] {
    background: #08111f !important; border: 1px solid #172540 !important;
    border-radius: 10px !important; box-shadow: 0 12px 40px rgba(0,0,0,0.7) !important;
}
[data-baseweb="menu"] li, [role="option"], [data-baseweb="menu"] [role="option"] {
    background: #08111f !important; color: #a8bcd4 !important;
}
[data-baseweb="menu"] li:hover, [role="option"]:hover,
[data-baseweb="menu"] [role="option"]:hover {
    background: #0e2040 !important; color: #fff !important;
}
[data-baseweb="menu"] [aria-selected="true"],
[role="option"][aria-selected="true"] {
    background: #0e2040 !important; color: #00b4ff !important;
}
[data-baseweb="popover"] input, [data-baseweb="menu"] input {
    background: #05080f !important; border: 1px solid #172540 !important;
    color: #c0d0e8 !important; border-radius: 6px !important;
}

/* ── Sidebar text / number / date inputs ───────────────────────────────────── */
[data-testid="stSidebar"] .stTextInput input,
[data-testid="stSidebar"] input[type="number"],
[data-testid="stSidebar"] [data-testid="stDateInput"] input {
    background: #080f1e !important; border: 1px solid #172540 !important;
    border-radius: 8px !important; color: #c8d8f0 !important;
}
[data-testid="stSidebar"] .stTextInput input:focus,
[data-testid="stSidebar"] input[type="number"]:focus,
[data-testid="stSidebar"] [data-testid="stDateInput"] input:focus {
    border-color: #0047e1 !important;
    box-shadow: 0 0 0 2px rgba(0,71,225,0.2) !important;
}

/* ── Main content inputs ───────────────────────────────────────────────────── */
.stMultiSelect [data-baseweb="select"] { background: #080f1e !important; border-color: #122035 !important; }
.stSelectbox [data-baseweb="select"] > div { background: #080f1e !important; border-color: #122035 !important; }
.stTextInput input {
    background: #080f1e !important; border: 1px solid #122035 !important;
    border-radius: 8px !important; color: #c8d8f0 !important;
}
.stTextInput input:focus { border-color: #0047e1 !important; }

/* ── Tabs ──────────────────────────────────────────────────────────────────── */
.stTabs [data-baseweb="tab-list"] {
    background: #080f1e !important; border-radius: 11px !important;
    padding: 4px !important; gap: 3px !important;
    border: 1px solid #122035 !important;
    box-shadow: inset 0 1px 3px rgba(0,0,0,0.4) !important;
}
.stTabs [data-baseweb="tab"] {
    background: transparent !important; color: #334d6e !important;
    border-radius: 8px !important; font-family: 'Syne', sans-serif !important;
    font-weight: 600 !important; font-size: .78rem !important;
    letter-spacing: .04em !important; padding: .42rem 1rem !important;
}
.stTabs [data-baseweb="tab"]:hover:not([aria-selected="true"]) {
    background: rgba(0,71,225,0.1) !important; color: #c0d0e8 !important;
}
.stTabs [aria-selected="true"] {
    background: linear-gradient(135deg,#0044d4,#0090e0) !important;
    color: #fff !important;
    box-shadow: 0 2px 10px rgba(0,71,225,0.35) !important;
}
.stTabs [data-baseweb="tab-panel"] { padding-top: 1.1rem !important; }

/* ── Download buttons ──────────────────────────────────────────────────────── */
.stDownloadButton button {
    background: linear-gradient(135deg, #002a0a, #004d14) !important;
    color: #00e676 !important; border: 1px solid #009e40 !important;
    border-radius: 9px !important; font-family: 'Syne', sans-serif !important;
    font-weight: 700 !important; letter-spacing: .04em !important;
    box-shadow: 0 2px 10px rgba(0,150,70,0.2) !important;
}
.stDownloadButton button:hover {
    opacity: .85 !important;
    box-shadow: 0 4px 16px rgba(0,150,70,0.35) !important;
    transform: translateY(-1px) !important;
}

/* ── Banner ────────────────────────────────────────────────────────────────── */
.gl-banner {
    background: linear-gradient(135deg, #060e1c 0%, #09183a 50%, #060e1c 100%);
    border: 1px solid #102035; border-radius: 18px;
    padding: 2rem 2.5rem; margin-bottom: 1.6rem;
    position: relative; overflow: hidden;
    box-shadow: 0 4px 40px rgba(0,71,225,0.08);
}
.gl-banner::before {
    content: ''; position: absolute; top: -90px; right: -50px;
    width: 340px; height: 340px;
    background: radial-gradient(circle, rgba(0,71,225,0.18) 0%, transparent 68%);
    border-radius: 50%;
}
.gl-banner::after {
    content: ''; position: absolute; bottom: -60px; left: 20%;
    width: 260px; height: 260px;
    background: radial-gradient(circle, rgba(0,180,255,0.1) 0%, transparent 68%);
    border-radius: 50%;
}
.banner-eyebrow {
    font-family: 'DM Mono', monospace; font-size: .65rem;
    letter-spacing: .22em; color: #00b4ff;
    text-transform: uppercase; margin-bottom: .5rem;
    display: flex; align-items: center; gap: .5rem;
}
.banner-eyebrow::before {
    content: ''; display: inline-block; width: 6px; height: 6px;
    background: #00b4ff; border-radius: 50%;
    animation: pulse-dot 2s ease-in-out infinite;
}
@keyframes pulse-dot {
    0%, 100% { opacity: 1; transform: scale(1); }
    50% { opacity: 0.4; transform: scale(0.7); }
}
.banner-title {
    font-family: 'Syne', sans-serif; font-size: 2rem;
    font-weight: 800; color: #fff; line-height: 1.1; margin-bottom: .35rem;
}
.banner-title span { color: #00b4ff; }
.banner-sub { font-size: .82rem; color: #5a7298; font-weight: 300; line-height: 1.6; }
.banner-ts {
    font-family: 'DM Mono', monospace; font-size: .65rem;
    color: #1e3356; margin-top: .8rem; letter-spacing: .06em;
}

/* ── Section headings ──────────────────────────────────────────────────────── */
.sec-head {
    font-family: 'Syne', sans-serif; font-size: .85rem; font-weight: 700;
    color: #a8bcd4; letter-spacing: .08em; text-transform: uppercase;
    border-left: 3px solid #0047e1; padding-left: .75rem;
    margin: 1.8rem 0 1rem 0;
    position: relative;
}
.sec-head::after {
    content: ''; position: absolute; left: 0; bottom: -6px;
    width: 40px; height: 1px;
    background: linear-gradient(90deg, #0047e1, transparent);
}
.sec-head:hover { border-left-color: #00b4ff !important; color: #fff !important; cursor: default; }

/* ── Pills ─────────────────────────────────────────────────────────────────── */
.pill-row { display: flex; flex-wrap: wrap; gap: .45rem; margin-bottom: 1.4rem; }
.pill {
    background: #080f1e; border: 1px solid #122035; border-radius: 20px;
    padding: .35rem .9rem; font-size: .76rem; color: #5a7298;
    display: flex; align-items: center; gap: .38rem;
}
.pill b { color: #e2eaf8; }
.pill-dot { width: 6px; height: 6px; border-radius: 50%; background: #0047e1; flex-shrink: 0; }
.pill:hover {
    border-color: #0047e1 !important; background: #0d1e3c !important;
    color: #c0d0e8 !important; cursor: default;
}

/* ── Global misc ───────────────────────────────────────────────────────────── */
hr { border-color: #0f1e35 !important; }
#MainMenu, footer, header { visibility: hidden; }

/* ── KPI cards ─────────────────────────────────────────────────────────────── */
.kpi-card {
    transition: transform 0.18s ease, border-color 0.18s ease, box-shadow 0.2s ease !important;
}
.kpi-card:hover {
    transform: translateY(-3px) !important; border-color: #0047e1 !important;
    box-shadow: 0 10px 36px rgba(0,71,225,0.22) !important;
}

/* ── Table ─────────────────────────────────────────────────────────────────── */
.dc-table tr:hover td { background: #0c1d38 !important; }

/* ── Hover cards ───────────────────────────────────────────────────────────── */
.hover-card { transition: transform 0.18s ease, border-color 0.18s ease, box-shadow 0.2s ease !important; }
.hover-card:hover {
    transform: translateY(-2px) !important; border-color: #0047e1 !important;
    box-shadow: 0 6px 28px rgba(0,71,225,0.18) !important;
}
.feature-card { transition: transform 0.18s ease, border-color 0.2s ease, box-shadow 0.2s ease !important; cursor: default; }
.feature-card:hover {
    transform: translateY(-4px) !important; border-color: #0047e1 !important;
    box-shadow: 0 12px 40px rgba(0,71,225,0.2) !important;
}
.saved-scan-card { transition: border-color 0.18s ease, box-shadow 0.18s ease !important; }
.saved-scan-card:hover {
    border-color: #a855f7 !important; box-shadow: 0 4px 20px rgba(168,85,247,0.15) !important;
}
.article-card-wrap { transition: border-color 0.18s ease, box-shadow 0.2s ease, transform 0.15s ease !important; }
.article-card-wrap:hover {
    border-color: #0047e1 !important; box-shadow: 0 6px 28px rgba(0,71,225,0.14) !important;
    transform: translateY(-1px) !important;
}
.score-badge { transition: filter 0.15s ease, transform 0.15s ease !important; }
.score-badge:hover { filter: brightness(1.3) !important; transform: scale(1.06) !important; }

/* ── Plotly chart wrapper ──────────────────────────────────────────────────── */
[data-testid="stPlotlyChart"] {
    border-radius: 12px !important;
    border: 1px solid #0f1e35 !important;
    overflow: hidden !important;
    transition: border-color 0.2s ease, box-shadow 0.25s ease !important;
}
[data-testid="stPlotlyChart"]:hover {
    border-color: #1a3060 !important;
    box-shadow: 0 2px 32px rgba(0,71,225,0.1) !important;
}

/* ── Sidebar run button special override ────────────────────────────────────── */
[data-testid="stSidebar"] .stButton:last-of-type button {
    background: linear-gradient(135deg, #0044d4 0%, #0090e0 100%) !important;
    font-size: .95rem !important;
    padding: .7rem 1rem !important;
    letter-spacing: .06em !important;
}

/* ── Progress bar ──────────────────────────────────────────────────────────── */
[data-testid="stProgressBar"] > div > div {
    background: linear-gradient(90deg, #0047e1, #00b4ff) !important;
    border-radius: 4px !important;
}

/* ── Info / warning / error alerts ────────────────────────────────────────── */
[data-testid="stAlert"] {
    background: #080f1e !important; border-radius: 10px !important;
    border: 1px solid #172540 !important;
}
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
    "Grid / ISO / RTO": "#9c27b0",
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


# ═══════════════════════════════════════════════════════════════════════════════
#  ISO / RTO / GRID-OPERATOR REGISTRY
# ═══════════════════════════════════════════════════════════════════════════════
US_ISO_RTO = {
    "PJM": {
        "full_name": "PJM Interconnection", "type": "RTO",
        "states": ["Delaware","Illinois","Indiana","Kentucky","Maryland","Michigan",
                   "New Jersey","North Carolina","Ohio","Pennsylvania","Tennessee",
                   "Virginia","West Virginia","Washington D.C.","Northern Virginia"],
        "keywords": ["PJM","PJM Interconnection","PJM queue","PJM RTO","Northern Virginia queue",
                     "Dominion Energy","AEP Ohio","Commonwealth Edison","PPL","Allegheny","PSEG",
                     "BGE","Pepco","Duquesne","FirstEnergy","MISO-PJM seam","PJM capacity market",
                     "PJM RTEP","PJM large load","PJM interconnection queue","PJM ATSI",
                     "PJM COMED","PJM DEOK","PJM GI queue","PJM cluster study"],
    },
    "ERCOT": {
        "full_name": "Electric Reliability Council of Texas", "type": "ISO (non-FERC)",
        "states": ["Texas"],
        "keywords": ["ERCOT","Electric Reliability Council of Texas","ERCOT queue",
                     "ERCOT large load study","ERCOT interconnection","Texas grid","Oncor",
                     "AEP Texas","CenterPoint Energy","LCRA","ERCOT LaaS","ERCOT SDT",
                     "ERCOT CRDP","ERCOT 4CP","ERCOT nodal","ERCOT NOIE","ERCOT West",
                     "ERCOT Houston Load Zone","ERCOT North Load Zone","ERCOT South Load Zone",
                     "Texas data center load","ERCOT generation interconnection"],
    },
    "CAISO": {
        "full_name": "California Independent System Operator", "type": "ISO",
        "states": ["California"],
        "keywords": ["CAISO","California ISO","CAISO queue","NP-15","SP-15","ZP-26",
                     "Pacific Gas and Electric","PG&E","SCE","Southern California Edison",
                     "SDG&E","IID","CAISO WEIM","Western EIM","CAISO large load",
                     "California grid","CAISO interconnection","California data center load"],
    },
    "MISO": {
        "full_name": "Midcontinent Independent System Operator", "type": "ISO",
        "states": ["Arkansas","Illinois","Indiana","Iowa","Kentucky","Louisiana","Michigan",
                   "Minnesota","Mississippi","Missouri","Montana","North Dakota","South Dakota",
                   "Texas","Wisconsin"],
        "keywords": ["MISO","Midcontinent ISO","MISO queue","MISO South","MISO North",
                     "MISO Central","MISO interconnection","Entergy","Ameren","Consumers Energy",
                     "DTE Energy","Evergy","ITC Holdings","MISO LRTP","MISO futures",
                     "MISO large load","MISO GI queue","Midwest grid"],
    },
    "NYISO": {
        "full_name": "New York Independent System Operator", "type": "ISO",
        "states": ["New York"],
        "keywords": ["NYISO","New York ISO","NYISO queue","Con Edison","National Grid NY",
                     "PSEG Long Island","NYSEG","RG&E","Central Hudson","Orange and Rockland",
                     "NYISO Zone J","NYISO Zone K","NYISO interconnection","New York grid",
                     "NYC data center queue","New York data center load"],
    },
    "ISO-NE": {
        "full_name": "ISO New England", "type": "ISO",
        "states": ["Connecticut","Maine","Massachusetts","New Hampshire","Rhode Island","Vermont"],
        "keywords": ["ISO-NE","ISO New England","New England ISO","ISO-NE queue","Eversource",
                     "National Grid NE","Avangrid","Green Mountain Power","Unitil",
                     "ISO-NE FCM","Forward Capacity Market","New England grid","NEPOOL",
                     "New England interconnection","New England data center load"],
    },
    "SPP": {
        "full_name": "Southwest Power Pool", "type": "RTO",
        "states": ["Kansas","Nebraska","Oklahoma","New Mexico","Wyoming","South Dakota",
                   "North Dakota","Montana","Colorado","Texas"],
        "keywords": ["SPP","Southwest Power Pool","SPP RTO","SPP queue","Western SPP",
                     "SPP interconnection","Evergy SPP","OG&E","SWEPCO","Westar",
                     "SPP market","SPP DISIS","Great Plains grid"],
    },
    "WECC": {
        "full_name": "Western Electricity Coordinating Council", "type": "Regional Entity",
        "keywords": ["WECC","Western Interconnection","WECC reliability","Desert Southwest",
                     "Western grid","Nevada grid","Arizona grid","WECC interconnection"],
    },
    "SERC": {
        "full_name": "SERC Reliability Corporation", "type": "Regional Entity",
        "keywords": ["SERC","SERC Reliability","Southeast grid","Southeastern grid",
                     "Southern Company","Duke Energy","Dominion South","TVA"],
    },
    "NERC": {
        "full_name": "North American Electric Reliability Corporation", "type": "Reliability Org",
        "keywords": ["NERC","North American Electric Reliability","NERC standard",
                     "NERC compliance","bulk electric system","BES","grid reliability"],
    },
    "BPA": {
        "full_name": "Bonneville Power Administration", "type": "Federal PMA",
        "keywords": ["BPA","Bonneville Power","BPA interconnection",
                     "Pacific Northwest grid","Columbia River hydropower"],
    },
    "TVA": {
        "full_name": "Tennessee Valley Authority", "type": "Federal Utility",
        "keywords": ["TVA","Tennessee Valley Authority","TVA power","TVA rate",
                     "TVA large power", "TVA data center"],
    },
}

GLOBAL_GRID_OPERATORS = {
    "IESO": {"country":"Canada","full_name":"Independent Electricity System Operator",
             "keywords":["IESO","Ontario grid","Ontario electricity","Ontario power"]},
    "AESO": {"country":"Canada","full_name":"Alberta Electric System Operator",
             "keywords":["AESO","Alberta grid","Alberta electricity","Alberta interconnection"]},
    "BC Hydro": {"country":"Canada","full_name":"BC Hydro",
             "keywords":["BC Hydro","British Columbia grid","BCUC","Site C dam"]},
    "Hydro-Québec": {"country":"Canada","full_name":"Hydro-Québec",
             "keywords":["Hydro-Québec","Hydro Quebec","Quebec grid","Quebec electricity"]},
    "National Grid ESO": {"country":"United Kingdom",
             "full_name":"National Grid Electricity System Operator",
             "keywords":["National Grid ESO","UK grid","UK transmission","Ofgem","DNO",
                         "National Grid UK","electricity connection UK","UK grid queue",
                         "UK TEC queue","TEC holder UK","UK data center grid"]},
    "RTE": {"country":"France","full_name":"Réseau de Transport d'Électricité",
             "keywords":["RTE","French grid","French TSO","France electricity transmission",
                         "RTE connection","France data center grid"]},
    "TenneT": {"country":"Germany/Netherlands","full_name":"TenneT TSO",
             "keywords":["TenneT","German TSO","Netherlands TSO","Dutch grid operator",
                         "TenneT connection"]},
    "50Hertz": {"country":"Germany","full_name":"50Hertz Transmission",
             "keywords":["50Hertz","50Hertz grid","Berlin electricity","East Germany grid"]},
    "Amprion": {"country":"Germany","full_name":"Amprion GmbH",
             "keywords":["Amprion","NRW grid","Germany electricity","Amprion connection"]},
    "TransnetBW": {"country":"Germany","full_name":"TransnetBW GmbH",
             "keywords":["TransnetBW","Baden-Württemberg grid","Stuttgart electricity"]},
    "Elia": {"country":"Belgium","full_name":"Elia Transmission Belgium",
             "keywords":["Elia","Belgian grid","Belgium electricity","Brussels grid",
                         "Elia connection"]},
    "REE": {"country":"Spain","full_name":"Red Eléctrica de España",
             "keywords":["REE","Spanish grid","Spain TSO","Spain electricity","REE connection"]},
    "Terna": {"country":"Italy","full_name":"Terna SpA",
             "keywords":["Terna","Italian grid","Italy TSO","Italy electricity","Terna connection"]},
    "EirGrid": {"country":"Ireland","full_name":"EirGrid plc",
             "keywords":["EirGrid","Irish grid","Ireland TSO","Ireland electricity","SEM",
                         "Single Electricity Market Ireland","CRU Ireland",
                         "Ireland grid moratorium","Dublin data center moratorium",
                         "EirGrid connection","Ireland grid capacity"]},
    "Swissgrid": {"country":"Switzerland","full_name":"Swissgrid AG",
             "keywords":["Swissgrid","Swiss grid","Switzerland electricity"]},
    "Fingrid": {"country":"Finland","full_name":"Fingrid Oyj",
             "keywords":["Fingrid","Finnish grid","Finland electricity","Helsinki grid"]},
    "Svenska Kraftnät": {"country":"Sweden","full_name":"Svenska kraftnät",
             "keywords":["Svenska Kraftnät","Swedish grid","Sweden electricity","SvK"]},
    "Statnett": {"country":"Norway","full_name":"Statnett SF",
             "keywords":["Statnett","Norwegian grid","Norway electricity","Oslo grid"]},
    "Energinet": {"country":"Denmark","full_name":"Energinet",
             "keywords":["Energinet","Danish grid","Denmark electricity","Copenhagen grid"]},
    "PSE": {"country":"Poland","full_name":"Polskie Sieci Elektroenergetyczne",
             "keywords":["PSE","Polish grid","Poland electricity","Warsaw grid",
                         "PSE connection"]},
    "ENTSO-E": {"country":"Europe","full_name":"European Network of TSOs",
             "keywords":["ENTSO-E","European grid","pan-European electricity","EU TSO",
                         "European electricity market","EU power market"]},
    "SEC Saudi": {"country":"Saudi Arabia","full_name":"Saudi Electricity Company",
             "keywords":["Saudi Electricity Company","SEC Saudi","Saudi grid",
                         "Saudi power","KSA grid","KSA electricity"]},
    "DEWA": {"country":"UAE","full_name":"Dubai Electricity and Water Authority",
             "keywords":["DEWA","Dubai electricity","Dubai grid","UAE power",
                         "DEWA connection","DEWA data center"]},
    "TAQA": {"country":"UAE","full_name":"Abu Dhabi National Energy Company",
             "keywords":["TAQA","ADNOC utilities","Abu Dhabi electricity","UAE grid",
                         "ADNOC power"]},
    "KAHRAMAA": {"country":"Qatar","full_name":"Qatar General Electricity & Water Corporation",
             "keywords":["KAHRAMAA","Qatar electricity","Doha grid","Qatar power"]},
    "AEMO": {"country":"Australia","full_name":"Australian Energy Market Operator",
             "keywords":["AEMO","Australian grid","NEM","National Electricity Market Australia",
                         "WEM","Wholesale Electricity Market Australia","AEMO connection",
                         "Australia electricity queue","Sydney grid","Melbourne grid",
                         "AEMO data center load"]},
    "PowerGrid India": {"country":"India","full_name":"Power Grid Corporation of India",
             "keywords":["PowerGrid India","Power Grid Corporation","Indian grid",
                         "India electricity transmission","PGCIL","CTUIL",
                         "India power infrastructure","Mumbai grid","Delhi grid",
                         "Bangalore grid","Chennai grid"]},
    "KEPCO": {"country":"South Korea","full_name":"Korea Electric Power Corporation",
             "keywords":["KEPCO","Korea electric","South Korea grid","Korean power",
                         "KEPCO connection"]},
    "State Grid China": {"country":"China","full_name":"State Grid Corporation of China",
             "keywords":["State Grid China","SGCC","China grid","China electricity",
                         "China Southern Grid","CSG","State Grid data center"]},
    "TNB": {"country":"Malaysia","full_name":"Tenaga Nasional Berhad",
             "keywords":["TNB","Tenaga Nasional","Malaysian grid","Malaysia electricity",
                         "Johor grid","Kuala Lumpur grid","TNB connection"]},
    "PLN": {"country":"Indonesia","full_name":"Perusahaan Listrik Negara",
             "keywords":["PLN Indonesia","Indonesian grid","Java-Bali grid",
                         "Jakarta electricity","PLN power"]},
    "TEPCO": {"country":"Japan","full_name":"Tokyo Electric Power Company",
             "keywords":["TEPCO","Tokyo Electric","Japan grid","Japanese power",
                         "Kansai Electric","Chubu Electric","Tokyo grid"]},
    "EMA Singapore": {"country":"Singapore","full_name":"Energy Market Authority",
             "keywords":["EMA Singapore","Singapore grid","Singapore electricity",
                         "SP Group","Singapore power","EMA moratorium",
                         "Singapore data center electricity"]},
    "Transpower": {"country":"New Zealand","full_name":"Transpower New Zealand",
             "keywords":["Transpower NZ","New Zealand grid","NZ electricity","Auckland grid"]},
    "ANEEL": {"country":"Brazil","full_name":"Agência Nacional de Energia Elétrica",
             "keywords":["ANEEL","Brazilian grid","Brazil electricity","ONS Brazil",
                         "Brazil power","CCEE","São Paulo grid","ONS interconnection"]},
    "CENACE": {"country":"Mexico","full_name":"Centro Nacional de Control de Energía",
             "keywords":["CENACE","Mexican grid","Mexico electricity","SENER Mexico",
                         "Mexico City grid","Monterrey grid"]},
    "Eskom": {"country":"South Africa","full_name":"Eskom Holdings SOC Ltd",
             "keywords":["Eskom","South Africa grid","South African power",
                         "Johannesburg electricity","Cape Town grid","load shedding",
                         "Eskom data center"]},
    "Kenya Power": {"country":"Kenya","full_name":"Kenya Power and Lighting Company",
             "keywords":["KPLC","Kenya Power","Nairobi grid","Kenya electricity"]},
    "NERC Nigeria": {"country":"Nigeria","full_name":"Nigerian Electricity Regulatory Commission",
             "keywords":["NERC Nigeria","Nigerian grid","Nigeria electricity","Lagos power"]},
}

ISO_RTO_KEYWORDS = {}
for _iso, _data in US_ISO_RTO.items():
    for _kw in _data.get("keywords", []):
        ISO_RTO_KEYWORDS[_kw.lower()] = _iso
for _op, _data in GLOBAL_GRID_OPERATORS.items():
    for _kw in _data.get("keywords", []):
        ISO_RTO_KEYWORDS[_kw.lower()] = _op

ALL_ISO_RTO_KW_LIST = sorted(set(ISO_RTO_KEYWORDS.keys()))

def detect_iso_rto(text):
    """Return the ISO/RTO/grid-operator relevant to this article, or empty string."""
    t = text.lower()
    for kw, iso in ISO_RTO_KEYWORDS.items():
        if kw in t:
            return iso
    return ""

COUNTRY_KEYWORDS = {k: [k] for k in COUNTRY_TO_REGION}
COUNTRY_KEYWORDS.update({
    "United States": ["United States", "U.S.", r"\bUS\b", "America", "American",
                      # States
                      "Virginia", "Texas", "California", "Georgia", "Ohio",
                      "Indiana", "Nevada", "Arizona", "Illinois", "Pennsylvania",
                      "North Carolina", "Florida", "Oregon", "Washington",
                      "Colorado", "Utah", "Wyoming", "Montana", "Idaho",
                      "Minnesota", "Wisconsin", "Michigan", "Iowa", "Kansas",
                      "Missouri", "Oklahoma", "Arkansas", "Louisiana",
                      "Mississippi", "Alabama", "Tennessee", "Kentucky",
                      "Maryland", "New Jersey", "New York", "Massachusetts",
                      "Connecticut", "Maine", "New Mexico", "Alaska", "Hawaii",
                      "Delaware", "Rhode Island", "Vermont", "New Hampshire",
                      "South Carolina", "South Dakota", "North Dakota",
                      "Nebraska", "West Virginia",
                      # Major metros / cities
                      "Dallas", "Austin", "Chicago", "Phoenix", "Atlanta",
                      "Seattle", "Denver", "Reno", "Ashburn", "Portland",
                      "Las Vegas", "San Jose", "San Francisco", "Los Angeles",
                      "Houston", "Miami", "Boston", "Columbus", "Nashville",
                      "Charlotte", "Northern Virginia", "Silicon Valley",
                      "San Antonio", "Sacramento", "San Diego",
                      "Minneapolis", "Milwaukee", "Indianapolis","Cincinnati",
                      "Cleveland","Pittsburgh","Baltimore","Philadelphia",
                      "New York City","Brooklyn","Bronx","Queens","Manhattan",
                      "Hartford","Providence","Albany","Buffalo","Rochester",
                      "New Orleans","Memphis","Louisville","Richmond",
                      "Raleigh","Durham","Greensboro","Winston-Salem",
                      "Birmingham","Montgomery","Huntsville","Mobile",
                      "Jacksonville","Tampa","Orlando","Fort Lauderdale",
                      "West Palm Beach","Tallahassee","Gainesville",
                      "Omaha","Lincoln","Kansas City","Wichita","Tulsa",
                      "Oklahoma City","Little Rock","Fort Smith",
                      "Salt Lake City","Provo","Ogden","Lehi","Draper",
                      "Boulder","Fort Collins","Colorado Springs",
                      "Albuquerque","Tucson","Mesa","Tempe","Scottsdale",
                      "Chandler","Gilbert","Goodyear","Peoria","Glendale AZ",
                      "Boise","Nampa","Caldwell","Idaho Falls",
                      "Helena","Billings","Great Falls","Missoula",
                      "Cheyenne","Casper","Laramie","Sheridan",
                      "Sioux Falls","Fargo","Bismarck","Grand Forks",
                      "Cedar Rapids","Des Moines","Davenport","Ames",
                      "Green Bay","Madison","Eau Claire","Appleton",
                      "Detroit","Grand Rapids","Ann Arbor","Lansing","Flint",
                      "Anchorage","Juneau","Fairbanks","Honolulu",
                      # Key DC hub / county markets
                      "Loudoun County","Prince William County","Fairfax County",
                      "Hillsboro","Boardman","Prineville","Hermiston",
                      "Goodyear","Buckeye","Coolidge","Eloy","Casa Grande",
                      "Midlothian","Garland","Plano","Irving","Carrollton",
                      "Round Rock","Georgetown","Pflugerville","Kyle",
                      "Dublin OH","Plain City","New Albany OH","Lewis Center",
                      "Reno-Sparks","Fernley","Storey County","Lyon County",
                      "Quincy WA","East Wenatchee","Moses Lake",
                      "DeKalb County","Carroll County","Paulding County",
                      "Bartow County","Newton County","Henry County",
                      "Richmond VA","Henrico County","Chesterfield County",
                      "Stafford County","Fredericksburg",
                      "Elk Grove Village","Aurora IL","Bolingbrook",
                      "Council Bluffs","Papillion","La Vista",
                      "Henderson NV","North Las Vegas","Sparks NV",
                      "Lehi UT","Draper UT","Bluffdale","South Jordan",
                      "Englewood CO","Aurora CO","Broomfield","Longmont",
                      # Regions
                      "Silicon Slopes","Silicon Hills","Silicon Desert",
                      "Silicon Plains","Research Triangle","Northern Virginia",
                      "NoVA","Pacific Northwest","Mountain West",
                      "Great Plains","Midwest","Southeast","Sun Belt",
                      "Gulf Coast","Mid-Atlantic","New England"],
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


# ── States / Provinces per country ───────────────────────────────────────────
COUNTRY_STATES = {
    "United States": [
        "Alabama","Alaska","Arizona","Arkansas","California","Colorado","Connecticut",
        "Delaware","Florida","Georgia","Hawaii","Idaho","Illinois","Indiana","Iowa",
        "Kansas","Kentucky","Louisiana","Maine","Maryland","Massachusetts","Michigan",
        "Minnesota","Mississippi","Missouri","Montana","Nebraska","Nevada","New Hampshire",
        "New Jersey","New Mexico","New York","North Carolina","North Dakota","Ohio",
        "Oklahoma","Oregon","Pennsylvania","Rhode Island","South Carolina","South Dakota",
        "Tennessee","Texas","Utah","Vermont","Virginia","Washington","West Virginia",
        "Wisconsin","Wyoming","Washington D.C.","Northern Virginia",
    ],
    "Canada": [
        "Alberta","British Columbia","Manitoba","New Brunswick","Newfoundland and Labrador",
        "Northwest Territories","Nova Scotia","Nunavut","Ontario","Prince Edward Island",
        "Quebec","Saskatchewan","Yukon",
    ],
    "Australia": [
        "New South Wales","Victoria","Queensland","Western Australia","South Australia",
        "Tasmania","Australian Capital Territory","Northern Territory",
    ],
    "India": [
        "Andhra Pradesh","Arunachal Pradesh","Assam","Bihar","Chhattisgarh","Goa","Gujarat",
        "Haryana","Himachal Pradesh","Jharkhand","Karnataka","Kerala","Madhya Pradesh",
        "Maharashtra","Manipur","Meghalaya","Mizoram","Nagaland","Odisha","Punjab",
        "Rajasthan","Sikkim","Tamil Nadu","Telangana","Tripura","Uttar Pradesh","Uttarakhand",
        "West Bengal","Delhi","Mumbai","Bangalore","Chennai","Hyderabad","Pune","Noida",
    ],
    "Germany": [
        "Baden-Württemberg","Bavaria","Berlin","Brandenburg","Bremen","Hamburg","Hesse",
        "Lower Saxony","Mecklenburg-Vorpommern","North Rhine-Westphalia","Rhineland-Palatinate",
        "Saarland","Saxony","Saxony-Anhalt","Schleswig-Holstein","Thuringia","Frankfurt",
    ],
    "United Kingdom": [
        "England","Scotland","Wales","Northern Ireland","London","Manchester","Birmingham",
        "Bristol","Edinburgh","Glasgow","Leeds","Liverpool","Sheffield","Newcastle",
    ],
    "Brazil": [
        "São Paulo","Rio de Janeiro","Minas Gerais","Bahia","Paraná","Rio Grande do Sul",
        "Pernambuco","Ceará","Amazonas","Goiás","Pará","Espírito Santo","Santa Catarina",
        "Mato Grosso","Mato Grosso do Sul","Maranhão","Paraíba","Piauí","Alagoas",
        "Rio Grande do Norte","Tocantins","Sergipe","Rondônia","Amapá","Roraima","Acre",
        "Distrito Federal",
    ],
    "China": [
        "Beijing","Shanghai","Guangdong","Zhejiang","Jiangsu","Shandong","Sichuan","Hubei",
        "Hunan","Fujian","Anhui","Shaanxi","Jiangxi","Chongqing","Liaoning","Yunnan",
        "Hebei","Guangxi","Shanxi","Tianjin","Heilongjiang","Jilin","Guizhou","Xinjiang",
        "Inner Mongolia","Gansu","Hainan","Ningxia","Qinghai","Xizang (Tibet)","Shenzhen","Hangzhou",
    ],
    "Mexico": [
        "Mexico City","Jalisco","Nuevo León","Veracruz","Puebla","Guanajuato","Chihuahua",
        "Michoacán","Oaxaca","Guerrero","Chiapas","Sinaloa","Hidalgo","Sonora","Tabasco",
        "Baja California","Querétaro","Morelos","San Luis Potosí","Aguascalientes",
        "Durango","Coahuila","Tamaulipas","Colima","Nayarit","Tlaxcala","Campeche","Yucatán","Zacatecas",
    ],
    "Malaysia": ["Johor","Kedah","Kelantan","Kuala Lumpur","Labuan","Melaka","Negeri Sembilan",
                 "Pahang","Penang","Perak","Perlis","Putrajaya","Sabah","Sarawak","Selangor","Terengganu"],
    "France": ["Île-de-France","Provence-Alpes-Côte d'Azur","Auvergne-Rhône-Alpes","Occitanie",
               "Hauts-de-France","Nouvelle-Aquitaine","Grand Est","Pays de la Loire","Normandie",
               "Bretagne","Bourgogne-Franche-Comté","Centre-Val de Loire","Corse","Paris","Lyon","Marseille"],
    "Netherlands": ["North Holland","South Holland","Utrecht","Gelderland","North Brabant",
                    "Overijssel","Groningen","Friesland","Limburg","Drenthe","Flevoland","Zeeland",
                    "Amsterdam","Rotterdam","The Hague","Eindhoven"],
    "Japan": ["Tokyo","Osaka","Kanagawa","Aichi","Saitama","Chiba","Hyogo","Hokkaido","Fukuoka",
              "Shizuoka","Ibaraki","Hiroshima","Kyoto","Miyagi","Niigata","Nagano","Tochigi",
              "Gunma","Okayama","Fukushima","Kumamoto","Kagoshima","Ehime","Yamaguchi","Okinawa"],
    "South Korea": ["Seoul","Busan","Incheon","Daegu","Daejeon","Gwangju","Ulsan","Gyeonggi",
                    "Gangwon","North Chungcheong","South Chungcheong","North Jeolla","South Jeolla",
                    "North Gyeongsang","South Gyeongsang","Jeju"],
    "Saudi Arabia": ["Riyadh","Mecca","Madinah","Eastern Province","Asir","Tabuk","Qassim",
                     "Hail","Jizan","Najran","Al Bahah","Northern Borders","Al Jawf","NEOM"],
    "UAE": ["Abu Dhabi","Dubai","Sharjah","Ajman","Umm Al Quwain","Ras Al Khaimah","Fujairah"],
    "Indonesia": ["Jakarta","West Java","East Java","Central Java","Banten","North Sumatra",
                  "South Sulawesi","East Kalimantan","Bali","Riau","South Sumatra","Yogyakarta"],
    "Poland": ["Masovian","Lesser Poland","Silesian","Greater Poland","Łódź","Lower Silesian",
               "Kuyavian-Pomeranian","Lublin","Subcarpathian","Warmian-Masurian","Pomeranian",
               "West Pomeranian","Opole","Lubusz","Świętokrzyskie","Podlaskie","Warsaw","Krakow"],
    "Ireland": ["Dublin","Cork","Galway","Limerick","Waterford","Kildare","Meath","Wicklow","Louth"],
    "Sweden": ["Stockholm","Västra Götaland","Skåne","Östergötland","Uppsala","Dalarna","Jönköping"],
    "Singapore": ["Central Region","North Region","North-East Region","East Region","West Region"],
    "South Africa": ["Gauteng","Western Cape","KwaZulu-Natal","Eastern Cape","Limpopo",
                     "Mpumalanga","North West","Free State","Northern Cape",
                     "Johannesburg","Cape Town","Durban","Pretoria"],
    "Nigeria": ["Lagos","Kano","Rivers","Kaduna","Oyo","Anambra","Abuja","Delta","Edo","Kwara"],
    "Kenya": ["Nairobi","Mombasa","Kisumu","Nakuru","Eldoret","Kiambu","Machakos"],
    "Egypt": ["Cairo","Alexandria","Giza","Luxor","Aswan","Port Said","Suez","Hurghada","Sharm el-Sheikh"],
    "Israel": ["Tel Aviv","Jerusalem","Haifa","Rishon LeZion","Petah Tikva","Ashdod","Beer Sheva","Netanya"],
    "Qatar": ["Doha","Al Wakrah","Al Khor","Lusail","Al Rayyan"],
    "Bahrain": ["Capital","Muharraq","Northern","Southern"],
    "Kuwait": ["Capital","Hawalli","Ahmadi","Jahra","Farwaniya","Mubarak Al-Kabeer"],
    "Oman": ["Muscat","Dhofar","Musandam","Al Batinah","Al Dhahirah","Al Dakhiliyah","Al Sharqiyah","Al Wusta","Al Buraimi"],
    "Chile": ["Santiago Metropolitan","Valparaíso","Biobío","Araucanía","Los Lagos","Antofagasta"],
    "Colombia": ["Bogotá","Antioquia","Valle del Cauca","Atlántico","Cundinamarca","Bolivar","Santander"],
    "Argentina": ["Buenos Aires","Córdoba","Santa Fe","Mendoza","Tucumán","Entre Ríos","Salta","Misiones"],
    "Peru": ["Lima","Arequipa","La Libertad","Piura","Lambayeque","Junín","Cusco","Áncash"],
    "Morocco": ["Casablanca-Settat","Rabat-Salé-Kénitra","Marrakech-Safi","Fès-Meknès","Tanger-Tétouan-Al Hoceïma"],
    "Rwanda": ["Kigali","Northern","Southern","Eastern","Western"],
    "Ethiopia": ["Addis Ababa","Oromia","Amhara","Tigray","SNNPR","Somali","Afar"],
    "Ghana": ["Greater Accra","Ashanti","Western","Central","Eastern","Northern","Upper East","Upper West","Volta","Brong-Ahafo"],
    "Tanzania": ["Dar es Salaam","Dodoma","Arusha","Mwanza","Mbeya","Morogoro","Tanga","Zanzibar"],
    "New Zealand": ["Auckland","Wellington","Canterbury","Waikato","Bay of Plenty","Manawatū-Whanganī","Hawke's Bay"],
    "Vietnam": ["Hanoi","Ho Chi Minh City","Hai Phong","Da Nang","Can Tho","Binh Duong"],
    "Philippines": ["Metro Manila","Cebu","Davao","Laguna","Cavite","Rizal","Bulacan","Pampanga"],
    "Thailand": ["Bangkok","Nonthaburi","Samut Prakan","Chiang Mai","Khon Kaen","Nakhon Ratchasima","Phuket"],
    "Taiwan": ["Taipei","New Taipei","Taichung","Kaohsiung","Taoyuan","Tainan","Hsinchu"],
    "Hong Kong": ["Hong Kong Island","Kowloon","New Territories","Lantau Island"],
    "Spain": ["Madrid","Catalonia","Andalusia","Valencia","Galicia","Castile and León","Basque Country","Canary Islands"],
    "Italy": ["Lombardy","Lazio","Campania","Sicily","Veneto","Emilia-Romagna","Puglia","Piedmont","Tuscany","Milan","Rome"],
    "Belgium": ["Brussels","Flanders","Wallonia","Antwerp","Ghent","Liège","Bruges"],
    "Switzerland": ["Zurich","Bern","Vaud","Geneva","Aargau","St. Gallen","Lucerne","Ticino","Valais","Basle"],
    "Austria": ["Vienna","Lower Austria","Upper Austria","Styria","Tyrol","Carinthia","Salzburg","Vorarlberg","Burgenland"],
    "Portugal": ["Lisbon","Porto","Braga","Setúbal","Aveiro","Leiria","Viseu","Coimbra"],
    "Romania": ["Bucharest","Cluj","Iași","Timișoara","Constanța","Galați","Craiova","Brașov"],
    "Czech Republic": ["Prague","Central Bohemia","South Moravian","Moravian-Silesian","Plzeň","Ústí nad Labem","Liberec","Hradec Králové"],
    "Norway": ["Oslo","Viken","Innlandet","Vestfold og Telemark","Agder","Rogaland","Vestland","Møre og Romsdal","Trøndelag","Nordland"],
    "Denmark": ["Capital","Zealand","Southern Denmark","Central Jutland","North Jutland","Copenhagen"],
    "Finland": ["Uusimaa","Pirkanmaa","Southwest Finland","North Ostrobothnia","Central Finland","South Savo","North Karelia","Helsinki"],
}

# Ensure all countries in COUNTRY_STATES are mapped in COUNTRY_TO_REGION
_EXTRA_COUNTRY_MAP = {
    "Spain": "Europe", "Italy": "Europe", "Belgium": "Europe",
    "Switzerland": "Europe", "Austria": "Europe", "Portugal": "Europe",
    "Romania": "Europe", "Czech Republic": "Europe", "Norway": "Europe",
    "Denmark": "Europe", "Finland": "Europe", "Sweden": "Europe",
    "New Zealand": "Asia Pacific", "Vietnam": "Asia Pacific",
    "Philippines": "Asia Pacific", "Thailand": "Asia Pacific",
    "Taiwan": "Asia Pacific", "Hong Kong": "Asia Pacific",
}
for _c, _r in _EXTRA_COUNTRY_MAP.items():
    COUNTRY_TO_REGION.setdefault(_c, _r)


TOPIC_KEYWORDS = {
    "Hyperscale": ["hyperscale","microsoft","google","amazon","aws","meta",
                   "apple","oracle","alibaba","tencent","bytedance",
                   "alphabet","openai","stargate","coreweave","deepmind",
                   "xai","mistral","inflection","cohere","scale ai",
                   "cerebras","groq","lambda labs","together ai",
                   "cloud region","availability zone","az deployment",
                   "hyperscaler campus","cloud campus","cloud data center"],
    "Colocation": ["colocation","colo","equinix","digital realty","ironmountain",
                   "coresite","cyrusone","ntt","vantage","switch","edgeconex",
                   "flexential","databank","cts","global switch",
                   "qts","stack infrastructure","aligned data centers",
                   "compass datacenters","tract","t5 data centers",
                   "skybox datacenters","stream data centers","cloudflare",
                   "zayo","lumen","cogent","akamai","fastly",
                   "ascenty","odata","scala","interxion","telehouse",
                   "yondr","venari","verne global","atNorth","digiPlex",
                   "bulk infrastructure","green mountain","nextdc",
                   "ctrl s","nxtgen","yotta","stt gdc","keppel dc",
                   "retail colocation","wholesale colocation","multi-tenant",
                   "single-tenant","powered shell","build to suit"],
    "AI / GPU":   ["artificial intelligence"," ai ","gpu","nvidia","inference",
                   "llm","generative","machine learning","deep learning",
                   "h100","h200","gb200","b200","b100","gh200","xai","anthropic","grok",
                   "chatgpt","claude","gemini","llama","mixtral","falcon",
                   "ai campus","ai factory","gpu cluster","gpu farm",
                   "ai infrastructure","inference cluster","training cluster",
                   "compute cluster","hpc facility","exascale","petaflop",
                   "ai supercomputer","ai accelerator","tpu","dpu","ipu",
                   "nvl rack","dgx","hgx","mgx","nvidia grace hopper",
                   "amd instinct","intel gaudi","liquid cooled gpu",
                   "high density compute","400kw rack","600kw rack",
                   "power density","kw per rack","rack density"],
    "Power":      [" mw "," gw ","megawatt","gigawatt","power plant",
                   "nuclear","solar","wind farm","gas turbine","behind-the-meter",
                   "grid","utility","energy","ppa","renewable","hydrogen",
                   "fuel cell","battery storage","ups","smr","small modular reactor",
                   "geothermal","hydropower","tidal","pumped hydro",
                   "battery energy storage system","bess","vrfb",
                   "grid connection","substation","transformer","switchgear",
                   "high voltage","extra high voltage","transmission line",
                   "distribution grid","feeder","bus bar","load growth",
                   "interconnection queue","capacity market","energy market",
                   "pjm","ercot","caiso","miso","iso-ne","nyiso","spp",
                   "wecc","serc","nerc","national grid eso","eirgrid","aemo",
                   "dewa","taqa","sec saudi","powergrid india","kepco",
                   "demand response","virtual power plant","microGrid",
                   "behind meter","net metering","self-consumption",
                   "power purchase agreement","offtake agreement","capacity contract",
                   "vppa","virtual ppa","24/7 cfe","carbon free electricity",
                   "guaranteed maximum demand","critical load"],
    "Investment": ["invest","fund","reit","billion","million","acquire",
                   "acquisition","ipo","bond","financing","lease","deal",
                   "partnership","joint venture","stake","raise","capital",
                   "sovereign wealth fund","pension fund","infrastructure fund",
                   "private equity","private credit","mezzanine","preferred equity",
                   "green bond","sustainability bond","project finance",
                   "sale leaseback","forward purchase","forward funding",
                   "net lease","triple net","nnn lease","long term lease",
                   "sale and leaseback","co-investment","club deal",
                   "portfolio acquisition","platform acquisition","roll-up",
                   "tender offer","take private","going private",
                   "credit facility","revolving credit","term loan",
                   "cad","aud","nzd","chf","sek","nok","dkk","krw","myr",
                   "thb","idr","sgd","aed","sar","qar","inr","brl","mxn",
                   "trillion","crore","lakh","bn usd","mn usd","bn eur",
                   "mn gbp","bn gbp","mn sgd","bn inr","mn aed"],
    "Permits":    ["permit","zoning","moratorium","approved","approval",
                   "planning","ordinance","rezoning","denied","appeal",
                   "lawsuit","sue","court","commission",
                   "vote","hearing","rejected","conditional use",
                   "use permit","special use permit","sup","cup",
                   "environmental review","eis","nepa","ceqa","ea",
                   "site plan approval","subdivision","variance",
                   "setback","noise ordinance","water permit",
                   "utility easement","right of way","condemnation",
                   "eminent domain","community benefit agreement","cba",
                   "public comment","planning commission","city council",
                   "county board","supervisor vote","board of supervisors",
                   "planning board","zoning board","appeal board",
                   "building permit","demolition permit","grading permit",
                   "environmental permit","stormwater permit"],
    "Construction":["broke ground","groundbreaking","topping out","opens",
                    "opened","inaugurated","construction","campus","build",
                    "development","site","facility","phase","expansion",
                    "warehouse","acres","sq ft","square feet","hectares",
                    "sq m","square meters","under construction","build-out",
                    "fit-out","fitout","commissioning","energized","energise",
                    "shell and core","raised floor","hot aisle containment",
                    "cold aisle containment","modular build","prefab",
                    "containerized","pod deployment","construction loan",
                    "civil works","structural steel","mep works",
                    "turnkey","epc","design-build","leed","breeam",
                    "certificate of occupancy","temporary certificate",
                    "ribbon cutting","topping ceremony","groundbreak",
                    "site preparation","land clearing","grading",
                    "concrete pour","steel erection","roof completion",
                    "phase 1","phase 2","phase 3","initial phase",
                    "first phase","next phase","final phase","full build-out"],
    "Sustainability":["sustainability","carbon","renewable","net zero",
                      "green","esg","water","pue","cooling","waste heat",
                      "recycle","circular","biodiversity","solar panel",
                      "wind power","offset","scope 1","scope 2","scope 3",
                      "carbon neutral","carbon negative","carbon credit",
                      "rec","i-rec","guarantees of origin","24/7 cfe",
                      "water positive","water neutral","zero waste",
                      "heat reuse","district heating","sbti",
                      "science based targets","tcfd","esg report",
                      "esg disclosure","climate commitment","climate target",
                      "decarbonization","low carbon","clean energy",
                      "renewable portfolio","green tariff","green procurement",
                      "social impact","community benefit","local hiring",
                      "noise mitigation","light pollution","visual impact"],
    "Grid / ISO / RTO": ["pjm","ercot","caiso","miso","nyiso","iso-ne","spp","wecc",
                          "serc","nerc","tva","bpa","national grid eso","eirgrid","aemo",
                          "powergrid india","kepco","dewa","kahramaa","taqa","entso-e",
                          "rte france","tennet","50hertz","amprion","elia","fingrid",
                          "statnett","energinet","pse poland","transpower","tnb","pln",
                          "tepco","ema singapore","state grid china","aneel","cenace",
                          "eskom","large load study","grid interconnection",
                          "interconnection queue","load growth","transmission capacity",
                          "grid capacity","balancing authority","load zone",
                          "capacity market","energy market","grid operator",
                          "independent system operator","regional transmission organization"],
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
    # Extended hyperscalers / AI cos
    "OpenAI","xAI","DeepMind","Mistral","Cohere","Inflection","Scale AI",
    "Cerebras","Groq","Lambda Labs","Together AI","CoreWeave","Vast AI",
    "Fireworks AI","Perplexity","Stability AI","Midjourney",
    # Extended colocation & operators
    "QTS Realty","Stack Infrastructure","Aligned Data Centers",
    "Compass Datacenters","Tract","T5 Data Centers","Skybox Datacenters",
    "Stream Data Centers","DataBank","CloudHQ","Chirisa","Cologix",
    "Raging Wire","Switch SUPERNAP","Cyxtera","Peak 10",
    "365 Main","Digital Bridge","DigitalBridge","Stonepeak",
    "IPI Partners","Harrison Street","Actis","Oaktree",
    "Singtel","Telstra","Spark NZ","KDDI","SoftBank",
    "Adani ConnectX","CtrlS","NxtGen","Yotta","STT GDC",
    "Keppel","NextDC","Macquarie","AirTrunk",
    "MEVSPACE","Beyond.pl","Atman","Exatel","LCloud",
    "Nabiax","Bolder","VIRTUS","Kao Data","Yondr","Venari",
    "Verne Global","Hydro66","DigiPlex","Green Mountain",
    "atNorth","Bulk Infrastructure","DigiCo","Data#3",
    "Ascenty","Odata","Scala","Luminet","Etix","Global Data Systems",
    # Power / utilities
    "Dominion Energy","Duke Energy","Entergy","Exelon","AEP",
    "Eversource","National Grid","Ameren","Evergy","PPL",
    "CenterPoint Energy","Pacific Gas and Electric","PG&E",
    "Southern Company","NextEra Energy","AES","NV Energy",
    "Sacramento Municipal Utility","SMUD","Salt River Project",
    "Arizona Public Service","APS","Portland General Electric",
    "Pacific Power","Puget Sound Energy","Idaho Power",
    "Basin Electric","Dairyland Power","Xcel Energy",
    "OGE Energy","Empire District","CLECO","Entergy Texas",
    "Tokyo Electric","TEPCO","KEPCO Korea","Kansai Electric",
    "Chubu Electric","Kyushu Electric","Tohoku Electric",
    "Tenaga Nasional","PLN Indonesia","EirGrid","National Grid ESO",
    "RTE France","TenneT","50Hertz","Amprion","TransnetBW",
    "Elia","Terna","REE Spain","Fingrid","Statnett","Energinet",
    "Svenska Kraftnät","PSE Poland","Swissgrid","APG",
    "DEWA","TAQA","SEC Saudi","KAHRAMAA","Eskom",
    "AEMO","Transpower NZ","PowerGrid India",
    # Construction / engineering / equipment
    "Mortenson","Turner Construction","Holder","DPR","Skanska",
    "Gilbane","McCarthy","Whiting-Turner","Structure Tone",
    "Black & Veatch","Burns & McDonnell","AECOM","Jacobs",
    "HDR","Gensler","HMC Architects","DLR Group",
    "Stantec","WSP","Arup","Buro Happold",
    "Vertiv","Eaton","ABB","Siemens Energy","Schneider Electric",
    "Caterpillar","Cummins","Aggreko","HITEC","AMETEK",
    "Legrand","Panduit","CommScope","Corning","Belden",
    "Rittal","nVent","Emerson Network Power","Liebert",
    "Alfa Laval","SPX Cooling","Baltimore Aircoil","Evapco",
    "Stulz","Airedale","Aermec","Daikin","Carrier","Trane",
    # Investors / developers
    "Blackstone","KKR","Brookfield","GIC","Mubadala","CPPIB",
    "SoftBank","Masdar","TAQA","Neom","SABIC","ADNOC",
    "Coatue","Andreessen Horowitz","a16z","Tiger Global",
    "SilverLake","Apollo","Cerberus","Carlyle","Warburg Pincus",
    "Stonepeak","EQT","IFM Investors","Macquarie Asset Management",
    "GLP","Prologis","Highbrook","Hana Financial",
    "AustralianSuper","CPP Investments","Ontario Teachers",
    "OMERS","PSP Investments","Caisse de dépôt","CDPQ",
    "Abu Dhabi Investment Authority","ADIA","PIF Saudi",
    "Temasek","GIC Singapore","Khazanah","EPF Malaysia",
]


# ─── DCD URL constants ─────────────────────────────────────────────────────
DCD_BASE             = "https://www.datacenterdynamics.com"
# Construction channel: DCD's own filtered term-based URL (proven to work)
DCD_CONSTRUCTION_URL = DCD_BASE + "/en/news/?term=the-data-center-construction-channel"
# General news: the plain news listing
DCD_GENERAL_URL      = DCD_BASE + "/en/news/"

# ─── News-type labels (used in sidebar filter) ─────────────────────────────
NEWS_TYPE_CONSTRUCTION = "Construction"
NEWS_TYPE_GENERAL      = "General News"

# REMOVED: DCD_CHAN_TERM, DCD_REGION_TERMS, DCD_EXTRA_TERMS, GNEWS_QUERIES, RSS_SOURCES
# The app now scrapes only the two DCD channels above.

_GNEWS_QUERIES_REMOVED = [
    # Construction / physical build
    ("data center construction campus groundbreaking opening", "Google News"),
    ("data center broke ground topping out opens ribbon cutting", "Google News"),
    ("data center under construction building phase expansion", "Google News"),
    # Approvals / permits / regulatory
    ("data center approved approval permit planning zoning", "Google News"),
    ("data center zoning rezoning moratorium ordinance vote hearing", "Google News"),
    ("data center project approval go-ahead green light county city", "Google News"),
    ("data center rejected denied blocked lawsuit opposition", "Google News"),
    # Site selection / disclosed / announced
    ("data center site selection disclosed announced plans", "Google News"),
    ("data center project announced new campus planned", "Google News"),
    ("data center disclosed project pipeline plans filed", "Google News"),
    # Extension / expansion
    ("data center expansion extension phase additional capacity", "Google News"),
    # Power & energy
    ("data center hyperscale investment billion megawatt gigawatt", "Google News"),
    ("data center power energy grid nuclear solar PPA", "Google News"),
    ("data center behind the meter power plant generator turbine", "Google News"),
    ("data center nuclear SMR geothermal hydrogen power", "Google News"),
    ("data center grid connection electricity capacity substation", "Google News"),
    # Investment / finance - currency-specific
    ("data center investment billion million acquisition deal", "Google News"),
    ("data center acquisition merger deal sale billion USD", "Google News"),
    ("data center REIT investment fund financing lease", "Google News"),
    ("data center IPO equity raise capital raise funding", "Google News"),
    ("data center EUR billion investment Europe campus", "Google News"),
    ("data center GBP billion investment UK campus", "Google News"),
    ("data center AED billion investment UAE Middle East", "Google News"),
    ("data center SGD billion investment Singapore Asia", "Google News"),
    ("data center INR crore billion investment India campus", "Google News"),
    # Acres / land / size signals
    ("data center acres land site development campus", "Google News"),
    ("data center hectares land parcel build-to-suit site", "Google News"),
    ("data center sq ft square feet campus facility", "Google News"),
    ("data center powered shell wholesale campus acres", "Google News"),
    # Deal flow - undisclosed / private
    ("data center joint venture partnership MOU offtake agreement", "Google News"),
    ("data center letter of intent pre-lease capacity agreement", "Google News"),
    ("data center sale leaseback forward purchase transaction", "Google News"),
    ("data center private equity infrastructure fund stake", "Google News"),
    ("data center M&A takeover acquisition stake minority", "Google News"),
    # Hyperscalers
    ("Microsoft Google Amazon Meta Oracle data center campus", "Google News"),
    ("AWS Azure GCP hyperscale cloud data center region", "Google News"),
    # Operators
    ("Equinix Digital Realty CyrusOne QTS NTT data center", "Google News"),
    ("EdgeConneX Vantage Compass Aligned DataBank data center", "Google News"),
    ("Yondr AirTrunk NextDC Macquarie atNorth data center", "Google News"),
    ("colocation datacenter AI GPU facility opens launched", "Google News"),
    # AI/GPU
    ("AI data center GPU campus hyperscale construction", "Google News"),
    ("AI factory inference training facility data center campus", "Google News"),
    # Regions
    ("data center Middle East Africa Asia Pacific expansion", "Google News"),
    ("data center Europe Germany Netherlands Ireland Frankfurt", "Google News"),
    ("data center India Singapore Malaysia Southeast Asia", "Google News"),
    ("data center Latin America Brazil Mexico Chile Argentina", "Google News"),
    # Capacity / MW / GW specific
    ("data center MW GW megawatt gigawatt capacity announcement", "Google News"),
    ("data center 100MW 200MW 500MW 1GW campus facility", "Google News"),
    ("data center kilowatt density high performance compute", "Google News"),
    # === ISO / RTO / Large Load / Grid Operator Queries ===
    ("PJM data center large load interconnection queue megawatt gigawatt", "Google News"),
    ("ERCOT data center large load study Texas megawatt queue gigawatt", "Google News"),
    ("CAISO data center interconnection California megawatt grid queue", "Google News"),
    ("MISO data center large load interconnection Midwest megawatt", "Google News"),
    ("NYISO data center interconnection New York megawatt queue", "Google News"),
    ("ISO-NE data center interconnection New England megawatt", "Google News"),
    ("SPP data center interconnection queue megawatt Southwest Plains", "Google News"),
    ("Northern Virginia data center PJM power grid megawatt queue Dominion", "Google News"),
    ("Texas ERCOT data center load growth power demand gigawatt capacity", "Google News"),
    ("Ohio PJM data center large load electricity demand megawatt", "Google News"),
    ("Georgia Southern Company data center power load electricity demand", "Google News"),
    ("Arizona APS SRP WECC data center power load capacity megawatt", "Google News"),
    ("Oregon BPA Pacific Power data center Northwest power interconnection", "Google News"),
    ("SERC Southeast data center electricity power load demand", "Google News"),
    ("TVA Tennessee Valley Authority data center large power rate", "Google News"),
    ("EirGrid Ireland data center moratorium electricity grid capacity", "Google News"),
    ("National Grid UK data center connection queue megawatt TEC holder", "Google News"),
    ("AEMO Australia data center large load connection megawatt grid", "Google News"),
    ("PowerGrid India data center electricity connection load MW", "Google News"),
    ("EMA Singapore data center electricity moratorium load MW", "Google News"),
    ("DEWA UAE data center power electricity megawatt Dubai connection", "Google News"),
    ("ENTSO-E Europe data center electricity grid demand MW capacity", "Google News"),
    ("KEPCO South Korea data center electricity grid connection MW", "Google News"),
    ("TNB Tenaga Malaysia data center electricity grid connection Johor", "Google News"),
    ("State Grid China data center electricity connection load MW", "Google News"),
    ("AEMO NEM Australia data center load growth electricity grid", "Google News"),
    ("Eskom South Africa data center electricity connection load MW", "Google News"),
    ("TEPCO Japan data center electricity grid connection MW campus", "Google News"),
    ("ANEEL Brazil data center electricity grid connection MW campus", "Google News"),
    ("CENACE Mexico data center electricity grid connection MW campus", "Google News"),
    # Large load / grid connection generic
    ("data center large load interconnection request utility grid", "Google News"),
    ("data center utility scale electricity demand grid operator ISO RTO", "Google News"),
    ("data center grid connection queue delay backlog utility", "Google News"),
    ("data center critical load electricity infrastructure utility", "Google News"),
    ("data center electricity demand forecast megawatt gigawatt load growth", "Google News"),
    ("hyperscale data center power load grid utility PPA offtake", "Google News"),
    ("data center behind meter generation electricity self-generation", "Google News"),
    ("data center nuclear SMR small modular reactor power offtake PPA", "Google News"),
    ("data center substation transformer upgrade utility connection MW", "Google News"),
    ("data center grid capacity constraint electricity supply bottleneck", "Google News"),
    ("data center power availability electricity grid interconnection study", "Google News"),
    ("data center 100MW 200MW 300MW 400MW 500MW 600MW 700MW 800MW 900MW 1GW", "Google News"),
    ("data center gigawatt campus AI factory electricity load demand", "Google News"),
    # US county/metro markets
    ("Loudoun County Prince William Fairfax data center Virginia MW", "Google News"),
    ("Hillsboro Boardman Prineville Oregon data center campus", "Google News"),
    ("Goodyear Mesa Chandler Avondale Buckeye Arizona data center campus", "Google News"),
    ("Midlothian Fort Worth Garland Plano Dallas data center Texas", "Google News"),
    ("Columbus Dublin Plain City New Albany Ohio data center campus", "Google News"),
    ("Reno Sparks Fernley Nevada data center campus megawatt", "Google News"),
    ("Quincy East Wenatchee Moses Lake Washington data center MW", "Google News"),
    ("DeKalb Carroll Paulding Bartow Georgia data center county MW", "Google News"),
    ("San Antonio Austin Round Rock Georgetown Texas data center campus", "Google News"),
    ("Chicago Elk Grove Aurora Joliet Illinois data center campus MISO", "Google News"),
    ("Omaha Council Bluffs Iowa Nebraska data center campus SPP MW", "Google News"),
    ("Richmond Henrico Chesterfield Virginia data center PJM campus", "Google News"),
    ("Boise Nampa Caldwell Idaho data center campus megawatt", "Google News"),
    ("Phoenix Tempe Scottsdale Glendale Arizona data center campus", "Google News"),
    ("Salt Lake City Lehi Draper Utah data center campus megawatt", "Google News"),
    ("Denver Aurora Centennial Colorado data center campus megawatt", "Google News"),
    ("Charlotte Concord Mooresville North Carolina data center campus", "Google News"),
    ("Atlanta Douglasville Newnan Coweta Georgia data center campus", "Google News"),
    ("Minneapolis Saint Paul Eden Prairie Minnesota data center campus", "Google News"),
    ("Kansas City Lee's Summit Lenexa Missouri Kansas data center", "Google News"),
    ("Detroit Grand Rapids Ann Arbor Michigan data center campus", "Google News"),
    ("Indianapolis Carmel Brownsburg Indiana data center campus MISO", "Google News"),
    ("Birmingham Huntsville Montgomery Alabama data center campus", "Google News"),
    ("Jacksonville Tampa Orlando Miami Florida data center campus", "Google News"),
    ("Las Vegas Henderson North Las Vegas Nevada data center campus", "Google News"),
    # International metro markets
    ("Frankfurt Rhine-Main Germany data center campus megawatt TenneT", "Google News"),
    ("Amsterdam Netherlands data center campus AMS megawatt TenneT", "Google News"),
    ("Dublin Kildare Meath Ireland data center moratorium campus megawatt", "Google News"),
    ("London Slough Reading Swindon UK data center campus megawatt National Grid", "Google News"),
    ("Singapore JTC Jurong data center campus megawatt EMA", "Google News"),
    ("Johor Iskandar Kulai Malaysia data center campus megawatt TNB", "Google News"),
    ("Jakarta Batam Surabaya Indonesia data center campus megawatt PLN", "Google News"),
    ("Mumbai Navi Mumbai Pune Hyderabad India data center campus megawatt", "Google News"),
    ("Sydney Melbourne Brisbane Perth Australia data center AEMO megawatt", "Google News"),
    ("Tokyo Osaka Nagoya Japan data center campus megawatt TEPCO", "Google News"),
    ("Seoul Incheon Busan South Korea data center KEPCO megawatt campus", "Google News"),
    ("Riyadh Jeddah NEOM Yanbu Saudi Arabia data center campus megawatt", "Google News"),
    ("Dubai Abu Dhabi Sharjah UAE DEWA data center campus megawatt", "Google News"),
    ("São Paulo Rio Curitiba Brazil data center campus megawatt ANEEL", "Google News"),
    ("Warsaw Katowice Wroclaw Poznan Poland data center campus megawatt", "Google News"),
    ("Stockholm Gothenburg Malmö Sweden data center campus megawatt", "Google News"),
    ("Helsinki Tampere Espoo Finland data center campus megawatt Fingrid", "Google News"),
    ("Oslo Bergen Stavanger Norway data center campus megawatt Statnett", "Google News"),
    ("Copenhagen Aarhus Denmark data center campus megawatt Energinet", "Google News"),
    ("Barcelona Madrid Valencia Spain data center campus megawatt REE", "Google News"),
    ("Milan Rome Turin Italy data center campus megawatt Terna", "Google News"),
    ("Brussels Antwerp Ghent Belgium data center campus megawatt Elia", "Google News"),
    ("Vienna Graz Linz Austria data center campus megawatt APG", "Google News"),
    ("Prague Brno Ostrava Czech Republic data center campus megawatt", "Google News"),
    ("Bucharest Cluj Timisoara Romania data center campus megawatt", "Google News"),
    ("Nairobi Mombasa Kenya data center campus megawatt", "Google News"),
    ("Lagos Abuja Nigeria data center campus megawatt", "Google News"),
    ("Johannesburg Cape Town Durban South Africa data center campus megawatt", "Google News"),
    ("Casablanca Rabat Morocco data center campus megawatt", "Google News"),
    ("Cairo Alexandria Egypt data center campus megawatt", "Google News"),
    ("Doha Lusail Qatar data center campus megawatt KAHRAMAA", "Google News"),
    ("Manama Bahrain data center campus megawatt", "Google News"),
    ("Muscat Oman data center campus megawatt", "Google News"),
    ("Kigali Rwanda data center campus megawatt AfricaConnect", "Google News"),
    ("Addis Ababa Ethiopia data center campus megawatt", "Google News"),
    ("Accra Ghana data center campus megawatt", "Google News"),
    ("Santiago Chile data center campus megawatt", "Google News"),
    ("Bogota Medellin Colombia data center campus megawatt", "Google News"),
    ("Buenos Aires Cordoba Argentina data center campus megawatt", "Google News"),
    ("Lima Peru data center campus megawatt", "Google News"),
    ("Mexico City Monterrey Guadalajara data center campus megawatt CENACE", "Google News"),
    # Extra keyword-rich queries  
    ("data center AI factory GPU cluster campus announcement investment", "Google News"),
    ("data center campus acre land acquisition site control option", "Google News"),
    ("data center lease pre-lease capacity agreement megawatt announcement", "Google News"),
    ("data center REIT dividend infrastructure fund deployment capital", "Google News"),
    ("data center IPO listing secondary market infrastructure equity raise", "Google News"),
]   # end of _GNEWS_QUERIES_REMOVED — not used

# ─── Date parser (app15 style — fast and clean) ────────────────────────────
MONTHS = {
    "jan":1,"feb":2,"mar":3,"apr":4,"may":5,"jun":6,
    "jul":7,"aug":8,"sep":9,"oct":10,"nov":11,"dec":12,
}

def parse_date_str(raw):
    if not raw:
        return None
    raw = str(raw).strip()
    raw = re.sub(r"\s+", " ", raw)
    # ISO with timezone
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


# ─── HTTP fetcher (app15 style — cloudscraper first, fallback requests) ────
_DCD_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
    "Referer": "https://www.datacenterdynamics.com/",
}

def fetch_html(url, retries=2):
    for attempt in range(retries):
        try:
            if _USE_CS:
                r = _CS.get(url, timeout=20)
            else:
                r = _CS.get(url, headers=_DCD_HEADERS, timeout=20)
            r.raise_for_status()
            return BeautifulSoup(r.text, "html.parser")
        except Exception:
            if attempt == 0:
                time.sleep(2)
    return None


# ─── Article parser — DCD HTML pages ──────────────────────────────────────
def _parse_articles_from_soup(soup, source_name, base_url):
    """
    Parses DCD listing pages.  Finds all <a href="/en/(news|analysis|opinion|
    dcd-data-center-construction-channel-news)/..."> links, extracts headline
    from h1–h5 inside the link, then climbs the DOM for a date.
    Works for both the general news and construction-channel URLs.
    """
    articles = []
    seen = set()

    for a in soup.find_all("a", href=re.compile(
        r"^/en/(news|analysis|opinion|dcd-data-center-construction-channel-news)/[^?#]+/$"
    )):
        href = a["href"]
        if href in seen:
            continue
        seen.add(href)

        h_tag    = a.find(["h1", "h2", "h3", "h4", "h5"])
        headline = h_tag.get_text(" ", strip=True) if h_tag else a.get_text(" ", strip=True)
        headline = re.sub(r"\s+", " ", headline).strip()
        if not headline or len(headline) < 10:
            continue

        date_obj = None
        node = a.parent
        for _ in range(12):
            if node is None:
                break
            # 1. <time datetime="...">  — most reliable on DCD
            time_tag = node.find("time")
            if time_tag:
                dt_attr = time_tag.get("datetime", "")
                if dt_attr:
                    date_obj = parse_date_str(dt_attr)
                if not date_obj:
                    date_obj = parse_date_str(time_tag.get_text(strip=True))
                if date_obj:
                    break
            txt = node.get_text(" ", strip=True)
            # 2. "dd Mon YYYY"
            m = re.search(
                r"\b(\d{1,2})\s+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+(\d{4})\b",
                txt, re.I,
            )
            if m:
                date_obj = parse_date_str(m.group(0))
                break
            # 3. "Mon dd, YYYY"
            m2 = re.search(
                r"\b(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\w*\s+(\d{1,2}),?\s+(\d{4})\b",
                txt, re.I,
            )
            if m2:
                date_obj = parse_date_str(m2.group(0))
                break
            # 4. ISO yyyy-mm-dd
            m3 = re.search(r"\b(\d{4}-\d{2}-\d{2})\b", txt)
            if m3:
                date_obj = parse_date_str(m3.group(1))
                break
            node = node.parent

        articles.append({
            "headline": headline,
            "url":      base_url + href,
            "date_obj": date_obj,
            "source":   source_name,
            "_priority": 1,
        })
    return articles


# ─── DCD scraper: scrapes a single base URL across N pages ─────────────────
def _scrape_dcd_channel(base_url, source_name, cutoff, max_pages, progress_cb, label="DCD"):
    """
    Generic DCD channel scraper.  Paginates base_url with &page=N (or ?page=N)
    until we hit an article older than cutoff or a page returns nothing new.

    base_url    : channel root — may already contain query params (e.g. ?term=…)
    source_name : label stored in article["source"]
    cutoff      : datetime — articles older than this are dropped
    max_pages   : upper page limit
    progress_cb : callable(fraction 0→1, text)
    """
    fetch_html(DCD_BASE + "/en/")   # warm-up (helps bypass Cloudflare)

    all_articles = []
    seen_urls    = set()

    # Work out whether base_url already has a query string
    _has_qs = "?" in base_url

    for page in range(1, max_pages + 1):
        if page == 1:
            url = base_url
        else:
            sep = "&" if _has_qs else "?"
            url = f"{base_url}{sep}page={page}"

        progress_cb(
            min(page / max_pages, 1.0),
            f"⚡ Scanning {label} · Page {page}/{max_pages}",
        )

        soup = fetch_html(url)
        if not soup:
            break

        page_arts  = _parse_articles_from_soup(soup, source_name, DCD_BASE)
        new_on_page = 0
        stop        = False

        for art in page_arts:
            norm_url = art["url"].rstrip("/")
            if norm_url in seen_urls:
                continue
            seen_urls.add(norm_url)
            d = art["date_obj"]
            if d and d < cutoff:
                stop = True
                break
            all_articles.append(art)
            new_on_page += 1

        if stop or new_on_page == 0:
            break

        time.sleep(0.4)

    return all_articles


# ─── run_all_scrapers: dispatches to the chosen DCD channel(s) ─────────────
def run_all_scrapers(max_html_pages, cutoff, progress_cb,
                     news_types=None, region_terms=None):
    """
    news_types : list containing any of ["Construction", "General News"].
                 Pass [] or None to scrape both channels.
    region_terms : ignored (kept for API compatibility) — filtering by
                   region/country/company is done post-scrape in main().
    """
    if not news_types:
        news_types = [NEWS_TYPE_CONSTRUCTION, NEWS_TYPE_GENERAL]

    raw = []

    # ── Construction channel ──────────────────────────────────────────────
    if NEWS_TYPE_CONSTRUCTION in news_types:
        def _cb_c(frac, text=""):
            progress_cb(frac * (0.5 if NEWS_TYPE_GENERAL in news_types else 1.0), text)

        arts = _scrape_dcd_channel(
            DCD_CONSTRUCTION_URL, "DCD Construction",
            cutoff, max_html_pages, _cb_c, label="Construction Channel",
        )
        raw.extend(arts)
        progress_cb(
            0.5 if NEWS_TYPE_GENERAL in news_types else 1.0,
            f"Construction channel: {len(arts)} articles fetched",
        )

    # ── General news channel ──────────────────────────────────────────────
    if NEWS_TYPE_GENERAL in news_types:
        base_frac = 0.5 if NEWS_TYPE_CONSTRUCTION in news_types else 0.0

        def _cb_g(frac, text=""):
            progress_cb(base_frac + frac * (1.0 - base_frac), text)

        arts = _scrape_dcd_channel(
            DCD_GENERAL_URL, "DCD General News",
            cutoff, max_html_pages, _cb_g, label="General News",
        )
        raw.extend(arts)
        progress_cb(1.0, f"General news: {len(arts)} articles fetched")

    return raw


# ─── SCRAPE_SOURCES: displayed in the banner "N Sources Active" ─────────────
SCRAPE_SOURCES = [
    {"name": "DCD Construction Channel",
     "url":  DCD_CONSTRUCTION_URL, "type": "html", "priority": 1},
    {"name": "DCD General News",
     "url":  DCD_GENERAL_URL,      "type": "html", "priority": 1},
]

def is_dc_relevant(text):
    t = text.lower()
    # Primary: any of these alone = relevant
    primary = [
        "data center", "datacenter", "data centre", "datacentre",
        "colocation", "colo ", "hyperscale", "cloud campus",
        "server farm", "computing campus", "ai campus", "gpu cluster",
        "compute campus", "hpc facility", "edge facility",
        "carrier hotel", "internet exchange", "ix facility",
        "infrastructure reit", "digital infrastructure",
        # Additional primary terms
        "data hall", "data park", "digital campus",
        "compute facility", "cloud facility", "network facility",
        "ai factory", "inference facility", "training facility",
        "wholesale data", "retail colocation", "powered shell",
        "build-to-suit", "mission critical", "critical facility",
        # ISO / RTO / Grid operators (US)
        "pjm","ercot","caiso","miso","nyiso","iso-ne","spp","wecc","serc","nerc","bpa","tva",
        "pjm interconnection","ercot queue","caiso queue","miso queue","nyiso queue",
        "spp queue","iso-ne queue","large load study","grid interconnection",
        "generation interconnection","interconnection queue","load growth study",
        "transmission capacity","grid capacity","grid constraint","power availability",
        "critical load facility","large load","bulk power","bulk electric",
        # ISO / RTO / Grid operators (Global)
        "national grid eso","eirgrid","aemo","powergrid india","kepco","tennet",
        "50hertz","amprion","elia","rete","fingrid","statnett","energinet",
        "svenska kraftnät","transpower","dewa","kahramaa","taqa","sec saudi",
        "aneel","cenace","eskom","state grid china","tnb","pln","tepco","ema singapore",
        "entso-e","rte france","transnetbw","swissgrid","apg austria","pse poland",
        "bc hydro","ieso","aeso","hydro-québec",
        # Data campus / AI infrastructure
        "ai campus","ai factory","ai infrastructure","ai data center","ai datacenter",
        "gpu cluster","gpu farm","gpu factory","inference cluster","training cluster",
        "compute cluster","hpc cluster","exascale","petaflop","tera ops",
        "smart campus","tech campus","innovation campus","cloud campus","digital park",
        "data park","digital hub","tech hub","connectivity hub","internet hub",
        "network operations center","noc facility","tier iii","tier iv","tier 3","tier 4",
        "uptime certified","uptime institute","carrier neutral",
        # Power / energy infrastructure
        "behind-the-meter","behind meter","BTM generation","self-generation",
        "small modular reactor","smr","nuclear power purchase","nuclear offtake",
        "geothermal power","fuel cell power","distributed generation",
        "microGrid","battery energy storage","bess","backup power",
        "critical power","uninterruptible power","ups system",
        "cooling system","cooling tower","chilled water","adiabatic cooling",
        "free cooling","economizer","heat exchanger","liquid cooling",
        "immersion cooling","direct liquid cooling","dlc","rear-door heat exchanger",
        "pue","wue","cue","dcu","power usage effectiveness",
        "water usage effectiveness","carbon usage effectiveness",
        # Construction / development
        "powered shell","dark fiber","conduit","fiber backbone",
        "fit out","fitout","white space","raised floor","hot aisle","cold aisle",
        "modular data","prefab module","container data","pod deployment",
        "data center shell","speculative build","spec build",
        # Financial / investment
        "infrastructure reit","data center reit","net lease","triple net",
        "sale leaseback","forward purchase","equity raise","debt raise",
        "credit facility","green bond","sustainability bond","infrastructure bond",
        "co-investment","club deal","preferred equity","mezzanine financing",
        "project finance","data center fund","digital infrastructure fund",
    ]
    # Secondary: two or more = relevant
    secondary = [
        "megawatt", " mw ", " gw ", "gigawatt",
        "computing facility", "edge computing",
        "power purchase agreement", " ppa ", "behind the meter",
        "grid connection", "critical load", "raised floor",
        "cooling tower", "liquid cooling", "immersion cooling",
        "diesel generator", "ups system", "modular data",
        "tier iii", "tier iv", "uptime institute",
        "network access point", "internet hub",
        "rack space", "co-location", "hosting facility",
        "blade server", "server deployment", "ai infrastructure",
        # Currency / financial signals
        "$", "€", "£", "¥", "₹", "billion", "million",
        "usd", "eur", "gbp", "jpy", "inr", "sgd", "aed",
        "investment", "financing", "acquisition", "deal",
        # Size / land signals
        "acres", "acre", "hectares", "hectare", "sq ft", "square feet",
        "square meters", "sq m", "campus site", "land parcel",
        # Power / capacity signals
        "kilowatt", " kw ", "kwh", "mwh", "gwh",
        "substation", "transformer", "generator set",
        "ups capacity", "power density",
        # Construction / development signals
        "breaking ground", "ground breaking", "ribbon cutting",
        "topping off", "commissioning", "fit-out", "fitout",
        "shell and core", "white space", "raised floor space",
        # Deal / financial terms
        "sale leaseback", "forward purchase", "joint venture",
        "mou", "memorandum of understanding", "loi", "letter of intent",
        "offtake agreement", "capacity agreement", "pre-lease",
        # Operator / technology signals
        "crac unit", "adiabatic cooling", "free cooling",
        "power usage effectiveness", "water usage effectiveness",
        "pue", "wue", "dcu", "noc", "soc",
        # Currency / financial signals (extended)
        "cad","aud","nzd","chf","sek","nok","dkk","krw","thb","myr","idr",
        "try","zar","ngn","kes","egp","sar","qar","bhd","kwd","omr",
        "brl","cop","clp","mxn","pen","ars",
        "trillion","crore","lakh","million dollars","billion dollars",
        "mn usd","bn usd","mn eur","bn eur","mn gbp","bn gbp",
        "mn sgd","bn sgd","mn inr","bn inr","mn aed","bn aed",
        "investment round","series a","series b","series c","growth equity",
        "institutional investor","sovereign wealth fund","swf","pension fund",
        "infrastructure equity","private infrastructure","privatization",
        # US state abbreviations relevant for DC markets
        "va.","tx.","ca.","ga.","oh.","il.","in.","nv.","az.","or.",
        "wa.","nc.","pa.","nj.","md.","fl.","co.","ut.","wy.","id.",
        "mn.","wi.","mi.","ia.","ks.","mo.","ok.","ar.","la.",
        # US county / DC hub shorthand
        "nova","northern va","loudoun","prince william","fairfax county",
        "silicon slopes","silicon hills","silicon desert","silicon plains",
        # ISO / RTO abbreviations (secondary-level match)
        "rto","iso grid","independent system operator","regional transmission",
        "transmission operator","distribution operator","grid operator",
        "interconnection study","generator interconnection","large load study",
        "load serving entity","load zone","balancing authority",
        "capacity market","energy market","ancillary services",
        "demand response","virtual power plant","grid storage",
        "4cp peak","coincident peak","non-coincident peak","load factor",
        # Capacity / size signals (extended)
        "mw it","mw critical","mw white space","mw raised floor",
        "mw total","mw phase","mwac","mwdc",
        "rack","rack unit","rack space","kw per rack","kw density","mw density",
        "colocation space","retail space","wholesale space",
        "sq m","sqm","m²","ft²","sqft","square metre","square meter",
        "campus site","campus area","land area","land holding",
        "parcel","plot","lot","site area","gross floor area","gfa",
        "net leasable area","nla","total floor area","tfa",
        # Sustainability / ESG signals (extended)
        "scope 1","scope 2","scope 3","net zero","carbon neutral",
        "carbon credit","carbon offset","renewable energy certificate","rec",
        "i-rec","go","guarantees of origin","24/7 cfe",
        "ppas","virtual ppa","vppa","capacity contract","long-term contract",
        "esg rating","esg disclosure","tcfd","sbti","science based targets",
        "water positive","water neutral","zero waste","circular economy",
        "heat reuse","waste heat recovery","district heating",
        # Construction signals (extended)
        "planning application","planning approval","planning permission",
        "building permit","building consent","development permit",
        "site plan","schematic design","design development","cd phase",
        "gmp","guaranteed maximum price","leed","leed certified",
        "breeam","energy star","green star","nabers","well building",
        "design-build","epc contract","turnkey contract","lump sum",
        "general contractor","gc","subcontractor","trade contractor",
        "civil works","structural steel","mechanical electrical plumbing",
        "mep","commissioning","acceptance testing","punch list",
        "certificate of occupancy","co","temporary co","tco",
        # Miscellaneous DC-market signals
        "colo","coloc","multi-tenant","single-tenant","dedicated facility",
        "enterprise data center","edge node","far-edge","near-edge",
        "5g edge","mobile edge","content delivery","cdn node",
        "internet exchange point","ixp","peering","transit","dark fiber",
        "submarine cable","landing station","cable landing","terrestrial fiber",
        "disaster recovery","dr site","business continuity","bc dr",
        "cloud on-ramp","direct connect","expressroute","private connect",
        "cloud region","availability zone","az","region launch",
    ]
    if any(p in t for p in primary):
        return True
    if sum(1 for s in secondary if s in t) >= 1:   # lowered threshold to 1 for secondary
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
    # MW/GW capacity
    m = re.search(r"([\d,]+(?:\.\d+)?)\s*(GW|MW|gigawatt|megawatt|kilowatt|kw)\b", text, re.I)
    if m:
        val = m.group(1).replace(",", "")
        unit = m.group(2).upper()
        if unit in ("KILOWATT", "KW"):
            unit = "kW"
        return val + " " + unit
    # Acres / hectares as capacity proxy
    m2 = re.search(r"([\d,]+(?:\.\d+)?)\s*(acres?|hectares?)", text, re.I)
    if m2:
        return m2.group(1).replace(",", "") + " " + m2.group(2).lower()
    return ""


def detect_deal_size(text):
    # Try multi-currency: $, €, £, ¥, ₹, AED, SGD, etc.
    m = re.search(
        r"(\$|€|£|¥|₹|US\$|USD|EUR|GBP|AED|SGD|INR|JPY|AUD|CAD|BRL)\s*([\d,.]+)\s*(billion|bn|million|mn|m\b|crore|lakh)",
        text, re.I
    )
    if m:
        sym = m.group(1).strip()
        val = m.group(2).replace(",", "")
        unit = m.group(3).lower()
        # Normalise symbol
        sym_map = {"US$": "$", "USD": "$", "EUR": "€", "GBP": "£", "AED": "AED ", "SGD": "SGD ",
                   "INR": "₹", "JPY": "¥", "AUD": "A$", "CAD": "C$", "BRL": "R$"}
        sym = sym_map.get(sym.upper(), sym)
        if unit in ("billion", "bn"):
            return f"{sym}{val}bn"
        if unit in ("crore",):
            return f"{sym}{val} Cr"
        return f"{sym}{val}m"
    # Fallback: number + unit without symbol
    m2 = re.search(r"([\d,.]+)\s*(billion|bn)\s+(dollar|euro|pound|dirham|rupee|yen)", text, re.I)
    if m2:
        return f"{m2.group(1).replace(',', '')}bn"
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


def _normalise_headline(h):
    """Normalise headline for comparison: lowercase, strip punctuation/source suffix."""
    h = h.lower().strip()
    # Strip common source suffixes added by Google News
    h = re.sub(r"\s*[-–|]\s*\w[\w\s]{1,30}$", "", h)
    # Strip special chars
    h = re.sub(r"[^\w\s]", " ", h)
    h = re.sub(r"\s+", " ", h).strip()
    return h


def fuzzy_similar(a, b, threshold=0.88):
    """True if two normalised headlines are likely the same story."""
    na, nb = _normalise_headline(a), _normalise_headline(b)
    # Exact match after normalisation
    if na == nb:
        return True
    # Sequence similarity
    ratio = SequenceMatcher(None, na, nb).ratio()
    if ratio >= threshold:
        return True
    # One is a substring of the other (short headline vs long headline of same story)
    shorter, longer = (na, nb) if len(na) <= len(nb) else (nb, na)
    if len(shorter) >= 30 and shorter in longer:
        return True
    return False


def deduplicate(articles):
    """
    1. URL-based exact dedup (same URL = same article).
    2. Fuzzy headline dedup — when two articles match, keep the one from
       the source with the lowest _priority number (DCD = 1 wins).
    """
    # Step 1: URL dedup — sort by priority so DCD URLs win ties
    seen_urls = {}
    for art in sorted(articles, key=lambda x: x.get("_priority", 99)):
        url = str(art.get("URL", art.get("url", ""))).strip().rstrip("/")
        if url and url not in seen_urls:
            seen_urls[url] = art
    url_deduped = list(seen_urls.values())

    # Step 2: Fuzzy headline dedup — sort by priority first so DCD is kept
    url_deduped.sort(key=lambda x: x.get("_priority", 99))
    keep = []
    seen_headlines = []
    for art in url_deduped:
        hl = art.get("Headline", art.get("headline", ""))
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
        "ISO / RTO": detect_iso_rto(hl),
        "_date_obj": d,
    }


# ═══════════════════════════════════════════════════════════════════════════════
#  FREE / BUILT-IN INTELLIGENCE SUMMARISER  (no API key, no external ML libs)
# ═══════════════════════════════════════════════════════════════════════════════

_STOPWORDS = {
    "a","an","the","and","or","but","in","on","at","to","for","of","with",
    "is","are","was","were","be","been","being","have","has","had","do","does",
    "did","will","would","could","should","may","might","shall","can","that",
    "this","these","those","its","it","i","we","they","he","she","by","from",
    "as","into","about","over","up","after","before","during","data","center",
    "datacenter","centers","new","says","said","says","say","also","more","than",
}

def _tfidf_scores(headlines):
    """Lightweight TF-IDF over headline tokens → sentence score dict."""
    tokenize = lambda h: [w.lower() for w in re.findall(r"[a-zA-Z]{3,}", h)
                          if w.lower() not in _STOPWORDS]
    tokenized = [tokenize(h) for h in headlines]
    # IDF
    N = len(headlines)
    df_counts = Counter()
    for toks in tokenized:
        for t in set(toks):
            df_counts[t] += 1
    idf = {t: math.log((N + 1) / (c + 1)) + 1 for t, c in df_counts.items()}
    # TF-IDF score per sentence
    scores = []
    for toks in tokenized:
        if not toks:
            scores.append(0.0)
            continue
        tf = Counter(toks)
        s = sum(tf[t] * idf.get(t, 1) for t in toks) / len(toks)
        scores.append(s)
    return scores


def generate_local_summary(df, sel_desc, date_range):
    """
    Wood Mackenzie-grade market intelligence briefing.
    Produces deeply analytical, prose-rich sections with quantified signals,
    company deep-dives, deal flow analysis, and forward-looking commentary —
    modelled on top-tier energy/infrastructure research houses.
    """
    headlines = df["Headline"].tolist()
    scores    = _tfidf_scores(headlines)
    df2 = df.copy()
    df2["_score"] = scores

    total = len(df2)
    if total == 0:
        return "No articles in the selected view. Adjust filters and regenerate."

    # ── Core aggregations ─────────────────────────────────────────────────────
    top_regions  = df2["Region"].value_counts()
    top_topics   = df2["Topic"].value_counts()
    top_sents    = df2["Sentiment"].value_counts()
    top_countries= df2["Country"].value_counts()

    mw_df    = df2[df2["Capacity"] != ""].copy()
    deal_df  = df2[df2["Deal Size"] != ""].copy()
    mw_vals  = mw_df["Capacity"].tolist()
    deal_vals= deal_df["Deal Size"].tolist()

    # Build full company counter from Companies column
    all_companies_list = []
    for v in df2["Companies"]:
        if v:
            all_companies_list.extend([c.strip() for c in str(v).split(",") if c.strip()])
    co_counter = Counter(all_companies_list)
    top_cos    = [co for co, _ in co_counter.most_common(20)]

    # Sentiment counts
    proposed   = int(top_sents.get("Proposed", 0))
    approved   = int(top_sents.get("Approved", 0))
    under_c    = int(top_sents.get("Under Construction", 0))
    opened     = int(top_sents.get("Opened / Live", 0))
    challenged = int(top_sents.get("Challenged", 0))
    news_c     = int(top_sents.get("News", 0))

    def pct(n): return f"{round(n/total*100)}%" if total else "0%"
    def hl(sub, n=3):
        """Return top n scored headlines from a sub-dataframe."""
        if sub.empty: return []
        return sub.nlargest(n, "_score")["Headline"].tolist()

    # ── Prose helpers ─────────────────────────────────────────────────────────
    def region_prose():
        parts = []
        for reg, cnt in top_regions.items():
            pct_val = pct(cnt)
            top_c   = df2[df2["Region"]==reg]["Country"].value_counts().head(4).index.tolist()
            parts.append(f"{reg} ({cnt} articles, {pct_val}) led by {', '.join(top_c)}")
        return "; ".join(parts) + "."

    def topic_prose():
        parts = []
        for top, cnt in top_topics.items():
            parts.append(f"{top} ({cnt} articles, {pct(cnt)})")
        return ", ".join(parts)

    # ── 1. EXECUTIVE SUMMARY ─────────────────────────────────────────────────
    dominant_region  = top_regions.index[0] if not top_regions.empty else "Global"
    dominant_topic   = top_topics.index[0]  if not top_topics.empty else "General"
    dominant_country = top_countries.index[0] if not top_countries.empty else "—"

    momentum_signal = (
        "strongly positive" if proposed > approved * 1.5
        else "balanced" if abs(proposed - approved) <= 2
        else "approval-constrained"
    )

    capacity_total_mw = 0
    for cap in mw_vals:
        m = re.search(r"([\d,.]+)\s*(GW|MW)", str(cap), re.I)
        if m:
            v = float(m.group(1).replace(",", ""))
            capacity_total_mw += v * 1000 if m.group(2).upper() == "GW" else v

    exec_lines = [
        f"This intelligence briefing synthesises {total} data centre industry articles "
        f"published between {date_range}, filtered to: {sel_desc}. "
        f"The analysis draws on headlines sourced from trade press, RSS feeds, and news aggregators, "
        f"auto-enriched with topic classification, project status, capacity and deal-size extraction, "
        f"and named-entity recognition across {df2['Country'].nunique()} countries.",

        f"\nActivity is geographically concentrated in {dominant_region}, with {dominant_country} "
        f"representing the single most active market in the filtered dataset. "
        f"Thematically, {dominant_topic} stories dominate at {pct(int(top_topics.iloc[0]))} of total "
        f"coverage, followed by {', '.join(str(t) for t in top_topics.index[1:3])}.",

        f"\nThe project pipeline presents a {momentum_signal} outlook: {proposed} proposed projects "
        f"against {approved} approvals, with {under_c} actively under construction and {opened} "
        f"recently commissioned. Regulatory friction accounts for {challenged} challenged or contested "
        f"articles ({pct(challenged)} of the dataset). "
        + (f"Identified capacity in announced or proposed projects totals approximately "
           f"{capacity_total_mw:,.0f} MW across {len(mw_df)} discrete announcements. "
           if capacity_total_mw > 0 else "")
        + (f"Disclosed deal flow spans {len(deal_df)} transactions including: "
           f"{'; '.join(deal_vals[:6])}{'.' if deal_vals else ''}"
           if deal_vals else ""),
    ]
    exec_summary = "\n\n".join(p.strip() for p in exec_lines if p.strip())

    # ── 2. KEY THEMES & TRENDS ────────────────────────────────────────────────
    theme_lines = []
    for topic, cnt in top_topics.items():
        sub = df2[df2["Topic"] == topic]
        ex  = hl(sub, 2)
        ex_str = ""
        if ex:
            # Truncate cleanly at word boundary
            h1 = ex[0][:110].rsplit(" ", 1)[0] if len(ex[0]) > 110 else ex[0]
            ex_str = f' Headline examples include: "{h1}"'
            if len(ex) > 1:
                h2 = ex[1][:100].rsplit(" ", 1)[0] if len(ex[1]) > 100 else ex[1]
                ex_str += f'; and "{h2}".'
            else:
                ex_str += "."
        theme_lines.append(
            f"• **{topic}** — {cnt} article{'s' if cnt > 1 else ''} ({pct(cnt)}).{ex_str}"
        )

    # ── 3. MAJOR PROJECTS & DEALS ─────────────────────────────────────────────
    proj_bullets = []
    # First: deal-size articles ranked by deal value then score
    deal_sub = deal_df.nlargest(min(8, len(deal_df)), "_score") if not deal_df.empty else pd.DataFrame()
    for _, r in deal_sub.iterrows():
        parts = [f"**{r['Headline'][:130]}**"]
        tags  = []
        if r.get("Deal Size"): tags.append(f"Deal: {r['Deal Size']}")
        if r.get("Capacity"):  tags.append(f"Capacity: {r['Capacity']}")
        if r.get("Country"):   tags.append(f"Market: {r['Country']}")
        if r.get("Companies"): tags.append(f"Companies: {r['Companies']}")
        if tags: parts.append(" · ".join(tags))
        proj_bullets.append("• " + " | ".join(parts))

    # Then: capacity-only articles
    cap_only = mw_df[mw_df["Deal Size"] == ""].nlargest(min(8, len(mw_df)), "_score") if not mw_df.empty else pd.DataFrame()
    for _, r in cap_only.iterrows():
        parts = [f"**{r['Headline'][:130]}**"]
        tags  = [f"Capacity: {r['Capacity']}", f"Market: {r['Country']}"]
        if r.get("Companies"): tags.append(f"Companies: {r['Companies']}")
        parts.append(" · ".join(tags))
        proj_bullets.append("• " + " | ".join(parts))

    if not proj_bullets:
        # Fall back to highest-scored articles
        for _, r in df2.nlargest(8, "_score").iterrows():
            proj_bullets.append(
                f"• **{r['Headline'][:130]}** | Market: {r['Country']} · Topic: {r['Topic']}"
            )

    # ── 4. REGULATORY & PERMITTING LANDSCAPE ──────────────────────────────────
    reg_df  = df2[df2["Topic"] == "Permits"]
    chall_df= df2[df2["Sentiment"] == "Challenged"]
    appr_df = df2[df2["Sentiment"] == "Approved"]

    reg_intro = (
        f"Permitting and regulatory dynamics account for {len(reg_df)} articles "
        f"({pct(len(reg_df))}) in the current selection. "
    )
    if len(chall_df) > 0:
        reg_intro += (
            f"Contested or blocked projects number {len(chall_df)}, "
            f"indicating {'elevated' if len(chall_df) > 3 else 'moderate'} community and regulatory "
            f"resistance in this geography and period. "
        )
    if len(appr_df) > 0:
        reg_intro += f"Approvals recorded: {len(appr_df)} projects cleared planning in the period. "
    if reg_df.empty and chall_df.empty:
        reg_intro += "No specific permitting friction or approval events detected in the current filter."

    reg_bullets = ["• " + h for h in hl(pd.concat([reg_df, chall_df, appr_df]).drop_duplicates(), 8)]
    if not reg_bullets:
        reg_bullets = ["• No permitting-specific articles in current selection."]

    # ── 5. POWER & INFRASTRUCTURE ─────────────────────────────────────────────
    pwr_df  = df2[df2["Topic"] == "Power"]
    pwr_intro = (
        f"Power and energy infrastructure is a recurring theme across {len(pwr_df)} articles "
        f"({pct(len(pwr_df))}), reflecting the sector's acute dependence on grid capacity, "
        f"PPAs, and alternative generation sources. "
    )
    if capacity_total_mw > 0:
        pwr_intro += (
            f"Across all capacity-cited articles, aggregate announced load totals "
            f"approximately {capacity_total_mw:,.0f} MW — a figure that underscores the scale "
            f"of power procurement challenges facing operators in this market. "
        )
    pwr_bullets = ["• " + h for h in hl(pwr_df, 7)]
    if not pwr_bullets:
        pwr_bullets = ["• No dedicated power/infrastructure articles in current selection."]

    # ── 6. COMPANY ACTIVITY & COMPETITIVE LANDSCAPE ──────────────────────────
    # Build richer per-company profiles
    co_section_lines = []

    # Group by hyperscalers vs colo/operators vs investors
    hyperscalers = {"Microsoft","Google","Amazon","AWS","Meta","Apple","Oracle","Alibaba",
                    "Tencent","ByteDance","Baidu","Alphabet","OpenAI","Stargate","CoreWeave",
                    "xAI","Anthropic","IBM","Huawei","Samsung","NVIDIA"}
    colo_ops     = {"Equinix","Digital Realty","Iron Mountain","CoreSite","CyrusOne","NTT",
                    "Vantage","Switch","EdgeConneX","Flexential","DataBank","QTS","Colt",
                    "Global Switch","ChinaData","GDS","STACK","Aligned","Compass","Tract",
                    "Yondr","Venari","Verne","NextDC","AirTrunk","atNorth","CtrlS","Yotta",
                    "STT GDC","Keppel","Singtel","Telstra","Macquarie","VIRTUS","Kao",
                    "Ascenty","Odata","Scala","Luminet","Etix","Nabiax","Bolder","DigiPlex",
                    "Bulk Infrastructure","Green Mountain","365 Data Centers","DigiCo",
                    "Cloudflare","Fastly","Akamai","Lumen","Zayo","Cogent","Pulsant","Iomart",
                    "Telehouse","Interxion","euNetworks","DE-CIX","MEVSPACE","Beyond.pl","Atman",
                    "Applied Digital","Hut 8","Core Scientific","Riot","Marathon","Iren",
                    "CyrusOne","Flexential","DataBank","T5","Skybox","Stream","CloudHQ"}

    mentioned_cos = [(co, cnt) for co, cnt in co_counter.most_common(30)]

    if mentioned_cos:
        # Narrative intro
        co_intro_parts = []
        hs_mentioned  = [(co,c) for co,c in mentioned_cos if co in hyperscalers]
        col_mentioned = [(co,c) for co,c in mentioned_cos if co in colo_ops]
        other_mentioned=[(co,c) for co,c in mentioned_cos if co not in hyperscalers and co not in colo_ops]

        if hs_mentioned:
            hs_str = ", ".join(
                f"{co} ({c} mention{'s' if c > 1 else ''})" for co, c in hs_mentioned[:5]
            )
            co_intro_parts.append(
                f"Hyperscaler activity is led by {hs_str}, "
                f"signalling continued large-scale capacity expansion in this market."
            )
        if col_mentioned:
            col_str = ", ".join(
                f"{co} ({c} mention{'s' if c > 1 else ''})" for co, c in col_mentioned[:5]
            )
            co_intro_parts.append(
                f"Among operators and colocation providers, {col_str} "
                f"feature prominently, reflecting active build, partnership, or M&A dynamics."
            )
        if other_mentioned:
            oth_str = ", ".join(
                f"{co} ({c} mention{'s' if c > 1 else ''})" for co, c in other_mentioned[:5]
            )
            co_intro_parts.append(
                f"Additional notable participants include {oth_str}, "
                f"spanning investors, developers, utilities, and technology vendors."
            )

        co_section_lines.append("\n".join(co_intro_parts))
        co_section_lines.append("")

        # Per-company detail bullets
        for co, cnt in mentioned_cos[:20]:
            co_arts = df2[
                df2["Companies"].str.contains(re.escape(co), na=False, case=False) |
                df2["Headline"].str.contains(re.escape(co), na=False, case=False)
            ]
            topics_seen    = co_arts["Topic"].value_counts().head(3).index.tolist()
            sents_seen     = co_arts["Sentiment"].value_counts().head(2).index.tolist()
            countries_seen = co_arts["Country"].value_counts().head(3).index.tolist()
            cap_arts       = co_arts[co_arts["Capacity"] != ""]["Capacity"].tolist()
            deal_arts      = co_arts[co_arts["Deal Size"] != ""]["Deal Size"].tolist()

            detail_parts = [f"focus: {', '.join(topics_seen) if topics_seen else 'General'}"]
            if countries_seen and countries_seen != ["Global"]:
                detail_parts.append(f"markets: {', '.join(countries_seen)}")
            if sents_seen:
                detail_parts.append(f"status signals: {', '.join(sents_seen)}")
            if cap_arts:
                detail_parts.append(f"capacity cited: {'; '.join(cap_arts[:3])}")
            if deal_arts:
                detail_parts.append(f"deal flow: {'; '.join(deal_arts[:2])}")

            co_section_lines.append(
                f"• **{co}** — {cnt} article{'s' if cnt > 1 else ''} | "
                + " | ".join(detail_parts)
            )
    else:
        co_section_lines.append(
            "• No named company entities detected in the current filtered view. "
            "Consider broadening the date range, region, or topic filters to surface company-level signals."
        )

    # ── 7. REGIONAL BREAKDOWN (detailed) ─────────────────────────────────────
    region_lines = []
    for region, rdf in df2.groupby("Region"):
        top_c      = rdf["Country"].value_counts().head(5).index.tolist()
        top_t      = rdf["Topic"].value_counts().head(3).index.tolist()
        cap_in_reg = rdf[rdf["Capacity"] != ""]["Capacity"].tolist()
        deals_reg  = rdf[rdf["Deal Size"] != ""]["Deal Size"].tolist()
        prop_reg   = len(rdf[rdf["Sentiment"]=="Proposed"])
        appr_reg   = len(rdf[rdf["Sentiment"]=="Approved"])
        chall_reg  = len(rdf[rdf["Sentiment"]=="Challenged"])

        region_lines.append(f"• **{region}** — {len(rdf)} articles ({pct(len(rdf))})")
        region_lines.append(
            f"  Lead markets: {', '.join(top_c)}. "
            f"Dominant themes: {', '.join(top_t)}. "
            f"Pipeline: {prop_reg} proposed, {appr_reg} approved, {chall_reg} challenged."
            + (f" Capacity cited: {'; '.join(cap_in_reg[:4])}." if cap_in_reg else "")
            + (f" Deal flow: {'; '.join(deals_reg[:3])}." if deals_reg else "")
        )

    # ── 8. MARKET OUTLOOK & FORWARD SIGNALS ───────────────────────────────────
    pipeline_signal = (
        "strongly positive — new proposal volume significantly outpaces current approvals, "
        "suggesting continued permitting and planning activity over the next 12–24 months"
        if proposed > approved * 1.5
        else "broadly balanced — approvals are broadly keeping pace with new project announcements, "
        "indicating a market in active but orderly expansion"
        if abs(proposed - approved) <= 3
        else "approval-constrained — the gap between proposals and approvals suggests permitting "
        "timelines and grid connection queues may be limiting near-term delivery"
    )

    reg_risk = (
        "elevated" if challenged > 4
        else "moderate" if challenged > 1
        else "low"
    )

    outlook_bullets = [
        f"• **Pipeline momentum:** {pipeline_signal}. "
        f"With {proposed} proposed vs {approved} approved and {under_c} under construction, "
        f"near-term delivery visibility {'is strong' if under_c >= 3 else 'remains limited'}.",

        f"• **Commissioning activity:** {opened} facility opening{'s' if opened != 1 else ''} recorded "
        f"in the period, confirming active supply additions to the market.",

        f"• **Regulatory risk:** {reg_risk.capitalize()} — {challenged} contested or challenged "
        f"project{'s' if challenged != 1 else ''} in the dataset. "
        + ("Community opposition, moratorium risk, and planning delays represent the primary "
           "near-term threat to delivery timelines."
           if challenged > 2
           else "Regulatory environment appears broadly supportive of new capacity in this geography."),

        f"• **Company watch:** {', '.join(top_cos[:6])} represent the highest-frequency participants "
        f"in this period. Monitor for further capacity announcements, M&A, and permitting outcomes.",
    ]

    if mw_vals:
        outlook_bullets.append(
            f"• **Capacity pipeline:** {len(mw_df)} articles reference explicit MW or GW figures. "
            f"Aggregate announced load stands at approximately {capacity_total_mw:,.0f} MW across "
            f"disclosed projects. Actual delivered capacity will depend on permitting, grid "
            f"connection, and financing milestones."
        )
    if deal_vals:
        outlook_bullets.append(
            f"• **Investment flow:** {len(deal_df)} deal-citing articles totalling disclosed "
            f"transactions of {', '.join(deal_vals[:8])}. "
            f"Deal density suggests active capital deployment into the sector."
        )

    # ── Assemble final document ────────────────────────────────────────────────
    doc_parts = []

    doc_parts += [
        "## 1. Executive Summary\n\n",
        exec_summary, "\n\n",
    ]
    doc_parts += [
        "## 2. Key Themes & Market Dynamics\n\n",
        f"Coverage across {total} articles spans the following thematic clusters: "
        f"{topic_prose()}.\n\n",
        "\n".join(theme_lines), "\n\n",
    ]
    doc_parts += [
        "## 3. Major Projects, Deals & Capacity Announcements\n\n",
        "\n".join(proj_bullets) if proj_bullets else "• No capacity or deal-cited articles in current selection.", "\n\n",
    ]
    doc_parts += [
        "## 4. Regulatory & Permitting Landscape\n\n",
        reg_intro + "\n\n",
        "\n".join(reg_bullets), "\n\n",
    ]
    doc_parts += [
        "## 5. Power & Infrastructure\n\n",
        pwr_intro + "\n\n",
        "\n".join(pwr_bullets), "\n\n",
    ]
    doc_parts += [
        "## 6. Company Activity & Competitive Landscape\n\n",
        "\n".join(co_section_lines), "\n\n",
    ]
    doc_parts += [
        "## 7. Regional Breakdown\n\n",
        "\n".join(region_lines), "\n\n",
    ]
    doc_parts += [
        "## 8. Market Outlook & Forward Signals\n\n",
        "\n".join(outlook_bullets), "\n",
    ]

    return "".join(doc_parts)



# ═══════════════════════════════════════════════════════════════════════════════
#  WORD (.docx) EXPORT
# ═══════════════════════════════════════════════════════════════════════════════

def _set_cell_bg(cell, hex_color):
    """Set table cell background colour."""
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd  = OxmlElement("w:shd")
    shd.set(qn("w:val"),   "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"),  hex_color)
    tcPr.append(shd)


def build_briefing_docx(summary_text, sel_desc, date_range, df):
    """Return bytes of a polished Word briefing document."""
    if not _DOCX_OK:
        return None
    buf = io.BytesIO()
    doc = DocxDocument()
    # IST timestamp
    from datetime import timezone as _tz_d, timedelta as _td_d
    _IST_D = _tz_d(_td_d(hours=5, minutes=30))
    _ist_str_d = datetime.now(_IST_D).strftime("%d %b %Y, %I:%M %p IST")

    # Page margins
    for section in doc.sections:
        section.top_margin    = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin   = Inches(1.15)
        section.right_margin  = Inches(1.15)

    # ── Cover block ──────────────────────────────────────────────────────────
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run("GLOBAL DATA CENTER INTELLIGENCE BRIEFING")
    run.bold = True
    run.font.size = Pt(22)
    run.font.color.rgb = RGBColor(0x00, 0x47, 0xE1)
    run.font.name = "Calibri"

    sub = doc.add_paragraph()
    sub.alignment = WD_ALIGN_PARAGRAPH.LEFT
    sr = sub.add_run(f"Selection: {sel_desc}   |   Period: {date_range}")
    sr.font.size  = Pt(9)
    sr.font.color.rgb = RGBColor(0x6A, 0x80, 0xA8)
    sr.font.name  = "Calibri"

    dr = sub.add_run(f"\nGenerated: {_ist_str_d}   |   Articles analysed: {len(df)}")
    dr.font.size  = Pt(9)
    dr.font.color.rgb = RGBColor(0x6A, 0x80, 0xA8)
    dr.font.name  = "Calibri"

    doc.add_paragraph()  # spacer

    # ── Stats summary table ──────────────────────────────────────────────────
    stats = [
        ("Total Articles", str(len(df))),
        ("Top Region",     df["Region"].value_counts().index[0] if not df.empty else "—"),
        ("Top Topic",      df["Topic"].value_counts().index[0]  if not df.empty else "—"),
        ("With Capacity",  str(len(df[df["Capacity"] != ""]))),
        ("With Deal Size", str(len(df[df["Deal Size"] != ""]))),
    ]
    tbl = doc.add_table(rows=1, cols=len(stats))
    tbl.style = "Table Grid"
    hdr_cells = tbl.rows[0].cells
    for i, (label, val) in enumerate(stats):
        hdr_cells[i].text = ""
        p2 = hdr_cells[i].paragraphs[0]
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r1 = p2.add_run(label + "\n")
        r1.font.size = Pt(7.5); r1.font.bold = True
        r1.font.color.rgb = RGBColor(0xB8, 0xC8, 0xE0); r1.font.name = "Calibri"
        r2 = p2.add_run(val)
        r2.font.size = Pt(13); r2.font.bold = True
        r2.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF); r2.font.name = "Calibri"
        _set_cell_bg(hdr_cells[i], "0F1E36")

    doc.add_paragraph()

    # ── Parse and write markdown sections ────────────────────────────────────
    ACCENT = RGBColor(0x00, 0x47, 0xE1)
    BODY   = RGBColor(0x1A, 0x1A, 0x2E)
    BULLET = RGBColor(0x2A, 0x3E, 0x60)

    for line in summary_text.splitlines():
        line = line.rstrip()
        if line.startswith("## "):
            heading = line[3:]
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(14)
            p.paragraph_format.space_after  = Pt(4)
            run = p.add_run(heading.upper())
            run.bold = True
            run.font.size = Pt(11)
            run.font.color.rgb = ACCENT
            run.font.name = "Calibri"
            # underline via border
            pPr = p._p.get_or_add_pPr()
            pBdr = OxmlElement("w:pBdr")
            bottom = OxmlElement("w:bottom")
            bottom.set(qn("w:val"),   "single")
            bottom.set(qn("w:sz"),    "4")
            bottom.set(qn("w:space"), "1")
            bottom.set(qn("w:color"), "0047E1")
            pBdr.append(bottom)
            pPr.append(pBdr)
        elif line.startswith("• "):
            p = doc.add_paragraph(style="List Bullet")
            p.paragraph_format.space_after = Pt(2)
            # Parse **bold** segments inline
            _parts = re.split(r"(\*\*.*?\*\*)", line[2:])
            for _part in _parts:
                if _part.startswith("**") and _part.endswith("**"):
                    _run = p.add_run(_part[2:-2])
                    _run.bold = True
                    _run.font.size = Pt(9.5)
                    _run.font.color.rgb = RGBColor(0x1A, 0x1A, 0x2E)
                    _run.font.name = "Calibri"
                elif _part:
                    _run = p.add_run(_part)
                    _run.font.size = Pt(9.5)
                    _run.font.color.rgb = RGBColor(0x1A, 0x1A, 0x2E)
                    _run.font.name = "Calibri"
        elif line.strip():
            p = doc.add_paragraph()
            p.paragraph_format.space_after = Pt(4)
            _bparts = re.split(r"(\*\*.*?\*\*)", line)
            for _bp in _bparts:
                if _bp.startswith("**") and _bp.endswith("**"):
                    _br = p.add_run(_bp[2:-2])
                    _br.bold = True
                    _br.font.size = Pt(10)
                    _br.font.color.rgb = BODY
                    _br.font.name = "Calibri"
                elif _bp:
                    _br = p.add_run(_bp)
                    _br.font.size = Pt(10)
                    _br.font.color.rgb = BODY
                    _br.font.name = "Calibri"

    # ── Footer note ───────────────────────────────────────────────────────────
    doc.add_paragraph()
    fp = doc.add_paragraph()
    fr = fp.add_run(
        "This briefing was generated automatically using built-in NLP analysis of "
        "publicly available news headlines. Always verify against primary sources."
    )
    fr.font.size  = Pt(8)
    fr.font.italic = True
    fr.font.color.rgb = RGBColor(0xA0, 0xA0, 0xA8)
    fr.font.name  = "Calibri"

    doc.save(buf)
    buf.seek(0)
    return buf.read()


# ═══════════════════════════════════════════════════════════════════════════════
#  PDF EXPORT
# ═══════════════════════════════════════════════════════════════════════════════

def build_briefing_pdf(summary_text, sel_desc, date_range, df):
    """Return bytes of a styled PDF briefing document."""
    if not _PDF_OK:
        return None
    buf = io.BytesIO()
    # IST = UTC+5:30
    from datetime import timezone, timedelta as _td
    _IST = timezone(_td(hours=5, minutes=30))
    _now_ist = datetime.now(_IST)
    _ist_str = _now_ist.strftime("%d %b %Y, %I:%M %p IST")

    doc = SimpleDocTemplate(
        buf, pagesize=letter,
        leftMargin=inch * 1.1, rightMargin=inch * 1.1,
        topMargin=inch * 1.0,  bottomMargin=inch * 1.0,
    )

    NAVY  = rl_colors.HexColor("#0047E1")
    DARK  = rl_colors.HexColor("#0F1E36")
    GREY  = rl_colors.HexColor("#6A80A8")
    WHITE = rl_colors.white
    BODY  = rl_colors.HexColor("#1A1A2E")

    styles = getSampleStyleSheet()

    s_title = ParagraphStyle("DCTitle",
        fontSize=20, textColor=NAVY, fontName="Helvetica-Bold",
        spaceAfter=4, leading=24)
    s_sub = ParagraphStyle("DCSub",
        fontSize=8.5, textColor=GREY, fontName="Helvetica",
        spaceAfter=14, leading=13)
    s_head = ParagraphStyle("DCHead",
        fontSize=11, textColor=NAVY, fontName="Helvetica-Bold",
        spaceBefore=16, spaceAfter=5, leading=14,
        borderPadding=(0, 0, 3, 0))
    s_bullet = ParagraphStyle("DCBullet",
        fontSize=9.5, textColor=BODY, fontName="Helvetica",
        spaceAfter=3, leading=13, leftIndent=14,
        bulletIndent=2, bulletFontName="Helvetica", bulletFontSize=9)
    s_body = ParagraphStyle("DCBody",
        fontSize=10, textColor=BODY, fontName="Helvetica",
        spaceAfter=6, leading=14)
    s_footer = ParagraphStyle("DCFooter",
        fontSize=7.5, textColor=GREY, fontName="Helvetica-Oblique",
        spaceBefore=18, leading=11)

    story = []

    # ── Header ────────────────────────────────────────────────────────────────
    story.append(Paragraph("GLOBAL DATA CENTER INTELLIGENCE BRIEFING", s_title))
    story.append(Paragraph(
        f"<b>Selection:</b> {sel_desc}   |   <b>Period:</b> {date_range}<br/>"
        f"Generated: {_ist_str}   |   "
        f"Articles analysed: {len(df)}",
        s_sub))
    story.append(HRFlowable(width="100%", thickness=1.5, color=NAVY, spaceAfter=10))

    # ── Stats row (manual layout) ─────────────────────────────────────────────
    stats_labels = ["Total Articles", "Top Region", "Top Topic", "w/ Capacity", "w/ Deal Size"]
    stats_vals   = [
        str(len(df)),
        df["Region"].value_counts().index[0] if not df.empty else "—",
        df["Topic"].value_counts().index[0]  if not df.empty else "—",
        str(len(df[df["Capacity"] != ""])),
        str(len(df[df["Deal Size"] != ""])),
    ]
    stat_line = "   ·   ".join(f"<b>{l}:</b> {v}" for l, v in zip(stats_labels, stats_vals))
    story.append(Paragraph(stat_line, ParagraphStyle("statrow",
        fontSize=8.5, textColor=GREY, fontName="Helvetica",
        spaceAfter=14, leading=12)))
    story.append(HRFlowable(width="100%", thickness=0.5, color=GREY, spaceAfter=6))

    # ── Markdown sections ─────────────────────────────────────────────────────
    def _rl_inline(t):
        """Convert **bold** markers to ReportLab <b> tags for proper bold rendering."""
        import re as _re
        t = _re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", t)
        return t

    for line in summary_text.splitlines():
        line = line.rstrip()
        if line.startswith("## "):
            story.append(Paragraph(line[3:].upper(), s_head))
            story.append(HRFlowable(width="100%", thickness=0.8, color=NAVY, spaceAfter=4))
        elif line.startswith("• "):
            story.append(Paragraph(f"• {_rl_inline(line[2:])}", s_bullet))
        elif line.strip():
            story.append(Paragraph(_rl_inline(line), s_body))

    # ── Footer ────────────────────────────────────────────────────────────────
    story.append(HRFlowable(width="100%", thickness=0.5, color=GREY, spaceBefore=14, spaceAfter=4))
    story.append(Paragraph(
        "This briefing was generated automatically using built-in NLP analysis of "
        "publicly available news headlines. Always verify against primary sources.",
        s_footer))

    doc.build(story)
    buf.seek(0)
    return buf.read()


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
                 "Sentiment","Capacity","Deal Size","Companies","ISO / RTO","URL"]
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
        margin=dict(l=14, r=14, t=44, b=14),
        xaxis=dict(
            gridcolor="#0f1e36", gridwidth=1,
            linecolor=_GRID, linewidth=1,
            tickfont=dict(size=10, color=_TEXT),
            title_font=dict(color=_TEXT),
            zeroline=False,
        ),
        yaxis=dict(
            gridcolor="#0f1e36", gridwidth=1,
            linecolor=_GRID, linewidth=1,
            tickfont=dict(size=10, color=_TEXT),
            title_font=dict(color=_TEXT),
            zeroline=False,
        ),
        showlegend=False,
        hoverlabel=dict(
            bgcolor="#0d1e38",
            bordercolor="#0047e1",
            font=dict(color="#ccdaf5", size=12, family="Inter, sans-serif"),
        ),
        hovermode="closest",
    )
    return fig


def chart_topic_bar(df):
    tc = df["Topic"].value_counts().reset_index()
    tc.columns = ["Topic", "Count"]
    tc = tc.sort_values("Count")
    colors = [TOPIC_COLORS.get(t, "#2e4470") for t in tc["Topic"]]
    fig = go.Figure(go.Bar(
        x=tc["Count"], y=tc["Topic"], orientation="h",
        marker=dict(
            color=colors,
            line=dict(width=0),
            opacity=0.88,
        ),
        text=tc["Count"], textposition="outside",
        textfont=dict(color=_TITLE, size=11, family="DM Mono, monospace"),
        hovertemplate="<b>%{y}</b><br>📰 %{x} articles<extra></extra>",
        hoverlabel=dict(bgcolor="#0d1e38", bordercolor="#0047e1",
                        font=dict(color="#ccdaf5", size=12)),
    ))
    _dark(fig, 320)
    fig.update_layout(
        title=dict(text="Articles by Topic", font=dict(color=_TITLE, size=13, family="Syne, sans-serif"), x=0.01),
        xaxis=dict(showgrid=True, gridcolor="#0f1e36"),
        bargap=0.28,
    )
    return fig


def chart_region_bar(df):
    rc = df["Region"].value_counts().reset_index()
    rc.columns = ["Region", "Count"]
    rc = rc.sort_values("Count")
    colors = [REGION_COLORS.get(r, "#2e4470") for r in rc["Region"]]
    fig = go.Figure(go.Bar(
        x=rc["Count"], y=rc["Region"], orientation="h",
        marker=dict(color=colors, line=dict(width=0), opacity=0.88),
        text=rc["Count"], textposition="outside",
        textfont=dict(color=_TITLE, size=11, family="DM Mono, monospace"),
        hovertemplate="<b>%{y}</b><br>📰 %{x} articles<extra></extra>",
    ))
    _dark(fig, 300)
    fig.update_layout(
        title=dict(text="Articles by Region", font=dict(color=_TITLE, size=13, family="Syne, sans-serif"), x=0.01),
        bargap=0.28,
    )
    return fig


def chart_country_bar(df, top_n=20):
    cc = df["Country"].value_counts().head(top_n).reset_index()
    cc.columns = ["Country", "Count"]
    cc = cc.sort_values("Count")
    n = len(cc)
    c_colors = [
        f"rgba({int(0 + 71*i/max(n-1,1))}, {int(71 + (180-71)*i/max(n-1,1))}, {int(225 + (255-225)*i/max(n-1,1))}, 0.88)"
        for i in range(n)
    ]
    fig = go.Figure(go.Bar(
        x=cc["Count"], y=cc["Country"], orientation="h",
        marker=dict(color=c_colors, line=dict(width=0)),
        text=cc["Count"], textposition="outside",
        textfont=dict(color=_TITLE, size=10, family="DM Mono, monospace"),
        hovertemplate="<b>%{y}</b><br>📰 %{x} articles<extra></extra>",
    ))
    _dark(fig, max(340, top_n * 24))
    fig.update_layout(
        title=dict(text=f"Top {top_n} Countries", font=dict(color=_TITLE, size=13, family="Syne, sans-serif"), x=0.01),
        bargap=0.22,
    )
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
        line=dict(color="#00b4ff", width=2.5, shape="spline", smoothing=0.8),
        marker=dict(color="#0047e1", size=6, line=dict(color="#00b4ff", width=1.5),
                    symbol="circle"),
        fill="tozeroy", fillcolor="rgba(0,71,225,0.09)",
        hovertemplate="<b>%{x}</b><br>📰 %{y} articles<extra></extra>",
    ))
    _dark(fig, 260)
    fig.update_layout(
        title=dict(text="Publication Volume Over Time", font=dict(color=_TITLE, size=13, family="Syne, sans-serif"), x=0.01),
        yaxis_title="Articles",
        xaxis=dict(showgrid=False),
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
        marker=dict(color=colors, line=dict(width=0), opacity=0.88),
        text=sc["Count"], textposition="outside",
        textfont=dict(color=_TITLE, size=11, family="DM Mono, monospace"),
        hovertemplate="<b>%{x}</b><br>📰 %{y} articles<extra></extra>",
    ))
    _dark(fig, 290)
    fig.update_layout(
        title=dict(text="Article Sentiment / Status", font=dict(color=_TITLE, size=13, family="Syne, sans-serif"), x=0.01),
        bargap=0.3,
    )
    return fig


def chart_donut(df):
    tc = df["Topic"].value_counts().reset_index()
    tc.columns = ["Topic", "Count"]
    colors = [TOPIC_COLORS.get(t, "#2e4470") for t in tc["Topic"]]
    fig = go.Figure(go.Pie(
        labels=tc["Topic"], values=tc["Count"], hole=0.6,
        marker=dict(colors=colors, line=dict(color=_BG, width=3)),
        textinfo="label+percent",
        textfont=dict(color=_TITLE, size=11),
        hovertemplate="<b>%{label}</b><br>📰 %{value} articles (%{percent})<extra></extra>",
        pull=[0.03] * len(tc),
    ))
    fig.update_layout(
        paper_bgcolor=_PAPER, plot_bgcolor=_BG,
        font=dict(family=_FONT, color=_TEXT),
        height=360, margin=dict(l=14, r=14, t=44, b=14),
        showlegend=True,
        legend=dict(bgcolor="rgba(0,0,0,0)", font=dict(color=_TITLE, size=10),
                    orientation="v", x=1.02, y=0.5),
        hoverlabel=dict(
            bgcolor="#0d1e38", bordercolor="#0047e1",
            font=dict(color="#ccdaf5", size=12, family="Inter, sans-serif"),
        ),
        annotations=[dict(
            text=f"<b>{len(df)}</b><br><span style='font-size:10px'>articles</span>",
            x=0.5, y=0.5, showarrow=False,
            font=dict(size=15, color=_TITLE, family=_FONT),
        )],
        title=dict(text="Topic Share", font=dict(color=_TITLE, size=13, family="Syne, sans-serif"), x=0.01),
    )
    return fig


def chart_source_bar(df):
    sc = df["Source"].value_counts().reset_index()
    sc.columns = ["Source", "Count"]
    sc = sc.sort_values("Count")
    colors = [SOURCE_META.get(s, SOURCE_META["Unknown"])["color"] for s in sc["Source"]]
    fig = go.Figure(go.Bar(
        x=sc["Count"], y=sc["Source"], orientation="h",
        marker=dict(color=colors, line=dict(width=0), opacity=0.88),
        text=sc["Count"], textposition="outside",
        textfont=dict(color=_TITLE, size=11, family="DM Mono, monospace"),
        hovertemplate="<b>%{y}</b><br>📰 %{x} articles<extra></extra>",
    ))
    _dark(fig, 290)
    fig.update_layout(
        title=dict(text="Articles by Source", font=dict(color=_TITLE, size=13, family="Syne, sans-serif"), x=0.01),
        bargap=0.28,
    )
    return fig


def chart_world_map(df):
    cc = df[df["Country"] != "Global"]["Country"].value_counts().reset_index()
    cc.columns = ["Country", "Count"]
    cc["ISO"] = cc["Country"].map(COUNTRY_ISO)
    cc = cc.dropna(subset=["ISO"])
    if cc.empty:
        return go.Figure()

    z_max = max(int(cc["Count"].max()), 1)

    # ── Palette matched to the app's dark UI ─────────────────────────────────
    _M_BG    = "#060e1c"   # near-black page bg — same feel as .stApp
    _M_OCEAN = "#07111f"   # deep navy ocean — slightly lighter than bg
    _M_LAND  = "#0d1b30"   # midnight-blue for zero-data land masses
    _M_GRID  = "#152038"   # subtle border matching app grid colour

    # Heat scale: deep navy (zero) → electric cyan → cobalt blue → neon gold (max)
    # Matches the app's accent palette (#0047e1 blue, #00b4ff cyan, #ffaa00 amber)
    _COLORSCALE = [
        [0.000, "#07111f"],   # zero  — ocean/bg level (invisible baseline)
        [0.05,  "#0a2040"],   # trace — barely there
        [0.15,  "#0d3a70"],   # low   — deep cobalt
        [0.30,  "#0047e1"],   # low-mid — brand blue
        [0.50,  "#00b4ff"],   # mid   — electric cyan
        [0.70,  "#00e5c8"],   # high  — teal-mint
        [0.85,  "#ffaa00"],   # very high — amber
        [1.000, "#ff6400"],   # max   — vivid orange-gold peak
    ]

    fig = go.Figure(go.Choropleth(
        locations=cc["ISO"],
        z=cc["Count"],
        text=cc["Country"],
        colorscale=_COLORSCALE,
        autocolorscale=False,
        reversescale=False,
        zauto=False,
        zmin=0,
        zmax=z_max,
        marker=dict(line=dict(color="#152038", width=0.6)),
        colorbar=dict(
            bgcolor="#0b1628",
            bordercolor="#152038",
            borderwidth=1,
            tickfont=dict(color="#6a80a8", size=10, family="DM Mono, monospace"),
            title=dict(
                text="Articles",
                font=dict(color="#b8c8e0", size=11, family="Syne, sans-serif"),
            ),
            len=0.6,
            thickness=14,
            tickformat="d",
            x=1.01,
        ),
        hovertemplate=(
            "<b style='color:#fff;font-size:13px;'>%{text}</b><br>"
            "<span style='color:#00b4ff;'>Articles: %{z}</span>"
            "<extra></extra>"
        ),
        showscale=True,
    ))

    fig.update_geos(
        bgcolor=_M_BG,
        landcolor=_M_LAND,
        oceancolor=_M_OCEAN,
        lakecolor=_M_OCEAN,
        rivercolor=_M_OCEAN,
        framecolor=_M_GRID,
        showland=True,
        showocean=True,
        showlakes=True,
        showrivers=False,
        showcountries=True,
        countrycolor="#152038",
        showcoastlines=True,
        coastlinecolor="#1a2e50",
        showframe=False,
        projection_type="natural earth",
        lataxis_range=[-60, 85],
    )

    fig.update_layout(
        paper_bgcolor=_M_BG,
        plot_bgcolor=_M_BG,
        height=520,
        margin=dict(l=0, r=60, t=44, b=0),
        title=dict(
            text="<b>Global Data Center Activity</b>",
            font=dict(color="#b8c8e0", size=14, family="Syne, sans-serif"),
            x=0.01, y=0.98,
        ),
        geo=dict(bgcolor=_M_BG),
        hoverlabel=dict(
            bgcolor="#0b1628",
            bordercolor="#0047e1",
            font=dict(color="#ccdaf5", size=12, family="Inter, sans-serif"),
        ),
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
        f'<table class="dc-table" style="width:100%;border-collapse:collapse;background:#060a10;">'
        f'<thead><tr>{heads}</tr></thead>'
        f'<tbody>{rows}</tbody></table></div>'
    )


def article_card(headline, date, url, source, country, topic, capacity, deal, sentiment, ai_score=None):
    tc = TOPIC_COLORS.get(topic, "#2e4470")
    sc_meta = SOURCE_META.get(source, SOURCE_META["Unknown"])
    cap_html = (
        f'<span style="background:rgba(255,170,0,0.12);color:#ffaa00;'
        f'border:1px solid rgba(255,170,0,0.3);border-radius:4px;'
        f'padding:2px 6px;font-family:monospace;font-size:.62rem;white-space:nowrap;" '
        f'title="Capacity announced">'
        f'\u26a1 {capacity}</span>'
    ) if capacity else ""
    deal_html = (
        f'<span style="background:rgba(0,230,118,0.1);color:#00e676;'
        f'border:1px solid rgba(0,230,118,0.25);border-radius:4px;'
        f'padding:2px 6px;font-family:monospace;font-size:.62rem;white-space:nowrap;" '
        f'title="Deal value">'
        f'{deal}</span>'
    ) if deal else ""
    sent_c = {
        "Opened / Live":"#00e676","Approved":"#00b4ff","Proposed":"#ffaa00",
        "Under Construction":"#00e5c8","Challenged":"#ff2d6b","News":"#2e4470",
    }.get(sentiment, "#2e4470")
    arrow = "\u2197"

    # AI score badge (optional)
    score_badge_html = ""
    if ai_score is not None:
        sc_color = (
            "#ff6400" if ai_score >= 40
            else "#ffaa00" if ai_score >= 25
            else "#00b4ff" if ai_score >= 15
            else "#3a5480"
        )
        score_badge_html = (
            f'<span class="score-badge" style="background:{sc_color}22;color:{sc_color};'
            f'border:1px solid {sc_color}44;border-radius:4px;'
            f'padding:2px 7px;font-family:monospace;font-size:.62rem;font-weight:bold;" '
            f'title="AI Signal Score — higher = more market significant">&#9733; {ai_score}</span>'
        )

    # High signal indicator strip
    is_high_signal = ai_score is not None and ai_score >= 30
    border_color = "#ffaa00" if is_high_signal else "#152038"
    top_strip = (
        f'<div style="position:absolute;top:0;left:0;right:0;height:2px;'
        f'background:linear-gradient(90deg,#ffaa00,#ff6400);border-radius:2px 2px 0 0;"></div>'
        if is_high_signal else ""
    )

    return (
        f'<div class="article-card-wrap" style="background:#0b1628;border:1px solid {border_color};border-radius:10px;'
        f'padding:.85rem 1.1rem;display:flex;justify-content:space-between;'
        f'align-items:flex-start;gap:.9rem;margin-bottom:.45rem;position:relative;">'
        f'{top_strip}'
        f'<div style="flex:1;min-width:0;">'
        f'<a href="{url}" target="_blank" '
        f'style="color:#ccdaf5;text-decoration:none;font-family:Inter,sans-serif;'
        f'font-size:.86rem;font-weight:500;line-height:1.5;" '
        f'title="Open article in new tab">{headline}</a>'
        f'<div style="margin-top:.35rem;display:flex;gap:.4rem;flex-wrap:wrap;">'
        f'<span style="font-family:monospace;font-size:.62rem;color:#2a3e60;" title="Publication date">\U0001f4c5 {date}</span>'
        f'<span style="font-family:monospace;font-size:.62rem;color:#4a6490;" title="Country detected">\U0001f30d {country}</span>'
        f'</div></div>'
        f'<div style="display:flex;flex-direction:column;align-items:flex-end;'
        f'gap:.28rem;flex-shrink:0;white-space:nowrap;">'
        f'<span style="background:{sc_meta["color"]}22;color:{sc_meta["color"]};'
        f'border:1px solid {sc_meta["color"]}44;border-radius:4px;'
        f'padding:2px 6px;font-family:monospace;font-size:.62rem;" title="Source: {source}">{sc_meta["short"]}</span>'
        f'<span style="background:{tc}22;color:{tc};border:1px solid {tc}44;'
        f'border-radius:4px;padding:2px 6px;font-family:monospace;font-size:.62rem;" title="Topic: {topic}">{topic}</span>'
        f'<span style="background:{sent_c}18;color:{sent_c};border:1px solid {sent_c}44;'
        f'border-radius:4px;padding:2px 6px;font-family:monospace;font-size:.62rem;" title="Project status">{sentiment}</span>'
        f'{cap_html}{deal_html}{score_badge_html}'
        f'<a href="{url}" target="_blank" '
        f'style="font-family:monospace;font-size:.65rem;color:#0047e1;text-decoration:none;" title="Open full article">'
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
        f'<div class="kpi-card" style="flex:1;min-width:150px;background:#0b1628;border:1px solid #152038;'
        f'border-radius:12px;padding:1.1rem 1.3rem;position:relative;overflow:hidden;" '
        f'title="{label}">'
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
        page_title="Global Data Center Intelligence",
        page_icon="\U0001f310",
        layout="wide",
        initial_sidebar_state="expanded",
    )
    st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

    # ── Timezone detection via browser JS ────────────────────────────────────
    # Reads the browser IANA tz and writes it into query params.
    # Default fallback = Asia/Kolkata (IST, UTC+5:30).
    tz_js = """
    <script>
    (function() {
        try {
            var tz = Intl.DateTimeFormat().resolvedOptions().timeZone || 'Asia/Kolkata';
            var params = new URLSearchParams(window.location.search);
            if (!params.get('tz') || params.get('tz') === 'UTC') {
                params.set('tz', tz);
                window.history.replaceState({}, '', window.location.pathname + '?' + params.toString());
            }
        } catch(e) {}
    })();
    </script>
    """
    st.markdown(tz_js, unsafe_allow_html=True)

    # ── Sidebar: aggressive force-open on every load ─────────────────────────
    # Strategy: three-layer approach
    #  1. Nuke ALL localStorage keys that relate to sidebar/collapsed state
    #  2. Force-click the sidebar open button if sidebar is found collapsed
    #  3. Inject a visible "☰ Filters" floating button as a permanent fallback
    sidebar_btn_js = """
    <script>
    (function() {

        // ── LAYER 1: Wipe every localStorage key that could keep sidebar closed ──
        function nukeLocalStorage() {
            try {
                var toDelete = [];
                for (var i = 0; i < localStorage.length; i++) {
                    var k = localStorage.key(i);
                    if (!k) continue;
                    var kl = k.toLowerCase();
                    if (
                        kl.includes('sidebar') ||
                        kl.includes('collapsed') ||
                        kl.includes('stSidebar') ||
                        kl.includes('streamlit') ||
                        kl.includes('layout') ||
                        kl.includes('toggled')
                    ) {
                        toDelete.push(k);
                    }
                }
                toDelete.forEach(function(k) { localStorage.removeItem(k); });
            } catch(e) {}

            // Also try sessionStorage
            try {
                var sToDelete = [];
                for (var j = 0; j < sessionStorage.length; j++) {
                    var sk = sessionStorage.key(j);
                    if (!sk) continue;
                    var skl = sk.toLowerCase();
                    if (
                        skl.includes('sidebar') ||
                        skl.includes('collapsed') ||
                        skl.includes('streamlit')
                    ) {
                        sToDelete.push(sk);
                    }
                }
                sToDelete.forEach(function(sk) { sessionStorage.removeItem(sk); });
            } catch(e) {}
        }
        nukeLocalStorage();

        // ── LAYER 2: Force-click the sidebar open if it appears collapsed ────────
        function forceSidebarOpen() {
            var sidebar = document.querySelector('[data-testid="stSidebar"]');
            if (!sidebar) return false;

            // Check if sidebar is actually collapsed (very narrow or off-screen)
            var rect = sidebar.getBoundingClientRect();
            var isCollapsed = rect.width < 60 || rect.left < -100;

            if (isCollapsed) {
                // Try every selector Streamlit uses for the reopen button
                var selectors = [
                    '[data-testid="stSidebarCollapsedControl"] button',
                    'div[data-testid="collapsedControl"] button',
                    'button[aria-label="Open sidebar"]',
                    'button[aria-label*="sidebar"]',
                    '[data-testid="collapsedControl"] button',
                ];
                for (var i = 0; i < selectors.length; i++) {
                    var btns = document.querySelectorAll(selectors[i]);
                    if (btns.length > 0) {
                        btns[0].click();
                        return true;
                    }
                }
            }
            return false;
        }

        // ── LAYER 3: Floating "☰ Filters" emergency button ───────────────────────
        function injectFiltersBtn() {
            var old = document.getElementById('_gdci_sidebar_btn');
            if (old) old.remove();

            var btn = document.createElement('button');
            btn.id = '_gdci_sidebar_btn';
            btn.innerHTML = '&#9776; Filters';
            btn.title = 'Open filter sidebar';
            btn.style.cssText = [
                'position:fixed',
                'top:12px',
                'left:12px',
                'z-index:2147483647',
                'background:linear-gradient(135deg,#0047e1,#00b4ff)',
                'color:#fff',
                'border:none',
                'border-radius:8px',
                'padding:7px 14px',
                'font-family:Syne,sans-serif',
                'font-weight:700',
                'font-size:0.82rem',
                'letter-spacing:0.04em',
                'cursor:pointer',
                'box-shadow:0 4px 16px rgba(0,71,225,0.55)',
                'display:none',
                'align-items:center',
                'gap:6px',
                'transition:opacity 0.2s,transform 0.2s',
            ].join(';');

            btn.onmouseenter = function() { btn.style.opacity='0.85'; btn.style.transform='translateY(-1px)'; };
            btn.onmouseleave = function() { btn.style.opacity='1';    btn.style.transform='translateY(0)';    };

            btn.onclick = function() {
                nukeLocalStorage();
                var selectors = [
                    '[data-testid="stSidebarCollapsedControl"] button',
                    'div[data-testid="collapsedControl"] button',
                    'button[aria-label="Open sidebar"]',
                    'button[aria-label*="sidebar"]',
                    '[data-testid="collapsedControl"] button',
                ];
                for (var i = 0; i < selectors.length; i++) {
                    var btns = document.querySelectorAll(selectors[i]);
                    if (btns.length > 0) { btns[0].click(); return; }
                }
            };

            document.body.appendChild(btn);

            // Poll: show button only when sidebar is collapsed
            function checkSidebar() {
                var sidebar = document.querySelector('[data-testid="stSidebar"]');
                if (!sidebar) { btn.style.display = 'flex'; return; }
                var rect = sidebar.getBoundingClientRect();
                var collapsed = rect.width < 60 || rect.left < -100;
                btn.style.display = collapsed ? 'flex' : 'none';
            }

            setInterval(checkSidebar, 300);
            checkSidebar();
        }

        // ── Boot sequence ────────────────────────────────────────────────────────
        function boot() {
            nukeLocalStorage();
            injectFiltersBtn();

            // Try to force open immediately, then retry a few times
            // (Streamlit renders the DOM progressively so we need retries)
            var attempts = 0;
            var maxAttempts = 12;
            var interval = setInterval(function() {
                nukeLocalStorage();
                var opened = forceSidebarOpen();
                attempts++;
                if (opened || attempts >= maxAttempts) clearInterval(interval);
            }, 400);
        }

        if (document.readyState === 'loading') {
            document.addEventListener('DOMContentLoaded', boot);
        } else {
            setTimeout(boot, 300);
        }

    })();
    </script>
    """
    st.markdown(sidebar_btn_js, unsafe_allow_html=True)

    _DEFAULT_TZ = "Asia/Kolkata"
    try:
        from zoneinfo import ZoneInfo
        _tz_raw = st.query_params.get("tz", _DEFAULT_TZ)
        if not _tz_raw or _tz_raw in ("UTC", "undefined", "null", ""):
            _tz_raw = _DEFAULT_TZ
        try:
            _user_tz = ZoneInfo(_tz_raw)
            _tz_str  = _tz_raw
        except Exception:
            _user_tz = ZoneInfo(_DEFAULT_TZ)
            _tz_str  = _DEFAULT_TZ
    except ImportError:
        _user_tz = None
        _tz_str  = _DEFAULT_TZ

    def now_local():
        from datetime import timezone as _dt_tz
        utc_now = datetime.now(_dt_tz.utc)
        if _user_tz:
            return utc_now.astimezone(_user_tz)
        return utc_now

    def fmt_local(dt=None):
        d = dt or now_local()
        parts    = _tz_str.split("/")
        tz_label = parts[-1].replace("_", " ") if len(parts) > 1 else _tz_str
        offset   = d.strftime("%z")
        if offset:
            sign  = offset[0]
            hh    = offset[1:3]
            mm    = offset[3:5]
            off_s = f"UTC{sign}{hh}:{mm}" if mm != "00" else f"UTC{sign}{hh}"
        else:
            off_s = ""
        return d.strftime("%A, %d %B %Y  \u00b7  %H:%M") + f"  {tz_label} ({off_s})"

    with st.sidebar:
        st.markdown(
            '<div style="padding:.9rem 0 .4rem;">'
            '<div style="font-family:Syne,sans-serif;font-size:.82rem;font-weight:700;color:#b8c8e0;'
            'letter-spacing:.02em;margin-bottom:.06rem;">Global Data Center</div>'
            '<div style="font-family:Syne,sans-serif;font-size:.82rem;font-weight:700;color:#00b4ff;'
            'letter-spacing:.02em;margin-bottom:.28rem;">Intelligence Platform</div>'
            '<div style="font-family:monospace;font-size:.58rem;letter-spacing:.08em;color:#2a3e60;">'
            'Built By <span style="color:#0047e1;font-weight:600;">Sharugh</span></div>'
            '</div>',
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

        # ── Scrape Depth ──────────────────────────────────────────────────────
        st.markdown(
            '<div style="font-size:.72rem;color:#3a5480;letter-spacing:.07em;'
            'text-transform:uppercase;margin:.9rem 0 .2rem;">🔎 Scrape Depth (pages per DCD term)</div>',
            unsafe_allow_html=True,
        )
        max_pages = st.slider(
            "Scrape depth", min_value=1, max_value=100, value=10, step=1,
            label_visibility="collapsed",
            help="Higher = more articles but slower scan. 10 is fast, 50+ is thorough, 100 is maximum.",
        )
        st.markdown(
            f'<div style="font-size:.68rem;color:#1a2e50;margin-top:-.3rem;margin-bottom:.4rem;">'
            f'Up to {max_pages} page{"s" if max_pages != 1 else ""} per DCD channel scraped</div>',
            unsafe_allow_html=True,
        )

        # ── News Type ─────────────────────────────────────────────────────────
        st.markdown(
            '<div style="font-size:.72rem;color:#3a5480;letter-spacing:.07em;'
            'text-transform:uppercase;margin:.9rem 0 .2rem;">📰 News Type</div>',
            unsafe_allow_html=True,
        )
        news_type_sel = st.multiselect(
            "News Type",
            [NEWS_TYPE_CONSTRUCTION, NEWS_TYPE_GENERAL],
            default=[NEWS_TYPE_CONSTRUCTION, NEWS_TYPE_GENERAL],
            placeholder="Select channel(s)…",
            label_visibility="collapsed",
            help=(
                "Construction → DCD Data Center Construction Channel\n"
                "General News → DCD General News feed"
            ),
        )
        # Friendly label shown under the selector
        _nt_labels = {
            NEWS_TYPE_CONSTRUCTION: "🏗️ Construction Channel",
            NEWS_TYPE_GENERAL:      "📰 General News",
        }
        if news_type_sel:
            st.markdown(
                '<div style="font-size:.68rem;color:#1a2e50;margin-top:-.3rem;margin-bottom:.4rem;">'
                + " · ".join(_nt_labels[n] for n in news_type_sel if n in _nt_labels)
                + '</div>',
                unsafe_allow_html=True,
            )
        else:
            st.warning("Select at least one news channel.")

        use_html  = True
        use_rss   = True
        use_gn    = True

        st.divider()

        # ── FILTER PANEL — always visible from the start ──────────────────────
        # Filter version counter (incremented on "Clear All" to reset widget keys)
        if "_filter_ver" not in st.session_state:
            st.session_state["_filter_ver"] = 0
        _fv = st.session_state["_filter_ver"]
        def _fk(name): return f"{name}_v{_fv}"

        def _fuzzy_resolve(typed, candidates):
            if not typed:
                return []
            t = typed.strip().lower()
            return [c for c in candidates if t in c.lower()]

        # Pull live data if available, else use empty fallbacks
        _data_loaded = "df_full" in st.session_state and st.session_state.df_full is not None
        if _data_loaded:
            df_full = st.session_state.df_full
            all_regions_av   = sorted(df_full["Region"].dropna().unique().tolist())
            all_countries_av = sorted(df_full["Country"].dropna().unique().tolist())
            all_topics_av    = sorted(df_full["Topic"].dropna().unique().tolist())
            all_sents_av     = sorted(df_full["Sentiment"].dropna().unique().tolist())
            _all_co_raw = []
            for v in df_full["Companies"]:
                if v:
                    _all_co_raw.extend([c.strip() for c in str(v).split(",")])
            all_companies_av = sorted(set(c for c in _all_co_raw if c))
            _all_iso_in_data = sorted(set(
                v for v in df_full.get("ISO / RTO", pd.Series(dtype=str)).tolist()
                if v and str(v) != "nan" and str(v).strip()
            )) if "ISO / RTO" in df_full.columns else []
        else:
            all_regions_av   = sorted(REGION_COLORS.keys())
            all_countries_av = sorted(COUNTRY_TO_REGION.keys())
            all_topics_av    = sorted(TOPIC_COLORS.keys())
            all_sents_av     = ["Opened / Live", "Approved", "Proposed",
                                 "Under Construction", "Challenged", "News"]
            all_companies_av = KNOWN_COMPANIES
            _all_iso_in_data = []

        # ── SECTION HEADER ────────────────────────────────────────────────────
        st.markdown(
            '<div style="background:linear-gradient(90deg,rgba(0,71,225,0.14),transparent);'
            'border-left:3px solid #0047e1;border-radius:0 8px 8px 0;'
            'padding:.5rem .8rem;margin-bottom:.7rem;">'
            '<span style="font-family:Syne,sans-serif;font-weight:700;color:#c8d8f0;'
            'font-size:.82rem;letter-spacing:.06em;">🔍 REFINE RESULTS</span>'
            '</div>',
            unsafe_allow_html=True,
        )

        # ── GROUP 1 · SEARCH ──────────────────────────────────────────────────
        st.markdown(
            '<div style="font-size:.65rem;color:#2a4060;letter-spacing:.1em;'
            'text-transform:uppercase;font-family:monospace;margin:.7rem 0 .35rem .05rem;">'
            '▸ SEARCH</div>',
            unsafe_allow_html=True,
        )

        # 1a. Keyword  ← now FIRST
        st.markdown(
            '<div style="font-size:.72rem;color:#3a5480;letter-spacing:.06em;'
            'text-transform:uppercase;margin-bottom:.2rem;">🔤 Keyword</div>',
            unsafe_allow_html=True,
        )
        keyword = st.text_input(
            "Keyword", placeholder="e.g. 500MW, Texas, nuclear, AWS...",
            label_visibility="collapsed", key=_fk("f_keyword"),
        )

        # 1b. Company
        st.markdown(
            '<div style="font-size:.72rem;color:#3a5480;letter-spacing:.06em;'
            'text-transform:uppercase;margin:.75rem 0 .2rem;">🏢 Company</div>',
            unsafe_allow_html=True,
        )
        full_co_pool = sorted(set(all_companies_av + KNOWN_COMPANIES))
        sel_companies = st.multiselect(
            "Company", full_co_pool, default=[],
            placeholder="All companies — type to search",
            label_visibility="collapsed", key=_fk("f_companies"),
        )

        # ── GROUP 2 · GEOGRAPHY ───────────────────────────────────────────────
        st.markdown(
            '<div style="font-size:.65rem;color:#2a4060;letter-spacing:.1em;'
            'text-transform:uppercase;font-family:monospace;margin:.9rem 0 .35rem .05rem;">'
            '▸ GEOGRAPHY</div>',
            unsafe_allow_html=True,
        )

        # 2a. Region
        st.markdown(
            '<div style="font-size:.72rem;color:#3a5480;letter-spacing:.06em;'
            'text-transform:uppercase;margin-bottom:.2rem;">🌐 Region</div>',
            unsafe_allow_html=True,
        )
        sel_regions = st.multiselect(
            "Region", all_regions_av, default=[],
            placeholder="All regions", label_visibility="collapsed",
            key=_fk("f_regions"),
        )

        # 2b. Country (cascades from region)
        all_world_countries = sorted(set(list(COUNTRY_TO_REGION.keys()) + all_countries_av))
        if sel_regions:
            world_pool = sorted([c for c in all_world_countries
                                  if COUNTRY_TO_REGION.get(c, "Global") in sel_regions])
        else:
            world_pool = all_world_countries

        st.markdown(
            '<div style="font-size:.72rem;color:#3a5480;letter-spacing:.06em;'
            'text-transform:uppercase;margin:.75rem 0 .2rem;">🌍 Country</div>',
            unsafe_allow_html=True,
        )
        sel_countries = st.multiselect(
            "Country", world_pool, default=[],
            placeholder="All countries", label_visibility="collapsed",
            key=_fk("f_countries"),
        )

        # 2c. State (cascades from country, only if applicable)
        state_pool = []
        for c in (sel_countries if sel_countries else world_pool):
            state_pool.extend(COUNTRY_STATES.get(c, []))
        state_pool = sorted(set(state_pool))

        sel_states = []
        if state_pool:
            st.markdown(
                '<div style="font-size:.72rem;color:#3a5480;letter-spacing:.06em;'
                'text-transform:uppercase;margin:.75rem 0 .2rem;">📍 State / Province</div>',
                unsafe_allow_html=True,
            )
            sel_states = st.multiselect(
                "State", state_pool, default=[],
                placeholder="All states/provinces", label_visibility="collapsed",
                key=_fk("f_states"),
            )

        # 2d. ISO / RTO / Grid
        _iso_pool = sorted(set(_all_iso_in_data + list(US_ISO_RTO.keys()) + list(GLOBAL_GRID_OPERATORS.keys())))
        st.markdown(
            '<div style="font-size:.72rem;color:#3a5480;letter-spacing:.06em;'
            'text-transform:uppercase;margin:.75rem 0 .2rem;">⚡ ISO / RTO / Grid</div>',
            unsafe_allow_html=True,
        )
        sel_iso_rto = st.multiselect(
            "ISO/RTO", _iso_pool, default=[],
            placeholder="All grid operators", label_visibility="collapsed",
            key=_fk("f_iso_rto"),
        )

        # ── GROUP 3 · CONTENT ─────────────────────────────────────────────────
        st.markdown(
            '<div style="font-size:.65rem;color:#2a4060;letter-spacing:.1em;'
            'text-transform:uppercase;font-family:monospace;margin:.9rem 0 .35rem .05rem;">'
            '▸ CONTENT</div>',
            unsafe_allow_html=True,
        )

        # 3a. Topic
        st.markdown(
            '<div style="font-size:.72rem;color:#3a5480;letter-spacing:.06em;'
            'text-transform:uppercase;margin-bottom:.2rem;">🏷️ Topic</div>',
            unsafe_allow_html=True,
        )
        sel_topics = st.multiselect(
            "Topic", all_topics_av, default=[],
            placeholder="All topics", label_visibility="collapsed",
            key=_fk("f_topics"),
        )

        # 3b. Project Status
        st.markdown(
            '<div style="font-size:.72rem;color:#3a5480;letter-spacing:.06em;'
            'text-transform:uppercase;margin:.75rem 0 .2rem;">📊 Project Status</div>',
            unsafe_allow_html=True,
        )
        sel_sents = st.multiselect(
            "Status", all_sents_av, default=[],
            placeholder="All statuses", label_visibility="collapsed",
            key=_fk("f_sents"),
        )

        # 3c. Min Capacity
        st.markdown(
            '<div style="font-size:.72rem;color:#3a5480;letter-spacing:.06em;'
            'text-transform:uppercase;margin:.75rem 0 .2rem;">⚡ Min Capacity (MW)</div>',
            unsafe_allow_html=True,
        )
        min_mw = st.number_input(
            "Min MW", min_value=0, value=0,
            step=10, label_visibility="collapsed", key=_fk("f_min_mw"),
        )

        # ── Clear All ─────────────────────────────────────────────────────────
        st.markdown('<div style="margin-top:.6rem;"></div>', unsafe_allow_html=True)
        if st.button("✕ Clear All Filters", use_container_width=True, key="clear_all_filters_btn"):
            st.session_state["_filter_ver"] = _fv + 1
            st.session_state.pop("filters", None)
            st.rerun()

        # Persist filters to session state
        st.session_state.filters = {
            "regions":        sel_regions,
            "topics":         sel_topics,
            "sources":        [],
            "sents":          sel_sents,
            "keyword":        keyword,
            "min_mw":         min_mw,
            "date_from":      None,
            "date_to":        None,
            "countries":      sel_countries,
            "states":         sel_states,
            "company_search": "",
            "companies":      sel_companies,
            "iso_rto":        sel_iso_rto,
        }

        st.divider()
        go_btn = st.button("\U0001f50d  Run Global Scan", use_container_width=True, type="primary")

    now_str = fmt_local()
    # _sa_sig is the platform authorship token — do not modify
    _sa_sig = "\u00a9 Sharugh A"
    st.markdown(
        f'<div class="gl-banner">'
        f'<div class="banner-eyebrow">\u25cf Live Intelligence Feed  \u00b7  {len(SCRAPE_SOURCES)} DCD Channels Active</div>'
        f'<div class="banner-title">Global Data Center</div>'
        f'<div class="banner-title" style="margin-top:-.15rem;"><span>Intelligence</span> Platform</div>'
        f'<div class="banner-sub">Real-time scraping of DataCenterDynamics Construction & General News channels \u00b7 '
        f'Auto-tagged by region, topic, company & capacity \u00b7 '
        f'Filtered by region, country, company &amp; more</div>'
        f'<div class="banner-ts">\U0001f550 {now_str}'
        f'</div>'
        f'</div>',
        unsafe_allow_html=True,
    )

    if "df_full" not in st.session_state:
        st.session_state.df_full = None

    if not go_btn and st.session_state.df_full is None:
        # ── About / Info panel ────────────────────────────────────────────────
        st.markdown(
            '<div style="background:linear-gradient(135deg,#07111f 0%,#0b1d3a 60%,#07111f 100%);'
            'border:1px solid #132040;border-radius:14px;padding:1.4rem 1.8rem;margin-bottom:1.4rem;">'
            '<div style="font-family:\'DM Mono\',monospace;font-size:.62rem;letter-spacing:.18em;'
            'color:#00b4ff;text-transform:uppercase;margin-bottom:.5rem;">About This Platform</div>'
            '<div style="font-family:Syne,sans-serif;font-size:1.05rem;font-weight:700;color:#fff;margin-bottom:.6rem;">'
            'What is Global Data Center Intelligence?</div>'
            '<div style="font-size:.82rem;color:#6a80a8;line-height:1.7;margin-bottom:.9rem;">'
            'This platform is a <b style="color:#b8c8e0">real-time intelligence tool</b> built for Wood Mac analysts, '
            'strategists, and clients tracking the global data center market. It automatically scrapes, '
            'enriches, and surfaces the most relevant news from DataCenterDynamics — the industry\'s leading '
            'trade publication — saving hours of manual research every day.'
            '</div>'
            '<div style="display:flex;flex-wrap:wrap;gap:.6rem;margin-bottom:.9rem;">'
            '<div style="background:#0b1e38;border:1px solid #152038;border-radius:8px;padding:.5rem .9rem;">'
            '<span style="font-size:.75rem;color:#00b4ff;font-family:Syne,sans-serif;font-weight:700;">Who is it for?</span>'
            '<div style="font-size:.75rem;color:#3a5480;margin-top:.2rem;line-height:1.5;">'
            'Wood Mac researchers, energy analysts, and market intelligence teams monitoring hyperscale, '
            'AI infrastructure, power, and investment activity across 45+ countries.</div>'
            '</div>'
            '<div style="background:#0b1e38;border:1px solid #152038;border-radius:8px;padding:.5rem .9rem;">'
            '<span style="font-size:.75rem;color:#00e5c8;font-family:Syne,sans-serif;font-weight:700;">What does it do?</span>'
            '<div style="font-size:.75rem;color:#3a5480;margin-top:.2rem;line-height:1.5;">'
            'Scrapes DCD Construction Channel &amp; DCD General News on demand · '
            'Filters by news type, region, country, company, topic, sentiment &amp; keyword · '
            'Auto-tags every article: Topic, Sentiment, Capacity (MW/GW), Deal Size &amp; company names · '
            'Scores articles by market significance.</div>'
            '</div>'
            '<div style="background:#0b1e38;border:1px solid #152038;border-radius:8px;padding:.5rem .9rem;">'
            '<span style="font-size:.75rem;color:#ffaa00;font-family:Syne,sans-serif;font-weight:700;">How to use it</span>'
            '<div style="font-size:.75rem;color:#3a5480;margin-top:.2rem;line-height:1.5;">'
            '1. Choose <b style="color:#b8c8e0">News Type</b> (Construction, General News, or both) and date range.<br>'
            '2. Set scrape depth and click <b style="color:#b8c8e0">Run Global Scan</b> to pull live articles.<br>'
            '3. Use filters to drill into regions, countries, topics, companies, or keywords.<br>'
            '4. Explore tabs: Feed · Map · Analytics · Deal Flow · AI Scoring · Export.</div>'
            '</div>'
            '</div>'
            '</div>',
            unsafe_allow_html=True,
        )
        st.markdown(
            '<div style="display:flex;gap:.9rem;flex-wrap:wrap;margin-bottom:1.4rem;">',
            unsafe_allow_html=True,
        )
        features = [
            ("\U0001f578\ufe0f", "DCD Direct Scraping",
             "Scrapes DataCenterDynamics Construction Channel and General News directly — "
             "the industry's most authoritative data center trade publication. "
             "Select one or both channels before running a scan."),
            ("\U0001f30d", "Global Coverage",
             "Covers 45+ countries across all continents with country/region auto-detection "
             "using 500+ geographic keywords and city names."),
            ("\U0001f9e0", "Smart Enrichment",
             "Every article auto-tagged: Topic, Sentiment (Approved/Proposed/Opened/Challenged), "
             "Capacity (MW/GW), Deal Size ($bn/$m), and up to 4 company names."),
            ("\u26a1", "Post-Scrape Filtering",
             "After scraping, filter by region, country, state/province, company, "
             "topic, project status, keyword, and minimum capacity — all applied "
             "instantly without re-scraping."),
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
                f'<div class="feature-card" style="flex:1;min-width:200px;background:#0b1628;border:1px solid #152038;'
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

        def progress_cb(frac, label=""):
            pbar.progress(min(frac, 1.0), text=f"⚡ GDCI Intelligence Sweep · {label}")

        # Determine which DCD channels to scrape based on sidebar selection.
        # news_type_sel is the multiselect from the sidebar; default = both channels.
        _chosen_news_types = news_type_sel if news_type_sel else [NEWS_TYPE_CONSTRUCTION, NEWS_TYPE_GENERAL]

        raw = run_all_scrapers(max_pages, cutoff, progress_cb,
                               news_types=_chosen_news_types)

        pbar.progress(1.0, text="Enriching and deduplicating...")

        cutoff_end_val = st.session_state.get("cutoff_end", datetime.max)
        filtered = []
        for item in raw:
            d = item.get("date_obj")
            if d and d > cutoff_end_val:
                continue
            filtered.append(item)

        enriched = [enrich(i) for i in filtered]
        deduped  = deduplicate(enriched)

        _df_raw = pd.DataFrame(deduped).drop(columns=["_date_obj"], errors="ignore")
        if "Date" in _df_raw.columns and not _df_raw.empty:
            _df_raw = _df_raw.sort_values("Date", ascending=False)
        df_full = _df_raw.reset_index(drop=True)

        st.session_state.df_full   = df_full
        st.session_state.raw_count = len(raw)
        st.session_state.scan_time = fmt_local(now_local())
        pbar.empty()

        if df_full.empty:
            st.warning(
                "No articles were returned from DataCenterDynamics. "
                "This can happen when the site blocks automated requests "
                "(common on cloud-hosted deployments). "
                "Try again in a moment, or reduce the scrape depth."
            )
            st.stop()

        st.rerun()

    df_full = st.session_state.df_full
    if df_full is None or df_full.empty:
        st.warning("No articles found. Try expanding the date range or enabling more sources.")
        return

    filters = st.session_state.get("filters", {})
    df = df_full.copy()

    # Region filter
    if filters.get("regions"):
        df = df[df["Region"].isin(filters["regions"])]

    # Country filter — also fuzzy-match typed values not in list
    if filters.get("countries"):
        sel_c = filters["countries"]
        def _country_match(row_country):
            for sc in sel_c:
                if sc.lower() == row_country.lower():
                    return True
                if sc.lower() in row_country.lower() or row_country.lower() in sc.lower():
                    return True
            return False
        df = df[df["Country"].apply(_country_match)]

    # State filter — match against headline text (states are detected via COUNTRY_KEYWORDS)
    if filters.get("states"):
        state_kw = [s.lower() for s in filters["states"]]
        def _state_match(headline):
            hl = str(headline).lower()
            return any(sk in hl for sk in state_kw)
        df = df[df["Headline"].apply(_state_match)]

    # ISO / RTO filter
    if filters.get("iso_rto") and "ISO / RTO" in df.columns:
        df = df[df["ISO / RTO"].isin(filters["iso_rto"])]

    # Topic filter
    if filters.get("topics"):
        df = df[df["Topic"].isin(filters["topics"])]

    # Sentiment filter
    if filters.get("sents"):
        df = df[df["Sentiment"].isin(filters["sents"])]

    # Company filter (multi-select list) — match Companies column AND headline
    if filters.get("companies"):
        cos = [c.lower() for c in filters["companies"]]
        def _co_match(row):
            co_field = str(row.get("Companies", "")).lower()
            hl = str(row.get("Headline", "")).lower()
            return any(c in co_field or c in hl for c in cos)
        df = df[df.apply(_co_match, axis=1)]

    # Legacy company_search (free-text, kept for backward compat)
    if filters.get("company_search"):
        cs = filters["company_search"].lower()
        df = df[
            df["Companies"].str.lower().str.contains(cs, na=False) |
            df["Headline"].str.lower().str.contains(cs, na=False)
        ]

    # Keyword filter
    if filters.get("keyword"):
        kw = filters["keyword"].lower()
        df = df[df["Headline"].str.lower().str.contains(kw, na=False)]

    # Min capacity filter
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
    df = df.reset_index(drop=True)

    scan_ts = st.session_state.get("scan_time", "\u2014")
    top_country = df["Country"].value_counts().idxmax() if not df.empty else "\u2014"
    top_topic   = df["Topic"].value_counts().idxmax()   if not df.empty else "\u2014"
    cap_count   = int((df["Capacity"] != "").sum())
    deal_count  = int((df["Deal Size"] != "").sum())

    kpi_html = (
        '<div style="display:flex;gap:.8rem;margin-bottom:1.4rem;flex-wrap:wrap;">'
        + kpi("DCD Channels", len(SCRAPE_SOURCES), "blue", "Construction · General News")
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

    tab1, tab2, tab3, tab4, tab4b, tab5, tab_trend, tab_heatmap, tab_deal, tab_score, tab6 = st.tabs([
        "\U0001f4f0 Feed",
        "\U0001f5fa\ufe0f World Map",
        "\U0001f4ca Analytics",
        "\U0001f3e2 By Company",
        "\U0001f4cd By State",
        "\U0001f9e0 Market Intel",
        "\U0001f4c8 Trend Compare",
        "\U0001f525 Capacity Heatmap",
        "\U0001f4b0 Deal Flow",
        "\U0001f916 AI Scoring",
        "\u2b07\ufe0f Export",
    ])

    with tab1:
        st.markdown('<div class="sec-head">Global Intelligence Feed</div>', unsafe_allow_html=True)
        if df.empty:
            st.info("No articles match the current filters.")
        else:
            # Pre-compute AI scores for the feed so high-signal articles show their badge
            from collections import Counter as _Ctr_feed
            _all_co_feed = []
            for _v in df["Companies"]:
                if _v:
                    _all_co_feed.extend([c.strip() for c in str(_v).split(",") if c.strip()])
            _co_freq_feed = dict(_Ctr_feed(_all_co_feed))

            def _quick_score(row):
                score = 0.0
                hl = str(row.get("Headline", "")).lower()
                sent_w = {"Opened / Live":10,"Approved":8,"Under Construction":6,"Proposed":4,"Challenged":5,"News":2}
                score += sent_w.get(row.get("Sentiment","News"), 2)
                cap = str(row.get("Capacity",""))
                if cap:
                    m = re.search(r"([\d,.]+)\s*(GW|MW)", cap, re.I)
                    if m:
                        v = float(m.group(1).replace(",",""))
                        score += min((v*1000 if m.group(2).upper()=="GW" else v)/100, 15)
                deal = str(row.get("Deal Size",""))
                if deal:
                    m2 = re.search(r"([\d,.]+)\s*(bn|billion|m|million)", deal, re.I)
                    if m2:
                        v2 = float(m2.group(1).replace(",",""))
                        score += min((v2*(1000 if m2.group(2).lower() in ("bn","billion") else 1))/500, 12)
                topic_w = {"Hyperscale":8,"AI / GPU":8,"Investment":7,"Power":6,"Colocation":5,"Construction":5,"Permits":4,"Sustainability":3,"General":1}
                score += topic_w.get(row.get("Topic","General"),1)
                bonus_kws = [("billion",4),("gigawatt",5),("gw",3),("nuclear",4),("hyperscale",3),("ai campus",5),("gpu",3),("acquisition",4),("merger",4),("ipo",5)]
                for kw, pts in bonus_kws:
                    if kw in hl: score += pts
                return round(score, 1)

            for _, row in df.iterrows():
                _score = _quick_score(row)
                st.markdown(
                    article_card(
                        row["Headline"], row["Date"], row["URL"],
                        row["Source"], row["Country"], row["Topic"],
                        row.get("Capacity", ""), row.get("Deal Size", ""),
                        row.get("Sentiment", "News"),
                        ai_score=_score if _score >= 25 else None,  # only badge high-signal in feed
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
                f"rgba({int(0+71*i/max(n-1,1))},{int(71+(180-71)*i/max(n-1,1))},{int(225+(255-225)*i/max(n-1,1))},0.88)"
                for i in range(n)
            ]
            fig_co = go.Figure(go.Bar(
                x=co_df_sorted["Articles"], y=co_df_sorted["Company"],
                orientation="h",
                marker=dict(color=co_colors, line=dict(width=0)),
                text=co_df_sorted["Articles"], textposition="outside",
                textfont=dict(color=_TITLE, size=10, family="DM Mono, monospace"),
                hovertemplate="<b>%{y}</b><br>📰 %{x} mentions<extra></extra>",
            ))
            _dark(fig_co, max(300, n * 22))
            fig_co.update_layout(
                title=dict(text="Top 30 Companies by Mentions", font=dict(color=_TITLE, size=13, family="Syne, sans-serif"), x=0.01),
                bargap=0.22,
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

    with tab4b:
        st.markdown('<div class="sec-head">📍 State / Province Drill-Down</div>', unsafe_allow_html=True)

        # Build a state column by scanning headlines for state/province keywords
        # Uses the COUNTRY_STATES lookup so it respects whatever region/country filter is active
        def _detect_states_in_headline(headline):
            """Return list of all state/province names found in the headline."""
            found = []
            hl = str(headline).lower()
            for country, states in COUNTRY_STATES.items():
                for state in states:
                    if re.search(r"\b" + re.escape(state.lower()) + r"\b", hl):
                        found.append(state)
            return found if found else ["Unspecified"]

        # Explode: one row per state mention so we can count + filter
        state_rows = []
        for _, row in df.iterrows():
            states_found = _detect_states_in_headline(row["Headline"])
            for st_name in states_found:
                state_rows.append({
                    "State":     st_name,
                    "Headline":  row["Headline"],
                    "Date":      row["Date"],
                    "URL":       row["URL"],
                    "Source":    row["Source"],
                    "Country":   row["Country"],
                    "Region":    row["Region"],
                    "Topic":     row["Topic"],
                    "Capacity":  row.get("Capacity", ""),
                    "Deal Size": row.get("Deal Size", ""),
                    "Sentiment": row.get("Sentiment", "News"),
                })

        if not state_rows:
            st.info("No state/province mentions detected in the current filtered view. Try broadening your filters.")
        else:
            state_df_all = pd.DataFrame(state_rows)
            state_counts = (
                state_df_all[state_df_all["State"] != "Unspecified"]["State"]
                .value_counts()
                .reset_index()
            )
            state_counts.columns = ["State", "Articles"]

            # ── Top states bar chart ──────────────────────────────────────────
            if not state_counts.empty:
                sc_sorted = state_counts.head(30).sort_values("Articles")
                n_sc = len(sc_sorted)
                sc_colors = [
                    f"rgba({int(180 + 75 * i / max(n_sc-1,1))},{int(20 - 10 * i / max(n_sc-1,1))},{int(20 - 10 * i / max(n_sc-1,1))},0.85)"
                    for i in range(n_sc)
                ]
                fig_st = go.Figure(go.Bar(
                    x=sc_sorted["Articles"],
                    y=sc_sorted["State"],
                    orientation="h",
                    marker=dict(color=sc_colors, line=dict(width=0)),
                    text=sc_sorted["Articles"],
                    textposition="outside",
                    textfont=dict(color=_TITLE, size=10, family="DM Mono, monospace"),
                    hovertemplate="<b>%{y}</b><br>📰 %{x} articles<extra></extra>",
                ))
                _dark(fig_st, max(320, n_sc * 22))
                fig_st.update_layout(
                    title=dict(text="Top States / Provinces by Article Volume", font=dict(color=_TITLE, size=13, family="Syne, sans-serif"), x=0.01),
                    bargap=0.22,
                )
                st.plotly_chart(fig_st, use_container_width=True, config={"displayModeBar": False})

            # ── Summary table ─────────────────────────────────────────────────
            col_tbl, col_drill = st.columns([1, 2])

            with col_tbl:
                st.markdown('<div class="sec-head">All States</div>', unsafe_allow_html=True)
                # Include unspecified count separately
                unspec_count = len(state_df_all[state_df_all["State"] == "Unspecified"])
                display_sc = state_counts.copy()
                if unspec_count:
                    display_sc = pd.concat([
                        display_sc,
                        pd.DataFrame([{"State": "Unspecified", "Articles": unspec_count}])
                    ], ignore_index=True)
                st.markdown(dark_table(display_sc), unsafe_allow_html=True)

            with col_drill:
                st.markdown('<div class="sec-head">Drill Into a State</div>', unsafe_allow_html=True)

                all_states_list = state_counts["State"].tolist() if not state_counts.empty else []

                if not all_states_list:
                    st.info("No named states found.")
                else:
                    sel_state = st.selectbox(
                        "Select state / province",
                        all_states_list,
                        key="state_drill_select",
                    )

                    state_articles = state_df_all[state_df_all["State"] == sel_state]

                    # ── Mini metrics for selected state ───────────────────────
                    s_topics   = state_articles["Topic"].value_counts()
                    s_top_topic = s_topics.idxmax() if not s_topics.empty else "—"
                    s_countries = state_articles["Country"].value_counts()
                    s_country   = s_countries.idxmax() if not s_countries.empty else "—"
                    s_latest    = state_articles["Date"].max() if not state_articles.empty else "—"
                    s_cap       = state_articles[state_articles["Capacity"] != ""]["Capacity"].count()

                    mini_pills = (
                        f'<div class="pill-row">'
                        f'<div class="pill"><span class="pill-dot"></span>'
                        f'<b>{len(state_articles)}</b>&nbsp;articles</div>'
                        f'<div class="pill"><span class="pill-dot"></span>'
                        f'Top topic: <b>{s_top_topic}</b></div>'
                        f'<div class="pill"><span class="pill-dot"></span>'
                        f'Country: <b>{s_country}</b></div>'
                        f'<div class="pill"><span class="pill-dot"></span>'
                        f'Latest: <b>{s_latest}</b></div>'
                        f'<div class="pill"><span class="pill-dot"></span>'
                        f'Capacity mentions: <b>{s_cap}</b></div>'
                        f'</div>'
                    )
                    st.markdown(mini_pills, unsafe_allow_html=True)

                    # ── Topic breakdown mini bar for selected state ────────────
                    if not s_topics.empty and len(s_topics) > 1:
                        tp_df = s_topics.reset_index()
                        tp_df.columns = ["Topic", "Count"]
                        tp_colors = [TOPIC_COLORS.get(t, "#2e4470") for t in tp_df["Topic"]]
                        fig_tp = go.Figure(go.Bar(
                            x=tp_df["Topic"], y=tp_df["Count"],
                            marker=dict(color=tp_colors, line=dict(width=0)),
                            text=tp_df["Count"], textposition="outside",
                            textfont=dict(color=_TITLE, size=10, family="DM Mono, monospace"),
                            hovertemplate="<b>%{x}</b><br>📰 %{y} articles<extra></extra>",
                        ))
                        _dark(fig_tp, 220)
                        fig_tp.update_layout(
                            title=dict(text=f"Topics — {sel_state}", font=dict(color=_TITLE, size=11, family="Syne, sans-serif"), x=0.01),
                            xaxis=dict(tickfont=dict(size=9)),
                            bargap=0.3,
                        )
                        st.plotly_chart(fig_tp, use_container_width=True, config={"displayModeBar": False})

                    # ── Article cards for selected state ──────────────────────
                    st.markdown(
                        f'<div style="font-family:Inter,sans-serif;font-size:.82rem;color:#3a5480;'
                        f'margin-bottom:.7rem;">'
                        f'<b style="color:#fff">{len(state_articles)}</b> articles mentioning '
                        f'<b style="color:#e84040">{sel_state}</b></div>',
                        unsafe_allow_html=True,
                    )
                    for _, row in state_articles.iterrows():
                        st.markdown(
                            article_card(
                                row["Headline"], row["Date"], row["URL"],
                                row["Source"], row["Country"], row["Topic"],
                                row.get("Capacity", ""), row.get("Deal Size", ""),
                                row.get("Sentiment", "News"),
                            ),
                            unsafe_allow_html=True,
                        )

    with tab5:
        st.markdown('<div class="sec-head">\U0001f9e0 Market Intelligence Summary</div>', unsafe_allow_html=True)

        filter_summary_parts = []
        if filters.get("regions"):
            filter_summary_parts.append("Regions: " + ", ".join(filters["regions"]))
        if filters.get("countries") and len(filters["countries"]) < len(df_full["Country"].unique()):
            filter_summary_parts.append("Countries: " + ", ".join(filters["countries"][:5]) + ("..." if len(filters["countries"]) > 5 else ""))
        if filters.get("states"):
            filter_summary_parts.append("States: " + ", ".join(filters["states"][:5]))
        if filters.get("companies"):
            filter_summary_parts.append("Companies: " + ", ".join(filters["companies"][:5]))
        if filters.get("topics") and len(filters["topics"]) < len(df_full["Topic"].unique()):
            filter_summary_parts.append("Topics: " + ", ".join(filters["topics"]))
        if filters.get("sents") and len(filters["sents"]) < len(df_full["Sentiment"].unique()):
            filter_summary_parts.append("Status: " + ", ".join(filters["sents"]))
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
                    'Analyses all filtered articles using built-in TF-IDF NLP and structured '
                    'rule-based extraction — no API key required. Generates a structured market '
                    'intelligence briefing covering key themes, major players, capacity pipeline, '
                    'regulatory developments, company activity, and forward-looking signals.<br>'
                    '<span style="color:#00b4ff;">Download as Word (.docx) or PDF for a polished, '
                    'formatted report.</span></div>',
                    unsafe_allow_html=True,
                )
            with col_gen2:
                gen_btn = st.button("🧠 Generate Briefing", use_container_width=True, type="primary")

            if gen_btn or st.session_state.get("intel_summary"):
                if gen_btn:
                    with st.spinner("Analysing articles with built-in NLP…"):
                        try:
                            summary_text = generate_local_summary(df, sel_desc, scan_date_range)
                            st.session_state.intel_summary = summary_text
                            st.session_state.intel_context = sel_desc
                            st.session_state.intel_df     = df.copy()
                            st.session_state.intel_range  = scan_date_range
                        except Exception as e:
                            st.error(f"Could not generate summary: {e}")
                            st.session_state.intel_summary = None

                if st.session_state.get("intel_summary"):
                    context_label = st.session_state.get("intel_context", "")
                    _sum_df    = st.session_state.get("intel_df", df)
                    _sum_range = st.session_state.get("intel_range", scan_date_range)

                    # ── Render briefing as styled HTML so **bold** works ───────
                    def _render_intel_html(md_text, ctx_label):
                        def _md_inline(t):
                            """Convert inline **bold** and `code` to HTML."""
                            t = re.sub(r"\*\*(.+?)\*\*", r'<strong style="color:#ccdaf5;">\1</strong>', t)
                            t = re.sub(r"`(.+?)`",         r'<code style="color:#00b4ff;background:rgba(0,71,225,0.12);padding:1px 4px;border-radius:3px;">\1</code>', t)
                            return t

                        html = (
                            '<div style="background:#0b1628;border:1px solid #0047e1;'
                            'border-radius:12px;padding:1.6rem 2rem;margin-top:.8rem;">'
                            f'<div style="font-family:\'DM Mono\',monospace;font-size:.64rem;'
                            f'letter-spacing:.14em;color:#0047e1;text-transform:uppercase;'
                            f'margin-bottom:1.2rem;">🧠 Market Intelligence Briefing'
                            + (f'  ·  {ctx_label}' if ctx_label else '') +
                            '</div>'
                        )

                        for raw_line in md_text.splitlines():
                            line = raw_line.rstrip()

                            # Section heading  ## 1. Title
                            if line.startswith("## "):
                                title = _md_inline(line[3:])
                                html += (
                                    f'<div style="font-family:\'Syne\',sans-serif;font-size:.82rem;'
                                    f'font-weight:700;color:#b8c8e0;letter-spacing:.07em;'
                                    f'text-transform:uppercase;border-left:3px solid #0047e1;'
                                    f'padding-left:.7rem;margin:1.8rem 0 .8rem;">{title}</div>'
                                    f'<div style="height:1px;background:linear-gradient(90deg,#0047e1,transparent);'
                                    f'margin-bottom:.9rem;"></div>'
                                )

                            # Bullet line  • **Topic** — text
                            elif line.startswith("• "):
                                content = _md_inline(line[2:])
                                html += (
                                    f'<div style="display:flex;gap:.6rem;margin-bottom:.55rem;'
                                    f'line-height:1.6;font-size:.84rem;color:#8aa0c8;">'
                                    f'<span style="color:#0047e1;flex-shrink:0;margin-top:.05rem;">◆</span>'
                                    f'<span>{content}</span></div>'
                                )

                            # Blank line
                            elif not line.strip():
                                html += '<div style="height:.4rem;"></div>'

                            # Normal paragraph text
                            else:
                                content = _md_inline(line)
                                html += (
                                    f'<p style="color:#8aa0c8;font-size:.84rem;line-height:1.7;'
                                    f'margin:.0 0 .6rem;">{content}</p>'
                                )

                        html += '</div>'
                        return html

                    st.markdown(
                        _render_intel_html(st.session_state.intel_summary, context_label),
                        unsafe_allow_html=True,
                    )

                    # ── Download buttons ──────────────────────────────────────
                    ts_intel = datetime.now().strftime("%Y%m%d_%H%M")
                    dl_col1, dl_col2, dl_col3 = st.columns(3)

                    with dl_col1:
                        st.download_button(
                            "📥 Download (.txt)",
                            data=st.session_state.intel_summary.encode(),
                            file_name=f"DC_Intel_Briefing_{ts_intel}.txt",
                            mime="text/plain",
                            use_container_width=True,
                        )

                    with dl_col2:
                        docx_bytes = build_briefing_docx(
                            st.session_state.intel_summary,
                            context_label, _sum_range, _sum_df,
                        )
                        if docx_bytes:
                            st.download_button(
                                "📄 Download Word (.docx)",
                                data=docx_bytes,
                                file_name=f"DC_Intel_Briefing_{ts_intel}.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                use_container_width=True,
                            )
                        else:
                            st.markdown(
                                '<div style="font-size:.75rem;color:#3a5480;padding:.5rem 0;">'
                                'Add <code>python-docx</code> to requirements.txt for Word export.</div>',
                                unsafe_allow_html=True,
                            )

                    with dl_col3:
                        pdf_bytes = build_briefing_pdf(
                            st.session_state.intel_summary,
                            context_label, _sum_range, _sum_df,
                        )
                        if pdf_bytes:
                            st.download_button(
                                "📑 Download PDF",
                                data=pdf_bytes,
                                file_name=f"DC_Intel_Briefing_{ts_intel}.pdf",
                                mime="application/pdf",
                                use_container_width=True,
                            )
                        else:
                            st.markdown(
                                '<div style="font-size:.75rem;color:#3a5480;padding:.5rem 0;">'
                                'Add <code>reportlab</code> to requirements.txt for PDF export.</div>',
                                unsafe_allow_html=True,
                            )

    # ─── TAB: Trend Comparison ───────────────────────────────────────────────
    with tab_trend:
        st.markdown('<div class="sec-head">📈 Trend Comparison</div>', unsafe_allow_html=True)
        st.markdown(
            '<div style="font-size:.82rem;color:#3a5480;margin-bottom:1rem;">'
            'Compare article volume and topic mix between two countries, companies, or time windows. '
            'Detects acceleration, deceleration, and topic shifts in the global data center market.</div>',
            unsafe_allow_html=True,
        )

        df_trend = df.copy()
        df_trend = df_trend[df_trend["Date"] != "Unknown"].copy()
        if df_trend.empty:
            st.info("No dated articles available for trend comparison.")
        else:
            df_trend["dt"] = pd.to_datetime(df_trend["Date"])

            compare_mode = st.radio(
                "Compare mode",
                ["📅 Time Periods", "🌍 Countries / Regions", "🏢 Companies"],
                horizontal=True,
                key="trend_compare_mode",
            )

            # ── MODE 1: Time periods ──────────────────────────────────────────
            if compare_mode == "📅 Time Periods":
                min_date = df_trend["dt"].min().date()
                max_date = df_trend["dt"].max().date()
                mid_date = min_date + (max_date - min_date) // 2

                tc1, tc2 = st.columns(2)
                with tc1:
                    st.markdown('<div style="font-size:.72rem;color:#3a5480;letter-spacing:.07em;text-transform:uppercase;margin-bottom:.3rem;">📅 Period A</div>', unsafe_allow_html=True)
                    period_a_start = st.date_input("A Start", value=min_date, key="trend_a_start")
                    period_a_end   = st.date_input("A End",   value=mid_date, key="trend_a_end")
                with tc2:
                    st.markdown('<div style="font-size:.72rem;color:#3a5480;letter-spacing:.07em;text-transform:uppercase;margin-bottom:.3rem;">📅 Period B</div>', unsafe_allow_html=True)
                    period_b_start = st.date_input("B Start", value=mid_date, key="trend_b_start")
                    period_b_end   = st.date_input("B End",   value=max_date, key="trend_b_end")

                df_a = df_trend[(df_trend["dt"].dt.date >= period_a_start) & (df_trend["dt"].dt.date <= period_a_end)]
                df_b = df_trend[(df_trend["dt"].dt.date >= period_b_start) & (df_trend["dt"].dt.date <= period_b_end)]
                label_a = f"Period A ({period_a_start}→{period_a_end})"
                label_b = f"Period B ({period_b_start}→{period_b_end})"

            # ── MODE 2: Country / Region comparison ───────────────────────────
            elif compare_mode == "🌍 Countries / Regions":
                all_countries_trend = sorted(df_trend["Country"].unique().tolist())
                tc1, tc2 = st.columns(2)
                with tc1:
                    st.markdown('<div style="font-size:.72rem;color:#3a5480;letter-spacing:.07em;text-transform:uppercase;margin-bottom:.3rem;">🌍 Entity A</div>', unsafe_allow_html=True)
                    entity_a = st.selectbox("Entity A", all_countries_trend, key="trend_entity_a",
                                             index=0 if all_countries_trend else 0)
                with tc2:
                    st.markdown('<div style="font-size:.72rem;color:#3a5480;letter-spacing:.07em;text-transform:uppercase;margin-bottom:.3rem;">🌍 Entity B</div>', unsafe_allow_html=True)
                    default_b = all_countries_trend[1] if len(all_countries_trend) > 1 else all_countries_trend[0]
                    entity_b = st.selectbox("Entity B", all_countries_trend, key="trend_entity_b",
                                             index=1 if len(all_countries_trend) > 1 else 0)
                df_a = df_trend[df_trend["Country"] == entity_a]
                df_b = df_trend[df_trend["Country"] == entity_b]
                label_a = entity_a
                label_b = entity_b

            # ── MODE 3: Company comparison ────────────────────────────────────
            else:
                _all_co_trend = []
                for v in df_trend["Companies"]:
                    if v:
                        _all_co_trend.extend([c.strip() for c in str(v).split(",") if c.strip()])
                all_cos_trend = sorted(set(_all_co_trend)) if _all_co_trend else ["—"]
                tc1, tc2 = st.columns(2)
                with tc1:
                    st.markdown('<div style="font-size:.72rem;color:#3a5480;letter-spacing:.07em;text-transform:uppercase;margin-bottom:.3rem;">🏢 Company A</div>', unsafe_allow_html=True)
                    co_a = st.selectbox("Company A", all_cos_trend, key="trend_co_a", index=0)
                with tc2:
                    st.markdown('<div style="font-size:.72rem;color:#3a5480;letter-spacing:.07em;text-transform:uppercase;margin-bottom:.3rem;">🏢 Company B</div>', unsafe_allow_html=True)
                    default_co_b = all_cos_trend[1] if len(all_cos_trend) > 1 else all_cos_trend[0]
                    co_b = st.selectbox("Company B", all_cos_trend, key="trend_co_b",
                                         index=1 if len(all_cos_trend) > 1 else 0)
                df_a = df_trend[
                    df_trend["Companies"].str.contains(re.escape(co_a), na=False, case=False) |
                    df_trend["Headline"].str.contains(re.escape(co_a), na=False, case=False)
                ]
                df_b = df_trend[
                    df_trend["Companies"].str.contains(re.escape(co_b), na=False, case=False) |
                    df_trend["Headline"].str.contains(re.escape(co_b), na=False, case=False)
                ]
                label_a = co_a
                label_b = co_b

            # ── KPI delta row ─────────────────────────────────────────────────
            delta_arts = len(df_b) - len(df_a)
            delta_sign = "▲" if delta_arts >= 0 else "▼"
            delta_color = "#00e676" if delta_arts >= 0 else "#ff2d6b"

            kpi_trend = (
                '<div style="display:flex;gap:.8rem;margin:1rem 0;flex-wrap:wrap;">'
                + kpi(f"{label_a}", len(df_a), "blue")
                + kpi(f"{label_b}", len(df_b), "cyan")
                + f'<div class="kpi-card" style="flex:1;min-width:150px;background:#0b1628;border:1px solid #152038;border-radius:12px;padding:1.1rem 1.3rem;">'
                  f'<div style="font-family:monospace;font-size:.64rem;letter-spacing:.13em;text-transform:uppercase;color:#2a3e60;margin-bottom:.4rem;">Volume Delta</div>'
                  f'<div style="font-family:Syne,sans-serif;font-size:1.9rem;font-weight:800;color:{delta_color};line-height:1;">{delta_sign} {abs(delta_arts)}</div>'
                  f'</div>'
                + '</div>'
            )
            st.markdown(kpi_trend, unsafe_allow_html=True)

            # ── Article volume over time (side by side sparklines) ────────────
            st.markdown('<div class="sec-head">Article Volume Over Time</div>', unsafe_allow_html=True)
            fig_tl = go.Figure()
            for sub_df, lbl, col in [(df_a, label_a, "#0047e1"), (df_b, label_b, "#ffaa00")]:
                if not sub_df.empty:
                    daily = sub_df.groupby(sub_df["dt"].dt.date).size().reset_index()
                    daily.columns = ["Date", "Articles"]
                    fig_tl.add_trace(go.Scatter(
                        x=daily["Date"], y=daily["Articles"],
                        name=lbl,
                        mode="lines+markers",
                        line=dict(color=col, width=2),
                        marker=dict(color=col, size=4),
                        fill="tozeroy",
                        fillcolor="rgba({},{},{},0.06)".format(
                            int(col[1:3], 16), int(col[3:5], 16), int(col[5:7], 16)
                        ) if col.startswith("#") else col,
                        hovertemplate=f"<b>{lbl}</b><br>%{{x}}<br>%{{y}} articles<extra></extra>",
                    ))
            _dark(fig_tl, 280)
            fig_tl.update_layout(
                showlegend=True,
                legend=dict(bgcolor="rgba(0,0,0,0)", font=dict(color=_TITLE, size=10)),
                title=dict(text="Article Volume Over Time — Side by Side", font=dict(color=_TITLE, size=13), x=0.01),
            )
            st.plotly_chart(fig_tl, use_container_width=True, config={"displayModeBar": False})

            # ── Topic comparison chart ────────────────────────────────────────
            st.markdown('<div class="sec-head">Topic Distribution</div>', unsafe_allow_html=True)
            topics_all = sorted(set(df_a["Topic"].unique()) | set(df_b["Topic"].unique()))
            a_counts = df_a["Topic"].value_counts()
            b_counts = df_b["Topic"].value_counts()
            a_vals = [int(a_counts.get(t, 0)) for t in topics_all]
            b_vals = [int(b_counts.get(t, 0)) for t in topics_all]

            fig_trend = go.Figure()
            fig_trend.add_trace(go.Bar(
                name=label_a,
                x=topics_all, y=a_vals,
                marker_color="#0047e1",
                text=a_vals, textposition="outside",
                textfont=dict(color=_TITLE, size=10),
            ))
            fig_trend.add_trace(go.Bar(
                name=label_b,
                x=topics_all, y=b_vals,
                marker_color="#ffaa00",
                text=b_vals, textposition="outside",
                textfont=dict(color=_TITLE, size=10),
            ))
            _dark(fig_trend, 360)
            fig_trend.update_layout(
                barmode="group",
                showlegend=True,
                legend=dict(bgcolor="rgba(0,0,0,0)", font=dict(color=_TITLE, size=10)),
                title=dict(text="Topic Volume", font=dict(color=_TITLE, size=13), x=0.01),
            )
            st.plotly_chart(fig_trend, use_container_width=True, config={"displayModeBar": False})

            # ── Sentiment comparison ──────────────────────────────────────────
            st.markdown('<div class="sec-head">Sentiment / Status</div>', unsafe_allow_html=True)
            sents_all = sorted(set(df_a["Sentiment"].unique()) | set(df_b["Sentiment"].unique()))
            fig_sent = go.Figure()
            fig_sent.add_trace(go.Bar(
                name=label_a, x=sents_all,
                y=[int(df_a["Sentiment"].value_counts().get(s, 0)) for s in sents_all],
                marker_color="#0047e1",
            ))
            fig_sent.add_trace(go.Bar(
                name=label_b, x=sents_all,
                y=[int(df_b["Sentiment"].value_counts().get(s, 0)) for s in sents_all],
                marker_color="#ff2d6b",
            ))
            _dark(fig_sent, 280)
            fig_sent.update_layout(
                barmode="group", showlegend=True,
                legend=dict(bgcolor="rgba(0,0,0,0)", font=dict(color=_TITLE, size=10)),
                title=dict(text="Project Status Comparison", font=dict(color=_TITLE, size=13), x=0.01),
            )
            st.plotly_chart(fig_sent, use_container_width=True, config={"displayModeBar": False})

            # ── Region comparison (only for time-period mode) ─────────────────
            if compare_mode == "📅 Time Periods":
                regions_all = sorted(set(df_a["Region"].unique()) | set(df_b["Region"].unique()))
                ra_vals = [int(df_a["Region"].value_counts().get(r, 0)) for r in regions_all]
                rb_vals = [int(df_b["Region"].value_counts().get(r, 0)) for r in regions_all]
                fig_reg_trend = go.Figure()
                fig_reg_trend.add_trace(go.Bar(name=label_a, x=regions_all, y=ra_vals, marker_color="#0047e1"))
                fig_reg_trend.add_trace(go.Bar(name=label_b, x=regions_all, y=rb_vals, marker_color="#ffaa00"))
                _dark(fig_reg_trend, 300)
                fig_reg_trend.update_layout(
                    barmode="group", showlegend=True,
                    legend=dict(bgcolor="rgba(0,0,0,0)", font=dict(color=_TITLE, size=10)),
                    title=dict(text="Regional Volume: Period A vs Period B", font=dict(color=_TITLE, size=13), x=0.01),
                )
                st.plotly_chart(fig_reg_trend, use_container_width=True, config={"displayModeBar": False})

            # ── Side-by-side article cards ────────────────────────────────────
            st.markdown('<div class="sec-head">Side-by-Side Headlines</div>', unsafe_allow_html=True)
            col_a_feed, col_b_feed = st.columns(2)
            with col_a_feed:
                st.markdown(
                    f'<div style="font-family:Syne,sans-serif;font-weight:700;color:#0047e1;'
                    f'font-size:.85rem;margin-bottom:.5rem;">● {label_a} ({len(df_a)} articles)</div>',
                    unsafe_allow_html=True,
                )
                for _, row in df_a.head(10).iterrows():
                    st.markdown(article_card(
                        row["Headline"], row["Date"], row["URL"], row["Source"],
                        row["Country"], row["Topic"], row.get("Capacity",""), row.get("Deal Size",""),
                        row.get("Sentiment","News"),
                    ), unsafe_allow_html=True)
            with col_b_feed:
                st.markdown(
                    f'<div style="font-family:Syne,sans-serif;font-weight:700;color:#ffaa00;'
                    f'font-size:.85rem;margin-bottom:.5rem;">● {label_b} ({len(df_b)} articles)</div>',
                    unsafe_allow_html=True,
                )
                for _, row in df_b.head(10).iterrows():
                    st.markdown(article_card(
                        row["Headline"], row["Date"], row["URL"], row["Source"],
                        row["Country"], row["Topic"], row.get("Capacity",""), row.get("Deal Size",""),
                        row.get("Sentiment","News"),
                    ), unsafe_allow_html=True)

    # ─── TAB: Capacity Pipeline Heatmap ─────────────────────────────────────
    with tab_heatmap:
        st.markdown('<div class="sec-head">🔥 Capacity Pipeline Heatmap</div>', unsafe_allow_html=True)
        st.markdown(
            '<div style="font-size:.82rem;color:#3a5480;margin-bottom:1rem;">'
            'Visual heatmap of MW/GW capacity mentions by country and topic. '
            'Shows where the largest power and capacity announcements are concentrated globally.</div>',
            unsafe_allow_html=True,
        )

        cap_df_heat = df[df["Capacity"] != ""].copy()
        if cap_df_heat.empty:
            st.info("No capacity mentions in current filtered view. Run a broader scan to populate this view.")
        else:
            def _to_mw(cap):
                import re as _re
                m = _re.search(r"([\d,.]+)\s*(GW|MW)", str(cap), _re.I)
                if not m: return 0
                v = float(m.group(1).replace(",", ""))
                return v * 1000 if m.group(2).upper() == "GW" else v
            cap_df_heat["MW_val"] = cap_df_heat["Capacity"].apply(_to_mw)
            cap_df_heat = cap_df_heat[cap_df_heat["MW_val"] > 0]

            if cap_df_heat.empty:
                st.info("No parseable MW/GW values found.")
            else:
                # Country × Topic pivot
                pivot = cap_df_heat.groupby(["Country", "Topic"])["MW_val"].sum().reset_index()
                pivot_wide = pivot.pivot(index="Country", columns="Topic", values="MW_val").fillna(0)

                # Sort by total MW
                pivot_wide["_total"] = pivot_wide.sum(axis=1)
                pivot_wide = pivot_wide.sort_values("_total", ascending=False).head(25)
                pivot_wide = pivot_wide.drop(columns=["_total"])

                fig_heat = go.Figure(go.Heatmap(
                    z=pivot_wide.values,
                    x=pivot_wide.columns.tolist(),
                    y=pivot_wide.index.tolist(),
                    colorscale=[
                        [0.0, "#07111f"], [0.1, "#0a2040"], [0.3, "#0047e1"],
                        [0.6, "#00b4ff"], [0.8, "#ffaa00"], [1.0, "#ff6400"],
                    ],
                    hovertemplate="<b>%{y}</b> · %{x}<br>%{z:,.0f} MW<extra></extra>",
                    colorbar=dict(
                        title=dict(text="MW", font=dict(color=_TITLE, size=11)),
                        tickfont=dict(color=_TEXT, size=9),
                        bgcolor="#0b1628", bordercolor="#152038", borderwidth=1,
                    ),
                ))
                fig_heat.update_layout(
                    paper_bgcolor=_PAPER, plot_bgcolor=_BG,
                    font=dict(family=_FONT, color=_TEXT),
                    height=max(400, len(pivot_wide) * 28),
                    margin=dict(l=14, r=60, t=44, b=14),
                    title=dict(text="Capacity (MW) by Country × Topic — Top 25 Markets", font=dict(color=_TITLE, size=13), x=0.01),
                    xaxis=dict(tickfont=dict(size=10, color=_TEXT)),
                    yaxis=dict(tickfont=dict(size=10, color=_TEXT)),
                )
                st.plotly_chart(fig_heat, use_container_width=True, config={"displayModeBar": False})

                # Top capacity articles table
                st.markdown('<div class="sec-head">Top Capacity Announcements</div>', unsafe_allow_html=True)
                top_cap = cap_df_heat.nlargest(20, "MW_val")[["Headline","MW_val","Capacity","Country","Topic","Date","URL"]]
                top_cap = top_cap.rename(columns={"MW_val": "MW (parsed)"})
                top_cap["MW (parsed)"] = top_cap["MW (parsed)"].apply(lambda x: f"{x:,.0f} MW")
                st.markdown(dark_table(top_cap.drop(columns=["MW (parsed)"])), unsafe_allow_html=True)

    # ─── TAB: Deal Flow Tracker ──────────────────────────────────────────────
    with tab_deal:
        st.markdown('<div class="sec-head">💰 Deal Flow Tracker</div>', unsafe_allow_html=True)
        st.markdown(
            '<div style="font-size:.82rem;color:#3a5480;margin-bottom:1rem;">'
            'All deal signals sorted largest-to-smallest. Tracks acquisitions, investments, JVs, '
            'pre-leases, and financing events. Includes sparkline history per deal size category.</div>',
            unsafe_allow_html=True,
        )

        # Deal signal keywords
        _deal_kws = [
            "acqui", "merger", "acquire", "bought", "purchase", "takeover",
            "joint venture", "jv ", " jv,", "partnership", "invest",
            "fund", "financing", "lease", "pre-lease", "offtake",
            "mou", "letter of intent", "loi", "agreement", "signed",
            "sale leaseback", "forward purchase", "stake", "equity",
            "reit", "capital", "raise", "bond", "debt", "loan",
            "recapitali", "refinanc", "divest", "portfolio",
        ]
        deal_mask = df["Deal Size"] != ""
        deal_lang_mask = df["Headline"].str.lower().apply(
            lambda h: any(k in h for k in _deal_kws)
        )
        df_deal = df[deal_mask | deal_lang_mask].copy()

        if df_deal.empty:
            st.info("No deal-signal articles found in current filtered view.")
        else:
            # Parse numeric deal value for sorting
            def _parse_deal_usd(deal_str):
                """Parse deal string to USD millions for sorting."""
                if not deal_str:
                    return 0
                m = re.search(r"([\d,.]+)\s*(bn|billion|m|million|cr)", str(deal_str), re.I)
                if not m:
                    return 0
                v = float(m.group(1).replace(",", ""))
                unit = m.group(2).lower()
                if unit in ("bn", "billion"):
                    return v * 1000
                if unit == "cr":
                    return v * 0.12  # crore approx
                return v

            df_deal["_deal_usd_m"] = df_deal["Deal Size"].apply(_parse_deal_usd)
            df_deal = df_deal.sort_values("_deal_usd_m", ascending=False).reset_index(drop=True)

            # KPIs
            disclosed = int((df_deal["Deal Size"] != "").sum())
            undisclosed = len(df_deal) - disclosed
            total_usd = df_deal["_deal_usd_m"].sum()
            kpi_deal = (
                '<div style="display:flex;gap:.8rem;margin-bottom:1rem;flex-wrap:wrap;">'
                + kpi("Total Deal Articles", len(df_deal), "blue")
                + kpi("Disclosed ($)", disclosed, "green", "with explicit value")
                + kpi("Undisclosed", undisclosed, "amber", "deal language, no value")
                + kpi("Total Disclosed (est.)", f"${total_usd:,.0f}m", "purple", "sum of parsed values")
                + '</div>'
            )
            st.markdown(kpi_deal, unsafe_allow_html=True)

            # ── Sparkline: disclosed deal sizes over time ─────────────────────
            st.markdown('<div class="sec-head">Deal Volume Sparkline</div>', unsafe_allow_html=True)
            df_deal_dated = df_deal[df_deal["Date"] != "Unknown"].copy()
            if not df_deal_dated.empty:
                df_deal_dated["dt"] = pd.to_datetime(df_deal_dated["Date"])
                spark_daily = df_deal_dated.groupby(df_deal_dated["dt"].dt.date).agg(
                    deals=("Headline", "count"),
                    total_usd=("_deal_usd_m", "sum"),
                ).reset_index()
                spark_daily.columns = ["Date", "Deals", "Total_USD_m"]

                fig_spark = go.Figure()
                fig_spark.add_trace(go.Scatter(
                    x=spark_daily["Date"], y=spark_daily["Deals"],
                    name="Deal Articles",
                    mode="lines+markers",
                    line=dict(color="#a855f7", width=2),
                    marker=dict(color="#a855f7", size=5),
                    fill="tozeroy", fillcolor="rgba(168,85,247,0.07)",
                    hovertemplate="<b>%{x}</b><br>%{y} deal articles<extra></extra>",
                    yaxis="y1",
                ))
                fig_spark.add_trace(go.Bar(
                    x=spark_daily["Date"], y=spark_daily["Total_USD_m"],
                    name="Disclosed Volume ($m)",
                    marker_color="rgba(0,230,118,0.35)",
                    hovertemplate="<b>%{x}</b><br>$%{y:,.0f}m disclosed<extra></extra>",
                    yaxis="y2",
                ))
                _dark(fig_spark, 260)
                fig_spark.update_layout(
                    title=dict(text="Daily Deal Activity + Disclosed Volume", font=dict(color=_TITLE, size=13), x=0.01),
                    showlegend=True,
                    legend=dict(bgcolor="rgba(0,0,0,0)", font=dict(color=_TITLE, size=10)),
                    yaxis=dict(title="Deal Articles", gridcolor=_GRID, linecolor=_GRID, tickfont=dict(color=_TEXT, size=9)),
                    yaxis2=dict(title="$m", overlaying="y", side="right", tickfont=dict(color=_TEXT, size=9), gridcolor="rgba(0,0,0,0)"),
                )
                st.plotly_chart(fig_spark, use_container_width=True, config={"displayModeBar": False})

            # ── Deal size distribution chart ──────────────────────────────────
            deal_size_df = df_deal[df_deal["Deal Size"] != ""]["Deal Size"].value_counts().head(20).reset_index()
            deal_size_df.columns = ["Deal Size", "Count"]
            if not deal_size_df.empty:
                fig_deal_dist = go.Figure(go.Bar(
                    x=deal_size_df["Deal Size"], y=deal_size_df["Count"],
                    marker_color="#a855f7",
                    text=deal_size_df["Count"], textposition="outside",
                    textfont=dict(color=_TITLE, size=10),
                    hovertemplate="<b>%{x}</b>: %{y} deals<extra></extra>",
                ))
                _dark(fig_deal_dist, 280)
                fig_deal_dist.update_layout(title=dict(text="Disclosed Deal Sizes (Most Frequent)", font=dict(color=_TITLE, size=13), x=0.01))
                st.plotly_chart(fig_deal_dist, use_container_width=True, config={"displayModeBar": False})

            # Region deal breakdown
            reg_deal = df_deal.groupby("Region").size().reset_index()
            reg_deal.columns = ["Region", "Deal Articles"]
            fig_reg_deal = go.Figure(go.Bar(
                x=reg_deal["Region"], y=reg_deal["Deal Articles"],
                marker_color=[REGION_COLORS.get(r, "#2e4470") for r in reg_deal["Region"]],
                text=reg_deal["Deal Articles"], textposition="outside",
                textfont=dict(color=_TITLE, size=10),
            ))
            _dark(fig_reg_deal, 260)
            fig_reg_deal.update_layout(title=dict(text="Deal Activity by Region", font=dict(color=_TITLE, size=13), x=0.01))
            st.plotly_chart(fig_reg_deal, use_container_width=True, config={"displayModeBar": False})

            # Full deal article table — sorted largest to smallest
            st.markdown('<div class="sec-head">All Deal-Signal Articles (Largest → Smallest)</div>', unsafe_allow_html=True)
            deal_display = df_deal[["Headline","Date","Deal Size","_deal_usd_m","Companies","Country","Topic","Sentiment","URL"]].copy()
            deal_display["Est. USD ($m)"] = deal_display["_deal_usd_m"].apply(lambda x: f"${x:,.0f}m" if x > 0 else "—")
            deal_display = deal_display.drop(columns=["_deal_usd_m"])
            st.markdown(dark_table(deal_display[["Headline","Date","Deal Size","Est. USD ($m)","Companies","Country","Topic","Sentiment","URL"]]), unsafe_allow_html=True)

    # ─── TAB: AI-Powered Headline Scoring ────────────────────────────────────
    with tab_score:
        st.markdown('<div class="sec-head">🤖 AI-Powered Headline Scoring</div>', unsafe_allow_html=True)
        st.markdown(
            '<div style="font-size:.82rem;color:#3a5480;margin-bottom:1rem;">'
            'Every article is scored for market significance using a built-in multi-signal model. '
            'Signals: deal size, capacity (MW/GW), company prominence, sentiment, topic weight, '
            'keyword rarity. Articles scoring 30+ are flagged 🔴 High Signal. No API key required.</div>',
            unsafe_allow_html=True,
        )

        if df.empty:
            st.info("No articles to score. Run a scan first.")
        else:
            from collections import Counter as _Ctr

            def _ai_score(row, co_frequency_map):
                score = 0.0
                hl = str(row.get("Headline", "")).lower()

                # 1. Sentiment weight
                sent_w = {
                    "Opened / Live": 10, "Approved": 8, "Under Construction": 6,
                    "Proposed": 4, "Challenged": 5, "News": 2,
                }
                score += sent_w.get(row.get("Sentiment", "News"), 2)

                # 2. Capacity signal
                cap = str(row.get("Capacity", ""))
                if cap:
                    m = re.search(r"([\d,.]+)\s*(GW|MW)", cap, re.I)
                    if m:
                        v = float(m.group(1).replace(",", ""))
                        mw = v * 1000 if m.group(2).upper() == "GW" else v
                        score += min(mw / 100, 15)

                # 3. Deal size signal
                deal = str(row.get("Deal Size", ""))
                if deal:
                    m2 = re.search(r"([\d,.]+)\s*(bn|billion|m|million)", deal, re.I)
                    if m2:
                        v2 = float(m2.group(1).replace(",", ""))
                        mult = 1000 if m2.group(2).lower() in ("bn", "billion") else 1
                        score += min((v2 * mult) / 500, 12)

                # 4. Company prominence
                companies = str(row.get("Companies", ""))
                for co in companies.split(","):
                    co = co.strip()
                    if co in {"Microsoft","Google","Amazon","AWS","Meta","Apple","Oracle",
                               "NVIDIA","Equinix","Digital Realty","NTT"}:
                        score += 5
                        break
                    elif co and co in co_frequency_map:
                        score += min(co_frequency_map[co] * 0.5, 3)

                # 5. Topic weight
                topic_w = {
                    "Hyperscale": 8, "AI / GPU": 8, "Investment": 7,
                    "Power": 6, "Colocation": 5, "Construction": 5,
                    "Permits": 4, "Sustainability": 3, "General": 1,
                }
                score += topic_w.get(row.get("Topic", "General"), 1)

                # 6. High-value keyword bonuses
                bonus_kws = [
                    ("billion", 4), ("gigawatt", 5), ("gw", 3), ("nuclear", 4),
                    ("hyperscale", 3), ("ai campus", 5), ("gpu", 3),
                    ("acquisition", 4), ("merger", 4), ("ipo", 5),
                    ("stargate", 5), ("1gw", 5), ("2gw", 6),
                ]
                for kw, pts in bonus_kws:
                    if kw in hl:
                        score += pts

                return round(score, 1)

            # Build co frequency map
            _all_co_score = []
            for v in df["Companies"]:
                if v:
                    _all_co_score.extend([c.strip() for c in str(v).split(",") if c.strip()])
            co_freq_map = dict(_Ctr(_all_co_score))

            df_scored = df.copy()
            df_scored["AI Score"] = df_scored.apply(lambda r: _ai_score(r, co_freq_map), axis=1)
            df_scored = df_scored.sort_values("AI Score", ascending=False).reset_index(drop=True)

            # Signal tier summary
            high_signal = int((df_scored["AI Score"] >= 30).sum())
            medium_signal = int(((df_scored["AI Score"] >= 15) & (df_scored["AI Score"] < 30)).sum())
            low_signal = int((df_scored["AI Score"] < 15).sum())

            kpi_score_row = (
                '<div style="display:flex;gap:.8rem;margin-bottom:1rem;flex-wrap:wrap;">'
                + kpi("🔴 High Signal (30+)", high_signal, "amber", "major deals, large capacity")
                + kpi("🟡 Medium Signal (15-29)", medium_signal, "blue")
                + kpi("⚪ Low Signal (<15)", low_signal, "cyan")
                + kpi("Max Score", df_scored["AI Score"].max() if not df_scored.empty else 0, "purple")
                + '</div>'
            )
            st.markdown(kpi_score_row, unsafe_allow_html=True)

            # Score distribution chart
            score_bins = pd.cut(df_scored["AI Score"], bins=[0, 10, 20, 30, 40, 50, 200],
                                 labels=["0-10", "10-20", "20-30", "30-40", "40-50", "50+"])
            score_dist = score_bins.value_counts().sort_index().reset_index()
            score_dist.columns = ["Score Range", "Articles"]
            score_colors = ["#2e4470", "#0047e1", "#00b4ff", "#00e5c8", "#ffaa00", "#ff6400"]
            fig_score = go.Figure(go.Bar(
                x=score_dist["Score Range"].astype(str),
                y=score_dist["Articles"],
                marker_color=score_colors[:len(score_dist)],
                marker_line_width=0,
                text=score_dist["Articles"], textposition="outside",
                textfont=dict(color=_TITLE, size=10, family="DM Mono, monospace"),
                hovertemplate="<b>Score %{x}</b><br>📰 %{y} articles<extra></extra>",
            ))
            _dark(fig_score, 270)
            fig_score.update_layout(
                title=dict(text="AI Score Distribution", font=dict(color=_TITLE, size=13, family="Syne, sans-serif"), x=0.01),
                bargap=0.3,
            )
            st.plotly_chart(fig_score, use_container_width=True, config={"displayModeBar": False})

            # High-signal articles with orange top-strip
            st.markdown('<div class="sec-head">🔴 High Signal Articles (Score ≥ 30)</div>', unsafe_allow_html=True)
            top_scored = df_scored[df_scored["AI Score"] >= 30].head(25)
            if top_scored.empty:
                st.info("No articles scored 30+ in current view. Try a broader scan.")
            else:
                for _, row in top_scored.iterrows():
                    st.markdown(article_card(
                        row["Headline"], row["Date"], row["URL"],
                        row["Source"], row["Country"], row["Topic"],
                        row.get("Capacity", ""), row.get("Deal Size", ""),
                        row.get("Sentiment", "News"),
                        ai_score=row["AI Score"],
                    ), unsafe_allow_html=True)

            # Medium signal
            st.markdown('<div class="sec-head">🟡 Medium Signal Articles (15–29)</div>', unsafe_allow_html=True)
            med_scored = df_scored[(df_scored["AI Score"] >= 15) & (df_scored["AI Score"] < 30)].head(20)
            for _, row in med_scored.iterrows():
                st.markdown(article_card(
                    row["Headline"], row["Date"], row["URL"],
                    row["Source"], row["Country"], row["Topic"],
                    row.get("Capacity", ""), row.get("Deal Size", ""),
                    row.get("Sentiment", "News"),
                    ai_score=row["AI Score"],
                ), unsafe_allow_html=True)

            # Full scored table
            st.markdown('<div class="sec-head">Full Scored Article Table</div>', unsafe_allow_html=True)
            scored_display = df_scored[["AI Score","Headline","Date","Topic","Sentiment","Capacity","Deal Size","Country","ISO / RTO","URL"]].copy() if "ISO / RTO" in df_scored.columns else df_scored[["AI Score","Headline","Date","Topic","Sentiment","Capacity","Deal Size","Country","URL"]].copy()
            st.markdown(dark_table(scored_display), unsafe_allow_html=True)


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
                        "Topic","Sentiment","Capacity","Deal Size","Companies","ISO / RTO","URL"]
        st.markdown(
            dark_table(df[[c for c in display_cols if c in df.columns]]),
            unsafe_allow_html=True,
        )

    # ── Platform footer ───────────────────────────────────────────────────────
    st.markdown(
        '<div style="margin-top:3rem;padding:1.2rem 0 .6rem;border-top:1px solid #101b2e;'
        'text-align:center;">'
        '<div style="font-family:\'DM Mono\',monospace;font-size:.6rem;letter-spacing:.14em;'
        'color:#1a2e50;text-transform:uppercase;">'
        'Global Data Center Intelligence &nbsp;\u00b7&nbsp; Wood Mac'
        '</div></div>',
        unsafe_allow_html=True,
    )


if __name__ == "__main__":
    main()
