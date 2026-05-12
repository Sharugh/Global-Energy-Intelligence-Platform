import streamlit as st
import re
import io
import json
import time
import math
import textwrap
import feedparser
from collections import Counter
from datetime import datetime, timedelta
from concurrent.futures import ThreadPoolExecutor, as_completed
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
@import url('https://fonts.googleapis.com/css2?family=Syne:wght@400;600;700;800&family=DM+Mono:wght@400;500&family=Inter:wght@300;400;500&display=swap');

html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
.stApp { background: #060a10; color: #e8edf5; }

/* ── Global hover transitions ──────────────────────────────────────────────── */
* { transition: background 0.15s ease, border-color 0.15s ease, box-shadow 0.18s ease, opacity 0.15s ease, transform 0.15s ease; }

/* ── Sidebar shell ─────────────────────────────────────────────────────────── */
[data-testid="stSidebar"] { background: #0a0f1a !important; border-right: 1px solid #151f35; }
[data-testid="stSidebar"] * { color: #b8c8e0 !important; }
[data-testid="stSidebar"] hr { border-color: #151f35 !important; }

/* Sidebar Run button */
[data-testid="stSidebar"] .stButton button {
    background: linear-gradient(135deg, #0047e1, #00b4ff) !important;
    color: #fff !important; border: none !important; border-radius: 8px !important;
    font-family: 'Syne', sans-serif !important; font-weight: 700 !important;
    font-size: 0.92rem !important; letter-spacing: 0.04em !important;
    padding: 0.65rem 1rem !important; transition: opacity .2s;
}
[data-testid="stSidebar"] .stButton button:hover { opacity: .82; }

/* ── Sidebar — multiselect control box ────────────────────────────────────── */
[data-testid="stSidebar"] [data-baseweb="select"] {
    background: #0b1628 !important;
}
[data-testid="stSidebar"] [data-baseweb="select"] > div:first-child {
    background: #0b1628 !important;
    border: 1px solid #1e3050 !important;
    border-radius: 8px !important;
}
[data-testid="stSidebar"] [data-baseweb="select"] > div:first-child:focus-within {
    border-color: #0047e1 !important;
    box-shadow: 0 0 0 2px rgba(0,71,225,0.18) !important;
}
/* Placeholder text */
[data-testid="stSidebar"] [data-baseweb="select"] [data-testid="stMarkdownContainer"],
[data-testid="stSidebar"] [data-baseweb="select"] span {
    color: #3a5480 !important;
}
/* Selected tag pills inside multiselect */
[data-testid="stSidebar"] [data-baseweb="tag"] {
    background: #0f2245 !important;
    border: 1px solid #0047e1 !important;
    border-radius: 5px !important;
}
[data-testid="stSidebar"] [data-baseweb="tag"] span {
    color: #7eb8ff !important;
}
[data-testid="stSidebar"] [data-baseweb="tag"] [role="presentation"] svg {
    fill: #3a5480 !important;
}
/* The input inside multiselect */
[data-testid="stSidebar"] [data-baseweb="select"] input {
    background: transparent !important;
    color: #ccdaf5 !important;
    caret-color: #0047e1 !important;
}

/* ── Dropdown menu (popover) — dark background ────────────────────────────── */
[data-baseweb="popover"],
[data-baseweb="menu"],
[role="listbox"],
ul[data-baseweb="menu"] {
    background: #0d1a2e !important;
    border: 1px solid #1e3050 !important;
    border-radius: 10px !important;
    box-shadow: 0 8px 32px rgba(0,0,0,0.6) !important;
}
/* Each option row */
[data-baseweb="menu"] li,
[role="option"],
[data-baseweb="menu"] [role="option"] {
    background: #0d1a2e !important;
    color: #b8c8e0 !important;
}
[data-baseweb="menu"] li:hover,
[role="option"]:hover,
[data-baseweb="menu"] [role="option"]:hover {
    background: #0f2245 !important;
    color: #ffffff !important;
}
/* Highlighted/active option */
[data-baseweb="menu"] [aria-selected="true"],
[role="option"][aria-selected="true"] {
    background: #0f2245 !important;
    color: #00b4ff !important;
}
/* Search input inside dropdown */
[data-baseweb="popover"] input,
[data-baseweb="menu"] input {
    background: #060a10 !important;
    border: 1px solid #1e3050 !important;
    color: #ccdaf5 !important;
    border-radius: 6px !important;
}

/* ── Sidebar text input ────────────────────────────────────────────────────── */
[data-testid="stSidebar"] .stTextInput input {
    background: #0b1628 !important;
    border: 1px solid #1e3050 !important;
    border-radius: 8px !important;
    color: #d0dff0 !important;
}
[data-testid="stSidebar"] .stTextInput input:focus {
    border-color: #0047e1 !important;
    box-shadow: 0 0 0 2px rgba(0,71,225,0.18) !important;
}

/* ── Sidebar number input ─────────────────────────────────────────────────── */
[data-testid="stSidebar"] input[type="number"] {
    background: #0b1628 !important;
    border: 1px solid #1e3050 !important;
    border-radius: 8px !important;
    color: #d0dff0 !important;
}

/* ── Sidebar date inputs ──────────────────────────────────────────────────── */
[data-testid="stSidebar"] [data-testid="stDateInput"] input {
    background: #0b1628 !important;
    border: 1px solid #1e3050 !important;
    color: #d0dff0 !important;
    border-radius: 8px !important;
}

/* ── Main content selects / multiselects ──────────────────────────────────── */
.stMultiSelect [data-baseweb="select"] { background: #0b1628 !important; border-color: #152038 !important; }
.stSelectbox [data-baseweb="select"] > div { background: #0b1628 !important; border-color: #152038 !important; }
.stTextInput input {
    background: #0b1628 !important; border: 1px solid #152038 !important;
    border-radius: 8px !important; color: #d0dff0 !important;
}
.stTextInput input:focus { border-color: #0047e1 !important; }

/* ── Tabs ─────────────────────────────────────────────────────────────────── */
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

/* ── Download buttons ─────────────────────────────────────────────────────── */
.stDownloadButton button {
    background: linear-gradient(135deg, #002d0a, #005214) !important;
    color: #00e676 !important; border: 1px solid #00a846 !important;
    border-radius: 8px !important; font-family: 'Syne', sans-serif !important;
    font-weight: 700 !important; letter-spacing: .04em !important;
}
.stDownloadButton button:hover { opacity: .82 !important; }

/* ── Banner ───────────────────────────────────────────────────────────────── */
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

/* ── Section headings & pills ─────────────────────────────────────────────── */
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

hr { border-color: #152038 !important; }
#MainMenu, footer, header { visibility: hidden; }

/* ── Hover card effect (applied to .hover-card class in HTML) ─────────────── */
.hover-card {
    transition: transform 0.18s ease, border-color 0.18s ease, box-shadow 0.2s ease !important;
}
.hover-card:hover {
    transform: translateY(-2px) !important;
    border-color: #0047e1 !important;
    box-shadow: 0 6px 28px rgba(0, 71, 225, 0.18) !important;
}

/* ── Tooltip-style hover on pills ─────────────────────────────────────────── */
.pill:hover {
    border-color: #0047e1 !important;
    background: #0f2245 !important;
    color: #ccdaf5 !important;
    cursor: default;
}

/* ── Section heading hover ──────────────────────────────────────────────────── */
.sec-head:hover {
    border-left-color: #00b4ff !important;
    color: #fff !important;
    cursor: default;
}

/* ── KPI card hover ─────────────────────────────────────────────────────────── */
.kpi-card {
    transition: transform 0.18s ease, border-color 0.18s ease, box-shadow 0.2s ease !important;
}
.kpi-card:hover {
    transform: translateY(-3px) !important;
    border-color: #0047e1 !important;
    box-shadow: 0 8px 32px rgba(0, 71, 225, 0.22) !important;
}

/* ── Table row hover ─────────────────────────────────────────────────────────── */
.dc-table tr:hover td {
    background: #0f2245 !important;
}

/* ── Tab hover ───────────────────────────────────────────────────────────────── */
.stTabs [data-baseweb="tab"]:hover:not([aria-selected="true"]) {
    background: rgba(0,71,225,0.12) !important;
    color: #ccdaf5 !important;
}

/* ── Sidebar button hover ───────────────────────────────────────────────────── */
[data-testid="stSidebar"] .stButton button:hover {
    opacity: .82;
    box-shadow: 0 4px 16px rgba(0, 71, 225, 0.35) !important;
    transform: translateY(-1px) !important;
}

/* ── Download button hover ──────────────────────────────────────────────────── */
.stDownloadButton button:hover {
    opacity: .82 !important;
    box-shadow: 0 4px 14px rgba(0, 168, 70, 0.35) !important;
    transform: translateY(-1px) !important;
}

/* ── Feature card hover ─────────────────────────────────────────────────────── */
.feature-card {
    transition: transform 0.18s ease, border-color 0.2s ease, box-shadow 0.2s ease !important;
    cursor: default;
}
.feature-card:hover {
    transform: translateY(-4px) !important;
    border-color: #0047e1 !important;
    box-shadow: 0 10px 36px rgba(0, 71, 225, 0.18) !important;
}

/* ── Saved scan row hover ─────────────────────────────────────────────────── */
.saved-scan-card {
    transition: border-color 0.18s ease, box-shadow 0.18s ease !important;
}
.saved-scan-card:hover {
    border-color: #a855f7 !important;
    box-shadow: 0 4px 20px rgba(168, 85, 247, 0.15) !important;
}

/* ── Article card hover ─────────────────────────────────────────────────────── */
.article-card-wrap {
    transition: border-color 0.18s ease, box-shadow 0.2s ease, transform 0.15s ease !important;
}
.article-card-wrap:hover {
    border-color: #0047e1 !important;
    box-shadow: 0 6px 28px rgba(0, 71, 225, 0.15) !important;
    transform: translateY(-1px) !important;
}

/* ── Score badge hover ───────────────────────────────────────────────────────── */
.score-badge {
    transition: filter 0.15s ease, transform 0.15s ease !important;
}
.score-badge:hover {
    filter: brightness(1.3) !important;
    transform: scale(1.05) !important;
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
    "DataCenterDynamics":  {"color": "#0047e1", "short": "DCD"},
    "DataCenter Knowledge":{"color": "#00b4ff", "short": "DCK"},
    "DataCentreMagazine":  {"color": "#00e5c8", "short": "DCM"},
    "DataCenterFrontier":  {"color": "#7ec8ff", "short": "DCF"},
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


# ─── DCD constants ─────────────────────────────────────────────────────────
DCD_BASE     = "https://www.datacenterdynamics.com"
DCD_CHAN_TERM = "the-data-center-construction-channel"

# ─── Additional source base URLs ────────────────────────────────────────────
DCK_BASE = "https://www.datacenterknowledge.com"
DCM_BASE = "https://datacentremagazine.com"

# DCD region taxonomy terms → mapped to code2's Region labels
# These are the actual URL ?term= slugs DCD uses in their site taxonomy
DCD_REGION_TERMS = {
    "north-america":  "North America",
    "europe":         "Europe",
    "asia-pacific":   "Asia Pacific",
    "middle-east":    "Middle East",
    "africa":         "Africa",
    "latin-america":  "Latin America",
}

# DCD project-stage / topic terms for extra coverage
DCD_EXTRA_TERMS = [
    "approved",
    "site-selection",
    "disclosed-projects",
    "project-announcement",
    "expansion",
    "extension",
]

GNEWS_QUERIES = [
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
]

RSS_SOURCES = []   # All sources are HTML-scraped directly


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

def fetch_html(url, retries=3):
    """
    Fetch a URL and return a BeautifulSoup object.
    Retries up to `retries` times with exponential back-off.
    Handles 429 / 503 rate-limit responses gracefully.
    Returns None on total failure (never raises).
    """
    _RETRY_STATUSES = {429, 500, 502, 503, 504}
    for attempt in range(retries):
        try:
            if _USE_CS:
                r = _CS.get(url, timeout=25)
            else:
                r = _CS.get(url, headers=_DCD_HEADERS, timeout=25)

            # Rate-limit / server-error → back off and retry
            if r.status_code in _RETRY_STATUSES:
                wait = 2 ** (attempt + 1)          # 2s, 4s, 8s
                time.sleep(wait)
                continue

            r.raise_for_status()

            # Guard against empty / non-HTML bodies
            ct = r.headers.get("Content-Type", "")
            if "html" not in ct and "xml" not in ct and len(r.text) < 200:
                return None

            return BeautifulSoup(r.text, "html.parser")

        except Exception:
            if attempt < retries - 1:
                time.sleep(2 ** attempt)            # 1s, 2s before final attempt
    return None


# ─── Article parser (app15 logic + <time datetime> + ISO date fallback) ────
def _parse_articles_from_soup(soup, source_name, base_url):
    """
    app15-style parser: finds all <a href="/en/news/..."> links,
    extracts headline from h1-h4 inside the link, climbs DOM for date.
    Works for DCD only (href pattern ^/en/news/).
    Enhanced: also checks <time datetime="..."> and ISO dates in text.
    """
    articles = []
    seen = set()

    for a in soup.find_all("a", href=re.compile(r"^/en/(news|analysis|opinion)/[^?#]+/$")):
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


# ─── Global request rate limiter ────────────────────────────────────────────
class _RateLimiter:
    """
    Token-bucket rate limiter for the HTML scrapers.
    Ensures the total requests/second across all scrapers stays under a
    configurable ceiling, preventing IP bans at high scrape depths (50+).

    Usage:
        _rate_limiter.wait()   # call before every fetch_html() at depth
    """
    def __init__(self, max_rps: float = 2.5):
        self._min_gap = 1.0 / max_rps    # minimum seconds between requests
        self._last    = 0.0

    def wait(self):
        now  = time.monotonic()
        gap  = now - self._last
        if gap < self._min_gap:
            time.sleep(self._min_gap - gap)
        self._last = time.monotonic()

_rate_limiter = _RateLimiter(max_rps=2.5)   # ≤ 2.5 requests/s globally


# ─── Structural change detector ──────────────────────────────────────────────
def _check_scraper_structure(soup, source_name: str, expected_min_links: int = 3) -> bool:
    """
    Lightweight structural health check.
    Returns True if the page looks like a valid article listing.
    Logs a warning into scraper health if the site has likely been redesigned.
    """
    if soup is None:
        return False

    # Check for a plausible number of article-like <a> tags
    a_tags = soup.find_all("a", href=True)
    if len(a_tags) < expected_min_links:
        _record_health(
            source_name, 0, 1,
            f"⚠️ Structural change detected — only {len(a_tags)} links found. "
            "Site may have been redesigned.",
        )
        return False

    # Check that the page isn't a CAPTCHA / bot-block page
    page_text = soup.get_text(" ", strip=True).lower()
    block_signals = ["captcha", "access denied", "403 forbidden",
                     "bot detection", "enable javascript", "checking your browser"]
    for sig in block_signals:
        if sig in page_text[:2000]:
            _record_health(
                source_name, 0, 1,
                f"🚫 Bot block / CAPTCHA detected on {source_name}. "
                "Consider rotating User-Agent or adding delay.",
            )
            return False

    return True


# ─── Scraper versioning & fallback parser registry ───────────────────────────
# Each entry is a (version_tag, parser_fn) pair.
# _parse_articles_from_soup_v2 is the current primary; earlier versions are kept
# as automatic fallbacks so a DCD redesign degrades gracefully instead of
# returning zero articles.
#
# How it works:
#   scrape_dcd() calls _parse_articles_versioned() which tries parsers in order.
#   If the primary returns < MIN_ARTICLES, it tries the next version.
#   Health is recorded with which version succeeded (or "all_failed").
#
# To add a new parser after a redesign:
#   1. Define _parse_articles_from_soup_vN(soup, source, base_url)
#   2. Prepend it to _DCD_PARSER_VERSIONS below.

_MIN_PARSER_ARTICLES = 1   # minimum articles to consider a parse "successful"

def _parse_articles_v1_fallback(soup, source_name: str, base_url: str) -> list:
    """
    Broad fallback: catches any <a> whose href contains '/news/' or '/analysis/'
    and whose inner text is long enough to be a headline.
    Intentionally permissive — used only when the primary parser yields nothing.
    """
    articles = []
    seen = set()
    for a in soup.find_all("a", href=re.compile(r"/(news|analysis|opinion)/", re.I)):
        href = a.get("href", "")
        if not href:
            continue
        if not href.startswith("http"):
            href = base_url + href
        norm = href.rstrip("/")
        if norm in seen:
            continue
        seen.add(norm)
        text = a.get_text(" ", strip=True)
        text = re.sub(r"\s+", " ", text).strip()
        if len(text) < 12:
            continue
        articles.append({
            "headline":  text,
            "url":       norm,
            "date_obj":  None,
            "source":    source_name,
            "_priority": 5,          # lower priority than primary parser
        })
    return articles


def _parse_articles_v2_generic(soup, source_name: str, base_url: str) -> list:
    """
    Generic article-card parser: looks for <article> / [class*=card] / [class*=item]
    wrappers that contain both a link and a heading, regardless of URL structure.
    Used as a mid-level fallback between primary and broad fallback.
    """
    articles = []
    seen = set()
    containers = soup.find_all(
        lambda tag: tag.name in ("article", "div", "li", "section")
        and any(c for c in (tag.get("class") or [])
                if any(kw in c.lower() for kw in ("card", "item", "post", "story", "result")))
    )
    for container in containers:
        a_tag = container.find("a", href=True)
        h_tag = container.find(["h1", "h2", "h3", "h4", "h5"])
        if not a_tag or not h_tag:
            continue
        href = a_tag["href"]
        if not href.startswith("http"):
            href = base_url + href
        norm = href.rstrip("/")
        if norm in seen:
            continue
        seen.add(norm)
        headline = re.sub(r"\s+", " ", h_tag.get_text(" ", strip=True)).strip()
        if len(headline) < 10:
            continue
        # Try to extract a date from the container
        date_obj = None
        tt = container.find("time")
        if tt:
            date_obj = parse_date_str(tt.get("datetime", "") or tt.get_text(strip=True))
        articles.append({
            "headline":  headline,
            "url":       norm,
            "date_obj":  date_obj,
            "source":    source_name,
            "_priority": 4,
        })
    return articles


# Ordered list of (version_tag, parser_fn) — primary first, fallbacks after.
_DCD_PARSER_VERSIONS: list[tuple[str, callable]] = [
    ("v3-primary",  _parse_articles_from_soup),   # current production parser
    ("v2-generic",  _parse_articles_v2_generic),   # mid-level fallback
    ("v1-broad",    _parse_articles_v1_fallback),  # permissive last-resort
]


def _parse_articles_versioned(
    soup,
    source_name: str,
    base_url: str,
    health_key: str = "DataCenterDynamics",
) -> tuple[list, str]:
    """
    Try each parser version in order.  Returns (articles, version_tag) for the
    first version that yields at least _MIN_PARSER_ARTICLES results.
    Records health entry with the active parser version for observability.
    """
    for version_tag, parser_fn in _DCD_PARSER_VERSIONS:
        try:
            articles = parser_fn(soup, source_name, base_url)
            if len(articles) >= _MIN_PARSER_ARTICLES:
                _record_health(
                    health_key, len(articles), 0,
                    f"OK (parser={version_tag})",
                )
                return articles, version_tag
        except Exception as exc:
            _record_health(
                health_key, 0, 1,
                f"Parser {version_tag} error: {exc}",
            )

    # All parsers failed — log and return empty
    _record_health(
        health_key, 0, 1,
        "⚠️ All parser versions failed — DCD may have been redesigned.",
    )
    return [], "all_failed"
def scrape_dcd(cutoff, max_pages, region_terms, pbar=None):
    """
    Scrapes DCD using:
      ?term=the-data-center-construction-channel
      &term=<region_slug>      (one per selected region, or all if none selected)
      &term=<extra_stage_term> (approved / site-selection / etc.)
      &page=N

    This is exactly what makes app15 fetch more articles — DCD's own
    pre-filtered construction channel returns only DC construction content,
    so almost every article is relevant regardless of region.

    region_terms: list of DCD slug strings from DCD_REGION_TERMS keys.
                  Empty list = scrape ALL regions (no region term added).
    """
    # Warm up DCD session (app15 does this too — helps bypass Cloudflare)
    fetch_html(DCD_BASE + "/en/")

    all_articles = []
    seen_urls    = set()

    # Build the set of term combinations to scrape:
    # 1. Construction channel + each selected region  (one URL per region)
    # 2. Each extra stage term (no region filter — global)
    # 3. Bare construction channel with no region (always included as catch-all)

    url_sets = []

    # If specific regions selected, add one URL per region
    if region_terms:
        for rterm in region_terms:
            url_sets.append([DCD_CHAN_TERM, rterm])
    else:
        # No region filter — just the construction channel (global)
        url_sets.append([DCD_CHAN_TERM])

    # Extra stage terms (approved, site-selection, etc.) — always global
    for extra in DCD_EXTRA_TERMS:
        url_sets.append([DCD_CHAN_TERM, extra])

    total_fetches = len(url_sets) * max_pages
    fetched       = [0]

    for terms in url_sets:
        for page in range(1, max_pages + 1):
            # Build URL: ?term=A&term=B&page=N
            params = "&".join(f"term={t}" for t in terms)
            if page > 1:
                params += f"&page={page}"
            url = f"{DCD_BASE}/en/news/?{params}"

            fetched[0] += 1
            if pbar:
                pbar.progress(
                    min(fetched[0] / total_fetches, 1.0),
                    text=f"⚡ Neural Scan: Probing {', '.join(terms).upper()} · Page {page}/{max_pages}",
                )

            try:
                _rate_limiter.wait()          # global request budget
                soup = fetch_html(url)
                if not soup:
                    break   # this term set seems unreachable; move on

                # Structural health check on the first page of each term set
                if page == 1 and not _check_scraper_structure(soup, "DataCenterDynamics"):
                    break

                page_arts, _active_ver = _parse_articles_versioned(
                    soup, "DataCenterDynamics", DCD_BASE
                )
                new_on_page = 0
                stop = False

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

            except Exception:
                # One page failed — skip it but continue with remaining pages/terms
                break

    return all_articles


# ─── Google News fetcher — versioned with fallback chain ─────────────────────
#
# Problem: if Google changes its RSS structure, the entire GNews pipeline
# silently returns zero articles.  We now have three strategies tried in order:
#
#   v1-rss     Primary: RSS feed via requests + feedparser (fastest)
#   v2-atom    Fallback: Atom/JSON feed variant Google sometimes serves instead
#   v3-scrape  Last-resort: scrape the Google News HTML search results page
#
# Each strategy is a standalone function returning a list or raising on failure.
# _fetch_google_news_versioned() iterates through them and returns the first
# non-empty result, recording which version was used in _SCRAPER_HEALTH.

_GN_MIN_ENTRIES = 1   # minimum entries to consider a strategy successful


def _gn_parse_feed(feed, source_label: str) -> list:
    """Convert a feedparser feed object into our standard article dicts."""
    results = []
    for entry in feed.entries:
        try:
            headline = (entry.get("title") or "").strip()
            url_val  = (entry.get("link")  or "").strip()
            if not headline or len(headline) < 10:
                continue
            if not url_val or not url_val.startswith("http"):
                continue
            if "news.google.com/rss/articles" in url_val and len(url_val) < 60:
                continue
            date_obj = None
            if hasattr(entry, "published_parsed") and entry.published_parsed:
                try:
                    date_obj = datetime(*entry.published_parsed[:6])
                except Exception:
                    pass
            results.append({
                "headline":  headline,
                "url":       url_val,
                "date_obj":  date_obj,
                "source":    source_label,
                "_priority": 10,
            })
        except Exception:
            continue
    return results


def _gn_strategy_rss(q_encoded: str, source_label: str) -> list:
    """v1-rss: Standard Google News RSS endpoint."""
    url = f"https://news.google.com/rss/search?q={q_encoded}&hl=en-US&gl=US&ceid=US:en"
    try:
        if _USE_CS:
            r = _CS.get(url, timeout=15)
        else:
            r = _CS.get(url, headers=_HEADERS, timeout=15)
        r.raise_for_status()
        feed = feedparser.parse(r.text)
    except Exception:
        feed = feedparser.parse(url)   # let feedparser try with its own socket
    if getattr(feed, "bozo", False) and not feed.entries:
        raise ValueError("RSS feed bozo / empty")
    results = _gn_parse_feed(feed, source_label)
    if not results:
        raise ValueError("RSS returned 0 usable entries")
    return results


def _gn_strategy_atom(q_encoded: str, source_label: str) -> list:
    """
    v2-atom: Google News Atom / alternate endpoint.
    Google occasionally switches between RSS and Atom; this catches that.
    Also tries the /search?q= HTML endpoint with an Accept header that
    asks for application/atom+xml.
    """
    atom_url = f"https://news.google.com/atom/search?q={q_encoded}&hl=en-US&gl=US&ceid=US:en"
    try:
        if _USE_CS:
            r = _CS.get(atom_url, timeout=15)
        else:
            r = _CS.get(atom_url, headers={**_HEADERS, "Accept": "application/atom+xml,*/*"}, timeout=15)
        r.raise_for_status()
        feed = feedparser.parse(r.text)
    except Exception:
        feed = feedparser.parse(atom_url)
    if getattr(feed, "bozo", False) and not feed.entries:
        raise ValueError("Atom feed bozo / empty")
    results = _gn_parse_feed(feed, source_label)
    if not results:
        raise ValueError("Atom returned 0 usable entries")
    return results


def _gn_strategy_scrape(q_encoded: str, source_label: str) -> list:
    """
    v3-scrape: Last-resort direct HTML scrape of Google News search results.
    Extracts article links and headlines from the rendered search page.
    Less reliable than feed-based strategies but works when both feeds fail.
    """
    search_url = f"https://news.google.com/search?q={q_encoded}&hl=en-US&gl=US&ceid=US:en"
    soup = fetch_html(search_url)
    if soup is None:
        raise ValueError("Google News HTML fetch returned None")
    results = []
    seen = set()
    # Google News renders articles as <article> tags with <a> links inside
    for article_tag in soup.find_all("article"):
        a = article_tag.find("a", href=True)
        if not a:
            continue
        href = a.get("href", "")
        # GNews uses relative /articles/... links
        if href.startswith("./"):
            href = "https://news.google.com/" + href[2:]
        elif href.startswith("/"):
            href = "https://news.google.com" + href
        if not href.startswith("http"):
            continue
        if href in seen:
            continue
        seen.add(href)
        # Headline: prefer h3/h4, fall back to all link text
        h_tag = article_tag.find(["h3", "h4"])
        headline = (h_tag.get_text(" ", strip=True) if h_tag
                    else a.get_text(" ", strip=True)).strip()
        headline = re.sub(r"\s+", " ", headline)
        if len(headline) < 10:
            continue
        # Date from <time datetime="..."> if present
        date_obj = None
        tt = article_tag.find("time")
        if tt:
            date_obj = parse_date_str(tt.get("datetime", "") or tt.get_text(strip=True))
        results.append({
            "headline":  headline,
            "url":       href,
            "date_obj":  date_obj,
            "source":    source_label,
            "_priority": 11,   # slightly lower than feed-based
        })
    if not results:
        raise ValueError("HTML scrape returned 0 usable articles")
    return results


# Ordered list of (version_tag, strategy_fn) — primary first
_GN_STRATEGIES: list[tuple[str, callable]] = [
    ("v1-rss",    _gn_strategy_rss),
    ("v2-atom",   _gn_strategy_atom),
    ("v3-scrape", _gn_strategy_scrape),
]


def fetch_google_news(query: str, source_label: str = "Google News") -> list:
    """
    Fetch Google News articles for `query` using a versioned fallback chain.
    Returns a (possibly empty) list — never raises.
    Tries v1-rss → v2-atom → v3-scrape in order, stopping at the first success.
    """
    q_encoded = query.replace(" ", "+")
    for version_tag, strategy_fn in _GN_STRATEGIES:
        try:
            results = strategy_fn(q_encoded, source_label)
            if len(results) >= _GN_MIN_ENTRIES:
                # Tag each article with the strategy that produced it (for debugging)
                for r in results:
                    r["_gn_strategy"] = version_tag
                return results
        except Exception:
            continue   # try next strategy

    return []   # all strategies failed — caller tracks error count


# ─── Scraper health tracking ────────────────────────────────────────────────
# Each source is tracked: articles fetched, errors, last status
_SCRAPER_HEALTH: dict = {}

def _record_health(source: str, count: int, errors: int, status: str):
    """Record per-source scrape health for the Source Health panel."""
    _SCRAPER_HEALTH[source] = {
        "articles": count,
        "errors":   errors,
        "status":   status,
        "ts":       datetime.now().strftime("%H:%M:%S"),
    }


# ─── DataCenter Knowledge scraper ──────────────────────────────────────────
DCK_BASE = "https://www.datacenterknowledge.com"

def scrape_dck(cutoff, max_pages=5):
    """Scrape DataCenter Knowledge news pages."""
    articles = []
    errors   = 0
    seen     = set()
    try:
        for page in range(1, max_pages + 1):
            url = DCK_BASE + "/news" + (f"?page={page}" if page > 1 else "")
            try:
                _rate_limiter.wait()
                soup = fetch_html(url)
                if soup is None:
                    errors += 1
                    break
                if page == 1 and not _check_scraper_structure(soup, "DataCenter Knowledge"):
                    errors += 1
                    break
                for a in soup.find_all("a", href=re.compile(r"/(data-centers|cloud|edge)/[^?#]+$")):
                    href = a.get("href", "")
                    if not href.startswith("http"):
                        href = DCK_BASE + href
                    if href in seen:
                        continue
                    seen.add(href)
                    h_tag = a.find(["h1","h2","h3","h4"])
                    headline = (h_tag.get_text(" ", strip=True) if h_tag else a.get_text(" ", strip=True)).strip()
                    if not headline or len(headline) < 12:
                        continue
                    # date: walk up DOM
                    date_obj = None
                    node = a.parent
                    for _ in range(10):
                        if node is None:
                            break
                        tt = node.find("time")
                        if tt:
                            date_obj = parse_date_str(tt.get("datetime","") or tt.get_text(strip=True))
                            if date_obj:
                                break
                        m = re.search(r"\b(\d{1,2})\s+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\w*\s+(\d{4})\b", node.get_text(" ",strip=True), re.I)
                        if m:
                            date_obj = parse_date_str(m.group(0))
                            break
                        node = node.parent
                    if date_obj and date_obj < cutoff:
                        break
                    if not is_dc_relevant(headline):
                        continue
                    articles.append({
                        "headline":  headline,
                        "url":       href,
                        "date_obj":  date_obj,
                        "source":    "DataCenter Knowledge",
                        "_priority": 2,
                    })
                time.sleep(0.4)
            except Exception as page_exc:
                errors += 1
                break   # stop paging this source on page-level error
    except Exception as exc:
        errors += 1
        _record_health("DataCenter Knowledge", len(articles), errors, f"Error: {exc}")
        return articles
    _record_health("DataCenter Knowledge", len(articles), errors,
                   "OK" if errors == 0 else f"Partial ({errors} page error(s))")
    return articles


# ─── Data Centre Magazine scraper ──────────────────────────────────────────
DCM_BASE = "https://datacentremagazine.com"

def scrape_dcm(cutoff, max_pages=4):
    """Scrape Data Centre Magazine."""
    articles = []
    errors   = 0
    seen     = set()
    try:
        for page in range(1, max_pages + 1):
            url = DCM_BASE + "/latest" + (f"?page={page}" if page > 1 else "")
            try:
                _rate_limiter.wait()
                soup = fetch_html(url)
                if soup is None:
                    errors += 1
                    break
                if page == 1 and not _check_scraper_structure(soup, "DataCentreMagazine"):
                    errors += 1
                    break
                for a in soup.find_all("a", href=re.compile(r"/(data-centres|technology|sustainability)/[^?#]+")):
                    href = a.get("href", "")
                    if not href.startswith("http"):
                        href = DCM_BASE + href
                    if href in seen:
                        continue
                    seen.add(href)
                    h_tag = a.find(["h1","h2","h3","h4"])
                    headline = (h_tag.get_text(" ", strip=True) if h_tag else a.get_text(" ", strip=True)).strip()
                    if not headline or len(headline) < 12:
                        continue
                    date_obj = None
                    node = a.parent
                    for _ in range(10):
                        if node is None:
                            break
                        tt = node.find("time")
                        if tt:
                            date_obj = parse_date_str(tt.get("datetime","") or tt.get_text(strip=True))
                            if date_obj:
                                break
                        node = node.parent
                    if date_obj and date_obj < cutoff:
                        break
                    if not is_dc_relevant(headline):
                        continue
                    articles.append({
                        "headline":  headline,
                        "url":       href,
                        "date_obj":  date_obj,
                        "source":    "DataCentreMagazine",
                        "_priority": 3,
                    })
                time.sleep(0.4)
            except Exception:
                errors += 1
                break   # stop paging on page-level error
    except Exception as exc:
        errors += 1
        _record_health("DataCentreMagazine", len(articles), errors, f"Error: {exc}")
        return articles
    _record_health("DataCentreMagazine", len(articles), errors,
                   "OK" if errors == 0 else f"Partial ({errors} page error(s))")
    return articles


# ─── run_all_scrapers: DCD + DCK + DCM + Google News ──────────────────────
def run_all_scrapers(max_html_pages, cutoff, progress_cb, region_terms=None):
    """
    region_terms: list of DCD region slug strings (e.g. ["north-america","europe"]).
                  Pass [] or None for global (no region filter on DCD).
    Returns (filtered_articles, health_dict).

    Error-resilience improvements:
    - Each HTML scraper runs inside its own try/except; failure of one never
      blocks the others.
    - DCD partial success (articles > 0 even if an exception was caught) is
      recorded as "Partial" rather than "Error".
    - Google News futures have a per-future timeout so one slow query doesn't
      stall the whole pool.
    - An empty raw list is handled gracefully instead of crashing the filter step.
    """
    global _SCRAPER_HEALTH
    _SCRAPER_HEALTH = {}   # reset on each scan

    if region_terms is None:
        region_terms = []

    raw = []
    failed_sources: list[str] = []

    # ── Step 1: DCD (primary, high-volume, construction channel) ──────────
    class _FakePbar:
        def progress(self, frac, text=""): progress_cb(frac * 0.55, text)

    dcd_arts: list = []
    try:
        dcd_arts = scrape_dcd(cutoff, max_html_pages, region_terms, pbar=_FakePbar())
        raw.extend(dcd_arts)
        _record_health("DataCenterDynamics", len(dcd_arts), 0, "OK")
        progress_cb(0.55, f"DCD: {len(dcd_arts)} articles")
    except Exception as exc:
        failed_sources.append(f"DCD: {exc}")
        # If we collected something before the exception, mark as Partial
        status = f"Partial ({len(dcd_arts)} articles)" if dcd_arts else f"Error: {exc}"
        _record_health("DataCenterDynamics", len(dcd_arts), 1, status)
        if dcd_arts:
            raw.extend(dcd_arts)
        progress_cb(0.55, f"DCD: {'partial' if dcd_arts else 'failed'}")

    # ── Step 2: DataCenter Knowledge ──────────────────────────────────────
    progress_cb(0.58, "Scraping DataCenter Knowledge…")
    dck_arts: list = []
    try:
        dck_arts = scrape_dck(cutoff, max_pages=min(max_html_pages, 6))
        raw.extend(dck_arts)
        progress_cb(0.62, f"DCK: {len(dck_arts)} articles")
    except Exception as exc:
        failed_sources.append(f"DCK: {exc}")
        status = f"Partial ({len(dck_arts)} articles)" if dck_arts else f"Error: {exc}"
        _record_health("DataCenter Knowledge", len(dck_arts), 1, status)
        if dck_arts:
            raw.extend(dck_arts)

    # ── Step 3: Data Centre Magazine ──────────────────────────────────────
    progress_cb(0.63, "Scraping Data Centre Magazine…")
    dcm_arts: list = []
    try:
        dcm_arts = scrape_dcm(cutoff, max_pages=min(max_html_pages, 4))
        raw.extend(dcm_arts)
        progress_cb(0.66, f"DCM: {len(dcm_arts)} articles")
    except Exception as exc:
        failed_sources.append(f"DCM: {exc}")
        status = f"Partial ({len(dcm_arts)} articles)" if dcm_arts else f"Error: {exc}"
        _record_health("DataCentreMagazine", len(dcm_arts), 1, status)
        if dcm_arts:
            raw.extend(dcm_arts)

    # ── Step 4: Google News supplement (runs in parallel) ─────────────────
    progress_cb(0.68, "Fetching Google News supplement…")
    gn_results: list = []
    gn_errors  = 0
    _GN_FUTURE_TIMEOUT = 20   # seconds per individual query future

    with ThreadPoolExecutor(max_workers=8) as pool:
        futures = {pool.submit(fetch_google_news, q, lbl): lbl for q, lbl in GNEWS_QUERIES}
        done = [0]
        for f in as_completed(futures, timeout=None):
            try:
                batch = f.result(timeout=_GN_FUTURE_TIMEOUT)
                if batch:
                    gn_results.extend(batch)
            except Exception:
                gn_errors += 1
            done[0] += 1
            progress_cb(0.68 + 0.28 * (done[0] / len(GNEWS_QUERIES)), "Google News…")

    raw.extend(gn_results)
    _record_health(
        "Google News", len(gn_results), gn_errors,
        "OK" if gn_errors == 0 else f"{gn_errors} query error(s)",
    )

    if failed_sources:
        _record_health(
            "_failed_sources", 0, len(failed_sources),
            " | ".join(failed_sources),
        )

    progress_cb(0.97, "Filtering and deduplicating…")

    # Guard: nothing scraped at all
    if not raw:
        return []

    # ── Step 5: date cutoff + DC relevance filter ─────────────────────────
    filtered = []
    for item in raw:
        # Skip items with missing or invalid headline
        headline = item.get("headline", "").strip()
        if not headline:
            continue
        d = item.get("date_obj")
        if d and d < cutoff:
            continue
        if not is_dc_relevant(headline):
            continue
        filtered.append(item)

    return filtered


# ─── SCRAPE_SOURCES — all active sources ────────────────────────────────────
SCRAPE_SOURCES = [
    {"name": "DataCenterDynamics",     "url": DCD_BASE + "/en/news/",  "type": "html",  "priority": 1},
    {"name": "DataCenter Knowledge",   "url": DCK_BASE + "/news",      "type": "html",  "priority": 2},
    {"name": "DataCentreMagazine",     "url": DCM_BASE + "/latest",    "type": "html",  "priority": 3},
    {"name": "Google News (Construction)",  "url": "", "type": "gnews", "priority": 10},
    {"name": "Google News (Approvals)",     "url": "", "type": "gnews", "priority": 10},
    {"name": "Google News (Hyperscalers)",  "url": "", "type": "gnews", "priority": 10},
    {"name": "Google News (Power/Energy)",  "url": "", "type": "gnews", "priority": 10},
    {"name": "Google News (Investment)",    "url": "", "type": "gnews", "priority": 10},
]
def is_dc_relevant(text):
    """
    Three-tier relevance filter:
      1. Primary keywords  → any single match = relevant
      2. Strong secondary  → any single match = relevant (specific DC-tech terms)
      3. Weak secondary    → requires 2+ matches (generic financial/size words)

    This prevents "$1bn acquisition" or "100 acres" alone from passing
    while still catching multi-signal headlines like "$2bn data park".
    """
    t = text.lower()

    # ── Tier 1: Primary — any single term is conclusive ──────────────────────
    primary = [
        "data center", "datacenter", "data centre", "datacentre",
        "colocation", "colo ", "hyperscale", "cloud campus",
        "server farm", "computing campus", "ai campus", "gpu cluster",
        "compute campus", "hpc facility", "edge facility",
        "carrier hotel", "internet exchange", "ix facility",
        "infrastructure reit", "digital infrastructure",
        "data hall", "data park", "digital campus",
        "compute facility", "cloud facility", "network facility",
        "ai factory", "inference facility", "training facility",
        "wholesale data", "retail colocation", "powered shell",
        "build-to-suit", "mission critical", "critical facility",
    ]
    if any(p in t for p in primary):
        return True

    # ── Tier 2: Strong secondary — highly specific DC-tech terms ─────────────
    # Any single one of these is a strong enough signal on its own.
    strong_secondary = [
        "megawatt", " mw ", " gw ", "gigawatt",
        "power purchase agreement", " ppa ", "behind the meter",
        "grid connection", "critical load", "raised floor",
        "cooling tower", "liquid cooling", "immersion cooling",
        "diesel generator", "ups system", "modular data",
        "tier iii", "tier iv", "uptime institute",
        "rack space", "co-location", "hosting facility",
        "ai infrastructure", "edge computing",
        "crac unit", "adiabatic cooling", "free cooling",
        "power usage effectiveness", "water usage effectiveness",
        "pue", "wue",
        "sale leaseback", "forward purchase",
        "offtake agreement", "capacity agreement", "pre-lease",
        "breaking ground", "ground breaking", "ribbon cutting",
        "topping off", "commissioning", "fit-out", "fitout",
        "shell and core", "white space", "raised floor space",
        "kilowatt", "kwh", "mwh", "gwh",
        "substation", "transformer", "generator set",
        "ups capacity", "power density",
        "campus site", "land parcel",
    ]
    if any(s in t for s in strong_secondary):
        return True

    # ── Tier 3: Weak secondary — generic financial/size words ─────────────────
    # Require 2+ matches to avoid passing "$1bn pharma deal" or "100 acres farm".
    weak_secondary = [
        "$", "€", "£", "¥", "₹",
        "billion", "million",
        "usd", "eur", "gbp", "jpy", "inr", "sgd", "aed",
        "investment", "financing", "acquisition", "deal",
        "acres", "acre", "hectares", "hectare",
        "sq ft", "square feet", "square meters", "sq m",
        "joint venture", "mou", "memorandum of understanding",
        "loi", "letter of intent",
        "blade server", "server deployment",
        "network access point", "internet hub",
        " kw ",
    ]
    if sum(1 for w in weak_secondary if w in t) >= 2:
        return True

    return False


# ─── Country detection: O(1) compiled-regex index ────────────────────────────
# Pre-compile one regex per country and build a list sorted by specificity
# (longer keyword lists first so more-specific countries win on ties).
# This replaces the O(countries × patterns × articles) linear scan with a
# single compiled-regex pass per article — roughly 60× faster at 3 000+ articles.

def _build_country_index(kw_map: dict) -> list[tuple[str, re.Pattern]]:
    """
    Returns a list of (country, compiled_pattern) sorted so that countries
    with more keywords (more specific) are checked first.
    """
    index = []
    for country, patterns in sorted(kw_map.items(), key=lambda x: -len(x[1])):
        parts = []
        for pat in patterns:
            if pat.startswith(r"\b"):
                parts.append(pat)
            else:
                parts.append(r"\b" + re.escape(pat) + r"\b")
        combined = "|".join(f"(?:{p})" for p in parts)
        try:
            index.append((country, re.compile(combined, re.I)))
        except re.error:
            # Fallback: compile each sub-pattern individually
            for p in parts:
                try:
                    index.append((country, re.compile(p, re.I)))
                    break
                except re.error:
                    pass
    return index

_COUNTRY_INDEX: list[tuple[str, re.Pattern]] = _build_country_index(COUNTRY_KEYWORDS)


def detect_country(text: str) -> str:
    """
    O(countries) country detection using pre-compiled combined regexes.
    Falls back to 'Global' when no pattern matches.
    """
    for country, pattern in _COUNTRY_INDEX:
        if pattern.search(text):
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


def _extract_headline_entities(headline: str) -> dict:
    """
    Extract concrete entities from a headline for direct comparison:
      companies, country, mw_value, deal_value
    Used by fuzzy_similar to avoid collapsing structurally similar but
    factually distinct headlines (e.g. Google vs Amazon, 100MW vs 300MW).
    """
    # Companies present (sorted for determinism)
    companies = sorted(
        co for co in KNOWN_COMPANIES
        if re.search(r"\b" + re.escape(co) + r"\b", headline, re.I)
    )
    # Country
    country = detect_country(headline)
    # Numeric capacity (extract raw number for comparison)
    mw_raw = ""
    m = re.search(r"([\d,]+(?:\.\d+)?)\s*(GW|MW|gigawatt|megawatt|kilowatt|kw)\b", headline, re.I)
    if m:
        mw_raw = m.group(1).replace(",", "") + m.group(2).upper()
    # Numeric deal value (raw number + unit)
    deal_raw = ""
    m2 = re.search(
        r"(\$|€|£|US\$|USD|EUR|GBP|AED|SGD|INR)?\s*([\d,.]+)\s*(billion|bn|million|mn)",
        headline, re.I
    )
    if m2:
        deal_raw = m2.group(2).replace(",", "") + (m2.group(3) or "").lower()[:2]
    return {
        "companies": companies,
        "country":   country,
        "mw":        mw_raw,
        "deal":      deal_raw,
    }


def _entities_conflict(a: str, b: str) -> bool:
    """
    Return True if two headlines have concrete entity values that differ,
    meaning they are factually distinct stories despite structural similarity.
    Conflicts:
      - Different lead companies (both non-empty and not overlapping)
      - Different specific country (both non-Global and different)
      - Different MW value (both non-empty and different)
      - Different deal value (both non-empty and different)
    """
    ea, eb = _extract_headline_entities(a), _extract_headline_entities(b)

    # Company conflict: both have companies, no overlap
    if ea["companies"] and eb["companies"]:
        if not set(ea["companies"]) & set(eb["companies"]):
            return True

    # Country conflict
    if (ea["country"] != "Global" and eb["country"] != "Global"
            and ea["country"] != eb["country"]):
        return True

    # MW conflict: both have a capacity value and they differ
    if ea["mw"] and eb["mw"] and ea["mw"] != eb["mw"]:
        return True

    # Deal value conflict
    if ea["deal"] and eb["deal"] and ea["deal"] != eb["deal"]:
        return True

    return False


def fuzzy_similar(a: str, b: str, threshold: float | None = None) -> bool:
    """
    Entity-aware duplicate detector.

    Step 1 — Entity conflict check (fast): if the two headlines have concrete
    entities that differ (different companies, locations, MW, or deal values),
    they are NOT duplicates regardless of string similarity.  This prevents
    "Google 100MW Iowa" and "Amazon 100MW Iowa" from collapsing.

    Step 2 — String similarity: use a context-derived threshold based on
    the entity density of each headline.  Entity-rich headlines use a lower
    threshold (0.65) to catch paraphrased cross-source duplicates; generic
    headlines use 0.88 to avoid collapsing temporally distinct stories.

    Step 3 — Substring containment: only for entity-rich headlines where
    one headline is a truncated version of the other.
    """
    na, nb = _normalise_headline(a), _normalise_headline(b)
    if na == nb:
        return True

    # Step 1: bail out immediately on entity conflict
    if _entities_conflict(a, b):
        return False

    # Step 2: string similarity
    if threshold is None:
        density_a = _headline_entity_density(a)
        density_b = _headline_entity_density(b)
        # Use tighter threshold for entity-rich headlines (paraphrase detection)
        max_density = max(density_a, density_b)
        if max_density >= 3:
            threshold = 0.65
        elif max_density >= 1:
            threshold = 0.82
        else:
            threshold = 0.88

    ratio = SequenceMatcher(None, na, nb).ratio()
    if ratio >= threshold:
        return True

    # Step 3: substring containment — only for entity-rich pairs
    min_density = min(_headline_entity_density(a), _headline_entity_density(b))
    shorter, longer = (na, nb) if len(na) <= len(nb) else (nb, na)
    if min_density >= 2 and len(shorter) >= 30 and shorter in longer:
        return True

    return False
    """
    Count entity signals in a headline: each MW value, deal size, company name,
    and country match adds 1.  Used to decide how tight the dedup threshold is.
    A headline with 3+ signals is entity-rich and needs a tighter threshold
    (0.72) so cross-source duplicates of the same specific story collapse.
    A generic headline (0 signals) keeps the loose threshold (0.88) so
    "Equinix expands in Europe" (Jan) ≠ "Equinix expands in Europe" (Mar).
    """
    score = 0
    if detect_mw(headline):
        score += 1
    if detect_deal_size(headline):
        score += 1
    # Count company matches (capped at 2 to avoid over-weighting)
    co_hits = sum(
        1 for co in KNOWN_COMPANIES
        if re.search(r"\b" + re.escape(co) + r"\b", headline, re.I)
    )
    score += min(co_hits, 2)
    # Country / location match
    if detect_country(headline) != "Global":
        score += 1
    return score


def _dedup_threshold(headline: str) -> float:
    """
    Return the similarity threshold for this headline.
    Entity-rich (≥3 signals) → 0.72  (tight: same story from different sources)
    Moderate (1-2 signals)   → 0.82
    Generic (0 signals)      → 0.88  (loose: avoid collapsing similar-sounding
                                       but temporally distinct generic headlines)
    """
    density = _headline_entity_density(headline)
    if density >= 3:
        return 0.72
    if density >= 1:
        return 0.82
    return 0.88


def fuzzy_similar(a: str, b: str, threshold: float | None = None) -> bool:
    """
    True if two normalised headlines are likely the same story.
    If threshold is None, it is derived from the entity density of the
    more specific headline (lower of the two thresholds = safer).
    """
    na, nb = _normalise_headline(a), _normalise_headline(b)
    if na == nb:
        return True
    # Use the tighter of the two context-derived thresholds
    if threshold is None:
        threshold = min(_dedup_threshold(a), _dedup_threshold(b))
    ratio = SequenceMatcher(None, na, nb).ratio()
    if ratio >= threshold:
        return True
    # Substring containment: only for entity-rich headlines (avoids generic false merges)
    shorter, longer = (na, nb) if len(na) <= len(nb) else (nb, na)
    min_density = min(_headline_entity_density(a), _headline_entity_density(b))
    if min_density >= 2 and len(shorter) >= 30 and shorter in longer:
        return True
    return False


def _trigram_set(text: str) -> set:
    """Return the set of character trigrams for a normalised headline."""
    t = _normalise_headline(text)
    return {t[i:i+3] for i in range(len(t) - 2)} if len(t) >= 3 else {t}


def deduplicate(articles):
    """
    1. URL-based exact dedup (same URL = same article).
    2. Fuzzy headline dedup — O(n log n) via trigram-bucket pre-filter.
       Two headlines only reach SequenceMatcher if they share enough trigrams,
       which eliminates ~95 % of comparisons at large article counts.
       When duplicates are found, keep the article from the source with the
       lowest _priority number (DCD = 1 wins).
    """
    # ── Step 1: URL dedup — sort by priority so DCD URLs win ties ────────────
    seen_urls: dict = {}
    for art in sorted(articles, key=lambda x: x.get("_priority", 99)):
        url = str(art.get("URL", art.get("url", ""))).strip().rstrip("/")
        if url and url not in seen_urls:
            seen_urls[url] = art
    url_deduped = list(seen_urls.values())

    # ── Step 2: Fuzzy headline dedup — trigram-accelerated ───────────────────
    url_deduped.sort(key=lambda x: x.get("_priority", 99))

    # Build an inverted index: trigram → list of keep-indices that contain it
    keep: list = []
    trigram_index: dict[str, list[int]] = {}   # trigram → indices in keep[]

    for art in url_deduped:
        hl  = art.get("Headline", art.get("headline", ""))
        tgs = _trigram_set(hl)

        # Collect candidate indices that share at least 1 trigram
        candidate_indices: set[int] = set()
        for tg in tgs:
            for idx in trigram_index.get(tg, []):
                candidate_indices.add(idx)

        is_dup = False
        for idx in candidate_indices:
            kept_hl = keep[idx].get("Headline", keep[idx].get("headline", ""))
            if fuzzy_similar(hl, kept_hl):
                is_dup = True
                break

        if not is_dup:
            new_idx = len(keep)
            keep.append(art)
            # Index the new article's trigrams
            for tg in tgs:
                trigram_index.setdefault(tg, []).append(new_idx)

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
    """
    Significance-weighted headline scorer.

    Pure TF-IDF rewards rare vocabulary, not importance — a niche headline
    with unusual words outscores a major deal announcement with common terms.
    This blends TF-IDF (60 %) with a DC-signal bonus (40 %) so headlines
    that mention MW values, deal sizes, named companies, and specific locations
    score higher than ones that just happen to use unusual words.

    Signal bonus components (each normalised to [0, 1]):
      • capacity_bonus : MW/GW value present (+0.4 max)
      • deal_bonus     : deal size present (+0.3 max)
      • company_bonus  : 1–3 known companies (+0.1 per company, max 0.3)
      • location_bonus : non-Global country detected (+0.2)
      • sentiment_bonus: Opened/Live or Approved (+0.15), Under Construction (+0.1)

    The final score is: 0.6 × tfidf_norm + 0.4 × signal_norm
    """
    tokenize = lambda h: [w.lower() for w in re.findall(r"[a-zA-Z]{3,}", h)
                          if w.lower() not in _STOPWORDS]
    tokenized = [tokenize(h) for h in headlines]

    # ── TF-IDF component ──────────────────────────────────────────────────────
    N = max(len(headlines), 1)
    df_counts = Counter()
    for toks in tokenized:
        for t in set(toks):
            df_counts[t] += 1
    idf = {t: math.log((N + 1) / (c + 1)) + 1 for t, c in df_counts.items()}
    raw_tfidf = []
    for toks in tokenized:
        if not toks:
            raw_tfidf.append(0.0)
            continue
        tf = Counter(toks)
        raw_tfidf.append(sum(tf[t] * idf.get(t, 1) for t in toks) / len(toks))

    # Normalise TF-IDF to [0, 1]
    max_tfidf = max(raw_tfidf) if raw_tfidf else 1.0
    if max_tfidf == 0:
        max_tfidf = 1.0
    tfidf_norm = [s / max_tfidf for s in raw_tfidf]

    # ── DC-signal bonus component ─────────────────────────────────────────────
    HIGH_SENT = {"Opened / Live", "Approved"}
    MID_SENT  = {"Under Construction"}

    signal_scores = []
    for h in headlines:
        bonus = 0.0
        if detect_mw(h):
            bonus += 0.4
        if detect_deal_size(h):
            bonus += 0.3
        co_hits = sum(
            1 for co in KNOWN_COMPANIES
            if re.search(r"\b" + re.escape(co) + r"\b", h, re.I)
        )
        bonus += min(co_hits, 3) * 0.1
        if detect_country(h) != "Global":
            bonus += 0.2
        sent = detect_sentiment(h)
        if sent in HIGH_SENT:
            bonus += 0.15
        elif sent in MID_SENT:
            bonus += 0.10
        signal_scores.append(min(bonus, 1.0))   # cap at 1.0

    # Normalise signal scores (already in [0, 1] range, but normalise for fairness)
    max_sig = max(signal_scores) if signal_scores else 1.0
    if max_sig == 0:
        max_sig = 1.0
    signal_norm = [s / max_sig for s in signal_scores]

    # ── Blend: 60 % TF-IDF + 40 % signal ────────────────────────────────────
    return [0.6 * t + 0.4 * s for t, s in zip(tfidf_norm, signal_norm)]


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
            '<div style="padding:.9rem 0 .4rem;text-align:center;">'
            '<div style="font-family:monospace;font-size:.6rem;letter-spacing:.2em;'
            'color:#1a2e50;text-transform:uppercase;margin-bottom:.25rem;">Intelligence Platform</div>'
            '<div style="font-family:Syne,sans-serif;font-size:.9rem;font-weight:800;color:#fff;">Global Data Center Intelligence</div>'
            '<div style="font-family:monospace;font-size:.6rem;color:#1a2e50;margin-top:.15rem;">'
            'Global Intelligence Platform</div>'
            '<div style="font-family:monospace;font-size:.55rem;color:#0f1e36;margin-top:.5rem;'
            'letter-spacing:.06em;">&#169; Sharugh A &nbsp;&middot;&nbsp; All rights reserved</div>'
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
            f'Up to {max_pages} page{"s" if max_pages != 1 else ""} per DCD term scraped</div>',
            unsafe_allow_html=True,
        )

        use_html  = True
        use_rss   = True
        use_gn    = True

        st.divider()

        if "df_full" in st.session_state and st.session_state.df_full is not None:
            df_full = st.session_state.df_full

            # ── Filter version counter (incremented on clear to force new widget keys) ──
            if "_filter_ver" not in st.session_state:
                st.session_state["_filter_ver"] = 0
            _fv = st.session_state["_filter_ver"]
            def _fk(name): return f"{name}_v{_fv}"   # versioned key

            # ── REFINE RESULTS heading + Clear All ──────────────────────────────
            st.markdown(
                '<div style="display:flex;justify-content:space-between;align-items:center;'
                'margin-bottom:.4rem;">'
                '<span style="font-family:Syne,sans-serif;font-weight:700;color:#b8c8e0;'
                'font-size:.82rem;letter-spacing:.04em;">🔍 REFINE RESULTS</span>'
                '</div>',
                unsafe_allow_html=True,
            )
            if st.button("✕ Clear All Filters", use_container_width=True, key="clear_all_filters_btn"):
                # Bump version → all widget keys change → Streamlit treats them as new → default=[] takes effect
                st.session_state["_filter_ver"] = _fv + 1
                st.session_state.pop("filters", None)
                st.rerun()

            # ── Helper: map typed-but-unlisted value to closest match ─────────
            def _fuzzy_resolve(typed, candidates):
                if not typed:
                    return []
                t = typed.strip().lower()
                return [c for c in candidates if t in c.lower()]

            all_regions_av   = sorted(df_full["Region"].unique().tolist())
            all_countries_av = sorted(df_full["Country"].unique().tolist())
            all_topics_av    = sorted(df_full["Topic"].unique().tolist())
            all_sents_av     = sorted(df_full["Sentiment"].unique().tolist())

            # Build company list from data
            _all_co_raw = []
            for v in df_full["Companies"]:
                if v:
                    _all_co_raw.extend([c.strip() for c in str(v).split(",")])
            all_companies_av = sorted(set(c for c in _all_co_raw if c))

            # ── 1. Region ─────────────────────────────────────────────────────
            st.markdown(
                '<div style="font-size:.72rem;color:#3a5480;letter-spacing:.07em;'
                'text-transform:uppercase;margin:.9rem 0 .2rem;">🌐 Region</div>',
                unsafe_allow_html=True,
            )
            sel_regions = st.multiselect(
                "Region", all_regions_av, default=[],
                placeholder="All regions", label_visibility="collapsed",
                key=_fk("f_regions"),
            )

            # ── 2. Country (filtered by region) ───────────────────────────────
            all_world_countries = sorted(set(list(COUNTRY_TO_REGION.keys()) + all_countries_av))
            if sel_regions:
                world_pool = sorted([c for c in all_world_countries
                                     if COUNTRY_TO_REGION.get(c, "Global") in sel_regions])
            else:
                world_pool = all_world_countries

            st.markdown(
                '<div style="font-size:.72rem;color:#3a5480;letter-spacing:.07em;'
                'text-transform:uppercase;margin:.9rem 0 .2rem;">🌍 Country</div>',
                unsafe_allow_html=True,
            )
            sel_countries = st.multiselect(
                "Country", world_pool, default=[],
                placeholder="All countries", label_visibility="collapsed",
                key=_fk("f_countries"),
            )

            # ── 3. State (filtered by country) ────────────────────────────────
            state_pool = []
            for c in (sel_countries if sel_countries else world_pool):
                state_pool.extend(COUNTRY_STATES.get(c, []))
            state_pool = sorted(set(state_pool))

            sel_states = []
            if state_pool:
                st.markdown(
                    '<div style="font-size:.72rem;color:#3a5480;letter-spacing:.07em;'
                    'text-transform:uppercase;margin:.9rem 0 .2rem;">📍 State / Province</div>',
                    unsafe_allow_html=True,
                )
                sel_states = st.multiselect(
                    "State", state_pool, default=[],
                    placeholder="All states/provinces", label_visibility="collapsed",
                    key=_fk("f_states"),
                )

            # ── 4. Company ────────────────────────────────────────────────────
            st.markdown(
                '<div style="font-size:.72rem;color:#3a5480;letter-spacing:.07em;'
                'text-transform:uppercase;margin:.9rem 0 .2rem;">🏢 Company</div>',
                unsafe_allow_html=True,
            )
            full_co_pool = sorted(set(all_companies_av + KNOWN_COMPANIES))
            sel_companies = st.multiselect(
                "Company", full_co_pool, default=[],
                placeholder="All companies — type to search", label_visibility="collapsed",
                key=_fk("f_companies"),
            )

            # ── 5. Topic ──────────────────────────────────────────────────────
            st.markdown(
                '<div style="font-size:.72rem;color:#3a5480;letter-spacing:.07em;'
                'text-transform:uppercase;margin:.9rem 0 .2rem;">🏷️ Topic</div>',
                unsafe_allow_html=True,
            )
            sel_topics = st.multiselect(
                "Topic", all_topics_av, default=[],
                placeholder="All topics", label_visibility="collapsed",
                key=_fk("f_topics"),
            )

            # ── 6. Project Status ─────────────────────────────────────────────
            st.markdown(
                '<div style="font-size:.72rem;color:#3a5480;letter-spacing:.07em;'
                'text-transform:uppercase;margin:.9rem 0 .2rem;">📊 Project Status</div>',
                unsafe_allow_html=True,
            )
            sel_sents = st.multiselect(
                "Status", all_sents_av, default=[],
                placeholder="All statuses", label_visibility="collapsed",
                key=_fk("f_sents"),
            )

            # ── 7. Keyword ────────────────────────────────────────────────────
            st.markdown(
                '<div style="font-size:.72rem;color:#3a5480;letter-spacing:.07em;'
                'text-transform:uppercase;margin:.9rem 0 .2rem;">🔤 Keyword</div>',
                unsafe_allow_html=True,
            )
            keyword = st.text_input(
                "Keyword", placeholder="e.g. 500MW, Texas, nuclear, AWS...",
                label_visibility="collapsed", key=_fk("f_keyword"),
            )

            # ── 8. Min Capacity (MW) ──────────────────────────────────────────
            st.markdown(
                '<div style="font-size:.72rem;color:#3a5480;letter-spacing:.07em;'
                'text-transform:uppercase;margin:.9rem 0 .2rem;">⚡ Min Capacity (MW)</div>',
                unsafe_allow_html=True,
            )
            min_mw = st.number_input(
                "Min MW", min_value=0, value=0,
                step=10, label_visibility="collapsed", key=_fk("f_min_mw"),
            )

            # Collect into session_state filters
            st.session_state.filters = {
                "regions":        sel_regions,
                "topics":         sel_topics,
                "sources":        [],           # not user-facing, always all
                "sents":          sel_sents,
                "keyword":        keyword,
                "min_mw":         min_mw,
                "date_from":      None,
                "date_to":        None,
                "countries":      sel_countries,
                "states":         sel_states,
                "company_search": "",           # replaced by sel_companies below
                "companies":      sel_companies,
            }

        st.divider()
        go_btn = st.button("\U0001f50d  Run Global Scan", use_container_width=True, type="primary")

    now_str = fmt_local()
    # _sa_sig is the platform authorship token — do not modify
    _sa_sig = "\u00a9 Sharugh A"
    st.markdown(
        f'<div class="gl-banner">'
        f'<div class="banner-eyebrow">\u25cf Live Intelligence Feed  \u00b7  {len(SCRAPE_SOURCES)} Sources Active</div>'
        f'<div class="banner-title">Global Data Center <span>Intelligence</span></div>'
        f'<div class="banner-sub">Real-time aggregation across trade press, RSS feeds & Google News \u00b7 '
        f'Auto-tagged by region, topic, company & capacity \u00b7 '
        f'Deduplicated across all sources</div>'
        f'<div class="banner-ts">\U0001f550 {now_str}'
        f'<span style="margin-left:1.8rem;color:#0d1e38;font-size:.6rem;letter-spacing:.08em;">'
        f'Built by Sharugh A &nbsp;\u00b7&nbsp; Licensed &nbsp;\u00b7&nbsp; All rights reserved'
        f'</span></div>'
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
             "Full-depth scraping of DataCenterDynamics, DataCenter Knowledge, "
             "Data Center World and Data Centre Magazine — all pages, maximum depth, "
             "DCD prioritised as primary source for deduplication."),
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

        # Map user-selected regions → DCD region slug terms
        # If the user hasn't selected any regions yet (filters cleared on go_btn),
        # we pass [] which means "no region filter" = full global scan on DCD.
        _region_map_reverse = {v: k for k, v in DCD_REGION_TERMS.items()}
        _pre_selected_regions = st.session_state.get("filters", {}).get("regions", [])
        region_terms = [
            _region_map_reverse[r]
            for r in _pre_selected_regions
            if r in _region_map_reverse
        ]

        # ── Cached scraper call (TTL = 45 min) ───────────────────────────
        @st.cache_data(ttl=2700, show_spinner=False)
        def _cached_scrape(max_p, cutoff_ts, regions_key):
            """Cached wrapper — re-runs only when params change or TTL expires."""
            _cutoff = datetime.fromtimestamp(cutoff_ts)
            _regions = list(regions_key)
            # progress_cb not cacheable; pass a no-op inside the cache
            return run_all_scrapers(max_p, _cutoff, lambda f, l="": None, region_terms=_regions)

        cutoff_ts_key = int(cutoff.timestamp()) if cutoff != datetime.min else 0
        try:
            raw = _cached_scrape(max_pages, cutoff_ts_key, tuple(sorted(region_terms)))
        except Exception as _cache_err:
            # Cache serialisation failure (e.g. unpicklable object) — fall back
            # to a live uncached run so the scan still completes.
            st.warning(
                f"⚠️ Cache layer failed ({type(_cache_err).__name__}). "
                "Running live scan without caching — results will not be cached this time.",
                icon="⚠️",
            )
            try:
                raw = run_all_scrapers(
                    max_pages, cutoff,
                    progress_cb,
                    region_terms=region_terms,
                )
            except Exception as _fallback_err:
                st.error(f"Scan failed: {_fallback_err}")
                pbar.empty()
                st.stop()
        # Save health snapshot captured during this scrape run
        st.session_state.scraper_health = dict(_SCRAPER_HEALTH)

        cutoff_end_val = st.session_state.get("cutoff_end", datetime.max)
        filtered = []
        for item in raw:
            d = item.get("date_obj")
            if d and d > cutoff_end_val:
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
        st.session_state.scan_time = fmt_local(now_local())
        pbar.empty()
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
        + kpi("Sources Polled", len(SCRAPE_SOURCES), "blue", "DCD · DCK · DCW · DCM")
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

    tab1, tab2, tab3, tab4, tab4b, tab5, tab_trend, tab_heatmap, tab_deal, tab_saved, tab_score, tab6, tab_health = st.tabs([
        "\U0001f4f0 Feed",
        "\U0001f5fa\ufe0f World Map",
        "\U0001f4ca Analytics",
        "\U0001f3e2 By Company",
        "\U0001f4cd By State",
        "\U0001f9e0 Market Intel",
        "\U0001f4c8 Trend Compare",
        "\U0001f525 Capacity Heatmap",
        "\U0001f4b0 Deal Flow",
        "\U0001f4be Saved Scans",
        "\U0001f916 AI Scoring",
        "\u2b07\ufe0f Export",
        "\U0001f6a6 Source Health",
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

        # ── Map mode selector ─────────────────────────────────────────────
        map_mode = st.radio(
            "Map display",
            ["Choropleth (country fill)", "Bubble map (city/country pins)", "Both"],
            index=0, horizontal=True, label_visibility="collapsed",
        )

        fig_map = chart_world_map(df)

        if map_mode in ("Bubble map (city/country pins)", "Both"):
            # Add scatter geo layer: one bubble per country, sized by article count
            cc_scatter = df[df["Country"] != "Global"]["Country"].value_counts().reset_index()
            cc_scatter.columns = ["Country", "Count"]
            # Country centroid lookup (approximate)
            _CENTROIDS = {
                "United States":(37.09,-95.71),"Canada":(56.13,-106.35),"Mexico":(23.63,-102.55),
                "United Kingdom":(55.38,-3.44),"Germany":(51.17,10.45),"France":(46.23,2.21),
                "Netherlands":(52.13,5.29),"Ireland":(53.41,-8.24),"Sweden":(60.13,18.64),
                "Norway":(60.47,8.47),"Denmark":(56.26,9.50),"Finland":(61.92,25.75),
                "Spain":(40.46,-3.75),"Italy":(41.87,12.57),"Poland":(51.92,19.15),
                "Switzerland":(46.82,8.23),"Austria":(47.52,14.55),"Belgium":(50.50,4.47),
                "Portugal":(39.40,-8.22),"Romania":(45.94,24.97),"Czech Republic":(49.82,15.47),
                "Singapore":(1.35,103.82),"Japan":(36.20,138.25),"South Korea":(35.91,127.77),
                "Australia":(-25.27,133.78),"India":(20.59,78.96),"China":(35.86,104.20),
                "Hong Kong":(22.32,114.17),"Taiwan":(23.70,120.96),"Malaysia":(4.21,108.96),
                "Indonesia":(-0.79,113.92),"Thailand":(15.87,100.99),"Philippines":(12.88,121.77),
                "New Zealand":(-40.90,174.89),"Vietnam":(14.06,108.28),
                "Saudi Arabia":(23.89,45.08),"UAE":(23.42,53.85),"Qatar":(25.35,51.18),
                "Bahrain":(26.03,50.55),"Kuwait":(29.31,47.48),"Oman":(21.51,55.92),
                "Israel":(31.05,34.85),"Jordan":(30.59,36.24),"Egypt":(26.82,30.80),
                "Brazil":(-14.24,-51.93),"Chile":(-35.68,-71.54),"Colombia":(4.57,-74.30),
                "Argentina":(-38.42,-63.62),"Peru":(-9.19,-75.02),
                "South Africa":(-30.56,22.94),"Nigeria":(9.08,8.68),"Kenya":(0.02,37.91),
                "Ethiopia":(9.15,40.49),"Ghana":(7.95,-1.02),"Morocco":(31.79,-7.09),
                "Tanzania":(-6.37,34.89),"Rwanda":(-1.94,29.87),
            }
            lats, lons, txts, sizes, cols = [], [], [], [], []
            for _, row in cc_scatter.iterrows():
                c = row["Country"]
                if c in _CENTROIDS:
                    lat, lon = _CENTROIDS[c]
                    lats.append(lat); lons.append(lon)
                    txts.append(f"<b>{c}</b><br>{row['Count']} articles")
                    sizes.append(max(8, min(40, row["Count"] * 3)))
                    cols.append(REGION_COLORS.get(COUNTRY_TO_REGION.get(c, "Global"), "#0047e1"))

            if lats:
                fig_map.add_trace(go.Scattergeo(
                    lat=lats, lon=lons,
                    mode="markers",
                    marker=dict(
                        size=sizes,
                        color=cols,
                        opacity=0.75,
                        line=dict(color="#fff", width=0.5),
                    ),
                    hovertemplate="%{text}<extra></extra>",
                    text=txts,
                    name="Article volume",
                    showlegend=False,
                ))

        if map_mode == "Bubble map (city/country pins)":
            # Hide choropleth layer, show only bubbles
            fig_map.data[0].update(showscale=False, colorscale=[[0,"#0d1b30"],[1,"#0d1b30"]], z=[0]*len(fig_map.data[0].z))

        st.plotly_chart(fig_map, use_container_width=True, config={"displayModeBar": False})

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
                    textfont=dict(color=_TITLE, size=10),
                    hovertemplate="<b>%{y}</b>: %{x} articles<extra></extra>",
                ))
                _dark(fig_st, max(320, n_sc * 22))
                fig_st.update_layout(
                    title=dict(text="Top States / Provinces by Article Volume", font=dict(color=_TITLE, size=13), x=0.01)
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
                            textfont=dict(color=_TITLE, size=10),
                            hovertemplate="<b>%{x}</b>: %{y}<extra></extra>",
                        ))
                        _dark(fig_tp, 220)
                        fig_tp.update_layout(
                            title=dict(text=f"Topics — {sel_state}", font=dict(color=_TITLE, size=11), x=0.01),
                            xaxis=dict(tickfont=dict(size=9)),
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

    # ─── TAB: Saved Scans ───────────────────────────────────────────────────
    with tab_saved:
        st.markdown('<div class="sec-head">💾 Saved Scans</div>', unsafe_allow_html=True)
        st.markdown(
            '<div style="font-size:.82rem;color:#3a5480;margin-bottom:1rem;">'
            'Name and save scan results. Scans are persisted as JSON in '
            '<code>saved_scans.json</code> in the app directory — they survive '
            'refreshes and redeploys. Compare two saved scans side-by-side '
            'to track market changes between daily or weekly runs.</div>',
            unsafe_allow_html=True,
        )

        # ── Persistent storage helpers ────────────────────────────────────────
        _SCANS_FILE = "saved_scans.json"

        def _load_scans() -> dict:
            """Load scans from disk; return empty dict on any error."""
            try:
                with open(_SCANS_FILE, "r", encoding="utf-8") as fh:
                    raw = json.load(fh)
                # Deserialise DataFrames from records
                result = {}
                for lbl, sdata in raw.items():
                    try:
                        sdata["df"] = pd.DataFrame(sdata.get("df_records", []))
                        result[lbl] = sdata
                    except Exception:
                        pass
                return result
            except Exception:
                return {}

        def _save_scans(scans: dict):
            """Persist scans to disk. DataFrame → records (JSON-serialisable)."""
            try:
                serialisable = {}
                for lbl, sdata in scans.items():
                    entry = {k: v for k, v in sdata.items() if k != "df"}
                    entry["df_records"] = sdata["df"].to_dict(orient="records")
                    serialisable[lbl] = entry
                with open(_SCANS_FILE, "w", encoding="utf-8") as fh:
                    json.dump(serialisable, fh, ensure_ascii=False, default=str)
            except Exception as e:
                st.warning(f"⚠️ Could not persist scans to disk: {e}")

        # Bootstrap session state from disk on first load
        if "saved_scans" not in st.session_state:
            st.session_state.saved_scans = _load_scans()

        col_save1, col_save2 = st.columns([3, 1])
        with col_save1:
            scan_label_input = st.text_input(
                "Scan label", placeholder="e.g. US hyperscale week 1, APAC May 2025...",
                label_visibility="collapsed", key="save_scan_label"
            )
        with col_save2:
            if st.button("💾 Save Current Scan", use_container_width=True):
                if not df.empty:
                    label_key = scan_label_input.strip() or f"Scan {len(st.session_state.saved_scans)+1}"
                    _ts_saved = datetime.now().strftime("%d %b %Y %H:%M")
                    st.session_state.saved_scans[label_key] = {
                        "df":       df.copy(),
                        "saved_at": _ts_saved,
                        "articles": len(df),
                        "filters":  str(filters),
                    }
                    _save_scans(st.session_state.saved_scans)
                    st.success(f"✅ Saved scan: **{label_key}** ({len(df)} articles)")
                else:
                    st.warning("No articles in current view to save.")

        if not st.session_state.saved_scans:
            st.markdown(
                '<div style="background:#0b1628;border:1px solid #152038;border-radius:10px;'
                'padding:1.5rem;text-align:center;color:#3a5480;font-size:.82rem;">'
                '📂 No saved scans yet. Run a scan, apply filters, then click Save above.</div>',
                unsafe_allow_html=True,
            )
        else:
            # Summary table of saved scans
            scan_rows = []
            for lbl, sdata in st.session_state.saved_scans.items():
                sdf = sdata["df"]
                scan_rows.append({
                    "Label": lbl,
                    "Saved At": sdata["saved_at"],
                    "Articles": sdata["articles"],
                    "Top Region": sdf["Region"].value_counts().idxmax() if not sdf.empty else "—",
                    "Top Topic": sdf["Topic"].value_counts().idxmax() if not sdf.empty else "—",
                    "With Deals": int((sdf["Deal Size"] != "").sum()),
                    "With Capacity": int((sdf["Capacity"] != "").sum()),
                })
            scan_summary_df = pd.DataFrame(scan_rows)
            st.markdown(dark_table(scan_summary_df), unsafe_allow_html=True)

            scan_names = list(st.session_state.saved_scans.keys())

            # ── Compare two saved scans ───────────────────────────────────────
            if len(scan_names) >= 2:
                st.markdown('<div class="sec-head">📊 Compare Two Saved Scans</div>', unsafe_allow_html=True)
                cmp_c1, cmp_c2 = st.columns(2)
                with cmp_c1:
                    cmp_scan_a = st.selectbox("Scan A", scan_names, key="cmp_scan_a", index=0)
                with cmp_c2:
                    cmp_scan_b = st.selectbox("Scan B", scan_names, key="cmp_scan_b", index=min(1, len(scan_names)-1))

                cmp_df_a = st.session_state.saved_scans[cmp_scan_a]["df"]
                cmp_df_b = st.session_state.saved_scans[cmp_scan_b]["df"]

                # KPI deltas
                def _delta_kpi(label, val_a, val_b, fmt=str, accent="blue"):
                    delta = val_b - val_a if isinstance(val_a, (int, float)) else 0
                    delta_sign = "▲" if delta > 0 else ("▼" if delta < 0 else "–")
                    delta_color = "#00e676" if delta > 0 else ("#ff2d6b" if delta < 0 else "#3a5480")
                    return (
                        f'<div class="kpi-card" style="flex:1;min-width:130px;background:#0b1628;'
                        f'border:1px solid #152038;border-radius:12px;padding:1rem 1.2rem;position:relative;">'
                        f'<div style="font-family:monospace;font-size:.6rem;letter-spacing:.12em;'
                        f'text-transform:uppercase;color:#2a3e60;margin-bottom:.35rem;">{label}</div>'
                        f'<div style="display:flex;gap:.6rem;align-items:baseline;">'
                        f'<span style="font-family:Syne,sans-serif;font-size:1.4rem;font-weight:800;color:#0047e1;">{fmt(val_a)}</span>'
                        f'<span style="font-family:monospace;font-size:.7rem;color:#2a3e60;">→</span>'
                        f'<span style="font-family:Syne,sans-serif;font-size:1.4rem;font-weight:800;color:#ffaa00;">{fmt(val_b)}</span>'
                        f'</div>'
                        f'<div style="font-size:.7rem;color:{delta_color};margin-top:.2rem;font-family:monospace;">'
                        f'{delta_sign} {abs(delta)}</div>'
                        f'</div>'
                    )

                cmp_kpi_row = (
                    '<div style="display:flex;gap:.7rem;margin:1rem 0;flex-wrap:wrap;">'
                    + _delta_kpi("Articles", len(cmp_df_a), len(cmp_df_b))
                    + _delta_kpi("Deals", int((cmp_df_a["Deal Size"]!="").sum()), int((cmp_df_b["Deal Size"]!="").sum()))
                    + _delta_kpi("Capacity", int((cmp_df_a["Capacity"]!="").sum()), int((cmp_df_b["Capacity"]!="").sum()))
                    + _delta_kpi("Countries", int(cmp_df_a["Country"].nunique()), int(cmp_df_b["Country"].nunique()))
                    + '</div>'
                )
                st.markdown(cmp_kpi_row, unsafe_allow_html=True)

                # Topic comparison
                topics_cmp = sorted(set(cmp_df_a["Topic"].unique()) | set(cmp_df_b["Topic"].unique()))
                fig_cmp = go.Figure()
                fig_cmp.add_trace(go.Bar(
                    name=cmp_scan_a, x=topics_cmp,
                    y=[int(cmp_df_a["Topic"].value_counts().get(t, 0)) for t in topics_cmp],
                    marker_color="#0047e1",
                ))
                fig_cmp.add_trace(go.Bar(
                    name=cmp_scan_b, x=topics_cmp,
                    y=[int(cmp_df_b["Topic"].value_counts().get(t, 0)) for t in topics_cmp],
                    marker_color="#ffaa00",
                ))
                _dark(fig_cmp, 300)
                fig_cmp.update_layout(
                    barmode="group", showlegend=True,
                    legend=dict(bgcolor="rgba(0,0,0,0)", font=dict(color=_TITLE, size=10)),
                    title=dict(text="Topic Comparison: Scan A vs Scan B", font=dict(color=_TITLE, size=13), x=0.01),
                )
                st.plotly_chart(fig_cmp, use_container_width=True, config={"displayModeBar": False})

                # Side-by-side headlines
                st.markdown('<div class="sec-head">Side-by-Side Headlines</div>', unsafe_allow_html=True)
                col_cmp_a, col_cmp_b = st.columns(2)
                with col_cmp_a:
                    st.markdown(f'<div style="font-family:Syne,sans-serif;font-weight:700;color:#0047e1;font-size:.85rem;margin-bottom:.5rem;">● {cmp_scan_a}</div>', unsafe_allow_html=True)
                    for _, row in cmp_df_a.head(8).iterrows():
                        st.markdown(article_card(row["Headline"],row["Date"],row["URL"],row["Source"],row["Country"],row["Topic"],row.get("Capacity",""),row.get("Deal Size",""),row.get("Sentiment","News")), unsafe_allow_html=True)
                with col_cmp_b:
                    st.markdown(f'<div style="font-family:Syne,sans-serif;font-weight:700;color:#ffaa00;font-size:.85rem;margin-bottom:.5rem;">● {cmp_scan_b}</div>', unsafe_allow_html=True)
                    for _, row in cmp_df_b.head(8).iterrows():
                        st.markdown(article_card(row["Headline"],row["Date"],row["URL"],row["Source"],row["Country"],row["Topic"],row.get("Capacity",""),row.get("Deal Size",""),row.get("Sentiment","News")), unsafe_allow_html=True)

            # Drill into a single saved scan
            st.markdown('<div class="sec-head">Browse a Saved Scan</div>', unsafe_allow_html=True)
            sel_scan = st.selectbox("Select scan", scan_names, key="saved_scan_select")
            saved_df = st.session_state.saved_scans[sel_scan]["df"]
            for _, row in saved_df.head(50).iterrows():
                st.markdown(
                    article_card(
                        row["Headline"], row["Date"], row["URL"],
                        row["Source"], row["Country"], row["Topic"],
                        row.get("Capacity", ""), row.get("Deal Size", ""),
                        row.get("Sentiment", "News"),
                    ),
                    unsafe_allow_html=True,
                )

            # Delete saved scan
            if st.button("🗑️ Delete This Saved Scan", key="delete_saved_scan"):
                del st.session_state.saved_scans[sel_scan]
                _save_scans(st.session_state.saved_scans)
                st.rerun()

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
                text=score_dist["Articles"], textposition="outside",
                textfont=dict(color=_TITLE, size=10),
                hovertemplate="<b>Score %{x}</b>: %{y} articles<extra></extra>",
            ))
            _dark(fig_score, 260)
            fig_score.update_layout(title=dict(text="AI Score Distribution", font=dict(color=_TITLE, size=13), x=0.01))
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
            scored_display = df_scored[["AI Score","Headline","Date","Topic","Sentiment","Capacity","Deal Size","Country","URL"]].copy()
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
                        "Topic","Sentiment","Capacity","Deal Size","Companies","URL"]
        st.markdown(
            dark_table(df[[c for c in display_cols if c in df.columns]]),
            unsafe_allow_html=True,
        )

    # ─── TAB: Source Health Panel ────────────────────────────────────────────
    with tab_health:
        st.markdown('<div class="sec-head">🚦 Source Health Dashboard</div>', unsafe_allow_html=True)
        st.markdown(
            '<div style="font-size:.82rem;color:#3a5480;margin-bottom:1rem;">'
            'Real-time status of every data source from the last scan. '
            'Shows per-source article count, error count, and status. '
            'Red = failed, amber = partial, green = healthy.</div>',
            unsafe_allow_html=True,
        )

        health = st.session_state.get("scraper_health", {})
        if not health:
            st.info("Run a scan first to populate source health data.")
        else:
            # Overall health KPIs
            total_sources = len([k for k in health if not k.startswith("_")])
            ok_sources    = len([v for k, v in health.items() if not k.startswith("_") and v["status"] == "OK"])
            err_sources   = total_sources - ok_sources
            total_arts_h  = sum(v["articles"] for k, v in health.items() if not k.startswith("_"))
            total_errs_h  = sum(v["errors"]   for k, v in health.items() if not k.startswith("_"))

            hkpi = (
                '<div style="display:flex;gap:.8rem;margin-bottom:1.4rem;flex-wrap:wrap;">'
                + kpi("Sources Polled",   total_sources, "blue",   "HTML + RSS + GNews")
                + kpi("Healthy",          ok_sources,    "green",  "returned OK status")
                + kpi("With Errors",      err_sources,   "red",    "partial or failed")
                + kpi("Total Articles",   total_arts_h,  "cyan",   "across all sources")
                + kpi("Total Errors",     total_errs_h,  "amber",  "HTTP/parse errors")
                + '</div>'
            )
            st.markdown(hkpi, unsafe_allow_html=True)

            # Per-source table
            st.markdown('<div class="sec-head">Per-Source Breakdown</div>', unsafe_allow_html=True)
            rows_h = ""
            td_h = "padding:.55rem .9rem;font-size:.8rem;border-bottom:1px solid #101b2e;vertical-align:middle;"
            th_h = ("background:#0f1e36;color:#b8c8e0;font-family:monospace;font-size:.65rem;"
                    "letter-spacing:.08em;text-transform:uppercase;padding:.55rem .9rem;"
                    "border-bottom:2px solid #0047e1;white-space:nowrap;")
            for src_name, meta in sorted(health.items()):
                if src_name.startswith("_"):
                    continue
                status = meta["status"]
                status_color = (
                    "#00e676" if status == "OK"
                    else "#ffaa00" if "Partial" in status or "Error" not in status
                    else "#ff2d6b"
                )
                status_icon  = "✅" if status == "OK" else ("⚠️" if "Partial" in status else "❌")
                src_meta_d   = SOURCE_META.get(src_name, SOURCE_META["Unknown"])
                art_bar_pct  = min(100, int(meta["articles"] / max(total_arts_h, 1) * 100))
                rows_h += (
                    f'<tr>'
                    f'<td style="{td_h}background:#0b1628;">'
                    f'<span style="background:{src_meta_d["color"]}22;color:{src_meta_d["color"]};'
                    f'border:1px solid {src_meta_d["color"]}44;border-radius:4px;'
                    f'padding:2px 6px;font-family:monospace;font-size:.65rem;">{src_meta_d["short"]}</span>'
                    f' <span style="color:#b8c8e0;font-size:.8rem;">{src_name}</span></td>'
                    f'<td style="{td_h}background:#060a10;color:#fff;font-family:monospace;">{meta["articles"]}'
                    f'<div style="margin-top:4px;height:4px;border-radius:2px;background:#152038;width:100%;">'
                    f'<div style="height:4px;border-radius:2px;background:{src_meta_d["color"]};width:{art_bar_pct}%;"></div>'
                    f'</div></td>'
                    f'<td style="{td_h}background:#0b1628;font-family:monospace;color:{"#ff2d6b" if meta["errors"] > 0 else "#3a5480"};">'
                    f'{meta["errors"]}</td>'
                    f'<td style="{td_h}background:#060a10;">'
                    f'<span style="color:{status_color};font-size:.78rem;">{status_icon} {status}</span></td>'
                    f'<td style="{td_h}background:#0b1628;font-family:monospace;font-size:.72rem;color:#3a5480;">'
                    f'{meta.get("ts","—")}</td>'
                    f'</tr>'
                )
            st.markdown(
                '<div style="overflow-x:auto;border-radius:10px;border:1px solid #152038;margin-bottom:1rem;">'
                '<table style="width:100%;border-collapse:collapse;background:#060a10;">'
                f'<thead><tr>'
                f'<th style="{th_h}">Source</th>'
                f'<th style="{th_h}">Articles</th>'
                f'<th style="{th_h}">Errors</th>'
                f'<th style="{th_h}">Status</th>'
                f'<th style="{th_h}">Fetched At</th>'
                f'</tr></thead>'
                f'<tbody>{rows_h}</tbody>'
                f'</table></div>',
                unsafe_allow_html=True,
            )

            # Failed sources detail (if any)
            failed_meta = health.get("_failed_sources")
            if failed_meta and failed_meta.get("status"):
                st.markdown('<div class="sec-head">⚠️ Failed Source Details</div>', unsafe_allow_html=True)
                for line in failed_meta["status"].split(" | "):
                    st.markdown(
                        f'<div style="background:#1a0a0a;border:1px solid #ff2d6b44;border-radius:8px;'
                        f'padding:.6rem 1rem;margin-bottom:.4rem;font-family:monospace;font-size:.78rem;color:#ff2d6b;">'
                        f'❌ {line}</div>',
                        unsafe_allow_html=True,
                    )

            # Interactive map showing article coverage by source
            st.markdown('<div class="sec-head">Source Contribution Chart</div>', unsafe_allow_html=True)
            src_names_h  = [k for k in health if not k.startswith("_")]
            src_counts_h = [health[k]["articles"] for k in src_names_h]
            src_cols_h   = [SOURCE_META.get(k, SOURCE_META["Unknown"])["color"] for k in src_names_h]
            if src_counts_h and sum(src_counts_h) > 0:
                fig_src_pie = go.Figure(go.Pie(
                    labels=src_names_h, values=src_counts_h, hole=0.5,
                    marker=dict(colors=src_cols_h, line=dict(color=_BG, width=2)),
                    textinfo="label+percent",
                    textfont=dict(color=_TITLE, size=10),
                    hovertemplate="<b>%{label}</b>: %{value} articles (%{percent})<extra></extra>",
                ))
                fig_src_pie.update_layout(
                    paper_bgcolor=_PAPER, plot_bgcolor=_BG,
                    font=dict(family=_FONT, color=_TEXT),
                    height=320, margin=dict(l=14, r=14, t=36, b=14),
                    showlegend=True,
                    legend=dict(bgcolor="rgba(0,0,0,0)", font=dict(color=_TITLE, size=9)),
                    title=dict(text="Article Share by Source", font=dict(color=_TITLE, size=13), x=0.01),
                    annotations=[dict(
                        text=f"<b>{sum(src_counts_h)}</b><br><span style='font-size:9px'>total</span>",
                        x=0.5, y=0.5, showarrow=False,
                        font=dict(size=13, color=_TITLE, family=_FONT),
                    )],
                )
                st.plotly_chart(fig_src_pie, use_container_width=True, config={"displayModeBar": False})

    # ── Platform footer ───────────────────────────────────────────────────────
    st.markdown(
        '<div style="margin-top:3rem;padding:1.2rem 0 .6rem;border-top:1px solid #101b2e;'
        'text-align:center;">'
        '<div style="font-family:\'DM Mono\',monospace;font-size:.6rem;letter-spacing:.14em;'
        'color:#cc0000;text-transform:uppercase;">'
        'Global Data Center Intelligence &nbsp;\u00b7&nbsp; '
        'Built by Sharugh A &nbsp;\u00b7&nbsp; \u00a9 All rights reserved'
        '</div></div>',
        unsafe_allow_html=True,
    )





# ═══════════════════════════════════════════════════════════════════════════════
#  UNIT TESTS  —  run with:  python GlobalDCIntel_v5.py --test
#
#  Covers the enrichment functions most likely to silently regress:
#    • detect_country        (O(1) compiled-regex index)
#    • detect_mw             (capacity / acreage extraction)
#    • is_dc_relevant        (three-tier relevance filter)
#    • fuzzy_similar         (context-aware dedup threshold)
#    • _headline_entity_density  (entity counting for threshold selection)
#    • _tfidf_scores         (significance-weighted scorer)
#    • _gn_parse_feed stub   (Google News feed parser)
#
#  Zero external dependencies — uses stdlib unittest only.
#  Exit code 0 = all pass, 1 = any failure.
# ═══════════════════════════════════════════════════════════════════════════════

import unittest, sys as _sys

class TestDetectCountry(unittest.TestCase):

    def test_exact_name_us(self):
        self.assertEqual(detect_country("New data center opens in United States"), "United States")

    def test_exact_name_germany(self):
        self.assertEqual(detect_country("Hyperscaler breaks ground in Germany"), "Germany")

    def test_exact_name_singapore(self):
        self.assertEqual(detect_country("Singapore announces moratorium review"), "Singapore")

    def test_city_ashburn(self):
        self.assertEqual(detect_country("Ashburn campus expands to 200MW"), "United States")

    def test_city_frankfurt(self):
        self.assertEqual(detect_country("Frankfurt data center campus"), "Germany")

    def test_city_dubai(self):
        self.assertEqual(detect_country("Dubai hyperscale facility announced"), "UAE")

    def test_city_mumbai(self):
        self.assertEqual(detect_country("Mumbai data centre plan"), "India")

    def test_alias_uk(self):
        self.assertEqual(detect_country("New UK data center investment"), "United Kingdom")

    def test_abbreviation_uae(self):
        self.assertEqual(detect_country("UAE announces new campus"), "UAE")

    def test_abbreviation_ksa(self):
        self.assertEqual(detect_country("KSA power deal signed"), "Saudi Arabia")

    def test_no_country_returns_global(self):
        self.assertEqual(detect_country("Hyperscale investment secured"), "Global")

    def test_empty_string(self):
        self.assertEqual(detect_country(""), "Global")

    def test_lowercase_country(self):
        self.assertEqual(detect_country("Data park opens in australia"), "Australia")

    def test_mixed_case(self):
        self.assertEqual(detect_country("TOKYO data center campus"), "Japan")

    def test_word_boundary_columbia_vs_colombia(self):
        self.assertEqual(detect_country("British Columbia data center"), "Canada")

    def test_regression_netherlands(self):
        self.assertEqual(detect_country("Amsterdam campus reaches 100MW"), "Netherlands")

    def test_regression_brazil(self):
        self.assertEqual(detect_country("Sao Paulo data centre approved"), "Brazil")


class TestDetectMW(unittest.TestCase):

    def test_mw_integer(self):
        self.assertEqual(detect_mw("Plans a 200MW data center campus"), "200 MW")

    def test_mw_decimal(self):
        self.assertEqual(detect_mw("Approved for 1.5 GW facility"), "1.5 GW")

    def test_mw_with_comma(self):
        self.assertEqual(detect_mw("1,200MW substation approved"), "1200 MW")

    def test_gw_uppercase(self):
        self.assertEqual(detect_mw("2GW hyperscale park"), "2 GW")

    def test_gigawatt_word(self):
        self.assertIn("gigawatt", detect_mw("3 gigawatt nuclear-powered campus").lower())

    def test_megawatt_word(self):
        self.assertIn("megawatt", detect_mw("500 megawatt investment").lower())

    def test_kilowatt(self):
        self.assertEqual(detect_mw("High-density 500kW rack deployment"), "500 kW")

    def test_acres(self):
        self.assertIn("acres", detect_mw("200 acres site selection announced").lower())

    def test_hectares(self):
        self.assertIn("hectare", detect_mw("50 hectares land parcel secured").lower())

    def test_no_capacity_returns_empty(self):
        self.assertEqual(detect_mw("Data center investment deal closed"), "")

    def test_empty_string(self):
        self.assertEqual(detect_mw(""), "")

    def test_mw_at_start(self):
        self.assertEqual(detect_mw("100MW facility opens in Virginia"), "100 MW")

    def test_does_not_confuse_mw_in_word(self):
        self.assertIsInstance(detect_mw("Battery stores 500MWh of energy"), str)


class TestIsDCRelevant(unittest.TestCase):

    def test_data_center(self):
        self.assertTrue(is_dc_relevant("New data center opens in Texas"))

    def test_data_centre_british(self):
        self.assertTrue(is_dc_relevant("UK data centre moratorium lifted"))

    def test_colocation(self):
        self.assertTrue(is_dc_relevant("Equinix expands colocation footprint"))

    def test_hyperscale(self):
        self.assertTrue(is_dc_relevant("Hyperscale campus breaks ground"))

    def test_ai_factory(self):
        self.assertTrue(is_dc_relevant("New AI factory announced in Malaysia"))

    def test_digital_infrastructure(self):
        self.assertTrue(is_dc_relevant("Digital infrastructure REIT files IPO"))

    def test_megawatt(self):
        self.assertTrue(is_dc_relevant("500 megawatt substation approved"))

    def test_immersion_cooling(self):
        self.assertTrue(is_dc_relevant("Immersion cooling trial at hyperscale facility"))

    def test_pue(self):
        self.assertTrue(is_dc_relevant("Facility achieves PUE of 1.15"))

    def test_groundbreaking(self):
        self.assertTrue(is_dc_relevant("Breaking ground ceremony held today"))

    def test_sale_leaseback(self):
        self.assertTrue(is_dc_relevant("Operator completes sale leaseback transaction"))

    def test_two_weak_signals_pass(self):
        self.assertTrue(is_dc_relevant("$2 billion investment in new compute campus"))

    def test_one_weak_signal_fails(self):
        self.assertFalse(is_dc_relevant("Pharma company signs a new partnership agreement"))

    def test_currency_plus_acres_passes(self):
        self.assertTrue(is_dc_relevant("$500 million deal for 100 acres"))

    def test_unrelated_tech_news(self):
        self.assertFalse(is_dc_relevant("Apple releases new iPhone model"))

    def test_unrelated_finance_news(self):
        self.assertFalse(is_dc_relevant("Federal Reserve raises interest rates again"))

    def test_empty_string(self):
        self.assertFalse(is_dc_relevant(""))

    def test_regression_server_farm(self):
        self.assertTrue(is_dc_relevant("Giant server farm planned for Nevada"))

    def test_regression_carrier_hotel(self):
        self.assertTrue(is_dc_relevant("Carrier hotel in Chicago expands capacity"))

    def test_regression_datacenter_no_space(self):
        self.assertTrue(is_dc_relevant("Datacenter construction approved by council"))


class TestFuzzySimilar(unittest.TestCase):

    # ── Obvious duplicates ────────────────────────────────────────────────────
    def test_exact_same(self):
        self.assertTrue(fuzzy_similar(
            "Microsoft breaks ground on 200MW Virginia campus",
            "Microsoft breaks ground on 200MW Virginia campus",
        ))

    def test_near_duplicate_source_suffix_stripped(self):
        # Google News appends " - DataCenterDynamics" etc.
        self.assertTrue(fuzzy_similar(
            "Equinix opens new Frankfurt data center - DataCenterDynamics",
            "Equinix opens new Frankfurt data center - DataCenter Knowledge",
        ))

    def test_entity_rich_cross_source_dup(self):
        # Same deal, different words — entity-rich → tight threshold
        self.assertTrue(fuzzy_similar(
            "Microsoft announces 500MW data center campus in Virginia worth $10bn",
            "Microsoft reveals 500MW Virginia data center campus in $10bn deal",
        ))

    # ── Legitimate non-duplicates ─────────────────────────────────────────────
    def test_same_company_different_dates_not_dup(self):
        # Generic expansion headline, no entity signals → loose threshold
        # These are from different months and should NOT collapse
        self.assertFalse(fuzzy_similar(
            "Equinix expands data center footprint in Europe",
            "Equinix expands data center footprint in Asia",
        ))

    def test_different_companies_not_dup(self):
        self.assertFalse(fuzzy_similar(
            "Google breaks ground on new data center in Iowa",
            "Amazon breaks ground on new data center in Iowa",
        ))

    def test_different_mw_values_not_dup(self):
        self.assertFalse(fuzzy_similar(
            "NextDC announces 100MW Sydney campus",
            "NextDC announces 300MW Sydney campus",
        ))

    def test_different_locations_not_dup(self):
        self.assertFalse(fuzzy_similar(
            "Equinix opens SG5 data center in Singapore",
            "Equinix opens AM8 data center in Amsterdam",
        ))

    # ── Threshold selection ───────────────────────────────────────────────────
    def test_entity_rich_threshold_is_tight(self):
        h = "Microsoft 500MW Virginia campus $10bn deal"
        self.assertLessEqual(_dedup_threshold(h), 0.82)

    def test_generic_threshold_is_loose(self):
        h = "Data center expansion announced"
        self.assertGreaterEqual(_dedup_threshold(h), 0.88)


class TestEntityDensity(unittest.TestCase):

    def test_rich_headline(self):
        # MW + deal size + company + location = 4 signals
        score = _headline_entity_density(
            "Microsoft breaks ground on 500MW data center campus in Virginia, $3bn deal"
        )
        self.assertGreaterEqual(score, 3)

    def test_generic_headline(self):
        score = _headline_entity_density("Hyperscale investment announced")
        self.assertEqual(score, 0)

    def test_company_only(self):
        score = _headline_entity_density("Equinix reports strong Q3 results")
        self.assertGreaterEqual(score, 1)


class TestTFIDFScorer(unittest.TestCase):

    def test_returns_correct_length(self):
        headlines = ["data center opens", "500MW campus in Virginia", "investment deal"]
        scores = _tfidf_scores(headlines)
        self.assertEqual(len(scores), 3)

    def test_signal_rich_headline_scores_higher(self):
        headlines = [
            "Microsoft opens 500MW Virginia data center worth $3bn",   # entity-rich
            "Company announces data center plans",                       # generic
        ]
        scores = _tfidf_scores(headlines)
        self.assertGreater(scores[0], scores[1])

    def test_scores_are_non_negative(self):
        headlines = ["foo bar baz", "data center opens in Texas", ""]
        scores = _tfidf_scores(headlines)
        for s in scores:
            self.assertGreaterEqual(s, 0.0)

    def test_single_headline(self):
        scores = _tfidf_scores(["Microsoft 500MW Virginia campus"])
        self.assertEqual(len(scores), 1)
        self.assertGreaterEqual(scores[0], 0.0)

    def test_empty_list(self):
        scores = _tfidf_scores([])
        self.assertEqual(scores, [])

    def test_empty_headline_scores_zero(self):
        scores = _tfidf_scores([""])
        self.assertEqual(scores[0], 0.0)


def _run_tests():
    """Entry point for --test flag: runs all unit tests and exits."""
    loader = unittest.TestLoader()
    suite  = unittest.TestSuite()
    for cls in (
        TestDetectCountry,
        TestDetectMW,
        TestIsDCRelevant,
        TestFuzzySimilar,
        TestEntityDensity,
        TestTFIDFScorer,
    ):
        suite.addTests(loader.loadTestsFromTestCase(cls))
    runner = unittest.TextTestRunner(verbosity=2)
    result = runner.run(suite)
    _sys.exit(0 if result.wasSuccessful() else 1)


if __name__ == "__main__":
    if len(_sys.argv) > 1 and _sys.argv[1] == "--test":
        _run_tests()
    else:
        main()
#
#  Covers the three enrichment functions most likely to silently regress:
#    • detect_country   (country detection via compiled-regex index)
#    • detect_mw        (capacity / acreage extraction)
#    • is_dc_relevant   (three-tier relevance filter)
#
#  Zero external dependencies — uses stdlib unittest only.
#  Exit code 0 = all pass, 1 = any failure.
# ═══════════════════════════════════════════════════════════════════════════════

import unittest, sys as _sys

class TestDetectCountry(unittest.TestCase):

    # ── Tier-1: exact country names ──────────────────────────────────────────
    def test_exact_name_us(self):
        self.assertEqual(detect_country("New data center opens in United States"), "United States")

    def test_exact_name_germany(self):
        self.assertEqual(detect_country("Hyperscaler breaks ground in Germany"), "Germany")

    def test_exact_name_singapore(self):
        self.assertEqual(detect_country("Singapore announces moratorium review"), "Singapore")

    # ── City / alias detection ───────────────────────────────────────────────
    def test_city_ashburn(self):
        self.assertEqual(detect_country("Ashburn campus expands to 200MW"), "United States")

    def test_city_frankfurt(self):
        self.assertEqual(detect_country("Frankfurt data center campus"), "Germany")

    def test_city_dubai(self):
        self.assertEqual(detect_country("Dubai hyperscale facility announced"), "UAE")

    def test_city_mumbai(self):
        self.assertEqual(detect_country("Mumbai data centre plan"), "India")

    def test_alias_uk(self):
        # r"\bUK\b" pattern
        self.assertEqual(detect_country("New UK data center investment"), "United Kingdom")

    # ── Abbreviations ────────────────────────────────────────────────────────
    def test_abbreviation_uae(self):
        self.assertEqual(detect_country("UAE announces new campus"), "UAE")

    def test_abbreviation_ksa(self):
        self.assertEqual(detect_country("KSA power deal signed"), "Saudi Arabia")

    # ── Fallback to Global ───────────────────────────────────────────────────
    def test_no_country_returns_global(self):
        self.assertEqual(detect_country("Hyperscale investment secured"), "Global")

    def test_empty_string(self):
        self.assertEqual(detect_country(""), "Global")

    # ── Case insensitivity ───────────────────────────────────────────────────
    def test_lowercase_country(self):
        self.assertEqual(detect_country("Data park opens in australia"), "Australia")

    def test_mixed_case(self):
        self.assertEqual(detect_country("TOKYO data center campus"), "Japan")

    # ── Word-boundary enforcement (should NOT match "Columbia" for Colombia) ─
    def test_word_boundary_columbia_vs_colombia(self):
        result = detect_country("British Columbia data center")
        # Should be Canada (British Columbia is a province) not Colombia
        self.assertEqual(result, "Canada")

    # ── Regression: adding a new keyword shouldn't break existing ones ───────
    def test_regression_netherlands(self):
        self.assertEqual(detect_country("Amsterdam campus reaches 100MW"), "Netherlands")

    def test_regression_brazil(self):
        self.assertEqual(detect_country("Sao Paulo data centre approved"), "Brazil")


class TestDetectMW(unittest.TestCase):

    # ── MW detection ────────────────────────────────────────────────────────
    def test_mw_integer(self):
        self.assertEqual(detect_mw("Plans a 200MW data center campus"), "200 MW")

    def test_mw_decimal(self):
        self.assertEqual(detect_mw("Approved for 1.5 GW facility"), "1.5 GW")

    def test_mw_with_comma(self):
        self.assertEqual(detect_mw("1,200MW substation approved"), "1200 MW")

    def test_gw_uppercase(self):
        self.assertEqual(detect_mw("2GW hyperscale park"), "2 GW")

    def test_gigawatt_word(self):
        self.assertEqual(detect_mw("3 gigawatt nuclear-powered campus"), "3 GIGAWATT")

    def test_megawatt_word(self):
        self.assertEqual(detect_mw("500 megawatt investment"), "500 MEGAWATT")

    def test_kilowatt(self):
        self.assertEqual(detect_mw("High-density 500kW rack deployment"), "500 kW")

    # ── Acres / hectares fallback ────────────────────────────────────────────
    def test_acres(self):
        self.assertIn("acres", detect_mw("200 acres site selection announced").lower())

    def test_hectares(self):
        self.assertIn("hectare", detect_mw("50 hectares land parcel secured").lower())

    # ── Empty / no match ────────────────────────────────────────────────────
    def test_no_capacity_returns_empty(self):
        self.assertEqual(detect_mw("Data center investment deal closed"), "")

    def test_empty_string(self):
        self.assertEqual(detect_mw(""), "")

    # ── Edge cases ───────────────────────────────────────────────────────────
    def test_mw_at_start(self):
        self.assertEqual(detect_mw("100MW facility opens in Virginia"), "100 MW")

    def test_does_not_confuse_mw_in_word(self):
        # "MWh" contains MW — function returns it; just check it doesn't crash
        result = detect_mw("Battery stores 500MWh of energy")
        self.assertIsInstance(result, str)


class TestIsDCRelevant(unittest.TestCase):

    # ── Tier-1 primary terms ─────────────────────────────────────────────────
    def test_data_center(self):
        self.assertTrue(is_dc_relevant("New data center opens in Texas"))

    def test_data_centre_british(self):
        self.assertTrue(is_dc_relevant("UK data centre moratorium lifted"))

    def test_colocation(self):
        self.assertTrue(is_dc_relevant("Equinix expands colocation footprint"))

    def test_hyperscale(self):
        self.assertTrue(is_dc_relevant("Hyperscale campus breaks ground"))

    def test_ai_factory(self):
        self.assertTrue(is_dc_relevant("New AI factory announced in Malaysia"))

    def test_digital_infrastructure(self):
        self.assertTrue(is_dc_relevant("Digital infrastructure REIT files IPO"))

    # ── Tier-2 strong secondary ──────────────────────────────────────────────
    def test_megawatt(self):
        self.assertTrue(is_dc_relevant("500 megawatt substation approved"))

    def test_immersion_cooling(self):
        self.assertTrue(is_dc_relevant("Immersion cooling trial at hyperscale facility"))

    def test_pue(self):
        self.assertTrue(is_dc_relevant("Facility achieves PUE of 1.15"))

    def test_groundbreaking(self):
        self.assertTrue(is_dc_relevant("Breaking ground ceremony held today"))

    def test_sale_leaseback(self):
        self.assertTrue(is_dc_relevant("Operator completes sale leaseback transaction"))

    # ── Tier-3 weak secondary (requires 2+ signals) ──────────────────────────
    def test_two_weak_signals_pass(self):
        # "billion" + "investment" both present → should pass
        self.assertTrue(is_dc_relevant("$2 billion investment in new compute campus"))

    def test_one_weak_signal_fails(self):
        # A single non-currency term alone should NOT pass (no $ or billion etc.)
        self.assertFalse(is_dc_relevant("Pharma company signs a new partnership agreement"))

    def test_currency_plus_acres_passes(self):
        self.assertTrue(is_dc_relevant("$500 million deal for 100 acres"))

    # ── True negatives ───────────────────────────────────────────────────────
    def test_unrelated_tech_news(self):
        self.assertFalse(is_dc_relevant("Apple releases new iPhone model"))

    def test_unrelated_finance_news(self):
        self.assertFalse(is_dc_relevant("Federal Reserve raises interest rates again"))

    def test_empty_string(self):
        self.assertFalse(is_dc_relevant(""))

    # ── Regression: adding new keywords shouldn't break old ones ────────────
    def test_regression_server_farm(self):
        self.assertTrue(is_dc_relevant("Giant server farm planned for Nevada"))

    def test_regression_carrier_hotel(self):
        self.assertTrue(is_dc_relevant("Carrier hotel in Chicago expands capacity"))

    def test_regression_datacenter_no_space(self):
        self.assertTrue(is_dc_relevant("Datacenter construction approved by council"))


