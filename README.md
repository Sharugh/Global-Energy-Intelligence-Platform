# ⚡ Global Energy Intelligence Platform

> **Real-time market intelligence for Data Centers & Renewables Power Markets — built for analysts, strategists, and executives who need structured, actionable insight from live industry news.**

[![Streamlit](https://img.shields.io/badge/Built%20with-Streamlit-FF4B4B?style=flat&logo=streamlit&logoColor=white)](https://streamlit.io)
[![Python](https://img.shields.io/badge/Python-3.10%2B-3776AB?style=flat&logo=python&logoColor=white)](https://python.org)
[![License](https://img.shields.io/badge/License-Proprietary-lightgrey?style=flat)](./LICENSE)
[![Platform](https://img.shields.io/badge/Platform-Streamlit%20Cloud-FF4B4B?style=flat)](https://streamlit.io/cloud)

---

## 📌 Overview

The **Global Energy Intelligence Platform** is a Streamlit web application that continuously scrapes, enriches, and analyses live news from the global data center and renewable energy sectors. It transforms raw industry headlines into structured, senior-analyst-grade market intelligence — covering deal flow, capacity pipelines, company activity, regulatory dynamics, and regional breakdowns — without requiring any external AI API key.

The platform has two fully independent intelligence modules, switchable from the sidebar:

| Module | Coverage |
|---|---|
| 🏢 **Data Center Markets** | Hyperscale, colocation, AI/GPU campuses, construction, power procurement, permitting |
| 🌱 **Renewables Power Markets** | Solar, Offshore Wind, Onshore Wind, Energy Storage/BESS, Hydrogen, Other Renewables |

---

## ✨ Key Features

### 🔍 Live Data Ingestion
- Scrapes **DataCenterDynamics** (Construction Channel + General News) directly via CloudScraper for the DC module
- Pulls **100+ specialist RSS feeds** for renewables: PV Magazine, PV Tech, Wind Power Monthly, Offshore Wind Biz, Hydrogen Insight, Energy Storage News, Electrek, CleanTechnica, Canary Media, IEEFA, Carbon Brief, and more
- Fetches **Google News query feeds** across 50+ targeted search terms per sector
- All ingestion runs **concurrently** via `ThreadPoolExecutor` for fast scan times
- **Fuzzy deduplication** removes duplicate stories across sources using TF-IDF normalisation + SequenceMatcher similarity

### 🧠 Automatic Article Enrichment
Every article is auto-tagged with:

| Field | Method |
|---|---|
| **Country** | 500+ geographic keyword patterns including city names, states, regions |
| **Region** | Mapped from country (Americas, Europe, Asia-Pacific, MEA, etc.) |
| **Topic / Sector** | Rule-based keyword classifier (Hyperscale, AI/GPU, Power, Construction, Solar, Wind, Storage, etc.) |
| **Project Status / Deal Status** | Keyword detection (Proposed, Approved, Under Construction, Commissioned, Tendered, Contracted, Challenged, Financed) |
| **Capacity (MW/GW)** | Regex extraction from headline text |
| **Deal Size ($bn/$m)** | Multi-currency regex extraction |
| **Companies** | Named-entity matching against 200+ known industry companies |
| **ISO / RTO / Grid Operator** | Detected from 50+ global grid operators (PJM, ERCOT, CAISO, National Grid, TenneT, AEMO, etc.) |
| **Deal Type (RE)** | PPA, MOU, Tender/Auction, AOR/Offtake, Investment/IPO, M&A, Commissioning, Policy/Reg |

### 🧠 AI Summarize — Market Intelligence Briefing
Both modules generate a **Wood Mackenzie / BloombergNEF-grade structured briefing** using built-in TF-IDF NLP — no external API key required.

**Data Center Briefing sections:**
1. Executive Summary
2. Key Themes & Market Dynamics
3. Major Projects, Deals & Capacity Announcements
4. Regulatory & Permitting Landscape
5. Power & Infrastructure
6. Company Activity & Competitive Landscape
7. Regional Breakdown
8. Market Outlook & Forward Signals

**Renewables Briefing sections:**
1. Executive Summary
2. Sector Dynamics (Solar / Wind / Storage / Hydrogen / OSW)
3. Deal Flow & Transaction Analysis (PPA / Tender / MOU / AOR / Investment / M&A)
4. Major Projects, Deals & Capacity Announcements
5. Policy, Regulation & Procurement
6. Developer, Offtaker & Investor Activity
7. Regional Breakdown
8. Market Outlook & Forward Signals

Briefings are downloadable as **`.txt`**, **`.docx` (Word)**, and **`.pdf`** with professional formatting.

### 📊 Analytics & Visualisations
- **World choropleth map** — article volume heatmap by country (dark-themed, Plotly)
- **Sector / Topic distribution** bar charts
- **Regional distribution** bar charts
- **Country breakdown** bar charts (top 20)
- **Publication volume over time** — spline area chart
- **Sentiment / Project status** distribution
- **Topic share** donut chart
- **AI Signal Score distribution** (DC module)
- **Capacity pipeline heatmap** — MW/GW by country × topic (DC module)
- **Deal flow breakdown** by type
- **Trend comparison** — compare two time periods, countries, or companies side-by-side (DC module)

### 🏢 By Company
- Top 30 companies by article mention count (bar chart)
- Drill-down panel: select any company to view all articles mentioning it
- Full article cards with headline, date, source badge, country, topic, capacity, deal size, sentiment

### 📍 By State / Province
Both modules include a state/province drill-down:
- Auto-detects US states, Australian states, Indian states, Canadian provinces, and more from headline text
- Bar chart of top states by article volume
- All-states summary table
- Drill-down: select any state to view metrics (top topic, latest date, capacity mentions) and all associated article cards

### 💰 Deal Flow (Renewables)
Dedicated tab surfacing PPA, Tender/Auction, MOU, AOR/Offtake, Investment, and M&A events with KPI summary and per-deal-type article listings.

### 🤖 AI Signal Scoring (Data Centers)
Proprietary multi-factor scoring model ranks every article by market significance:
- **Sentiment weight** (Opened/Live = 10, Approved = 8, Under Construction = 6…)
- **Capacity size** (scaled MW/GW value up to 15 pts)
- **Deal value** (scaled $bn/$m value up to 12 pts)
- **Recency bonus** (articles within 7 days)
- **Hyperscaler involvement** (Microsoft, Google, Amazon, Meta, etc.)
- **Topic weight** (Hyperscale = 8, AI/GPU = 8, Investment = 7, Power = 6…)
- **High-value keyword bonuses** (billion, gigawatt, nuclear, Stargate, AI campus…)

High-signal articles (score ≥ 30) are surfaced with orange highlight strips and score badges.

### 📤 Export
- **Excel (.xlsx)** — multi-sheet workbook: All Articles · By Country · By Region · By Topic · By Company — colour-coded badges, auto-filter, frozen headers, clickable URLs
- **CSV** — flat export of the filtered view, ready for Excel / Python / PowerBI / Tableau
- **Word (.docx)** — formatted intelligence briefing with cover block, stats table, section headings, bullet points
- **PDF** — styled briefing with matching layout

---

## 🗂️ Platform Structure

```
app.py
│
├── INTELLIGENCE MODULES
│   ├── 🏢 Data Center Markets
│   │   ├── 📰 Feed
│   │   ├── 🗺️ World Map
│   │   ├── 📊 Analytics
│   │   ├── 🏢 By Company
│   │   ├── 📍 By State
│   │   ├── 🧠 AI Summarize          ← Market Intelligence Briefing
│   │   ├── 📈 Trend Compare
│   │   ├── 🔥 Capacity Heatmap
│   │   ├── 💰 Deal Flow
│   │   ├── 🤖 AI Scoring
│   │   └── ⬇️ Export
│   │
│   └── 🌱 Renewables Power Markets
│       ├── 📰 Feed
│       ├── 🗺️ World Map
│       ├── 📊 Analytics
│       ├── 🏢 By Company
│       ├── 📍 By State
│       ├── 💰 Deal Flow
│       ├── 🧠 AI Summarize          ← Renewables Intelligence Briefing
│       └── ⬇️ Export
│
├── SIDEBAR FILTERS
│   ├── Module selector (DC / Renewables)
│   ├── Time range (Latest / 30d / 14d / 7d / Custom)
│   ├── Scrape depth slider (1–100 pages)
│   ├── Geography: Region → Country → State/Province
│   └── Content filters (mode-specific)
│       ├── DC: News Type · ISO/RTO · Keyword · Company · Topic · Status · Min MW
│       └── RE: Sector · State · Deal Type · Project Status · Keyword
│
└── CORE ENGINE
    ├── run_all_scrapers()         — DC scraper (DCD direct + RSS)
    ├── run_re_scrapers()          — RE scraper (specialist RSS + Google News)
    ├── enrich() / enrich_re()     — enrichment pipeline
    ├── deduplicate()              — fuzzy dedup engine
    ├── generate_local_summary()   — DC NLP briefing generator
    ├── generate_re_summary()      — RE NLP briefing generator
    ├── build_briefing_docx()      — Word export
    ├── build_briefing_pdf()       — PDF export
    └── build_excel()              — Excel export
```

---

## 🚀 Getting Started

### Prerequisites

- Python 3.10 or higher
- Git

### Installation

```bash
# 1. Clone the repository
git clone https://github.com/YOUR_USERNAME/YOUR_REPO_NAME.git
cd YOUR_REPO_NAME

# 2. Create and activate a virtual environment
python -m venv venv
source venv/bin/activate        # macOS/Linux
venv\Scripts\activate           # Windows

# 3. Install dependencies
pip install -r requirements.txt
```

### Running Locally

```bash
streamlit run app.py
```

The app opens at `http://localhost:8501` in your browser.

---

## 📦 Dependencies

```txt
streamlit>=1.35.0
pandas>=2.0.0
plotly>=5.18.0
openpyxl>=3.1.0
beautifulsoup4>=4.12.0
cloudscraper>=1.2.71
python-docx>=1.1.0
reportlab>=4.1.0
requests>=2.31.0
```

> **Note:** `cloudscraper` is the preferred HTTP client for bypassing Cloudflare-protected pages. If it's unavailable, the app falls back to `requests`.

---

## ☁️ Deploying to Streamlit Cloud

1. Push your code to a public or private GitHub repository
2. Go to [share.streamlit.io](https://share.streamlit.io) and click **New app**
3. Select your repository, branch (`main`), and set the **Main file path** to `app.py`
4. Click **Deploy**

**`requirements.txt`** is automatically picked up by Streamlit Cloud. No additional configuration needed.

> ⚠️ **Cloudflare note:** Some news sources (particularly DataCenterDynamics) deploy Cloudflare protection that may block scraping on shared cloud infrastructure. If no DC articles are returned, try reducing scrape depth or re-running the scan. The RSS-based Renewables module is generally unaffected.

---

## 🔧 Configuration & Customisation

### Adding New RSS Feeds (Renewables)
Edit the `RE_FEED_REGISTRY` dictionary in `app.py`. Each entry is a list of feed dicts under a sector key:

```python
RE_FEED_REGISTRY = {
    "Solar": [
        {"url": "https://www.example.com/feed/", "source": "Example Solar News", "weight": 8},
        ...
    ],
    ...
}
```

### Adding New Companies
Append to the `KNOWN_COMPANIES` list for DC, or extend the `developers` / `offtakers` / `financiers` sets inside `generate_re_summary()` for RE.

### Adding New Countries / States
- Countries: add to `COUNTRY_TO_REGION` and `COUNTRY_KEYWORDS` dicts
- States/Provinces: add to `COUNTRY_STATES` dict

### Adjusting the AI Signal Scoring (DC)
Modify the `_ai_score()` function inside the AI Scoring tab section. Weight maps for topics, sentiments, and keyword bonuses are all clearly labelled.

---

## 📐 Architecture

```
Browser
  └── Streamlit Frontend (app.py)
        ├── Sidebar: filters + scan trigger
        ├── Session state: df_full / re_df_full (persists across reruns)
        └── Main area: tabbed intelligence views
              │
              ▼
        Scraper Layer
        ├── CloudScraper / requests → DCD HTML pages (DC)
        └── ThreadPoolExecutor → 100+ RSS/Atom feeds (RE)
              │
              ▼
        Enrichment Pipeline
        ├── Country / Region detection
        ├── Topic / Sector classification
        ├── Status / Deal Type detection
        ├── Capacity & Deal Size regex extraction
        ├── Company NER matching
        └── ISO/RTO grid operator detection
              │
              ▼
        Deduplication
        ├── URL-based exact dedup
        └── Fuzzy headline similarity (SequenceMatcher ≥ 0.88)
              │
              ▼
        Analytics & Export Layer
        ├── Plotly charts
        ├── TF-IDF NLP summariser
        ├── Word / PDF / Excel / CSV export
        └── Streamlit UI rendering
```

---

## 📊 Data Sources

### Data Center Module
| Source | Type | Coverage |
|---|---|---|
| DataCenterDynamics — Construction Channel | Direct HTML scrape | Global DC construction news |
| DataCenterDynamics — General News | Direct HTML scrape | Global DC industry news |

### Renewables Module
| Sector | Key Sources |
|---|---|
| Solar | PV Magazine, PV Tech, PV Magazine USA, Solar Power World, Google News queries |
| Wind | Wind Power Monthly, Wind Energy News, Windpower Engineering, Google News queries |
| Offshore Wind | Offshore Wind Biz, 4C Offshore, Google News queries |
| Energy Storage | Energy Storage News, Storage Daily, Google News queries |
| Hydrogen | Hydrogen Insight, H2 View, Fuel Cells Works, Hydrogen Fuel News, Google News queries |
| Other Renewables | Recharge News, Enerdata, Google News queries (geothermal, hydro, tidal, biomass, CSP) |
| Cross-sector | CleanTechnica, Electrek, Canary Media, IEEFA, Carbon Brief, GreenBiz |

---

## 🛡️ Legal & Usage Notes

- All data is sourced from **publicly available** RSS feeds and news websites
- This platform aggregates **headlines and metadata only** — it does not reproduce full article text
- Intended for **internal research, market monitoring, and business intelligence** purposes
- Always verify critical decisions against primary source publications
- The AI briefings are generated from headline-level signals and should be treated as **intelligence summaries**, not definitive financial or investment advice

---

## 🗺️ Roadmap

- [ ] Email / Slack alert system for high-signal articles (score ≥ 40)
- [ ] Persistent article history with cross-session trend tracking
- [ ] Company-level profile pages with historical activity timeline
- [ ] PPA price index tracking
- [ ] Nuclear & Gas power markets module
- [ ] Carbon markets & voluntary offset module
- [ ] User-configurable watchlists (companies, countries, keywords)
- [ ] Scheduled auto-scan with Streamlit background jobs

---

## 👤 Author

Built by **Sharugh A**

*Global Energy Intelligence Platform — Data Centers · Renewables Power Markets*

---

## 📄 License

© Sharugh A. All rights reserved.

Unauthorised redistribution, resale, or removal of authorship metadata is prohibited under applicable intellectual property law. This software is provided for authorised use only.
