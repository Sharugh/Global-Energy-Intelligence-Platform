# 🌐 Global Data Center Construction Intelligence

A real-time Streamlit intelligence platform that scrapes, enriches, and visualises global data center industry news — covering construction projects, site selections, approvals, expansions, investments, and power infrastructure across 45+ countries.

---

## Features

### 🕸️ Smart Scraping Engine
- **Primary source:** DataCenterDynamics (DCD) — scraped via DCD's own Construction Channel (`?term=the-data-center-construction-channel`), the highest-density feed of DC construction news on the web
- **Dynamic region targeting:** When you select regions in the sidebar, the scraper appends DCD's own region taxonomy terms (`?term=europe`, `?term=asia-pacific`, etc.) so only regionally relevant articles are fetched — faster and more accurate than post-hoc filtering
- **Project-stage coverage:** Automatically scrapes DCD's tagged pages for `approved`, `site-selection`, `disclosed-projects`, `project-announcement`, `expansion`, and `extension` — so the full project lifecycle is covered
- **Google News supplement:** 30 targeted queries covering construction, approvals, hyperscalers, operators, power, investment, and all global regions — runs in parallel after DCD scrape completes
- **Cloudflare bypass:** Uses `cloudscraper` when available, falls back to `requests` with browser-like headers

### 🌍 Global Coverage
- 45+ countries across all continents with auto-detection using 500+ geographic keywords, city names, and state/province names
- Regions: North America · Europe · Asia Pacific · Middle East · Africa · Latin America
- Country → Region → State/Province drill-down hierarchy in both filters and tabs

### 🧠 Auto-Enrichment (every article)
| Field | What it detects |
|---|---|
| **Country** | 45+ countries via keyword + city matching |
| **Region** | Auto-mapped from country |
| **Topic** | Hyperscale · Colocation · AI/GPU · Power · Investment · Permits · Construction · Sustainability |
| **Sentiment** | Opened/Live · Approved · Proposed · Under Construction · Challenged · News |
| **Capacity** | MW / GW figures extracted from headline |
| **Deal Size** | $bn / $m figures extracted from headline |
| **Companies** | Up to 4 named companies matched against 80+ known DC operators/vendors |

### 🔁 Deduplication
- URL-based exact dedup first
- Fuzzy headline matching (88% similarity threshold via `difflib.SequenceMatcher`)
- When duplicates found across sources, DCD version is always kept (priority 1)

---

## Tabs

| Tab | Description |
|---|---|
| 📰 **Feed** | Chronological article cards with headline, date, source, topic badge, sentiment, capacity, deal size, and direct link |
| 🗺️ **World Map** | Red-heat choropleth map — countries glow brighter red the more articles they have |
| 📊 **Analytics** | Topic breakdown · Top countries · Publication volume over time · Topic share donut · Capacity mentions table · Source bar |
| 🏢 **By Company** | Top 30 companies by mention count (bar chart) + drill-down to read all articles mentioning a selected company |
| 📍 **By State** | State/province article volume bar chart + drill-down: select any state to see its topic breakdown mini-chart and all matching articles |
| 🧠 **Market Intel** | Built-in TF-IDF NLP briefing — no API key needed. Covers themes, capacity pipeline, regulatory activity, company signals, and forward-looking indicators. Export as `.txt`, `.docx`, or `.pdf` |
| ⬇️ **Export** | Excel (5 sheets: Articles · By Country · By Region · By Topic · By Company) + CSV |

---

## Sidebar Filters

Filters apply **after** the scan — no re-scrape needed:

- **Date range** — Latest (all), 7d, 14d, 30d, 90d, or custom from/to dates
- **🌐 Region** — multiselect; also dynamically narrows the DCD scrape on next run
- **🌍 Country** — filtered by selected regions
- **📍 State / Province** — filtered by selected countries; draws from per-country state lists
- **🏢 Company** — search across 80+ known DC companies + any detected in results
- **🏷️ Topic** — Hyperscale, Colocation, AI/GPU, Power, Investment, Permits, Construction, Sustainability
- **📊 Project Status** — Opened/Live, Approved, Proposed, Under Construction, Challenged, News
- **🔤 Keyword** — free-text search across headlines
- **⚡ Min Capacity (MW)** — filter to only articles with a MW/GW figure above a threshold

---

## Installation

```bash
pip install streamlit cloudscraper beautifulsoup4 feedparser pandas \
            plotly openpyxl python-docx reportlab
```

> `cloudscraper` and `python-docx` and `reportlab` are optional — the app degrades gracefully if they are absent.

### `requirements.txt`
```
streamlit
cloudscraper
beautifulsoup4
feedparser
pandas
plotly
openpyxl
python-docx
reportlab
```

---

## Deploying to Streamlit Cloud

1. Push `code.py` (rename if needed) and `requirements.txt` to a GitHub repository
2. Go to [share.streamlit.io](https://share.streamlit.io) → **New app**
3. Select your repo, branch, and set **Main file path** to `code.py`
4. Click **Deploy** — no secrets or environment variables required

---

## How the Scraper Works

```
Run Global Scan
    │
    ├── DCD Construction Channel (primary)
    │     ├── ?term=the-data-center-construction-channel
    │     ├── + ?term=<region> for each selected region (e.g. europe)
    │     ├── + ?term=approved  (project stage pages)
    │     ├── + ?term=site-selection
    │     ├── + ?term=disclosed-projects
    │     ├── + ?term=project-announcement
    │     ├── + ?term=expansion
    │     └── + ?term=extension
    │     Each paginated up to max_pages (default 10 × ~28 articles/page)
    │
    ├── Google News RSS (30 targeted queries, parallel)
    │     ├── Construction / groundbreaking / phase / expansion
    │     ├── Approvals / permits / site selection / disclosed
    │     ├── Hyperscalers (Microsoft, Google, Amazon, Meta, Oracle…)
    │     ├── Operators (Equinix, Digital Realty, CyrusOne, QTS…)
    │     ├── Power (nuclear, SMR, PPA, grid connection, MW/GW)
    │     ├── Investment (acquisition, REIT, IPO, financing)
    │     └── Regions (Europe, APAC, Middle East, LatAm, Africa)
    │
    └── Merge → DC relevance filter → Date cutoff filter
              → Enrich (country/topic/sentiment/capacity/deal/companies)
              → Fuzzy deduplicate
              → Display
```

---

## Data Sources

| Source | Type | Notes |
|---|---|---|
| [DataCenterDynamics](https://www.datacenterdynamics.com) | HTML scrape | Primary — Construction Channel + stage tags |
| [Google News RSS](https://news.google.com) | RSS/feedparser | 30 targeted DC queries, runs in parallel |

---

## Notes

- No API keys required for any functionality
- The app stores all data in Streamlit session state — refreshing the page clears results
- Scrape depth is fixed at 10 pages per DCD term (≈280 articles per term, ≈2,240 total from DCD alone before dedup)
- Google News typically adds 300–600 additional unique articles depending on query overlap
- The Market Intel briefing uses pure Python TF-IDF + rule-based extraction — no LLM, no external API

---

## License

MIT
