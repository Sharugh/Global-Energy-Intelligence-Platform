# 🌐 Global Data Center Intelligence Platform

A Streamlit-based market intelligence dashboard that scrapes, enriches, and analyses data center news in real time — built for the Wood Mackenzie team to track global DC activity across deals, capacity, construction, and investment.

---

## What it does

The app pulls live articles from Data Center Dynamics, classifies each one by topic, sentiment, region, and company, then surfaces the results across a set of analytical views. Everything runs in the browser — no database, no backend service required.

---

## Features

**Live scraper** — fetches the latest articles from DCD's news and construction channels with Cloudflare bypass via `cloudscraper`.

**Auto-enrichment** — each article is automatically tagged with:
- Topic (Hyperscale, AI/GPU, Investment, Power, Colocation, Construction, Sustainability, etc.)
- Sentiment (Bullish, Bearish, Neutral, News)
- Region and ISO/RTO zone
- Capacity (MW) and Deal Size extracted from headline text
- Company names matched against a known operator/hyperscaler list

**AI Scoring engine** — proprietary scoring model that ranks articles by signal strength across recency, capacity, deal size, company tier, topic weight, and high-value keywords. Articles are tiered into High (30+), Medium (15–29), and Low (<15) signal bands.

**11 analysis tabs:**

| Tab | Description |
|---|---|
| 📰 Feed | Chronological article cards with all enrichment tags |
| 🗺️ World Map | Choropleth map of article volume by country |
| 📊 Analytics | KPIs, topic breakdown, sentiment distribution, top companies |
| 🏢 By Company | Per-operator deep dive with article timelines |
| 📍 By State | US state-level breakdown with ISO/RTO mapping |
| 🧠 Market Intel | Auto-generated Wood Mackenzie-style prose briefing |
| 📈 Trend Compare | Multi-topic trend lines over time |
| 🔥 Capacity Heatmap | MW capacity heatmap by country and topic |
| 💰 Deal Flow | Deal size tracker and financial signal analysis |
| 🤖 AI Scoring | Full scored article table with distribution chart |
| ⬇️ Export | Download filtered data as Excel (.xlsx) or CSV |

**Export options:**
- Excel report with 5 sheets (All Articles, By Country, By Region, By Topic, By Company), colour-coded badges, auto-filter, frozen headers
- CSV flat export for use in Python, PowerBI, or Tableau
- PDF and Word (.docx) report generation

**Sidebar filters** — filter live data by date range, region, country, topic, sentiment, company, keyword, and capacity range.

---

## Tech stack

- [Streamlit](https://streamlit.io/) — UI framework
- [cloudscraper](https://github.com/VeNoMouS/cloudscraper) + [BeautifulSoup](https://www.crummy.com/software/BeautifulSoup/) — scraping and HTML parsing
- [pandas](https://pandas.pydata.org/) — data manipulation
- [Plotly](https://plotly.com/python/) — interactive charts and maps
- [openpyxl](https://openpyxl.readthedocs.io/) — Excel report generation
- [reportlab](https://www.reportlab.com/) — PDF export
- [python-docx](https://python-docx.readthedocs.io/) — Word document export

---

## Getting started

**1. Clone the repo**
```bash
git clone https://github.com/your-username/global-dc-intelligence.git
cd global-dc-intelligence
```

**2. Install dependencies**
```bash
pip install -r requirements.txt
```

**3. Run the app**
```bash
streamlit run app.py
```

The app opens at `http://localhost:8501`.

---

## Requirements

```
streamlit
pandas
plotly
openpyxl
beautifulsoup4
cloudscraper
requests
reportlab
python-docx
```

---

## Deployment

The app is deployed on [Streamlit Community Cloud](https://streamlit.io/cloud). To deploy your own instance:

1. Push the repo to GitHub
2. Go to [share.streamlit.io](https://share.streamlit.io) and connect the repo
3. Set the main file path to `app.py`
4. Deploy

No environment variables or secrets are required — the scraper uses public endpoints only.

---

## Notes

- The scraper targets public DCD news pages. Run times vary depending on the date range and number of pages fetched.
- The Market Intel tab generates a prose briefing using the enriched article data — useful for quick team updates or client-facing summaries.
- All filtering is done client-side on the fetched dataset; re-running the scan refreshes the data.

---

*Global Data Center Intelligence · Wood Mac*
