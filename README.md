# DCD US Construction News Scraper

Streamlit app that scrapes the [Data Center Dynamics Construction Channel](https://www.datacenterdynamics.com/en/news/?term=the-data-center-construction-channel) and exports US-state-filtered articles to a formatted Excel file.

## Features
- Filter by **Latest**, **Past 30 days**, or **Past 10 days**
- Pre-filters by North America region tag (optional)
- Post-filters by 50 US states + abbreviations + ~80 major US data center cities/counties
- Exports a styled `.xlsx` with clickable hyperlinks, auto-filter, frozen header
- Uses `cloudscraper` to bypass Cloudflare protection

## Deploy on Streamlit Community Cloud

1. Push this repo to GitHub
2. Go to [share.streamlit.io](https://share.streamlit.io) → **New app**
3. Select your repo, branch `main`, file `app.py`
4. Click **Deploy**

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Why cloudscraper?

DCD uses Cloudflare, which returns a 403 to plain `requests` calls on some hosts.
`cloudscraper` mimics a real browser handshake and solves JS challenges automatically.
If it's not installed the app falls back to `requests` with browser-like headers.

## Excel output columns

| Column   | Description                       |
|----------|-----------------------------------|
| #        | Row number                        |
| Headline | Article title                     |
| Date     | Publication date (YYYY-MM-DD)     |
| URL      | Direct link (clickable in Excel)  |
