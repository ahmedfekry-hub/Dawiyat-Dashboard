# Dawiyat Project Intelligence Dashboard

This package contains a Streamlit dashboard built for the uploaded **Dawiyat Master Sheet.xlsx**.

## Files
- `app.py` — main dashboard application
- `requirements.txt` — Python dependencies
- `Dawiyat Master Sheet.xlsx` — sample data file used as default source

## Run locally
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Refresh with a new daily file
Open the app, then use the **Upload refreshed Dawiyat workbook** control in the sidebar.
The dashboard will recalculate automatically.

## Main features
- Dark and light theme
- Executive KPI cards
- Region / City / District / Stage / Year / Status / Type / Class / Project / Subclass filters
- Planned vs actual progress tracking
- Approximate SPI / CPI based on available sheet fields
- Penalties and forecast delay exposure
- Work-order detail table for management review
