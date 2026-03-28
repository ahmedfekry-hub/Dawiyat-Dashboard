# Dawiyat Executive PMO Dashboard

This package contains the refreshed Streamlit dashboard for the Dawiyat project.

## Files to upload to GitHub
- `app.py`
- `requirements.txt`
- `README.md`
- `Dawiyat Master Sheet.xlsx`

Do not upload `__pycache__`.

## Main updates in this version
- Returned to the original professional executive dashboard style
- Removed CPI and SPI from all KPI cards, charts, and detail tables
- Uses the latest updated workbook and penalties sheet
- Fixed city and district mapping using the District sheet and link-code fallback
- Added Link Code as a cascading filter
- Preserved dark and light themes
- Preserved daily refresh through workbook upload in the sidebar

## Run locally
```bash
pip install -r requirements.txt
streamlit run app.py
```
