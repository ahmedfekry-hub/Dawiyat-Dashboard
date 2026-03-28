# Dawiyat Executive PMO Dashboard

Executive Streamlit dashboard for Dawiyat project monitoring using the `Dawaiyat Service Tool`, `District`, and `Penalties` sheets.

## Files to upload to GitHub
Upload these files only:
- `app.py`
- `requirements.txt`
- `README.md`
- `Dawiyat Master Sheet.xlsx` (recommended as the default demo dataset)

Do **not** upload `__pycache__`.

## Main features
- Executive dark and light themes
- Cascading filters: Region > City > District and remaining categories adapt automatically
- Safe district cleaning without moving records across cities; city-to-district relationships remain exactly as provided in the Dawiyat workbook
- Planned vs actual progress analysis
- Civil and fiber completion KPIs
- City and district performance views
- Penalties and quality exposure analysis
- Approximate SPI / CPI / EAC for executive monitoring
- Daily refresh by uploading a new workbook from the sidebar

## Run locally
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Notes
- Keep the same workbook sheet names and general column structure for best results.
- SPI, CPI, and EAC are management approximations based on available workbook fields.
- Link Code is included as a dashboard filter, and all filters cascade from the actual row relationships in the workbook.
