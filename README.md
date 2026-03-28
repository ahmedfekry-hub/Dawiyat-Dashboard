# Dawiyat Project Intelligence Dashboard

GitHub / Streamlit upload files:
- app.py
- requirements.txt
- README.md
- Dawiyat Master Sheet.xlsx

## Run locally
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Notes
- Upload the latest updated Excel workbook from the sidebar to refresh the dashboard.
- The app uses these sheets automatically:
  - `Dawaiyat Service Tool`
  - `District`
  - `Penalties`
- City and District are mapped from the `District` sheet using `Link Code`.
- If a Link Code is missing from the `District` sheet, the app will show `Unknown` in the Data Quality page.
