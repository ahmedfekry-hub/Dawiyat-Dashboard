# Dawiyat Project Intelligence Dashboard

Files to upload to GitHub:
- app.py
- requirements.txt
- README.md
- Dawiyat Master Sheet.xlsx

Notes:
- Keep Streamlit Python version on 3.12.
- Region is normalized in the dashboard to Western / Southern / Eastern / Northern.
- City and District are taken from the District sheet first, then from Link Code fallback.
- Penalties analysis separates total deviation count from total deducted penalty amount.
- Use the sidebar uploader whenever a refreshed workbook is available.
