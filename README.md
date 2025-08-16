
# World Analogs — Streamlit App (Legacy-compatible, works with `.xls` on Pandas 2.x)

This build supports **`.xls`** by bypassing Pandas' xlrd dependency and using **pyexcel-xls** to read legacy Excel files.  
It also supports **`.xlsx`** via `openpyxl`.

## Why this build?
- On modern environments, Pandas 2.x and xlrd conflict for `.xls` (xlrd>=2.0.0 dropped `.xls` support).
- We avoid that entirely by using `pyexcel-xls` for `.xls` and `openpyxl` for `.xlsx`.

## Deploy on Streamlit Community Cloud
1. Push these files to a public GitHub repo.
2. On https://share.streamlit.io, create a new app with `app.py` as the entry point.
3. Upload your workbook (`.xls` or `.xlsx`) in the app.

## Run locally
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Tech notes
- `.xls` → `pyexcel-xls`
- `.xlsx` → `openpyxl` via `pandas.ExcelFile`
