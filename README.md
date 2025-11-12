
# Upload-Based Sales Dashboard (No Monthly, No Folders)

Client workflow:
1) Open the web app.
2) Upload two Excel files (Sales + Weight).
3) Select filters (Location + Order Date range).
4) Download filtered Excel and a PDF report.

Runs on Streamlit Cloud (no install) or locally (optional).

## Streamlit Cloud (recommended)
- Create a free account at https://streamlit.io/cloud
- Create a new app and upload these files.
- Set main file to `app.py`. Use `requirements.txt` provided.

## Local run (optional)
```
python -m venv .venv && . .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

## Expected columns
Sales: Date (or Order Date), Quantity, Location, Product, Size
Weight: Product, Weight of Indv. Product (lb)
