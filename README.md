# Outlook Calendar Auditor

A Streamlit app to audit Outlook calendar events for issues like:
- After-hours meetings
- Missing location
- Overlaps (to be extended)

## Run locally

```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```

## Deploy to Streamlit Cloud
1. Push this repo to GitHub.
2. Go to https://share.streamlit.io and link your repo.
3. Deploy with `streamlit_app.py` as the entry point.
