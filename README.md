# M365 Unified Audit Log Viewer

Interactive Streamlit app for browsing and analyzing Microsoft 365 Unified Audit Log exports.

## Features

- **CSV Upload** — Select any UAL export CSV with an `AuditData` JSON column
- **Sidebar Filters** — Filter by date range, operation, user, workload, or free-text search
- **Summary Dashboard** — Metric cards and bar charts for quick insights
- **Timeline View** — Scatter or daily-aggregated chart of events over time
- **Filterable Log Table** — Sortable table of all entries
- **Entry Detail** — Structured breakdown of Actor, Target, Modified Properties, Extended Properties, and full JSON

## Quick Start (Local)

```bash
pip install -r requirements.txt
streamlit run app.py
```

Then open http://localhost:8501 and upload a CSV.

## Docker

```bash
docker build -t ual-viewer .
docker run -p 8501:8501 ual-viewer
```

## Streamlit Community Cloud

1. Push this repo to GitHub
2. Go to [share.streamlit.io](https://share.streamlit.io)
3. Point it at this repo with `app.py` as the main file

## CSV Format

The app expects the standard Unified Audit Log export format with these columns:

| Column | Description |
|--------|-------------|
| RecordId | Unique record identifier |
| CreationDate | Timestamp of the event |
| RecordType | Numeric record type |
| Operation | Action performed |
| UserId | User or service principal |
| AuditData | JSON blob with full event details |
