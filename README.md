# VisaBulletineDashboard

Static visa bulletin dashboard with filters and charts (ECharts via CDN) using local data in `visa_bulletins.all.json`.

## Run

This dashboard loads JSON via `fetch`, so use a local server:

```
python -m http.server
```

Then open `http://localhost:8000` in a browser.
