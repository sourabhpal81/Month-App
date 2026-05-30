# 🎯 Matka Tracker

A web app for tracking daily jodis across 31 Matka markets, with live dashboard, Markov predictions, monthly reports, weekday patterns, and Excel/PDF export.

## ✨ Features

- 📊 **Live dashboard** — Sum 10-14 + Digit 0-9 distribution charts, streaks, anomalies, recent activity
- 📝 **Quick data entry** — single, bulk paste, CSV upload, photo OCR
- 🔮 **Markov chain predictions** — weekday-aware, top-10 suggested jodis with score
- 🧪 **Backtest** — walk-forward accuracy on last 30 days
- 🧩 **Pattern discovery** — weekday heatmap, cycle detection, market correlations
- 📑 **Reports** — Excel and PDF export, any date range, per-market history
- 📈 **Calendar view** per market, color-coded by sum group
- 🔄 **Auto-fetch** — scrape daily results from source site automatically
- 👥 **Multi-user login** with role-based access
- 💾 **Backup / restore** the SQLite database
- 📱 **PWA** — install on Android, iOS, Windows like a real app

## 🚀 Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

Opens at `http://localhost:8501`. Login: `admin` / `admin123` (change in Settings after first login).

## ☁️ Deploy free

This repo is ready for [Streamlit Community Cloud](https://share.streamlit.io):
1. Fork this repo
2. Connect at share.streamlit.io
3. Click "Deploy"
4. Public URL ready in ~2 minutes

## 🧮 The "Sum 10-14" algorithm

For any 2-digit jodi `AB`:
```
Sum Group = (A + B) mod 5 + 10
```
Always returns 10, 11, 12, 13, or 14.

Examples:
- Jodi `45` → 4+5=9, 9%5=4, +10 = **Sum 14**
- Jodi `91` → 9+1=10, 10%5=0, +10 = **Sum 10**

Color convention throughout app: 10=blue · 11=green · 12=yellow · 13=orange · 14=red.

## 📦 Project structure

```
matka-tracker/
├── app.py              ← main Streamlit app
├── scraper.py          ← daily auto-fetch from source site
├── matka.db            ← SQLite DB, pre-loaded with Jan-May 2026 data
├── requirements.txt    ← Python dependencies
├── .streamlit/
│   └── config.toml     ← production config
└── static/
    ├── manifest.json   ← PWA manifest
    ├── icon-192.png
    └── icon-512.png
```

## 🛠️ Tech

Python · Streamlit · SQLite · openpyxl · plotly · reportlab · Pillow

## ⚖️ License

Personal use only. Predictions are statistical analysis, not gambling advice.
