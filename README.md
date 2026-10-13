# Interactive Data Analytics Dashboard (Flask + Chart.js)

Covers all 7 modules from the project proposal: Authentication, Dataset
Management, Preprocessing, Analysis, Visualization, Dashboard Management,
Report Generation.

## Setup

```bash
python -m venv venv
source venv/bin/activate        # Windows: venv\Scripts\activate
pip install -r requirements.txt
python app.py
```

Open `http://127.0.0.1:5000`. **Demo login:** `admin` / `admin123`

## Fixes applied while testing against the real Kaggle Superstore CSV

1. **CSV encoding** — the Kaggle file isn't UTF-8; `load_dataframe()` now
   tries `utf-8` → `cp1252` → `latin1` in order.
2. **Charts render blank** — canvases had no fixed-height container, so
   Chart.js (with `responsive: true`) couldn't compute a size. Every canvas
   is now wrapped in a `.chart-wrap` div with an explicit height, and all
   charts set `maintainAspectRatio: false`.
3. **No internet on some networks** — Bootstrap and Chart.js are bundled
   locally under `static/css` and `static/js` instead of loaded from a CDN,
   so the dashboard works even if CDN access is blocked.
4. **Line chart said "No date column available"** — on pandas ≥ 2.x with
   the string-dtype backend (default in pandas 3.0), text columns get a
   `str` dtype instead of the classic `object` dtype. The old check
   `df[col].dtype == object` silently failed to match, so `Order Date` /
   `Ship Date` never got parsed to datetime. Fixed by using
   `pd.api.types.is_numeric_dtype` / `is_datetime64_any_dtype` instead of
   direct dtype comparisons, which work correctly regardless of pandas
   version or string backend.
5. **Row ID / Postal Code showing up as KPIs and in the stats table** —
   these are identifiers, not metrics. `numeric_columns()` now excludes any
   column whose name contains a token like `id`, `code`, `zip`, `no`,
   `number` (case-insensitive, checked as whole words) so summary stats and
   KPI cards only show meaningful numeric fields (Sales, Profit, Quantity,
   Discount, ...).
6. **Default chart columns were high-cardinality IDs** — bar/pie charts
   defaulted to whatever categorical column came first (often `Order ID`,
   9,994 unique values). `pick_default_categorical()` now prefers a column
   with 2–20 unique values (e.g. Category, Region, Segment) for a
   meaningful chart out of the box.

## Project structure

```
dashboard_project/
├── app.py
├── requirements.txt
├── templates/
│   ├── base.html          # local Bootstrap/Chart.js, no CDN dependency
│   ├── login.html
│   ├── upload.html
│   └── dashboard.html
├── static/
│   ├── css/{bootstrap.min.css, style.css}
│   └── js/{bootstrap.bundle.min.js, chart.umd.min.js}
└── uploads/
```

## Troubleshooting

If a chart still doesn't render, it will now show a red error message
inline instead of staying blank — check that message, and check the
browser console (F12) and the Flask terminal for anything unexpected."# Mini_project" 
"# Mini_project" 
"# Mini_project" 
"# Mini_project" 
