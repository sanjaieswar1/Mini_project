"""
Interactive Data Analytics Dashboard
Python Flask + Chart.js

Modules implemented (per project spec):
  1. User Authentication
  2. Dataset Management (upload CSV/Excel)
  3. Data Preprocessing (missing values, duplicates, type conversion)
  4. Data Analysis (statistical summaries, patterns)
  5. Data Visualization (bar, line, pie, scatter, histogram via Chart.js)
  6. Dashboard Management (KPI cards, filters)
  7. Report Generation (export filtered/analyzed data as CSV)
"""

import os
import io
import re
import json
import math
import uuid
import warnings
from functools import wraps
from urllib.parse import urlencode

import pandas as pd
import numpy as np
from flask import (
    Flask, render_template, request, redirect, url_for,
    session, flash, jsonify, send_file
)
from werkzeug.utils import secure_filename
from werkzeug.security import generate_password_hash, check_password_hash

# --------------------------------------------------------------------------
# App configuration
# --------------------------------------------------------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")
ALLOWED_EXTENSIONS = {"csv", "xlsx", "xls"}

app = Flask(__name__)
app.config["SECRET_KEY"] = os.environ.get("SECRET_KEY", "dev-secret-change-me")
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024  # 25 MB upload limit

os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# --------------------------------------------------------------------------
# Module 1: User Authentication
# --------------------------------------------------------------------------
# The single admin account stays hardcoded and separate from self-registered
# users - "admin" is a reserved login id that signup can never create.
ADMIN_USERNAME = "admin"
ADMIN_PASSWORD_HASH = generate_password_hash("admin123")

# Self-registered users (students/staff) are stored keyed by lowercase
# email, and persisted to users.json so accounts survive an app restart.
# This is still a simple file-backed store for demo/academic purposes -
# replace with a real database (SQLite/MySQL) for production use.
USERS_FILE = os.path.join(BASE_DIR, "users.json")


def load_users():
    if os.path.exists(USERS_FILE):
        try:
            with open(USERS_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except (json.JSONDecodeError, OSError):
            pass
    return {}


def save_users(users):
    with open(USERS_FILE, "w", encoding="utf-8") as f:
        json.dump(users, f, indent=2)


USERS = load_users()


def login_required(view):
    @wraps(view)
    def wrapped(*args, **kwargs):
        if "username" not in session:
            flash("Please log in to continue.", "warning")
            return redirect(url_for("login"))
        return view(*args, **kwargs)
    return wrapped


@app.route("/", methods=["GET"])
def index():
    if "username" in session:
        return redirect(url_for("dashboard"))
    return redirect(url_for("login"))


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        identifier = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        key = identifier.lower()

        if key == ADMIN_USERNAME:
            if check_password_hash(ADMIN_PASSWORD_HASH, password):
                session["username"] = ADMIN_USERNAME
                session["display_name"] = "Admin"
                session["role"] = "admin"
                flash("Welcome back, Admin!", "success")
                return redirect(url_for("dashboard"))
        else:
            user = USERS.get(key)
            if user and check_password_hash(user["password_hash"], password):
                session["username"] = key
                session["display_name"] = user["name"]
                session["role"] = "user"
                flash(f"Welcome back, {user['name']}!", "success")
                return redirect(url_for("dashboard"))

        flash("Invalid username/email or password.", "danger")
    return render_template("login.html")


@app.route("/register", methods=["GET", "POST"])
def register():
    """Self-service signup for regular (non-admin) users.

    Collects name, designation, college, email and password, and gives
    the new account the exact same dashboard/upload experience as admin -
    there is no separate "user view"; login_required is the only gate,
    so any logged-in account (admin or self-registered) sees the same
    dashboard, upload, chart and report pages.
    """
    if request.method == "POST":
        name = request.form.get("name", "").strip()
        designation = request.form.get("designation", "").strip()
        college = request.form.get("college", "").strip()
        email = request.form.get("email", "").strip().lower()
        password = request.form.get("password", "")
        confirm_password = request.form.get("confirm_password", "")

        errors = []
        if not all([name, designation, college, email, password]):
            errors.append("Please fill in all fields.")
        if email and "@" not in email:
            errors.append("Enter a valid email address.")
        if password and len(password) < 6:
            errors.append("Password must be at least 6 characters.")
        if password != confirm_password:
            errors.append("Passwords do not match.")
        if email == ADMIN_USERNAME:
            errors.append("That email is reserved.")
        if email in USERS:
            errors.append("An account with this email already exists.")

        if errors:
            for err in errors:
                flash(err, "danger")
            return render_template("register.html", form=request.form)

        USERS[email] = {
            "name": name,
            "designation": designation,
            "college": college,
            "password_hash": generate_password_hash(password),
        }
        save_users(USERS)
        flash("Account created successfully. Please log in.", "success")
        return redirect(url_for("login"))

    return render_template("register.html", form={})


@app.route("/logout")
def logout():
    session.clear()
    flash("Logged out successfully.", "info")
    return redirect(url_for("login"))


# --------------------------------------------------------------------------
# Module 2 & 3: Dataset Management + Preprocessing
# --------------------------------------------------------------------------
def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def load_dataframe(filepath):
    """Load CSV or Excel into a pandas DataFrame.

    - Tries several encodings (Kaggle CSVs are often cp1252 / latin1, and
      Excel-exported CSVs carry a BOM that would otherwise glue "\ufeff" onto
      the first column name).
    - If a CSV loads as a single column, retries with ';', tab and '|' so
      semicolon / tab / pipe separated files also work.
    """
    ext = filepath.rsplit(".", 1)[1].lower()
    if ext != "csv":
        return pd.read_excel(filepath)

    last_error = None
    for enc in ("utf-8-sig", "cp1252", "latin1"):
        try:
            df = pd.read_csv(filepath, encoding=enc)
            if df.shape[1] == 1:
                # One column usually means a ';' / tab / '|' separated file.
                # Only real delimiters are tried (never a guessed letter).
                for sep in (";", "\t", "|"):
                    try:
                        alt = pd.read_csv(filepath, encoding=enc, sep=sep)
                    except Exception:
                        continue
                    if alt.shape[1] > 1:
                        df = alt
                        break
            return df
        except UnicodeDecodeError as exc:
            last_error = exc
    raise last_error


def _tokens(name):
    """'Row ID' / 'row_id' / 'RowID' / 'Sub-Category' -> lowercase word list."""
    name = re.sub(r"([a-z0-9])([A-Z])", r"\1 \2", str(name))
    return [t for t in re.split(r"[^A-Za-z0-9]+", name.lower()) if t]


# Column-name words that mark a column as an IDENTIFIER (Row ID, Postal Code,
# Customer ID, Phone, SKU ...) rather than a metric. Summing/averaging or
# charting an identifier is meaningless, so these are never picked for charts.
STRONG_ID_TOKENS = {"id", "uuid", "guid", "sku", "code", "zip", "zipcode", "postal",
                    "pincode", "pin", "phone", "mobile", "contact", "unnamed"}
# "no / num / number" only mean identifier when they END the name
# ("Invoice No", "Roll Number") - not in "Number of Orders" or "No. of Items".
WEAK_ID_TOKENS = {"no", "num", "number"}


def is_identifier_column(col_name):
    toks = _tokens(col_name)
    if not toks:
        return False
    return any(t in STRONG_ID_TOKENS for t in toks) or toks[-1] in WEAK_ID_TOKENS


def coerce_numeric_columns(df, min_success_ratio=0.8):
    """
    Module 3 (generalized fix): Numeric-format cleanup.

    Real-world CSVs (Kaggle scrapes, e-commerce/finance exports, etc.)
    routinely store genuinely numeric fields as TEXT because of:
      - currency symbols       ("₹399", "$1,099")
      - thousands separators   ("24,269")
      - percent signs          ("64%")
      - a handful of stray/garbage tokens in an otherwise numeric column
        (e.g. a single rogue "|" in a ratings column)

    pd.read_csv only infers a numeric dtype when EVERY value in a column
    parses cleanly as a number. One bad cell, or a symbol present on every
    row, forces the WHOLE column to load as object/string dtype. Since
    numeric_columns() (and therefore the KPI cards, bar/scatter/histogram
    charts, and summary stats) only look at columns pandas already
    recognizes as numeric, a column that is semantically numeric but
    textually "dirty" becomes invisible to the entire visualization layer.
    This isn't specific to one dataset - it affects any dataset with this
    kind of formatting.

    Fix: for every non-numeric column, strip common non-numeric
    formatting and try to parse it as a number. If a high-enough fraction
    of non-empty cells parse successfully, commit the conversion
    (unparseable cells become NaN, handled by the missing-value step
    below). Columns that are mostly text (IDs, names, free-form
    descriptions) are left untouched, so this won't misclassify a
    column like "product_id" as numeric.
    """
    # Only strip known numeric-formatting symbols (currency signs, thousands
    # commas, percent signs, whitespace). Anything else left over (letters,
    # unexpected punctuation) is intentionally NOT stripped, so genuinely
    # alphanumeric columns (product IDs like "B07JW9H4J1", user IDs, etc.)
    # correctly fail to parse instead of being forced into a number by
    # deleting their letters.
    currency_symbols = re.compile(r"[₹$€£¥]")

    def clean_series(raw):
        cleaned = raw.str.replace(currency_symbols, "", regex=True)
        cleaned = cleaned.str.replace(",", "", regex=False)
        cleaned = cleaned.str.replace("%", "", regex=False)
        return cleaned.str.strip()

    for col in df.columns:
        if pd.api.types.is_numeric_dtype(df[col]) or pd.api.types.is_datetime64_any_dtype(df[col]):
            continue

        raw = df[col].astype(str).str.strip()
        non_empty = raw[(raw != "") & (raw.str.lower() != "nan")]
        if non_empty.empty:
            continue

        numeric = pd.to_numeric(clean_series(non_empty), errors="coerce")

        success_ratio = numeric.notna().sum() / len(non_empty)
        if success_ratio >= min_success_ratio:
            df[col] = pd.to_numeric(clean_series(raw), errors="coerce")

    return df


def _try_parse_dates(values):
    """Parse a text Series as dates. Returns (parsed, success_ratio, kwargs).

    First tries pandas' normal format inference; if only part of the values
    parse (a column that mixes e.g. 2016-11-08 and 11/8/2016), retries with
    format="mixed" and keeps whichever works better.
    """
    best = (None, 0.0, {})
    for kwargs in ({}, {"format": "mixed"}):
        if kwargs and not (0.3 <= best[1] < 0.85):
            break
        try:
            with warnings.catch_warnings():
                warnings.simplefilter("ignore")
                parsed = pd.to_datetime(values, errors="coerce", **kwargs)
        except Exception:
            continue
        ratio = float(parsed.notna().mean()) if len(values) else 0.0
        if ratio > best[1]:
            best = (parsed, ratio, kwargs)
    return best


def preprocess_dataframe(df):
    """
    Module 3: Data Preprocessing
      - Tidy column names, drop fully-empty columns and duplicate rows
      - Coerce numeric-looking text columns (currency / % / commas / stray junk)
      - Handle missing values (numeric -> median, categorical -> mode)
      - Convert date-like text columns to datetime (content-based, not name-based)

    NOTE: pandas >= 2.x/3.x can back text columns with a dedicated "str"
    dtype instead of numpy "object", so we only use pandas' dtype helpers
    (is_numeric_dtype / is_datetime64_any_dtype), never `dtype == object`.
    """
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(axis=1, how="all")
    df = df.drop_duplicates()
    df = coerce_numeric_columns(df)

    for col in df.columns:
        if pd.api.types.is_numeric_dtype(df[col]):
            if df[col].isnull().any():
                df[col] = df[col].fillna(df[col].median())
        else:
            if df[col].isnull().any():
                mode = df[col].mode()
                fill_val = mode.iloc[0] if not mode.empty else ""
                df[col] = df[col].fillna(fill_val)

    # Date detection: look at the CONTENT of every remaining text column, so
    # "Order Date", "Timestamp", "Joined_On", "Created" ... are all found.
    # Guards against false positives: most values must contain a digit, be
    # reasonably long (so "May"/"Mon" are not dates), and >=85% of a sample
    # must parse. Identifier-style columns are skipped.
    for col in df.columns:
        if pd.api.types.is_numeric_dtype(df[col]) or pd.api.types.is_datetime64_any_dtype(df[col]):
            continue
        if is_identifier_column(col):
            continue

        raw = df[col].astype(str).str.strip()
        non_empty = raw[(raw != "") & (raw.str.lower() != "nan")]
        if non_empty.empty:
            continue
        if non_empty.str.contains(r"\d", regex=True).mean() < 0.8:
            continue
        if non_empty.str.len().median() < 6:
            continue

        sample = non_empty.sample(min(200, len(non_empty)), random_state=1)
        _, ratio, kwargs = _try_parse_dates(sample)
        if ratio < 0.85:
            continue

        try:
            with warnings.catch_warnings():
                warnings.simplefilter("ignore")
                parsed_full = pd.to_datetime(raw, errors="coerce", **kwargs)
        except Exception:
            continue
        if parsed_full.notna().any():
            df[col] = parsed_full

    return df.reset_index(drop=True)


# --------------------------------------------------------------------------
# Dataset cache: every chart endpoint needs the SAME cleaned dataframe, so
# load + preprocess + profile it once per uploaded file instead of once per
# request (the dashboard fires 5 chart requests + the page itself).
# --------------------------------------------------------------------------
_DATASET_CACHE = {}
_DATASET_CACHE_MAX = 4


def get_dataset():
    """Return {"df", "profile", "picks"} for this session's file, or None."""
    filepath = session.get("current_file")
    if not filepath or not os.path.exists(filepath):
        return None
    stat = os.stat(filepath)
    key = (filepath, stat.st_mtime_ns, stat.st_size)
    entry = _DATASET_CACHE.get(key)
    if entry is None:
        df = preprocess_dataframe(load_dataframe(filepath))
        profile = profile_columns(df)
        entry = {"df": df, "profile": profile, "picks": choose_defaults(df, profile)}
        while len(_DATASET_CACHE) >= _DATASET_CACHE_MAX:
            _DATASET_CACHE.pop(next(iter(_DATASET_CACHE)))
        _DATASET_CACHE[key] = entry
    return entry


def get_current_dataframe():
    """Load the currently active dataset for this session."""
    entry = get_dataset()
    return entry["df"] if entry else None


@app.route("/upload", methods=["GET", "POST"])
@login_required
def upload():
    if request.method == "POST":
        file = request.files.get("dataset")
        if not file or file.filename == "":
            flash("No file selected.", "warning")
            return redirect(url_for("upload"))

        if not allowed_file(file.filename):
            flash("Only CSV and Excel (.csv, .xlsx, .xls) files are supported.", "danger")
            return redirect(url_for("upload"))

        filename = secure_filename(file.filename)
        unique_name = f"{uuid.uuid4().hex}_{filename}"
        filepath = os.path.join(app.config["UPLOAD_FOLDER"], unique_name)
        file.save(filepath)

        # Validate it actually loads
        try:
            df = load_dataframe(filepath)
        except Exception as exc:
            os.remove(filepath)
            flash(f"Could not read the file: {exc}", "danger")
            return redirect(url_for("upload"))

        session["current_file"] = filepath
        session["original_filename"] = filename
        flash(f"Dataset '{filename}' uploaded successfully ({df.shape[0]} rows, {df.shape[1]} columns).", "success")
        return redirect(url_for("dashboard"))

    return render_template("upload.html")


# --------------------------------------------------------------------------
# Module 4: Data Analysis - column profiling + automatic field selection
# --------------------------------------------------------------------------
# Instead of "first numeric column / first text column", every column is first
# PROFILED and given a role:
#   date        real datetime column
#   metric      numeric measure worth summing/averaging (Sales, Price, Rating)
#   ordinal     numeric time-part (Year, Month ...) - an axis, never a measure
#   category    text/bool column with a modest number of distinct values
#   identifier  Row ID, Customer ID, Postal Code, dense 1..N sequences, ...
#   constant    only one distinct value - carries no information
#   text        free text / names / links / very high-cardinality text
# Each chart then picks ONLY from the roles that make sense for it, so IDs,
# constants and free text can never end up on an axis or in a KPI card.
TIME_PART_TOKENS = {"year", "yr", "month", "day", "week", "quarter", "hour", "fy"}
# Measures that should be AVERAGED (summing them is meaningless)...
AVERAGE_TOKENS = {"rating", "rate", "percent", "percentage", "pct", "discount", "score",
                  "ratio", "avg", "average", "mean", "margin", "age", "temperature",
                  "temp", "price", "salary", "gpa", "cgpa", "height", "weight",
                  "latitude", "longitude", "lat", "lon", "probability", "density",
                  "marks", "mark", "attendance"}
# ...and measures that are naturally SUMMED (the LAST word decides: rating_count).
ADDITIVE_LAST_TOKENS = {"count", "total", "qty", "quantity", "units", "votes", "reviews",
                        "orders", "amount", "sold", "sales", "revenue", "profit", "cost",
                        "income", "expense", "spend"}
ADDITIVE_ANY_TOKENS = {"sales", "revenue", "profit", "income", "expense", "spend",
                       "quantity", "qty", "turnover"}
# Words that make a measure a good "headline" metric for the first chart/KPI.
PRIMARY_TOKENS = {"sales", "revenue", "amount", "profit", "income", "price", "cost",
                  "total", "value", "salary", "marks", "score", "quantity", "expense",
                  "spend", "turnover", "population", "gdp"}
# Words that make a text column a good "group by" candidate.
CATEGORY_PREF_TOKENS = {"category", "region", "segment", "type", "status", "department",
                        "dept", "gender", "brand", "class", "group", "genre", "country",
                        "state", "channel", "course", "branch", "college", "city"}
# Words that make a date column a poor primary time axis (ship date, due date ...).
SECONDARY_DATE_TOKENS = {"ship", "shipping", "delivery", "delivered", "due", "expiry",
                         "end", "return", "returned"}
MAX_CATEGORY_UNIQUE = 60


def _is_numeric(s):
    return pd.api.types.is_numeric_dtype(s) and not pd.api.types.is_bool_dtype(s)


def numeric_columns(df, exclude_identifiers=True):
    """Numeric columns, optionally excluding ID/code-like columns
    (Row ID, Postal Code, ...) that aren't meaningful to sum/average."""
    cols = [c for c in df.columns if _is_numeric(df[c])]
    if exclude_identifiers:
        cols = [c for c in cols if not is_identifier_column(c)]
    return cols


def categorical_columns(df):
    return [c for c in df.columns
            if not _is_numeric(df[c]) and not pd.api.types.is_datetime64_any_dtype(df[c])]


def date_columns(df):
    return [c for c in df.columns if pd.api.types.is_datetime64_any_dtype(df[c])]


def _category_series(s):
    """Text -> labels to group by. Hierarchy paths like 'A|B|C' (Amazon-style
    categories) are collapsed to their first level so the chart shows a handful
    of meaningful groups instead of hundreds of unique paths."""
    s = s.astype(str).str.strip()
    if len(s) and s.str.contains("|", regex=False).mean() > 0.8:
        first = s.str.split("|", n=1).str[0].str.strip()
        if first.nunique() < s.nunique():
            return first
    return s


def _is_dense_integer_sequence(s, n):
    """Unique integers that fill a 1..N-like range are row numbers, not data."""
    if n < 50 or not pd.api.types.is_integer_dtype(s) or s.nunique() != n:
        return False
    if not (s.is_monotonic_increasing or s.is_monotonic_decreasing):
        return False
    return (int(s.max()) - int(s.min()) + 1) <= 1.5 * n


def profile_columns(df):
    n = len(df)
    profile = {}
    for col in df.columns:
        s = df[col]
        if pd.api.types.is_datetime64_any_dtype(s):
            role, nun = "date", int(s.nunique())
        elif pd.api.types.is_bool_dtype(s):
            nun = int(s.nunique())
            role = "category" if nun >= 2 else "constant"
        elif _is_numeric(s):
            nun = int(s.nunique())
            if nun < 2:
                role = "constant"
            elif is_identifier_column(col) or _is_dense_integer_sequence(s, n):
                role = "identifier"
            elif set(_tokens(col)) & TIME_PART_TOKENS:
                role = "ordinal"
            else:
                role = "metric"
        else:
            cs = _category_series(s)
            nun = int(cs.nunique())
            avg_len = float(cs.str.len().mean()) if n else 0.0
            if nun < 2:
                role = "constant"
            elif is_identifier_column(col):
                role = "identifier"
            elif avg_len > 60 or nun > MAX_CATEGORY_UNIQUE or nun > 0.5 * n:
                role = "text"
            else:
                role = "category"
        profile[col] = {"role": role, "nunique": nun}
    return profile


def is_additive(df, col):
    """True if values of `col` should be summed (Sales, Quantity), False if
    they should be averaged (Rating, Price, Discount %, Age ...)."""
    toks = _tokens(col)
    if not toks:
        return True
    if "total" in toks or "sum" in toks or toks[-1] in ADDITIVE_LAST_TOKENS:
        return True
    if set(toks) & AVERAGE_TOKENS:
        return False
    if set(toks) & ADDITIVE_ANY_TOKENS:
        return True
    s = df[col]
    if len(s) and s.min() >= 0 and s.max() <= 1.0:      # looks like a ratio
        return False
    return True


def rank_metrics(df, profile):
    """Metric columns, best 'headline' measure first."""
    order = {c: i for i, c in enumerate(df.columns)}
    metrics = [c for c in df.columns if profile[c]["role"] == "metric"]

    def key(c):
        hit = bool(set(_tokens(c)) & PRIMARY_TOKENS)
        return (0 if hit else 1, 0 if is_additive(df, c) else 1, order[c])

    return sorted(metrics, key=key)


def pick_category(df, profile, exclude=None, max_unique=15):
    """Best column to group by: a real category with ~3-15 groups, preferring
    business-like names (Category, Region, Segment ...)."""
    order = {c: i for i, c in enumerate(df.columns)}
    cands = [c for c in df.columns if profile[c]["role"] == "category" and c != exclude]
    if not cands:
        return None

    def key(c):
        nun = profile[c]["nunique"]
        in_range = 2 <= nun <= max_unique
        pref = bool(set(_tokens(c)) & CATEGORY_PREF_TOKENS)
        return (0 if in_range else 1, 0 if pref else 1, abs(nun - 6), order[c])

    return min(cands, key=key)


def pick_date(df, profile):
    order = {c: i for i, c in enumerate(df.columns)}
    dates = [c for c in df.columns if profile[c]["role"] == "date"]
    if not dates:
        return None
    return min(dates, key=lambda c: (1 if set(_tokens(c)) & SECONDARY_DATE_TOKENS else 0, order[c]))


def pick_scatter_pair(df, metrics, max_cols=6):
    """The two measures that move together most strongly (highest |correlation|)."""
    cols = [c for c in metrics if df[c].nunique() > 5][:max_cols]
    if len(cols) < 2:
        cols = metrics[:max_cols]
    if len(cols) < 2:
        return None
    corr = df[cols].corr().abs().fillna(0)
    best, best_val = (cols[0], cols[1]), -1.0
    for i in range(len(cols)):
        for j in range(i + 1, len(cols)):
            v = float(corr.iloc[i, j])
            if v > best_val:
                best, best_val = (cols[i], cols[j]), v
    return best


def choose_defaults(df, profile):
    """Decide, once per dataset, which columns every chart should use."""
    metrics = rank_metrics(df, profile)
    y = metrics[0] if metrics else None
    bar_x = pick_category(df, profile)
    pie_col = pick_category(df, profile, exclude=bar_x) or bar_x

    line_x = pick_date(df, profile)
    ordinals = [c for c in df.columns if profile[c]["role"] == "ordinal"]
    numeric_x = None
    if not line_x and not ordinals:
        pool = [c for c in metrics if c != y and profile[c]["nunique"] >= 5]
        if pool:
            numeric_x = min(pool, key=lambda c: profile[c]["nunique"])

    # Used to build groups when the dataset has no category column at all.
    bin_pool = [c for c in metrics + ordinals if c != y and profile[c]["nunique"] > 2]
    bin_x = bin_pool[0] if bin_pool else y

    return {
        "metrics": metrics,
        "y": y,
        "bar_x": bar_x,
        "pie": pie_col,
        "line_x": line_x,
        "line_ordinal": ordinals[0] if ordinals else None,
        "line_numeric_x": numeric_x,
        "bin_x": bin_x,
        "scatter": pick_scatter_pair(df, metrics) or ((ordinals[0], y) if ordinals and y else None),
        "hist": next((c for c in metrics if df[c].nunique() > 12), metrics[0] if metrics else None),
    }


def _num(v, nd=2):
    """JSON-safe number: NaN / inf become None (JSON has no NaN)."""
    try:
        f = float(v)
    except (TypeError, ValueError):
        return None
    return round(f, nd) if math.isfinite(f) else None


def _num0(v):
    out = _num(v)
    return 0.0 if out is None else out


def compute_kpis(df, ds):
    """Module 6: KPI cards for the top-3 measures. Additive measures show a
    total (Sales, Quantity); rates/ratings/prices show an average."""
    kpis = {"total_records": int(df.shape[0]), "total_columns": int(df.shape[1])}
    full = ds["df"]
    for col in ds["picks"]["metrics"][:3]:
        if is_additive(full, col):
            kpis[f"{col}_total"] = _num0(df[col].sum())
        else:
            kpis[f"{col}_avg"] = _num0(df[col].mean())
    return kpis


def compute_summary_stats(df, ds):
    """Statistical measures for the real measures (IDs/constants/text excluded)."""
    cols = ds["picks"]["metrics"]
    if not cols or df.empty:
        return {}
    desc = df[cols].describe().round(2)
    return {c: {m: ("-" if pd.isna(desc.loc[m, c]) else float(desc.loc[m, c])) for m in desc.index}
            for c in desc.columns}


def build_filter_options(df, profile, max_cols=5, max_values=50):
    cands = [c for c in df.columns if profile[c]["role"] == "category"]
    cands.sort(key=lambda c: profile[c]["nunique"])
    return {c: sorted(_category_series(df[c]).dropna().unique().tolist())[:max_values]
            for c in cands[:max_cols]}


# --------------------------------------------------------------------------
# Module 6: Dashboard Management (with filtering)
# --------------------------------------------------------------------------
@app.route("/dashboard")
@login_required
def dashboard():
    ds = get_dataset()
    if ds is None:
        flash("Please upload a dataset first.", "info")
        return redirect(url_for("upload"))

    df = ds["df"]
    filter_col = request.args.get("filter_col")
    filter_val = request.args.get("filter_val")
    filtered_df = apply_filter(df)
    filter_active = bool(filter_col and filter_val and filter_col in df.columns)

    return render_template(
        "dashboard.html",
        filename=session.get("original_filename"),
        num_cols=numeric_columns(df),
        cat_cols=categorical_columns(df),
        kpis=compute_kpis(filtered_df, ds),
        stats=compute_summary_stats(filtered_df, ds),
        filter_options=build_filter_options(df, ds["profile"]),
        active_filter_col=filter_col,
        active_filter_val=filter_val,
        filter_qs=urlencode({"filter_col": filter_col, "filter_val": filter_val}) if filter_active else "",
        row_count=filtered_df.shape[0],
    )


# --------------------------------------------------------------------------
# Module 5: Data Visualization - JSON endpoints consumed by Chart.js
# --------------------------------------------------------------------------
def apply_filter(df):
    filter_col = request.args.get("filter_col")
    filter_val = request.args.get("filter_val")
    if filter_col and filter_val and filter_col in df.columns:
        return df[_category_series(df[filter_col]) == filter_val]
    return df


def _resolve_col(df, requested, default):
    """Use the ?param column if it really exists, otherwise the auto-picked one."""
    return requested if requested in df.columns else default


def _numeric_to_category(s, max_bins=8):
    """Numeric column -> (group labels, ordered label list). Few distinct values
    are used as-is; otherwise split into equal-count ranges."""
    s = s.replace([np.inf, -np.inf], np.nan)
    if s.nunique() <= 12:
        values = sorted(s.dropna().unique().tolist())
        names = {v: f"{v:g}" for v in values}
        return s.map(names), [names[v] for v in values]
    binned = pd.qcut(s, q=max_bins, duplicates="drop")
    intervals = list(binned.cat.categories)
    names = {iv: f"{iv.left:,.1f} - {iv.right:,.1f}" for iv in intervals}
    return binned.map(names).astype(object), [names[iv] for iv in intervals]


def _category_labels(df, col):
    s = df[col]
    if _is_numeric(s):
        return _numeric_to_category(s)
    if pd.api.types.is_datetime64_any_dtype(s):
        labels = s.dt.strftime("%Y-%m")
        return labels, sorted(labels.dropna().unique().tolist())
    return _category_series(s), None


def _freq(new_alias, old_alias):
    """'ME'/'YE' exist from pandas 2.2; fall back to 'M'/'Y' on older pandas."""
    try:
        pd.tseries.frequencies.to_offset(new_alias)
        return new_alias
    except ValueError:
        return old_alias


def _need_dataset():
    ds = get_dataset()
    if ds is None:
        return None, (jsonify({"error": "No dataset loaded"}), 400)
    return ds, None


@app.route("/api/chart/bar")
@login_required
def chart_bar():
    """Measure grouped by a real category (sum for amounts, mean for
    ratings/prices/percentages). With no measure: record count per group."""
    ds, err = _need_dataset()
    if err:
        return err
    picks = ds["picks"]
    df = apply_filter(ds["df"])
    if df.empty:
        return jsonify({"error": "No rows match the current filter"}), 400

    x_col = _resolve_col(df, request.args.get("x"), picks["bar_x"] or picks["bin_x"])
    y_col = _resolve_col(df, request.args.get("y"), picks["y"])
    if x_col is None:
        return jsonify({"error": "No usable column to group by (every column is an ID, constant or free text)"}), 400

    labels_s, ordered = _category_labels(df, x_col)
    if y_col and y_col != x_col and _is_numeric(df[y_col]):
        # Groups made by binning a numeric column (no real category exists)
        # are compared by average; real categories use sum/mean by measure type.
        agg = "mean" if _is_numeric(df[x_col]) else ("sum" if is_additive(ds["df"], y_col) else "mean")
        grouped = df[y_col].groupby(labels_s).agg(agg)
        y_label = f"{y_col} ({agg})"
    else:
        grouped = labels_s.value_counts()
        y_label = "Record count"

    if ordered:
        grouped = grouped.reindex(ordered).dropna()
    else:
        grouped = grouped.sort_values(ascending=False).head(12)

    return jsonify({
        "labels": grouped.index.astype(str).tolist(),
        "values": [_num(v) for v in grouped.values],
        "x_col": x_col,
        "y_col": y_label,
        "caption": f"{y_label} by {x_col}",
    })


@app.route("/api/chart/pie")
@login_required
def chart_pie():
    """Share of records per group of a real category (top 7 + 'Other')."""
    ds, err = _need_dataset()
    if err:
        return err
    picks = ds["picks"]
    df = apply_filter(ds["df"])
    if df.empty:
        return jsonify({"error": "No rows match the current filter"}), 400

    col = _resolve_col(df, request.args.get("col"), picks["pie"] or picks["bin_x"])
    if col is None:
        return jsonify({"error": "No usable column for a pie chart (every column is an ID, constant or free text)"}), 400

    labels_s, ordered = _category_labels(df, col)
    counts = labels_s.value_counts()
    if ordered:
        counts = counts.reindex(ordered).dropna()
    top = 7
    if len(counts) > top + 1:
        other = counts.iloc[top:].sum()
        counts = counts.iloc[:top]
        counts["Other"] = other

    return jsonify({
        "labels": counts.index.astype(str).tolist(),
        "values": [int(v) for v in counts.values],
        "col": col,
        "caption": f"Share of records by {col}",
    })


@app.route("/api/chart/line")
@login_required
def chart_line():
    """Trend of a measure.
      1. a real date column  -> day / month / year buckets (by date range)
      2. a Year/Month column -> grouped by that column
      3. neither             -> average of the measure across ranges of the most
                                suitable other numeric column
      4. only one numeric column -> its average across record order (last resort)
    """
    ds, err = _need_dataset()
    if err:
        return err
    picks = ds["picks"]
    df = apply_filter(ds["df"])
    if df.empty:
        return jsonify({"error": "No rows match the current filter"}), 400

    y_col = _resolve_col(df, request.args.get("y"), picks["y"])
    if y_col is None or not _is_numeric(df[y_col]):
        return jsonify({"error": "This chart needs a numeric column - none was found in this dataset"}), 400
    agg = "sum" if is_additive(ds["df"], y_col) else "mean"

    x_arg = request.args.get("x")
    x_col = x_arg if x_arg in df.columns else picks["line_x"]
    labels, values = [], []

    if x_col is not None and pd.api.types.is_datetime64_any_dtype(df[x_col]):
        dated = df[[x_col, y_col]].dropna()
        if dated.empty:
            return jsonify({"error": "No dated rows to plot"}), 400
        span_days = (dated[x_col].max() - dated[x_col].min()).days
        if span_days <= 62:
            freq, fmt = "D", "%Y-%m-%d"
        elif span_days <= 2200:
            freq, fmt = _freq("ME", "M"), "%Y-%m"
        else:
            freq, fmt = _freq("YE", "Y"), "%Y"
        series = dated.groupby(pd.Grouper(key=x_col, freq=freq))[y_col].agg(agg)
        series = series[series.index.notna()]
        if agg == "mean":
            series = series.dropna()
        labels = [d.strftime(fmt) for d in series.index]
        values = [_num(v) for v in series.values]
        x_label, caption = x_col, f"{y_col} ({agg}) over {x_col}"
    else:
        ord_col = x_arg if (x_arg in df.columns and _is_numeric(df[x_arg])) else \
            (picks["line_ordinal"] or picks["line_numeric_x"])
        if ord_col is not None and ord_col != y_col and _is_numeric(df[ord_col]):
            is_time_part = ord_col == picks["line_ordinal"]
            agg2 = agg if is_time_part else "mean"
            xs = df[ord_col].replace([np.inf, -np.inf], np.nan)
            if xs.nunique() <= 40:
                series = df[y_col].groupby(xs).agg(agg2).sort_index()
                labels = [f"{v:g}" for v in series.index]
            else:
                binned = pd.qcut(xs, q=15, duplicates="drop")
                series = df[y_col].groupby(binned, observed=True).agg(agg2)
                labels = [f"{iv.left:,.1f} - {iv.right:,.1f}" for iv in series.index]
            values = [_num(v) for v in series.values]
            x_label, caption = ord_col, f"{y_col} ({agg2}) by {ord_col}"
        else:
            data = df[y_col].replace([np.inf, -np.inf], np.nan).dropna().reset_index(drop=True)
            if data.empty:
                return jsonify({"error": f"'{y_col}' has no data to plot"}), 400
            start = 0
            for b in np.array_split(data.values, min(30, len(data))):
                if len(b):
                    values.append(_num(np.mean(b)))
                    labels.append(f"{start + 1}-{start + len(b)}")
                    start += len(b)
            x_label = "Record order (no date column found)"
            caption = f"{y_col} (mean) across record order"

    return jsonify({
        "labels": labels,
        "values": values,
        "x_col": x_label,
        "y_col": f"{y_col} ({agg})",
        "caption": caption,
    })


@app.route("/api/chart/scatter")
@login_required
def chart_scatter():
    """The two measures that are most strongly related (highest correlation)."""
    ds, err = _need_dataset()
    if err:
        return err
    picks = ds["picks"]
    df = apply_filter(ds["df"])
    if df.empty:
        return jsonify({"error": "No rows match the current filter"}), 400

    pair = picks["scatter"]
    x_col = request.args.get("x")
    y_col = request.args.get("y")
    if x_col in df.columns and y_col in df.columns and _is_numeric(df[x_col]) and _is_numeric(df[y_col]):
        pass
    elif pair:
        x_col, y_col = pair
    elif picks["y"]:
        y_col, x_col = picks["y"], None
    else:
        return jsonify({"error": "This chart needs numeric columns - none were found in this dataset"}), 400

    if x_col is None:      # a single numeric column: plot it against record order
        data = df[y_col].replace([np.inf, -np.inf], np.nan).dropna().reset_index(drop=True)
        sample = data.sample(min(500, len(data)), random_state=1).sort_index()
        points = [{"x": int(i), "y": _num(v)} for i, v in sample.items()]
        x_col = "Record order"
    else:
        clean = df[[x_col, y_col]].replace([np.inf, -np.inf], np.nan).dropna()
        if clean.empty:
            return jsonify({"error": "No complete rows to plot"}), 400
        sample = clean.sample(min(500, len(clean)), random_state=1)
        points = [{"x": _num(a), "y": _num(b)} for a, b in zip(sample[x_col], sample[y_col])]

    return jsonify({
        "points": points,
        "x_col": x_col,
        "y_col": y_col,
        "caption": f"{y_col} vs {x_col}",
    })


@app.route("/api/chart/histogram")
@login_required
def chart_histogram():
    ds, err = _need_dataset()
    if err:
        return err
    picks = ds["picks"]
    df = apply_filter(ds["df"])
    if df.empty:
        return jsonify({"error": "No rows match the current filter"}), 400

    col = _resolve_col(df, request.args.get("col"), picks["hist"])
    if col is None or not _is_numeric(df[col]):
        return jsonify({"error": "This chart needs a numeric column - none was found in this dataset"}), 400

    data = df[col].replace([np.inf, -np.inf], np.nan).dropna()
    if data.empty:
        return jsonify({"error": f"'{col}' has no data to plot"}), 400

    use_log = False
    if data.nunique() <= 12:
        # Few distinct values (ratings, quantities ...): one bar per value.
        vc = data.value_counts().sort_index()
        labels = [f"{v:g}" for v in vc.index]
        counts = vc.values
    else:
        # Strictly positive, heavily right-skewed data (Sales, Profit ...) is
        # binned on a log scale so the shape of the distribution is visible.
        use_log = bool((data > 0).all()) and float(data.skew()) > 2
        if use_log:
            counts, log_edges = np.histogram(np.log10(data), bins=10)
            bin_edges = 10 ** log_edges
        else:
            counts, bin_edges = np.histogram(data, bins=10)
        labels = [f"{bin_edges[i]:,.1f}-{bin_edges[i + 1]:,.1f}" for i in range(len(bin_edges) - 1)]

    return jsonify({
        "labels": labels,
        "values": [int(c) for c in counts],
        "col": col,
        "log_scale": use_log,
        "caption": f"Distribution of {col}" + (" (log-scale bins)" if use_log else ""),
    })


# --------------------------------------------------------------------------
# Module 7: Report Generation
# --------------------------------------------------------------------------
@app.route("/report/download")
@login_required
def download_report():
    df = get_current_dataframe()
    if df is None:
        flash("No dataset loaded.", "warning")
        return redirect(url_for("upload"))

    df = apply_filter(df)

    buffer = io.BytesIO()
    df.to_csv(buffer, index=False)
    buffer.seek(0)

    out_name = f"report_{session.get('original_filename', 'dataset')}"
    if not out_name.lower().endswith(".csv"):
        out_name = os.path.splitext(out_name)[0] + ".csv"

    return send_file(
        buffer,
        mimetype="text/csv",
        as_attachment=True,
        download_name=out_name,
    )


@app.route("/api/columns")
@login_required
def api_columns():
    """Helper endpoint: column roles and the fields each chart uses by default."""
    ds, err = _need_dataset()
    if err:
        return err
    df = ds["df"]
    return jsonify({
        "numeric": numeric_columns(df),
        "categorical": categorical_columns(df),
        "date": date_columns(df),
        "roles": ds["profile"],
        "defaults": ds["picks"],
    })


if __name__ == "__main__":
    app.run(debug=True)