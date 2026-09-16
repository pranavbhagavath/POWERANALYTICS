import io
import re
import math
from typing import Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go


# ============================================================
# POWERANALYTICS
# Graph-first Excel analytics dashboard
# ============================================================

st.set_page_config(
    page_title="POWERANALYTICS",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded",
)

# -----------------------------
# Theme
# -----------------------------
st.markdown("""
<style>
:root{
    --bg:#061525; --panel:#0c2743; --panel2:#102f50;
    --line:#1d527d; --text:#f4f8ff; --muted:#8fa8c2;
    --cyan:#20d6ea; --green:#26df69; --orange:#ffad3d;
    --purple:#a77cff; --red:#ff5d6c;

    /* Spacing scale — used consistently across every component below
       so vertical rhythm stays uniform instead of ad-hoc pixel values. */
    --space-1:4px;  --space-2:8px;  --space-3:12px; --space-4:16px;
    --space-5:20px; --space-6:28px; --space-7:36px; --space-8:48px;
}
*{box-sizing:border-box}
.stApp{
    background:
        radial-gradient(circle at 88% 0%, rgba(31,101,154,.18), transparent 27%),
        linear-gradient(135deg,#061525 0%,#071827 58%,#091426 100%);
    color:var(--text);
}
.block-container{
    max-width:1540px;
    padding:var(--space-6) var(--space-5) var(--space-8) var(--space-5);
}
h1,h2,h3,h4,p,div,span,label,button{font-family:Arial,sans-serif}
h1,h2,h3,h4{color:var(--text)!important}

/* -----------------------------
   Sidebar
   ----------------------------- */
section[data-testid="stSidebar"]{
    background:#041321;
    border-right:1px solid #173d5d;
}
section[data-testid="stSidebar"] > div{padding:var(--space-5) var(--space-4) var(--space-6) var(--space-4)}
section[data-testid="stSidebar"] div[data-testid="stVerticalBlock"]{gap:var(--space-3)}
section[data-testid="stSidebar"] .stButton button{
    min-height:44px;border-radius:9px;
    background:#0d2d4d;border:1px solid #245b84;
    margin-top:var(--space-2);
}
section[data-testid="stSidebar"] .stFileUploader > div{
    border:1px solid #24577e;border-radius:10px;background:#071d31;
}
section[data-testid="stSidebar"] .stSelectbox > div > div{
    background:#081a2b!important;border:1px solid #24577e!important;
    border-radius:9px!important;
}
section[data-testid="stSidebar"] h3{
    margin:var(--space-5) 0 var(--space-2) 0!important;
    font-size:12px!important;letter-spacing:.6px;color:#8fa8c2!important;
}

/* -----------------------------
   Header / meta bar
   ----------------------------- */
.power-header{
    background:linear-gradient(100deg,#0d3153 0%,#0a2038 100%);
    border:1px solid #22547d;border-radius:14px;
    padding:24px 28px;margin:0 0 var(--space-4) 0;
    box-shadow:0 8px 28px rgba(0,0,0,.12);
}
.brand{font-size:28px;line-height:1.05;font-weight:850;letter-spacing:.35px}
.brand .cyan{color:var(--cyan)}
.subtitle{
    color:#8fa8c2;font-size:11.5px;line-height:1.5;
    margin-top:var(--space-2);letter-spacing:.45px;
}
.meta{
    background:#08213a;border:1px solid #1d527d;border-radius:8px;
    padding:var(--space-3) var(--space-4);color:#9bb4cc;font-size:10.5px;line-height:1.6;
    margin:0 0 var(--space-4) 0;
}

/* -----------------------------
   Section headings & titles
   ----------------------------- */
.section-heading{
    font-size:15px;font-weight:850;letter-spacing:.15px;
    margin:var(--space-7) 0 var(--space-3) 0;
}
.section-heading:first-of-type{margin-top:var(--space-5)}
.section-title{
    font-size:13px;line-height:1.3;font-weight:850;
    color:#f5f8fc;margin:0 0 var(--space-3) 0;
}
.section-title small{color:#7898b7;font-size:10px;font-weight:500;margin-left:var(--space-1)}

/* -----------------------------
   KPI cards
   ----------------------------- */
.kpi{
    background:linear-gradient(145deg,#103456 0%,#0b2743 100%);
    border:1px solid #1e557f;border-radius:12px;min-height:120px;
    padding:var(--space-4) var(--space-5);box-sizing:border-box;
    box-shadow:0 5px 18px rgba(0,0,0,.08);
}
.kpi .label{
    color:#8ca8c3;font-size:9.5px;line-height:1.3;
    font-weight:850;letter-spacing:.75px;text-transform:uppercase;
}
.kpi .value{color:#fff;font-size:26px;line-height:1.15;font-weight:850;margin-top:var(--space-3)}
.kpi .sub{color:#7fa2c2;font-size:9.5px;line-height:1.5;margin-top:var(--space-2)}

/* -----------------------------
   Snapshot list items
   ----------------------------- */
.snapshot-item{
    background:#0d2946;border:1px solid #1b4d74;border-radius:10px;
    padding:var(--space-3) var(--space-4);margin-bottom:var(--space-3);
}
.snapshot-label{color:#8da7bf;font-size:9px;font-weight:850;letter-spacing:.55px}
.snapshot-value{color:#f7fbff;font-size:22px;font-weight:850;line-height:1.2;margin-top:var(--space-2)}

/* -----------------------------
   Streamlit layout primitives
   ----------------------------- */
div[data-testid="stHorizontalBlock"]{gap:var(--space-5)}
div[data-testid="stVerticalBlock"]{gap:var(--space-2)}
div[data-testid="stMetric"]{background:transparent}
div[data-testid="stPlotlyChart"]{margin:0!important}
div[data-testid="stPlotlyChart"] > div{margin:0!important}
hr{border-color:#163b59!important;margin:var(--space-4) 0!important}
.stCaption{color:#7898b7!important;margin-bottom:var(--space-2)!important}
.stButton button{border-radius:8px}
.stAlert{padding:var(--space-3) var(--space-4)!important}

@media (max-width:900px){
    .block-container{padding:var(--space-4) var(--space-3) var(--space-6) var(--space-3)}
    .power-header{padding:18px 18px}
    .brand{font-size:23px}
    .meta{font-size:9.5px}
    div[data-testid="stHorizontalBlock"]{gap:var(--space-3)}
}
</style>
""", unsafe_allow_html=True)


# ============================================================
# Authentication
# ============================================================
USERS = {"prahal": "12345678"}

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    st.markdown("""
    <div style="max-width:520px;margin:12vh auto 24px auto;
         background:#0c2743;border:1px solid #22547d;border-radius:14px;padding:36px 34px;">
      <div class="brand">POWER<span style="color:#21d4e8">ANALYTICS</span></div>
      <div class="subtitle" style="margin-top:10px">UTILITY • RENEWABLE ENERGY • PRODUCTION • PERFORMANCE INTELLIGENCE</div>
    </div>
    """, unsafe_allow_html=True)
    st.markdown('<div style="max-width:520px;margin:0 auto">', unsafe_allow_html=True)
    with st.form("login"):
        u = st.text_input("Username")
        p = st.text_input("Password", type="password")
        st.markdown('<div style="height:4px"></div>', unsafe_allow_html=True)
        if st.form_submit_button("SIGN IN", use_container_width=True):
            if USERS.get(u) == p:
                st.session_state.authenticated = True
                st.session_state.username = u
                st.rerun()
            else:
                st.error("Invalid username or password.")
    st.markdown('</div>', unsafe_allow_html=True)
    st.stop()


# ============================================================
# Domain / semantic vocabulary
# ============================================================
ALIASES = {
    "date": ["date", "day", "time", "timestamp", "datetime"],
    "generation": ["generation", "generated", "production", "produced", "output", "energy generated"],
    "consumption": ["consumption", "consumed", "usage", "use", "water use", "energy use"],
    "flow": ["flow", "throughput", "feed", "inlet", "outlet", "received", "treated"],
    "level": ["level", "tank level", "water level", "storage level"],
    "efficiency": ["efficiency", "eff", "actual efficiency", "performance ratio", "pr"],
    "target": ["target", "setpoint", "required", "goal", "planned"],
    "recovery": ["recovery", "ro recovery", "recovered"],
    "quality": ["quality", "conductivity", "tds", "ph", "turbidity", "svi"],
    "downtime": ["downtime", "down time", "stoppage", "idle"],
    "availability": ["availability", "uptime", "operating"],
    "temperature": ["temperature", "temp"],
    "pressure": ["pressure", "bar", "psi"],
    "power": ["power", "kw", "mw", "load"],
    "voltage": ["voltage", "volt", "kv"],
    "current": ["current", "amp", "amps"],
    "frequency": ["frequency", "hz"],
    "power_factor": ["power factor", "pf"],
    "rainfall": ["rain", "rainfall", "precipitation"],
    "shift": ["shift", "a shift", "b shift", "c shift"],
    "status": ["status", "condition", "state", "running", "stopped"],
    "source": ["source", "category", "type", "material", "feed type"],
    "equipment": ["equipment", "machine", "pump", "motor", "unit"],
    "loss": ["loss", "losses", "waste"],
}

DOMAIN_WORDS = {
    "ETP": ["etp", "effluent", "treated", "ro", "di", "conductivity"],
    "WTP": ["wtp", "water treatment", "raw water", "clarifier", "filter"],
    "STP": ["stp", "sewage", "svi", "sludge", "mlss", "mld"],
    "CTP": ["ctp", "cooling", "tower", "circulation"],
    "RO & UF": ["ro", "uf", "membrane", "permeate", "reject", "recovery"],
    "Storage Tank": ["storage", "tank", "level"],
    "Rain Water": ["rain", "rainwater", "rain water", "harvesting"],
    "MEE": ["mee", "evaporator", "evaporation"],
    "Composting": ["compost", "composting", "organic", "waste"],
    "Efficiency": ["efficiency", "actual efficiency", "target efficiency"],
    "Electrical": ["electrical", "voltage", "current", "power", "frequency", "energy"],
}


def norm(x) -> str:
    return re.sub(r"[^a-z0-9]+", " ", str(x).lower()).strip()


def semantic(col: str) -> Tuple[str, float]:
    n = norm(col)
    best, score = "metric", 0.0
    for key, words in ALIASES.items():
        for w in words:
            nw = norm(w)
            if n == nw:
                return key, 1.0
            if nw and nw in n:
                s = min(0.96, 0.72 + 0.035 * len(nw.split()))
                if s > score:
                    best, score = key, s
    return best, score


def detect_domain(sheet_name: str, columns: List[str]) -> str:
    text = norm(sheet_name + " " + " ".join(map(str, columns)))
    scores = {}
    for domain, words in DOMAIN_WORDS.items():
        scores[domain] = sum(1 for w in words if norm(w) in text)
    best = max(scores, key=scores.get)
    return best if scores[best] else "General Operations"


# ============================================================
# Excel parsing
# Handles normal tables AND matrix-style worksheets such as:
# Target efficiency | 70 | 70 | 70 | ...
# Actual efficiency | 58 | 1.4 | ...
# ============================================================
def make_unique_columns(cols):
    out, seen = [], {}
    for c in cols:
        c = str(c).strip() if str(c).strip() else "Column"
        k = seen.get(c, 0)
        out.append(c if k == 0 else f"{c}_{k+1}")
        seen[c] = k + 1
    return out


def clean_frame(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df = df.dropna(axis=0, how="all").dropna(axis=1, how="all")
    if df.empty:
        return df
    df.columns = make_unique_columns(df.columns)
    return df.reset_index(drop=True)


def detect_header(raw: pd.DataFrame) -> int:
    if raw.empty:
        return 0
    best_i, best_score = 0, -1
    for i in range(min(len(raw), 25)):
        vals = [x for x in raw.iloc[i].tolist() if pd.notna(x)]
        if not vals:
            continue
        strings = sum(isinstance(x, str) for x in vals)
        unique = len(set(map(str, vals)))
        score = strings * 2 + unique * .3
        if i > 0:
            score += 0.3
        if score > best_score:
            best_i, best_score = i, score
    return best_i


def parse_matrix(raw: pd.DataFrame) -> Optional[pd.DataFrame]:
    """
    Detect transposed/day-matrix sheets where the first column contains
    metric names and the remaining columns are days/numbers.
    """
    if raw.shape[0] < 3 or raw.shape[1] < 4:
        return None

    first = raw.iloc[:, 0].astype(str).str.strip()
    numeric_cols = []
    for c in raw.columns[1:]:
        vals = pd.to_numeric(raw[c], errors="coerce")
        numeric_cols.append(vals.notna().mean())

    if not numeric_cols:
        return None

    numeric_density = float(np.mean(numeric_cols))
    metric_text_density = float(first.replace({"nan": ""}).str.len().gt(0).mean())

    if numeric_density < 0.45 or metric_text_density < 0.55:
        return None

    # Day-like columns: 1..31, dates, or short sequential labels
    labels = list(raw.columns[1:])
    day_like = 0
    for x in labels:
        s = str(x).strip()
        try:
            v = float(s)
            if 1 <= v <= 31 and v.is_integer():
                day_like += 1
        except Exception:
            if re.search(r"\b(day|date)\b", s.lower()):
                day_like += 1

    if day_like < max(3, int(len(labels) * .45)):
        return None

    records = []
    for _, row in raw.iterrows():
        metric_name = str(row.iloc[0]).strip()
        if not metric_name or metric_name.lower() == "nan":
            continue
        vals = []
        for col in labels:
            v = pd.to_numeric(row[col], errors="coerce")
            vals.append(v)
        if sum(pd.notna(vals)) >= 2:
            records.append((metric_name, vals))

    if not records:
        return None

    # Convert to long-ish daily table.
    out = {}
    day_values = []
    for col in labels:
        s = str(col).strip()
        try:
            v = float(s)
            day_values.append(int(v) if v.is_integer() else v)
        except Exception:
            day_values.append(s)

    out["Day"] = day_values
    for metric, vals in records:
        name = metric
        if name in out:
            name = name + "_2"
        out[name] = vals

    df = pd.DataFrame(out)
    # Preserve month from workbook filename later.
    return clean_frame(df)


def parse_sheet(xls: pd.ExcelFile, sheet: str) -> pd.DataFrame:
    raw0 = pd.read_excel(xls, sheet_name=sheet, header=None)
    raw0 = clean_frame(raw0)
    if raw0.empty:
        return raw0

    matrix = parse_matrix(raw0)
    if matrix is not None:
        return matrix

    h = detect_header(raw0)
    df = pd.read_excel(xls, sheet_name=sheet, header=h)
    return clean_frame(df)


def find_date_column(df: pd.DataFrame) -> Optional[str]:
    for c in df.columns:
        s, conf = semantic(c)
        if s == "date" and conf >= .72:
            parsed = pd.to_datetime(df[c], errors="coerce", dayfirst=True)
            if parsed.notna().mean() >= .35:
                return c
    # fallback: test all object columns
    for c in df.columns:
        parsed = pd.to_datetime(df[c], errors="coerce", dayfirst=True)
        if parsed.notna().mean() >= .70:
            return c
    return None


def materialize_matrix_dates(df: pd.DataFrame, filename: str) -> pd.DataFrame:
    if "Day" not in df.columns:
        return df
    d = pd.to_numeric(df["Day"], errors="coerce")
    if d.notna().mean() < .7:
        return df

    m = re.search(r"(January|February|March|April|May|June|July|August|September|October|November|December)", filename, re.I)
    y = re.search(r"(20\d{2})", filename)
    if not m:
        return df

    month = pd.to_datetime(m.group(1), format="%B").month
    year = int(y.group(1)) if y else pd.Timestamp.today().year
    out = df.copy()
    out["Date"] = pd.to_datetime(
        {"year": year, "month": month, "day": d.astype("Int64")},
        errors="coerce"
    )
    return out


# ============================================================
# Data classification
# ============================================================
def numeric_columns(df):
    cols = []
    for c in df.columns:
        s = pd.to_numeric(df[c], errors="coerce")
        if s.notna().sum() >= 3 and s.nunique(dropna=True) >= 2:
            cols.append(c)
    return cols


def profile(df: pd.DataFrame) -> List[dict]:
    result = []
    for c in df.columns:
        sem, conf = semantic(c)
        n = pd.to_numeric(df[c], errors="coerce")
        result.append({
            "column": c,
            "semantic": sem,
            "confidence": conf,
            "numeric": n.notna().sum() >= 3,
            "nunique": n.nunique(dropna=True),
            "mean": float(n.mean()) if n.notna().any() else np.nan,
        })
    return result


def choose_metrics(df: pd.DataFrame):
    prof = profile(df)
    nums = [p for p in prof if p["numeric"]]

    # Remove obvious serial/index columns unless their names strongly imply a KPI.
    usable = []
    for p in nums:
        n = norm(p["column"])
        if n in {"s n", "sn", "serial no", "serial number", "no", "sr no", "index"}:
            continue
        if p["nunique"] <= 1:
            continue
        usable.append(p)

    semantic_groups = {}
    for p in usable:
        semantic_groups.setdefault(p["semantic"], []).append(p)

    def first_sem(keys):
        for k in keys:
            if k in semantic_groups:
                return semantic_groups[k][0]["column"]
        return None

    primary = first_sem([
        "generation", "flow", "efficiency", "level", "consumption",
        "recovery", "power", "quality", "rainfall", "availability"
    ])

    secondary = first_sem([
        "efficiency", "recovery", "consumption", "flow", "level",
        "quality", "power", "temperature", "availability"
    ])

    if secondary == primary:
        secondary = None

    # For sheets with no semantic hits, use the strongest-looking columns,
    # but NEVER serial/index.
    if primary is None and usable:
        primary = usable[0]["column"]
    if secondary is None:
        for p in usable:
            if p["column"] != primary:
                secondary = p["column"]
                break

    return usable, primary, secondary


# ============================================================
# Chart intelligence
# IMPORTANT: chart type is selected per segment, not globally.
# ============================================================
def chart_title(col, semantic_name):
    labels = {
        "generation": "Energy / Production Generation",
        "consumption": "Consumption",
        "flow": "Process Flow / Throughput",
        "level": "Tank / Storage Level",
        "efficiency": "Efficiency Performance",
        "target": "Target",
        "recovery": "Recovery",
        "quality": "Quality",
        "downtime": "Downtime",
        "availability": "Availability",
        "temperature": "Temperature",
        "pressure": "Pressure",
        "power": "Power",
        "voltage": "Voltage",
        "current": "Current",
        "frequency": "Frequency",
        "power_factor": "Power Factor",
        "rainfall": "Rainfall",
        "loss": "Loss",
    }
    return labels.get(semantic_name, str(col))


def base_fig(fig, height=300):
    fig.update_layout(
        height=height,
        margin=dict(l=38, r=12, t=38, b=34),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(color="#dce9f7", size=11),
        legend=dict(orientation="h", y=1.08, x=0),
        xaxis=dict(gridcolor="#173b5a", zerolinecolor="#173b5a"),
        yaxis=dict(gridcolor="#173b5a", zerolinecolor="#173b5a"),
    )
    return fig


def make_line(df, date_col, metric):
    x = pd.to_datetime(df[date_col], errors="coerce")
    y = pd.to_numeric(df[metric], errors="coerce")
    d = pd.DataFrame({"Date": x, "Value": y}).dropna()
    if len(d) < 2:
        return None
    # Guard against columns that are actually time-only (e.g. "23:59:59"),
    # which pandas anchors to the 1970-01-01 epoch and produces a
    # meaningless one-day chart instead of a real trend.
    if d["Date"].max() - d["Date"].min() < pd.Timedelta(hours=36) and \
       d["Date"].dt.year.between(1969, 1970).all():
        return None
    d = d.sort_values("Date")
    fig = px.line(d, x="Date", y="Value", markers=True, title=chart_title(metric, semantic(metric)[0]))
    return base_fig(fig, 315)


def make_bar(df, cat_col, metric):
    if cat_col == metric:
        return None
    d = df[[cat_col, metric]].copy()
    d.columns = ["__cat__", "__val__"]
    d["__val__"] = pd.to_numeric(d["__val__"], errors="coerce")
    d = d.dropna().groupby("__cat__", as_index=False)["__val__"].sum()
    d.columns = [cat_col, metric]
    if len(d) < 2:
        return None
    d = d.sort_values(metric, ascending=False).head(15)
    fig = px.bar(d, x=cat_col, y=metric, title=chart_title(metric, semantic(metric)[0]), text_auto=".2s")
    return base_fig(fig, 300)


def make_donut(df, cat_col, metric):
    if cat_col == metric:
        return None
    d = df[[cat_col, metric]].copy()
    d.columns = ["__cat__", "__val__"]
    d["__val__"] = pd.to_numeric(d["__val__"], errors="coerce")
    d = d.dropna().groupby("__cat__", as_index=False)["__val__"].sum()
    d.columns = [cat_col, metric]
    if len(d) < 2 or d[metric].sum() == 0:
        return None
    d = d.sort_values(metric, ascending=False).head(10)
    fig = px.pie(d, names=cat_col, values=metric, hole=.58,
                 title=f"{chart_title(metric, semantic(metric)[0])} — Contribution")
    fig.update_traces(textposition="inside", textinfo="percent")
    return base_fig(fig, 315)


def make_scatter(df, xcol, ycol):
    if xcol == ycol:
        return None
    d = df[[xcol, ycol]].copy()
    d.columns = ["__x__", "__y__"]
    d["__x__"] = pd.to_numeric(d["__x__"], errors="coerce")
    d["__y__"] = pd.to_numeric(d["__y__"], errors="coerce")
    d = d.dropna()
    if len(d) < 5:
        return None
    d.columns = [xcol, ycol]
    fig = px.scatter(d, x=xcol, y=ycol, trendline=None,
                     title=f"{chart_title(ycol, semantic(ycol)[0])} vs {chart_title(xcol, semantic(xcol)[0])}")
    return base_fig(fig, 300)


def make_hist(df, metric):
    y = pd.to_numeric(df[metric], errors="coerce").dropna()
    if len(y) < 8 or y.nunique() < 4:
        return None
    tmp = pd.DataFrame({"Value": y})
    fig = px.histogram(tmp, x="Value", nbins=min(20, max(6, int(math.sqrt(len(y))))),
                       title=f"{chart_title(metric, semantic(metric)[0])} Distribution")
    return base_fig(fig, 290)


def make_gauge(value, title, minimum=None, maximum=None, target=None):
    if not np.isfinite(value):
        return None
    if minimum is None or maximum is None:
        if target is not None and np.isfinite(target):
            minimum = 0
            maximum = max(target * 1.35, value * 1.15, 1)
        else:
            minimum = 0
            maximum = max(abs(value) * 1.35, 1)

    fig = go.Figure(go.Indicator(
        mode="gauge+number",
        value=value,
        title={"text": title, "font": {"size": 14}},
        gauge={
            "axis": {"range": [minimum, maximum], "tickcolor": "#89a6c2"},
            "bar": {"color": "#20d4e8"},
            "bgcolor": "#0a2036",
            "bordercolor": "#1d527d",
            "steps": [
                {"range": [minimum, minimum + (maximum-minimum)*.65], "color": "#0d3451"},
                {"range": [minimum + (maximum-minimum)*.65, maximum], "color": "#183b50"},
            ],
        }
    ))
    return base_fig(fig, 275)


# ============================================================
# KPI / target logic
# ============================================================
def fmt_value(v):
    if not np.isfinite(v):
        return "—"
    if abs(v) >= 1000000:
        return f"{v/1000000:.2f}M"
    if abs(v) >= 1000:
        return f"{v:,.0f}"
    if abs(v) >= 100:
        return f"{v:,.1f}"
    return f"{v:,.2f}"


def find_col_by_sem(df, key):
    candidates = []
    for c in df.columns:
        s, conf = semantic(c)
        if s == key:
            candidates.append((conf, c))
    return max(candidates)[1] if candidates else None


def target_pair(df):
    target = find_col_by_sem(df, "target")
    actual = None
    for key in ["efficiency", "generation", "flow", "level", "consumption", "recovery", "power"]:
        c = find_col_by_sem(df, key)
        if c and c != target:
            actual = c
            break
    return actual, target


def target_pct(df):
    actual, target = target_pair(df)
    if not actual or not target:
        return None
    a = pd.to_numeric(df[actual], errors="coerce").mean()
    t = pd.to_numeric(df[target], errors="coerce").mean()
    if not np.isfinite(a) or not np.isfinite(t) or t == 0:
        return None
    return float(a / t * 100), actual, target, a, t


def condition_value(df):
    for key in ["availability", "quality", "efficiency", "recovery"]:
        c = find_col_by_sem(df, key)
        if c:
            y = pd.to_numeric(df[c], errors="coerce").dropna()
            if len(y):
                return float(y.iloc[-1]), c
    return None


# ============================================================
# Dashboard sections
# ============================================================
def kpi_card(label, value, sub="", accent="cyan"):
    st.markdown(f"""
    <div class="kpi">
      <div class="label">{label}</div>
      <div class="value">{value}</div>
      <div class="sub">{sub}</div>
    </div>
    """, unsafe_allow_html=True)


def render_kpis(df, primary, secondary):
    nums = numeric_columns(df)
    vals = {}
    for c in nums:
        vals[c] = pd.to_numeric(df[c], errors="coerce").dropna()

    if primary and primary in vals and len(vals[primary]):
        pv = vals[primary].mean()
        psem = semantic(primary)[0]
        ptitle = chart_title(primary, psem)
    else:
        pv, ptitle = np.nan, "Business Metric"

    if secondary and secondary in vals and len(vals[secondary]):
        sv = vals[secondary].mean()
        stitle = chart_title(secondary, semantic(secondary)[0])
    else:
        sv, stitle = np.nan, "Secondary Metric"

    c1,c2,c3,c4 = st.columns(4)
    with c1: kpi_card(ptitle, fmt_value(pv) if np.isfinite(pv) else "—", "Average / selected period")
    with c2: kpi_card(stitle, fmt_value(sv) if np.isfinite(sv) else "—", "Average / selected period")
    with c3: kpi_card("Data Points", f"{len(df):,}", "Usable records")
    with c4:
        tp = target_pct(df)
        if tp:
            pct, actual, target, a, t = tp
            kpi_card("Target Achievement", f"{pct:.0f}%", f"{chart_title(actual, semantic(actual)[0])} vs target")
        else:
            kpi_card("Metrics Found", f"{len(nums)}", "Usable quantitative fields")


def render_main_row(df, domain, primary, secondary, date_col):
    # Left: operational snapshot
    # Center: primary performance
    # Right: target/condition
    left, mid, right = st.columns([1.05, 1.85, 1.0])

    with left:
        st.markdown('<div class="section-title">Operational Snapshot <small>selected worksheet</small></div>', unsafe_allow_html=True)
        candidates = []
        for c in numeric_columns(df):
            s = semantic(c)[0]
            if c == primary or s in {"consumption","flow","level","recovery","availability","power"}:
                y = pd.to_numeric(df[c], errors="coerce").dropna()
                if len(y):
                    candidates.append((chart_title(c,s), float(y.mean()), c))
        if candidates:
            for label, val, _ in candidates[:3]:
                st.markdown(f"""
                <div class="snapshot-item">
                  <div class="snapshot-label">{label.upper()}</div>
                  <div class="snapshot-value">{fmt_value(val)}</div>
                </div>
                """, unsafe_allow_html=True)
        else:
            st.info("No quantitative operating metric was found.")

    with mid:
        title = f"{domain} Performance"
        st.markdown(f'<div class="section-title">{title} <small>primary trend / comparison</small></div>', unsafe_allow_html=True)

        fig = None
        if primary:
            if date_col:
                fig = make_line(df, date_col, primary)
                if fig is None:
                    fig = make_hist(df, primary)
            else:
                # If the sheet is category-driven, use bar.
                cats = [c for c in df.columns if c != primary and df[c].dtype == "object" and df[c].nunique() <= 20]
                if cats:
                    fig = make_bar(df, cats[0], primary)
                else:
                    fig = make_hist(df, primary)

        if fig is not None:
            st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False}, key=f"main_primary_{sheet}")
        else:
            st.info("Not enough data for a primary performance chart.")

    with right:
        st.markdown('<div class="section-title">Target Achievement <small>actual vs target</small></div>', unsafe_allow_html=True)
        tp = target_pct(df)
        if tp:
            pct, actual, target, a, t = tp
            fig = go.Figure(go.Indicator(
                mode="number+gauge",
                value=pct,
                number={"suffix":"%", "font":{"size":32}},
                gauge={
                    "axis":{"range":[0, max(100, min(150, pct*1.25))]},
                    "bar":{"color":"#24df69"},
                    "bgcolor":"#0a2036",
                    "bordercolor":"#1d527d",
                    "steps":[
                        {"range":[0,70],"color":"#162d43"},
                        {"range":[70,90],"color":"#153f42"},
                        {"range":[90,150],"color":"#123b31"},
                    ],
                }
            ))
            fig.update_layout(title=f"{chart_title(actual, semantic(actual)[0])} / Target",
                              title_font_size=13)
            st.plotly_chart(base_fig(fig,270), use_container_width=True, config={"displayModeBar":False}, key=f"main_target_{sheet}")
        else:
            cv = condition_value(df)
            if cv:
                value, col = cv
                fig = make_gauge(value, chart_title(col, semantic(col)[0]))
                st.plotly_chart(fig, use_container_width=True, config={"displayModeBar":False}, key=f"main_condition_{sheet}")
            else:
                st.info("No reliable target / condition metric.")


def render_secondary_row(df, primary, secondary, date_col):
    a,b,c,d = st.columns([1.15,1.15,1.15,1.35])

    # 1) Donut: only categorical contribution data
    with a:
        st.markdown('<div class="section-title">Source-wise Contribution <small>composition</small></div>', unsafe_allow_html=True)
        metric_for_donut = primary or (numeric_columns(df)[0] if numeric_columns(df) else None)
        cats = []
        for c0 in df.columns:
            if c0 != metric_for_donut and df[c0].dtype == "object" and 1 < df[c0].nunique() <= 12:
                cats.append(c0)
        fig = make_donut(df, cats[0], metric_for_donut) if cats and metric_for_donut else None
        if fig:
            st.plotly_chart(fig, use_container_width=True, config={"displayModeBar":False}, key=f"sec_donut_{sheet}")
        else:
            st.info("No meaningful composition segment in this worksheet.")

    # 2) Secondary trend / comparison
    with b:
        st.markdown('<div class="section-title">Secondary Performance <small>best additional metric</small></div>', unsafe_allow_html=True)
        fig = None
        if secondary:
            if date_col:
                fig = make_line(df, date_col, secondary)
                if fig is None:
                    fig = make_hist(df, secondary)
            else:
                cats = [c0 for c0 in df.columns if c0 != secondary and df[c0].dtype == "object" and df[c0].nunique() <= 20]
                if cats:
                    fig = make_bar(df, cats[0], secondary)
                else:
                    fig = make_hist(df, secondary)
        if fig:
            st.plotly_chart(fig, use_container_width=True, config={"displayModeBar":False}, key=f"sec_trend_{sheet}")
        else:
            st.info("No second domain metric is available.")

    # 3) Condition gauge — only if a meaningful condition metric exists
    with c:
        st.markdown('<div class="section-title">Water / Operating Quality <small>condition</small></div>', unsafe_allow_html=True)
        qcol = find_col_by_sem(df, "quality")
        if qcol:
            y = pd.to_numeric(df[qcol], errors="coerce").dropna()
            if len(y):
                fig = make_gauge(float(y.iloc[-1]), chart_title(qcol, "quality"))
                st.plotly_chart(fig, use_container_width=True, config={"displayModeBar":False}, key=f"sec_quality_{sheet}")
            else:
                st.info("No usable condition values.")
        else:
            av = find_col_by_sem(df, "availability")
            if av:
                y = pd.to_numeric(df[av], errors="coerce").dropna()
                if len(y):
                    fig = make_gauge(float(y.iloc[-1]), "Availability")
                    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar":False}, key=f"sec_availability_{sheet}")
                else:
                    st.info("No condition metric.")
            else:
                st.info("No condition metric for this worksheet.")

    # 4) relationship — only domain-compatible pairs
    with d:
        st.markdown('<div class="section-title">Performance Relationship <small>domain-relevant</small></div>', unsafe_allow_html=True)
        pair = None
        semantic_pairs = [
            ("rainfall","level"), ("temperature","efficiency"),
            ("consumption","generation"), ("flow","efficiency"),
            ("availability","generation"), ("downtime","generation"),
            ("flow","recovery"), ("power","generation"),
        ]
        for xsem,ysem in semantic_pairs:
            x = find_col_by_sem(df, xsem)
            y = find_col_by_sem(df, ysem)
            if x and y and x != y:
                pair = (x,y)
                break
        if pair:
            fig = make_scatter(df, pair[0], pair[1])
            if fig:
                st.plotly_chart(fig, use_container_width=True, config={"displayModeBar":False}, key=f"sec_relationship_{sheet}")
            else:
                st.info("Not enough observations for a relationship.")
        else:
            st.info("No defensible domain relationship exists in this worksheet.")


def render_operations(df, primary, secondary, date_col):
    st.markdown('<div class="section-title" style="font-size:16px">Operations at a Glance</div>', unsafe_allow_html=True)
    left, right = st.columns([1.15, 1.0])

    with left:
        st.markdown('<div class="section-title">Trend Detail <small>selected primary metric</small></div>', unsafe_allow_html=True)
        if primary:
            fig = make_line(df, date_col, primary) if date_col else None
            if fig is None:
                fig = make_hist(df, primary)
            if fig:
                st.plotly_chart(fig, use_container_width=True, config={"displayModeBar":False}, key=f"ops_trend_{sheet}")
            else:
                st.info("No usable trend.")
        else:
            st.info("No primary metric.")

    with right:
        st.markdown('<div class="section-title">Business Metric Comparison <small>same worksheet</small></div>', unsafe_allow_html=True)
        nums = numeric_columns(df)
        if len(nums) >= 2:
            # Compare only semantically meaningful pair first.
            x = primary or nums[0]
            y = secondary or nums[1]
            if x != y:
                d = pd.DataFrame({
                    "Metric": [chart_title(x,semantic(x)[0]), chart_title(y,semantic(y)[0])],
                    "Average": [
                        pd.to_numeric(df[x],errors="coerce").mean(),
                        pd.to_numeric(df[y],errors="coerce").mean()
                    ]
                })
                fig = px.bar(d, x="Metric", y="Average", text_auto=".3s", title="Average Metric Comparison")
                st.plotly_chart(base_fig(fig,290), use_container_width=True, config={"displayModeBar":False}, key=f"ops_comparison_{sheet}")
            else:
                st.info("No independent metric pair.")
        else:
            st.info("This worksheet has fewer than two quantitative metrics.")


def render_extra_segments(df, primary, secondary, date_col):
    """
    The remaining data is segmented into additional visual panels.
    Each metric gets a chart type appropriate to its own structure.
    """
    nums = numeric_columns(df)
    used = {primary, secondary}
    extras = [c for c in nums if c not in used]

    if not extras:
        return

    st.markdown('<div class="section-heading">ADDITIONAL MEASUREMENTS <small>all remaining quantitative fields, automatically segmented</small></div>', unsafe_allow_html=True)

    # Up to 6 additional charts per page; no single chart type is forced.
    cols = st.columns(3, gap="large")
    for i, metric in enumerate(extras):
        with cols[i % 3]:
            if i >= 3:
                st.markdown('<div style="height:var(--space-5)"></div>', unsafe_allow_html=True)
            sem = semantic(metric)[0]
            st.markdown(f'<div class="section-title">{chart_title(metric,sem)} <small>{sem}</small></div>', unsafe_allow_html=True)

            # Metric-specific choice:
            if date_col:
                fig = make_line(df, date_col, metric)
                if fig is None:
                    fig = make_hist(df, metric)
            else:
                cats = [c for c in df.columns if c != metric and df[c].dtype == "object" and 1 < df[c].nunique() <= 15]
                if cats:
                    fig = make_bar(df, cats[0], metric)
                else:
                    # Distributions are useful for standalone numeric fields.
                    fig = make_hist(df, metric)

            if fig:
                st.plotly_chart(fig, use_container_width=True, config={"displayModeBar":False}, key=f"extra_{sheet}_{metric}_{i}")



# ============================================================
# Sidebar / workbook
# ============================================================
with st.sidebar:
    st.markdown('<div class="brand" style="font-size:21px">POWER<span class="cyan">ANALYTICS</span></div>', unsafe_allow_html=True)
    st.caption(f"Signed in as {st.session_state.get('username','prahal')}")
    st.markdown('<div style="height:14px"></div>', unsafe_allow_html=True)
    upload = st.file_uploader("UPLOAD EXCEL WORKBOOK", type=["xlsx","xlsm","xls"])
    st.markdown('<div style="height:6px"></div>', unsafe_allow_html=True)

    if st.button("LOG OUT", use_container_width=True):
        st.session_state.authenticated = False
        st.rerun()

if upload is None:
    st.markdown("""
    <div class="power-header">
      <div class="brand">POWER<span class="cyan">ANALYTICS</span></div>
      <div class="subtitle">UTILITY • RENEWABLE ENERGY • PRODUCTION • PERFORMANCE INTELLIGENCE</div>
      <div style="margin-top:22px;color:#2bd4e8;font-size:16px">
        Upload an Excel workbook to generate the graphical management dashboard.
      </div>
    </div>
    """, unsafe_allow_html=True)
    st.stop()

file_bytes = upload.getvalue()
xls = pd.ExcelFile(io.BytesIO(file_bytes))
sheets = xls.sheet_names

if "selected_sheet" not in st.session_state or st.session_state.selected_sheet not in sheets:
    st.session_state.selected_sheet = sheets[0]

with st.sidebar:
    st.markdown("### Workbook Tabs")
    selected = st.selectbox(
        "WORKSHEET",
        sheets,
        index=sheets.index(st.session_state.selected_sheet),
        label_visibility="collapsed"
    )
    st.session_state.selected_sheet = selected

# ============================================================
# Main
# ============================================================
sheet = st.session_state.selected_sheet
df = parse_sheet(xls, sheet)
df = materialize_matrix_dates(df, upload.name)

if df.empty:
    st.error(f"'{sheet}' does not contain usable data.")
    st.stop()

date_col = find_date_column(df)
if date_col:
    parsed = pd.to_datetime(df[date_col], errors="coerce", dayfirst=True)
    df = df.copy()
    df[date_col] = parsed
    df = df.dropna(subset=[date_col], how="all")
    df = df.sort_values(date_col)

usable, primary, secondary = choose_metrics(df)
domain = detect_domain(sheet, list(df.columns))

# Header
st.markdown(f"""
<div class="power-header">
  <div style="display:flex;justify-content:space-between;gap:28px;align-items:center">
    <div>
      <div class="brand">POWER<span class="cyan">ANALYTICS</span></div>
      <div class="subtitle">UTILITY • RENEWABLE ENERGY • PRODUCTION • PERFORMANCE INTELLIGENCE</div>
    </div>
    <div style="text-align:right">
      <div style="font-weight:800">{sheet}</div>
      <div class="subtitle">{st.session_state.get('username','prahal')} • Management Console</div>
    </div>
  </div>
</div>
<div class="meta">
  SOURCE <b>{upload.name}</b>
  &nbsp;&nbsp;•&nbsp;&nbsp; SHEET <b>{sheet}</b>
  &nbsp;&nbsp;•&nbsp;&nbsp; {len(df):,} ROWS
  &nbsp;&nbsp;•&nbsp;&nbsp; {len(df.columns):,} COLUMNS
  &nbsp;&nbsp;•&nbsp;&nbsp; DOMAIN <b>{domain}</b>
</div>
<div class="meta" style="margin-bottom:0">
  <b>WORKSHEET VIEW</b>&nbsp;&nbsp; {sheet}
  &nbsp;&nbsp;•&nbsp;&nbsp;<b>GRAPHICAL ANALYSIS</b>
  &nbsp;&nbsp;•&nbsp;&nbsp;<b>{len(sheets)} TABS</b> IN WORKBOOK
</div>
""", unsafe_allow_html=True)

st.markdown('<div class="section-heading" style="margin-top:var(--space-6)">REPORTING PERIOD</div>', unsafe_allow_html=True)

if date_col:
    dates = pd.to_datetime(df[date_col], errors="coerce").dropna()
    if len(dates):
        min_d, max_d = dates.min(), dates.max()
        st.caption(f"{min_d.strftime('%d %b %Y')} → {max_d.strftime('%d %b %Y')}   •   {len(usable)} quantitative measurements identified")
else:
    st.caption(f"All available records   •   {len(usable)} quantitative measurements identified")

st.markdown('<div style="height:var(--space-3)"></div>', unsafe_allow_html=True)
render_kpis(df, primary, secondary)

st.markdown('<div class="section-heading">OPERATIONS OVERVIEW</div>', unsafe_allow_html=True)
render_main_row(df, domain, primary, secondary, date_col)

st.markdown('<div class="section-heading">PERFORMANCE MIX</div>', unsafe_allow_html=True)
render_secondary_row(df, primary, secondary, date_col)

st.markdown('<div class="section-heading">DETAILED PERFORMANCE</div>', unsafe_allow_html=True)
render_operations(df, primary, secondary, date_col)

render_extra_segments(df, primary, secondary, date_col)


# Footer
st.markdown("""
<hr style="margin-top:var(--space-6)!important">
<div style="text-align:center;color:#607b96;font-size:10.5px;letter-spacing:.3px;padding:var(--space-2) 0 var(--space-3)">
POWERANALYTICS  •  Graphical workbook intelligence  •  Each worksheet analysed independently
</div>
""", unsafe_allow_html=True)