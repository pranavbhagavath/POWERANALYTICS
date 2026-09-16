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
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

:root{
    /* Surfaces — layered from page background up to elevated cards */
    --bg-0:#070f1c; --bg-1:#0a1626;
    --surface:#0e1f34; --surface-2:#122844; --surface-hover:#16304f;

    /* Borders */
    --border:#1c3752; --border-soft:#152a41; --border-strong:#2c5478;

    /* Text */
    --text:#eef4fb; --text-dim:#9fb4cb; --text-faint:#6d869e;

    /* Accents */
    --accent:#2dd4ee; --accent-2:#0ea5c4; --accent-soft:rgba(45,212,238,.12);
    --green:#34e0a1; --orange:#ffb547; --red:#ff6b7a; --purple:#a685ff;

    /* Spacing scale */
    --space-1:4px;  --space-2:8px;  --space-3:12px; --space-4:16px;
    --space-5:20px; --space-6:28px; --space-7:36px; --space-8:48px;

    --radius:12px; --radius-sm:8px; --radius-lg:16px;
    --shadow:0 6px 20px rgba(0,0,0,.22);
    --shadow-lg:0 14px 38px rgba(0,0,0,.28);
}
*{box-sizing:border-box}
html,body,[class*="css"]{
    font-family:'Inter',-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Arial,sans-serif;
}
.stApp{
    background:
        radial-gradient(circle at 85% -5%, rgba(45,212,238,.10), transparent 32%),
        radial-gradient(circle at 0% 100%, rgba(94,64,255,.06), transparent 40%),
        linear-gradient(180deg,#070f1c 0%,#081221 60%,#070f1c 100%);
    color:var(--text);
}
.block-container{
    max-width:1560px;
    padding:var(--space-5) var(--space-5) var(--space-8) var(--space-5);
}
h1,h2,h3,h4,p,div,span,label,button{font-family:'Inter',Arial,sans-serif}
h1,h2,h3,h4{color:var(--text)!important}

/* thin themed scrollbar */
::-webkit-scrollbar{width:9px;height:9px}
::-webkit-scrollbar-track{background:transparent}
::-webkit-scrollbar-thumb{background:#1e3c5a;border-radius:8px}
::-webkit-scrollbar-thumb:hover{background:#2c5478}

/* -----------------------------
   Sidebar
   ----------------------------- */
section[data-testid="stSidebar"]{
    background:linear-gradient(180deg,#050c17 0%,#060f1d 100%);
    border-right:1px solid var(--border-soft);
}
section[data-testid="stSidebar"] > div{padding:var(--space-5) var(--space-4) var(--space-6) var(--space-4)}
section[data-testid="stSidebar"] div[data-testid="stVerticalBlock"]{gap:var(--space-2)}
section[data-testid="stSidebar"] .stButton button{
    min-height:42px;border-radius:var(--radius-sm);
    background:var(--surface);border:1px solid var(--border-strong);
    color:var(--text)!important;font-weight:600;letter-spacing:.2px;
    margin-top:var(--space-2);transition:background .15s ease,border-color .15s ease;
}
section[data-testid="stSidebar"] .stButton button:hover{
    background:var(--surface-hover);border-color:var(--accent-2);
}
section[data-testid="stSidebar"] .stFileUploader > div{
    border:1px dashed var(--border-strong);border-radius:var(--radius-sm);background:var(--surface);
}
section[data-testid="stSidebar"] .stSelectbox > div > div{
    background:var(--surface)!important;border:1px solid var(--border-strong)!important;
    border-radius:var(--radius-sm)!important;color:var(--text)!important;
}
section[data-testid="stSidebar"] h3{
    margin:var(--space-6) 0 var(--space-2) 0!important;
    font-size:11px!important;letter-spacing:1px;color:var(--text-faint)!important;
    text-transform:uppercase;font-weight:700!important;
}
.sb-divider{height:1px;background:var(--border-soft);margin:var(--space-4) 0}
.sb-eyebrow{
    font-size:9.5px;letter-spacing:1.2px;color:var(--text-faint);
    font-weight:700;text-transform:uppercase;margin-bottom:2px;
}

/* -----------------------------
   Top navigation
   ----------------------------- */
.nav-bar{
    display:flex;justify-content:space-between;align-items:center;
    padding:var(--space-4) 0;border-bottom:1px solid var(--border-soft);
    margin-bottom:var(--space-5);
}
.brand{font-size:22px;line-height:1;font-weight:800;letter-spacing:.2px}
.brand .cyan{color:var(--accent)}
.nav-right{display:flex;align-items:center;gap:var(--space-3)}
.nav-chip{
    background:var(--surface);border:1px solid var(--border);border-radius:999px;
    padding:6px 14px;font-size:11.5px;color:var(--text-dim);font-weight:600;
    display:flex;align-items:center;gap:6px;
}
.nav-chip b{color:var(--text);font-weight:700}
.nav-chip .dot{width:6px;height:6px;border-radius:50%;background:var(--green)}

.meta-row{display:flex;flex-wrap:wrap;gap:var(--space-2);margin:0 0 var(--space-5) 0}
.meta-chip{
    background:var(--surface);border:1px solid var(--border);border-radius:8px;
    padding:7px 12px;font-size:10.5px;color:var(--text-dim);letter-spacing:.2px;
    font-weight:600;
}
.meta-chip b{color:var(--accent);font-weight:700}

/* -----------------------------
   Section headings & titles
   ----------------------------- */
.section-heading{
    display:flex;align-items:center;gap:10px;
    font-size:13px;font-weight:800;letter-spacing:1px;text-transform:uppercase;
    color:var(--text-dim);
    margin:var(--space-7) 0 var(--space-4) 0;
}
.section-heading::before{
    content:"";width:3px;height:15px;border-radius:2px;
    background:linear-gradient(180deg,var(--accent),var(--accent-2));
}
.section-heading small{
    text-transform:none;font-weight:500;letter-spacing:0;color:var(--text-faint);font-size:10.5px;
}
.section-heading:first-of-type{margin-top:var(--space-3)}
.section-title{
    font-size:13px;line-height:1.3;font-weight:700;
    color:var(--text);margin:0 0 var(--space-3) 0;
}
.section-title small{color:var(--text-faint);font-size:10px;font-weight:500;margin-left:var(--space-1)}

/* -----------------------------
   Chart / panel cards
   Streamlit's real bordered container (st.container(border=True))
   is restyled here so every chart sits inside a genuine card.
   ----------------------------- */
div[data-testid="stVerticalBlockBorderWrapper"]:has(div[data-testid="stPlotlyChart"]),
div[data-testid="stVerticalBlockBorderWrapper"].chart-card-wrap{
    background:var(--surface)!important;
    border:1px solid var(--border)!important;
    border-radius:var(--radius)!important;
    box-shadow:var(--shadow);
    padding:var(--space-4) var(--space-4) var(--space-2) var(--space-4)!important;
    transition:border-color .15s ease;
}
div[data-testid="stVerticalBlockBorderWrapper"]:has(div[data-testid="stPlotlyChart"]):hover{
    border-color:var(--border-strong)!important;
}

/* -----------------------------
   KPI cards
   ----------------------------- */
.kpi{
    position:relative;overflow:hidden;
    background:var(--surface);
    border:1px solid var(--border);border-radius:var(--radius);min-height:118px;
    padding:var(--space-4) var(--space-5);box-sizing:border-box;
    box-shadow:var(--shadow);
}
.kpi::before{
    content:"";position:absolute;left:0;top:0;bottom:0;width:3px;
    background:linear-gradient(180deg,var(--accent),var(--accent-2));
}
.kpi .label{
    color:var(--text-faint);font-size:9.5px;line-height:1.3;
    font-weight:700;letter-spacing:.9px;text-transform:uppercase;
}
.kpi .value{color:var(--text);font-size:27px;line-height:1.15;font-weight:800;margin-top:var(--space-3);
    font-variant-numeric:tabular-nums;}
.kpi .sub{color:var(--text-faint);font-size:9.5px;line-height:1.5;margin-top:var(--space-2)}

/* -----------------------------
   Snapshot list items
   ----------------------------- */
.snapshot-item{
    background:var(--surface-2);border:1px solid var(--border);border-radius:var(--radius-sm);
    padding:var(--space-3) var(--space-4);margin-bottom:var(--space-3);
}
.snapshot-label{color:var(--text-faint);font-size:9px;font-weight:700;letter-spacing:.6px;text-transform:uppercase}
.snapshot-value{color:var(--text);font-size:22px;font-weight:800;line-height:1.2;margin-top:var(--space-2);
    font-variant-numeric:tabular-nums;}

/* -----------------------------
   Login form
   ----------------------------- */
div[data-testid="stForm"]{
    background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);
    padding:22px 24px;box-shadow:var(--shadow-lg);
}
div[data-testid="stForm"] label{
    color:var(--text-dim)!important;font-size:12px!important;font-weight:600!important;
}
div[data-testid="stForm"] input{
    background:var(--bg-1)!important;border:1px solid var(--border-strong)!important;
    color:var(--text)!important;border-radius:var(--radius-sm)!important;
}
div[data-testid="stForm"] input:focus{
    border-color:var(--accent)!important;box-shadow:0 0 0 1px var(--accent)!important;
}
div[data-testid="stForm"] .stButton button,
div[data-testid="stForm"] button[kind="formSubmit"]{
    background:linear-gradient(100deg,var(--accent-2) 0%,#0d6b85 100%)!important;
    border:1px solid var(--accent)!important;border-radius:var(--radius-sm)!important;
    color:#03151c!important;font-weight:800!important;letter-spacing:.6px!important;
    margin-top:var(--space-1);
}

/* -----------------------------
   Streamlit layout primitives
   ----------------------------- */
div[data-testid="stHorizontalBlock"]{gap:var(--space-5)}
div[data-testid="stVerticalBlock"]{gap:var(--space-2)}
div[data-testid="stMetric"]{background:transparent}
div[data-testid="stPlotlyChart"]{margin:0!important}
div[data-testid="stPlotlyChart"] > div{margin:0!important}
hr{border-color:var(--border-soft)!important;margin:var(--space-4) 0!important}
.stCaption{color:var(--text-faint)!important;margin-bottom:var(--space-2)!important;font-size:11.5px!important}
.stButton button{border-radius:var(--radius-sm)}
.stAlert{
    background:var(--surface)!important;border:1px solid var(--border)!important;
    border-radius:var(--radius-sm)!important;
    padding:var(--space-3) var(--space-4)!important;font-size:12.5px!important;
}

@media (max-width:900px){
    .block-container{padding:var(--space-4) var(--space-3) var(--space-6) var(--space-3)}
    .nav-bar{flex-direction:column;align-items:flex-start;gap:var(--space-3)}
    .brand{font-size:20px}
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
    _, mid, _ = st.columns([1, 1.05, 1])
    with mid:
        st.markdown("""
        <div style="margin:14vh 0 18px 0;
             background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);padding:22px 24px;
             box-shadow:var(--shadow-lg)">
          <div class="brand" style="font-size:21px">POWER<span class="cyan">ANALYTICS</span></div>
          <div style="color:var(--text-faint);font-size:10px;line-height:1.5;margin-top:8px;letter-spacing:.4px">UTILITY • RENEWABLE ENERGY • PRODUCTION • PERFORMANCE INTELLIGENCE</div>
        </div>
        """, unsafe_allow_html=True)
        with st.form("login"):
            u = st.text_input("Username")
            p = st.text_input("Password", type="password")
            st.markdown('<div style="height:2px"></div>', unsafe_allow_html=True)
            if st.form_submit_button("SIGN IN", use_container_width=True):
                if USERS.get(u) == p:
                    st.session_state.authenticated = True
                    st.session_state.username = u
                    st.rerun()
                else:
                    st.error("Invalid username or password.")
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
        margin=dict(l=38, r=12, t=14, b=34),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(color="#dce9f7", size=11, family="Inter, Arial, sans-serif"),
        legend=dict(orientation="h", y=1.1, x=0, font=dict(size=10.5)),
        xaxis=dict(gridcolor="#173b5a", zerolinecolor="#173b5a"),
        yaxis=dict(gridcolor="#173b5a", zerolinecolor="#173b5a"),
        colorway=["#2dd4ee","#a685ff","#34e0a1","#ffb547","#ff6b7a","#5eead4","#f472b6","#93c5fd"],
        hoverlabel=dict(bgcolor="#0e2038", bordercolor="#2c5478",
                        font=dict(color="#eef4fb", size=11, family="Inter, Arial, sans-serif")),
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
    fig = px.line(d, x="Date", y="Value", markers=True)
    fig.update_traces(line=dict(width=2.4, color="#2dd4ee"), marker=dict(size=5))
    return base_fig(fig, 300)


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
    fig = px.bar(d, x=cat_col, y=metric, text_auto=".2s")
    fig.update_traces(marker_color="#2dd4ee")
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
    fig = px.pie(d, names=cat_col, values=metric, hole=.62)
    fig.update_traces(textposition="inside", textinfo="percent",
                      marker=dict(line=dict(color="#0e1f34", width=2)))
    return base_fig(fig, 300)


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
    fig = px.scatter(d, x=xcol, y=ycol, trendline=None)
    fig.update_traces(marker=dict(color="#a685ff", size=7, opacity=.85))
    return base_fig(fig, 300)


def make_hist(df, metric):
    y = pd.to_numeric(df[metric], errors="coerce").dropna()
    if len(y) < 8 or y.nunique() < 4:
        return None
    tmp = pd.DataFrame({"Value": y})
    fig = px.histogram(tmp, x="Value", nbins=min(20, max(6, int(math.sqrt(len(y))))))
    fig.update_traces(marker_color="#2dd4ee")
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
        number={"font": {"size": 30, "color": "#eef4fb"}},
        gauge={
            "axis": {"range": [minimum, maximum], "tickcolor": "#6d869e", "tickfont":{"size":9}},
            "bar": {"color": "#2dd4ee", "thickness": .82},
            "bgcolor": "#0a1626",
            "bordercolor": "#1c3752",
            "borderwidth": 1,
            "steps": [
                {"range": [minimum, minimum + (maximum-minimum)*.65], "color": "#0e2038"},
                {"range": [minimum + (maximum-minimum)*.65, maximum], "color": "#122844"},
            ],
        }
    ))
    return base_fig(fig, 240)


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
        with st.container(border=True):
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
        with st.container(border=True):
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
        with st.container(border=True):
            st.markdown('<div class="section-title">Target Achievement <small>actual vs target</small></div>', unsafe_allow_html=True)
            tp = target_pct(df)
            if tp:
                pct, actual, target, a, t = tp
                fig = go.Figure(go.Indicator(
                    mode="number+gauge",
                    value=pct,
                    number={"suffix":"%", "font":{"size":30,"color":"#eef4fb"}},
                    gauge={
                        "axis":{"range":[0, max(100, min(150, pct*1.25))], "tickcolor":"#6d869e","tickfont":{"size":9}},
                        "bar":{"color":"#34e0a1", "thickness": .82},
                        "bgcolor":"#0a1626",
                        "bordercolor":"#1c3752",
                        "borderwidth": 1,
                        "steps":[
                            {"range":[0,70],"color":"#0e2038"},
                            {"range":[70,90],"color":"#123b31"},
                            {"range":[90,150],"color":"#123b31"},
                        ],
                    }
                ))
                st.plotly_chart(base_fig(fig,240), use_container_width=True, config={"displayModeBar":False}, key=f"main_target_{sheet}")
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
        with st.container(border=True):
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
        with st.container(border=True):
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
        with st.container(border=True):
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
        with st.container(border=True):
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
        with st.container(border=True):
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
        with st.container(border=True):
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
                    fig = px.bar(d, x="Metric", y="Average", text_auto=".3s")
                    fig.update_traces(marker_color=["#2dd4ee", "#a685ff"])
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
            with st.container(border=True):
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
    st.markdown('<div class="brand" style="font-size:20px">POWER<span class="cyan">ANALYTICS</span></div>', unsafe_allow_html=True)
    st.markdown(f'<div style="color:var(--text-faint);font-size:11.5px;margin-top:2px">Signed in as <b style="color:var(--text-dim)">{st.session_state.get("username","prahal")}</b></div>', unsafe_allow_html=True)
    st.markdown('<div class="sb-divider"></div>', unsafe_allow_html=True)
    st.markdown('<div class="sb-eyebrow">Workbook</div>', unsafe_allow_html=True)
    upload = st.file_uploader("UPLOAD EXCEL WORKBOOK", type=["xlsx","xlsm","xls"], label_visibility="collapsed")

    if st.button("LOG OUT", use_container_width=True):
        st.session_state.authenticated = False
        st.rerun()

if upload is None:
    st.markdown("""
    <div style="margin-top:8vh;text-align:center">
      <div class="brand" style="font-size:30px;justify-content:center">POWER<span class="cyan">ANALYTICS</span></div>
      <div style="color:var(--text-faint);font-size:11.5px;margin-top:10px;letter-spacing:.4px">
        UTILITY • RENEWABLE ENERGY • PRODUCTION • PERFORMANCE INTELLIGENCE
      </div>
      <div style="max-width:420px;margin:28px auto 0 auto;background:var(--surface);border:1px solid var(--border);
           border-radius:var(--radius);padding:20px 22px;color:var(--text-dim);font-size:13px;box-shadow:var(--shadow)">
        Upload an Excel workbook from the sidebar to generate the graphical management dashboard.
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
    st.markdown('<div class="sb-divider"></div>', unsafe_allow_html=True)
    st.markdown("### Worksheet")
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
<div class="nav-bar">
  <div>
    <div class="brand">POWER<span class="cyan">ANALYTICS</span></div>
    <div style="color:var(--text-faint);font-size:10.5px;margin-top:4px;letter-spacing:.4px">
      UTILITY • RENEWABLE ENERGY • PRODUCTION • PERFORMANCE INTELLIGENCE
    </div>
  </div>
  <div class="nav-right">
    <div class="nav-chip"><span class="dot"></span> {domain}</div>
    <div class="nav-chip">{st.session_state.get('username','prahal')} • Management Console</div>
  </div>
</div>
<div class="meta-row">
  <div class="meta-chip">SOURCE&nbsp; <b>{upload.name}</b></div>
  <div class="meta-chip">SHEET&nbsp; <b>{sheet}</b></div>
  <div class="meta-chip">{len(df):,} <b>ROWS</b></div>
  <div class="meta-chip">{len(df.columns):,} <b>COLUMNS</b></div>
  <div class="meta-chip">{len(sheets)} <b>TABS IN WORKBOOK</b></div>
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
<div style="text-align:center;color:var(--text-faint);font-size:10.5px;letter-spacing:.3px;padding:var(--space-2) 0 var(--space-3)">
POWERANALYTICS  •  Graphical workbook intelligence  •  Each worksheet analysed independently
</div>
""", unsafe_allow_html=True)