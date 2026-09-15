# -*- coding: utf-8 -*-
"""
POWERANALYTICS
Enterprise Utility, Renewable Energy & Generation Performance Intelligence
---------------------------------------------------------------------------
A Streamlit dashboard that connects directly to a company Excel workbook
(Daily Report) and turns every worksheet into a professional, Power-BI style
analytics view. The workbook is the single source of truth - nothing is
fabricated. If a metric cannot be found in the uploaded file the app shows
"-" and a short note instead of inventing a number.

Run with:
    streamlit run POWERANALYTICS_EXCEL_CONNECTED.py
"""

import io
import re
from datetime import datetime, date

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import streamlit as st

# =============================================================================
# 0. PAGE CONFIG
# =============================================================================
st.set_page_config(
    page_title="POWERANALYTICS",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded",
)

# =============================================================================
# 1. CREDENTIALS (prototype only - move to st.secrets for production)
# =============================================================================
USERS = {"prahal": "12345678"}

# =============================================================================
# 2. CORPORATE THEME / CSS
# =============================================================================
NAVY = "#0B2545"
NAVY_DARK = "#081A33"
TEAL = "#0E8C8C"
TEAL_DARK = "#0B6E6E"
BG = "#F2F4F7"
CARD = "#FFFFFF"
BORDER = "#E1E5EA"
TEXT_DARK = "#101828"
TEXT_MUTED = "#5B6472"
GREEN = "#12805C"
GREEN_BG = "#E7F6EF"
RED = "#B3261E"
RED_BG = "#FBEAE9"

CUSTOM_CSS = f"""
<style>
    #MainMenu {{visibility: hidden;}}
    footer {{visibility: hidden;}}
    header[data-testid="stHeader"] {{
        background: transparent;
        box-shadow: none;
        height: 2.75rem;
    }}
    /* Keep the sidebar open/close arrow visible and easy to see on the light background */
    [data-testid="collapsedControl"], [data-testid="stSidebarCollapseButton"],
    button[title="Open sidebar"], button[title="Close sidebar"] {{
        visibility: visible !important;
        opacity: 1 !important;
        z-index: 999999 !important;
    }}
    [data-testid="collapsedControl"] svg, [data-testid="stSidebarCollapseButton"] svg {{
        color: {NAVY} !important;
        fill: {NAVY} !important;
    }}

    .stApp {{
        background-color: {BG};
    }}
    section[data-testid="stSidebar"] {{
        background-color: {NAVY_DARK};
    }}
    section[data-testid="stSidebar"] * {{
        color: #E7ECF3 !important;
    }}
    section[data-testid="stSidebar"] .stSelectbox label,
    section[data-testid="stSidebar"] .stRadio label {{
        color: #AEB9CB !important;
        font-weight: 600;
        font-size: 0.78rem;
        text-transform: uppercase;
        letter-spacing: .04em;
    }}
    div.block-container {{
        padding-top: 0.6rem;
        padding-bottom: 2rem;
        max-width: 1400px;
    }}

    /* Corporate header banner */
    .pa-header {{
        background: linear-gradient(90deg, {NAVY} 0%, {NAVY_DARK} 100%);
        border-radius: 10px;
        padding: 18px 28px;
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-bottom: 6px;
        box-shadow: 0 2px 10px rgba(0,0,0,0.15);
    }}
    .pa-header .pa-title {{
        color: #FFFFFF;
        font-size: 1.55rem;
        font-weight: 800;
        letter-spacing: .03em;
        margin: 0;
    }}
    .pa-header .pa-sub {{
        color: #B9E3E1;
        font-size: 0.82rem;
        font-weight: 500;
        margin-top: 2px;
    }}
    .pa-header .pa-user {{
        color: #E7ECF3;
        font-size: 0.85rem;
        text-align: right;
    }}
    .pa-ribbon {{
        background: {TEAL};
        color: #ffffff;
        border-radius: 8px;
        padding: 8px 22px;
        font-size: 0.82rem;
        font-weight: 600;
        margin-bottom: 18px;
        display: flex;
        justify-content: space-between;
    }}

    /* Section headers */
    .pa-section-title {{
        font-size: 1.05rem;
        font-weight: 700;
        color: {TEXT_DARK};
        border-left: 5px solid {TEAL};
        padding-left: 10px;
        margin: 22px 0 12px 0;
    }}

    /* KPI Cards */
    .pa-kpi {{
        background: {CARD};
        border: 1px solid {BORDER};
        border-radius: 10px;
        padding: 14px 16px 12px 16px;
        box-shadow: 0 1px 3px rgba(16,24,40,0.06);
        height: 100%;
    }}
    .pa-kpi-label {{
        font-size: 0.72rem;
        font-weight: 700;
        color: {TEXT_MUTED};
        text-transform: uppercase;
        letter-spacing: .04em;
        margin-bottom: 6px;
    }}
    .pa-kpi-value {{
        font-size: 1.5rem;
        font-weight: 800;
        color: {TEXT_DARK};
        line-height: 1.1;
    }}
    .pa-kpi-sub {{
        font-size: 0.74rem;
        color: {TEXT_MUTED};
        margin-top: 4px;
    }}
    .pa-kpi-na {{
        font-size: 1.3rem;
        font-weight: 800;
        color: #B7BEC9;
    }}

    .pa-status-green {{
        background: {GREEN_BG};
        border: 1px solid #B8E3CF;
        color: {GREEN};
        border-radius: 10px;
        padding: 14px 16px;
        font-weight: 800;
        text-align:center;
    }}
    .pa-status-red {{
        background: {RED_BG};
        border: 1px solid #F2C3C0;
        color: {RED};
        border-radius: 10px;
        padding: 14px 16px;
        font-weight: 800;
        text-align:center;
    }}
    .pa-status-title {{font-size:0.7rem; text-transform:uppercase; letter-spacing:.05em; font-weight:700; opacity:0.8;}}
    .pa-status-value {{font-size:1.35rem; margin-top:4px;}}

    .pa-note {{
        font-size: 0.78rem;
        color: {TEXT_MUTED};
        background: #F7F9FB;
        border: 1px dashed {BORDER};
        border-radius: 8px;
        padding: 8px 12px;
        margin: 6px 0 14px 0;
    }}

    .pa-welcome {{
        background: linear-gradient(135deg, {NAVY} 0%, {TEAL_DARK} 120%);
        border-radius: 16px;
        padding: 60px 50px;
        text-align: center;
        color: white;
        margin-top: 40px;
    }}
    .pa-welcome h1 {{
        font-size: 2.4rem;
        font-weight: 900;
        letter-spacing: .04em;
        margin-bottom: 6px;
    }}
    .pa-welcome p {{
        font-size: 1.05rem;
        color: #CFEFEC;
        font-weight: 400;
    }}

    div[data-testid="stMetric"] {{
        background: {CARD};
        border: 1px solid {BORDER};
        border-radius: 10px;
        padding: 10px 14px;
    }}
    .stTabs [data-baseweb="tab-list"] {{ gap: 4px; }}
    .stTabs [data-baseweb="tab"] {{
        background: #E9EDF2;
        border-radius: 8px 8px 0 0;
        padding: 8px 16px;
        font-weight: 600;
        color: {TEXT_DARK};
    }}
    .stTabs [aria-selected="true"] {{
        background: {TEAL} !important;
        color: white !important;
    }}
    thead tr th {{ background-color: {NAVY} !important; color: white !important; }}
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

CHART_TEMPLATE = "plotly_white"
CHART_FONT = dict(family="Segoe UI, Arial", size=12, color=TEXT_DARK)
PALETTE = [TEAL, NAVY, "#E38B29", "#7A5CC0", "#C0447A", "#3E8E7E", "#B3261E", "#5B6472"]


def style_fig(fig, title=None, y_title=None, x_title=None, height=380):
    fig.update_layout(
        template=CHART_TEMPLATE,
        font=CHART_FONT,
        title=dict(text=title, font=dict(size=15, color=TEXT_DARK, family="Segoe UI, Arial")) if title else None,
        plot_bgcolor="white",
        paper_bgcolor="white",
        height=height,
        margin=dict(l=10, r=10, t=50 if title else 20, b=10),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1, font=dict(size=11)),
        hoverlabel=dict(bgcolor="white", font_size=12),
    )
    if y_title:
        fig.update_yaxes(title_text=y_title, gridcolor="#EEF1F5", zeroline=False)
    else:
        fig.update_yaxes(gridcolor="#EEF1F5", zeroline=False)
    if x_title:
        fig.update_xaxes(title_text=x_title)
    return fig


_chart_counter = {"n": 0}


def chart_key(prefix):
    _chart_counter["n"] += 1
    return f"{prefix}_{_chart_counter['n']}"


# =============================================================================
# 3. AUTH
# =============================================================================
def login_gate() -> bool:
    if st.session_state.get("authenticated"):
        return True

    left, mid, right = st.columns([1, 1.1, 1])
    with mid:
        st.markdown(
            f"""
            <div style="text-align:center; margin-top:60px; margin-bottom: 10px;">
                <div style="font-size:2rem;">⚡</div>
                <div style="font-size:1.7rem; font-weight:900; color:{NAVY}; letter-spacing:.04em;">POWERANALYTICS</div>
                <div style="font-size:0.85rem; color:{TEXT_MUTED}; margin-top:4px;">
                    Enterprise Utility, Renewable Energy &amp; Generation Performance Intelligence
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        with st.form("login_form", clear_on_submit=False):
            st.markdown("##### Secure Sign-In")
            username = st.text_input("Username")
            password = st.text_input("Password", type="password")
            submitted = st.form_submit_button("Sign In", width='stretch')
            if submitted:
                if USERS.get(username) == password:
                    st.session_state["authenticated"] = True
                    st.session_state["username"] = username
                    st.rerun()
                else:
                    st.error("Invalid username or password.")
        st.caption("Prototype credentials are managed in-app. Production deployments should use Streamlit secrets.")
    return False


# =============================================================================
# 4. WORKBOOK LOADING (fast metadata + lazy, cached sheet parsing)
# =============================================================================
@st.cache_data(show_spinner=False)
def get_sheet_names(file_bytes: bytes):
    import openpyxl
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
    names = list(wb.sheetnames)
    wb.close()
    return names


@st.cache_data(show_spinner=False)
def load_raw_sheet(file_bytes: bytes, sheet_name: str):
    """Return the sheet as a list-of-lists of raw cell values (uniform width)."""
    import openpyxl
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb[sheet_name]
    rows = [list(r) for r in ws.iter_rows(values_only=True)]
    wb.close()
    if not rows:
        return []
    ncols = max(len(r) for r in rows)
    rows = [r + [None] * (ncols - len(r)) for r in rows]

    # Trim trailing columns/rows that are entirely empty. Some worksheets report
    # an inflated used-range (e.g. formatting applied across full Excel columns)
    # which would otherwise explode the parsed table to tens of thousands of columns.
    last_col = -1
    for r in rows:
        for c in range(len(r) - 1, last_col, -1):
            if r[c] is not None:
                last_col = max(last_col, c)
                break
    if last_col < 0:
        return []
    rows = [r[: last_col + 1] for r in rows]

    last_row = -1
    for i in range(len(rows) - 1, -1, -1):
        if any(v is not None for v in rows[i]):
            last_row = i
            break
    rows = rows[: last_row + 1]
    return rows


# =============================================================================
# 5. GENERIC PARSING ENGINE
# =============================================================================
_MONTHS = "Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec"
_DATE_STR_RE = re.compile(
    rf"(\d{{1,2}}[-/.]\d{{1,2}}[-/.]\d{{2,4}})|(\d{{1,2}}[-\s]?({_MONTHS})\b)|(\b({_MONTHS})[-\s]?\d{{1,2}}\b)",
    re.IGNORECASE,
)


def _looks_datey(v):
    if isinstance(v, (datetime, date)):
        return True
    if isinstance(v, str):
        return bool(_DATE_STR_RE.search(v.strip()))
    return False


def _clean_label(v, fallback="Column"):
    if v is None:
        return fallback
    s = re.sub(r"\s+", " ", str(v)).strip()
    return s if s else fallback


def _dedupe_columns(cols):
    seen, out = {}, []
    for c in cols:
        c = _clean_label(c)
        if c in seen:
            seen[c] += 1
            out.append(f"{c} ({seen[c]})")
        else:
            seen[c] = 1
            out.append(c)
    return out


def find_header_row(rows, max_scan=12):
    best_idx, best_score = 0, -1
    for i in range(min(max_scan, len(rows))):
        score = sum(1 for v in rows[i] if isinstance(v, str) and v.strip())
        if score > best_score:
            best_score = score
            best_idx = i
    return best_idx


def find_best_date_row(rows, center, window=3):
    """Look near the text-header row for the row carrying the most real datetime cells
    (handles sheets where dates sit on a different row than the text sub-header)."""
    lo, hi = max(0, center - window), min(len(rows), center + window + 1)
    best_i, best_n = None, 0
    for i in range(lo, hi):
        n = sum(1 for v in rows[i] if isinstance(v, (datetime, date)))
        if n > best_n:
            best_n = n
            best_i = i
    return best_i, best_n


def detect_date_col(rows, header_idx, ncols, sample=100):
    data = rows[header_idx + 1: header_idx + 1 + sample]
    best_col, best_frac = None, 0
    for c in range(min(ncols, 4)):
        vals = [r[c] for r in data if c < len(r) and r[c] is not None]
        if len(vals) < 3:
            continue
        datey = [v for v in vals if _looks_datey(v)]
        frac = len(datey) / len(vals)
        uniq = len(set(str(v) for v in datey))
        if frac > 0.4 and uniq >= 3 and frac > best_frac:
            best_frac = frac
            best_col = c
    return best_col


def _to_numeric(series):
    return pd.to_numeric(series, errors="coerce")


class ParsedSheet:
    """Container describing how a worksheet was interpreted."""

    def __init__(self, mode, df=None, note=None, id_cols=None, value_col=None):
        self.mode = mode  # "long" | "wide" | "categorical" | "empty"
        self.df = df
        self.note = note
        self.id_cols = id_cols or []
        self.value_col = value_col


def parse_sheet_generic(file_bytes: bytes, sheet_name: str) -> ParsedSheet:
    rows = load_raw_sheet(file_bytes, sheet_name)
    if not rows or all(all(v is None for v in r) for r in rows):
        return ParsedSheet("empty", note="This worksheet has no readable data.")

    ncols = max(len(r) for r in rows)
    header_idx = find_header_row(rows)

    # ---- Wide date-table check (dates spread across a header row) ----
    date_row_idx, n_date_cells = find_best_date_row(rows, header_idx, window=3)
    if n_date_cells >= 3:
        date_row = rows[date_row_idx]
        date_cols = [c for c in range(ncols) if c < len(date_row) and isinstance(date_row[c], (datetime, date))]
        id_col_count = min(date_cols) if date_cols else 0
        label_row = rows[header_idx] if header_idx != date_row_idx else rows[max(0, date_row_idx - 1)]
        sub_row = rows[date_row_idx + 1] if date_row_idx + 1 < len(rows) else [None] * ncols
        id_names = []
        for c in range(id_col_count):
            name = label_row[c] if c < len(label_row) and label_row[c] not in (None, "") else None
            if name is None and c < len(sub_row):
                name = sub_row[c]
            id_names.append(_clean_label(name, f"Col{c+1}"))
        id_names = _dedupe_columns(id_names)

        data_start = max(header_idx, date_row_idx) + 2
        records = []
        for r in rows[data_start:]:
            if all(v is None for v in r):
                continue
            id_vals = r[:id_col_count]
            if all(v is None for v in id_vals):
                continue
            base = dict(zip(id_names, id_vals))
            for c in date_cols:
                val = r[c] if c < len(r) else None
                if isinstance(val, (int, float)):
                    rec = dict(base)
                    rec["Date"] = pd.to_datetime(date_row[c], errors="coerce")
                    rec["Value"] = val
                    records.append(rec)
        if records:
            df = pd.DataFrame(records)
            df = df.dropna(subset=["Date"])
            return ParsedSheet("wide", df=df, id_cols=id_names, value_col="Value")

    # ---- Long format (date column present) ----
    date_col = detect_date_col(rows, header_idx, ncols)
    if date_col is not None:
        header = [rows[header_idx][c] if c < len(rows[header_idx]) else None for c in range(ncols)]
        cols = _dedupe_columns(header)
        data = rows[header_idx + 1:]
        df = pd.DataFrame(data, columns=cols)
        date_col_name = cols[date_col]
        df["Date"] = pd.to_datetime(df[date_col_name], errors="coerce", dayfirst=True)
        if date_col_name != "Date":
            df = df.drop(columns=[date_col_name])
        df = df.dropna(subset=["Date"]).reset_index(drop=True)
        num_cols = []
        for c in df.columns:
            if c == "Date":
                continue
            conv = _to_numeric(df[c])
            if conv.notna().sum() >= max(1, int(0.3 * len(df))):
                df[c] = conv
                num_cols.append(c)
            else:
                df = df.drop(columns=[c])
        df = df.dropna(axis=1, how="all")
        if not num_cols:
            return ParsedSheet("empty", note="Dates were found, but no numeric measures could be identified.")
        df = df.sort_values("Date").reset_index(drop=True)
        return ParsedSheet("long", df=df)

    # ---- Categorical fallback (no time dimension detected) ----
    header = [rows[header_idx][c] if c < len(rows[header_idx]) else None for c in range(ncols)]
    cols = _dedupe_columns(header)
    data = rows[header_idx + 1:]
    df = pd.DataFrame(data, columns=cols)
    df = df.dropna(axis=0, how="all").reset_index(drop=True)
    num_cols, label_cols = [], []
    for c in df.columns:
        conv = _to_numeric(df[c])
        if conv.notna().sum() >= max(1, int(0.3 * len(df))):
            df[c] = conv
            num_cols.append(c)
        elif df[c].notna().sum() >= max(1, int(0.3 * len(df))) and len(label_cols) < 3:
            label_cols.append(c)
    if not num_cols:
        return ParsedSheet("empty", note="No structured numeric data could be identified on this worksheet.")
    df = df[label_cols + num_cols].dropna(axis=1, how="all")
    label_cols = [c for c in label_cols if c in df.columns]
    num_cols = [c for c in num_cols if c in df.columns]
    return ParsedSheet("categorical", df=df, id_cols=label_cols, value_col=num_cols)


def find_plan_actual_pairs(columns):
    """Match e.g. 'Plan' <-> 'Actual', 'Target ...' <-> 'Actual ...' by shared context words."""
    plans = [c for c in columns if re.search(r"\bplan\b|\btarget\b", c, re.I)]
    actuals = [c for c in columns if re.search(r"\bactual\b", c, re.I)]
    pairs = []
    for p in plans:
        best, best_score = None, 0
        p_tokens = set(re.findall(r"[a-zA-Z]+", p.lower()))
        for a in actuals:
            a_tokens = set(re.findall(r"[a-zA-Z]+", a.lower()))
            score = len(p_tokens & a_tokens)
            if score > best_score:
                best_score, best = score, a
        if best:
            pairs.append((p, best))
    return pairs


def apply_grain(df: pd.DataFrame, grain: str, numeric_cols) -> pd.DataFrame:
    d = df.sort_values("Date").copy()
    if grain == "Day-wise" or d.empty:
        return d
    freq = "W" if grain == "Week-wise" else "M"
    agg = d.set_index("Date")[numeric_cols].resample(freq).sum(min_count=1).reset_index()
    return agg


def filter_period(df: pd.DataFrame, period_label: str) -> pd.DataFrame:
    if period_label in (None, "All available data") or "Date" not in df.columns:
        return df
    try:
        per = pd.Period(period_label, freq="M")
    except Exception:
        return df
    return df[df["Date"].dt.to_period("M") == per]


def top_numeric_columns(df, n=4, exclude=()):
    numeric_cols = [c for c in df.columns if c != "Date" and c not in exclude and pd.api.types.is_numeric_dtype(df[c])]
    sums = {c: df[c].abs().sum() for c in numeric_cols}
    ranked = sorted(sums, key=sums.get, reverse=True)
    return ranked[:n]


# =============================================================================
# 6. UI HELPERS
# =============================================================================
def kpi_card(col, label, value, sub=None, unit=""):
    with col:
        if value is None or (isinstance(value, float) and np.isnan(value)):
            st.markdown(
                f"""<div class="pa-kpi"><div class="pa-kpi-label">{label}</div>
                <div class="pa-kpi-na">—</div>
                <div class="pa-kpi-sub">Source metric not available for this period</div></div>""",
                unsafe_allow_html=True,
            )
            return
        if isinstance(value, (int, float, np.floating, np.integer)):
            disp = f"{value:,.1f}" if abs(value) < 1000 else f"{value:,.0f}"
        else:
            disp = str(value)
        st.markdown(
            f"""<div class="pa-kpi"><div class="pa-kpi-label">{label}</div>
            <div class="pa-kpi-value">{disp}{unit}</div>
            <div class="pa-kpi-sub">{sub or ""}</div></div>""",
            unsafe_allow_html=True,
        )


def status_card(col, label, is_green, value_text):
    with col:
        cls = "pa-status-green" if is_green else "pa-status-red"
        dot = "🟢" if is_green else "🔴"
        st.markdown(
            f"""<div class="{cls}"><div class="pa-status-title">{label}</div>
            <div class="pa-status-value">{dot} {value_text}</div></div>""",
            unsafe_allow_html=True,
        )


def section_title(text):
    st.markdown(f'<div class="pa-section-title">{text}</div>', unsafe_allow_html=True)


def note(text):
    st.markdown(f'<div class="pa-note">ℹ️ {text}</div>', unsafe_allow_html=True)


def source_data_expander(df, key):
    with st.expander("📄 View Source Data"):
        st.caption("Values shown exactly as parsed from the uploaded workbook.")
        safe_df = df.copy()
        for c in safe_df.columns:
            if safe_df[c].dtype == object:
                safe_df[c] = safe_df[c].apply(lambda v: v if (v is None or isinstance(v, (str, int, float))) else str(v))
                safe_df[c] = safe_df[c].astype(str).where(safe_df[c].notna(), None)
        st.dataframe(safe_df, width='stretch', height=320, key=f"src_{key}")


# =============================================================================
# 7. SPECIALIZED PARSER — SUMMARY SHEET (highest-value management sheet)
# =============================================================================
SUMMARY_MAP = {
    0: "Date", 5: "Power Consumption (kWh)", 6: "Recorded Demand (kVA)", 7: "Power Factor",
    9: "EB Line-123", 10: "EB Line-4", 11: "EB Total",
    12: "DG Line-123", 13: "DG Line-4", 14: "DG Total",
    15: "Solar Parking", 16: "Solar MS", 17: "Solar LOG", 18: "Solar MA",
    19: "Solar Line-4 (EB)", 20: "Solar Line-4 (DG)", 23: "Solar Total",
    24: "Total Power Consumption", 27: "MSB Total", 28: "% Loss",
    49: "Department Total Consumption",
    60: "5.4MW Wind Total (after loss)", 61: "8.1MW Wind Total (after loss)",
    66: "3.7MW Hybrid Gen (after loss)", 67: "Avg Wind Speed", 68: "Wind Speed/Gen Total",
}


@st.cache_data(show_spinner=False)
def parse_summary_sheet(file_bytes: bytes, sheet_name: str):
    rows = load_raw_sheet(file_bytes, sheet_name)
    if not rows:
        return None
    records = []
    for r in rows:
        d = r[0] if len(r) > 0 else None
        if not isinstance(d, (datetime, date)):
            continue
        rec = {}
        for idx, name in SUMMARY_MAP.items():
            rec[name] = r[idx] if idx < len(r) else None
        records.append(rec)
    if not records:
        return None
    df = pd.DataFrame(records)
    for c in df.columns:
        if c == "Date":
            df[c] = pd.to_datetime(df[c], errors="coerce")
        else:
            df[c] = pd.to_numeric(df[c], errors="coerce")
    df = df.dropna(subset=["Date"]).sort_values("Date").reset_index(drop=True)
    return df


# =============================================================================
# 8. SPECIALIZED PARSER — "Consump & Gen. Plan Vs actual" (Plan/Actual/Variance blocks)
# =============================================================================
@st.cache_data(show_spinner=False)
def parse_plan_vs_actual(file_bytes: bytes, sheet_name: str):
    rows = load_raw_sheet(file_bytes, sheet_name)
    if not rows:
        return None
    ncols = max(len(r) for r in rows)
    header_row = rows[0]
    records = []
    date_re = re.compile(r"(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{2,4})")
    for block_start in range(0, ncols, 5):
        header_val = header_row[block_start] if block_start < len(header_row) else None
        if not isinstance(header_val, str):
            continue
        m = date_re.search(header_val)
        if not m:
            continue
        day, mon, yr = m.groups()
        yr = int(yr) if len(yr) == 4 else 2000 + int(yr)
        try:
            block_date = pd.Timestamp(year=yr, month=int(mon), day=int(day))
        except Exception:
            continue
        shift_m = re.search(r"([A-Za-z])\s*shift", header_val, re.I)
        shift = shift_m.group(1).upper() if shift_m else "—"
        for r in rows[2:]:
            metric = r[block_start] if block_start < len(r) else None
            if metric is None or not isinstance(metric, str) or not metric.strip():
                continue
            plan = r[block_start + 1] if block_start + 1 < len(r) else None
            actual = r[block_start + 2] if block_start + 2 < len(r) else None
            if not isinstance(plan, (int, float)) and not isinstance(actual, (int, float)):
                continue
            records.append({
                "Date": block_date, "Shift": shift, "Metric": _clean_label(metric),
                "Plan": plan, "Actual": actual,
            })
    if not records:
        return None
    df = pd.DataFrame(records)
    df["Plan"] = pd.to_numeric(df["Plan"], errors="coerce")
    df["Actual"] = pd.to_numeric(df["Actual"], errors="coerce")
    df["Variance"] = df["Actual"] - df["Plan"]
    df["Achievement %"] = np.where(df["Plan"] != 0, df["Actual"] / df["Plan"] * 100, np.nan)
    return df.sort_values("Date").reset_index(drop=True)


# =============================================================================
# 9. DASHBOARD GROUPING
# =============================================================================
def group_sheets(sheet_names):
    groups = {"Executive": [], "Production": [], "Solar": [], "Wind": [],
              "Utility & Cost": [], "Analysis / Other": []}
    exec_keys = ["summary", "morning meeting"]
    util_keys = ["utility", "cost", "fixed power", "grid power", "ot breakup", "shift wise consumption"]
    prod_keys = ["line", "produc", "beki", "shift", "consump", "march month"]
    for name in sheet_names:
        low = name.lower()
        if "solar" in low:
            groups["Solar"].append(name)
        elif "wind" in low:
            groups["Wind"].append(name)
        elif any(k in low for k in exec_keys):
            groups["Executive"].append(name)
        elif any(k in low for k in util_keys):
            groups["Utility & Cost"].append(name)
        elif any(k in low for k in prod_keys):
            groups["Production"].append(name)
        else:
            groups["Analysis / Other"].append(name)
    return {k: v for k, v in groups.items() if v}


# =============================================================================
# 10. GENERIC SHEET RENDERER
# =============================================================================
def render_generic_sheet(file_bytes, sheet_name, grain, period_label, key_prefix):
    parsed = parse_sheet_generic(file_bytes, sheet_name)

    if parsed.mode == "empty":
        note(parsed.note or "No usable data found on this worksheet.")
        return

    if parsed.mode == "long":
        df = parsed.df.copy()
        df_period = filter_period(df, period_label)
        if df_period.empty:
            note("No rows fall inside the selected reporting period for this worksheet — showing full available range instead.")
            df_period = df
        numeric_cols = [c for c in df_period.columns if c != "Date" and pd.api.types.is_numeric_dtype(df_period[c])]
        agg = apply_grain(df_period, grain, numeric_cols)

        pairs = find_plan_actual_pairs(numeric_cols)
        top_cols = top_numeric_columns(df_period, n=4, exclude=[c for p in pairs for c in p])

        section_title("Key Metrics")
        kpi_cols = st.columns(4)
        shown = 0
        if pairs:
            p, a = pairs[0]
            plan_sum, act_sum = df_period[p].sum(), df_period[a].sum()
            kpi_card(kpi_cols[0], f"Total {a}", act_sum)
            kpi_card(kpi_cols[1], f"Total {p}", plan_sum)
            ach = (act_sum / plan_sum * 100) if plan_sum else np.nan
            kpi_card(kpi_cols[2], "Achievement %", ach, unit="%")
            shown = 3
        for c in top_cols:
            if shown >= 4:
                break
            kpi_card(kpi_cols[shown], c, df_period[c].sum(), sub="Sum over selected period")
            shown += 1
        while shown < 4:
            kpi_card(kpi_cols[shown], "—", None)
            shown += 1

        if pairs:
            p, a = pairs[0]
            plan_sum, act_sum = df_period[p].sum(), df_period[a].sum()
            is_green = act_sum >= plan_sum if plan_sum else False
            sc = st.columns([1, 3])
            status_card(sc[0], "Target Status", is_green, "Achieved" if is_green else "Below Target")
            with sc[1]:
                var = act_sum - plan_sum
                st.markdown(
                    f"<div class='pa-note'>Plan: <b>{plan_sum:,.0f}</b> &nbsp;|&nbsp; "
                    f"Actual: <b>{act_sum:,.0f}</b> &nbsp;|&nbsp; Variance: <b>{var:,.0f}</b> "
                    f"&nbsp;|&nbsp; Achievement: <b>{(act_sum/plan_sum*100 if plan_sum else 0):,.1f}%</b></div>",
                    unsafe_allow_html=True,
                )

        section_title(f"{grain} Trend")
        if not agg.empty:
            if pairs:
                p, a = pairs[0]
                fig = go.Figure()
                fig.add_bar(x=agg["Date"], y=agg[a], name=a, marker_color=TEAL)
                fig.add_scatter(x=agg["Date"], y=agg[p], name=p, mode="lines+markers", line=dict(color=NAVY, width=3))
                st.plotly_chart(style_fig(fig, f"{a} vs {p} ({grain})"), width='stretch', key=chart_key(f"{key_prefix}_pa"))
            elif top_cols:
                fig = go.Figure()
                for i, c in enumerate(top_cols):
                    fig.add_scatter(x=agg["Date"], y=agg[c], name=c, mode="lines+markers",
                                     line=dict(color=PALETTE[i % len(PALETTE)], width=2.5))
                st.plotly_chart(style_fig(fig, f"Trend — {sheet_name} ({grain})"), width='stretch', key=chart_key(f"{key_prefix}_trend"))
        else:
            note("Not enough data points in the selected period to draw a trend chart.")

        source_data_expander(df_period, key_prefix)

    elif parsed.mode == "wide":
        df = parsed.df.copy()
        df = filter_period(df, period_label) if not filter_period(df, period_label).empty else df
        section_title("Key Metrics")
        kc = st.columns(4)
        kpi_card(kc[0], "Total Value", df["Value"].sum())
        kpi_card(kc[1], "Average", df["Value"].mean())
        kpi_card(kc[2], "Records", int(df.shape[0]))
        kpi_card(kc[3], "Date Range", f"{df['Date'].min().date()} → {df['Date'].max().date()}" if not df.empty else None)

        section_title(f"{grain} Trend")
        d2 = df.groupby(pd.Grouper(key="Date", freq={"Day-wise": "D", "Week-wise": "W", "Month-wise": "M"}[grain]))["Value"].sum().reset_index()
        fig = px.bar(d2, x="Date", y="Value", color_discrete_sequence=[TEAL])
        st.plotly_chart(style_fig(fig, f"{sheet_name} — {grain}", y_title="Value"), width='stretch', key=chart_key(f"{key_prefix}_wide_trend"))

        if parsed.id_cols:
            id_col = parsed.id_cols[min(1, len(parsed.id_cols) - 1)]
            if id_col in df.columns:
                section_title("Contribution Breakdown")
                d3 = df.groupby(id_col)["Value"].sum().reset_index().sort_values("Value", ascending=False).head(12)
                fig2 = px.bar(d3, x=id_col, y="Value", color_discrete_sequence=[NAVY])
                st.plotly_chart(style_fig(fig2, f"By {id_col}"), width='stretch', key=chart_key(f"{key_prefix}_wide_contrib"))

        source_data_expander(df, key_prefix)

    elif parsed.mode == "categorical":
        df = parsed.df
        value_cols = parsed.value_col if isinstance(parsed.value_col, list) else [parsed.value_col]
        note("This worksheet does not contain a per-day date series — showing the values as recorded on the sheet.")
        section_title("Key Metrics")
        kc = st.columns(4)
        top = sorted(value_cols, key=lambda c: df[c].abs().sum(), reverse=True)[:4]
        for i, c in enumerate(top):
            kpi_card(kc[i], c, df[c].sum(), sub="Sum of column")
        for i in range(len(top), 4):
            kpi_card(kc[i], "—", None)

        section_title("Composition")
        label_col = parsed.id_cols[0] if parsed.id_cols else df.columns[0]
        if top:
            plot_df = df[[label_col, top[0]]].dropna().sort_values(top[0], ascending=False).head(15)
            fig = px.bar(plot_df, x=label_col, y=top[0], color_discrete_sequence=[TEAL])
            st.plotly_chart(style_fig(fig, f"{top[0]} by {label_col}"), width='stretch', key=chart_key(f"{key_prefix}_cat"))
        source_data_expander(df, key_prefix)


# =============================================================================
# 11. EXECUTIVE OVERVIEW
# =============================================================================
def render_executive_overview(file_bytes, sheet_names, grain, period_label):
    summary_name = next((s for s in sheet_names if s.strip().lower() == "summary"), None)
    plan_name = next((s for s in sheet_names if "plan vs actual" in s.lower()), None)

    sdf = parse_summary_sheet(file_bytes, summary_name) if summary_name else None
    pdf = parse_plan_vs_actual(file_bytes, plan_name) if plan_name else None

    if sdf is None and pdf is None:
        note("Could not locate a 'Summary' or 'Plan Vs actual' worksheet in this workbook — Executive Overview needs at least one of these sheets.")
        return

    if sdf is not None:
        sdf_p = filter_period(sdf, period_label)
        if sdf_p.empty:
            sdf_p = sdf

        section_title("Executive KPIs")
        c = st.columns(4)
        kpi_card(c[0], "Total Plant Consumption", sdf_p.get("Power Consumption (kWh)", pd.Series(dtype=float)).sum(), sub="kWh, meter-based")
        wind_cols = [x for x in ["5.4MW Wind Total (after loss)", "8.1MW Wind Total (after loss)"] if x in sdf_p.columns]
        wind_total = sdf_p[wind_cols].sum().sum() if wind_cols else np.nan
        kpi_card(c[1], "Wind Generation", wind_total, sub="kWh, sum of wind farms")
        solar_total = sdf_p.get("Solar Total", pd.Series(dtype=float)).sum()
        kpi_card(c[2], "Solar Rooftop Generation", solar_total, sub="kWh")
        renewable = (wind_total if not np.isnan(wind_total) else 0) + (solar_total if not np.isnan(solar_total) else 0)
        kpi_card(c[3], "Total Renewable Generation", renewable if renewable else np.nan, sub="Wind + Solar, kWh")

        c2 = st.columns(4)
        kpi_card(c2[0], "Peak Demand", sdf_p.get("Recorded Demand (kVA)", pd.Series(dtype=float)).max(), sub="kVA, max recorded")
        kpi_card(c2[1], "Average Power Factor", sdf_p.get("Power Factor", pd.Series(dtype=float)).mean())
        kpi_card(c2[2], "Operating Days", sdf_p["Date"].nunique(), sub="Days with a recorded reading")
        total_cons = sdf_p.get("Total Power Consumption", pd.Series(dtype=float)).sum()
        kpi_card(c2[3], "Plant Total Power", total_cons, sub="EB + DG + Solar (kWh)")

        if total_cons and renewable:
            section_title("Renewable Contribution")
            pct = renewable / total_cons * 100
            g = go.Figure(go.Indicator(
                mode="gauge+number", value=pct,
                number={"suffix": "%", "font": {"size": 32, "color": TEXT_DARK}},
                gauge={"axis": {"range": [0, 100]}, "bar": {"color": TEAL},
                       "steps": [{"range": [0, 50], "color": "#F1F4F7"}, {"range": [50, 100], "color": "#DFF3EF"}]},
            ))
            st.plotly_chart(style_fig(g, "Renewable Share of Total Plant Power", height=260),
                             width='stretch', key=chart_key("exec_gauge"))

        section_title("Daily Performance")
        cc = st.columns(2)
        with cc[0]:
            fig = go.Figure()
            fig.add_scatter(x=sdf_p["Date"], y=sdf_p.get("Power Consumption (kWh)"), name="Consumption",
                             mode="lines+markers", line=dict(color=NAVY, width=2.5))
            st.plotly_chart(style_fig(fig, "Daily Plant Consumption", y_title="kWh"), width='stretch', key=chart_key("exec_cons"))
        with cc[1]:
            fig2 = go.Figure()
            for name, color in [("EB Total", TEAL), ("DG Total", "#E38B29"), ("Solar Total", "#3E8E7E")]:
                if name in sdf_p.columns:
                    fig2.add_bar(x=sdf_p["Date"], y=sdf_p[name], name=name, marker_color=color)
            fig2.update_layout(barmode="stack")
            st.plotly_chart(style_fig(fig2, "Daily Generation / Source Mix", y_title="kWh"), width='stretch', key=chart_key("exec_gen"))

        cc2 = st.columns(2)
        with cc2[0]:
            mix = {}
            for name in ["EB Total", "DG Total", "Solar Total"]:
                if name in sdf_p.columns:
                    mix[name] = sdf_p[name].sum()
            if wind_total:
                mix["Wind Total"] = wind_total
            if mix:
                fig3 = px.pie(names=list(mix.keys()), values=list(mix.values()), hole=0.55,
                               color_discrete_sequence=PALETTE)
                st.plotly_chart(style_fig(fig3, "Energy Source Mix"), width='stretch', key=chart_key("exec_mix"))
        with cc2[1]:
            if "Power Factor" in sdf_p.columns:
                fig4 = go.Figure()
                fig4.add_scatter(x=sdf_p["Date"], y=sdf_p["Power Factor"], mode="lines+markers", line=dict(color=TEAL, width=2.5))
                st.plotly_chart(style_fig(fig4, "Power Factor Trend"), width='stretch', key=chart_key("exec_pf"))
            else:
                note("Power Factor not available for this period.")

        source_data_expander(sdf_p, "exec_summary")

    if pdf is not None:
        pdf_p = filter_period(pdf, period_label)
        if pdf_p.empty:
            pdf_p = pdf
        section_title("Plan vs Actual — Target Achievement")
        metrics = pdf_p["Metric"].unique().tolist()
        cols = st.columns(min(4, max(1, len(metrics))))
        for i, m in enumerate(metrics[:4]):
            sub = pdf_p[pdf_p["Metric"] == m]
            plan_sum, act_sum = sub["Plan"].sum(), sub["Actual"].sum()
            is_green = act_sum >= plan_sum if plan_sum else False
            status_card(cols[i], m, is_green,
                        f"{act_sum:,.0f} / {plan_sum:,.0f} ({(act_sum/plan_sum*100 if plan_sum else 0):,.0f}%)")

        section_title("Variance Trend (Plan vs Actual)")
        day_agg = pdf_p.groupby(["Date", "Metric"])[["Plan", "Actual"]].sum().reset_index()
        fig5 = go.Figure()
        for i, m in enumerate(metrics[:5]):
            sub = day_agg[day_agg["Metric"] == m]
            fig5.add_scatter(x=sub["Date"], y=sub["Actual"] - sub["Plan"], name=m, mode="lines+markers",
                              line=dict(color=PALETTE[i % len(PALETTE)], width=2.5))
        fig5.add_hline(y=0, line_dash="dash", line_color=TEXT_MUTED)
        st.plotly_chart(style_fig(fig5, "Daily Variance (Actual − Plan)", y_title="Variance"), width='stretch', key=chart_key("exec_variance"))

        source_data_expander(pdf_p, "exec_plan_actual")


# =============================================================================
# 12. MAIN APP
# =============================================================================
def main():
    if not login_gate():
        return

    with st.sidebar:
        st.markdown("### ⚡ POWERANALYTICS")
        st.caption(f"Signed in as **{st.session_state.get('username','')}**")
        if st.button("Log out", width='stretch'):
            st.session_state.clear()
            st.rerun()
        st.divider()

        st.markdown("#### Company Excel Workbook")
        uploaded = st.file_uploader("Upload workbook (.xlsx)", type=["xlsx"], key="uploader")
        if uploaded is not None:
            st.session_state["file_bytes"] = uploaded.getvalue()
            st.session_state["file_name"] = uploaded.name

        file_bytes = st.session_state.get("file_bytes")

        period_label = "All available data"
        grain = "Day-wise"
        if file_bytes:
            sheet_names = get_sheet_names(file_bytes)
            summary_name = next((s for s in sheet_names if s.strip().lower() == "summary"), None)
            available_periods = ["All available data"]
            if summary_name:
                sdf = parse_summary_sheet(file_bytes, summary_name)
                if sdf is not None and not sdf.empty:
                    periods = sorted(sdf["Date"].dt.to_period("M").unique().astype(str))
                    available_periods += periods
            st.divider()
            st.markdown("#### Reporting Period")
            period_label = st.selectbox("Month", available_periods, index=len(available_periods) - 1)
            st.markdown("#### Analysis Grain")
            grain = st.radio("Grain", ["Day-wise", "Week-wise", "Month-wise"], label_visibility="collapsed")

    file_bytes = st.session_state.get("file_bytes")

    # ---- Header ----
    st.markdown(
        f"""<div class="pa-header">
            <div>
                <p class="pa-title">⚡ POWERANALYTICS</p>
                <p class="pa-sub">Enterprise Utility, Renewable Energy &amp; Generation Performance Intelligence</p>
            </div>
            <div class="pa-user">{st.session_state.get('username','').title()}<br/><span style="font-size:0.7rem;color:#9FB6D8;">Management Console</span></div>
        </div>""",
        unsafe_allow_html=True,
    )

    if not file_bytes:
        st.markdown(
            """<div class="pa-welcome">
                <h1>WELCOME TO POWERANALYTICS</h1>
                <p>Enterprise Utility, Renewable Energy &amp; Generation Performance Intelligence</p>
                <p style="margin-top:22px; font-size:0.95rem;">Upload your company Excel reporting workbook to begin.</p>
            </div>""",
            unsafe_allow_html=True,
        )
        st.write("")
        u_col1, u_col2, u_col3 = st.columns([1, 1.4, 1])
        with u_col2:
            main_upload = st.file_uploader(
                "Upload workbook (.xlsx)", type=["xlsx"], key="main_uploader",
                help="You can also upload from the sidebar (arrow at the top-left of the screen).",
            )
            if main_upload is not None:
                st.session_state["file_bytes"] = main_upload.getvalue()
                st.session_state["file_name"] = main_upload.name
                st.rerun()
        return

    sheet_names = get_sheet_names(file_bytes)
    groups = group_sheets(sheet_names)

    st.markdown(
        f"""<div class="pa-ribbon">
            <span>📁 {st.session_state.get('file_name','workbook.xlsx')} · {len(sheet_names)} worksheets detected</span>
            <span>🗓 {period_label} &nbsp;·&nbsp; 📊 {grain}</span>
        </div>""",
        unsafe_allow_html=True,
    )

    # ---- Navigation: group + sheet selectors render ONLY the chosen worksheet ----
    # (st.tabs would execute every tab's body on every rerun, which is exactly the
    # "scan everything on load" anti-pattern the workbook is too large for. A pair of
    # selectors keeps the app to a single worksheet's worth of work per interaction.)
    group_names = list(groups.keys())
    nav_cols = st.columns([1, 1, 3])
    with nav_cols[0]:
        gname = st.selectbox("Dashboard Group", group_names, key="nav_group")
    gsheets = groups[gname]
    sheet_options = (["Executive Overview"] + gsheets) if gname == "Executive" else gsheets
    with nav_cols[1]:
        sname = st.selectbox("Worksheet", sheet_options, key=f"nav_sheet_{gname}")

    st.markdown("---")

    if gname == "Executive" and sname == "Executive Overview":
        render_executive_overview(file_bytes, sheet_names, grain, period_label)
    else:
        st.markdown(f"##### 📑 {sname}")
        render_generic_sheet(file_bytes, sname, grain, period_label, f"{gname}_{sname}".replace(" ", "_"))

    st.markdown(
        f"""<div style="text-align:center; color:{TEXT_MUTED}; font-size:0.75rem; margin-top:26px;">
        POWERANALYTICS · Data sourced entirely from the uploaded workbook · No values are fabricated
        </div>""",
        unsafe_allow_html=True,
    )


if __name__ == "__main__":
    main()
