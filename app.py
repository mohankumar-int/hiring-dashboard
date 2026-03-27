import os
import json
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime

# ── Intuit brand palette ──────────────────────────────────────────────────────
BLUE    = "#0077C5"
DARK    = "#1A1A2E"
ORANGE  = "#FF6900"
GREEN   = "#2CA01C"
RED     = "#D0021B"
TEAL    = "#007D8A"
WHITE   = "#FFFFFF"
OFFWHITE= "#F8FAFB"
BORDER  = "#D1D9E0"

# ── Session persistence ───────────────────────────────────────────────────────
CACHE_DIR = os.path.expanduser("~/.hiring_dashboard")
os.makedirs(CACHE_DIR, exist_ok=True)

def _meta_path():
    return os.path.join(CACHE_DIR, "meta.json")

def load_meta() -> dict:
    p = _meta_path()
    if os.path.exists(p):
        with open(p) as f:
            return json.load(f)
    return {}

def save_to_cache(key: str, data: bytes):
    with open(os.path.join(CACHE_DIR, f"{key}.xlsx"), "wb") as f:
        f.write(data)
    meta = load_meta()
    meta[key] = datetime.now().strftime("%d %b %Y, %H:%M")
    with open(_meta_path(), "w") as f:
        json.dump(meta, f)

def load_from_cache(key: str):
    p = os.path.join(CACHE_DIR, f"{key}.xlsx")
    return open(p, "rb").read() if os.path.exists(p) else None

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Intuit Hiring Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(f"""
<style>
  /* ── Palette:  warm light background, always dark text, blue accents only ── */

  /* Page & main content */
  .stApp, .main, .block-container {{
    background-color: #F7F9FB !important;
  }}

  /* Sidebar */
  section[data-testid="stSidebar"],
  section[data-testid="stSidebar"] > div {{
    background-color: #EDF2F7 !important;
    border-right: 2px solid #C9D6E3 !important;
  }}

  /* ── Dark text everywhere — no exceptions ── */
  html, body, p, span, div, label, li, a, small,
  strong, em, h1, h2, h3, h4, h5,
  .stMarkdown, .stMarkdown *,
  [data-testid="stMarkdownContainer"],
  [data-testid="stMarkdownContainer"] *,
  section[data-testid="stSidebar"] *,
  .stSelectbox *, .stMultiSelect *,
  .stRadio *, .stFileUploader *,
  .stCaption, [data-testid="stCaptionContainer"] * {{
    color: #1C2B3A !important;
    font-family: Arial, sans-serif !important;
  }}

  /* ── Tabs: underline style — no color inversion ── */
  .stTabs [data-baseweb="tab-list"] {{
    background-color: #EDF2F7 !important;
    border-bottom: 3px solid {BLUE} !important;
    gap: 2px !important;
    padding-bottom: 0 !important;
  }}
  .stTabs [data-baseweb="tab"],
  .stTabs [data-baseweb="tab"] * {{
    font-size: 0.9rem !important;
    font-weight: 700 !important;
    color: #3A4A5C !important;
    background-color: #EDF2F7 !important;
    padding: 10px 28px !important;
    border-radius: 6px 6px 0 0 !important;
    border: 1px solid #C9D6E3 !important;
    border-bottom: none !important;
  }}
  /* Active tab: white card popping out of the blue bottom border */
  .stTabs [aria-selected="true"],
  .stTabs [aria-selected="true"] * {{
    color: #0077C5 !important;
    background-color: #FFFFFF !important;
    border-color: #C9D6E3 !important;
    font-weight: 800 !important;
  }}
  /* Tab content panel */
  .stTabs [data-baseweb="tab-panel"] {{
    background-color: #FFFFFF !important;
    border: 1px solid #C9D6E3 !important;
    border-top: none !important;
    border-radius: 0 0 8px 8px !important;
    padding: 24px !important;
  }}

  /* ── Metric cards ── */
  [data-testid="metric-container"] {{
    background: #FFFFFF !important;
    border: 1px solid #C9D6E3 !important;
    border-left: 5px solid {BLUE} !important;
    border-radius: 6px !important;
    padding: 14px 18px !important;
    box-shadow: 0 1px 4px rgba(0,60,120,.07) !important;
  }}
  [data-testid="metric-container"] [data-testid="stMetricLabel"] *  {{
    color: #3A4A5C !important;
    font-size: 0.78rem !important;
    font-weight: 700 !important;
    text-transform: uppercase !important;
    letter-spacing: .05em !important;
  }}
  [data-testid="metric-container"] [data-testid="stMetricValue"],
  [data-testid="metric-container"] [data-testid="stMetricValue"] * {{
    color: #1C2B3A !important;
    font-size: 1.6rem !important;
    font-weight: 800 !important;
  }}

  /* ── Headings ── */
  h1 {{ color: {BLUE} !important; font-size: 1.7rem !important; font-weight: 800 !important; }}
  h2, h3 {{ color: #1C2B3A !important; font-weight: 700 !important; }}

  /* ── Filter card ── */
  .filter-card {{
    background: #EDF2F7 !important;
    border: 1px solid #C9D6E3 !important;
    border-left: 5px solid {BLUE} !important;
    border-radius: 6px !important;
    padding: 16px 20px !important;
    margin-bottom: 20px !important;
  }}
  .filter-card * {{ color: #1C2B3A !important; }}

  /* ── Timestamp banner ── */
  .ts-banner {{
    background: #E8F0F8 !important;
    border: 1px solid #A8BDD0 !important;
    border-left: 5px solid {BLUE} !important;
    border-radius: 6px !important;
    padding: 10px 16px !important;
    font-size: 0.85rem !important;
    color: #1C2B3A !important;
    margin-bottom: 18px !important;
  }}
  .ts-banner * {{ color: #1C2B3A !important; }}
  .ts-banner strong {{ font-weight: 800 !important; }}

  /* ── Section label ── */
  .section-label {{
    font-size: 0.75rem !important;
    font-weight: 800 !important;
    text-transform: uppercase !important;
    letter-spacing: .07em !important;
    color: {BLUE} !important;
    margin-bottom: 4px !important;
  }}

  /* ── Download button ── */
  .stDownloadButton > button {{
    background-color: {BLUE} !important;
    color: #FFFFFF !important;
    border-radius: 4px !important;
    border: none !important;
    font-weight: 700 !important;
  }}
  .stDownloadButton > button * {{ color: #FFFFFF !important; }}
  .stDownloadButton > button:hover {{ background-color: #005EA6 !important; }}

  /* ── Multiselect tags: dark bg, white text (small pill — readable) ── */
  .stMultiSelect [data-baseweb="tag"] {{
    background-color: {BLUE} !important;
  }}
  .stMultiSelect [data-baseweb="tag"] span,
  .stMultiSelect [data-baseweb="tag"] svg {{
    color: #FFFFFF !important;
    fill: #FFFFFF !important;
  }}

  /* ── Dataframe ── */
  [data-testid="stDataFrame"] {{
    border: 1px solid #C9D6E3 !important;
    border-radius: 6px !important;
  }}

  /* ── Divider ── */
  hr {{ border-color: #C9D6E3 !important; margin: 20px 0 !important; }}

  /* ── Dropdown / select / multiselect overlays ── */

  /* The input box itself */
  [data-baseweb="select"] > div,
  [data-baseweb="select"] input {{
    background-color: #FFFFFF !important;
    color: #1C2B3A !important;
    border-color: #A8BDD0 !important;
  }}
  [data-baseweb="select"] > div *,
  [data-baseweb="select"] input * {{
    color: #1C2B3A !important;
  }}

  /* Placeholder text */
  [data-baseweb="select"] input::placeholder {{
    color: #6B7F96 !important;
  }}

  /* The floating dropdown popover */
  [data-baseweb="popover"],
  [data-baseweb="popover"] > div,
  ul[data-baseweb="menu"],
  [data-baseweb="menu"] {{
    background-color: #FFFFFF !important;
    border: 1px solid #A8BDD0 !important;
    border-radius: 6px !important;
    box-shadow: 0 4px 12px rgba(0,60,120,.12) !important;
  }}

  /* Each dropdown option */
  [role="option"],
  [data-baseweb="menu"] li,
  ul[data-baseweb="menu"] li {{
    background-color: #FFFFFF !important;
    color: #1C2B3A !important;
    font-family: Arial, sans-serif !important;
    font-size: 0.9rem !important;
  }}
  [role="option"] *,
  [data-baseweb="menu"] li * {{
    color: #1C2B3A !important;
  }}

  /* Hover state on option */
  [role="option"]:hover,
  [data-baseweb="menu"] li:hover {{
    background-color: #E8F0F8 !important;
    color: #0077C5 !important;
    cursor: pointer !important;
  }}
  [role="option"]:hover *,
  [data-baseweb="menu"] li:hover * {{
    color: #0077C5 !important;
  }}

  /* Selected / highlighted option */
  [aria-selected="true"][role="option"],
  [data-highlighted][role="option"] {{
    background-color: #D6E8F7 !important;
    color: #0077C5 !important;
  }}
  [aria-selected="true"][role="option"] *,
  [data-highlighted][role="option"] * {{
    color: #0077C5 !important;
  }}

  /* "No results" message */
  [data-baseweb="menu"] li[aria-disabled="true"],
  [data-baseweb="menu"] li[aria-disabled="true"] * {{
    color: #6B7F96 !important;
    background-color: #FFFFFF !important;
  }}
</style>
""", unsafe_allow_html=True)

# ── Column name constants ─────────────────────────────────────────────────────
R_PIPELINE_ID  = "Pipelines > Pipeline ID"
R_L2           = "DS - Join Supervisory Organization > Level 02 Organization Name"
R_L3           = "DS - Join Supervisory Organization > Level 03 Organization Name"
R_L4           = "DS - Join Supervisory Organization > Level 04 Organization Name"
R_JOB_LEVEL    = "Pipelines > Job level"
R_JOB_PROFILE  = "Pipelines > Job profile name"
R_JOB_TITLE    = "Pipelines > Job posting title"
R_JOB_FAMILY   = "Pipelines > Job family"
R_REMAINING    = "Pipelines > Remaining openings"
R_OPENINGS     = "Pipelines > # Openings"
R_STEP         = "Pipeline current step > Step"
R_TARGET_FY    = "Pipelines > Target fiscal year"
R_TARGET_QTR   = "Pipelines > Target quarter"
R_PRIORITY     = "Pipelines > Target priority level"
R_PRIORITY_CAT = "Pipelines > Target priority category"
R_RECRUITER    = "Pipelines > Recruiter"
R_TAM          = "Pipelines > TAM"
R_LOCATION     = "Pipelines > Work location"
R_COUNTRY      = "Pipelines > Location country"
R_DAYS_OPEN    = "Days Open"
R_JOB_FAM_GRP  = "Pipelines > Job family group"
R_WORKER_TYPE  = "Pipelines > Worker Type Classification - Employee Type"

H_PIPELINE_ID  = "Avature Pipeline ID"
H_L2           = "Org Level 2"
H_L3           = "Org Level 3"
H_L4           = "Org Level 4"
H_JOB_LEVEL    = "Job Level"
H_JOB_TITLE    = "Job Title"
H_JOB_FAMILY   = "Job Family"
H_JOB_FAM_GRP  = "Job Family Group"
H_START_DATE   = "Job Start Date"
H_HIRE_TYPE    = "Hire Type"
H_EMP_TYPE     = "Employee Type"
H_TARGET_FY    = "Target Fiscal Year"
H_TARGET_QTR   = "Target Quarter"
H_PRIORITY     = "Target Priority Level"
H_PRIORITY_CAT = "Target Priority Category"
H_RECRUITER    = "Recruiter Name"
H_TAM          = "TAM Name"
H_HM           = "Hiring Manager"
H_SITE         = "Business Site"
H_COUNTRY      = "Country"
H_CAREER_TRACK = "Career Track"
H_COMMUNITY    = "Community"
H_TECH_COMM    = "Tech Community"

# ── Data loading ──────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def load_open_reqs(raw: bytes) -> pd.DataFrame:
    df = pd.read_excel(raw)
    df.columns = [c.strip() for c in df.columns]
    return df

@st.cache_data(show_spinner=False)
def load_expected_hires(raw: bytes) -> pd.DataFrame:
    df = pd.read_excel(raw)
    df.columns = [c.strip() for c in df.columns]
    df[H_START_DATE] = pd.to_datetime(df[H_START_DATE], errors="coerce")
    df["_month_sort"] = df[H_START_DATE].dt.to_period("M")
    df["Month"]       = df["_month_sort"].astype(str)
    df["Month Label"] = df[H_START_DATE].dt.strftime("%b %Y")
    return df

# ── Helpers ───────────────────────────────────────────────────────────────────
def ms_filter(df, col, sel):
    return df[df[col].isin(sel)] if sel else df

def clean_options(series: pd.Series) -> list:
    return sorted(series.replace("-", pd.NA).dropna().unique().tolist())

def pivot_with_totals(df, index_col, col_col, val_col, aggfunc="sum") -> pd.DataFrame:
    piv = df.pivot_table(index=index_col, columns=col_col,
                         values=val_col, aggfunc=aggfunc, fill_value=0)
    piv = piv.reindex(sorted(piv.columns), axis=1)
    piv["Total"] = piv.sum(axis=1)
    piv = piv.sort_values("Total", ascending=False)
    grand = piv.sum()
    grand.name = "Grand Total"
    return pd.concat([piv, grand.to_frame().T])

def simple_summary(df, group_col, val_col="count") -> pd.DataFrame:
    if val_col == "count":
        s = df.groupby(group_col).size().reset_index(name="Count")
    else:
        s = df.groupby(group_col)[val_col].sum().reset_index()
    total_row = pd.DataFrame({group_col: ["Total"], s.columns[-1]: [s.iloc[:, -1].sum()]})
    return pd.concat([s.sort_values(s.columns[-1], ascending=False), total_row], ignore_index=True)

def chart_base():
    return dict(
        plot_bgcolor=WHITE,
        paper_bgcolor=WHITE,
        font=dict(family="Arial", color=DARK, size=12),
        margin=dict(t=40, b=20, l=10, r=10),
        xaxis=dict(gridcolor="#E8ECF0", linecolor=BORDER),
        yaxis=dict(gridcolor="#E8ECF0", linecolor=BORDER),
    )

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown(f"<h3 style='color:{BLUE};margin-top:0'>Data Sources</h3>", unsafe_allow_html=True)
    st.caption("Upload fresh files to refresh the dashboard. Previous data loads automatically on restart.")
    st.markdown("---")

    open_reqs_upload     = st.file_uploader("Open Hiring Requisitions (.xlsx)", type=["xlsx"], key="up_reqs")
    expected_hire_upload = st.file_uploader("Expected Hires (.xlsx)",           type=["xlsx"], key="up_hires")

    st.markdown("---")
    st.markdown(f"<small style='color:#888'>Intuit Talent Acquisition · FY26</small>", unsafe_allow_html=True)

# ── Handle uploads & cache ────────────────────────────────────────────────────
if open_reqs_upload:
    raw = open_reqs_upload.read()
    save_to_cache("reqs", raw)
    st.session_state["reqs_bytes"] = raw

if expected_hire_upload:
    raw = expected_hire_upload.read()
    save_to_cache("hires", raw)
    st.session_state["hires_bytes"] = raw

# On first load, pull from disk cache if session is empty
if "reqs_bytes" not in st.session_state:
    cached = load_from_cache("reqs")
    if cached:
        st.session_state["reqs_bytes"] = cached

if "hires_bytes" not in st.session_state:
    cached = load_from_cache("hires")
    if cached:
        st.session_state["hires_bytes"] = cached

reqs_bytes  = st.session_state.get("reqs_bytes")
hires_bytes = st.session_state.get("hires_bytes")

# ── Header ────────────────────────────────────────────────────────────────────
st.markdown("<h1>Intuit Hiring Dashboard</h1>", unsafe_allow_html=True)

# Timestamp banner
meta = load_meta()
ts_reqs  = meta.get("reqs",  "not loaded")
ts_hires = meta.get("hires", "not loaded")
st.markdown(
    f"<div class='ts-banner'>"
    f"Open Reqs data as of: <strong>{ts_reqs}</strong> &nbsp;|&nbsp; "
    f"Expected Hires data as of: <strong>{ts_hires}</strong>"
    f"</div>",
    unsafe_allow_html=True,
)

if not reqs_bytes and not hires_bytes:
    st.info("Upload one or both data files in the sidebar to get started.")
    st.stop()

df_reqs  = load_open_reqs(reqs_bytes)   if reqs_bytes  else None
df_hires = load_expected_hires(hires_bytes) if hires_bytes else None

# Pipeline IDs that have expected hires
pipeline_ids_with_offers: set = set()
if df_hires is not None:
    pipeline_ids_with_offers = set(df_hires[H_PIPELINE_ID].dropna().astype(int).unique())

if df_reqs is not None:
    df_reqs["Has Expected Offer"] = (
        df_reqs[R_PIPELINE_ID].isin(pipeline_ids_with_offers)
        .map({True: "Yes", False: "No"})
    )

# ══════════════════════════════════════════════════════════════════════════════
tab1, tab2 = st.tabs(["  Open Hiring Requisitions  ", "  Expected Hires  "])

# ─────────────────────────────────────────────────────────────────────────────
# TAB 1
# ─────────────────────────────────────────────────────────────────────────────
with tab1:
    if df_reqs is None:
        st.warning("Upload the **Open Hiring Requisitions** file in the sidebar.")
    else:
        # ── Filters ──────────────────────────────────────────────────────────
        st.markdown("<div class='filter-card'>", unsafe_allow_html=True)
        st.markdown("**Filters**")

        # Row 1 — Employee Type (always first, default = Employee) + Org hierarchy
        fr1a, fr1b, fr1c, fr1d = st.columns(4)
        with fr1a:
            wt_opts = clean_options(df_reqs[R_WORKER_TYPE])
            wt_default = ["Employee"] if "Employee" in wt_opts else []
            wt_filter = st.multiselect("Employee Type", wt_opts,
                                       default=wt_default, key="r_wt")
        with fr1b:
            l2_filter = st.multiselect("Level 2 Org", clean_options(df_reqs[R_L2]), key="r_l2")
        with fr1c:
            l3_filter = st.multiselect("Level 3 Org", clean_options(df_reqs[R_L3]), key="r_l3")
        with fr1d:
            l4_filter = st.multiselect("Level 4 Org", clean_options(df_reqs[R_L4]), key="r_l4")

        # Row 2 — Job dimensions
        fr2a, fr2b, fr2c, fr2d = st.columns(4)
        with fr2a:
            jfg_filter = st.multiselect("Job Family Group", clean_options(df_reqs[R_JOB_FAM_GRP]), key="r_jfg")
        with fr2b:
            jf_filter  = st.multiselect("Job Family",       clean_options(df_reqs[R_JOB_FAMILY]),  key="r_jf")
        with fr2c:
            jl_filter  = st.multiselect("Job Level",        clean_options(df_reqs[R_JOB_LEVEL]),   key="r_jl")
        with fr2d:
            jp_filter  = st.multiselect("Job Profile",      clean_options(df_reqs[R_JOB_PROFILE]), key="r_jp")

        # Row 3 — People & Priority
        fr3a, fr3b, fr3c, fr3d = st.columns(4)
        with fr3a:
            tam_filter  = st.multiselect("TAM",       clean_options(df_reqs[R_TAM]),          key="r_tam")
        with fr3b:
            rec_filter  = st.multiselect("Recruiter", clean_options(df_reqs[R_RECRUITER]),    key="r_rec")
        with fr3c:
            pri_filter  = st.multiselect("Priority Level",    clean_options(df_reqs[R_PRIORITY]),    key="r_pri")
        with fr3d:
            pricat_filter = st.multiselect("Priority Category", clean_options(df_reqs[R_PRIORITY_CAT]), key="r_pricat")

        # Row 4 — Step, FY, Offer toggle, Pipeline IDs
        fr4a, fr4b, fr4c, fr4d = st.columns(4)
        with fr4a:
            step_filter = st.multiselect("Pipeline Step", clean_options(df_reqs[R_STEP]), key="r_step")
        with fr4b:
            fy_filter   = st.multiselect("Target FY",    clean_options(df_reqs[R_TARGET_FY]), key="r_fy")
        with fr4c:
            offer_toggle = st.selectbox(
                "Has Expected Offer?", ["All", "Yes — has offer", "No — no offer"], key="r_offer"
            )
        with fr4d:
            if df_hires is not None and offer_toggle != "No — no offer":
                pid_filter = st.multiselect(
                    "Specific Pipeline IDs (Expected Hires)",
                    sorted(pipeline_ids_with_offers), key="r_pid",
                )
            else:
                pid_filter = []

        st.markdown("</div>", unsafe_allow_html=True)

        # Apply filters
        f = df_reqs.copy()
        f = ms_filter(f, R_WORKER_TYPE,  wt_filter)
        f = ms_filter(f, R_L2,           l2_filter)
        f = ms_filter(f, R_L3,           l3_filter)
        f = ms_filter(f, R_L4,           l4_filter)
        f = ms_filter(f, R_JOB_FAM_GRP,  jfg_filter)
        f = ms_filter(f, R_JOB_FAMILY,   jf_filter)
        f = ms_filter(f, R_JOB_LEVEL,    jl_filter)
        f = ms_filter(f, R_JOB_PROFILE,  jp_filter)
        f = ms_filter(f, R_TAM,          tam_filter)
        f = ms_filter(f, R_RECRUITER,    rec_filter)
        f = ms_filter(f, R_PRIORITY,     pri_filter)
        f = ms_filter(f, R_PRIORITY_CAT, pricat_filter)
        f = ms_filter(f, R_STEP,         step_filter)
        f = ms_filter(f, R_TARGET_FY,    fy_filter)
        if offer_toggle == "Yes — has offer":
            f = f[f["Has Expected Offer"] == "Yes"]
        elif offer_toggle == "No — no offer":
            f = f[f["Has Expected Offer"] == "No"]
        if pid_filter:
            f = f[f[R_PIPELINE_ID].isin(pid_filter)]

        # ── Metrics ──────────────────────────────────────────────────────────
        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("Total Requisitions",     f"{len(f):,}")
        m2.metric("Remaining Openings",     f"{int(f[R_REMAINING].sum()):,}")
        m3.metric("With Expected Offer",    f"{(f['Has Expected Offer']=='Yes').sum():,}")
        m4.metric("Without Expected Offer", f"{(f['Has Expected Offer']=='No').sum():,}")
        m5.metric("Avg Days Open",
                  f"{f[R_DAYS_OPEN].mean():.0f}" if f[R_DAYS_OPEN].notna().any() else "—")

        st.markdown("---")

        # ── Table 1: L2 Org — 3-column split ─────────────────────────────────
        st.markdown("<p class='section-label'>Remaining Openings by Level 2 Org</p>", unsafe_allow_html=True)
        l2_grp = (
            f.groupby([R_L2, "Has Expected Offer"])[R_REMAINING].sum()
            .unstack(fill_value=0)
        )
        for col in ["Yes", "No"]:
            if col not in l2_grp.columns:
                l2_grp[col] = 0
        l2_grp = l2_grp.rename(columns={"Yes": "With Offer", "No": "Without Offer"})
        l2_grp["Total Remaining"] = l2_grp["With Offer"] + l2_grp["Without Offer"]
        l2_grp = l2_grp[["With Offer", "Without Offer", "Total Remaining"]].sort_values("Total Remaining", ascending=False)
        grand_l2 = l2_grp.sum(); grand_l2.name = "Total"
        l2_grp = pd.concat([l2_grp, grand_l2.to_frame().T])
        l2_grp.index.name = "Level 2 Org"
        st.dataframe(l2_grp, use_container_width=True)

        st.markdown("---")

        # ── Table 2: L3 Org × Job Level — toggleable view ────────────────────
        st.markdown("<p class='section-label'>Remaining Openings by L3 Org × Job Level</p>", unsafe_allow_html=True)
        l3_view = st.radio(
            "Show openings:",
            ["Total", "With Offer", "Without Offer"],
            horizontal=True, key="l3_jl_toggle",
        )
        if l3_view == "With Offer":
            l3_src = f[f["Has Expected Offer"] == "Yes"]
        elif l3_view == "Without Offer":
            l3_src = f[f["Has Expected Offer"] == "No"]
        else:
            l3_src = f
        l3_jl_piv = pivot_with_totals(l3_src, R_L3, R_JOB_LEVEL, R_REMAINING)
        l3_jl_piv.index.name = "L3 Org"
        st.dataframe(l3_jl_piv, use_container_width=True)

        st.markdown("---")

        # ── Table 3: Job Level × Has Expected Offer (replaces chart) ─────────
        st.markdown("<p class='section-label'>Remaining Openings by Job Level — Offer Status</p>", unsafe_allow_html=True)

        jl_offer_piv = pivot_with_totals(f, R_JOB_LEVEL, "Has Expected Offer", R_REMAINING)
        jl_offer_piv.index.name = "Job Level"
        # Rename columns for clarity
        jl_offer_piv = jl_offer_piv.rename(columns={"No": "No Offer", "Yes": "Has Offer"})

        tc1, tc2 = st.columns([2, 3])
        with tc1:
            st.dataframe(jl_offer_piv, use_container_width=True)
        with tc2:
            # Keep one focused chart for visual split
            chart_src = (
                f.groupby([R_JOB_LEVEL, "Has Expected Offer"])[R_REMAINING]
                .sum().reset_index()
                .rename(columns={R_JOB_LEVEL: "Job Level", R_REMAINING: "Openings", "Has Expected Offer": "Offer"})
            )
            fig = px.bar(
                chart_src, x="Job Level", y="Openings", color="Offer",
                color_discrete_map={"Yes": GREEN, "No": ORANGE},
                barmode="stack",
                title="Job Level — Offer Status Split",
            )
            fig.update_layout(**chart_base(), height=320,
                              legend=dict(orientation="h", y=-0.2, title=""),
                              title_font=dict(size=13))
            st.plotly_chart(fig, use_container_width=True)

        st.markdown("---")

        # ── Table 4: Priority breakdown ───────────────────────────────────────
        st.markdown("<p class='section-label'>Remaining Openings by Priority Level</p>", unsafe_allow_html=True)
        pri_tbl = (
            f.dropna(subset=[R_PRIORITY])
            .groupby(R_PRIORITY)[R_REMAINING].sum()
            .reset_index().rename(columns={R_PRIORITY: "Priority", R_REMAINING: "Remaining Openings"})
            .sort_values("Priority")
        )
        pri_total = pd.DataFrame({"Priority": ["Total"], "Remaining Openings": [pri_tbl["Remaining Openings"].sum()]})
        st.dataframe(
            pd.concat([pri_tbl, pri_total], ignore_index=True),
            use_container_width=True, hide_index=True,
        )

        st.markdown("---")

        # ── Main detail table ─────────────────────────────────────────────────
        st.markdown("<p class='section-label'>All Open Requisitions</p>", unsafe_allow_html=True)
        st.caption(f"Sorted: L3 Org → L4 Org → Job Level  ·  {len(f):,} rows")

        col_map = {
            R_PIPELINE_ID:  "Pipeline ID",
            R_WORKER_TYPE:  "Employee Type",
            R_L2:           "L2 Org",
            R_L3:           "L3 Org",
            R_L4:           "L4 Org",
            R_JOB_FAM_GRP:  "Job Family Group",
            R_JOB_FAMILY:   "Job Family",
            R_JOB_LEVEL:    "Job Level",
            R_JOB_PROFILE:  "Job Profile",
            R_JOB_TITLE:    "Job Title",
            R_REMAINING:    "Remaining",
            R_OPENINGS:     "# Openings",
            R_STEP:         "Current Step",
            R_TARGET_FY:    "Target FY",
            R_TARGET_QTR:   "Quarter",
            R_PRIORITY:     "Priority",
            R_PRIORITY_CAT: "Priority Category",
            R_RECRUITER:    "Recruiter",
            R_TAM:          "TAM",
            R_LOCATION:     "Location",
            R_DAYS_OPEN:    "Days Open",
            "Has Expected Offer": "Has Offer?",
        }
        tbl = (
            f[[c for c in col_map if c in f.columns]]
            .rename(columns=col_map)
            .sort_values(["L3 Org", "L4 Org", "Job Level"])
            .reset_index(drop=True)
        )

        def row_style(row):
            if row.get("Has Offer?") == "Yes":
                return ["background-color:#EAF5EA; color:#1A1A2E"] * len(row)
            return [""] * len(row)

        st.dataframe(
            tbl.style.apply(row_style, axis=1),
            use_container_width=True, height=500,
        )
        st.download_button(
            "Download Filtered Data (CSV)",
            tbl.to_csv(index=False).encode("utf-8"),
            "open_reqs_filtered.csv", "text/csv",
        )

# ─────────────────────────────────────────────────────────────────────────────
# TAB 2
# ─────────────────────────────────────────────────────────────────────────────
with tab2:
    if df_hires is None:
        st.warning("Upload the **Expected Hires** file in the sidebar.")
    else:
        # ── Filters ──────────────────────────────────────────────────────────
        st.markdown("<div class='filter-card'>", unsafe_allow_html=True)
        st.markdown("**Filters**")

        # Row 1 — Employee Type (first, default Employee) + Org hierarchy
        hf1a, hf1b, hf1c, hf1d = st.columns(4)
        with hf1a:
            het_opts    = clean_options(df_hires[H_EMP_TYPE]) if H_EMP_TYPE in df_hires.columns else []
            het_default = ["Employee"] if "Employee" in het_opts else []
            het = st.multiselect("Employee Type", het_opts, default=het_default, key="h_et")
        with hf1b:
            hl2  = st.multiselect("Org Level 2", clean_options(df_hires[H_L2])  if H_L2  in df_hires.columns else [], key="h_l2")
        with hf1c:
            hl3  = st.multiselect("Org Level 3", clean_options(df_hires[H_L3]),  key="h_l3")
        with hf1d:
            hl4  = st.multiselect("Org Level 4", clean_options(df_hires[H_L4])  if H_L4  in df_hires.columns else [], key="h_l4")

        # Row 2 — Job dimensions
        hf2a, hf2b, hf2c, hf2d = st.columns(4)
        with hf2a:
            hjfg = st.multiselect("Job Family Group", clean_options(df_hires[H_JOB_FAM_GRP]) if H_JOB_FAM_GRP in df_hires.columns else [], key="h_jfg")
        with hf2b:
            hjf  = st.multiselect("Job Family",       clean_options(df_hires[H_JOB_FAMILY]),  key="h_jf")
        with hf2c:
            hjl  = st.multiselect("Job Level",        clean_options(df_hires[H_JOB_LEVEL]),   key="h_jl")
        with hf2d:
            hht  = st.multiselect("Hire Type",        clean_options(df_hires[H_HIRE_TYPE]),   key="h_ht")

        # Row 3 — People & Priority
        hf3a, hf3b, hf3c, hf3d = st.columns(4)
        with hf3a:
            htam = st.multiselect("TAM",               clean_options(df_hires[H_TAM]),          key="h_tam")
        with hf3b:
            hrec = st.multiselect("Recruiter",         clean_options(df_hires[H_RECRUITER]),    key="h_rec")
        with hf3c:
            hpp  = st.multiselect("Priority Level",    clean_options(df_hires[H_PRIORITY]),     key="h_pri")
        with hf3d:
            hppc = st.multiselect("Priority Category", clean_options(df_hires[H_PRIORITY_CAT]) if H_PRIORITY_CAT in df_hires.columns else [], key="h_pricat")

        # Row 4 — FY & Quarter
        hf4a, hf4b, _, _ = st.columns(4)
        with hf4a:
            hfy  = st.multiselect("Target FY", clean_options(df_hires[H_TARGET_FY]),  key="h_fy")
        with hf4b:
            hqtr = st.multiselect("Quarter",   clean_options(df_hires[H_TARGET_QTR]), key="h_qtr")

        # Row 5 — Career Track / Community / Tech Community
        hf5a, hf5b, hf5c, _ = st.columns(4)
        with hf5a:
            hct  = st.multiselect("Career Track",   clean_options(df_hires[H_CAREER_TRACK]) if H_CAREER_TRACK in df_hires.columns else [], key="h_ct")
        with hf5b:
            hcom = st.multiselect("Community",      clean_options(df_hires[H_COMMUNITY])    if H_COMMUNITY    in df_hires.columns else [], key="h_com")
        with hf5c:
            htc  = st.multiselect("Tech Community", clean_options(df_hires[H_TECH_COMM])    if H_TECH_COMM    in df_hires.columns else [], key="h_tc")

        st.markdown("</div>", unsafe_allow_html=True)

        fh = df_hires.copy()
        if het  and H_EMP_TYPE     in fh.columns: fh = ms_filter(fh, H_EMP_TYPE,     het)
        if hl2  and H_L2           in fh.columns: fh = ms_filter(fh, H_L2,           hl2)
        fh = ms_filter(fh, H_L3,         hl3)
        if hl4  and H_L4           in fh.columns: fh = ms_filter(fh, H_L4,           hl4)
        if hjfg and H_JOB_FAM_GRP  in fh.columns: fh = ms_filter(fh, H_JOB_FAM_GRP,  hjfg)
        fh = ms_filter(fh, H_JOB_FAMILY,  hjf)
        fh = ms_filter(fh, H_JOB_LEVEL,   hjl)
        fh = ms_filter(fh, H_HIRE_TYPE,   hht)
        fh = ms_filter(fh, H_TAM,         htam)
        fh = ms_filter(fh, H_RECRUITER,   hrec)
        fh = ms_filter(fh, H_PRIORITY,    hpp)
        if hppc and H_PRIORITY_CAT in fh.columns: fh = ms_filter(fh, H_PRIORITY_CAT, hppc)
        fh = ms_filter(fh, H_TARGET_FY,   hfy)
        fh = ms_filter(fh, H_TARGET_QTR,  hqtr)
        if hct  and H_CAREER_TRACK in fh.columns: fh = ms_filter(fh, H_CAREER_TRACK, hct)
        if hcom and H_COMMUNITY    in fh.columns: fh = ms_filter(fh, H_COMMUNITY,    hcom)
        if htc  and H_TECH_COMM    in fh.columns: fh = ms_filter(fh, H_TECH_COMM,    htc)

        # ── Metrics ──────────────────────────────────────────────────────────
        hm1, hm2, hm3, hm4 = st.columns(4)
        hm1.metric("Total Expected Hires", f"{len(fh):,}")
        hm2.metric("Unique Pipelines",     f"{fh[H_PIPELINE_ID].nunique():,}")
        hm3.metric("Months Covered",       f"{fh['Month'].nunique():,}")
        hm4.metric("L3 Orgs",              f"{fh[H_L3].nunique():,}")

        st.markdown("---")

        # ── Monthly trend chart ───────────────────────────────────────────────
        st.markdown("<p class='section-label'>Expected Hires by Month</p>", unsafe_allow_html=True)
        monthly = (
            fh.groupby(["Month", "Month Label"]).size()
            .reset_index(name="Hires").sort_values("Month")
        )
        fig_trend = px.bar(monthly, x="Month Label", y="Hires",
                           color_discrete_sequence=[BLUE])
        fig_trend.update_traces(marker_line_color=WHITE, marker_line_width=1)
        fig_trend.update_layout(**chart_base(), height=280,
                                xaxis_title="", yaxis_title="# Expected Hires",
                                bargap=0.25)
        st.plotly_chart(fig_trend, use_container_width=True)

        st.markdown("---")

        # ── Table 1: Month × L3 Org ───────────────────────────────────────────
        st.markdown("<p class='section-label'>Expected Hires by Month × L3 Org</p>", unsafe_allow_html=True)
        piv_l3 = fh.pivot_table(index=H_L3, columns="Month", aggfunc="size", fill_value=0)
        piv_l3 = piv_l3.reindex(sorted(piv_l3.columns), axis=1)
        piv_l3["Total"] = piv_l3.sum(axis=1)
        piv_l3 = piv_l3.sort_values("Total", ascending=False)
        grand_l3 = piv_l3.sum(); grand_l3.name = "Grand Total"
        piv_l3 = pd.concat([piv_l3, grand_l3.to_frame().T])
        piv_l3.index.name = "L3 Org"
        st.dataframe(piv_l3, use_container_width=True)

        st.markdown("---")

        # ── Table 2: Month × Job Level ────────────────────────────────────────
        st.markdown("<p class='section-label'>Expected Hires by Month × Job Level</p>", unsafe_allow_html=True)
        piv_jl = fh.pivot_table(index=H_JOB_LEVEL, columns="Month", aggfunc="size", fill_value=0)
        piv_jl = piv_jl.reindex(sorted(piv_jl.columns), axis=1)
        piv_jl["Total"] = piv_jl.sum(axis=1)
        piv_jl = piv_jl.sort_values("Total", ascending=False)
        grand_jl = piv_jl.sum(); grand_jl.name = "Grand Total"
        piv_jl = pd.concat([piv_jl, grand_jl.to_frame().T])
        piv_jl.index.name = "Job Level"
        st.dataframe(piv_jl, use_container_width=True)

        st.markdown("---")

        # ── Table 3: L3 Org × Job Level cross-tab ────────────────────────────
        st.markdown("<p class='section-label'>Expected Hires by L3 Org × Job Level</p>", unsafe_allow_html=True)
        piv_l3jl = fh.pivot_table(index=H_L3, columns=H_JOB_LEVEL, aggfunc="size", fill_value=0)
        piv_l3jl["Total"] = piv_l3jl.sum(axis=1)
        piv_l3jl = piv_l3jl.sort_values("Total", ascending=False)
        grand_l3jl = piv_l3jl.sum(); grand_l3jl.name = "Grand Total"
        piv_l3jl = pd.concat([piv_l3jl, grand_l3jl.to_frame().T])
        piv_l3jl.index.name = "L3 Org"
        st.dataframe(piv_l3jl, use_container_width=True)

        st.markdown("---")

        # ── Table 4: Hire Type summary ────────────────────────────────────────
        st.markdown("<p class='section-label'>Expected Hires by Hire Type & Priority</p>", unsafe_allow_html=True)
        tc1, tc2 = st.columns(2)
        with tc1:
            ht_tbl = simple_summary(fh.replace("-", pd.NA).dropna(subset=[H_HIRE_TYPE]), H_HIRE_TYPE)
            ht_tbl.columns = ["Hire Type", "Count"]
            st.dataframe(ht_tbl, use_container_width=True, hide_index=True)
        with tc2:
            pri_h_tbl = simple_summary(fh.replace("-", pd.NA).dropna(subset=[H_PRIORITY]), H_PRIORITY)
            pri_h_tbl.columns = ["Priority", "Count"]
            st.dataframe(pri_h_tbl, use_container_width=True, hide_index=True)

        st.markdown("---")

        # ── Detail table ──────────────────────────────────────────────────────
        st.markdown("<p class='section-label'>Expected Hires Detail</p>", unsafe_allow_html=True)
        st.caption(f"Sorted: L3 Org → Job Level → Month  ·  {len(fh):,} rows")

        hire_col_map = {
            H_PIPELINE_ID:  "Pipeline ID",
            H_L3:           "L3 Org",
            H_L4:           "L4 Org",
            H_JOB_LEVEL:    "Job Level",
            H_JOB_TITLE:    "Job Title",
            H_JOB_FAMILY:   "Job Family",
            "Month Label":  "Start Month",
            H_HIRE_TYPE:    "Hire Type",
            H_TARGET_FY:    "Target FY",
            H_TARGET_QTR:   "Quarter",
            H_PRIORITY:     "Priority",
            H_PRIORITY_CAT: "Priority Category",
            H_RECRUITER:    "Recruiter",
            H_TAM:          "TAM",
            H_HM:           "Hiring Manager",
            H_SITE:         "Site",
            H_COUNTRY:      "Country",
        }
        hcols = [c for c in hire_col_map if c in fh.columns]
        hire_tbl = (
            fh[hcols].rename(columns=hire_col_map)
            .sort_values(["L3 Org", "Job Level", "Start Month"])
            .reset_index(drop=True)
        )
        st.dataframe(hire_tbl, use_container_width=True, height=480)
        st.download_button(
            "Download Expected Hires (CSV)",
            hire_tbl.to_csv(index=False).encode("utf-8"),
            "expected_hires_filtered.csv", "text/csv",
        )
