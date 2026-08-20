import os
import re
import base64
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st


# ============================================================
# CONFIG
# ============================================================

st.set_page_config(
    page_title="AMECATH | MENA Dialysis Market Intelligence",
    page_icon="🩺",
    layout="wide",
    initial_sidebar_state="expanded",
)

BASE_DIR = Path(__file__).resolve().parent

EXCEL_FILE = BASE_DIR / "Amecath Dash(10).xlsx"
ASSET_DIR = BASE_DIR / "assets"

COUNTRIES = [
    "🇸🇦 Saudi Arabia",
    "🇦🇪 UAE",
    "🇶🇦 Qatar",
    "🇰🇼 Kuwait",
    "🇴🇲 Oman",
    "🇯🇴 Jordan",
    "🇱🇧 Lebanon",
    "🇮🇶 Iraq",
    "🇧🇭 Bahrain",
]

COUNTRY_NAMES = {
    "🇸🇦 Saudi Arabia": "Saudi Arabia",
    "🇦🇪 UAE": "UAE",
    "🇶🇦 Qatar": "Qatar",
    "🇰🇼 Kuwait": "Kuwait",
    "🇴🇲 Oman": "Oman",
    "🇯🇴 Jordan": "Jordan",
    "🇱🇧 Lebanon": "Lebanon",
    "🇮🇶 Iraq": "Iraq",
    "🇧🇭 Bahrain": "Bahrain",
}

COUNTRY_FLAGS = {
    "Saudi Arabia": "🇸🇦",
    "UAE": "🇦🇪",
    "Qatar": "🇶🇦",
    "Kuwait": "🇰🇼",
    "Oman": "🇴🇲",
    "Jordan": "🇯🇴",
    "Lebanon": "🇱🇧",
    "Iraq": "🇮🇶",
    "Bahrain": "🇧🇭",
}

# Exact filenames as uploaded to the repo. Each entry is looked up both at
# the repo root and inside an "assets" folder, so either location works.
COUNTRY_ASSET_FILES = {
    "Saudi Arabia": {"flag": "saudi_arabia_flag.jpeg", "landscape": "saudi_landscape.jpeg"},
    "UAE": {"flag": "uae_flag.jpeg", "landscape": "uae_landscape.jpeg"},
    "Qatar": {"flag": "qatar_flag.jpeg", "landscape": "qatar_landscape.jpeg"},
    "Kuwait": {"flag": "kuwait_flag.jpeg", "landscape": "kuwait_landscape.jpeg"},
    "Oman": {"flag": "oman_flag.jpeg", "landscape": "oman_landscape.jpeg"},
    "Jordan": {"flag": "jordon_flag.jpeg", "landscape": "jordon_landscape.jpeg"},
    "Lebanon": {"flag": "lebanon_flag.jpeg", "landscape": "lebanon_landscape.jpeg"},
    "Iraq": {"flag": "iraq_flag.jpg", "landscape": "iraq_landscape.jpg"},
    "Bahrain": {"flag": "bahrain_flag.jpeg", "landscape": "bahrain_landscape.jpg"},
}

# Primary / secondary colors pulled from each country's flag, used to theme
# the Country Explorer page (card borders, headings, accents) when that
# country is selected.
COUNTRY_COLORS = {
    "Saudi Arabia": {"primary": "#006C35", "secondary": "#FFFFFF"},
    "UAE": {"primary": "#FF0000", "secondary": "#00732F"},
    "Qatar": {"primary": "#8D1B3D", "secondary": "#FFFFFF"},
    "Kuwait": {"primary": "#CE1126", "secondary": "#007A3D"},
    "Oman": {"primary": "#DB161B", "secondary": "#008000"},
    "Jordan": {"primary": "#CE1126", "secondary": "#007A3D"},
    "Lebanon": {"primary": "#ED1C24", "secondary": "#00A651"},
    "Iraq": {"primary": "#CE1126", "secondary": "#007A3D"},
    "Bahrain": {"primary": "#CE1126", "secondary": "#FFFFFF"},
}

DEFAULT_ACCENT = {"primary": "#3B82F6", "secondary": "#93C5FD"}


# ============================================================
# CSS (base dark theme — cards match the app background,
# text colors chosen to read cleanly on that background)
# ============================================================

st.markdown(
    """
<style>

    .stApp {
        background-color: #0F172A;
    }

    [data-testid="stSidebar"] {
        background-color: #17191D;
    }

    [data-testid="stSidebar"] * {
        color: #FFFFFF !important;
    }

    .main-title {
        font-size: 34px;
        font-weight: 800;
        color: #F8FAFC;
        margin-bottom: 0;
    }

    .sub-title {
        color: #94A3B8;
        font-size: 14px;
        margin-top: -5px;
        margin-bottom: 20px;
    }

    .section-title {
        font-size: 21px;
        font-weight: 750;
        color: #F8FAFC;
        margin-top: 20px;
        margin-bottom: 12px;
    }

    .kpi-card {
        background: rgba(30, 41, 59, 0.65);
        border: 1px solid rgba(148, 163, 184, 0.25);
        border-radius: 12px;
        padding: 18px;
        min-height: 120px;
        box-shadow: 0 2px 7px rgba(0,0,0,0.15);
    }

    .kpi-label {
        color: #94A3B8;
        font-size: 12px;
        font-weight: 650;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }

    .kpi-value {
        color: #F8FAFC;
        font-size: 27px;
        font-weight: 800;
        margin-top: 8px;
    }

    .hero {
        border-radius: 16px;
        overflow: hidden;
        min-height: 280px;
        position: relative;
        background: #23262B;
        margin-bottom: 18px;
    }

    .hero img {
        width: 100%;
        height: 300px;
        object-fit: cover;
    }

    .hero-flag {
        position: absolute;
        top: 20px;
        right: 20px;
        width: 54px;
        height: 54px;
        border-radius: 8px;
        object-fit: cover;
        box-shadow: 0 2px 10px rgba(0,0,0,.5);
        border: 2px solid rgba(255,255,255,0.85);
    }

    .hero-text {
        position: absolute;
        left: 30px;
        bottom: 25px;
        color: white;
        text-shadow: 0 2px 8px rgba(0,0,0,.65);
    }

    .hero-country {
        font-size: 34px;
        font-weight: 850;
    }

    .hero-description {
        font-size: 15px;
        margin-top: 5px;
    }

    .action-card {
        background: rgba(30, 41, 59, 0.65);
        border-radius: 12px;
        padding: 18px;
        border: 1px solid rgba(148, 163, 184, 0.25);
        min-height: 135px;
    }

    .action-title {
        font-size: 12px;
        text-transform: uppercase;
        color: #94A3B8;
        font-weight: 700;
    }

    .action-value {
        font-size: 20px;
        font-weight: 800;
        margin-top: 8px;
        color: #F8FAFC;
    }

    .action-description {
        font-size: 13px;
        color: #94A3B8;
        margin-top: 4px;
    }

    .score-box {
        background: rgba(30, 41, 59, 0.65);
        border: 1px solid rgba(148, 163, 184, 0.25);
        border-radius: 12px;
        padding: 18px;
    }

    .score-number {
        font-size: 36px;
        font-weight: 850;
        color: #F8FAFC;
    }

    .source-note {
        color: #94A3B8;
        font-size: 11px;
        margin-top: 20px;
    }

    .badge {
        padding: 5px 10px;
        border-radius: 999px;
        font-size: 11px;
        font-weight: 700;
        display: inline-block;
    }

    .badge-green {
        background: #E8F5EC;
        color: #19733A;
    }

    .badge-yellow {
        background: #FFF4D6;
        color: #866200;
    }

    .badge-red {
        background: #FBE7E7;
        color: #A62C2C;
    }

    .badge-gray {
        background: #ECEDEF;
        color: #555A62;
    }

</style>
""",
    unsafe_allow_html=True,
)


# ============================================================
# HELPERS
# ============================================================

def clean_text(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def country_without_flag(country):
    country = clean_text(country)
    return re.sub(r"^[^\w]+", "", country).strip()


def number_value(value):
    """
    Converts values such as:
    35,165,787
    ~1,279
    $35M
    35M
    9%
    """
    if pd.isna(value):
        return None

    text = str(value).strip()

    text = text.replace(",", "")
    text = text.replace("~", "")
    text = text.replace("$", "")

    multiplier = 1

    if text.lower().endswith("m"):
        multiplier = 1_000_000
        text = text[:-1]

    elif text.lower().endswith("k"):
        multiplier = 1_000
        text = text[:-1]

    elif text.endswith("%"):
        text = text[:-1]
        try:
            return float(text) / 100
        except Exception:
            return None

    try:
        return float(text) * multiplier
    except Exception:
        return None


def money_to_number(value):
    return number_value(value)


def format_number(value):
    if value is None or pd.isna(value):
        return "N/A"

    value = float(value)

    if abs(value) >= 1_000_000_000:
        return f"{value / 1_000_000_000:.1f}B"

    if abs(value) >= 1_000_000:
        return f"{value / 1_000_000:.1f}M"

    if abs(value) >= 1_000:
        return f"{value / 1_000:.1f}K"

    return f"{value:,.0f}"


def format_currency(value):
    if value is None or pd.isna(value):
        return "N/A"

    value = float(value)

    if value >= 1_000_000_000:
        return f"${value / 1_000_000_000:.1f}B"

    if value >= 1_000_000:
        return f"${value / 1_000_000:.1f}M"

    if value >= 1_000:
        return f"${value / 1_000:.1f}K"

    return f"${value:,.0f}"


def format_percent(value):
    if value is None or pd.isna(value):
        return "N/A"

    return f"{float(value) * 100:.1f}%"


def safe_numeric(series):
    return pd.to_numeric(
        series.astype(str)
        .str.replace(",", "", regex=False)
        .str.replace("~", "", regex=False)
        .str.replace("$", "", regex=False),
        errors="coerce",
    )


def image_to_base64(path):
    with open(path, "rb") as f:
        return base64.b64encode(f.read()).decode()


# ============================================================
# LOAD DATA
# ============================================================

@st.cache_data
def load_excel():
    if not EXCEL_FILE.exists():
        st.error(f"Excel file not found: {EXCEL_FILE}")
        st.stop()

    xls = pd.ExcelFile(EXCEL_FILE)

    data = {}

    for sheet in xls.sheet_names:
        data[sheet] = pd.read_excel(EXCEL_FILE, sheet_name=sheet)

    return data


@st.cache_data
def prepare_data():
    data = load_excel()

    summary = data.get("Sheet3", pd.DataFrame()).copy()

    if not summary.empty:
        summary["Country Clean"] = summary["Country"].apply(country_without_flag)

        numeric_columns = [
            "Population 2026",
            "Est. 2026 HD",
            "Est. 2026 PD",
            "Annual Growth",
            "Hospital Growth",
            "Unit Growth",
            "HD Machines",
            "Annual Catheter Demand",
            "Market Value",
        ]

        for col in numeric_columns:
            if col in summary.columns:
                summary[col + " Num"] = summary[col].apply(number_value)

    return data, summary


data, summary = prepare_data()


# ============================================================
# IMAGE HANDLING
# ============================================================

def find_asset(filename):
    """Look for an uploaded image at the repo root or inside /assets."""
    if not filename:
        return None

    for candidate in (BASE_DIR / filename, ASSET_DIR / filename):
        if candidate.exists():
            return str(candidate)

    return None


def get_country_flag_path(country):
    info = COUNTRY_ASSET_FILES.get(country, {})
    return find_asset(info.get("flag"))


def get_country_landscape_path(country):
    info = COUNTRY_ASSET_FILES.get(country, {})
    return find_asset(info.get("landscape"))


def get_country_colors(country):
    return COUNTRY_COLORS.get(country, DEFAULT_ACCENT)


# ============================================================
# SIDEBAR
# ============================================================

st.sidebar.markdown(
    """
    <div style="
        font-size:28px;
        font-weight:850;
        margin-bottom:3px;
    ">
        AMECATH
    </div>

    <div style="
        font-size:11px;
        color:#AEB3BA;
        margin-bottom:25px;
    ">
        MENA DIALYSIS MARKET INTELLIGENCE
    </div>
    """,
    unsafe_allow_html=True,
)

page = st.sidebar.radio(
    "NAVIGATION",
    [
        "🌍 Regional Overview",
        "🇸🇦 Country Explorer",
        "📊 Data Quality",
        "📚 Sources",
    ],
)

st.sidebar.markdown("---")

st.sidebar.caption("2026 Market Intelligence")
st.sidebar.caption("Data period: 2025–2026")


# ============================================================
# HEADER
# ============================================================

def header(title, subtitle):
    st.markdown(
        f"""
        <div class="main-title">{title}</div>
        <div class="sub-title">{subtitle}</div>
        """,
        unsafe_allow_html=True,
    )


# ============================================================
# KPI
# ============================================================

def kpi(label, value):
    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-label">{label}</div>
            <div class="kpi-value">{value}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


# ============================================================
# COUNTRY THEME (flag colors + landscape background)
# ============================================================

def apply_country_theme(country):
    """
    Re-colors the Country Explorer page using the selected country's flag
    colors, and sets that country's landscape photo as the page background.
    Only injected while Country Explorer is the active page, so other pages
    are unaffected.
    """

    colors = get_country_colors(country)
    landscape_path = get_country_landscape_path(country)

    background_css = ""

    if landscape_path:
        b64 = image_to_base64(landscape_path)
        background_css = f"""
        [data-testid="stAppViewContainer"] > .main {{
            background-image:
                linear-gradient(rgba(15, 23, 42, 0.90), rgba(15, 23, 42, 0.94)),
                url("data:image/jpeg;base64,{b64}");
            background-size: cover;
            background-position: center;
            background-attachment: fixed;
        }}
        """

    st.markdown(
        f"""
        <style>
        {background_css}

        .section-title {{
            color: {colors['secondary']} !important;
            border-left: 4px solid {colors['primary']};
            padding-left: 10px;
        }}

        .kpi-card, .action-card, .score-box {{
            background: rgba(15, 23, 42, 0.72) !important;
            border: 1px solid {colors['primary']} !important;
        }}

        .kpi-value, .action-value, .score-number {{
            color: {colors['secondary']} !important;
        }}

        .hero-country {{
            color: {colors['secondary']};
        }}

        .stTabs [aria-selected="true"] {{
            color: {colors['secondary']} !important;
            border-bottom-color: {colors['primary']} !important;
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )


# ============================================================
# REGIONAL OVERVIEW
# ============================================================

def regional_overview():

    header(
        "AMECATH",
        "2026 MENA DIALYSIS MARKET INTELLIGENCE • 9-COUNTRY COMMERCIAL OPPORTUNITY",
    )

    if summary.empty:
        st.warning("No country summary data found.")
        return

    total_population = summary["Population 2026 Num"].sum()
    total_hd = summary["Est. 2026 HD Num"].sum()
    total_machines = summary["HD Machines Num"].sum()
    total_demand = summary["Annual Catheter Demand Num"].sum()
    total_market = summary["Market Value Num"].sum()

    row = st.columns(5)

    with row[0]:
        kpi("Population", format_number(total_population))

    with row[1]:
        kpi("HD Patients", format_number(total_hd))

    with row[2]:
        kpi("HD Machines", format_number(total_machines))

    with row[3]:
        kpi("Catheter Demand / Year", format_number(total_demand))

    with row[4]:
        kpi("Market Value", format_currency(total_market))

    st.markdown(
        '<div class="section-title">Where is the opportunity?</div>',
        unsafe_allow_html=True,
    )

    left, right = st.columns(2)

    chart_df = summary.copy()

    chart_df["Market Value Chart"] = chart_df["Market Value Num"]
    chart_df["Demand Chart"] = chart_df["Annual Catheter Demand Num"]

    with left:
        fig = px.bar(
            chart_df.sort_values("Market Value Chart"),
            x="Market Value Chart",
            y="Country Clean",
            orientation="h",
            title="Catheter Market Value by Country",
            labels={
                "Market Value Chart": "Market Value (USD)",
                "Country Clean": "",
            },
        )

        fig.update_layout(
            height=430,
            showlegend=False,
            margin=dict(l=10, r=10, t=50, b=10),
        )

        st.plotly_chart(fig, use_container_width=True)

    with right:
        fig = px.bar(
            chart_df.sort_values("Demand Chart"),
            x="Demand Chart",
            y="Country Clean",
            orientation="h",
            title="Annual Catheter Demand",
            labels={
                "Demand Chart": "Units / Year",
                "Country Clean": "",
            },
        )

        fig.update_layout(
            height=430,
            showlegend=False,
            margin=dict(l=10, r=10, t=50, b=10),
        )

        st.plotly_chart(fig, use_container_width=True)

    # --------------------------------------------------------
    # OPPORTUNITY SCORE
    # --------------------------------------------------------

    st.markdown(
        '<div class="section-title">MENA Opportunity Map</div>',
        unsafe_allow_html=True,
    )

    score_df = summary.copy()

    # Transparent scoring model based ONLY on workbook metrics.
    score_df["Demand Score"] = (
        score_df["Annual Catheter Demand Num"]
        / score_df["Annual Catheter Demand Num"].max()
        * 40
    )

    score_df["Market Score"] = (
        score_df["Market Value Num"]
        / score_df["Market Value Num"].max()
        * 30
    )

    score_df["Growth Score"] = (
        score_df["Annual Growth Num"]
        / score_df["Annual Growth Num"].max()
        * 20
    )

    score_df["Machine Score"] = (
        score_df["HD Machines Num"]
        / score_df["HD Machines Num"].max()
        * 10
    )

    score_df["Opportunity Score"] = (
        score_df["Demand Score"]
        + score_df["Market Score"]
        + score_df["Growth Score"]
        + score_df["Machine Score"]
    ).round(0)

    fig = px.bar(
        score_df.sort_values("Opportunity Score"),
        x="Opportunity Score",
        y="Country Clean",
        orientation="h",
        text="Opportunity Score",
        title="AMECATH Opportunity Score / 100",
    )

    fig.update_layout(
        height=420,
        xaxis=dict(range=[0, 100]),
        showlegend=False,
    )

    st.plotly_chart(fig, use_container_width=True)

    # --------------------------------------------------------
    # TOP 3 ACTIONS
    # --------------------------------------------------------

    highest_market = (
        summary.loc[
            summary["Market Value Num"].idxmax(),
            "Country Clean"
        ]
    )

    highest_demand = (
        summary.loc[
            summary["Annual Catheter Demand Num"].idxmax(),
            "Country Clean"
        ]
    )

    highest_growth = (
        summary.loc[
            summary["Annual Growth Num"].idxmax(),
            "Country Clean"
        ]
    )

    st.markdown(
        '<div class="section-title">What should AMECATH do?</div>',
        unsafe_allow_html=True,
    )

    c1, c2, c3 = st.columns(3)

    with c1:
        st.markdown(
            f"""
            <div class="action-card">
                <div class="action-title">🔴 Priority Market</div>
                <div class="action-value">{highest_market}</div>
                <div class="action-description">
                    Highest market value in the workbook.
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with c2:
        st.markdown(
            f"""
            <div class="action-card">
                <div class="action-title">📦 Volume Opportunity</div>
                <div class="action-value">{highest_demand}</div>
                <div class="action-description">
                    Highest annual catheter demand.
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with c3:
        st.markdown(
            f"""
            <div class="action-card">
                <div class="action-title">📈 Growth Opportunity</div>
                <div class="action-value">{highest_growth}</div>
                <div class="action-description">
                    Highest annual growth rate.
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )


# ============================================================
# COUNTRY SELECTOR
# ============================================================

def country_selector():

    selected = st.sidebar.selectbox(
        "COUNTRY",
        COUNTRIES,
        index=0,
    )

    return COUNTRY_NAMES[selected]


# ============================================================
# COUNTRY DATA
# ============================================================

def get_country_row(country):

    if summary.empty:
        return None

    result = summary[
        summary["Country Clean"].str.lower() == country.lower()
    ]

    if result.empty:
        return None

    return result.iloc[0]


# ============================================================
# HERO
# ============================================================

def country_hero(country, row):

    image_path = get_country_landscape_path(country)
    flag_path = get_country_flag_path(country)

    description = (
        "Regional dialysis market intelligence and commercial opportunity."
    )

    if row is not None:
        if row.get("Annual Catheter Demand Num") is not None:
            description = (
                f"{format_number(row['Annual Catheter Demand Num'])} "
                "estimated annual catheter demand • "
                f"{format_currency(row['Market Value Num'])} market value"
            )

    flag_html = ""

    if flag_path:
        flag_html = (
            f'<img class="hero-flag" '
            f'src="data:image/jpeg;base64,{image_to_base64(flag_path)}">'
        )

    if image_path:

        st.markdown(
            f"""
            <div class="hero">
                <img src="data:image/jpeg;base64,{image_to_base64(image_path)}">
                {flag_html}
                <div class="hero-text">
                    <div class="hero-country">
                        {COUNTRY_FLAGS.get(country, "")} {country}
                    </div>
                    <div class="hero-description">
                        {description}
                    </div>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    else:
        st.markdown(
            f"""
            <div style="
                background:#24272C;
                padding:40px;
                border-radius:16px;
                color:white;
                margin-bottom:18px;
                position:relative;
            ">
                {flag_html}
                <div style="font-size:34px;font-weight:850;">
                    {COUNTRY_FLAGS.get(country, "")} {country}
                </div>
                <div style="margin-top:8px;color:#D2D5D9;">
                    {description}
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )


# ============================================================
# COUNTRY KPI ROW
# ============================================================

def country_kpis(row):

    if row is None:
        return

    cols = st.columns(6)

    values = [
        ("Population", format_number(row["Population 2026 Num"])),
        ("HD Patients", format_number(row["Est. 2026 HD Num"])),
        ("PD Patients", format_number(row["Est. 2026 PD Num"])),
        ("HD Machines", format_number(row["HD Machines Num"])),
        ("Catheter Demand", format_number(row["Annual Catheter Demand Num"])),
        ("Market Value", format_currency(row["Market Value Num"])),
    ]

    for col, (label, value) in zip(cols, values):
        with col:
            kpi(label, value)


# ============================================================
# COUNTRY OVERVIEW
# ============================================================

def country_overview(country, row):

    st.markdown(
        '<div class="section-title">Country Overview</div>',
        unsafe_allow_html=True,
    )

    if row is None:
        st.warning("Country data not available.")
        return

    c1, c2 = st.columns(2)

    with c1:

        growth = row["Annual Growth Num"]

        fig = go.Figure()

        if growth is not None:
            base = row["Est. 2026 HD Num"]

            years = [2026, 2027, 2028, 2029]
            values = [
                base,
                base * (1 + growth),
                base * (1 + growth) ** 2,
                base * (1 + growth) ** 3,
            ]

            fig.add_trace(
                go.Scatter(
                    x=years,
                    y=values,
                    mode="lines+markers",
                    name="HD Patients",
                )
            )

        fig.update_layout(
            title="HD Patient Trend",
            height=350,
            xaxis_title="Year",
            yaxis_title="Patients",
        )

        st.plotly_chart(fig, use_container_width=True)

    with c2:

        market = row["Market Value Num"]
        demand = row["Annual Catheter Demand Num"]
        machines = row["HD Machines Num"]

        score = 0

        if market:
            score += (market / summary["Market Value Num"].max()) * 30

        if demand:
            score += (demand / summary["Annual Catheter Demand Num"].max()) * 40

        if growth:
            score += (growth / summary["Annual Growth Num"].max()) * 20

        if machines:
            score += (machines / summary["HD Machines Num"].max()) * 10

        score = min(round(score), 100)

        colors = get_country_colors(country)

        st.markdown(
            f"""
            <div class="score-box">
                <div style="color:#94A3B8;font-size:13px;font-weight:700;">
                    COUNTRY OPPORTUNITY SCORE
                </div>

                <div class="score-number">
                    {score}/100
                </div>

                <div style="
                    background:rgba(148,163,184,0.25);
                    height:10px;
                    border-radius:10px;
                    margin-top:10px;
                ">
                    <div style="
                        background:{colors['primary']};
                        width:{score}%;
                        height:10px;
                        border-radius:10px;
                    "></div>
                </div>

                <div style="
                    color:#94A3B8;
                    font-size:12px;
                    margin-top:10px;
                ">
                    Derived from demand, market value, growth and installed
                    HD-machine base in the workbook.
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.markdown("### Key conclusions")

    largest_demand_country = (
        summary.loc[
            summary["Annual Catheter Demand Num"].idxmax(),
            "Country Clean"
        ]
    )

    largest_market_country = (
        summary.loc[
            summary["Market Value Num"].idxmax(),
            "Country Clean"
        ]
    )

    st.markdown(
        f"""
        1. **Largest regional demand:** {largest_demand_country}

        2. **Largest market value:** {largest_market_country}

        3. **Selected country annual growth:** {format_percent(row["Annual Growth Num"])}
        """
    )


# ============================================================
# MARKET TAB
# ============================================================

def market_tab(country, row):

    st.markdown("### Market")

    if row is None:
        return

    cols = st.columns(5)

    metrics = [
        ("HD Patients", format_number(row["Est. 2026 HD Num"])),
        ("PD Patients", format_number(row["Est. 2026 PD Num"])),
        ("HD Growth", format_percent(row["Annual Growth Num"])),
        ("Catheter Demand", format_number(row["Annual Catheter Demand Num"])),
        ("Unit Growth", format_percent(row["Unit Growth Num"])),
    ]

    for col, (label, value) in zip(cols, metrics):
        with col:
            kpi(label, value)

    hd = row["Est. 2026 HD Num"] or 0
    pd_patients = row["Est. 2026 PD Num"] or 0

    mix = pd.DataFrame(
        {
            "Type": ["HD", "PD"],
            "Patients": [hd, pd_patients],
        }
    )

    fig = px.pie(
        mix,
        names="Type",
        values="Patients",
        hole=0.60,
        title="Dialysis Patient Mix",
    )

    fig.update_layout(height=380)

    st.plotly_chart(fig, use_container_width=True)

    st.markdown("### Catheter Demand Model")

    st.markdown(
        """
        <div style="
            background:rgba(30, 41, 59, 0.65);
            border:1px solid rgba(148, 163, 184, 0.25);
            border-radius:12px;
            padding:25px;
            text-align:center;
            font-size:18px;
            color:#F8FAFC;
        ">
            HD Population
            ↓
            Catheter Utilization
            ↓
            Replacement Frequency
            ↓
            Incident Patients
            ↓
            Acute Demand
            ↓
            <b>Annual Catheter Demand</b>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.metric(
        "Annual Catheter Demand",
        format_number(row["Annual Catheter Demand Num"]),
    )

    st.caption(
        "BASE CASE | Derived from the uploaded workbook's country-level estimate."
    )


# ============================================================
# INFRASTRUCTURE TAB
# ============================================================

def infrastructure_tab(country, row):

    st.markdown("### Infrastructure")

    if row is None:
        return

    c1, c2, c3 = st.columns(3)

    with c1:
        kpi(
            "HD Machines",
            format_number(row["HD Machines Num"]),
        )

    with c2:
        kpi(
            "HD Patients",
            format_number(row["Est. 2026 HD Num"]),
        )

    with c3:
        kpi(
            "Hospital Growth",
            format_percent(row["Hospital Growth Num"]),
        )

    hot = data.get("Hot Areas", pd.DataFrame())

    if not hot.empty:

        country_col = None

        for col in hot.columns:
            if country.lower() in str(col).lower():
                country_col = col
                break

        if country_col:

            st.markdown("### Hot Areas")

            rows = []

            for i, value in enumerate(
                hot[country_col].dropna().tolist(),
                start=1
            ):
                rows.append(
                    {
                        "Rank": i,
                        "Hot Area": value,
                    }
                )

            st.dataframe(
                pd.DataFrame(rows),
                use_container_width=True,
                hide_index=True,
            )

    st.info(
        "The workbook does not provide a standardized government/private "
        "facility split for every country, so the dashboard does not invent one."
    )


# ============================================================
# COMMERCIAL TAB
# ============================================================

def commercial_tab(country):

    st.markdown("### Commercial Intelligence")

    distributors = data.get("Distributors", pd.DataFrame()).copy()
    kols = data.get("KOLS", pd.DataFrame()).copy()

    # ---------------- DISTRIBUTORS ----------------

    st.markdown("#### Top Distributors")

    if not distributors.empty:

        # Find country section.
        country_idx = None

        for idx, row in distributors.iterrows():

            first_value = clean_text(row.iloc[0])

            if country.upper() in first_value.upper():
                country_idx = idx
                break

        if country_idx is not None:

            header_idx = country_idx + 1

            if header_idx < len(distributors):

                columns = [
                    clean_text(x)
                    for x in distributors.iloc[header_idx].tolist()
                ]

                rows = distributors.iloc[header_idx + 1:].copy()
                rows.columns = columns

                rows = rows.dropna(how="all")

                st.dataframe(
                    rows.head(10),
                    use_container_width=True,
                    hide_index=True,
                )

        else:
            st.warning("Distributor data for this country was not found.")

    # ---------------- KOLS ----------------

    st.markdown("#### Top KOLs")

    if not kols.empty:

        country_idx = None

        for idx, row in kols.iterrows():

            first_value = clean_text(row.iloc[0])

            if country.upper() in first_value.upper():
                country_idx = idx
                break

        if country_idx is not None:

            header_idx = country_idx + 1

            if header_idx < len(kols):

                columns = [
                    clean_text(x)
                    for x in kols.iloc[header_idx].tolist()
                ]

                rows = kols.iloc[header_idx + 1:].copy()
                rows.columns = columns

                rows = rows.dropna(how="all")

                st.dataframe(
                    rows.head(10),
                    use_container_width=True,
                    hide_index=True,
                )

        else:
            st.warning("KOL data for this country was not found.")


# ============================================================
# COMPETITION TAB
# ============================================================

def competition_tab(country):

    st.markdown("### Competitive Landscape")

    competitors = data.get("Compititors", pd.DataFrame()).copy()

    if competitors.empty:
        st.warning("Competitor data unavailable.")
        return

    country_idx = None

    for idx, row in competitors.iterrows():

        first_value = clean_text(row.iloc[0])

        if country.upper() in first_value.upper():
            country_idx = idx
            break

    if country_idx is None:
        st.warning(f"No competitor data found for {country}.")
        return

    header_idx = country_idx + 1

    if header_idx >= len(competitors):
        return

    columns = [
        clean_text(x)
        for x in competitors.iloc[header_idx].tolist()
    ]

    rows = competitors.iloc[header_idx + 1:].copy()
    rows.columns = columns
    rows = rows.dropna(how="all")

    # Rename to safer names.
    rename_map = {}

    for col in rows.columns:
        low = col.lower()

        if "competitor" in low:
            rename_map[col] = "Competitor"

        elif "market share" in low:
            rename_map[col] = "Market Share"

        elif "coverage" in low:
            rename_map[col] = "Coverage"

        elif "weakness" in low or "gap" in low:
            rename_map[col] = "Weakness / Gap"

        elif "advantage" in low and "main" in low:
            rename_map[col] = "Main Advantage"

        elif "special" in low:
            rename_map[col] = "Specializes In"

        elif "amecath" in low:
            rename_map[col] = "AMECATH Competitive Advantage"

    rows = rows.rename(columns=rename_map)

    display_cols = [
        c for c in [
            "Competitor",
            "Market Share",
            "Coverage",
            "Weakness / Gap",
            "Main Advantage",
            "Specializes In",
            "AMECATH Competitive Advantage",
        ]
        if c in rows.columns
    ]

    if display_cols:
        st.dataframe(
            rows[display_cols].head(10),
            use_container_width=True,
            hide_index=True,
        )
    else:
        st.dataframe(
            rows.head(10),
            use_container_width=True,
            hide_index=True,
        )

    st.markdown("### AMECATH Position")

    position = pd.DataFrame(
        {
            "Dimension": [
                "Price",
                "Availability",
                "Brand",
                "Clinical Evidence",
                "Customization",
            ],
            "Assessment": [
                "Competitive",
                "Competitive",
                "Developing",
                "Developing",
                "Competitive",
            ],
        }
    )

    st.dataframe(
        position,
        use_container_width=True,
        hide_index=True,
    )


# ============================================================
# TENDERS TAB
# ============================================================

def tenders_tab(country):

    st.markdown("### Tender Intelligence")

    tenders = data.get("tenders", pd.DataFrame()).copy()

    if tenders.empty:
        st.warning("Tender data unavailable.")
        return

    tenders["Country Clean"] = tenders["Country"].apply(country_without_flag)

    filtered = tenders[
        tenders["Country Clean"].str.lower() == country.lower()
    ].copy()

    c1, c2, c3 = st.columns(3)

    tenders_2025 = tenders[
        tenders["Published"].astype(str).str.contains("2025", na=False)
    ]

    tenders_2026 = tenders[
        tenders["Published"].astype(str).str.contains("2026", na=False)
    ]

    with c1:
        kpi("2025 Tenders", len(tenders_2025))

    with c2:
        kpi("2026 Tenders", len(tenders_2026))

    with c3:
        kpi("Selected Country", len(filtered))

    if filtered.empty:
        st.info("No tender rows found for this country.")
        return

    display_cols = [
        c for c in [
            "Tender Title (Short)",
            "Tender Ref / ID",
            "Issuing Entity",
            "Published",
            "Closing Date",
            "Tender Value (USD)",
            "Winner / Awarded To",
            "Notes (scope)",
            "Link",
        ]
        if c in filtered.columns
    ]

    st.dataframe(
        filtered[display_cols],
        use_container_width=True,
        hide_index=True,
    )

    st.markdown("### Pricing Evidence")

    st.markdown(
        """
        <span class="badge badge-green">Verified tender price</span>
        &nbsp;
        <span class="badge badge-yellow">Derived</span>
        &nbsp;
        <span class="badge badge-red">Estimated</span>
        &nbsp;
        <span class="badge badge-gray">Not disclosed</span>
        """,
        unsafe_allow_html=True,
    )

    st.caption(
        "Tender prices are not inferred where the workbook states that pricing "
        "is undisclosed."
    )


# ============================================================
# STRATEGY TAB
# ============================================================

def strategy_tab(country, row):

    st.markdown("### AMECATH Country Playbook")

    asp = data.get("ASP", pd.DataFrame())

    c1, c2, c3 = st.columns(3)

    with c1:
        st.markdown(
            """
            <div class="action-card">
                <div class="action-title">🎯 TARGET</div>
                <div class="action-value">Top Accounts</div>
                <div class="action-description">
                    Prioritize the highest-value dialysis clusters.
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with c2:
        st.markdown(
            """
            <div class="action-card">
                <div class="action-title">🤝 PARTNER</div>
                <div class="action-value">Distributor Network</div>
                <div class="action-description">
                    Evaluate tender access and geographic coverage.
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with c3:
        st.markdown(
            """
            <div class="action-card">
                <div class="action-title">👨‍⚕️ KOL</div>
                <div class="action-value">Top 3 KOLs</div>
                <div class="action-description">
                    Build clinical influence around dialysis access.
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    c4, c5, c6 = st.columns(3)

    with c4:
        st.markdown(
            """
            <div class="action-card">
                <div class="action-title">📑 TENDER</div>
                <div class="action-value">Next Opportunity</div>
                <div class="action-description">
                    Monitor upcoming public and hospital tenders.
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with c5:

        asp_value = "N/A"

        if not asp.empty:
            asp_country = asp[
                asp["Country"].apply(country_without_flag).str.lower()
                == country.lower()
            ]

            if not asp_country.empty:
                asp_value = clean_text(
                    asp_country.iloc[0]["Indicative HD catheter ASP (USD)"]
                )

        st.markdown(
            f"""
            <div class="action-card">
                <div class="action-title">💰 PRICE</div>
                <div class="action-value">{asp_value}</div>
                <div class="action-description">
                    Indicative HD catheter ASP range.
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with c6:
        st.markdown(
            """
            <div class="action-card">
                <div class="action-title">🥊 COMPETITOR</div>
                <div class="action-value">Competitive Landscape</div>
                <div class="action-description">
                    Focus on value, responsiveness and HD specialization.
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.markdown("### AMECATH 3-Year Scenario")

    if row is None:
        return

    market = row["Market Value Num"] or 0
    growth = row["Annual Growth Num"] or 0

    years = [2027, 2028, 2029]

    market_forecast = [
        market * (1 + growth),
        market * (1 + growth) ** 2,
        market * (1 + growth) ** 3,
    ]

    share = [0.02, 0.05, 0.08]

    revenue = [
        market_forecast[i] * share[i]
        for i in range(3)
    ]

    forecast = pd.DataFrame(
        {
            "Year": years,
            "Market": market_forecast,
            "AMECATH Share": share,
            "AMECATH Revenue": revenue,
        }
    )

    st.dataframe(
        forecast.assign(
            Market=forecast["Market"].map(format_currency),
            **{
                "AMECATH Share": forecast["AMECATH Share"].map(
                    lambda x: f"{x*100:.0f}%"
                ),
                "AMECATH Revenue": forecast["AMECATH Revenue"].map(
                    format_currency
                ),
            },
        ),
        use_container_width=True,
        hide_index=True,
    )

    fig = px.line(
        forecast,
        x="Year",
        y=["Market", "AMECATH Revenue"],
        markers=True,
        title="Market → AMECATH Revenue Scenario",
    )

    fig.update_layout(height=380)

    st.plotly_chart(fig, use_container_width=True)

    st.warning(
        "The 2% / 5% / 8% AMECATH share path is a scenario assumption, "
        "not a source-derived forecast."
    )


# ============================================================
# COUNTRY EXPLORER
# ============================================================

def country_explorer():

    header(
        "Country Explorer",
        "Country-level market, infrastructure, commercial and competitive intelligence",
    )

    country = country_selector()

    # Theme this page with the selected country's flag colors and
    # landscape background.
    apply_country_theme(country)

    row = get_country_row(country)

    country_hero(country, row)

    country_kpis(row)

    tabs = st.tabs(
        [
            "Overview",
            "Market",
            "Infrastructure",
            "Commercial",
            "Competition",
            "Tenders",
            "Strategy",
        ]
    )

    with tabs[0]:
        country_overview(country, row)

    with tabs[1]:
        market_tab(country, row)

    with tabs[2]:
        infrastructure_tab(country, row)

    with tabs[3]:
        commercial_tab(country)

    with tabs[4]:
        competition_tab(country)

    with tabs[5]:
        tenders_tab(country)

    with tabs[6]:
        strategy_tab(country, row)


# ============================================================
# DATA QUALITY
# ============================================================

def data_quality():

    header(
        "Data Quality",
        "Coverage, completeness and transparency of the AMECATH market intelligence dataset",
    )

    rows = []

    for sheet_name, df in data.items():

        total_cells = df.shape[0] * df.shape[1]

        if total_cells == 0:
            completeness = 0
        else:
            completeness = (
                df.notna().sum().sum()
                / total_cells
            )

        rows.append(
            {
                "Sheet": sheet_name,
                "Rows": df.shape[0],
                "Columns": df.shape[1],
                "Completeness": f"{completeness * 100:.1f}%",
            }
        )

    st.dataframe(
        pd.DataFrame(rows),
        use_container_width=True,
        hide_index=True,
    )

    st.markdown("### Important Data Rules")

    st.markdown(
        """
        - Values marked as **N/D** remain unavailable.
        - Values marked as **estimated** are not presented as verified.
        - Tender prices are not invented when not publicly disclosed.
        - Scenario forecasts are clearly labelled as scenarios.
        - Opportunity scoring is derived from workbook metrics and is not a
          published external market ranking.
        """
    )


# ============================================================
# SOURCES
# ============================================================

def sources_page():

    header(
        "Sources & Dataset",
        "Traceability of the dashboard to the uploaded AMECATH workbook",
    )

    st.markdown(
        """
        ### Primary dataset

        **Amecath Dash(10).xlsx**

        Dashboard source sheets:

        - Sheet3 — Country Summary
        - Hot Areas
        - Distributors
        - Compititors
        - KOLS
        - tenders
        - ASP

        ### Visual assets

        The dashboard loads each country's flag and landscape photo directly
        from the repo (root folder or an `/assets` subfolder).

        ### Important

        The dashboard does not silently replace missing workbook data with
        external assumptions. Where a calculation is generated by the
        application, it is explicitly labelled as **derived** or **scenario**.
        """
    )

    for sheet_name, df in data.items():

        with st.expander(sheet_name):

            st.write(
                f"{df.shape[0]:,} rows × {df.shape[1]:,} columns"
            )

            st.dataframe(
                df.head(20),
                use_container_width=True,
                hide_index=True,
            )


# ============================================================
# ROUTER
# ============================================================

if page == "🌍 Regional Overview":

    regional_overview()

elif page == "🇸🇦 Country Explorer":

    country_explorer()

elif page == "📊 Data Quality":

    data_quality()

elif page == "📚 Sources":

    sources_page()


# ============================================================
# FOOTER
# ============================================================

st.markdown(
    """
    <div class="source-note">
        AMECATH MENA Dialysis Market Intelligence • 2026
        <br>
        Built for executive commercial decision support.
    </div>
    """,
    unsafe_allow_html=True,
)
