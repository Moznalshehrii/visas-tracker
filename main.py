import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from groq import Groq
from dotenv import load_dotenv
import os
import io
import calendar

load_dotenv()

# ---------------------------------------------------------------------------
# Page config
# ---------------------------------------------------------------------------
st.set_page_config(page_title="Visas Tracker 2026", page_icon="🌍", layout="wide")

# ---------------------------------------------------------------------------
# Custom CSS
# ---------------------------------------------------------------------------
st.markdown("""
<style>
    .main-header {
        font-size: 2.2rem;
        font-weight: 700;
        background: linear-gradient(90deg, #1e3a5f, #4a90d9);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin-bottom: 0;
    }
    .sub-header {
        font-size: 1rem;
        color: #6c757d;
        margin-top: -10px;
        margin-bottom: 20px;
    }
    .metric-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 20px;
        border-radius: 12px;
        color: white;
        text-align: center;
    }
    .metric-card h3 {
        margin: 0;
        font-size: 2rem;
        font-weight: 700;
    }
    .metric-card p {
        margin: 5px 0 0 0;
        font-size: 0.9rem;
        opacity: 0.85;
    }
    .stTabs [data-baseweb="tab-list"] { gap: 8px; }
    .stTabs [data-baseweb="tab"] { border-radius: 8px 8px 0 0; padding: 10px 20px; }
    div[data-testid="stChatMessage"] { border-radius: 12px; }
</style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
GROQ_API_KEY = os.getenv("GROQ_API_KEY") 
CHART_TYPES = ["Bar", "Pie", "Donut", "Line", "Area", "Treemap", "Sunburst", "Funnel", "Scatter", "Histogram", "Heatmap"]
COLOR_SCALES = {
    "Default Blue": px.colors.sequential.Blues,
    "Viridis": px.colors.sequential.Viridis,
    "Plasma": px.colors.sequential.Plasma,
    "Inferno": px.colors.sequential.Inferno,
    "Magma": px.colors.sequential.Magma,
    "Cividis": px.colors.sequential.Cividis,
    "Turbo": px.colors.sequential.Turbo,
    "Rainbow": px.colors.qualitative.Set3,
    "Bold": px.colors.qualitative.Bold,
    "Pastel": px.colors.qualitative.Pastel,
    "Sunset": px.colors.sequential.Sunset,
    "Teal": px.colors.sequential.Teal,
    "Berry": px.colors.sequential.Magenta,
    "Earth": px.colors.sequential.Brwnyl,
    "Ice": px.colors.sequential.ice,
}

DARK_BLUE = "#1e3a5f"
NAVY_COLORS = ["#1e3a5f", "#2c5282", "#3182ce", "#4a90d9", "#63b3ed", "#90cdf4"]


# ---------------------------------------------------------------------------
# Helper: read sheets with correct header rows
# ---------------------------------------------------------------------------
@st.cache_data
def load_excel(file_bytes: bytes) -> dict[str, pd.DataFrame]:
    xls = pd.ExcelFile(io.BytesIO(file_bytes), engine="openpyxl")
    sheets: dict[str, pd.DataFrame] = {}

    # Map of key -> (partial sheet name match, header row 0-indexed)
    targets = {
        "business_visit": ("business visit", 2),
        "temp_work": ("temporary work", 3),
        "perm_work": ("permanent work", 3),
    }

    for key, (match_str, hdr_row) in targets.items():
        for sn in xls.sheet_names:
            if match_str in sn.lower():
                df = pd.read_excel(xls, sheet_name=sn, header=hdr_row)
                df.columns = [str(c).strip() for c in df.columns]
                # drop fully empty rows
                df = df.dropna(how="all")
                # drop rows where name column is empty (catches formula-only rows)
                name_col = None
                for candidate in ["Employee Name", "Name"]:
                    if candidate in df.columns:
                        name_col = candidate
                        break
                if name_col is None:
                    for c in df.columns:
                        if "name" in c.lower():
                            name_col = c
                            break
                if name_col:
                    df = df.dropna(subset=[name_col])
                    df = df[df[name_col].astype(str).str.strip() != ""]
                sheets[key] = df
                break

    return sheets


def clean_num(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce")


# ---------------------------------------------------------------------------
# Helper: build a Plotly figure
# ---------------------------------------------------------------------------
def make_chart(df, x, y, chart_type, color_scale, height, title=""):
    cs = color_scale

    if chart_type == "Bar":
        if y:
            fig = px.bar(df, x=x, y=y, color=x, color_discrete_sequence=cs, title=title)
        else:
            c = df[x].value_counts().reset_index(); c.columns = [x, "Count"]
            fig = px.bar(c, x=x, y="Count", color=x, color_discrete_sequence=cs, title=title)
    elif chart_type == "Pie":
        if y:
            fig = px.pie(df, names=x, values=y, color_discrete_sequence=cs, title=title)
        else:
            c = df[x].value_counts().reset_index(); c.columns = [x, "Count"]
            fig = px.pie(c, names=x, values="Count", color_discrete_sequence=cs, title=title)
    elif chart_type == "Donut":
        if y:
            fig = px.pie(df, names=x, values=y, color_discrete_sequence=cs, title=title, hole=0.45)
        else:
            c = df[x].value_counts().reset_index(); c.columns = [x, "Count"]
            fig = px.pie(c, names=x, values="Count", color_discrete_sequence=cs, title=title, hole=0.45)
    elif chart_type == "Line":
        if y:
            fig = px.line(df, x=x, y=y, color_discrete_sequence=cs, title=title, markers=True)
        else:
            c = df[x].value_counts().sort_index().reset_index(); c.columns = [x, "Count"]
            fig = px.line(c, x=x, y="Count", color_discrete_sequence=cs, title=title, markers=True)
    elif chart_type == "Area":
        if y:
            fig = px.area(df, x=x, y=y, color_discrete_sequence=cs, title=title)
        else:
            c = df[x].value_counts().sort_index().reset_index(); c.columns = [x, "Count"]
            fig = px.area(c, x=x, y="Count", color_discrete_sequence=cs, title=title)
    elif chart_type == "Treemap":
        if y:
            fig = px.treemap(df, path=[x], values=y, color_discrete_sequence=cs, title=title)
        else:
            c = df[x].value_counts().reset_index(); c.columns = [x, "Count"]
            fig = px.treemap(c, path=[x], values="Count", color_discrete_sequence=cs, title=title)
    elif chart_type == "Sunburst":
        if y:
            fig = px.sunburst(df, path=[x], values=y, color_discrete_sequence=cs, title=title)
        else:
            c = df[x].value_counts().reset_index(); c.columns = [x, "Count"]
            fig = px.sunburst(c, path=[x], values="Count", color_discrete_sequence=cs, title=title)
    elif chart_type == "Funnel":
        if y:
            fig = px.funnel(df, x=y, y=x, color=x, color_discrete_sequence=cs, title=title)
        else:
            c = df[x].value_counts().reset_index(); c.columns = [x, "Count"]
            fig = px.funnel(c, x="Count", y=x, color=x, color_discrete_sequence=cs, title=title)
    elif chart_type == "Scatter":
        if y:
            fig = px.scatter(df, x=x, y=y, color=x, color_discrete_sequence=cs, title=title, size=y)
        else:
            c = df[x].value_counts().reset_index(); c.columns = [x, "Count"]
            fig = px.scatter(c, x=x, y="Count", color=x, color_discrete_sequence=cs, title=title, size="Count")
    elif chart_type == "Histogram":
        fig = px.histogram(df, x=x, color_discrete_sequence=cs, title=title)
    elif chart_type == "Heatmap":
        if y:
            fig = px.density_heatmap(df, x=x, y=y, color_continuous_scale=cs[::-1] if len(cs) > 2 else "Blues", title=title)
        else:
            fig = px.histogram(df, x=x, color_discrete_sequence=cs, title=title)
    else:
        fig = px.bar(df, x=x, color_discrete_sequence=cs, title=title)

    fig.update_layout(
        height=height, plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
        font=dict(size=12), margin=dict(l=40, r=40, t=60, b=40), title_font_size=16,
    )
    return fig


# ---------------------------------------------------------------------------
# Helpers: extract columns by keyword
# ---------------------------------------------------------------------------
def find_col(df, *keywords):
    for kw in keywords:
        matches = [c for c in df.columns if kw.lower() in c.lower()]
        if matches:
            return matches[0]
    return None


def get_nationality(df):
    col = find_col(df, "national")
    if col is None:
        return pd.Series(dtype=str)
    return df[col].dropna().astype(str).str.strip()


def get_passport(df):
    col = find_col(df, "passport")
    if col is None:
        return pd.Series(dtype=str)
    return df[col].dropna().astype(str).str.strip()


def get_occupation(df):
    col = find_col(df, "occup", "profes")
    if col is None:
        return pd.Series(dtype=str)
    return df[col].dropna().astype(str).str.strip()


def get_date(df):
    col = find_col(df, "issuance", "issue date", "visa issue")
    if col is None:
        return pd.Series(dtype="datetime64[ns]")
    return pd.to_datetime(df[col], errors="coerce").dropna()


def get_name(df):
    col = find_col(df, "employee name", "name")
    if col is None:
        return pd.Series(dtype=str)
    return df[col].dropna().astype(str).str.strip()


# ---------------------------------------------------------------------------
# Groq Chat
# ---------------------------------------------------------------------------
def ask_groq(question: str, data_context: str) -> str:
    try:
        client = Groq(api_key=GROQ_API_KEY)
        resp = client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=[
                {"role": "system", "content": (
                    "You are a helpful data analyst assistant for 2026 visa tracking data. "
                    "The data covers Business Visit Visas, Temporary Work Visas, and Permanent Work Visas for 2026. "
                    "Answer accurately. Use markdown. Be precise with numbers."
                )},
                {"role": "user", "content": f"2026 visa data:\n\n{data_context}\n\nQuestion: {question}"},
            ],
            temperature=0.2, max_tokens=2048,
        )
        return resp.choices[0].message.content
    except Exception as e:
        return f"Error: {str(e)}"


def build_data_context(sheets):
    parts = []
    labels = {"business_visit": "Business Visit Visa 2026", "temp_work": "Temporary Work Visa 2026", "perm_work": "Permanent Work Visa 2026"}
    for key, df in sheets.items():
        name = labels.get(key, key)
        parts.append(f"## {name}\nRows: {len(df)}\nColumns: {', '.join(df.columns.tolist())}\n")
        sample = df.head(200).to_csv(index=False)
        if len(sample) > 12000:
            sample = sample[:12000] + "\n... (truncated)"
        parts.append(sample)
    return "\n\n".join(parts)


# ===========================================================================
# MAIN APP
# ===========================================================================
st.markdown('<p class="main-header">Visas Tracker 2026</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-header">Upload your Excel file to explore 2026 visa data with interactive charts and AI chat</p>', unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# Sidebar
# ---------------------------------------------------------------------------
with st.sidebar:
    st.image("https://img.icons8.com/fluency/96/passport.png", width=64)
    st.markdown("### Settings")
    uploaded = st.file_uploader("Upload Visas Tracker (.xlsx)", type=["xlsx"])
    st.divider()
    st.markdown("### Chart Preferences")
    default_chart = st.selectbox("Default chart type", CHART_TYPES, index=0)
    default_color_name = st.selectbox("Color palette", list(COLOR_SCALES.keys()), index=0)
    default_color = COLOR_SCALES[default_color_name]
    chart_height = st.slider("Chart height (px)", 300, 800, 450, step=50)
    st.divider()
    st.caption("Built with Streamlit, Plotly & Groq AI")

# ---------------------------------------------------------------------------
# No file
# ---------------------------------------------------------------------------
if uploaded is None:
    st.info("Upload your **Visas Tracker .xlsx** file in the sidebar to get started.")
    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown('<div class="metric-card"><h3>--</h3><p>Business Visit Visas</p></div>', unsafe_allow_html=True)
    with c2:
        st.markdown('<div class="metric-card"><h3>--</h3><p>Temporary Work Visas</p></div>', unsafe_allow_html=True)
    with c3:
        st.markdown('<div class="metric-card"><h3>--</h3><p>Permanent Work Visas</p></div>', unsafe_allow_html=True)
    st.stop()

# ---------------------------------------------------------------------------
# Load
# ---------------------------------------------------------------------------
file_bytes = uploaded.read()
sheets = load_excel(file_bytes)

if not sheets:
    st.error("Could not find the expected sheets. Check the file format.")
    st.stop()

bv = sheets.get("business_visit")
tw = sheets.get("temp_work")
pw = sheets.get("perm_work")

n_bv = len(bv) if bv is not None else 0
n_tw = len(tw) if tw is not None else 0
n_pw = len(pw) if pw is not None else 0
n_total = n_bv + n_tw + n_pw

# Collect all nationalities, passports, occupations, dates per type
all_nat, all_pass, all_occ = [], [], []
type_labels = {"business_visit": "Business", "temp_work": "Temporary", "perm_work": "Permanent"}
visa_type_list = []
monthly_records = []  # (month_num, type_label)

for key, df in sheets.items():
    label = type_labels[key]
    n = len(df)
    visa_type_list.extend([label] * n)

    nats = get_nationality(df)
    all_nat.extend(nats.tolist())

    passports = get_passport(df)
    all_pass.extend(passports.tolist())

    occs = get_occupation(df)
    all_occ.extend(occs.tolist())
    # for business visit, occupation is not specified
    if key == "business_visit":
        all_occ.extend(["Not Specified"] * (n - len(occs)))

    dates = get_date(df)
    for d in dates:
        monthly_records.append({"Year": d.year, "Month_Num": d.month, "Type": label})

n_nationalities = len(set(all_nat))
n_passports = len(set(all_pass))

# ---------------------------------------------------------------------------
# KPI row
# ---------------------------------------------------------------------------
c1, c2, c3 = st.columns(3)
with c1:
    st.markdown(f'<div class="metric-card"><h3>{n_bv}</h3><p>Business Visit Visas</p></div>', unsafe_allow_html=True)
with c2:
    st.markdown(f'<div class="metric-card"><h3>{n_tw}</h3><p>Temporary Work Visas</p></div>', unsafe_allow_html=True)
with c3:
    st.markdown(f'<div class="metric-card"><h3>{n_pw}</h3><p>Permanent Work Visas</p></div>', unsafe_allow_html=True)

st.markdown("---")

# ---------------------------------------------------------------------------
# Tabs
# ---------------------------------------------------------------------------
tabs = st.tabs(["Overview", "Business Visit", "Temporary Work", "Permanent Work", "AI Chat"])

# ===== TAB 0 : OVERVIEW ====================================================
with tabs[0]:
    st.subheader("2026 Overview Dashboard")

    oc1, oc2, oc3 = st.columns(3)
    with oc1:
        ov_chart = st.selectbox("Chart type", CHART_TYPES, index=0, key="ov_chart")
    with oc2:
        ov_cn = st.selectbox("Color", list(COLOR_SCALES.keys()), index=0, key="ov_color")
        ov_color = COLOR_SCALES[ov_cn]
    with oc3:
        ov_h = st.slider("Height", 300, 800, chart_height, 50, key="ov_h")

    # --- ROW 1: Summary (with numbers on top) ---
    st.markdown("#### Summary")
    sm1, sm2, sm3, sm4 = st.columns(4)
    summary_items = [
        ("Visa Types", 3),
        ("Nationalities", n_nationalities),
        ("Passport Numbers", n_passports),
        ("Visas", n_total),
    ]
    for col, (label, val) in zip([sm1, sm2, sm3, sm4], summary_items):
        with col:
            st.metric(label, val)

    st.markdown("---")

    # --- ROW 2: Nationality | Visa Type ---
    r2c1, r2c2 = st.columns(2)

    with r2c1:
        if all_nat:
            nat_df = pd.DataFrame({"Nationality": all_nat})
            fig = make_chart(nat_df, "Nationality", None, ov_chart, ov_color, ov_h, "Nationality")
            st.plotly_chart(fig, use_container_width=True)

    with r2c2:
        vtype_df = pd.DataFrame({"Visa Type": visa_type_list})
        fig = make_chart(vtype_df, "Visa Type", None, ov_chart, ov_color, ov_h, "Visa Type")
        st.plotly_chart(fig, use_container_width=True)

    # --- ROW 3: Occupations (full width, bigger) ---
    if all_occ:
        occ_df = pd.DataFrame({"Occupation": all_occ})
        fig = make_chart(occ_df, "Occupation", None, ov_chart, ov_color, ov_h + 100, "Occupations")
        st.plotly_chart(fig, use_container_width=True)

    st.markdown("---")

    # --- ROW 4: Redundancy table (with Name) ---
    st.markdown("#### Redundancy")
    pass_nat_records = []
    for key, df in sheets.items():
        p_col = find_col(df, "passport")
        n_col = find_col(df, "national")
        name_c = find_col(df, "employee name", "name")
        if p_col and n_col and name_c:
            for _, row in df[[p_col, n_col, name_c]].dropna().iterrows():
                pass_nat_records.append({
                    "Passport Number": str(row[p_col]).strip(),
                    "Name": str(row[name_c]).strip(),
                    "Nationality": str(row[n_col]).strip(),
                })

    if pass_nat_records:
        pn_df = pd.DataFrame(pass_nat_records)
        pn_counts = pn_df.groupby(["Passport Number", "Name", "Nationality"]).size().reset_index(name="Count")
        redundant = pn_counts[pn_counts["Count"] > 1].sort_values("Count", ascending=False).reset_index(drop=True)
        if not redundant.empty:
            st.dataframe(redundant, use_container_width=True, hide_index=True)
        else:
            st.info("No duplicate passport numbers found.")
    else:
        st.info("No passport data available.")

    st.markdown("---")

    # --- ROW 3: Monthly Visa Issuance (Oct 2025 – Dec 2026) ---
    st.markdown("#### Monthly Visa Issuance")
    # Build month slots from Oct 2025 to Dec 2026
    timeline = []
    for m in range(10, 13):  # Oct, Nov, Dec 2025
        timeline.append((2025, m))
    for m in range(1, 13):   # Jan–Dec 2026
        timeline.append((2026, m))
    timeline_labels = [f"{calendar.month_abbr[m].upper()} {y}" for y, m in timeline]

    monthly_df = pd.DataFrame(monthly_records) if monthly_records else pd.DataFrame(columns=["Year", "Month_Num", "Type"])
    fig = go.Figure()
    for vtype in ["Business", "Temporary", "Permanent"]:
        subset = monthly_df[monthly_df["Type"] == vtype] if not monthly_df.empty else pd.DataFrame()
        counts = []
        for y, m in timeline:
            if not subset.empty:
                cnt = len(subset[(subset["Year"] == y) & (subset["Month_Num"] == m)])
            else:
                cnt = 0
            counts.append(cnt)
        fig.add_trace(go.Scatter(
            x=timeline_labels, y=counts, mode="lines+markers", name=vtype,
            line=dict(width=2), marker=dict(size=8),
        ))

    fig.update_layout(
        title="MONTHLY VISA ISSUANCE (OCT 2025 – DEC 2026)",
        xaxis_title="", yaxis_title="VISAS ISSUED",
        height=ov_h + 50, plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
        margin=dict(l=40, r=40, t=60, b=40), title_font_size=16,
        legend=dict(orientation="h", yanchor="bottom", y=-0.2, xanchor="center", x=0.5),
    )
    st.plotly_chart(fig, use_container_width=True)


# ===== TAB 1 : BUSINESS VISIT ==============================================
with tabs[1]:
    if bv is not None:
        st.subheader("Business Visit Visa 2026")
        bc1, bc2, bc3 = st.columns(3)
        with bc1:
            bv_chart = st.selectbox("Chart type", CHART_TYPES, index=0, key="bv_ct")
        with bc2:
            bv_cn = st.selectbox("Color", list(COLOR_SCALES.keys()), index=0, key="bv_cl")
            bv_c = COLOR_SCALES[bv_cn]
        with bc3:
            bv_h = st.slider("Height", 300, 800, chart_height, 50, key="bv_h")

        st.dataframe(bv, use_container_width=True, height=300)

        b1, b2 = st.columns(2)
        nc = find_col(bv, "national")
        if nc:
            with b1:
                fig = make_chart(bv, nc, None, bv_chart, bv_c, bv_h, "By Nationality")
                st.plotly_chart(fig, use_container_width=True)
        rc = find_col(bv, "requester")
        if rc:
            with b2:
                fig = make_chart(bv, rc, None, bv_chart, bv_c, bv_h, "By Requester")
                st.plotly_chart(fig, use_container_width=True)
        hc = find_col(bv, "handle")
        if hc:
            fig = make_chart(bv, hc, None, bv_chart, bv_c, bv_h, "By Handler")
            st.plotly_chart(fig, use_container_width=True)
        cc = find_col(bv, "collect")
        if cc:
            fig = make_chart(bv, cc, None, bv_chart, bv_c, bv_h, "By Collection City")
            st.plotly_chart(fig, use_container_width=True)
        dc = find_col(bv, "issuance", "date")
        if dc:
            bv_d = bv.copy()
            bv_d[dc] = pd.to_datetime(bv_d[dc], errors="coerce")
            bv_d = bv_d.dropna(subset=[dc])
            if not bv_d.empty:
                bv_d["Month"] = bv_d[dc].dt.to_period("M").astype(str)
                fig = make_chart(bv_d, "Month", None,
                                 "Line" if bv_chart in ["Pie", "Donut", "Treemap", "Sunburst"] else bv_chart,
                                 bv_c, bv_h, "Issuance Trend by Month")
                st.plotly_chart(fig, use_container_width=True)
    else:
        st.warning("Business Visit Visa 2026 sheet not found.")


# ===== TAB 2 : TEMP WORK ===================================================
with tabs[2]:
    if tw is not None:
        st.subheader("Temporary Work Visa 2026")
        tc1, tc2, tc3 = st.columns(3)
        with tc1:
            tw_chart = st.selectbox("Chart type", CHART_TYPES, index=0, key="tw_ct")
        with tc2:
            tw_cn = st.selectbox("Color", list(COLOR_SCALES.keys()), index=0, key="tw_cl")
            tw_c = COLOR_SCALES[tw_cn]
        with tc3:
            tw_h = st.slider("Height", 300, 800, chart_height, 50, key="tw_h")

        st.dataframe(tw, use_container_width=True, height=300)

        t1, t2 = st.columns(2)
        nc = find_col(tw, "national")
        if nc:
            with t1:
                fig = make_chart(tw, nc, None, tw_chart, tw_c, tw_h, "By Nationality")
                st.plotly_chart(fig, use_container_width=True)
        oc = find_col(tw, "occup", "profes")
        if oc:
            with t2:
                fig = make_chart(tw, oc, None, tw_chart, tw_c, tw_h, "By Occupation")
                st.plotly_chart(fig, use_container_width=True)

        emb = find_col(tw, "embassy")
        if emb:
            fig = make_chart(tw, emb, None, tw_chart, tw_c, tw_h, "By Embassy")
            st.plotly_chart(fig, use_container_width=True)

        fee_cols = [c for c in tw.columns if "fee" in c.lower()]
        if fee_cols:
            tw_fees = tw[fee_cols].apply(pd.to_numeric, errors="coerce").sum().reset_index()
            tw_fees.columns = ["Fee Type", "Total"]
            tw_fees = tw_fees[tw_fees["Total"] > 0]
            if not tw_fees.empty:
                fig = make_chart(tw_fees, "Fee Type", "Total", tw_chart, tw_c, tw_h, "Fee Breakdown")
                st.plotly_chart(fig, use_container_width=True)
    else:
        st.warning("Temporary Work Visa 2026 sheet not found.")


# ===== TAB 3 : PERMANENT WORK ==============================================
with tabs[3]:
    if pw is not None:
        st.subheader("Permanent Work Visa 2026")
        pc1, pc2, pc3 = st.columns(3)
        with pc1:
            pw_chart = st.selectbox("Chart type", CHART_TYPES, index=0, key="pw_ct")
        with pc2:
            pw_cn = st.selectbox("Color", list(COLOR_SCALES.keys()), index=0, key="pw_cl")
            pw_c = COLOR_SCALES[pw_cn]
        with pc3:
            pw_h = st.slider("Height", 300, 800, chart_height, 50, key="pw_h")

        st.dataframe(pw, use_container_width=True, height=300)

        p1, p2 = st.columns(2)
        nc = find_col(pw, "national")
        if nc:
            with p1:
                fig = make_chart(pw, nc, None, pw_chart, pw_c, pw_h, "By Nationality")
                st.plotly_chart(fig, use_container_width=True)
        pc_col = find_col(pw, "project")
        if pc_col:
            with p2:
                fig = make_chart(pw, pc_col, None, pw_chart, pw_c, pw_h, "By Project")
                st.plotly_chart(fig, use_container_width=True)

        prof = find_col(pw, "profes", "occup")
        if prof:
            fig = make_chart(pw, prof, None, pw_chart, pw_c, pw_h, "By Profession")
            st.plotly_chart(fig, use_container_width=True)

        fee_names = ["MOI Fee", "COC Fee", "MOFA Fee"]
        found_fees = [c for c in pw.columns if any(f.lower() in c.lower() for f in fee_names)]
        if found_fees:
            pw_fees = pw[found_fees].apply(pd.to_numeric, errors="coerce").sum().reset_index()
            pw_fees.columns = ["Fee Type", "Total"]
            pw_fees = pw_fees[pw_fees["Total"] > 0]
            if not pw_fees.empty:
                fig = make_chart(pw_fees, "Fee Type", "Total", pw_chart, pw_c, pw_h, "Fee Breakdown")
                st.plotly_chart(fig, use_container_width=True)

        city = find_col(pw, "city")
        if city:
            fig = make_chart(pw, city, None, pw_chart, pw_c, pw_h, "By City")
            st.plotly_chart(fig, use_container_width=True)
    else:
        st.warning("Permanent Work Visa 2026 sheet not found.")


# ===== TAB 4 : AI CHAT =====================================================
with tabs[4]:
    st.subheader("AI Data Assistant")
    st.caption("Ask any question about your 2026 visa data. Powered by Groq (Llama 3.3 70B).")

    if "data_context" not in st.session_state:
        st.session_state.data_context = build_data_context(sheets)

    if "messages" not in st.session_state:
        st.session_state.messages = [
            {"role": "assistant", "content": "Hello! I'm your 2026 visa data assistant. Ask me anything about Business Visit, Temporary Work, or Permanent Work visas."}
        ]

    for msg in st.session_state.messages:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])

    if prompt := st.chat_input("Ask about your 2026 visa data..."):
        st.session_state.messages.append({"role": "user", "content": prompt})
        with st.chat_message("user"):
            st.markdown(prompt)
        with st.chat_message("assistant"):
            with st.spinner("Thinking..."):
                response = ask_groq(prompt, st.session_state.data_context)
            st.markdown(response)
        st.session_state.messages.append({"role": "assistant", "content": response})
