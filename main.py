import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from groq import Groq
from dotenv import load_dotenv
import os
import io
import re
import calendar
from datetime import datetime
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.lib.colors import HexColor
from reportlab.lib.enums import TA_LEFT, TA_CENTER
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Image, PageBreak, Table, TableStyle,
)

load_dotenv()

st.set_page_config(page_title="Visas Tracker 2026", page_icon="🌍", layout="wide")

st.markdown("""
<style>
    .main-header { font-size: 5rem !important; font-weight: 900 !important; color: #47D7AC !important; margin: 0 0 4px 0 !important; line-height: 1.05 !important; letter-spacing: -1px !important; }
    .sub-header { font-size: 1.1rem !important; color: #8893A5 !important; margin-top: 0 !important; margin-bottom: 24px !important; }
    .stTabs [data-baseweb="tab-panel"] h2,
    .stTabs [data-baseweb="tab-panel"] [data-testid="stHeadingWithActionElements"] h2 { font-size: 2.4rem !important; font-weight: 800 !important; }
    .metric-card { background: linear-gradient(135deg, #47D7AC 0%, #06041F 100%); padding: 20px; border-radius: 12px; color: white; text-align: center; }
    .metric-card h3 { margin: 0; font-size: 2rem; font-weight: 700; }
    .metric-card p { margin: 5px 0 0 0; font-size: 0.9rem; opacity: 0.85; }
    .stTabs [data-baseweb="tab-list"] { gap: 8px; }
    .stTabs [data-baseweb="tab"] { border-radius: 8px 8px 0 0; padding: 10px 20px; }
    div[data-testid="stChatMessage"] { border-radius: 12px; }
</style>
""", unsafe_allow_html=True)

GROQ_API_KEY = os.getenv("GROQ_API_KEY")
CHART_TYPES = ["Bar", "Pie", "Donut", "Line", "Area", "Treemap", "Sunburst", "Funnel", "Scatter", "Histogram", "Heatmap"]
NAGARRO_COLORS = ["#47D7AC", "#FBD872", "#F8485E", "#6240A8", "#8893A5", "#75E1C1"]

# Timeline Oct 2025 – Dec 2026
TIMELINE = [(2025, m) for m in range(10, 13)] + [(2026, m) for m in range(1, 13)]
TIMELINE_LABELS = [f"{calendar.month_abbr[m].upper()} {y}" for y, m in TIMELINE]
LINE_COLORS = {"Business": "#47D7AC", "Temporary": "#FBD872", "Permanent": "#F8485E"}


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
@st.cache_data
def load_excel(file_bytes: bytes) -> dict[str, pd.DataFrame]:
    xls = pd.ExcelFile(io.BytesIO(file_bytes), engine="openpyxl")
    sheets = {}
    targets = {"business_visit": ("business visit", 2), "temp_work": ("temporary work", 3), "perm_work": ("permanent work", 3)}
    for key, (match_str, hdr_row) in targets.items():
        for sn in xls.sheet_names:
            if match_str in sn.lower():
                df = pd.read_excel(xls, sheet_name=sn, header=hdr_row)
                df.columns = [str(c).strip() for c in df.columns]
                df = df.dropna(how="all")
                for col in df.columns:
                    if df[col].dtype == object:
                        df[col] = df[col].apply(lambda v: v.strip() if isinstance(v, str) else v)
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


def drop_time_cols(df):
    if df is None:
        return df
    df = df.drop(columns=[c for c in df.columns if "time" in str(c).lower()], errors="ignore")
    for c in df.columns:
        if pd.api.types.is_datetime64_any_dtype(df[c]):
            df[c] = df[c].dt.date
    return df


def find_col(df, *keywords):
    for kw in keywords:
        matches = [c for c in df.columns if kw.lower() in c.lower()]
        if matches:
            return matches[0]
    return None


def get_nationality(df):
    col = find_col(df, "national")
    return df[col].dropna().astype(str).str.strip() if col else pd.Series(dtype=str)


def get_passport(df):
    col = find_col(df, "passport")
    return df[col].dropna().astype(str).str.strip() if col else pd.Series(dtype=str)


def get_occupation(df):
    col = find_col(df, "occup", "profes")
    return df[col].dropna().astype(str).str.strip() if col else pd.Series(dtype=str)


def get_date(df):
    col = find_col(df, "issuance", "issue date", "visa issue")
    return pd.to_datetime(df[col], errors="coerce").dropna() if col else pd.Series(dtype="datetime64[ns]")


def get_name(df):
    col = find_col(df, "employee name", "name")
    return df[col].dropna().astype(str).str.strip() if col else pd.Series(dtype=str)


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
    fig.update_layout(height=height, plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
                      font=dict(size=12), margin=dict(l=40, r=40, t=60, b=40), title_font_size=16)
    if chart_type in ("Bar", "Funnel", "Scatter", "Histogram"):
        fig.update_xaxes(showticklabels=False, title_text="")
    return fig


def make_expense_line(records, title, height, color="#47D7AC"):
    """Build a line chart from a list of {Year, Month_Num, Cost} records over TIMELINE."""
    df = pd.DataFrame(records) if records else pd.DataFrame(columns=["Year", "Month_Num", "Cost"])
    costs = []
    for y, m in TIMELINE:
        if not df.empty:
            val = df[(df["Year"] == y) & (df["Month_Num"] == m)]["Cost"].sum()
        else:
            val = 0
        costs.append(val)
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=TIMELINE_LABELS, y=costs, mode="lines+markers", name="Cost",
        line=dict(width=3, color=color), marker=dict(size=8, color=color),
        text=[f"{c:,.0f}" for c in costs], textposition="top center",
    ))
    fig.update_layout(title=title, xaxis_title="", yaxis_title="SAR",
                      height=height, plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
                      margin=dict(l=40, r=40, t=60, b=40), title_font_size=16, showlegend=False)
    return fig


def make_multi_expense_line(series_dict, title, height):
    """series_dict: {name: [{Year, Month_Num, Cost}]}"""
    colors = ["#47D7AC", "#FBD872", "#F8485E", "#6240A8"]
    fig = go.Figure()
    for i, (name, records) in enumerate(series_dict.items()):
        df = pd.DataFrame(records) if records else pd.DataFrame(columns=["Year", "Month_Num", "Cost"])
        costs = []
        for y, m in TIMELINE:
            val = df[(df["Year"] == y) & (df["Month_Num"] == m)]["Cost"].sum() if not df.empty else 0
            costs.append(val)
        clr = colors[i % len(colors)]
        fig.add_trace(go.Scatter(
            x=TIMELINE_LABELS, y=costs, mode="lines+markers", name=name,
            line=dict(width=3, color=clr), marker=dict(size=8, color=clr),
        ))
    fig.update_layout(title=title, xaxis_title="", yaxis_title="SAR",
                      height=height, plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
                      margin=dict(l=40, r=40, t=60, b=40), title_font_size=16,
                      legend=dict(orientation="h", yanchor="bottom", y=-0.2, xanchor="center", x=0.5))
    return fig


# ---------------------------------------------------------------------------
# Expense helpers — extract cost records per visa type
# ---------------------------------------------------------------------------
def _safe_num(val):
    """Convert value to float; return 0 for anything non-numeric."""
    v = pd.to_numeric(val, errors="coerce")
    return 0.0 if pd.isna(v) else float(v)


def get_bv_expenses(df):
    """Business Visit: total = COC fee per visa, date = issuance date."""
    date_col = find_col(df, "issuance", "date")
    cost_col = find_col(df, "cost", "fee")
    records = []
    if cost_col:
        for _, row in df.iterrows():
            d = pd.to_datetime(row.get(date_col), errors="coerce") if date_col else pd.NaT
            c = _safe_num(row.get(cost_col, 0))
            if c > 0:
                if pd.isna(d):
                    records.append({"Year": 2026, "Month_Num": 1, "Cost": c})
                else:
                    records.append({"Year": d.year, "Month_Num": d.month, "Cost": c})
    return records


def get_tw_expenses(df):
    """Temporary Work: total = sum of all fee columns, date = Visa Issue Date."""
    date_col = find_col(df, "visa issue", "issue date")
    fee_cols = [c for c in df.columns if "fee" in c.lower()]
    total_col = find_col(df, "total")
    records = []
    for _, row in df.iterrows():
        d = pd.to_datetime(row.get(date_col), errors="coerce") if date_col else pd.NaT
        if total_col:
            cost = _safe_num(row[total_col])
        else:
            cost = sum(_safe_num(row.get(fc, 0)) for fc in fee_cols)
        if cost > 0:
            if pd.isna(d):
                records.append({"Year": 2026, "Month_Num": 1, "Cost": cost})
            else:
                records.append({"Year": d.year, "Month_Num": d.month, "Cost": cost})
    return records


def get_pw_expenses(df):
    """Permanent Work: returns (before_arrival, after_arrival, total) record lists."""
    # Try multiple date column fallbacks
    date_col = find_col(df, "visa issue") or find_col(df, "issue date") or find_col(df, "date")
    before_keys = ["moi fee", "coc fee", "mofa fee", "document shipping"]
    after_keys = ["medical", "work permit", "sce", "iqama", "health insurance"]
    before_cols = [c for c in df.columns if any(k in c.lower() for k in before_keys)]
    after_cols = [c for c in df.columns if any(k in c.lower() for k in after_keys)]
    total_col = find_col(df, "total")

    before_records, after_records, total_records = [], [], []
    for _, row in df.iterrows():
        d = pd.to_datetime(row.get(date_col), errors="coerce") if date_col else pd.NaT
        # Default undated rows to Jan 2026 so they still get counted
        if pd.isna(d):
            ym = {"Year": 2026, "Month_Num": 1}
        else:
            ym = {"Year": d.year, "Month_Num": d.month}
        b = sum(_safe_num(row.get(c, 0)) for c in before_cols)
        a = sum(_safe_num(row.get(c, 0)) for c in after_cols)
        if total_col:
            t = _safe_num(row[total_col])
            if t == 0:
                t = b + a
        else:
            t = b + a
        if b > 0:
            before_records.append({**ym, "Cost": b})
        if a > 0:
            after_records.append({**ym, "Cost": a})
        if t > 0:
            total_records.append({**ym, "Cost": t})
    return before_records, after_records, total_records


# ---------------------------------------------------------------------------
# PDF Summary Report Builder
# ---------------------------------------------------------------------------
NAGARRO_TEAL = HexColor("#47D7AC")
NAGARRO_NAVY = HexColor("#2E008B")
NAGARRO_GREY = HexColor("#4E5E78")


def _md_to_rl(text):
    """Convert a single line of markdown inline formatting to ReportLab markup."""
    # Escape ampersands and angle brackets that aren't part of our tags
    text = text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    # **bold**
    text = re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", text)
    # *italic* or _italic_
    text = re.sub(r"(?<!\*)\*(?!\*)(.+?)(?<!\*)\*(?!\*)", r"<i>\1</i>", text)
    text = re.sub(r"_(.+?)_", r"<i>\1</i>", text)
    # `code` → just bold-ish
    text = re.sub(r"`(.+?)`", r"<font face='Courier'>\1</font>", text)
    return text


def build_pdf_summary(report_md, metrics, chart_figs, logo_path=None):
    """
    Build a PDF summary report.
      report_md : AI-generated markdown text
      metrics   : dict of summary key→value pairs (strings)
      chart_figs: list of (title, plotly.Figure) tuples to embed
      logo_path : optional path to logo image
    Returns: PDF bytes
    """
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=1.8 * cm, rightMargin=1.8 * cm,
        topMargin=1.8 * cm, bottomMargin=1.8 * cm,
        title="Visas Tracker 2026 — Summary Report",
    )

    styles = getSampleStyleSheet()
    h1 = ParagraphStyle("H1", parent=styles["Heading1"], fontSize=20, textColor=NAGARRO_NAVY,
                        spaceAfter=8, alignment=TA_LEFT)
    h2 = ParagraphStyle("H2", parent=styles["Heading2"], fontSize=14, textColor=NAGARRO_NAVY,
                        spaceBefore=14, spaceAfter=6)
    h3 = ParagraphStyle("H3", parent=styles["Heading3"], fontSize=12, textColor=NAGARRO_GREY,
                        spaceBefore=10, spaceAfter=4)
    body = ParagraphStyle("Body", parent=styles["BodyText"], fontSize=10.5, leading=15,
                          textColor=HexColor("#222222"), spaceAfter=6)
    bullet = ParagraphStyle("Bullet", parent=body, leftIndent=16, bulletIndent=6, spaceAfter=3)
    subtitle = ParagraphStyle("Sub", parent=body, fontSize=9, textColor=NAGARRO_GREY)

    story = []

    # --- Header with logo ---
    if logo_path and os.path.exists(logo_path):
        try:
            logo = Image(logo_path, width=3 * cm, height=3 * cm, kind="proportional")
            story.append(logo)
        except Exception:
            pass

    story.append(Paragraph("Visas Tracker 2026", h1))
    story.append(Paragraph("Executive Summary Report", h3))
    story.append(Paragraph(
        f"Generated {datetime.now().strftime('%B %d, %Y at %H:%M')}", subtitle))
    story.append(Spacer(1, 0.4 * cm))

    # --- Key Metrics table ---
    if metrics:
        story.append(Paragraph("Key Metrics", h2))
        rows = [[Paragraph(f"<b>{k}</b>", body), Paragraph(str(v), body)] for k, v in metrics.items()]
        tbl = Table(rows, colWidths=[6 * cm, 10 * cm])
        tbl.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (0, -1), HexColor("#EFF1F4")),
            ("BOX", (0, 0), (-1, -1), 0.5, NAGARRO_GREY),
            ("INNERGRID", (0, 0), (-1, -1), 0.25, HexColor("#C4C9D2")),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING", (0, 0), (-1, -1), 8),
            ("RIGHTPADDING", (0, 0), (-1, -1), 8),
            ("TOPPADDING", (0, 0), (-1, -1), 6),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ]))
        story.append(tbl)
        story.append(Spacer(1, 0.4 * cm))

    # --- AI Report body (markdown → ReportLab) ---
    for raw in (report_md or "").splitlines():
        line = raw.rstrip()
        if not line.strip():
            story.append(Spacer(1, 0.15 * cm))
            continue
        if line.startswith("### "):
            story.append(Paragraph(_md_to_rl(line[4:]), h3))
        elif line.startswith("## "):
            story.append(Paragraph(_md_to_rl(line[3:]), h2))
        elif line.startswith("# "):
            story.append(Paragraph(_md_to_rl(line[2:]), h1))
        elif line.lstrip().startswith(("- ", "* ")):
            item = line.lstrip()[2:]
            story.append(Paragraph(_md_to_rl(item), bullet, bulletText="•"))
        elif re.match(r"^\s*\d+\.\s+", line):
            item = re.sub(r"^\s*\d+\.\s+", "", line)
            story.append(Paragraph(_md_to_rl(item), bullet, bulletText="•"))
        else:
            story.append(Paragraph(_md_to_rl(line), body))

    # --- Charts (ALWAYS embedded — never silently skipped) ---
    if chart_figs:
        story.append(PageBreak())
        story.append(Paragraph("Dashboard Charts", h2))
        story.append(Paragraph(
            f"The following {len(chart_figs)} charts from the Overview dashboard are included.", subtitle))
        story.append(Spacer(1, 0.3 * cm))

        for i, (name, fig) in enumerate(chart_figs):
            # Render chart to high-DPI PNG. If this fails, surface the error
            # loudly — do NOT silently skip (the user explicitly wants charts).
            png = fig.to_image(format="png", width=1400, height=700, scale=2)
            story.append(Paragraph(f"<b>{i + 1}. {name}</b>", h3))
            img = Image(io.BytesIO(png), width=17 * cm, height=8.5 * cm, kind="proportional")
            story.append(img)
            story.append(Spacer(1, 0.5 * cm))
            # Force page break between every 2 charts so they stay readable
            if (i + 1) % 2 == 0 and i < len(chart_figs) - 1:
                story.append(PageBreak())

    doc.build(story)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# Groq Chat
# ---------------------------------------------------------------------------
def ask_groq(question, data_context):
    try:
        client = Groq(api_key=GROQ_API_KEY)
        resp = client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=[
                {"role": "system", "content": "You are a helpful data analyst for 2026 visa tracking data. Answer accurately with markdown."},
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

with st.sidebar:
    st.image("assets/nagarro_logo.png", width=80)
    st.markdown("### Settings")
    uploaded = st.file_uploader("Upload Visas Tracker (.xlsx)", type=["xlsx"])
    st.divider()
    st.markdown("### Chart Preferences")
    default_chart = st.selectbox("Default chart type", CHART_TYPES, index=0)
    default_color = NAGARRO_COLORS
    chart_height = st.slider("Chart height (px)", 300, 800, 450, step=50)
    st.divider()
    st.caption("Built with Streamlit, Plotly & Groq AI")

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

# Collect metadata
all_nat, all_pass, all_occ = [], [], []
type_labels = {"business_visit": "Business", "temp_work": "Temporary", "perm_work": "Permanent"}
visa_type_list = []
monthly_records = []

for key, df in sheets.items():
    label = type_labels[key]
    n = len(df)
    visa_type_list.extend([label] * n)
    all_nat.extend(get_nationality(df).tolist())
    all_pass.extend(get_passport(df).tolist())
    occs = get_occupation(df)
    all_occ.extend(occs.tolist())
    if key == "business_visit":
        all_occ.extend(["Not Specified"] * (n - len(occs)))
    for d in get_date(df):
        monthly_records.append({"Year": d.year, "Month_Num": d.month, "Type": label})

n_nationalities = len(set(all_nat))
n_passports = len(set(all_pass))

# Pre-compute expenses
bv_exp = get_bv_expenses(bv) if bv is not None else []
tw_exp = get_tw_expenses(tw) if tw is not None else []
pw_before, pw_after, pw_total = get_pw_expenses(pw) if pw is not None else ([], [], [])
all_exp = bv_exp + tw_exp + pw_total  # combined

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
tabs = st.tabs(["Overview", "Business Visit", "Temporary Work", "Permanent Work", "Expenses", "AI Chat"])

# ===== TAB 0 : OVERVIEW ====================================================
with tabs[0]:
    st.subheader("2026 Overview Dashboard")
    oc1, oc2 = st.columns(2)
    with oc1:
        ov_chart = st.selectbox("Chart type", CHART_TYPES, index=0, key="ov_chart")
    with oc2:
        ov_h = st.slider("Height", 300, 800, chart_height, 50, key="ov_h")
    ov_color = NAGARRO_COLORS

    # Summary metrics
    st.markdown("#### Summary")
    sm1, sm2, sm3 = st.columns(3)
    for col, (label, val) in zip([sm1, sm2, sm3], [("Nationalities", n_nationalities), ("Passport Numbers", n_passports), ("Visas", n_total)]):
        with col:
            st.metric(label, val)
    st.markdown("---")

    # Nationality | Visa Type
    ov_figs = []  # collect for report
    r2c1, r2c2 = st.columns(2)
    with r2c1:
        if all_nat:
            fig_nat = make_chart(pd.DataFrame({"Nationality": all_nat}), "Nationality", None, ov_chart, ov_color, ov_h, "Nationality")
            st.plotly_chart(fig_nat, use_container_width=True)
            ov_figs.append(("Nationality", fig_nat))
    with r2c2:
        fig_vt = make_chart(pd.DataFrame({"Visa Type": visa_type_list}), "Visa Type", None, ov_chart, ov_color, ov_h, "Visa Type")
        st.plotly_chart(fig_vt, use_container_width=True)
        ov_figs.append(("Visa Type", fig_vt))

    # Professions
    if all_occ:
        fig_occ = make_chart(pd.DataFrame({"Profession": all_occ}), "Profession", None, ov_chart, ov_color, ov_h + 100, "Professions")
        st.plotly_chart(fig_occ, use_container_width=True)
        ov_figs.append(("Professions", fig_occ))
    st.markdown("---")

    # Frequency
    st.markdown("#### Frequency")
    pass_nat_records = []
    for key, df in sheets.items():
        p_col, n_col, name_c = find_col(df, "passport"), find_col(df, "national"), find_col(df, "employee name", "name")
        if p_col and n_col and name_c:
            for _, row in df[[p_col, n_col, name_c]].dropna().iterrows():
                pass_nat_records.append({"Passport Number": str(row[p_col]).strip(), "Name": str(row[name_c]).strip(), "Nationality": str(row[n_col]).strip()})
    if pass_nat_records:
        pn_df = pd.DataFrame(pass_nat_records)
        pn_counts = pn_df.groupby(["Passport Number", "Name", "Nationality"]).size().reset_index(name="Count")
        redundant = pn_counts[pn_counts["Count"] > 1].sort_values("Count", ascending=False).reset_index(drop=True)
        if not redundant.empty:
            st.dataframe(redundant, use_container_width=True, hide_index=True)
        else:
            st.info("No duplicate passport numbers found.")
    st.markdown("---")

    # Monthly Visa Issuance
    st.markdown("#### Monthly Visa Issuance")
    monthly_df = pd.DataFrame(monthly_records) if monthly_records else pd.DataFrame(columns=["Year", "Month_Num", "Type"])
    fig = go.Figure()
    for vtype in ["Business", "Temporary", "Permanent"]:
        subset = monthly_df[monthly_df["Type"] == vtype] if not monthly_df.empty else pd.DataFrame()
        counts = []
        for y, m in TIMELINE:
            counts.append(len(subset[(subset["Year"] == y) & (subset["Month_Num"] == m)]) if not subset.empty else 0)
        fig.add_trace(go.Scatter(x=TIMELINE_LABELS, y=counts, mode="lines+markers", name=vtype,
                                 line=dict(width=3, color=LINE_COLORS[vtype]), marker=dict(size=8, color=LINE_COLORS[vtype])))
    fig.update_layout(title="MONTHLY VISA ISSUANCE (OCT 2025 – DEC 2026)", xaxis_title="", yaxis_title="VISAS ISSUED",
                      height=ov_h + 50, plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
                      margin=dict(l=40, r=40, t=60, b=40), title_font_size=16,
                      legend=dict(orientation="h", yanchor="bottom", y=-0.2, xanchor="center", x=0.5))
    fig_issuance = fig
    st.plotly_chart(fig_issuance, use_container_width=True)
    ov_figs.append(("Monthly Visa Issuance", fig_issuance))
    st.markdown("---")

    # Monthly Expenses — all types on one chart
    st.markdown("#### Monthly Expenses by Visa Type")
    fig_exp = make_multi_expense_line(
        {"Business Visit": bv_exp, "Temporary Work": tw_exp, "Permanent Work": pw_total},
        "MONTHLY EXPENSES BY VISA TYPE (OCT 2025 – DEC 2026)", ov_h + 50)
    st.plotly_chart(fig_exp, use_container_width=True, key="ov_expense_chart")
    ov_figs.append(("Monthly Expenses", fig_exp))

    st.markdown("---")

    # Print Summary — generates a downloadable PDF
    if st.button("Print Summary", type="primary", use_container_width=True, key="print_report"):
        with st.spinner("Generating summary report..."):
            # Build summary data for the AI
            total_bv_cost = sum(r["Cost"] for r in bv_exp)
            total_tw_cost = sum(r["Cost"] for r in tw_exp)
            total_pw_cost = sum(r["Cost"] for r in pw_total)
            grand = total_bv_cost + total_tw_cost + total_pw_cost

            nat_summary = pd.Series(all_nat).value_counts().to_dict() if all_nat else {}
            occ_summary = pd.Series(all_occ).value_counts().to_dict() if all_occ else {}

            redundant_text = ""
            if pass_nat_records:
                pn_df2 = pd.DataFrame(pass_nat_records)
                pn_c2 = pn_df2.groupby(["Passport Number", "Name", "Nationality"]).size().reset_index(name="Count")
                red2 = pn_c2[pn_c2["Count"] > 1]
                if not red2.empty:
                    redundant_text = red2.to_string(index=False)

            report_prompt = f"""Generate a comprehensive executive summary report for the Visas Tracker 2026 dashboard.

Data Summary:
- Total visas: {n_total} (Business Visit: {n_bv}, Temporary Work: {n_tw}, Permanent Work: {n_pw})
- Nationalities: {n_nationalities} — breakdown: {nat_summary}
- Unique passport numbers: {n_passports}
- Professions: {occ_summary}

Expenses:
- Business Visit total: {total_bv_cost:,.0f} SAR
- Temporary Work total: {total_tw_cost:,.0f} SAR
- Permanent Work total: {total_pw_cost:,.0f} SAR
- Grand Total: {grand:,.0f} SAR

Frequent passports (appearing more than once):
{redundant_text if redundant_text else "None found"}

Please write a professional summary report with these sections:
1. Executive Summary
2. Visa Distribution Overview
3. Nationality & Workforce Analysis
4. Expense Analysis
5. Frequency Alerts
6. Key Insights & Recommendations

Format it nicely with markdown headers, bullet points, and bold key numbers."""

            report = ask_groq(report_prompt, "")

            # Build the PDF
            metrics_dict = {
                "Total Visas": f"{n_total}",
                "Business Visit": f"{n_bv}",
                "Temporary Work": f"{n_tw}",
                "Permanent Work": f"{n_pw}",
                "Nationalities": f"{n_nationalities}",
                "Unique Passports": f"{n_passports}",
                "Total Spent": f"{grand:,.0f} SAR",
            }
            logo = os.path.join("assets", "nagarro_logo.png")
            try:
                pdf_bytes = build_pdf_summary(report, metrics_dict, ov_figs, logo_path=logo)
                st.session_state["summary_pdf"] = pdf_bytes
                st.session_state["summary_md"] = report
                st.success(f"Report ready — includes {len(ov_figs)} chart(s). Click the download button below.")
            except Exception as e:
                import traceback
                st.error(f"PDF generation failed: {e}")
                st.code(traceback.format_exc())
                st.session_state["summary_pdf"] = None
                st.session_state["summary_md"] = report

    # Download button + on-screen preview (persists across reruns)
    if st.session_state.get("summary_pdf"):
        fname = f"visas_tracker_summary_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf"
        st.download_button(
            "⬇️ Download Summary PDF",
            data=st.session_state["summary_pdf"],
            file_name=fname,
            mime="application/pdf",
            type="primary",
            use_container_width=True,
            key="dl_summary_pdf",
        )
        st.markdown("---")
        st.markdown("## Summary Report Preview")
        st.markdown(st.session_state.get("summary_md", ""))


# ===== TAB 1 : BUSINESS VISIT ==============================================
with tabs[1]:
    if bv is not None:
        st.subheader("Business Visit Visa 2026")
        bc1, bc2 = st.columns(2)
        with bc1:
            bv_chart = st.selectbox("Chart type", CHART_TYPES, index=0, key="bv_ct")
        with bc2:
            bv_h = st.slider("Height", 300, 800, chart_height, 50, key="bv_h")
        bv_c = NAGARRO_COLORS

        st.dataframe(drop_time_cols(bv), use_container_width=True, height=300, hide_index=True)

        b1, b2 = st.columns(2)
        nc = find_col(bv, "national")
        if nc:
            with b1:
                st.plotly_chart(make_chart(bv, nc, None, bv_chart, bv_c, bv_h, "By Nationality"), use_container_width=True)
        rc = find_col(bv, "requester")
        if rc:
            with b2:
                st.plotly_chart(make_chart(bv, rc, None, bv_chart, bv_c, bv_h, "By Requester"), use_container_width=True)
        hc = find_col(bv, "handle")
        if hc:
            st.plotly_chart(make_chart(bv, hc, None, bv_chart, bv_c, bv_h, "By Handler"), use_container_width=True)
        cc = find_col(bv, "collect")
        if cc:
            st.plotly_chart(make_chart(bv, cc, None, bv_chart, bv_c, bv_h, "By Collection City"), use_container_width=True)
        dc = find_col(bv, "issuance", "date")
        if dc:
            bv_d = bv.copy()
            bv_d[dc] = pd.to_datetime(bv_d[dc], errors="coerce")
            bv_d = bv_d.dropna(subset=[dc])
            if not bv_d.empty:
                bv_d["Month"] = bv_d[dc].dt.to_period("M").astype(str)
                st.plotly_chart(make_chart(bv_d, "Month", None, "Line" if bv_chart in ["Pie", "Donut", "Treemap", "Sunburst"] else bv_chart, bv_c, bv_h, "Issuance Trend by Month"), use_container_width=True)

        # Expense line chart
        st.markdown("---")
        st.markdown("#### Total Expenses (COC Fee)")
        bv_total_cost = sum(r["Cost"] for r in bv_exp)
        st.metric("Total Business Visit Cost", f"{bv_total_cost:,.0f} SAR")
        fig = make_expense_line(bv_exp, "Business Visit Monthly Expenses", bv_h, "#47D7AC")
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.warning("Business Visit Visa 2026 sheet not found.")


# ===== TAB 2 : TEMP WORK ===================================================
with tabs[2]:
    if tw is not None:
        st.subheader("Temporary Work Visa 2026")
        tc1, tc2 = st.columns(2)
        with tc1:
            tw_chart = st.selectbox("Chart type", CHART_TYPES, index=0, key="tw_ct")
        with tc2:
            tw_h = st.slider("Height", 300, 800, chart_height, 50, key="tw_h")
        tw_c = NAGARRO_COLORS

        st.dataframe(drop_time_cols(tw), use_container_width=True, height=300, hide_index=True)

        t1, t2 = st.columns(2)
        nc = find_col(tw, "national")
        if nc:
            with t1:
                st.plotly_chart(make_chart(tw, nc, None, tw_chart, tw_c, tw_h, "By Nationality"), use_container_width=True)
        oc = find_col(tw, "occup", "profes")
        if oc:
            with t2:
                st.plotly_chart(make_chart(tw, oc, None, tw_chart, tw_c, tw_h, "By Profession"), use_container_width=True)
        emb = find_col(tw, "embassy")
        if emb:
            st.plotly_chart(make_chart(tw, emb, None, tw_chart, tw_c, tw_h, "By Embassy"), use_container_width=True)
        fee_cols = [c for c in tw.columns if "fee" in c.lower()]
        if fee_cols:
            tw_fees = tw[fee_cols].apply(pd.to_numeric, errors="coerce").sum().reset_index()
            tw_fees.columns = ["Fee Type", "Total"]
            tw_fees = tw_fees[tw_fees["Total"] > 0]
            if not tw_fees.empty:
                st.plotly_chart(make_chart(tw_fees, "Fee Type", "Total", tw_chart, tw_c, tw_h, "Fee Breakdown"), use_container_width=True)

        # Expense line chart
        st.markdown("---")
        st.markdown("#### Total Expenses")
        tw_total_cost = sum(r["Cost"] for r in tw_exp)
        st.metric("Total Temporary Work Cost", f"{tw_total_cost:,.0f} SAR")
        fig = make_expense_line(tw_exp, "Temporary Work Monthly Expenses", tw_h, "#FBD872")
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.warning("Temporary Work Visa 2026 sheet not found.")


# ===== TAB 3 : PERMANENT WORK ==============================================
with tabs[3]:
    if pw is not None:
        st.subheader("Permanent Work Visa 2026")
        pc1, pc2 = st.columns(2)
        with pc1:
            pw_chart = st.selectbox("Chart type", CHART_TYPES, index=0, key="pw_ct")
        with pc2:
            pw_h = st.slider("Height", 300, 800, chart_height, 50, key="pw_h")
        pw_c = NAGARRO_COLORS

        st.dataframe(drop_time_cols(pw), use_container_width=True, height=300, hide_index=True)

        p1, p2 = st.columns(2)
        nc = find_col(pw, "national")
        if nc:
            with p1:
                st.plotly_chart(make_chart(pw, nc, None, pw_chart, pw_c, pw_h, "By Nationality"), use_container_width=True)
        pc_col = find_col(pw, "project")
        if pc_col:
            with p2:
                st.plotly_chart(make_chart(pw, pc_col, None, pw_chart, pw_c, pw_h, "By Project"), use_container_width=True)
        prof = find_col(pw, "profes", "occup")
        if prof:
            st.plotly_chart(make_chart(pw, prof, None, pw_chart, pw_c, pw_h, "By Profession"), use_container_width=True)
        fee_names = ["MOI Fee", "COC Fee", "MOFA Fee"]
        found_fees = [c for c in pw.columns if any(f.lower() in c.lower() for f in fee_names)]
        if found_fees:
            pw_fees = pw[found_fees].apply(pd.to_numeric, errors="coerce").sum().reset_index()
            pw_fees.columns = ["Fee Type", "Total"]
            pw_fees = pw_fees[pw_fees["Total"] > 0]
            if not pw_fees.empty:
                st.plotly_chart(make_chart(pw_fees, "Fee Type", "Total", pw_chart, pw_c, pw_h, "Fee Breakdown"), use_container_width=True)
        city = find_col(pw, "city")
        if city:
            st.plotly_chart(make_chart(pw, city, None, pw_chart, pw_c, pw_h, "By City"), use_container_width=True)

        # 4 Expense line charts
        st.markdown("---")
        st.markdown("#### Permanent Work Expenses")
        pw_before_total = sum(r["Cost"] for r in pw_before)
        pw_after_total = sum(r["Cost"] for r in pw_after)
        pw_grand_total = sum(r["Cost"] for r in pw_total)
        em1, em2, em3 = st.columns(3)
        with em1:
            st.metric("Before Arrival", f"{pw_before_total:,.0f} SAR")
        with em2:
            st.metric("After Arrival", f"{pw_after_total:,.0f} SAR")
        with em3:
            st.metric("Grand Total", f"{pw_grand_total:,.0f} SAR")

        pe1, pe2 = st.columns(2)
        with pe1:
            fig = make_expense_line(pw_before, "Before Arrival to KSA (Monthly)", pw_h, "#47D7AC")
            st.plotly_chart(fig, use_container_width=True)
        with pe2:
            fig = make_expense_line(pw_after, "After Arrival in KSA (Monthly)", pw_h, "#FBD872")
            st.plotly_chart(fig, use_container_width=True)

        fig = make_expense_line(pw_total, "Total Permanent Work Expenses (Monthly)", pw_h, "#F8485E")
        st.plotly_chart(fig, use_container_width=True)

        # All 3 on one chart
        fig = make_multi_expense_line(
            {"Before Arrival": pw_before, "After Arrival": pw_after, "Total": pw_total},
            "Permanent Work — All Expense Categories", pw_h + 50)
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.warning("Permanent Work Visa 2026 sheet not found.")


# ===== TAB 4 : EXPENSES ====================================================
with tabs[4]:
    st.subheader("Expenses Overview")
    st.caption("Total spending across all visa types from Oct 2025 to Dec 2026")

    # Grand totals
    total_bv = sum(r["Cost"] for r in bv_exp)
    total_tw = sum(r["Cost"] for r in tw_exp)
    total_pw = sum(r["Cost"] for r in pw_total)
    grand_total = total_bv + total_tw + total_pw

    e1, e2, e3, e4 = st.columns(4)
    with e1:
        st.metric("Business Visit", f"{total_bv:,.0f} SAR")
    with e2:
        st.metric("Temporary Work", f"{total_tw:,.0f} SAR")
    with e3:
        st.metric("Permanent Work", f"{total_pw:,.0f} SAR")
    with e4:
        st.markdown(f'<div class="metric-card"><h3>{grand_total:,.0f}</h3><p>Total Spent (SAR)</p></div>', unsafe_allow_html=True)

    # Permanent Work breakdown: Before / After Arrival (totals)
    pw_before_sum = sum(r["Cost"] for r in pw_before)
    pw_after_sum = sum(r["Cost"] for r in pw_after)
    st.markdown("##### Permanent Work Breakdown")
    ba1, ba2 = st.columns(2)
    with ba1:
        st.metric("Before Arrival to KSA", f"{pw_before_sum:,.0f} SAR")
    with ba2:
        st.metric("After Arrival in KSA", f"{pw_after_sum:,.0f} SAR")

    # Month selector for Before / After Arrival
    st.markdown("##### Monthly Before & After Arrival Lookup")
    ba_month = st.selectbox("Select a month", TIMELINE_LABELS, index=len(TIMELINE_LABELS) - 1, key="exp_ba_month_sel")
    ba_idx = TIMELINE_LABELS.index(ba_month)
    ba_y, ba_m = TIMELINE[ba_idx]
    before_sel = sum(r["Cost"] for r in pw_before if r["Year"] == ba_y and r["Month_Num"] == ba_m)
    after_sel = sum(r["Cost"] for r in pw_after if r["Year"] == ba_y and r["Month_Num"] == ba_m)
    ba_total_sel = before_sel + after_sel

    ba3, ba4, ba5 = st.columns(3)
    with ba3:
        st.metric(f"Before Arrival ({ba_month})", f"{before_sel:,.0f} SAR")
    with ba4:
        st.metric(f"After Arrival ({ba_month})", f"{after_sel:,.0f} SAR")
    with ba5:
        st.metric(f"PW Total ({ba_month})", f"{ba_total_sel:,.0f} SAR")
    st.markdown("---")

    # Month selector — written totals (all visa types)
    st.markdown("#### Monthly Expense Lookup")
    sel_month = st.selectbox("Select a month", TIMELINE_LABELS, index=len(TIMELINE_LABELS) - 1, key="exp_month_sel")
    sel_idx = TIMELINE_LABELS.index(sel_month)
    sel_y, sel_m = TIMELINE[sel_idx]
    bv_sel = sum(r["Cost"] for r in bv_exp if r["Year"] == sel_y and r["Month_Num"] == sel_m)
    tw_sel = sum(r["Cost"] for r in tw_exp if r["Year"] == sel_y and r["Month_Num"] == sel_m)
    pw_sel = sum(r["Cost"] for r in pw_total if r["Year"] == sel_y and r["Month_Num"] == sel_m)
    total_sel = bv_sel + tw_sel + pw_sel

    ms1, ms2, ms3, ms4 = st.columns(4)
    with ms1:
        st.metric("Business Visit", f"{bv_sel:,.0f} SAR")
    with ms2:
        st.metric("Temporary Work", f"{tw_sel:,.0f} SAR")
    with ms3:
        st.metric("Permanent Work", f"{pw_sel:,.0f} SAR")
    with ms4:
        st.markdown(f'<div class="metric-card"><h3>{total_sel:,.0f}</h3><p>Total for {sel_month} (SAR)</p></div>', unsafe_allow_html=True)

    st.markdown("---")

    # Combined monthly line chart — all types + total
    st.markdown("#### Monthly Spending by Visa Type")
    fig = make_multi_expense_line(
        {"Business Visit": bv_exp, "Temporary Work": tw_exp, "Permanent Work": pw_total, "All Combined": all_exp},
        "Monthly Expenses — All Visa Types (OCT 2025 – DEC 2026)", 500)
    st.plotly_chart(fig, use_container_width=True)

    # Individual charts side by side
    st.markdown("---")
    st.markdown("#### Individual Visa Type Expenses")
    ie1, ie2, ie3 = st.columns(3)
    with ie1:
        fig = make_expense_line(bv_exp, "Business Visit", 400, "#47D7AC")
        st.plotly_chart(fig, use_container_width=True)
    with ie2:
        fig = make_expense_line(tw_exp, "Temporary Work", 400, "#FBD872")
        st.plotly_chart(fig, use_container_width=True)
    with ie3:
        fig = make_expense_line(pw_total, "Permanent Work", 400, "#F8485E")
        st.plotly_chart(fig, use_container_width=True)

    # Monthly total table
    st.markdown("---")
    st.markdown("#### Monthly Expense Table")
    table_rows = []
    for y, m in TIMELINE:
        label = f"{calendar.month_abbr[m]} {y}"
        bv_m = sum(r["Cost"] for r in bv_exp if r["Year"] == y and r["Month_Num"] == m)
        tw_m = sum(r["Cost"] for r in tw_exp if r["Year"] == y and r["Month_Num"] == m)
        pw_m = sum(r["Cost"] for r in pw_total if r["Year"] == y and r["Month_Num"] == m)
        total_m = bv_m + tw_m + pw_m
        if total_m > 0:
            table_rows.append({"Month": label, "Business Visit": f"{bv_m:,.0f}", "Temporary Work": f"{tw_m:,.0f}", "Permanent Work": f"{pw_m:,.0f}", "Total": f"{total_m:,.0f}"})
    if table_rows:
        st.dataframe(pd.DataFrame(table_rows), use_container_width=True, hide_index=True)
    else:
        st.info("No expense data available.")

    # Pie chart of total distribution
    st.markdown("---")
    st.markdown("#### Expense Distribution")
    if grand_total > 0:
        dist_df = pd.DataFrame({"Type": ["Business Visit", "Temporary Work", "Permanent Work"], "Cost": [total_bv, total_tw, total_pw]})
        dist_df = dist_df[dist_df["Cost"] > 0]
        fig = px.pie(dist_df, names="Type", values="Cost", color_discrete_sequence=NAGARRO_COLORS, title="Total Expense Distribution", hole=0.4)
        fig.update_traces(textposition="inside", textinfo="percent+value+label")
        fig.update_layout(height=450, plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)")
        st.plotly_chart(fig, use_container_width=True)


# ===== TAB 5 : AI CHAT =====================================================
with tabs[5]:
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
