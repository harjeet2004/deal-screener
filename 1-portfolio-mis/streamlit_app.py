"""
streamlit_app.py  —  MGA Ventures Portfolio MIS Dashboard

Two sections:
  1. Portfolio Overview  — status cards, RAG table, revenue & runway charts
  2. Company Detail      — dropdown, metric cards, analyst note, 4 trend charts

Driven entirely by file upload. No sample_data/ fallback.
All monetary values display in Indian Rupee format (₹1,23,45,678).
Chart axes are styled for a dark background: white bold text, visible gridlines.

Run: streamlit run streamlit_app.py
"""

import os
import sys
from io import BytesIO

import numpy as np
import pandas as pd
import streamlit as st
import plotly.graph_objects as go

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from mis_report import load_company, compute_metrics, company_name_from_path, fmt_inr

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="MGA Portfolio MIS",
    page_icon="📊",
    layout="wide",
)

CRORE = 10_000_000

# ── Colour palette (tuned for dark background) ────────────────────────────────
C = {
    "navy":    "#1F3864",
    "blue":    "#4A9FE0",
    "orange":  "#F0954A",
    "grey":    "#AAAAAA",
    "green":   "#5FC669",
    "amber":   "#FFC107",
    "red":     "#FF5252",
    "trend":   "#7EB3F5",
    "white":   "#FFFFFF",
    "subtext": "#C8C8D8",
}
STATUS_COLOR = {
    "Critical  (<2 qtrs)": C["red"],
    "Watch  (2-4 qtrs)":   C["amber"],
    "Healthy  (>4 qtrs)":  C["green"],
}
STATUS_SHORT = {
    "Critical  (<2 qtrs)": "At Risk",
    "Watch  (2-4 qtrs)":   "Watch",
    "Healthy  (>4 qtrs)":  "Healthy",
}


# ── Dark-theme chart styling ──────────────────────────────────────────────────

def _dark_axis(title=""):
    """Return axis config readable on a dark background."""
    return dict(
        title=title,
        gridcolor="#3A3A4C",
        linecolor="#66667A",
        linewidth=1.5,
        tickfont=dict(color="#FFFFFF", size=12, family="Arial, sans-serif"),
        title_font=dict(color="#FFFFFF", size=13, family="Arial, sans-serif"),
        showgrid=True,
        zeroline=False,
    )


def _apply_dark(fig, title="", height=350, xangle=-45):
    """Apply consistent dark-theme layout to any Plotly figure."""
    fig.update_layout(
        title=dict(
            text=title,
            font=dict(size=14, color="#FFFFFF", family="Arial, sans-serif"),
        ),
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
        font=dict(color="#FFFFFF", size=12, family="Arial, sans-serif"),
        legend=dict(
            font=dict(color="#FFFFFF", size=11),
            bgcolor="rgba(20,20,36,0.75)",
            bordercolor="#55556A",
            borderwidth=1,
        ),
        height=height,
        margin=dict(l=65, r=20, t=48, b=90),
    )
    fig.update_xaxes(tickangle=xangle, **_dark_axis())
    fig.update_yaxes(**_dark_axis())
    return fig


# ── File processing (cached per content hash) ─────────────────────────────────

@st.cache_data(show_spinner=False)
def _process_file(file_bytes: bytes, filename: str):
    name = company_name_from_path(filename)
    df   = load_company(BytesIO(file_bytes), fname=filename)
    if df is None:
        return name, None, None
    return name, df, compute_metrics(df, name)


# ── Analyst note generator ────────────────────────────────────────────────────

def generate_analyst_note(metrics):
    """3-sentence note driven entirely by actual numbers, not hardcoded strings."""
    name   = metrics["Company"]
    cagr   = metrics.get("_cagr_raw") or 0
    runway = metrics.get("_runway_raw", 0)
    gm     = metrics.get("_gm_raw") or 0
    status = metrics.get("Status", "")

    if cagr >= 0.30:
        s1 = (f"{name} has compounded revenue at {cagr*100:.1f}% annually over three years, "
              "placing it firmly in high-growth territory.")
    elif cagr >= 0.10:
        s1 = (f"{name} has grown revenue at a moderate {cagr*100:.1f}% CAGR — "
              "steady progress, but below high-growth benchmarks for its stage.")
    elif cagr > 0:
        s1 = (f"{name}'s {cagr*100:.1f}% revenue CAGR signals a plateauing trajectory; "
              "the business has not yet found a second growth gear.")
    else:
        s1 = (f"{name} is showing revenue contraction ({cagr*100:.1f}% CAGR), "
              "a fundamental concern at this lifecycle stage.")

    if "Critical" in status:
        s2 = (f"Cash is critical — only {runway:.1f} quarters of runway remain at current burn; "
              "immediate fundraising or deep cost action is required.")
    elif "Watch" in status:
        s2 = (f"With {runway:.1f} quarters of runway, fundraising conversations should begin "
              "now to avoid closing in a distressed position.")
    elif runway == float("inf") or runway > 20:
        s2 = ("The balance sheet is strong — the company is cash-flow positive "
              "and not dependent on external capital.")
    else:
        s2 = (f"The balance sheet is healthy at {runway:.1f} quarters of runway, "
              "providing a comfortable buffer for continued growth investment.")

    if gm >= 0.65:
        s3 = (f"At {gm*100:.0f}% gross margin, unit economics are strong — "
              "scale will meaningfully expand operating leverage.")
    elif gm >= 0.40:
        s3 = (f"Gross margins at {gm*100:.0f}% are acceptable but leave limited room "
              "to absorb opex growth as the business scales.")
    else:
        s3 = (f"Gross margins of {gm*100:.0f}% are structurally thin and must improve "
              "before sustainable profitability is achievable.")

    return f"{s1} {s2} {s3}"


# ── Chart builders ────────────────────────────────────────────────────────────

def _cr(values):
    return [v / CRORE for v in values]


def chart_revenue_trend(df):
    qtrs    = df["quarter_label"].tolist()
    revs    = df["revenue"].tolist()
    revs_cr = _cr(revs)
    x       = np.arange(len(revs_cr))
    slope, intercept = np.polyfit(x, revs_cr, 1)
    trend   = (slope * x + intercept).tolist()

    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=qtrs, y=revs_cr, mode="lines+markers", name="Revenue",
        line=dict(color=C["blue"], width=2.5),
        marker=dict(size=7, color=C["blue"]),
        customdata=[fmt_inr(r) for r in revs],
        hovertemplate="<b>%{x}</b><br>Revenue: %{customdata}<extra></extra>",
    ))
    fig.add_trace(go.Scatter(
        x=qtrs, y=trend, mode="lines", name="Trend",
        line=dict(color=C["trend"], width=1.5, dash="dash"),
        hoverinfo="skip",
    ))
    _apply_dark(fig, "Quarterly Revenue")
    fig.update_yaxes(title_text="₹ Crores", **_dark_axis("₹ Crores"))
    return fig


def chart_pnl_breakdown(df):
    qtrs = df["quarter_label"].tolist()
    series = [
        (df["revenue"].tolist(), "Revenue", C["blue"]),
        (df["cogs"].tolist(),    "COGS",    C["orange"]),
        (df["opex"].tolist(),    "OpEx",    C["grey"]),
    ]
    fig = go.Figure()
    for raw_vals, label, color in series:
        fig.add_trace(go.Scatter(
            x=qtrs, y=_cr(raw_vals), name=label,
            mode="lines+markers",
            line=dict(color=color, width=2),
            marker=dict(size=5),
            customdata=[fmt_inr(v) for v in raw_vals],
            hovertemplate=f"<b>%{{x}}</b><br>{label}: %{{customdata}}<extra></extra>",
        ))
    _apply_dark(fig, "Revenue vs COGS vs OpEx")
    fig.update_yaxes(title_text="₹ Crores", **_dark_axis("₹ Crores"))
    return fig


def chart_cash_balance(df, metrics):
    qtrs      = df["quarter_label"].tolist()
    cash_vals = df["cash"].tolist()
    cash_cr   = _cr(cash_vals)
    burn      = metrics.get("_burn_raw", 0)

    bar_colors = []
    for c in cash_vals:
        if burn <= 0:
            bar_colors.append(C["green"])
        elif c / burn > 4:
            bar_colors.append(C["green"])
        elif c / burn >= 2:
            bar_colors.append(C["amber"])
        else:
            bar_colors.append(C["red"])

    fig = go.Figure(go.Bar(
        x=qtrs, y=cash_cr,
        marker_color=bar_colors,
        customdata=[fmt_inr(v) for v in cash_vals],
        hovertemplate="<b>%{x}</b><br>Cash: %{customdata}<extra></extra>",
    ))
    _apply_dark(fig, "Cash Balance (coloured by runway status)")
    fig.update_yaxes(title_text="₹ Crores", **_dark_axis("₹ Crores"))
    return fig


def chart_qoq_growth(df):
    qtrs = df["quarter_label"].tolist()
    revs = df["revenue"].tolist()
    qoq  = [(revs[i] - revs[i-1]) / revs[i-1] * 100 for i in range(1, len(revs))]
    colors = [C["green"] if v >= 0 else C["red"] for v in qoq]

    fig = go.Figure(go.Bar(
        x=qtrs[1:], y=qoq,
        marker_color=colors,
        text=[f"{v:.1f}%" for v in qoq],
        textposition="outside",
        textfont=dict(color="white", size=11),
        hovertemplate="<b>%{x}</b><br>QoQ: %{y:.1f}%<extra></extra>",
    ))
    fig.add_hline(y=0, line_color="#AAAAAA", line_width=1.5)
    _apply_dark(fig, "Quarter-on-Quarter Revenue Growth")
    fig.update_yaxes(title_text="Growth %", ticksuffix="%", **_dark_axis("Growth %"))
    return fig


def chart_portfolio_revenue(all_data):
    companies = list(all_data.keys())
    rev_vals  = [all_data[c]["metrics"]["_latest_rev_raw"] for c in companies]
    rev_cr    = [v / CRORE for v in rev_vals]
    colors    = [STATUS_COLOR.get(all_data[c]["metrics"]["Status"], C["blue"]) for c in companies]

    fig = go.Figure(go.Bar(
        x=rev_cr, y=companies, orientation="h",
        marker_color=colors,
        customdata=[fmt_inr(v) for v in rev_vals],
        hovertemplate="<b>%{y}</b><br>Revenue: %{customdata}<extra></extra>",
        text=[f"₹{v:.1f} Cr" for v in rev_cr],
        textposition="outside",
        textfont=dict(color="white", size=11),
    ))
    h = max(280, len(companies) * 62)
    _apply_dark(fig, "Latest Quarterly Revenue", height=h, xangle=0)
    fig.update_xaxes(title_text="₹ Crores", **_dark_axis("₹ Crores"))
    fig.update_yaxes(autorange="reversed", **_dark_axis())
    fig.update_layout(margin=dict(l=130, r=95, t=48, b=50))
    return fig


def chart_portfolio_runway(all_data):
    companies = list(all_data.keys())
    runways, colors, texts = [], [], []
    for c in companies:
        r = all_data[c]["metrics"]["_runway_raw"]
        runways.append(min(r, 24) if r != float("inf") else 24)
        colors.append(STATUS_COLOR.get(all_data[c]["metrics"]["Status"], C["blue"]))
        texts.append("∞" if r == float("inf") else f"{r:.1f} qtrs")

    fig = go.Figure(go.Bar(
        x=runways, y=companies, orientation="h",
        marker_color=colors,
        text=texts,
        textposition="outside",
        textfont=dict(color="white", size=11),
        hovertemplate="<b>%{y}</b><br>Runway: %{text}<extra></extra>",
    ))
    for x_val, color, label in [(2, C["red"], "2 qtrs"), (4, C["amber"], "4 qtrs")]:
        fig.add_vline(
            x=x_val, line_dash="dash", line_color=color, line_width=1.5,
            annotation_text=label,
            annotation_font=dict(color="white", size=11),
            annotation_position="top",
        )
    h = max(280, len(companies) * 62)
    _apply_dark(fig, "Runway by Company (quarters)", height=h, xangle=0)
    fig.update_xaxes(title_text="Quarters", **_dark_axis("Quarters"))
    fig.update_yaxes(autorange="reversed", **_dark_axis())
    fig.update_layout(margin=dict(l=130, r=85, t=48, b=50))
    return fig


# ── Main layout ───────────────────────────────────────────────────────────────

def main():
    # ── Header ───────────────────────────────────────────────────────────────
    st.markdown(
        "<h1 style='color:#FFFFFF;margin-bottom:2px'>MGA Ventures — Portfolio MIS</h1>"
        f"<p style='color:{C['subtext']};font-size:13px;margin-top:0'>"
        "Quarterly data · FY23 Q1 – FY25 Q4 · All figures in Indian Rupees</p>",
        unsafe_allow_html=True,
    )
    st.divider()

    # ── Session state — tracks whether Generate has been clicked ─────────────
    if "dashboard_fp" not in st.session_state:
        st.session_state.dashboard_fp   = None   # fingerprint of last generated set
        st.session_state.show_dashboard = False

    # ── File upload (the only input — no sample_data fallback) ───────────────
    uploaded_files = st.file_uploader(
        "Upload company financial files (.xlsx)",
        type=["xlsx"],
        accept_multiple_files=True,
        label_visibility="visible",
    )

    if not uploaded_files:
        # Blank state: reset any previous dashboard and show instruction only
        st.session_state.show_dashboard = False
        st.session_state.dashboard_fp   = None
        st.markdown(
            f"<p style='color:{C['subtext']};font-size:14px;margin-top:8px'>"
            "Upload one or more company .xlsx files above to populate the dashboard.</p>",
            unsafe_allow_html=True,
        )
        return

    # Fingerprint the current upload set by name + size so we detect changes
    current_fp = tuple(sorted((uf.name, uf.size) for uf in uploaded_files))

    # If the file set changed since last generation, force the user to re-generate
    if current_fp != st.session_state.dashboard_fp:
        st.session_state.show_dashboard = False

    # ── Generate button row ───────────────────────────────────────────────────
    btn_col, info_col = st.columns([2, 5])
    with btn_col:
        if st.button("Generate Dashboard", type="primary", use_container_width=True):
            st.session_state.dashboard_fp   = current_fp
            st.session_state.show_dashboard = True

    with info_col:
        names_str = "  •  ".join(uf.name for uf in uploaded_files)
        st.markdown(
            f"<p style='color:{C['subtext']};font-size:13px;padding-top:8px'>"
            f"{len(uploaded_files)} file(s) selected:  {names_str}</p>",
            unsafe_allow_html=True,
        )

    if not st.session_state.show_dashboard:
        st.markdown(
            f"<p style='color:{C['subtext']};font-size:14px;margin-top:4px'>"
            "Files ready — click <b>Generate Dashboard</b> to compute metrics "
            "and build the report.</p>",
            unsafe_allow_html=True,
        )
        return   # nothing rendered below until Generate is clicked

    # ── Process uploaded files (cached per content) ───────────────────────────
    all_data, failed = {}, []
    with st.spinner("Processing files..."):
        for uf in uploaded_files:
            name, df, metrics = _process_file(uf.getvalue(), uf.name)
            if df is not None:
                all_data[name] = {"df": df, "metrics": metrics}
            else:
                failed.append(uf.name)

    if not all_data:
        st.error(
            "Could not parse any of the uploaded files. "
            "Check that they contain financial data with recognisable column headers."
        )
        return

    n_loaded     = len(all_data)
    company_list = ", ".join(all_data.keys())
    st.success(
        f"{n_loaded} {'file' if n_loaded == 1 else 'files'} loaded — "
        f"companies detected: **{company_list}**"
    )
    if failed:
        st.warning(f"Could not parse: {', '.join(failed)}")

    # ══════════════════════════════════════════════════════════════════════════
    # SECTION 1 — Portfolio Overview
    # ══════════════════════════════════════════════════════════════════════════
    st.markdown(
        "<h2 style='color:#FFFFFF;margin-top:24px;margin-bottom:16px'>Portfolio Overview</h2>",
        unsafe_allow_html=True,
    )

    statuses   = [d["metrics"]["Status"] for d in all_data.values()]
    n_healthy  = sum(1 for s in statuses if "Healthy"  in s)
    n_watch    = sum(1 for s in statuses if "Watch"    in s)
    n_critical = sum(1 for s in statuses if "Critical" in s)

    pill_col1, pill_col2, pill_col3, _ = st.columns([1, 1, 1, 3])
    for col, count, label, color in [
        (pill_col1, n_healthy,  "Healthy  (>4 qtrs)",  C["green"]),
        (pill_col2, n_watch,    "Watch  (2–4 qtrs)",   C["amber"]),
        (pill_col3, n_critical, "Critical  (<2 qtrs)", C["red"]),
    ]:
        with col:
            st.markdown(
                f"<div style='background:{color};padding:16px 12px;border-radius:8px;text-align:center'>"
                f"<div style='font-size:30px;font-weight:bold;color:white'>{count}</div>"
                f"<div style='color:white;font-size:12px;margin-top:3px'>{label}</div>"
                "</div>",
                unsafe_allow_html=True,
            )

    st.markdown("<br>", unsafe_allow_html=True)

    # RAG summary table
    TABLE_COLS = [
        "Company", "Latest Quarter", "Revenue", "QoQ Rev Growth",
        "YoY Rev Growth", "Gross Margin", "Revenue CAGR (3Y)",
        "Qtrly Burn", "Cash", "Runway (qtrs)", "Status",
    ]
    table_rows = [d["metrics"] for d in all_data.values()]
    table_df   = pd.DataFrame([{c: r.get(c, "") for c in TABLE_COLS} for r in table_rows])

    def _style_status(val):
        return {
            "Critical  (<2 qtrs)": "background-color:#FF5252;color:white;font-weight:bold",
            "Watch  (2-4 qtrs)":   "background-color:#FFC107;color:white;font-weight:bold",
            "Healthy  (>4 qtrs)":  "background-color:#5FC669;color:white;font-weight:bold",
        }.get(val, "")

    styled = (
        table_df.style
        .map(_style_status, subset=["Status"])
        .set_properties(**{"text-align": "center"})
        .set_table_styles([
            {"selector": "th", "props": [
                ("background-color", "#1F3864"), ("color", "white"),
                ("font-weight", "bold"), ("text-align", "center"), ("padding", "8px 10px"),
            ]},
            {"selector": "td", "props": [("padding", "6px 10px")]},
        ])
    )
    st.dataframe(styled, use_container_width=True, hide_index=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # Portfolio comparison charts
    ch_l, ch_r = st.columns(2)
    with ch_l:
        st.plotly_chart(chart_portfolio_revenue(all_data), use_container_width=True)
    with ch_r:
        st.plotly_chart(chart_portfolio_runway(all_data), use_container_width=True)

    # ══════════════════════════════════════════════════════════════════════════
    # SECTION 2 — Company Detail
    # ══════════════════════════════════════════════════════════════════════════
    st.divider()
    st.markdown(
        "<h2 style='color:#FFFFFF;margin-bottom:14px'>Company Detail</h2>",
        unsafe_allow_html=True,
    )

    selected = st.selectbox("Select a company", options=list(all_data.keys()))
    df_sel   = all_data[selected]["df"]
    m        = all_data[selected]["metrics"]
    status   = m["Status"]
    sc       = STATUS_COLOR.get(status, C["blue"])
    sl       = STATUS_SHORT.get(status, status)

    # Company name + coloured status pill (replaces the white callout box)
    st.markdown(
        f"<div style='margin:12px 0 8px 0'>"
        f"<span style='color:white;font-size:22px;font-weight:700;margin-right:10px'>"
        f"{selected}</span>"
        f"<span style='background:{sc};color:white;font-size:12px;font-weight:bold;"
        f"padding:4px 14px;border-radius:20px;vertical-align:middle;letter-spacing:0.03em'>"
        f"{sl}</span>"
        f"</div>",
        unsafe_allow_html=True,
    )

    # Metric cards
    mc1, mc2, mc3, mc4 = st.columns(4)
    with mc1:
        st.metric("Latest Revenue",    m["Revenue"])
    with mc2:
        gm_str = f"{m['_gm_raw']*100:.1f}%" if m.get("_gm_raw") is not None else "N/A"
        st.metric("Gross Margin",      gm_str)
    with mc3:
        rq = m["_runway_raw"]
        st.metric("Runway",            "∞" if rq == float("inf") else f"{rq:.1f} qtrs")
    with mc4:
        cagr_str = f"{m['_cagr_raw']*100:.1f}%" if m.get("_cagr_raw") is not None else "N/A"
        st.metric("Revenue CAGR (3Y)", cagr_str)

    # Analyst note — plain readable text directly on dark background, no container
    note = generate_analyst_note(m)
    st.markdown(
        f"<p style='color:{C['subtext']};font-size:14px;line-height:1.80;"
        f"margin:10px 0 22px 0;max-width:920px'>{note}</p>",
        unsafe_allow_html=True,
    )

    # 2×2 trend chart panel
    r1c1, r1c2 = st.columns(2)
    with r1c1:
        st.plotly_chart(chart_revenue_trend(df_sel),   use_container_width=True)
    with r1c2:
        st.plotly_chart(chart_pnl_breakdown(df_sel),   use_container_width=True)

    r2c1, r2c2 = st.columns(2)
    with r2c1:
        st.plotly_chart(chart_cash_balance(df_sel, m), use_container_width=True)
    with r2c2:
        st.plotly_chart(chart_qoq_growth(df_sel),      use_container_width=True)


if __name__ == "__main__":
    main()
