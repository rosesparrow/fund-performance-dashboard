"""
Fund Performance Attribution & Risk Dashboard
================================================
Interactive Streamlit dashboard for fund analysis.
Combines the key visuals from Parts 1-4 into a single web app.

TO RUN:
  pip install streamlit yfinance plotly pandas numpy
  streamlit run streamlit_dashboard.py

TO DEPLOY (free, public URL):
  1. Push this file + requirements.txt to a GitHub repo
  2. Go to share.streamlit.io
  3. Connect your repo → deploy
  4. Share the URL on your CV / portfolio
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
from datetime import datetime

# ══════════════════════════════════════════════════════════════════════════
# PAGE CONFIG — must be first Streamlit command
# ══════════════════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="Fund Performance Dashboard",
    page_icon="🔹",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ══════════════════════════════════════════════════════════════════════════
# STYLING — Fund Alpha colour palette
# ══════════════════════════════════════════════════════════════════════════
DARK_BLUE = "#1B2A4A"
MED_BLUE = "#2E5090"
LIGHT_BLUE = "#D6E4F0"
GREEN = "#27AE60"
RED = "#E74C3C"
ORANGE = "#E67E22"
PURPLE = "#8E44AD"
TEAL = "#1ABC9C"

FUND_COLOURS = {
    "Fund Alpha": MED_BLUE,
    "Peer 1 (Systematic)": GREEN,
    "Peer 2 (Replication)": ORANGE,
    "Peer 3 (Index-Based)": RED,
    "Global Equity Index": PURPLE,
}

# Custom CSS for dark navy theme
st.markdown("""
<style>
    /* Dark navy header bar */
    header[data-testid="stHeader"] {
        background-color: #1B2A4A;
    }
    
    /* Sidebar styling */
    section[data-testid="stSidebar"] {
        background-color: #1B2A4A;
    }
    section[data-testid="stSidebar"] > div:first-child {
        color: #FFFFFF;
    }
    section[data-testid="stSidebar"] h2,
    section[data-testid="stSidebar"] .stMarkdown p,
    section[data-testid="stSidebar"] .stMarkdown strong,
    section[data-testid="stSidebar"] span.st-emotion-cache-10trblm,
    section[data-testid="stSidebar"] label {
        color: #FFFFFF !important;
    }
    section[data-testid="stSidebar"] .stCaption p {
        color: #8899AA !important;
    }
    section[data-testid="stSidebar"] hr {
        border-color: #2E5090;
    }
    /* FIX: Keep date input text dark so it's readable */
    section[data-testid="stSidebar"] input {
        color: #1B2A4A !important;
    }
    
    /* Main content */
    .main .block-container { padding-top: 2rem; max-width: 1200px; }
    h1 { color: #1B2A4A !important; }
    h2 { color: #2E5090 !important; }
    h3 { color: #2E5090 !important; }
    
    /* Metric cards with navy accent */
    div[data-testid="stMetric"] {
        background-color: #f0f4f8;
        padding: 12px;
        border-radius: 8px;
        border-left: 4px solid #1B2A4A;
    }
    div[data-testid="stMetric"] label {
        color: #2E5090 !important;
    }
    
    /* Section banners */
    .section-banner {
        background: linear-gradient(135deg, #1B2A4A 0%, #2E5090 100%);
        color: white;
        padding: 14px 24px;
        border-radius: 8px;
        margin: 30px 0 16px 0;
        font-size: 1.3em;
        font-weight: 600;
        letter-spacing: 0.5px;
    }
    .section-banner .section-num {
        background-color: rgba(255,255,255,0.2);
        padding: 2px 10px;
        border-radius: 4px;
        margin-right: 10px;
        font-size: 0.85em;
    }
    
    /* Footer styling */
    .footer-text { text-align: center; color: #888; font-size: 0.85em; }
</style>
""", unsafe_allow_html=True)


def section_header(number, title):
    """Render a styled section banner."""
    st.markdown(
        f'<div class="section-banner">'
        f'<span class="section-num">{number}</span> {title}'
        f'</div>',
        unsafe_allow_html=True,
    )


# ══════════════════════════════════════════════════════════════════════════
# DATA LOADING — cached so it only downloads once
# ══════════════════════════════════════════════════════════════════════════
@st.cache_data(ttl=3600)  # refresh every hour
def load_data(start_date, end_date):
    """Download fund data from Yahoo Finance."""
    import yfinance as yf

    tickers = {
        "Fund Alpha": "AHLPX",
        "Peer 1 (Systematic)": "AQMIX",
        "Peer 2 (Replication)": "DBMF",
        "Peer 3 (Index-Based)": "KMLM",
        "Global Equity Index": "URTH",
    }

    try:
        prices = yf.download(
            list(tickers.values()),
            start=start_date,
            end=end_date,
            interval="1mo",
            progress=False,
        )["Close"]

        ticker_to_name = {v: k for k, v in tickers.items()}
        prices.columns = [ticker_to_name[col] for col in prices.columns]
        returns_df = prices.pct_change().dropna()

        # Reorder: funds first, benchmark last
        cols = [c for c in returns_df.columns if c != "Global Equity Index"] + ["Global Equity Index"]
        returns_df = returns_df[cols]
        return returns_df, None

    except Exception as e:
        return None, str(e)


def calc_metrics(r, bench, rf):
    """Calculate all metrics for a single fund."""
    n = len(r)
    ann_ret = (1 + r).prod() ** (12 / n) - 1
    ann_vol = r.std() * np.sqrt(12)
    dv = r[r < 0].std() * np.sqrt(12)
    cum = (1 + r).cumprod()
    peak = cum.cummax()
    dd = (cum - peak) / peak
    max_dd = dd.min()

    cov_fb = np.cov(r, bench)[0, 1]
    var_b = np.var(bench, ddof=1)
    beta = cov_fb / var_b if var_b != 0 else 0
    b_ann = (1 + bench).prod() ** (12 / len(bench)) - 1
    alpha = ann_ret - (rf + beta * (b_ann - rf))
    excess = r.values - bench.values
    te = np.std(excess, ddof=1) * np.sqrt(12)
    ir = (np.mean(excess) * 12) / te if te != 0 else 0

    return {
        "Annualised Return": ann_ret,
        "Annualised Volatility": ann_vol,
        "Sharpe Ratio": (ann_ret - rf) / ann_vol if ann_vol != 0 else 0,
        "Sortino Ratio": (ann_ret - rf) / dv if dv != 0 else 0,
        "Max Drawdown": max_dd,
        "Calmar Ratio": ann_ret / abs(max_dd) if max_dd != 0 else 0,
        "VaR (95%, 1M)": np.percentile(r, 5),
        "CVaR (95%, 1M)": r[r <= np.percentile(r, 5)].mean(),
        "Win Rate": (r > 0).mean(),
        "Beta": beta,
        "Alpha": alpha,
        "Tracking Error": te,
        "Information Ratio": ir,
    }


# ══════════════════════════════════════════════════════════════════════════
# SIDEBAR — controls
# ══════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown(
        "<div style='text-align: center; padding: 10px 0 5px 0;'>"
        "<span style='font-size: 2em;'>◆</span><br>"
        "<span style='font-size: 1.3em; font-weight: bold; letter-spacing: 2px;'>SB</span><br>"
        "<span style='font-size: 0.75em; letter-spacing: 1px; opacity: 0.7;'>PERFORMANCE ANALYTICS</span>"
        "</div>",
        unsafe_allow_html=True,
    )
    st.markdown("---")
    st.header("⚙️ Settings")

    start_date = st.date_input("Start Date", value=datetime(2019, 1, 1))
    end_date = st.date_input("End Date", value=datetime(2025, 12, 31))
    rf_rate = st.slider("Risk-Free Rate (%)", 0.0, 8.0, 4.5, 0.5) / 100

    st.markdown("---")
    st.markdown("**Built by:** Sewa B")
    st.markdown("**Data:** Yahoo Finance")
    st.markdown("**Framework:** Python + Streamlit")
    st.markdown("---")
    st.caption("Fund Performance Attribution & Risk Dashboard")


# ══════════════════════════════════════════════════════════════════════════
# LOAD DATA
# ══════════════════════════════════════════════════════════════════════════
returns_df, error = load_data(str(start_date), str(end_date))

if error or returns_df is None:
    st.error(f"❌ Failed to load data: {error}")
    st.info("Check your internet connection and try again.")
    st.stop()

bench = returns_df["Global Equity Index"]
fund_names = [c for c in returns_df.columns if c != "Global Equity Index"]
all_funds = list(returns_df.columns)

# Calculate metrics for all funds
all_metrics = {}
for fund in all_funds:
    all_metrics[fund] = calc_metrics(returns_df[fund], bench, rf_rate)
metrics_df = pd.DataFrame(all_metrics)

# Primary fund specific
ahl = returns_df["Fund Alpha"]
ahl_metrics = all_metrics["Fund Alpha"]


# ══════════════════════════════════════════════════════════════════════════
# HEADER
# ══════════════════════════════════════════════════════════════════════════
st.title("Fund Performance Attribution & Risk Dashboard")
st.markdown(f"**Fund Alpha vs Managed Futures Peers** | {returns_df.index[0].strftime('%b %Y')} to {returns_df.index[-1].strftime('%b %Y')} | {len(returns_df)} months")
st.markdown("---")


# ══════════════════════════════════════════════════════════════════════════
# SECTION 1: KEY METRICS (Fund Alpha headline numbers)
# ══════════════════════════════════════════════════════════════════════════
section_header("01", "Fund Alpha — Key Metrics")

col1, col2, col3, col4, col5, col6 = st.columns(6)
col1.metric("Ann. Return", f"{ahl_metrics['Annualised Return']:.1%}")
col2.metric("Ann. Volatility", f"{ahl_metrics['Annualised Volatility']:.1%}")
col3.metric("Sharpe Ratio", f"{ahl_metrics['Sharpe Ratio']:.2f}")
col4.metric("Max Drawdown", f"{ahl_metrics['Max Drawdown']:.1%}")
col5.metric("Beta", f"{ahl_metrics['Beta']:.2f}")
col6.metric("Alpha", f"{ahl_metrics['Alpha']:.1%}")

st.markdown("")


# ══════════════════════════════════════════════════════════════════════════
# SECTION 2: CUMULATIVE GROWTH
# ══════════════════════════════════════════════════════════════════════════
section_header("02", "Cumulative Growth of $100")
st.caption("How $100 invested at inception would have grown for each fund.")

growth = (1 + returns_df).cumprod() * 100

fig_growth = go.Figure()
for fund in all_funds:
    fig_growth.add_trace(go.Scatter(
        x=growth.index, y=growth[fund],
        name=fund, mode="lines",
        line=dict(color=FUND_COLOURS.get(fund, "grey"), width=2.5),
    ))
fig_growth.update_layout(
    height=450, template="plotly_white",
    yaxis_title="Value ($)", xaxis_title="",
    legend=dict(orientation="h", yanchor="bottom", y=-0.25, xanchor="center", x=0.5),
    hovermode="x unified",
)
fig_growth.add_hline(y=100, line_dash="dot", line_color="grey", opacity=0.5)
st.plotly_chart(fig_growth, use_container_width=True)


# ══════════════════════════════════════════════════════════════════════════
# SECTION 3: RISK / RETURN SCATTER
# ══════════════════════════════════════════════════════════════════════════
section_header("03", "Risk / Return — Peer Comparison")
st.caption("Top-left = best (high return, low risk). Bottom-right = worst.")

scatter_data = pd.DataFrame({
    "Fund": all_funds,
    "Volatility": [all_metrics[f]["Annualised Volatility"] for f in all_funds],
    "Return": [all_metrics[f]["Annualised Return"] for f in all_funds],
    "Sharpe": [all_metrics[f]["Sharpe Ratio"] for f in all_funds],
})

fig_scatter = go.Figure()
for _, row in scatter_data.iterrows():
    fig_scatter.add_trace(go.Scatter(
        x=[row["Volatility"]], y=[row["Return"]],
        name=row["Fund"], mode="markers+text",
        marker=dict(size=14, color=FUND_COLOURS.get(row["Fund"], "grey")),
        text=[row["Fund"]], textposition="top center",
        textfont=dict(size=11),
        hovertemplate=f"<b>{row['Fund']}</b><br>Vol: {row['Volatility']:.1%}<br>Return: {row['Return']:.1%}<br>Sharpe: {row['Sharpe']:.2f}<extra></extra>",
    ))
fig_scatter.update_layout(
    height=500, template="plotly_white",
    xaxis_title="Annualised Volatility (Risk)",
    yaxis_title="Annualised Return",
    xaxis_tickformat=".0%", yaxis_tickformat=".0%",
    showlegend=False,
)
st.plotly_chart(fig_scatter, use_container_width=True)


# ══════════════════════════════════════════════════════════════════════════
# SECTION 4: PEER COMPARISON TABLE
# ══════════════════════════════════════════════════════════════════════════
section_header("04", "Peer Comparison — All Metrics")

display_metrics = metrics_df.T.copy()
# Format for display
fmt_pct = ["Annualised Return", "Annualised Volatility", "Max Drawdown",
           "VaR (95%, 1M)", "CVaR (95%, 1M)", "Win Rate", "Alpha", "Tracking Error"]
fmt_ratio = ["Sharpe Ratio", "Sortino Ratio", "Calmar Ratio", "Beta", "Information Ratio"]

styled = display_metrics.style.format(
    {c: "{:.1%}" for c in fmt_pct if c in display_metrics.columns}
).format(
    {c: "{:.2f}" for c in fmt_ratio if c in display_metrics.columns}
).background_gradient(
    subset=["Sharpe Ratio"], cmap="RdYlGn", vmin=-0.5, vmax=1.0
).background_gradient(
    subset=["Max Drawdown"], cmap="RdYlGn", vmin=-0.4, vmax=0
).background_gradient(
    subset=["Annualised Return"], cmap="RdYlGn", vmin=-0.05, vmax=0.15
)

st.dataframe(styled, use_container_width=True, height=250)


# ══════════════════════════════════════════════════════════════════════════
# SECTION 5: CALENDAR YEAR RETURNS
# ══════════════════════════════════════════════════════════════════════════
section_header("05", "Calendar Year Returns")

yearly = returns_df.groupby(returns_df.index.year).apply(lambda x: (1 + x).prod() - 1)

fig_cal = go.Figure()
for fund in all_funds:
    fig_cal.add_trace(go.Bar(
        x=yearly.index.astype(str), y=yearly[fund],
        name=fund,
        marker_color=FUND_COLOURS.get(fund, "grey"),
    ))
fig_cal.update_layout(
    barmode="group", height=400, template="plotly_white",
    yaxis_title="Annual Return", yaxis_tickformat=".0%",
    legend=dict(orientation="h", yanchor="bottom", y=-0.3, xanchor="center", x=0.5),
)
fig_cal.add_hline(y=0, line_color="grey", line_width=1)
st.plotly_chart(fig_cal, use_container_width=True)

# Year table
st.dataframe(
    yearly.style.format("{:.1%}").applymap(
        lambda v: f"color: {GREEN}" if v > 0 else f"color: {RED}"),
    use_container_width=True, height=200,
)


# ══════════════════════════════════════════════════════════════════════════
# SECTION 6: DRAWDOWN COMPARISON
# ══════════════════════════════════════════════════════════════════════════
section_header("06", "Drawdown Comparison")
st.caption("How deep and how long each fund's losses were from their peak.")

dd_df = pd.DataFrame()
for col in returns_df.columns:
    cum = (1 + returns_df[col]).cumprod()
    peak = cum.cummax()
    dd_df[col] = (cum - peak) / peak

fig_dd = go.Figure()
for fund in all_funds:
    fig_dd.add_trace(go.Scatter(
        x=dd_df.index, y=dd_df[fund],
        name=fund, mode="lines",
        line=dict(color=FUND_COLOURS.get(fund, "grey"), width=2),
        fill="tozeroy" if fund == "Fund Alpha" else None,
        fillcolor="rgba(46,80,144,0.1)" if fund == "Fund Alpha" else None,
    ))
fig_dd.update_layout(
    height=400, template="plotly_white",
    yaxis_title="Drawdown", yaxis_tickformat=".0%",
    legend=dict(orientation="h", yanchor="bottom", y=-0.25, xanchor="center", x=0.5),
    hovermode="x unified",
)
st.plotly_chart(fig_dd, use_container_width=True)


# ══════════════════════════════════════════════════════════════════════════
# SECTION 7: ROLLING SHARPE
# ══════════════════════════════════════════════════════════════════════════
section_header("07", "Rolling 12-Month Sharpe Ratio")
st.caption("How consistently each fund delivers risk-adjusted returns over time.")

def rolling_sharpe(returns_df, window=12, rf=rf_rate):
    def sharpe_w(x):
        ann_ret = (1 + x).prod() ** (12 / len(x)) - 1
        vol = x.std() * np.sqrt(12)
        return (ann_ret - rf) / vol if vol != 0 else 0
    return returns_df.rolling(window).apply(sharpe_w, raw=False).dropna()

rolling_sh = rolling_sharpe(returns_df)

fig_rs = go.Figure()
for fund in all_funds:
    fig_rs.add_trace(go.Scatter(
        x=rolling_sh.index, y=rolling_sh[fund],
        name=fund, mode="lines",
        line=dict(color=FUND_COLOURS.get(fund, "grey"), width=2),
    ))
fig_rs.update_layout(
    height=400, template="plotly_white",
    yaxis_title="Sharpe Ratio",
    legend=dict(orientation="h", yanchor="bottom", y=-0.25, xanchor="center", x=0.5),
    hovermode="x unified",
)
fig_rs.add_hline(y=0, line_dash="dot", line_color="grey", opacity=0.5)
st.plotly_chart(fig_rs, use_container_width=True)


# ══════════════════════════════════════════════════════════════════════════
# SECTION 8: RISK LIMIT MONITOR
# ══════════════════════════════════════════════════════════════════════════
section_header("08", "Risk Limit Monitor — Fund Alpha")

limits = {
    "Max Drawdown": {"current": ahl_metrics["Max Drawdown"], "limit": -0.20, "negative": True},
    "Monthly VaR (95%)": {"current": ahl_metrics["VaR (95%, 1M)"], "limit": -0.08, "negative": True},
    "Ann. Volatility": {"current": ahl_metrics["Annualised Volatility"], "limit": 0.25, "negative": False},
    "Beta vs Equity": {"current": ahl_metrics["Beta"], "limit": 0.50, "negative": False},
}

limit_cols = st.columns(len(limits))
for i, (name, data) in enumerate(limits.items()):
    with limit_cols[i]:
        current = data["current"]
        limit = data["limit"]
        if data["negative"]:
            util = abs(current / limit) if limit != 0 else 0
        else:
            util = current / limit if limit != 0 else 0

        if util < 0.80:
            status, colour = "🟢 GREEN", GREEN
        elif util < 1.0:
            status, colour = "🟡 AMBER", ORANGE
        else:
            status, colour = "🔴 RED", RED

        st.markdown(f"**{name}**")
        if name in ["Max Drawdown", "Monthly VaR (95%)"]:
            st.markdown(f"Current: **{current:.1%}** | Limit: {limit:.0%}")
        elif name == "Ann. Volatility":
            st.markdown(f"Current: **{current:.1%}** | Limit: {limit:.0%}")
        else:
            st.markdown(f"Current: **{current:.2f}** | Limit: {limit:.2f}")
        st.markdown(f"Utilisation: **{util:.0%}** → {status}")


# ══════════════════════════════════════════════════════════════════════════
# FOOTER
# ══════════════════════════════════════════════════════════════════════════
st.markdown("---")
st.markdown(
    f"<div style='text-align: center; color: #888; font-size: 0.85em;'>"
    f"Fund Performance Attribution & Risk Dashboard | Data: Yahoo Finance | "
    f"Period: {returns_df.index[0].strftime('%b %Y')} – {returns_df.index[-1].strftime('%b %Y')} | "
    f"Built with Python, Streamlit & Plotly"
    f"</div>",
    unsafe_allow_html=True,
)
