# Fund Performance Attribution & Risk Dashboard

**Fund performance analytics · risk monitoring · peer group benchmarking**

Python · Streamlit · Excel · Yahoo Finance · Plotly

---

## Project overview

This project builds a fund analytics platform for a systematic trend-following hedge fund, benchmarked against managed futures peers and a global equity index. Using publicly available data from Yahoo Finance, the platform calculates 15 risk-adjusted performance metrics, conducts attribution analysis, produces institutional-style risk reports with limit monitoring, and presents everything through an interactive Streamlit web dashboard.

The goal was to replicate the analytical workflow of a portfolio analytics or risk team — from raw return data through to the outputs that go in front of a portfolio manager or investor.

> **Note on fund names:** The analysis uses publicly available mutual fund and ETF data as proxies for institutional managed futures strategies. Tickers and display names are configured in a single dictionary at the top of each file, making it straightforward to swap in different funds or asset classes.

## What it covers

| Part | Deliverable | What it demonstrates |
|------|------------|---------------------|
| **Part 1** | Performance Metrics Dashboard | 15 risk-adjusted metrics with formula-driven Excel workbook |
| **Part 2** | Attribution Analysis | Brinson-Fachler decomposition + factor attribution by asset class |
| **Part 3** | Peer Group Comparison | Ranked comparison across 5 funds on all metrics |
| **Part 4** | Monthly Risk Report | Risk report with limit monitoring, stress testing, and exposure breakdown |
| **Part 5** | Interactive Dashboard | Streamlit web app with live data, interactive charts, and risk limits |

## Key metrics calculated

The platform computes 15 performance and risk metrics across all funds:

- **Return and volatility** — annualised return, annualised volatility
- **Risk-adjusted ratios** — Sharpe, Sortino, and Calmar
- **Drawdown analysis** — maximum drawdown, drawdown duration
- **Tail risk** — Value at Risk (VaR) and Conditional VaR at the 95% confidence level
- **Market relationship** — CAPM beta, CAPM alpha, equity correlation
- **Relative performance** — tracking error, information ratio, win rate

The Excel workbooks use cell-level formulas throughout, so the methodology is fully transparent.

## Technical approach

The platform combines Python and Excel. Python handles data acquisition from Yahoo Finance, statistical calculations, and automated report generation. Excel provides formula-driven workbooks that recalculate automatically when the underlying data changes. Streamlit ties it together as a web application, allowing anyone to interact with the dashboard without installing anything.

All calculations use standard industry formulas applied to monthly returns. The Excel workbooks and the Python scripts are built independently from the same source data, which means agreement between them serves as a cross-check on the methodology.

| Component | Technology |
|-----------|-----------|
| Language | Python (pandas, numpy, openpyxl, plotly) |
| Spreadsheets | Excel (.xlsx) — formula-driven workbooks with conditional formatting and charts |
| Data source | Yahoo Finance API (yfinance) — real-time fund data |
| Dashboard | Streamlit + Plotly (interactive web application) |
| Notebook | Jupyter Notebook for reproducible analytical workflow |

## Repository structure

```
├── fund_performance_notebook.ipynb    # Master notebook — orchestrates Parts 1–4
├── build_dashboard_v2.py              # Part 1: Performance metrics → Excel
├── build_attribution.py               # Part 2: Brinson attribution → Excel
├── build_peer_comparison.py           # Part 3: Peer group analysis → Excel
├── build_risk_report.py               # Part 4: Monthly risk report → Excel
├── streamlit_dashboard.py             # Part 5: Interactive web dashboard (standalone)
├── requirements.txt                   # Python dependencies
├── .gitignore                         # Excludes generated Excel files
└── README.md
```

**How to run:**

1. Clone the repo and install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

2. Run the Jupyter notebook to generate all Excel workbooks:
   ```bash
   jupyter notebook fund_performance_notebook.ipynb
   ```
   Execute cells in order — the notebook pulls live data from Yahoo Finance and runs each build script.

3. Launch the interactive dashboard:
   ```bash
   streamlit run streamlit_dashboard.py
   ```

## Attribution analysis

The project implements two attribution models:

**Part A — Classic Brinson-Fachler Attribution** decomposes portfolio outperformance into allocation effect (sector weight decisions), selection effect (stock picking within sectors), and interaction effect. This uses a simulated equity portfolio vs a broad equity benchmark to demonstrate the standard model that underpins most institutional performance reporting.

**Part B — Factor Attribution for Managed Futures** decomposes returns into contributions from asset class factors: equities, bonds, commodities, and FX. This is more appropriate for a multi-asset trend-following fund that trades across asset classes rather than selecting securities within a single market.

## Risk report

The monthly risk report (Part 4) is modelled on what institutional risk teams produce. It includes an executive summary with NAV and period performance, a risk metrics snapshot with month-on-month changes, exposure breakdowns by asset class (long, short, and net), a drawdown monitor tracking current and historical drawdowns, scenario-based stress testing (2008 GFC, 2020 COVID, rate shock, etc.), and a traffic-light limit monitoring system that flags breaches against pre-set risk thresholds.

## Peer group analysis

The analysis compares the primary fund against managed futures peers over a 60-month period. The peer group represents different approaches to systematic trend-following, from pure momentum to more diversified multi-strategy. A global equity index is included not as a peer but as a reference point to illustrate managed futures diversification properties.

An important caveat: all funds in the analysis are retail-accessible proxies (mutual fund wrappers and listed ETFs) rather than direct institutional strategies. The wrapper and ETF structures impose their own constraints — lower leverage targets, daily liquidity requirements, and higher relative fee loads — which means the figures reflect the wrapper as much as the underlying strategy.

## Skills demonstrated

- **Excel** — formula-driven workbooks, conditional formatting, named ranges, multi-sheet architecture, chart formatting
- **Python** — pandas, numpy, openpyxl for programmatic Excel generation, yfinance for data acquisition, plotly for interactive visualisation
- **Risk analytics** — VaR/CVaR, stress testing, limit monitoring, drawdown analysis, Brinson attribution
- **Performance measurement** — Sharpe/Sortino/Calmar ratios, tracking error, information ratio, CAPM alpha and beta
- **Data visualisation** — Streamlit dashboards, Plotly charts, institutional-quality report formatting
- **Reproducibility** — notebook-driven workflow, formula transparency, automated data pipeline

## Live dashboard

The Streamlit dashboard is deployed at: [fund-performance-dashboard.streamlit.app](https://rs-fund-performance-dashboard.streamlit.app)

---

*Built as part of a portfolio analytics project demonstrating investment risk and performance reporting capabilities.*
