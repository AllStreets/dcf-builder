# DCF Builder - Design Document

## Overview

DCF Builder is a Python-powered Excel add-in that helps investment banking analysts build DCF models faster with automated data fetching, template generation, and error checking.

**Target user**: IB/M&A analysts who build valuation models in Excel

**Problem solved**:
- Model building is repetitive and slow
- Manual data entry is error-prone
- Error checking is tedious and often missed
- Pulling market data requires switching between sources

## Architecture

```
┌─────────────────────────────────────────────────┐
│                    Excel                        │
│  ┌──────────────┐  ┌─────────────────────────┐  │
│  │ Custom Ribbon│  │ Custom Functions        │  │
│  │ - Generate   │  │ =DCF_BETA("AAPL")       │  │
│  │ - Error Check│  │ =DCF_MARKET_CAP()       │  │
│  │ - Refresh    │  │ =DCF_RISK_FREE()        │  │
│  └──────────────┘  └─────────────────────────┘  │
└────────────────────────┬────────────────────────┘
                         │ xlwings
┌────────────────────────▼────────────────────────┐
│                 Python Backend                   │
│  ┌────────────┐ ┌────────────┐ ┌─────────────┐  │
│  │ Template   │ │ Data       │ │ Error       │  │
│  │ Generator  │ │ Fetcher    │ │ Checker     │  │
└────────────────────────┬────────────────────────┘
                         │
┌────────────────────────▼────────────────────────┐
│              External APIs                       │
│  yfinance │ FRED (treasury rates) │ SEC EDGAR   │
└─────────────────────────────────────────────────┘
```

## Technology Stack

| Component | Technology |
|-----------|------------|
| Excel integration | xlwings |
| Stock data | yfinance |
| Treasury rates | fredapi (FRED API) |
| SEC filings | sec-edgar-downloader |
| Data manipulation | pandas |
| Excel file creation | openpyxl |
| Testing | pytest |

## Features

### 1. Template Generator

Creates an 8-sheet DCF workbook with one click.

**Sheet structure**:

| Sheet | Purpose |
|-------|---------|
| Dashboard | Executive summary, football field chart, key metrics |
| Assumptions | All inputs, scenario toggle (Bull/Base/Bear) |
| Historical | 5 years of actual financials, growth rates, margins |
| Projections | Revenue → EBITDA → Unlevered FCF |
| Valuation | DCF math, terminal value, enterprise → equity value |
| Comps | Peer trading multiples, percentile rankings |
| Sensitivity | WACC vs. growth matrix, margin vs. growth matrix |
| Football Field | Data underlying the valuation range chart |

**Ribbon action**: "Generate DCF" → Prompts for ticker → Full model with live data

### 2. Data Fetcher

Pulls live market and company data into Excel.

**Data sources**:

| Source | Data |
|--------|------|
| yfinance | Market cap, beta, price, 52-week range, shares outstanding |
| SEC EDGAR | Historical financials (revenue, EBITDA, assets, liabilities) |
| FRED | 10-year Treasury rate (risk-free rate) |
| Defaults | Equity risk premium (~5.5%), market return assumptions |

**Custom Excel functions**:

```
=DCF_PRICE("AAPL")         → Current stock price
=DCF_MARKET_CAP("AAPL")    → Market cap in millions
=DCF_BETA("AAPL")          → 5-year monthly beta
=DCF_SHARES_OUT("AAPL")    → Shares outstanding
=DCF_52W_HIGH("AAPL")      → 52-week high
=DCF_52W_LOW("AAPL")       → 52-week low
=DCF_RISK_FREE()           → Current 10-year Treasury
=DCF_REVENUE("AAPL", 2023) → Revenue for given year
=DCF_EBITDA("AAPL", 2023)  → EBITDA for given year
```

**Ribbon actions**:
- "Refresh All Data" → Re-pulls all market data
- "Pull Historicals" → Fetches 5 years of financials

**Offline fallback**: Functions return last cached value if APIs fail.

### 3. Error Checker

Scans models for issues across four categories.

**Formula Errors**:
- Circular references (flags cell, suggests fix)
- #REF, #VALUE, #DIV/0 errors
- Broken external links
- Inconsistent formulas in ranges

**Sanity Checks**:
- Revenue growth > 50% or < -30%
- EBITDA margin outside 5-60%
- WACC outside 5-20%
- Terminal growth > risk-free rate
- Negative terminal year FCF

**Structural Issues**:
- Hardcoded numbers in formula rows
- Inconsistent time periods
- Unlinked assumptions

**Balance Sheet Checks**:
- Assets ≠ Liabilities + Equity
- Cash flow doesn't reconcile
- Retained earnings doesn't roll forward

**Output**: Results pane with errors sorted by severity (Critical / Warning / Info). Each error is clickable and jumps to the problem cell.

### 4. Scenario Manager

Toggle between Bull / Base / Bear cases.

**Behavior**:
- Dropdown on Assumptions sheet selects scenario
- Each scenario stores different growth rates, margins, terminal assumptions
- Changing dropdown updates: Projections → Valuation → Dashboard
- Football field shows all three scenarios

### 5. Comparable Company Analysis

Auto-populates peer multiples.

**Input**: User enters 5-10 peer tickers on Assumptions sheet

**Output table**:

| Company | Ticker | EV | Revenue | EBITDA | EV/Rev | EV/EBITDA | P/E |
|---------|--------|-----|---------|--------|--------|-----------|-----|
| Peer 1 | XXX | ... | ... | ... | ... | ... | ... |
| Peer 2 | YYY | ... | ... | ... | ... | ... | ... |
| **Median** | | | | | **X.Xx** | **X.Xx** | **X.Xx** |
| **Target** | ZZZ | ... | ... | ... | X.Xx | X.Xx | X.Xx |
| **Percentile** | | | | | Xth | Xth | Xth |

**Additional analysis**:
- Implied valuation at median multiples
- Scatter plot: EV/EBITDA vs. Revenue Growth
- Target row highlighting

### 6. Historical Analysis

Pulls and analyzes 5 years of financials.

**Data pulled**:
- Income statement: Revenue, COGS, Gross Profit, EBITDA, EBIT, Net Income
- Balance sheet: Assets, Liabilities, Equity, Cash, AR, Inventory, AP, Debt

**Calculated metrics**:

| Metric | Calculation |
|--------|-------------|
| Revenue Growth | YoY % change |
| Gross Margin | Gross Profit / Revenue |
| EBITDA Margin | EBITDA / Revenue |
| ROIC | NOPAT / Invested Capital |
| Debt/EBITDA | Total Debt / EBITDA |

**Integration**: Base case assumptions default to 5-year averages.

### 7. Executive Dashboard

One-page visual summary.

**Layout**:

```
┌─────────────────────────────────────────────────────────────────┐
│  COMPANY NAME (TICKER)                     Generated: Date      │
│  DCF Valuation Analysis                    Scenario: Base Case  │
├─────────────────────────────────────────────────────────────────┤
│  KEY METRICS                      VALUATION SUMMARY             │
│  ┌─────────────────────────┐     ┌─────────────────────────┐   │
│  │ Current Price           │     │ DCF Value               │   │
│  │ Market Cap              │     │ Comps Median            │   │
│  │ EV/EBITDA               │     │ 52-Week Range           │   │
│  │ Revenue Growth          │     │ Implied Upside          │   │
│  │ EBITDA Margin           │     │                         │   │
│  └─────────────────────────┘     └─────────────────────────┘   │
│                                                                 │
│  FOOTBALL FIELD CHART                                           │
│  ┌─────────────────────────────────────────────────────────┐   │
│  │ DCF - Bear    ████████                                  │   │
│  │ DCF - Base         ████████████                         │   │
│  │ DCF - Bull              ████████                        │   │
│  │ Comps              ████████                             │   │
│  │ 52-Week        ████████████████                         │   │
│  └─────────────────────────────────────────────────────────┘   │
│                                                                 │
│  SCENARIO COMPARISON TABLE                                      │
│  Bear / Base / Bull values for key metrics                      │
└─────────────────────────────────────────────────────────────────┘
```

**Dynamic elements**:
- All values linked to model
- Football field is an Excel chart (auto-updates)
- Conditional formatting: green if undervalued, red if overvalued

## Project Structure

```
dcf-builder/
├── README.md
├── requirements.txt
├── setup.py
├── dcf_builder/
│   ├── __init__.py
│   ├── main.py               # xlwings entry point, ribbon callbacks
│   ├── template_generator.py # Creates DCF workbook structure
│   ├── data_fetcher.py       # API calls (yfinance, FRED, SEC)
│   ├── error_checker.py      # All validation logic
│   ├── excel_functions.py    # Custom UDFs
│   ├── formatters.py         # Cell formatting, chart creation
│   └── config.py             # API keys, default assumptions
├── templates/
│   └── base_dcf.xlsx         # Pre-formatted template shell
├── tests/
│   ├── test_data_fetcher.py
│   ├── test_error_checker.py
│   └── test_template.py
└── dcf_builder.xlam          # Compiled Excel add-in
```

## Dependencies

```
xlwings>=0.30.0
yfinance>=0.2.0
fredapi>=0.5.0
pandas>=2.0.0
openpyxl>=3.1.0
requests>=2.31.0
pytest>=7.0.0
```

## Installation (End User)

1. Install Python 3.9+
2. `pip install dcf-builder`
3. `xlwings addin install`
4. Open Excel → DCF Builder tab appears in ribbon

## Interview Talking Points

This project demonstrates:

- **Python**: Backend logic, API integration, data manipulation
- **Excel**: Deep understanding of formulas, UDFs, workbook structure
- **APIs**: Working with yfinance, SEC EDGAR, FRED
- **Financial modeling**: DCF, comps, WACC, terminal value concepts
- **UX thinking**: Ribbon design, error messaging, dashboard layout
- **Software engineering**: Modular code, testing, packaging

Key phrases:
- "Built a Python Excel add-in that automates DCF modeling"
- "Pulls live data from Yahoo Finance, SEC EDGAR, and FRED"
- "Automated error checking catches circular refs and sanity issues"
- "Features scenario analysis and comparable company valuation"
