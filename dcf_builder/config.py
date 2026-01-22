"""Configuration and constants for DCF Builder."""

import os
from pathlib import Path

# API Keys (set FRED_API_KEY environment variable)
# Get a free key at: https://fred.stlouisfed.org/docs/api/api_key.html
FRED_API_KEY = os.environ.get("FRED_API_KEY")

# Cache settings
CACHE_DIR = Path.home() / ".dcf_builder_cache"
CACHE_TTL_MARKET_DATA = 15 * 60  # 15 minutes for live market data
CACHE_TTL_HISTORICAL = 24 * 60 * 60  # 24 hours for historical financials
CACHE_TTL_TREASURY = 60 * 60  # 1 hour for treasury rates

# Default DCF assumptions
DEFAULT_EQUITY_RISK_PREMIUM = 0.055  # 5.5%
DEFAULT_MARKET_RETURN = 0.10  # 10%
DEFAULT_TAX_RATE = 0.21  # 21% corporate tax rate
DEFAULT_PROJECTION_YEARS = 5
DEFAULT_TERMINAL_GROWTH = 0.025  # 2.5%

# Scenario multipliers (relative to base case)
SCENARIOS = {
    "Bull": {
        "revenue_growth_adj": 1.2,  # 20% higher growth
        "margin_adj": 1.1,  # 10% better margins
        "terminal_growth_adj": 1.1,
    },
    "Base": {
        "revenue_growth_adj": 1.0,
        "margin_adj": 1.0,
        "terminal_growth_adj": 1.0,
    },
    "Bear": {
        "revenue_growth_adj": 0.8,  # 20% lower growth
        "margin_adj": 0.9,  # 10% worse margins
        "terminal_growth_adj": 0.9,
    },
}

# Sanity check thresholds
SANITY_CHECKS = {
    "revenue_growth_max": 0.50,  # 50%
    "revenue_growth_min": -0.30,  # -30%
    "ebitda_margin_max": 0.60,  # 60%
    "ebitda_margin_min": 0.05,  # 5%
    "wacc_max": 0.20,  # 20%
    "wacc_min": 0.05,  # 5%
}

# Excel sheet names
SHEET_NAMES = [
    "Dashboard",
    "Assumptions",
    "Historical",
    "Projections",
    "Valuation",
    "Comps",
    "Sensitivity",
    "Football Field",
]

# Template path
TEMPLATE_DIR = Path(__file__).parent.parent / "templates"
BASE_TEMPLATE_PATH = TEMPLATE_DIR / "base_dcf.xlsx"
