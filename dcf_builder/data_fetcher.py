"""Data fetching with caching for market data, financials, and treasury rates."""

import json
import time
from pathlib import Path
from typing import Any, Optional

import pandas as pd
import yfinance as yf
from fredapi import Fred

from . import config


class Cache:
    """Simple JSON file cache with TTL support."""

    def __init__(self, cache_dir: Path = config.CACHE_DIR):
        self.cache_dir = cache_dir
        self.cache_dir.mkdir(parents=True, exist_ok=True)
        self.cache_file = self.cache_dir / "cache.json"
        self._cache = self._load_cache()

    def _load_cache(self) -> dict:
        if self.cache_file.exists():
            try:
                with open(self.cache_file, "r") as f:
                    return json.load(f)
            except (json.JSONDecodeError, IOError):
                return {}
        return {}

    def _save_cache(self) -> None:
        with open(self.cache_file, "w") as f:
            json.dump(self._cache, f, indent=2, default=str)

    def get(self, key: str, ttl: int) -> Optional[Any]:
        """Get cached value if not expired."""
        if key in self._cache:
            entry = self._cache[key]
            if time.time() - entry["timestamp"] < ttl:
                return entry["value"]
        return None

    def set(self, key: str, value: Any) -> None:
        """Cache a value with current timestamp."""
        self._cache[key] = {"value": value, "timestamp": time.time()}
        self._save_cache()

    def clear(self) -> None:
        """Clear all cached data."""
        self._cache = {}
        self._save_cache()


# Global cache instance
_cache = Cache()


class DataFetcher:
    """Fetches market data from yfinance and FRED with caching."""

    def __init__(self):
        self.cache = _cache
        self._fred = None

    @property
    def fred(self) -> Fred:
        """Lazy-load FRED client."""
        if self._fred is None:
            self._fred = Fred(api_key=config.FRED_API_KEY)
        return self._fred

    def get_stock_info(self, ticker: str) -> dict:
        """Get basic stock info (price, market cap, beta, etc.)."""
        cache_key = f"stock_info_{ticker}"
        cached = self.cache.get(cache_key, config.CACHE_TTL_MARKET_DATA)
        if cached:
            return cached

        try:
            stock = yf.Ticker(ticker)
            info = stock.info
            result = {
                "price": info.get("currentPrice") or info.get("regularMarketPrice"),
                "market_cap": info.get("marketCap"),
                "beta": info.get("beta"),
                "shares_outstanding": info.get("sharesOutstanding"),
                "fifty_two_week_high": info.get("fiftyTwoWeekHigh"),
                "fifty_two_week_low": info.get("fiftyTwoWeekLow"),
                "enterprise_value": info.get("enterpriseValue"),
                "trailing_pe": info.get("trailingPE"),
                "forward_pe": info.get("forwardPE"),
                "dividend_yield": info.get("dividendYield"),
                "name": info.get("longName") or info.get("shortName"),
                "sector": info.get("sector"),
                "industry": info.get("industry"),
            }
            self.cache.set(cache_key, result)
            return result
        except Exception as e:
            # Return cached value even if expired, or empty dict
            expired = self.cache._cache.get(cache_key, {}).get("value", {})
            if expired:
                return expired
            raise RuntimeError(f"Failed to fetch data for {ticker}: {e}")

    def get_price(self, ticker: str) -> Optional[float]:
        """Get current stock price."""
        return self.get_stock_info(ticker).get("price")

    def get_market_cap(self, ticker: str) -> Optional[float]:
        """Get market cap in millions."""
        mc = self.get_stock_info(ticker).get("market_cap")
        return mc / 1_000_000 if mc else None

    def get_beta(self, ticker: str) -> Optional[float]:
        """Get 5-year monthly beta."""
        return self.get_stock_info(ticker).get("beta")

    def get_shares_outstanding(self, ticker: str) -> Optional[float]:
        """Get shares outstanding."""
        return self.get_stock_info(ticker).get("shares_outstanding")

    def get_52_week_high(self, ticker: str) -> Optional[float]:
        """Get 52-week high price."""
        return self.get_stock_info(ticker).get("fifty_two_week_high")

    def get_52_week_low(self, ticker: str) -> Optional[float]:
        """Get 52-week low price."""
        return self.get_stock_info(ticker).get("fifty_two_week_low")

    def get_risk_free_rate(self) -> Optional[float]:
        """Get current 10-year Treasury rate from FRED."""
        cache_key = "risk_free_rate"
        cached = self.cache.get(cache_key, config.CACHE_TTL_TREASURY)
        if cached is not None:
            return cached

        try:
            # DGS10 is the 10-Year Treasury Constant Maturity Rate
            data = self.fred.get_series("DGS10")
            # Get most recent non-null value
            rate = data.dropna().iloc[-1] / 100  # Convert from percent
            self.cache.set(cache_key, rate)
            return rate
        except Exception as e:
            # Return cached value even if expired, or default
            expired = self.cache._cache.get(cache_key, {}).get("value")
            if expired is not None:
                return expired
            # Fallback to reasonable default
            return 0.04  # 4%

    def get_historical_financials(self, ticker: str) -> dict:
        """Get 5 years of historical financials from yfinance."""
        cache_key = f"financials_{ticker}"
        cached = self.cache.get(cache_key, config.CACHE_TTL_HISTORICAL)
        if cached:
            return cached

        try:
            stock = yf.Ticker(ticker)

            # Get income statement
            income_stmt = stock.financials
            # Get balance sheet
            balance_sheet = stock.balance_sheet

            result = {"income_statement": {}, "balance_sheet": {}, "years": []}

            if income_stmt is not None and not income_stmt.empty:
                # yfinance returns columns as dates
                years = [col.year for col in income_stmt.columns[:5]]
                result["years"] = years

                for col in income_stmt.columns[:5]:
                    year = col.year
                    result["income_statement"][year] = {
                        "revenue": self._safe_get(income_stmt, "Total Revenue", col),
                        "gross_profit": self._safe_get(income_stmt, "Gross Profit", col),
                        "ebitda": self._safe_get(income_stmt, "EBITDA", col),
                        "ebit": self._safe_get(income_stmt, "EBIT", col),
                        "net_income": self._safe_get(income_stmt, "Net Income", col),
                    }

            if balance_sheet is not None and not balance_sheet.empty:
                for col in balance_sheet.columns[:5]:
                    year = col.year
                    if year not in result["income_statement"]:
                        result["income_statement"][year] = {}
                    result["balance_sheet"][year] = {
                        "total_assets": self._safe_get(balance_sheet, "Total Assets", col),
                        "total_liabilities": self._safe_get(
                            balance_sheet, "Total Liabilities Net Minority Interest", col
                        ),
                        "total_equity": self._safe_get(
                            balance_sheet, "Total Equity Gross Minority Interest", col
                        ),
                        "cash": self._safe_get(
                            balance_sheet, "Cash And Cash Equivalents", col
                        ),
                        "total_debt": self._safe_get(balance_sheet, "Total Debt", col),
                    }

            self.cache.set(cache_key, result)
            return result
        except Exception as e:
            expired = self.cache._cache.get(cache_key, {}).get("value", {})
            if expired:
                return expired
            raise RuntimeError(f"Failed to fetch financials for {ticker}: {e}")

    def _safe_get(self, df: pd.DataFrame, row: str, col) -> Optional[float]:
        """Safely get a value from a DataFrame."""
        try:
            if row in df.index:
                val = df.loc[row, col]
                if pd.notna(val):
                    return float(val)
        except (KeyError, TypeError):
            pass
        return None

    def get_revenue(self, ticker: str, year: int) -> Optional[float]:
        """Get revenue for a specific year."""
        financials = self.get_historical_financials(ticker)
        return financials.get("income_statement", {}).get(year, {}).get("revenue")

    def get_ebitda(self, ticker: str, year: int) -> Optional[float]:
        """Get EBITDA for a specific year."""
        financials = self.get_historical_financials(ticker)
        return financials.get("income_statement", {}).get(year, {}).get("ebitda")

    def calculate_wacc(
        self,
        ticker: str,
        cost_of_debt: float = 0.05,
        tax_rate: float = config.DEFAULT_TAX_RATE,
    ) -> Optional[float]:
        """Calculate WACC for a company."""
        info = self.get_stock_info(ticker)
        financials = self.get_historical_financials(ticker)

        beta = info.get("beta")
        market_cap = info.get("market_cap")
        risk_free = self.get_risk_free_rate()

        if beta is None or market_cap is None or risk_free is None:
            return None

        # Get total debt from most recent year
        years = financials.get("years", [])
        total_debt = 0
        if years:
            latest_year = max(years)
            total_debt = (
                financials.get("balance_sheet", {}).get(latest_year, {}).get("total_debt")
                or 0
            )

        # Cost of equity using CAPM
        cost_of_equity = risk_free + beta * config.DEFAULT_EQUITY_RISK_PREMIUM

        # Capital structure weights
        total_capital = market_cap + total_debt
        if total_capital == 0:
            return cost_of_equity

        weight_equity = market_cap / total_capital
        weight_debt = total_debt / total_capital

        # WACC formula
        wacc = (weight_equity * cost_of_equity) + (
            weight_debt * cost_of_debt * (1 - tax_rate)
        )

        return wacc


# Convenience functions for direct access
_fetcher = DataFetcher()


def get_price(ticker: str) -> Optional[float]:
    return _fetcher.get_price(ticker)


def get_market_cap(ticker: str) -> Optional[float]:
    return _fetcher.get_market_cap(ticker)


def get_beta(ticker: str) -> Optional[float]:
    return _fetcher.get_beta(ticker)


def get_shares_outstanding(ticker: str) -> Optional[float]:
    return _fetcher.get_shares_outstanding(ticker)


def get_52_week_high(ticker: str) -> Optional[float]:
    return _fetcher.get_52_week_high(ticker)


def get_52_week_low(ticker: str) -> Optional[float]:
    return _fetcher.get_52_week_low(ticker)


def get_risk_free_rate() -> Optional[float]:
    return _fetcher.get_risk_free_rate()


def get_revenue(ticker: str, year: int) -> Optional[float]:
    return _fetcher.get_revenue(ticker, year)


def get_ebitda(ticker: str, year: int) -> Optional[float]:
    return _fetcher.get_ebitda(ticker, year)


def get_historical_financials(ticker: str) -> dict:
    return _fetcher.get_historical_financials(ticker)


def get_stock_info(ticker: str) -> dict:
    return _fetcher.get_stock_info(ticker)


def calculate_wacc(
    ticker: str,
    cost_of_debt: float = 0.05,
    tax_rate: float = config.DEFAULT_TAX_RATE,
) -> Optional[float]:
    return _fetcher.calculate_wacc(ticker, cost_of_debt, tax_rate)


def clear_cache() -> None:
    _fetcher.cache.clear()
