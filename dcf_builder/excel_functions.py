"""Custom Excel functions (UDFs) for DCF Builder.

These functions can be used directly in Excel cells when xlwings is configured.
They wrap the data_fetcher module to provide clean Excel function syntax.

Usage in Excel (after xlwings setup):
    =DCF_PRICE("AAPL")
    =DCF_MARKET_CAP("AAPL")
    =DCF_BETA("AAPL")
    =DCF_RISK_FREE()
    =DCF_REVENUE("AAPL", 2023)
"""

from typing import Optional, Union

from . import data_fetcher as df


def DCF_PRICE(ticker: str) -> Optional[float]:
    """Get current stock price.

    Args:
        ticker: Stock ticker symbol (e.g., "AAPL")

    Returns:
        Current stock price or None if unavailable
    """
    try:
        return df.get_price(ticker.upper())
    except Exception:
        return None


def DCF_MARKET_CAP(ticker: str) -> Optional[float]:
    """Get market capitalization in millions.

    Args:
        ticker: Stock ticker symbol (e.g., "AAPL")

    Returns:
        Market cap in millions or None if unavailable
    """
    try:
        return df.get_market_cap(ticker.upper())
    except Exception:
        return None


def DCF_BETA(ticker: str) -> Optional[float]:
    """Get 5-year monthly beta.

    Args:
        ticker: Stock ticker symbol (e.g., "AAPL")

    Returns:
        Beta or None if unavailable
    """
    try:
        return df.get_beta(ticker.upper())
    except Exception:
        return None


def DCF_SHARES_OUT(ticker: str) -> Optional[float]:
    """Get shares outstanding.

    Args:
        ticker: Stock ticker symbol (e.g., "AAPL")

    Returns:
        Shares outstanding or None if unavailable
    """
    try:
        return df.get_shares_outstanding(ticker.upper())
    except Exception:
        return None


def DCF_52W_HIGH(ticker: str) -> Optional[float]:
    """Get 52-week high price.

    Args:
        ticker: Stock ticker symbol (e.g., "AAPL")

    Returns:
        52-week high price or None if unavailable
    """
    try:
        return df.get_52_week_high(ticker.upper())
    except Exception:
        return None


def DCF_52W_LOW(ticker: str) -> Optional[float]:
    """Get 52-week low price.

    Args:
        ticker: Stock ticker symbol (e.g., "AAPL")

    Returns:
        52-week low price or None if unavailable
    """
    try:
        return df.get_52_week_low(ticker.upper())
    except Exception:
        return None


def DCF_RISK_FREE() -> Optional[float]:
    """Get current 10-year Treasury rate (risk-free rate).

    Returns:
        Current 10-year Treasury rate as decimal (e.g., 0.04 for 4%)
    """
    try:
        return df.get_risk_free_rate()
    except Exception:
        return None


def DCF_REVENUE(ticker: str, year: Union[int, float]) -> Optional[float]:
    """Get revenue for a specific year.

    Args:
        ticker: Stock ticker symbol (e.g., "AAPL")
        year: Fiscal year (e.g., 2023)

    Returns:
        Revenue in dollars or None if unavailable
    """
    try:
        return df.get_revenue(ticker.upper(), int(year))
    except Exception:
        return None


def DCF_EBITDA(ticker: str, year: Union[int, float]) -> Optional[float]:
    """Get EBITDA for a specific year.

    Args:
        ticker: Stock ticker symbol (e.g., "AAPL")
        year: Fiscal year (e.g., 2023)

    Returns:
        EBITDA in dollars or None if unavailable
    """
    try:
        return df.get_ebitda(ticker.upper(), int(year))
    except Exception:
        return None


def DCF_WACC(ticker: str) -> Optional[float]:
    """Calculate WACC for a company.

    Args:
        ticker: Stock ticker symbol (e.g., "AAPL")

    Returns:
        WACC as decimal (e.g., 0.10 for 10%)
    """
    try:
        return df.calculate_wacc(ticker.upper())
    except Exception:
        return None


def DCF_EV(ticker: str) -> Optional[float]:
    """Get enterprise value.

    Args:
        ticker: Stock ticker symbol (e.g., "AAPL")

    Returns:
        Enterprise value or None if unavailable
    """
    try:
        info = df.get_stock_info(ticker.upper())
        return info.get("enterprise_value")
    except Exception:
        return None


def DCF_PE(ticker: str) -> Optional[float]:
    """Get trailing P/E ratio.

    Args:
        ticker: Stock ticker symbol (e.g., "AAPL")

    Returns:
        Trailing P/E ratio or None if unavailable
    """
    try:
        info = df.get_stock_info(ticker.upper())
        return info.get("trailing_pe")
    except Exception:
        return None


# Export all functions for xlwings registration
__all__ = [
    "DCF_PRICE",
    "DCF_MARKET_CAP",
    "DCF_BETA",
    "DCF_SHARES_OUT",
    "DCF_52W_HIGH",
    "DCF_52W_LOW",
    "DCF_RISK_FREE",
    "DCF_REVENUE",
    "DCF_EBITDA",
    "DCF_WACC",
    "DCF_EV",
    "DCF_PE",
]
