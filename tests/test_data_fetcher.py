"""Tests for data_fetcher module."""

import pytest
from unittest.mock import patch, MagicMock
import tempfile
from pathlib import Path

from dcf_builder import data_fetcher as df
from dcf_builder.data_fetcher import Cache


class TestCache:
    """Tests for the Cache class."""

    def test_cache_set_and_get(self):
        """Test basic cache set and get operations."""
        with tempfile.TemporaryDirectory() as tmpdir:
            cache = Cache(Path(tmpdir))
            cache.set("test_key", {"value": 123})

            # Should retrieve within TTL
            result = cache.get("test_key", ttl=3600)
            assert result == {"value": 123}

    def test_cache_expiry(self):
        """Test that expired cache entries return None."""
        with tempfile.TemporaryDirectory() as tmpdir:
            cache = Cache(Path(tmpdir))
            cache.set("test_key", "test_value")

            # Should return None for 0 TTL (immediately expired)
            result = cache.get("test_key", ttl=0)
            assert result is None

    def test_cache_clear(self):
        """Test cache clearing."""
        with tempfile.TemporaryDirectory() as tmpdir:
            cache = Cache(Path(tmpdir))
            cache.set("key1", "value1")
            cache.set("key2", "value2")

            cache.clear()

            assert cache.get("key1", ttl=3600) is None
            assert cache.get("key2", ttl=3600) is None


class TestDataFetcher:
    """Tests for DataFetcher class."""

    @patch("dcf_builder.data_fetcher.yf.Ticker")
    def test_get_stock_info(self, mock_ticker):
        """Test fetching stock info."""
        # Setup mock
        mock_ticker_instance = MagicMock()
        mock_ticker_instance.info = {
            "currentPrice": 150.0,
            "marketCap": 2500000000000,
            "beta": 1.2,
            "sharesOutstanding": 16000000000,
            "fiftyTwoWeekHigh": 180.0,
            "fiftyTwoWeekLow": 120.0,
            "longName": "Apple Inc.",
            "sector": "Technology",
        }
        mock_ticker.return_value = mock_ticker_instance

        # Clear cache first
        df.clear_cache()

        fetcher = df.DataFetcher()
        result = fetcher.get_stock_info("AAPL")

        assert result["price"] == 150.0
        assert result["market_cap"] == 2500000000000
        assert result["beta"] == 1.2
        assert result["name"] == "Apple Inc."

    @patch("dcf_builder.data_fetcher.yf.Ticker")
    def test_get_price(self, mock_ticker):
        """Test get_price convenience function."""
        mock_ticker_instance = MagicMock()
        mock_ticker_instance.info = {"currentPrice": 150.0}
        mock_ticker.return_value = mock_ticker_instance

        df.clear_cache()
        price = df.get_price("AAPL")
        assert price == 150.0

    @patch("dcf_builder.data_fetcher.yf.Ticker")
    def test_get_market_cap(self, mock_ticker):
        """Test get_market_cap returns value in millions."""
        mock_ticker_instance = MagicMock()
        mock_ticker_instance.info = {"marketCap": 2500000000000}
        mock_ticker.return_value = mock_ticker_instance

        df.clear_cache()
        market_cap = df.get_market_cap("AAPL")
        assert market_cap == 2500000.0  # In millions

    @patch("dcf_builder.data_fetcher.Fred")
    def test_get_risk_free_rate_fallback(self, mock_fred_class):
        """Test that risk-free rate has a fallback value."""
        df.clear_cache()

        # Force fresh fetcher with mocked Fred
        mock_fred_instance = MagicMock()
        mock_fred_instance.get_series.side_effect = Exception("API Error")
        mock_fred_class.return_value = mock_fred_instance

        fetcher = df.DataFetcher()
        fetcher._fred = None  # Force re-creation
        rate = fetcher.get_risk_free_rate()

        # Should return fallback value
        assert rate == 0.04


class TestWACCCalculation:
    """Tests for WACC calculation."""

    @patch("dcf_builder.data_fetcher.yf.Ticker")
    def test_calculate_wacc(self, mock_ticker):
        """Test WACC calculation."""
        mock_ticker_instance = MagicMock()
        mock_ticker_instance.info = {
            "currentPrice": 150.0,
            "marketCap": 2500000000000,
            "beta": 1.2,
        }
        mock_ticker_instance.financials = MagicMock()
        mock_ticker_instance.financials.empty = True
        mock_ticker_instance.balance_sheet = MagicMock()
        mock_ticker_instance.balance_sheet.empty = True
        mock_ticker.return_value = mock_ticker_instance

        df.clear_cache()

        with patch.object(df._fetcher, "get_risk_free_rate", return_value=0.04):
            wacc = df.calculate_wacc("AAPL")

            # WACC should be a reasonable value
            assert wacc is not None
            assert 0 < wacc < 0.30  # Sanity check
