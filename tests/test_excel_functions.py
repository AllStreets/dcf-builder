"""Tests for excel_functions module."""

import pytest
from unittest.mock import patch

from dcf_builder import excel_functions as ef


class TestExcelFunctions:
    """Tests for custom Excel functions."""

    @patch("dcf_builder.excel_functions.df")
    def test_dcf_price(self, mock_df):
        """Test DCF_PRICE function."""
        mock_df.get_price.return_value = 150.0

        result = ef.DCF_PRICE("aapl")  # Test lowercase conversion
        mock_df.get_price.assert_called_with("AAPL")
        assert result == 150.0

    @patch("dcf_builder.excel_functions.df")
    def test_dcf_price_handles_error(self, mock_df):
        """Test DCF_PRICE returns None on error."""
        mock_df.get_price.side_effect = Exception("API Error")

        result = ef.DCF_PRICE("INVALID")
        assert result is None

    @patch("dcf_builder.excel_functions.df")
    def test_dcf_market_cap(self, mock_df):
        """Test DCF_MARKET_CAP function."""
        mock_df.get_market_cap.return_value = 2500000.0  # In millions

        result = ef.DCF_MARKET_CAP("AAPL")
        assert result == 2500000.0

    @patch("dcf_builder.excel_functions.df")
    def test_dcf_beta(self, mock_df):
        """Test DCF_BETA function."""
        mock_df.get_beta.return_value = 1.2

        result = ef.DCF_BETA("AAPL")
        assert result == 1.2

    @patch("dcf_builder.excel_functions.df")
    def test_dcf_risk_free(self, mock_df):
        """Test DCF_RISK_FREE function."""
        mock_df.get_risk_free_rate.return_value = 0.04

        result = ef.DCF_RISK_FREE()
        assert result == 0.04

    @patch("dcf_builder.excel_functions.df")
    def test_dcf_revenue(self, mock_df):
        """Test DCF_REVENUE function."""
        mock_df.get_revenue.return_value = 400000000000

        result = ef.DCF_REVENUE("AAPL", 2023)
        mock_df.get_revenue.assert_called_with("AAPL", 2023)
        assert result == 400000000000

    @patch("dcf_builder.excel_functions.df")
    def test_dcf_revenue_float_year(self, mock_df):
        """Test DCF_REVENUE handles float year from Excel."""
        mock_df.get_revenue.return_value = 400000000000

        # Excel sometimes passes years as floats
        result = ef.DCF_REVENUE("AAPL", 2023.0)
        mock_df.get_revenue.assert_called_with("AAPL", 2023)

    @patch("dcf_builder.excel_functions.df")
    def test_dcf_ebitda(self, mock_df):
        """Test DCF_EBITDA function."""
        mock_df.get_ebitda.return_value = 130000000000

        result = ef.DCF_EBITDA("AAPL", 2023)
        assert result == 130000000000

    @patch("dcf_builder.excel_functions.df")
    def test_dcf_wacc(self, mock_df):
        """Test DCF_WACC function."""
        mock_df.calculate_wacc.return_value = 0.10

        result = ef.DCF_WACC("AAPL")
        assert result == 0.10

    @patch("dcf_builder.excel_functions.df")
    def test_dcf_ev(self, mock_df):
        """Test DCF_EV function."""
        mock_df.get_stock_info.return_value = {"enterprise_value": 2600000000000}

        result = ef.DCF_EV("AAPL")
        assert result == 2600000000000

    @patch("dcf_builder.excel_functions.df")
    def test_dcf_pe(self, mock_df):
        """Test DCF_PE function."""
        mock_df.get_stock_info.return_value = {"trailing_pe": 25.5}

        result = ef.DCF_PE("AAPL")
        assert result == 25.5
