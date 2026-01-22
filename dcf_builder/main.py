"""Main entry point and xlwings ribbon callbacks for DCF Builder.

This module provides:
1. Command-line interface for generating DCF models
2. xlwings ribbon callbacks for Excel integration
3. UDF registration for custom Excel functions

For Excel integration, this module is called by xlwings when:
- The add-in is loaded
- Ribbon buttons are clicked
- Custom functions are evaluated
"""

import argparse
import sys
from pathlib import Path
from typing import Optional

# Try to import xlwings (may not be available in all environments)
try:
    import xlwings as xw
    XLWINGS_AVAILABLE = True
except ImportError:
    XLWINGS_AVAILABLE = False

from . import data_fetcher as df
from .template_generator import generate_dcf_model
from . import excel_functions


def generate_dcf(ticker: str, output_path: Optional[str] = None) -> Path:
    """Generate a DCF model for the given ticker.

    Args:
        ticker: Stock ticker symbol
        output_path: Optional output file path

    Returns:
        Path to the generated Excel file
    """
    path = Path(output_path) if output_path else None
    return generate_dcf_model(ticker, path)


def refresh_data() -> None:
    """Clear cache and refresh all data."""
    df.clear_cache()
    print("Cache cleared. Data will be refreshed on next request.")


# ============================================================================
# xlwings Ribbon Callbacks
# ============================================================================

if XLWINGS_AVAILABLE:

    @xw.sub
    def ribbon_generate_dcf():
        """Ribbon callback: Generate DCF model.

        Prompts user for ticker and generates a complete DCF model.
        """
        # Get the calling workbook
        wb = xw.Book.caller()

        # Prompt for ticker (using Excel's InputBox)
        ticker = wb.app.api.InputBox(
            "Enter stock ticker symbol:",
            "DCF Builder - Generate Model",
            Type=2,  # Text
        )

        if not ticker or ticker == "False":
            return

        ticker = str(ticker).upper().strip()

        # Show status
        wb.app.status_bar = f"Generating DCF model for {ticker}..."

        try:
            # Generate the model
            output_path = generate_dcf_model(ticker)

            # Open the generated workbook
            xw.Book(str(output_path))

            wb.app.status_bar = f"DCF model generated: {output_path}"

        except Exception as e:
            wb.app.api.Alert(
                f"Error generating DCF model: {str(e)}",
                "DCF Builder Error",
            )
            wb.app.status_bar = ""

    @xw.sub
    def ribbon_refresh_data():
        """Ribbon callback: Refresh all market data."""
        wb = xw.Book.caller()
        wb.app.status_bar = "Refreshing data..."

        df.clear_cache()

        # Recalculate the workbook to update all UDFs
        wb.app.calculate()

        wb.app.status_bar = "Data refreshed."

    @xw.sub
    def ribbon_clear_cache():
        """Ribbon callback: Clear the data cache."""
        df.clear_cache()

        wb = xw.Book.caller()
        wb.app.api.Alert(
            "Cache cleared. Data will be fetched fresh on next request.",
            "DCF Builder",
        )

    # Register UDFs with xlwings
    @xw.func
    def dcf_price(ticker: str) -> Optional[float]:
        """Excel UDF: Get current stock price."""
        return excel_functions.DCF_PRICE(ticker)

    @xw.func
    def dcf_market_cap(ticker: str) -> Optional[float]:
        """Excel UDF: Get market cap in millions."""
        return excel_functions.DCF_MARKET_CAP(ticker)

    @xw.func
    def dcf_beta(ticker: str) -> Optional[float]:
        """Excel UDF: Get beta."""
        return excel_functions.DCF_BETA(ticker)

    @xw.func
    def dcf_shares_out(ticker: str) -> Optional[float]:
        """Excel UDF: Get shares outstanding."""
        return excel_functions.DCF_SHARES_OUT(ticker)

    @xw.func
    def dcf_52w_high(ticker: str) -> Optional[float]:
        """Excel UDF: Get 52-week high."""
        return excel_functions.DCF_52W_HIGH(ticker)

    @xw.func
    def dcf_52w_low(ticker: str) -> Optional[float]:
        """Excel UDF: Get 52-week low."""
        return excel_functions.DCF_52W_LOW(ticker)

    @xw.func
    def dcf_risk_free() -> Optional[float]:
        """Excel UDF: Get risk-free rate."""
        return excel_functions.DCF_RISK_FREE()

    @xw.func
    def dcf_revenue(ticker: str, year: int) -> Optional[float]:
        """Excel UDF: Get revenue for year."""
        return excel_functions.DCF_REVENUE(ticker, year)

    @xw.func
    def dcf_ebitda(ticker: str, year: int) -> Optional[float]:
        """Excel UDF: Get EBITDA for year."""
        return excel_functions.DCF_EBITDA(ticker, year)

    @xw.func
    def dcf_wacc(ticker: str) -> Optional[float]:
        """Excel UDF: Calculate WACC."""
        return excel_functions.DCF_WACC(ticker)

    @xw.func
    def dcf_ev(ticker: str) -> Optional[float]:
        """Excel UDF: Get enterprise value."""
        return excel_functions.DCF_EV(ticker)

    @xw.func
    def dcf_pe(ticker: str) -> Optional[float]:
        """Excel UDF: Get P/E ratio."""
        return excel_functions.DCF_PE(ticker)


# ============================================================================
# Command-Line Interface
# ============================================================================

def main():
    """Command-line entry point."""
    parser = argparse.ArgumentParser(
        description="DCF Builder - Generate DCF valuation models",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  dcf-builder generate AAPL
  dcf-builder generate MSFT --output ~/Desktop/msft_dcf.xlsx
  dcf-builder refresh
  dcf-builder price AAPL

Excel Integration:
  After installing, run 'xlwings addin install' to add the ribbon to Excel.
        """,
    )

    subparsers = parser.add_subparsers(dest="command", help="Available commands")

    # Generate command
    gen_parser = subparsers.add_parser("generate", help="Generate a DCF model")
    gen_parser.add_argument("ticker", help="Stock ticker symbol")
    gen_parser.add_argument(
        "--output", "-o", help="Output file path (default: DCF_TICKER_DATE.xlsx)"
    )

    # Refresh command
    subparsers.add_parser("refresh", help="Clear data cache")

    # Price command
    price_parser = subparsers.add_parser("price", help="Get current stock price")
    price_parser.add_argument("ticker", help="Stock ticker symbol")

    # Info command
    info_parser = subparsers.add_parser("info", help="Get stock info")
    info_parser.add_argument("ticker", help="Stock ticker symbol")

    args = parser.parse_args()

    if args.command == "generate":
        print(f"Generating DCF model for {args.ticker}...")
        output = generate_dcf(args.ticker, args.output)
        print(f"Model generated: {output}")

    elif args.command == "refresh":
        refresh_data()

    elif args.command == "price":
        price = df.get_price(args.ticker.upper())
        if price:
            print(f"{args.ticker.upper()}: ${price:.2f}")
        else:
            print(f"Could not fetch price for {args.ticker}")
            sys.exit(1)

    elif args.command == "info":
        info = df.get_stock_info(args.ticker.upper())
        if info:
            print(f"\n{info.get('name', args.ticker.upper())} ({args.ticker.upper()})")
            print("-" * 40)
            print(f"Price:          ${info.get('price', 'N/A')}")
            print(f"Market Cap:     ${info.get('market_cap', 0) / 1e9:.2f}B")
            print(f"Beta:           {info.get('beta', 'N/A')}")
            print(f"52-Week High:   ${info.get('fifty_two_week_high', 'N/A')}")
            print(f"52-Week Low:    ${info.get('fifty_two_week_low', 'N/A')}")
            print(f"Sector:         {info.get('sector', 'N/A')}")
            print(f"Industry:       {info.get('industry', 'N/A')}")
        else:
            print(f"Could not fetch info for {args.ticker}")
            sys.exit(1)

    else:
        parser.print_help()


if __name__ == "__main__":
    main()
