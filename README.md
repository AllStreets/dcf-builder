# DCF Builder

Python-powered Excel add-in that helps build DCF valuation models with automated data fetching.

## Installation

```bash
# Install dependencies
pip install -e .

# Or install from requirements
pip install -r requirements.txt
```

## Setup

### FRED API Key (Required for Treasury Rates)

1. Get a free API key at: https://fred.stlouisfed.org/docs/api/api_key.html
2. Set the environment variable:

```bash
# macOS/Linux - add to ~/.bashrc or ~/.zshrc
export FRED_API_KEY="your_api_key_here"

# Windows
set FRED_API_KEY=your_api_key_here
```

## Command-Line Usage

```bash
# Generate a DCF model for Apple
dcf-builder generate AAPL

# Generate with custom output path
dcf-builder generate MSFT --output ~/Desktop/msft_dcf.xlsx

# Get current stock price
dcf-builder price AAPL

# Get stock info
dcf-builder info AAPL

# Clear cached data
dcf-builder refresh
```

## Python Usage

```python
from dcf_builder.template_generator import generate_dcf_model
from dcf_builder import data_fetcher as df

# Generate a complete DCF model
output_path = generate_dcf_model("AAPL")
print(f"Model saved to: {output_path}")

# Fetch individual data points
price = df.get_price("AAPL")
beta = df.get_beta("AAPL")
risk_free = df.get_risk_free_rate()
wacc = df.calculate_wacc("AAPL")

print(f"Price: ${price}")
print(f"Beta: {beta}")
print(f"Risk-Free Rate: {risk_free:.2%}")
print(f"WACC: {wacc:.2%}")
```

## Excel Integration (xlwings)

### Setup Steps

1. **Install xlwings add-in:**
   ```bash
   xlwings addin install
   ```

2. **Create an xlwings config file** in your project root or Excel file location:

   Create a file named `dcf_builder.conf`:
   ```ini
   [xlwings]
   PYTHONPATH=/path/to/dcf-builder
   UDF_MODULES=dcf_builder.main
   ```

3. **Open Excel** and you should see an "xlwings" tab in the ribbon.

4. **Import UDFs:**
   - Go to the xlwings tab
   - Click "Import Functions"
   - The DCF functions will now be available

### Custom Excel Functions

Once configured, you can use these functions in Excel cells:

| Function | Description | Example |
|----------|-------------|---------|
| `=dcf_price("AAPL")` | Current stock price | `$150.00` |
| `=dcf_market_cap("AAPL")` | Market cap in millions | `2,500,000` |
| `=dcf_beta("AAPL")` | 5-year monthly beta | `1.20` |
| `=dcf_shares_out("AAPL")` | Shares outstanding | `16,000,000,000` |
| `=dcf_52w_high("AAPL")` | 52-week high | `$180.00` |
| `=dcf_52w_low("AAPL")` | 52-week low | `$120.00` |
| `=dcf_risk_free()` | 10-year Treasury rate | `0.04` |
| `=dcf_revenue("AAPL", 2023)` | Revenue for year | `400,000,000,000` |
| `=dcf_ebitda("AAPL", 2023)` | EBITDA for year | `130,000,000,000` |
| `=dcf_wacc("AAPL")` | Calculated WACC | `0.10` |
| `=dcf_ev("AAPL")` | Enterprise value | `2,600,000,000,000` |
| `=dcf_pe("AAPL")` | Trailing P/E ratio | `25.5` |

### Creating a Custom Ribbon (Advanced)

To add a "DCF Builder" tab to your Excel ribbon:

1. Create an Excel file and save as `.xlsm` (macro-enabled)

2. Open the VBA editor (Alt+F11)

3. Add a reference to xlwings:
   - Tools → References
   - Check "xlwings"

4. Create a custom ribbon XML file or use xlwings quickstart:
   ```bash
   xlwings quickstart dcf_project
   ```

5. The quickstart creates a project with:
   - `dcf_project.xlsm` - Excel file with ribbon
   - `dcf_project.py` - Python backend

6. Modify `dcf_project.py` to import from dcf_builder:
   ```python
   from dcf_builder.main import ribbon_generate_dcf, ribbon_refresh_data
   ```

For detailed xlwings documentation, see: https://docs.xlwings.org/

## Project Structure

```
dcf-builder/
├── dcf_builder/
│   ├── __init__.py
│   ├── config.py           # Configuration and constants
│   ├── data_fetcher.py     # API calls with caching
│   ├── template_generator.py # Creates DCF workbooks
│   ├── excel_functions.py  # Custom Excel UDFs
│   └── main.py             # CLI and xlwings entry point
├── templates/
│   └── base_dcf.xlsx       # Base template (generated)
├── tests/
│   ├── test_data_fetcher.py
│   ├── test_template.py
│   └── test_excel_functions.py
├── requirements.txt
├── setup.py
└── README.md
```

## Data Sources

- **Stock Data**: Yahoo Finance (via yfinance)
- **Treasury Rates**: FRED API (10-year Treasury)
- **Financials**: Yahoo Finance (income statement, balance sheet)

## Caching

Data is cached to reduce API calls:
- Market data (price, beta, etc.): 15 minutes
- Historical financials: 24 hours
- Treasury rates: 1 hour

Cache location: `~/.dcf_builder_cache/`

To clear the cache:
```bash
dcf-builder refresh
```

Or in Python:
```python
from dcf_builder import data_fetcher as df
df.clear_cache()
```

## Running Tests

```bash
pytest tests/
```

## License

MIT
