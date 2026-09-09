# MIS Summarizer

A comprehensive data conversion tool with three powerful modes for processing financial data.

## Features

### 📊 MIS Summarizer
- Processes tax payment data from Excel files
- Groups transactions by financial year and section
- Generates formatted Excel reports with subtotals
- Automatically calculates financial years based on payment dates

### 🧾 Sale/Purchase Summarizer
- Processes invoice data from Excel files
- Groups by Buyer Name, Seller Name, or Tax Year
- Calculates subtotals and grand totals
- Generates formatted Excel reports with professional styling

### 🏦 Bank Statement Processor ⚠️ Meezan Bank Only
- **Extracts transactions from Meezan Bank statement PDFs**
- **Parses account information, dates, and balances**
- **Generates formatted Excel reports with:**
  - Account details and period information
  - All transactions with dates, descriptions, credits, debits, and running balance
  - Total credits and total debits
  - Opening and closing balances
- **Supported Format:** Meezan Bank online account statements (2025-2026 layout)
- **Other Banks:** Coming soon - easily extensible architecture

## Installation

This project uses [uv](https://github.com/astral-sh/uv) for dependency management.

```bash
# Clone the repository
git clone <repository-url>
cd MIS-Summarizer

# Install dependencies
uv sync
```

## Usage

### Running the Web Interface

```bash
uv run streamlit run src/mis_converter/app.py
```

The application will open in your browser at `http://localhost:8501`

### Using the Bank Statement Feature

⚠️ **Important**: Currently supports **Meezan Bank statement format only**. Other bank formats will be added in future updates.

1. **Select Mode**: Click on the "🏦 Meezan Bank" tab in the sidebar
2. **Upload PDF**: Upload your Meezan Bank statement PDF file
3. **Extract**: Click the "Extract" button to process the PDF
4. **Review**: View the extracted transactions in the table
5. **Download**: Click "Download Excel" to get your formatted report

### Excel Output Format

The bank statement Excel file includes:
- **Header**: Account title, number, IBAN, and period
- **Opening Balance**: Starting balance for the period
- **Transactions Table**: Date, Description, Credit, Debit, Balance columns
- **Totals**: Sum of all credits and debits
- **Closing Balance**: Ending balance for the period

## Dependencies

- Python >= 3.14
- streamlit >= 1.40.0
- pandas >= 3.0.3
- openpyxl >= 3.1.5
- pdfplumber >= 0.11.10
- xlrd >= 2.0.2

## Project Structure

```
MIS-Summarizer/
├── src/
│   └── mis_converter/
│       ├── app.py              # Streamlit web interface
│       ├── converter.py        # MIS data conversion logic
│       ├── sale_purchase.py    # Sale/Purchase conversion logic
│       └── bank_statement.py   # Bank statement PDF parser (NEW)
├── pyproject.toml
└── README.md
```

## Features in Detail

### Bank Statement Parser

The bank statement parser uses `pdfplumber` to extract text from PDF files and intelligently parses:

- **Account Information**: Account title, number, and IBAN
- **Date Range**: Statement period (from date to date)
- **Balances**: Opening and closing balances
- **Transactions**: 
  - Transaction date
  - Description (merchant name, transaction type)
  - Credit amounts (deposits, refunds)
  - Debit amounts (withdrawals, purchases)
  - Running balance after each transaction

**Current Support**: Meezan Bank statement format (tested with 2025-2026 statements)

The parser handles multi-line transaction descriptions and correctly identifies credits vs debits based on the `+` and `-` symbols in the PDF.

**Format-Specific**: This parser is tailored to Meezan Bank's specific PDF layout. Different banks use different formats for headers, dates, and transaction details.

## Development

### Adding Support for Other Bank Formats

The parser is designed to be extensible. To support additional banks:

**Step 1: Collect Sample PDFs**
- Get 2-3 sample statements from the target bank
- Anonymize sensitive information (account numbers, names, balances)

**Step 2: Analyze the Format**
- Check header layout (where account info appears)
- Identify date format (DD-MMM-YYYY, DD/MM/YYYY, etc.)
- Note transaction format (how credits/debits are marked)

**Step 3: Extend the Parser**

Modify `src/mis_converter/bank_statement.py`:

```python
def detect_bank_format(text: str) -> str:
    """Detect which bank format this PDF uses."""
    if "MEEZAN" in text.upper():
        return "meezan"
    elif "HBL" in text or "HABIB BANK" in text:
        return "hbl"
    elif "UBL" in text:
        return "ubl"
    return "unknown"

def extract_transactions(pdf_path):
    # Detect bank format
    first_page_text = ...
    bank_format = detect_bank_format(first_page_text)
    
    # Use appropriate parser
    if bank_format == "meezan":
        return _parse_meezan_format(pdf)
    elif bank_format == "hbl":
        return _parse_hbl_format(pdf)
    elif bank_format == "ubl":
        return _parse_ubl_format(pdf)
    else:
        raise ValueError(f"Unsupported bank format: {bank_format}")
```

**Step 4: Implement Bank-Specific Parsers**

1. Create `_parse_header_<bank>()` for that bank's header format
2. Create `_parse_transactions_<bank>()` for that bank's transaction format
3. Test with multiple statements from that bank
4. Update UI to show supported banks

**Step 5: Update UI**

Change the mode label and info messages to show multiple supported banks.

## License

[Add your license information here]

## Author

Muhammad Anus (mohdanus20@gmail.com)
