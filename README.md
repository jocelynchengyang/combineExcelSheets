# Excel Sheet Combiner

This Python script combines multiple sheets from an Excel file into a single spreadsheet or CSV file, transposing the data so that each sheet becomes a row.

## Requirements

- Python 3.6 or higher
- pandas library
- openpyxl library (for Excel file handling)

## Installation

Install the required libraries:

```bash
pip install pandas openpyxl --break-system-packages
```

## Usage

### Basic Usage

```bash
python combine_excel_sheets.py <input_excel_file> [output_file]
```

### Examples

1. **Create a CSV file (default):**
   ```bash
   python combine_excel_sheets.py your_data.xlsx
   ```
   Output: `combined_output.csv`

2. **Specify output CSV file name:**
   ```bash
   python combine_excel_sheets.py your_data.xlsx combined_data.csv
   ```

3. **Create an Excel output file:**
   ```bash
   python combine_excel_sheets.py your_data.xlsx combined_data.xlsx
   ```

