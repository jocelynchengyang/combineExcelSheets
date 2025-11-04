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

## How It Works

### Input Structure
Your Excel file has multiple sheets (e.g., "001", "002", ..., "017"), where each sheet contains:
- **Column A**: Nerve names (e.g., "T1Lt_Ant_Horn", "C8Lt_Ant_Horn", etc.)
- **Column B**: Measurement values

Example of original sheet "001":
```
| Nerve          | Value       |
|----------------|-------------|
| T1Lt_Ant_Horn  | 0.499393969 |
| C8Lt_Ant_Horn  | 0.523456789 |
| ...            | ...         |
```

### Output Structure
The script creates a single spreadsheet where:
- **Column A**: Sheet names (001, 002, 003, etc.)
- **Columns B onwards**: Each nerve measurement becomes its own column

Example output:
```
|     | T1Lt_Ant_Horn | C8Lt_Ant_Horn | ... |
|-----|---------------|---------------|-----|
| 001 | 0.499393969   | 0.523456789   | ... |
| 002 | 0.392274112   | 0.456789012   | ... |
| 003 | 0.387654321   | 0.498765432   | ... |
| ... | ...           | ...           | ... |
```

## Features

- ✅ Automatically processes all sheets in the Excel file
- ✅ Transposes data (sheets become rows, measurements become columns)
- ✅ Preserves all measurement values with full precision
- ✅ Can output as CSV or Excel format
- ✅ Provides progress feedback and data preview
- ✅ Error handling and validation

## Output Format

The first row (row 0) contains headers:
- A0: Empty (blank header for sheet names column)
- B0: First nerve name (e.g., "T1Lt_Ant_Horn")
- C0: Second nerve name (e.g., "C8Lt_Ant_Horn")
- And so on...

Subsequent rows contain:
- Column A: Sheet name
- Columns B onwards: Corresponding measurement values

## Troubleshooting

### Missing Dependencies
If you get an import error, install the required libraries:
```bash
pip install pandas openpyxl --break-system-packages
```

### File Not Found
Make sure the input Excel file path is correct and the file exists.

### Empty Output
Check that your Excel sheets have the expected structure:
- Column A with nerve names
- Column B with measurement values

## Notes

- The script preserves all numeric precision from the original data
- Sheet names should be numeric (001, 002, etc.) but the script works with any sheet names
- All sheets should have the same structure (same nerve names in the same order)
- The script automatically detects whether to save as CSV or Excel based on the output file extension

## Support

For issues or questions, check that:
1. Your Excel file is not corrupted
2. All sheets have the expected structure
3. You have the required Python libraries installed
4. You're using Python 3.6 or higher
