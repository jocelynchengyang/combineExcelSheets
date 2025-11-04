import pandas as pd
import sys
import os

def combine_excel_sheets(input_file, output_file='combined_output.csv', save_excel=False):
    """
    Combine multiple Excel sheets into one spreadsheet/CSV file.
    
    Each sheet becomes a row in the output, with:
    - Column A: Sheet name (e.g., "001", "002", etc.)
    - Columns B onwards: Combinations of Nerve + Metric names
      (e.g., "T1_Left_Ant_Horn_FA", "T1_Left_Hemicord_FA", "C8_Left_Ant_Horn_FA", etc.)
    
    Args:
        input_file: Path to the Excel file with multiple sheets
        output_file: Path for the output file (CSV or Excel)
        save_excel: If True, save as Excel instead of CSV
    """
    
    # Check if input file exists
    if not os.path.exists(input_file):
        raise FileNotFoundError(f"Input file not found: {input_file}")
    
    # Read all sheets from the Excel file
    print(f"Reading Excel file: {input_file}")
    excel_file = pd.ExcelFile(input_file)
    
    print(f"Found {len(excel_file.sheet_names)} sheets: {', '.join(excel_file.sheet_names)}")
    
    # List to store all rows of data
    all_data = []
    
    # Process each sheet
    for sheet_name in excel_file.sheet_names:
        print(f"Processing sheet: {sheet_name}")
        
        # Read the sheet
        df = pd.read_excel(excel_file, sheet_name=sheet_name)
        
        # Remove rows where Nerve is NaN (empty rows at the end)
        df = df.dropna(subset=['Nerve'])
        
        # Create a dictionary for this row starting with sheet name
        row_data = {}
        
        # Get all metric columns (everything except 'slice' and 'Nerve' and blank columns)
        metric_columns = [col for col in df.columns 
                         if col not in ['slice', 'Nerve'] 
                         and not col.startswith('__BLANK')]
        
        # For each nerve (row in the original sheet)
        for idx, row in df.iterrows():
            nerve_name = str(row['Nerve']).strip()
            
            # For each metric column, create a combined column name
            for metric_col in metric_columns:
                # Create column name as "Nerve_MetricName" keeping original names
                # e.g., "T1_Left_Ant_Horn_FA", "C8_Left_Hemicord_FA", etc.
                combined_col_name = f"{nerve_name}_{metric_col}"
                row_data[combined_col_name] = row[metric_col]
        
        # Add sheet name as the first column
        row_data = {'Sheet': sheet_name, **row_data}
        
        # Add this row to our collection
        all_data.append(row_data)
    
    # Create a DataFrame from all collected data
    combined_df = pd.DataFrame(all_data)
    
    # Sort by sheet name to get them in order (001, 002, etc.)
    combined_df = combined_df.sort_values('Sheet').reset_index(drop=True)
    
    # The 'Sheet' column should be first
    cols = ['Sheet'] + [col for col in combined_df.columns if col != 'Sheet']
    combined_df = combined_df[cols]
    
    # Rename 'Sheet' column to empty string per specification
    combined_df.rename(columns={'Sheet': ''}, inplace=True)
    
    # Save to file
    if save_excel or output_file.endswith('.xlsx'):
        combined_df.to_excel(output_file, index=False)
    else:
        combined_df.to_csv(output_file, index=False)
    
    print(f"\n✓ Successfully combined {len(excel_file.sheet_names)} sheets!")
    print(f"✓ Output saved to: {output_file}")
    print(f"✓ Output shape: {combined_df.shape[0]} rows × {combined_df.shape[1]} columns")
    print(f"\nFirst few column names:")
    for i, col in enumerate(combined_df.columns[:6]):
        if col == '':
            print(f"  Column {i}: [empty - contains sheet names]")
        else:
            print(f"  Column {i}: {col}")
    
    return combined_df

if __name__ == "__main__":
    # Check if input file is provided
    if len(sys.argv) < 2:
        print("=" * 70)
        print("Excel Sheet Combiner")
        print("=" * 70)
        print("\nUsage: python combine_excel_sheets.py <input_excel_file> [output_file]")
        print("\nExamples:")
        print("  python combine_excel_sheets.py data.xlsx")
        print("  python combine_excel_sheets.py data.xlsx combined_data.csv")
        print("  python combine_excel_sheets.py data.xlsx combined_data.xlsx")
        print("\nThe script will:")
        print("  • Read all sheets from the input Excel file")
        print("  • Transpose the data so each sheet becomes a row")
        print("  • Column A: Sheet names (001, 002, etc.)")
        print("  • Columns B+: Nerve measurements (T1Lt_Ant_Horn, C8Lt_Ant_Horn, etc.)")
        print("  • Save to CSV by default (or Excel if output ends with .xlsx)")
        print("=" * 70)
        sys.exit(1)
    
    input_file = sys.argv[1]
    output_file = sys.argv[2] if len(sys.argv) > 2 else 'combined_output.csv'
    
    try:
        print("=" * 70)
        print("Excel Sheet Combiner")
        print("=" * 70)
        print()
        
        combined_df = combine_excel_sheets(input_file, output_file)
        
        print("\n" + "=" * 70)
        print("✓ Process completed successfully!")
        print("=" * 70)
        
        # Show a preview of the data
        print("\nPreview of combined data (first 3 rows, first 5 columns):")
        print(combined_df.iloc[:3, :5].to_string())
        
    except Exception as e:
        print(f"\n{'=' * 70}")
        print(f"✗ Error: {e}")
        print("=" * 70)
        import traceback
        traceback.print_exc()
        sys.exit(1)
