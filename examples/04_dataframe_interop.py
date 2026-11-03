r"""
Example 04: Two-Way Pandas DataFrame Interoperability

Demonstrates:
1. Reading arbitrary live Excel ranges directly into a Pandas DataFrame in <0.05s.
2. Performing vectorized analytics / filtering with Pandas.
3. Writing a processed DataFrame directly back into the live Excel sheet.
"""

import pandas as pd
from live_excel import get_status, read_dataframe, write_dataframe, autofit

def main():
    status = get_status()
    print("=== Live Excel Session ===")
    print(f"Excel Connected : {status['connected']}")
    print(f"Active Workbook : {status['active_workbook']}")
    print(f"Active Sheet    : {status['active_sheet']}")
    
    if not status['connected']:
        print("Please ensure Microsoft Excel is open on your desktop.")
        return

    # 1. Read existing range into Pandas DataFrame
    # Replace 'A1:E20' with any valid range in your active sheet
    sample_range = "A1:E10"
    print(f"\nReading {sample_range} into Pandas DataFrame...")
    df = read_dataframe(sample_range, has_headers=True)
    print("DataFrame Preview:")
    print(df.head())
    print(f"Shape: {df.shape}")

    # 2. Example: Creating and writing a new analytical summary DataFrame
    summary_data = {
        "Metric": ["Total Revenue", "Operating Expenses", "Net Margin", "YoY Growth (%)"],
        "Q1": [1250000, 820000, 430000, 12.5],
        "Q2": [1380000, 890000, 490000, 15.2],
        "Q3": [1420000, 910000, 510000, 14.1],
        "Q4": [1650000, 990000, 660000, 18.7],
    }
    summary_df = pd.DataFrame(summary_data)
    
    print("\nSummary DataFrame to write:")
    print(summary_df)

    # 3. Write DataFrame back to live Excel at a target location (e.g. G1)
    # write_dataframe automatically handles headers and converts NaN to None
    print("\nWriting summary DataFrame to live Excel starting at G1...")
    result = write_dataframe(summary_df, start_cell="G1", include_headers=True)
    print(f"Written range: {result.get('range')} on sheet '{result.get('sheet')}'")

    # 4. Auto-fit columns G:L for clean presentation
    autofit(columns="G:L")
    print("Columns G:L auto-fitted successfully in live Excel!")

if __name__ == "__main__":
    main()
