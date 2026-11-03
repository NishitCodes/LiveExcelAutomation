"""
Example 01: Inspect Live Excel Application & Workbooks
"""

import sys
import os

# Add root folder to sys.path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import live_excel

def main():
    print("Checking live Excel status...")
    try:
        status = live_excel.get_status()
        print(f"Excel Version: {status['excel_version']}")
        print(f"Open Workbooks Count: {status['workbooks_open_count']}")
        print(f"Active Workbook: {status['active_workbook']}")
        print(f"Active Sheet: {status['active_sheet']}")
        
        print("\nOpen Workbooks:")
        for wb in status["workbooks"]:
            print(f" - {wb['name']} ({wb['sheets_count']} sheets, Saved: {wb['saved']})")
            
        if status["active_workbook"]:
            sheets_info = live_excel.list_sheets()
            print(f"\nSheets in {sheets_info['workbook']}:")
            for s in sheets_info["sheets"]:
                print(f" - {s['name']}: UsedRange = {s['used_range']}")
                
    except RuntimeError as e:
        print(f"Could not connect: {e}")
        print("Please ensure Microsoft Excel is running with a workbook open.")

if __name__ == "__main__":
    main()
