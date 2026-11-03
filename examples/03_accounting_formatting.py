"""
Example 03: Accounting Table Formatting
Demonstrates applying header fills, number formats, thin grid borders, and double underlines.
"""

import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import live_excel

def formatting_demo(app):
    wb, ws = live_excel.resolve_target(app)
    
    # 1. Header Row (A1:E1): Dark Blue Background, Bold White Font, Centered
    hdr = ws.Range("A1:E1")
    hdr.Interior.ColorIndex = 49  # Dark Blue
    hdr.Font.Color = 16777215     # Pure White
    hdr.Font.Bold = True
    hdr.Font.Name = "Arial"
    hdr.Font.Size = 11
    hdr.HorizontalAlignment = live_excel.xlCenter
    hdr.VerticalAlignment = live_excel.xlCenter
    
    # 2. Data Grid (A2:E20): Number Format & Thin Borders
    data = ws.Range("B2:E20")
    data.NumberFormat = "#,##0.00;(#,##0.00);\"-\""
    data.HorizontalAlignment = live_excel.xlRight
    
    for edge in [
        live_excel.xlEdgeTop, live_excel.xlEdgeBottom,
        live_excel.xlEdgeLeft, live_excel.xlEdgeRight,
        live_excel.xlInsideHorizontal, live_excel.xlInsideVertical
    ]:
        data.Borders(edge).LineStyle = live_excel.xlContinuous
        data.Borders(edge).Weight = live_excel.xlThin
        
    # 3. Total Row (A21:E21): Bold, Soft Blue Fill, Double Accounting Underline
    total = ws.Range("A21:E21")
    total.Font.Bold = True
    total.Interior.ColorIndex = 37 # Sky blue fill
    total.Borders(live_excel.xlEdgeTop).LineStyle = live_excel.xlContinuous
    total.Borders(live_excel.xlEdgeTop).Weight = live_excel.xlThin
    total.Borders(live_excel.xlEdgeBottom).LineStyle = live_excel.xlDouble # Accounting double line
    
    print("Formatting applied successfully.")

if __name__ == "__main__":
    live_excel.run_in_live_excel(formatting_demo)
