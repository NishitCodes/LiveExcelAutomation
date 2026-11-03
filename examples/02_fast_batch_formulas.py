"""
Example 02: High-Speed 2D Matrix Formula Injection
Demonstrates how to inject thousands of formula cells in under 20ms using 2D array assignments.
"""

import sys
import os
import time

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import live_excel

def batch_formula_demo(app):
    wb, ws = live_excel.resolve_target(app)
    print(f"Target Sheet: {ws.Name}")
    
    start_row = 5
    num_rows = 100
    
    # Construct a 2D matrix of formulas: (num_rows x 3 columns)
    # Column D = B + C, Column E = D * 1.18 (tax), Column F = SUM(B:E)
    matrix = []
    for r in range(start_row, start_row + num_rows):
        row_formulas = [
            f"=B{r}+C{r}",
            f"=D{r}*1.18",
            f"=SUM(B{r}:E{r})"
        ]
        matrix.append(row_formulas)
        
    t0 = time.time()
    # Turn off screen updating for maximum speed
    prev_updating = app.ScreenUpdating
    app.ScreenUpdating = False
    try:
        ws.Range(f"D{start_row}:F{start_row + num_rows - 1}").Formula = matrix
    finally:
        app.ScreenUpdating = prev_updating
        
    t1 = time.time()
    print(f"Injected {num_rows * 3} formulas in {(t1 - t0)*1000:.2f} ms!")
    
    # Recalculate
    app.CalculateFull()
    print("Recalculation complete.")

if __name__ == "__main__":
    live_excel.run_in_live_excel(batch_formula_demo)
