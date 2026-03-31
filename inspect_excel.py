import pandas as pd

try:
    # Read the excel file
    # We use engine='openpyxl' because it's an xlsx file
    df = pd.read_excel(r'c:\Users\ENS-1000\Documents\Antigravity\BOM3\25BCE01-포장지시서.xlsx', header=None)
    
    # Print some interesting cells mentioned by user
    # Note: pandas uses 0-based indexing for rows and columns.
    # A=0, B=1, C=2, D=3, E=4, ...
    # E4 -> row 3, col 4
    # A7 -> row 6, col 0
    # J7 -> row 6, col 9
    # N7 -> row 6, col 13
    # S7 -> row 6, col 18
    # Z7 -> row 6, col 25
    # AE9 -> row 8, col 30
    
    cells = [
        ('E4', 3, 4),
        ('A7', 6, 0),
        ('J7', 6, 9),
        ('N7', 6, 13),
        ('S7', 6, 18),
        ('Z7', 6, 25),
        ('AE9', 8, 30),
        ('L21', 20, 11),
        ('S21', 20, 18),
        ('X21', 20, 23),
        ('AI21', 20, 34)
    ]
    
    for label, r, c in cells:
        try:
            val = df.iloc[r, c]
            print(f"{label}: {val}")
        except Exception as e:
            print(f"{label}: Error - {e}")

except Exception as e:
    print(f"Error reading excel: {e}")
