import pandas as pd

try:
    df = pd.read_excel(r'c:\Users\ENS-1000\Documents\Antigravity\BOM3\25BCE01-포장지시서.xlsx', header=None)
    
    print("--- Row 1 to 10 ---")
    print(df.iloc[0:10, 0:35].to_string())
    
    print("\n--- Row 20 to 35 ---")
    print(df.iloc[20:35, 0:40].to_string())

except Exception as e:
    print(f"Error: {e}")
