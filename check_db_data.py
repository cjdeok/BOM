import sqlite3
import pandas as pd
import openpyxl
from openpyxl import load_workbook
import os

DB_PATH = 'bom.db'

def check_ema015():
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("SELECT code_no, description FROM item_master WHERE code_no LIKE 'EMA015%'")
    rows = cur.fetchall()
    print("EMA015 items in item_master:")
    for r in rows:
        print(r)
    
    # Check level1 for a sample lot
    cur.execute("SELECT DISTINCT \"상위Lot\" FROM level1 LIMIT 5")
    lots = cur.fetchall()
    for lot_tuple in lots:
        lot = lot_tuple[0]
        cur.execute("SELECT \"코드번호\", \"Lot No.\", \"포장시 요구량\" FROM level1 WHERE \"상위Lot\" = ? AND \"코드번호\" LIKE 'EMA015%'", (lot,))
        l1_rows = cur.fetchall()
        if l1_rows:
            print(f"\nFound EMA015 in level1 for lot {lot}:")
            for r in l1_rows:
                print(r)
    
    conn.close()

if __name__ == "__main__":
    check_ema015()
