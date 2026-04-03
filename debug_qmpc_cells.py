import os
import openpyxl

path = r"\\ens-nas918\회사공유폴더\회사공유폴더\품질경영시스템 표준문서\표준서\ESH-PC-BCE-01 품질관리 공정도\R2\ESH-PC-BCE01-01-R2 품질관리공정도(BCE01).xlsx"
wb = openpyxl.load_workbook(path, data_only=True)
ws = wb["표지 (2)"]
print(f"Sheet: {ws.title}")
for i, row in enumerate(ws.iter_rows(values_only=True)):
    if i > 50: break
    # B, C, D, E columns
    if any(row[1:5]):
        print(f"Row {i+1}: B={row[1]}, C={row[2]}, D={row[3]}, E={row[4]}")
