import os
import openpyxl
from unified_server import _qmpc_meta_from_xlsx, _build_qmpc_catalog_rows

path = r"\\ens-nas918\회사공유폴더\회사공유폴더\품질경영시스템 표준문서\표준서\ESH-PC-BCE-01 품질관리 공정도\R2\ESH-PC-BCE01-01-R2 품질관리공정도(BCE01).xlsx"
if os.path.isfile(path):
    meta = _qmpc_meta_from_xlsx(path)
    print("META:", meta)
    rows = _build_qmpc_catalog_rows("BCE01", [{"filename": "test.xlsx", "modified": "2024-01-01T00:00:00"}], 2, meta)
    print("CATALOG_ROWS:", rows)
else:
    print("File not found at", path)
