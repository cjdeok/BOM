from unified_server import _flexible_qmpc_date_to_iso
from datetime import datetime

v = "2024.09.23"
res = _flexible_qmpc_date_to_iso(v)
print(f"Input: {v} -> Result: {res}")

v2 = "2024-09-23 00:00:00"
res2 = _flexible_qmpc_date_to_iso(v2)
print(f"Input: {v2} -> Result: {res2}")
