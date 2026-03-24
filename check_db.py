import sqlite3

def print_schema():
    conn = sqlite3.connect('bom.db')
    cursor = conn.cursor()
    cursor.execute("SELECT name, sql FROM sqlite_master WHERE type='table'")
    tables = cursor.fetchall()
    for name, sql in tables:
        print(f"Table: {name}")
        print(f"Schema: {sql}\n")
    
    # 레코드 조금만 가져오기 (예: level1 등)
    if any(name == 'level1' for name, _ in tables):
        cursor.execute("SELECT * FROM level1 LIMIT 3")
        rows = cursor.fetchall()
        print("Level1 sample rows:", rows)
        
print_schema()
