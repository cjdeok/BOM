import sqlite3
from flask import Flask, jsonify, send_from_directory
import os

app = Flask(__name__, static_folder='.')

DB_PATH = 'bom.db'

def get_db_connection():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

@app.route('/')
def index():
    return send_from_directory('.', 'index.html')

@app.route('/api/bom-all')
def get_bom_all():
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        
        tables = ['level0', 'level1', 'level2', 'level3', 'instruction_summary']
        result = {}
        
        for table in tables:
            cursor.execute(f"SELECT * FROM {table}")
            rows = cursor.fetchall()
            result[table] = [dict(row) for row in rows]
            
        conn.close()
        return jsonify(result)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    # 모든 인터페이스(0.0.0.0)에서 9000번 포트로 수신 대기
    print("--- BOM Viewer Server Starting ---")
    print("Accessible at: http://192.168.1.151:9000 (Local Network)")
    print("Accessible at: http://localhost:9000 (Local Machine)")
    app.run(host='0.0.0.0', port=9000, debug=False)
