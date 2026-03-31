import sqlite3
import pandas as pd
from flask import Flask, jsonify, send_from_directory, request, send_file
import os
import io

app = Flask(__name__, static_folder='.')

# 디렉토리 설정
ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(ROOT_DIR, 'bom.db')

def get_db_connection():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

# --- 공통 및 정적 파일 서버 ---
@app.route('/')
def index():
    return send_from_directory('.', 'index.html')

@app.route('/styles.css')
def serve_css():
    return send_from_directory('.', 'styles.css')

@app.route('/app.js')
def serve_js():
    return send_from_directory('.', 'app.js')

# --- BOM 조회 (Viewer) API ---
@app.route('/api/bom-all')
def get_bom_all():
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        tables = ['level0', 'level1', 'level2', 'level3', 'instruction_summary']
        result = {}
        for table in tables:
            if table in ['level1', 'level2', 'level3']:
                name_col = '구성품 명칭' if table == 'level1' else '원재료명'
                cursor.execute(f'''
                    SELECT l.*, i.description as _master_name 
                    FROM {table} l 
                    LEFT JOIN item_master i ON l."코드번호" = i.code_no
                ''')
                rows = []
                for r in cursor.fetchall():
                    d = dict(r)
                    if d.get('_master_name'):
                        d[name_col] = d['_master_name']
                    rows.append(d)
                result[table] = rows
            else:
                cursor.execute(f"SELECT * FROM {table}")
                rows = cursor.fetchall()
                result[table] = [dict(row) for row in rows]
        conn.close()
        return jsonify(result)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/api/item_master/<code_no>')
def get_item_master(code_no):
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM item_master WHERE code_no = ?", (code_no,))
        row = cursor.fetchone()
        conn.close()
        if row:
            return jsonify(dict(row))
        return jsonify({"error": "Not found"}), 404
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/api/doc_master/<code_no>')
def get_doc_master(code_no):
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM instruction_doc_master WHERE code_no = ?", (code_no,))
        rows = cursor.fetchall()
        conn.close()
        return jsonify([dict(row) for row in rows])
    except Exception as e:
        return jsonify({"error": str(e)}), 500

# --- BOM 제조지시 실행 데이터 저장/조회 ---
@app.route('/api/save_instruction', methods=['POST'])
def save_instruction():
    data = request.json
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        
        # 1. level0 저장 (스키마에 맞춰 필드명 매핑)
        l0 = data.get('level0', {})
        
        # item_master에서 category 조회
        cursor.execute('SELECT category FROM item_master WHERE code_no = ?', (l0.get('modelName'),))
        row = cursor.fetchone()
        l0_category = row['category'] if row else None

        cursor.execute("""
            INSERT INTO level0 ("Level", "제품코드", "구분", "제품명", "제품정보", "LOT NO.", "생산 수량(kit)", "제품버전", "제조일자", "의뢰팀", "생산목적", "유효기간")
            VALUES (0, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (l0.get('modelName'), l0_category, l0.get('productName'), l0.get('productInfo'), l0.get('lotNo'), l0.get('targetQty'), l0.get('version'), l0.get('mfgDate'), l0.get('requestTeam'), l0.get('purpose'), l0.get('expiryDate')))
        
        # 2. level1, 2, 3 저장
        for lvl in [1, 2, 3]:
            rows = data.get(f'level{lvl}', [])
            for r in rows:
                code_no = r.get('Code No.')
                # item_master에서 category 및 manufacturer 조회
                cursor.execute('SELECT category, manufacturer FROM item_master WHERE code_no = ?', (code_no,))
                m_row = cursor.fetchone()
                category = m_row['category'] if m_row else None
                mfr = m_row['manufacturer'] if m_row else None

                if lvl == 1:
                    cursor.execute(f"""
                        INSERT INTO level1 ("Level", "상위Lot", "코드번호", "구분", "구성품 명칭", "Lot No.", "제조일자", "유효기간", "포장시 요구량", "단위")
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (lvl, r.get('상위 Lot'), code_no, category, r.get('명칭 / 구성품'), r.get('할당 Lot'), l0['mfgDate'], r.get('유효기간'), r.get('할당수량'), r.get('단위')))
                else:
                    cursor.execute(f"""
                        INSERT INTO level{lvl} ("Level", "상위Lot", "코드번호", "구분", "원재료명", "제조사", "Lot No.", "제조일자", "유효기간", "제조량", "단위")
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (lvl, r.get('상위 Lot'), code_no, category, r.get('명칭 / 구성품'), mfr, r.get('할당 Lot'), l0['mfgDate'], r.get('유효기간'), r.get('할당수량'), r.get('단위')))
        
        # 3. instruction_summary 저장
        semi_lots = data.get('instruction_summary', [])
        for s in semi_lots:
            cursor.execute("""
                INSERT INTO instruction_summary ("상위Lot", "약어", "제조지침서 No.", "Lot. No.", "생산량", "제조일자")
                VALUES (?, ?, ?, ?, ?, ?)
            """, (l0['lotNo'], s.get('division'), s.get('latest_doc_no'), s.get('calcLot'), l0['targetQty'], s.get('mfgDate')))
            
        conn.commit()
        conn.close()
        return jsonify({"status": "success"})
    except Exception as e:
        if 'conn' in locals(): conn.close()
        # 에러 발생 시 JSON으로 반환하도록 보장
        return jsonify({"error": str(e)}), 500

@app.route('/api/instruction_lots')
def get_instruction_lots():
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        # 필드명을 "LOT NO.", "제품명", "제조일자"로 변경
        cursor.execute('SELECT DISTINCT "LOT NO.", "제품명", "제조일자" FROM level0 ORDER BY "제조일자" DESC')
        rows = cursor.fetchall()
        conn.close()
        # 클라이언트(app.js)에서 기대하는 키값(lot_no, product_name, mfg_date)으로 변환하여 반환
        result = []
        for r in rows:
            result.append({
                "lot_no": r["LOT NO."],
                "product_name": r["제품명"],
                "mfg_date": r["제조일자"]
            })
        return jsonify(result)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/api/instruction_detail/<lot_no>')
def get_instruction_detail(lot_no):
    try:
        conn = get_db_connection()
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
        
        # 1. Level 0
        cursor.execute('SELECT * FROM level0 WHERE "LOT NO." = ?', (lot_no,))
        l0_row = cursor.fetchone()
        if not l0_row:
            conn.close()
            return jsonify({"error": "Lot not found"}), 404
        l0 = dict(l0_row)
        
        # 2. Level 1 (상위Lot이 Level 0의 LOT NO.인 항목)
        cursor.execute('''
            SELECT l.*, i.description as _master_name 
            FROM level1 l 
            LEFT JOIN item_master i ON l."코드번호" = i.code_no 
            WHERE l."상위Lot" = ?
        ''', (lot_no,))
        l1_rows = []
        for r in cursor.fetchall():
            d = dict(r)
            if d.get('_master_name'): d['구성품 명칭'] = d['_master_name']
            l1_rows.append(d)
        l1_lots = [r["Lot No."] for r in l1_rows if r["Lot No."]]
        
        # 3. Level 2 (상위Lot이 Level 1의 Lot No.들 중 하나인 항목)
        l2_rows = []
        if l1_lots:
            placeholders = ",".join(["?" for _ in l1_lots])
            cursor.execute(f'''
                SELECT l.*, i.description as _master_name 
                FROM level2 l 
                LEFT JOIN item_master i ON l."코드번호" = i.code_no 
                WHERE l."상위Lot" IN ({placeholders})
            ''', l1_lots)
            for r in cursor.fetchall():
                d = dict(r)
                if d.get('_master_name'): d['원재료명'] = d['_master_name']
                l2_rows.append(d)
        
        l2_lots = [r["Lot No."] for r in l2_rows if r["Lot No."]]
        
        # 4. Level 3 (상위Lot이 Level 2의 Lot No.들 중 하나인 항목)
        l3_rows = []
        if l2_lots:
            placeholders = ",".join(["?" for _ in l2_lots])
            cursor.execute(f'''
                SELECT l.*, i.description as _master_name 
                FROM level3 l 
                LEFT JOIN item_master i ON l."코드번호" = i.code_no 
                WHERE l."상위Lot" IN ({placeholders})
            ''', l2_lots)
            for r in cursor.fetchall():
                d = dict(r)
                if d.get('_master_name'): d['원재료명'] = d['_master_name']
                l3_rows.append(d)
            
        # 5. instruction_summary
        cursor.execute('SELECT * FROM instruction_summary WHERE "상위Lot" = ?', (lot_no,))
        summary = [dict(r) for r in cursor.fetchall()]
        
        conn.close()
        return jsonify({
            "level0": l0,
            "level1": l1_rows,
            "level2": l2_rows,
            "level3": l3_rows,
            "instruction_summary": summary
        })
    except Exception as e:
        if 'conn' in locals(): conn.close()
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    print("--- Unified BOM System Server Starting ---")
    print("Accessible at: http://localhost:9000")
    app.run(host='0.0.0.0', port=9000, debug=True)
