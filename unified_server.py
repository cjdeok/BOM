import sqlite3
import pandas as pd
from flask import Flask, jsonify, send_from_directory, request, send_file
import os
import io
import openpyxl
import tempfile
import urllib.parse

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

@app.route('/api/instruction_lots')
def get_instruction_lots():
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute('SELECT DISTINCT "LOT NO.", "제품명", "제조일자" FROM level0 ORDER BY "제조일자" DESC')
        rows = cursor.fetchall()
        conn.close()
        result = [{"lot_no": r["LOT NO."], "product_name": r["제품명"], "mfg_date": r["제조일자"]} for r in rows]
        return jsonify(result)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/api/instruction_detail/<lot_no>')
def get_instruction_detail(lot_no):
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute('SELECT * FROM level0 WHERE "LOT NO." = ?', (lot_no,))
        l0 = cursor.fetchone()
        if not l0:
            conn.close()
            return jsonify({"error": "Lot not found"}), 404
        l0_dict = dict(l0)
        
        cursor.execute('SELECT * FROM level1 WHERE "상위Lot" = ?', (lot_no,))
        l1 = [dict(r) for r in cursor.fetchall()]
        l1_lots = [r["Lot No."] for r in l1 if r["Lot No."]]
        
        l2 = []
        if l1_lots:
            p = ",".join(["?" for _ in l1_lots])
            cursor.execute(f'SELECT * FROM level2 WHERE "상위Lot" IN ({p})', l1_lots)
            l2 = [dict(r) for r in cursor.fetchall()]
        
        l2_lots = [r["Lot No."] for r in l2 if r["Lot No."]]
        l3 = []
        if l2_lots:
            p = ",".join(["?" for _ in l2_lots])
            cursor.execute(f'SELECT * FROM level3 WHERE "상위Lot" IN ({p})', l2_lots)
            l3 = [dict(r) for r in cursor.fetchall()]

        cursor.execute('SELECT * FROM instruction_summary WHERE "상위Lot" = ?', (lot_no,))
        summary = [dict(r) for r in cursor.fetchall()]
        conn.close()
        return jsonify({"level0": l0_dict, "level1": l1, "level2": l2, "level3": l3, "instruction_summary": summary})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

# --- 포장지시서 API ---
@app.route('/api/packaging_preview/<lot_no>')
def get_packaging_preview(lot_no):
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute('SELECT * FROM level0 WHERE "LOT NO." = ?', (lot_no,))
        l0 = cursor.fetchone()
        if not l0:
            conn.close()
            return jsonify({"error": "Lot not found"}), 404
        l0 = dict(l0)
        cursor.execute('SELECT * FROM level1 WHERE "상위Lot" = ?', (lot_no,))
        l1_items = [dict(r) for r in cursor.fetchall()]
        cursor.execute('SELECT doc_name FROM instruction_doc_master WHERE code_no = ? AND division LIKE "%PI%"', (l0['제품코드'],))
        doc = cursor.fetchone()
        doc_name = doc['doc_name'] if doc else ""
        conn.close()
        
        total_qty = l0.get('생산 수량(kit)') or 0
        return jsonify({
            "E4": doc_name, "A7": l0['제품명'], "J7": l0['제품버전'], "N7": total_qty,
            "S7": l0['제조일자'], "Z7": l0['유효기간'], "AE7": l0['LOT NO.'], "EMA015_items": l1_items
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/api/packaging_download/<lot_no>')
def download_packaging(lot_no):
    template_path = os.path.join(ROOT_DIR, '25BCE01-포장지시서.xlsx')
    if not os.path.exists(template_path):
        return jsonify({"error": "Template file not found"}), 404
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute('SELECT * FROM level0 WHERE "LOT NO." = ?', (lot_no,))
        l0 = cursor.fetchone()
        if not l0:
            conn.close()
            return jsonify({"error": "Lot not found"}), 404
        l0 = dict(l0)
        cursor.execute('SELECT * FROM level1 WHERE "상위Lot" = ?', (lot_no,))
        l1_items = [dict(r) for r in cursor.fetchall()]
        cursor.execute('SELECT doc_name FROM instruction_doc_master WHERE code_no = ? AND division LIKE "%PI%"', (l0['제품코드'],))
        doc = cursor.fetchone()
        doc_name = doc['doc_name'] if doc else ""
        conn.close()

        wb = openpyxl.load_workbook(template_path)
        ws = wb.active
        ws['E4']=doc_name; ws['A7']=l0['제품명']; ws['J7']=l0['제품버전']
        ws['S7']=l0['제조일자']; ws['Z7']=l0['유효기간']; ws['AE7']=l0['LOT NO.']
        
        mapping = {'EMA015':21,'EMA014':22,'CR(01)':23,'PC(01)':24,'NC(01)':25,'DA(01)':26,'RD(01)':27,'WS(01)':28,'TM(01)':29,'SS(01)':30,'EMA013':31,'PL(01)':32,'IFU':33}
        total_qty = l0.get('생산 수량(kit)') or 0

        for item in l1_items:
            code = str(item.get('코드번호') or '').strip()
            row_idx = next((row for key, row in mapping.items() if key in code), None)
            try:
                if row_idx and row_idx != 33:
                    ws[f'L{row_idx}'] = item.get('Lot No.')
                    ws[f'S{row_idx}'] = l0['제조일자']
                    ws[f'X{row_idx}'] = item.get('유효기간')
                    ws[f'AI{row_idx}'] = float(str(item.get('포장시 요구량') or 0).replace(',', ''))
            except: pass
            
        ws['L33']=""; ws['S33']=l0['제조일자']; ws['X33']=""; ws['AI33']=total_qty; ws['N7']=total_qty

        tmp_fd, tmp_name = tempfile.mkstemp(suffix='.xlsx')
        os.close(tmp_fd)
        wb.save(tmp_name); wb.close()
        return send_file(tmp_name, as_attachment=True, download_name=f'Packaging_Instruction_{lot_no}.xlsx', mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    except Exception as e:
        return jsonify({"error": str(e)}), 500

# --- 완제품 관리 API ---
@app.route('/api/product_management_preview/<lot_no>')
def get_product_management_preview(lot_no):
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute('SELECT * FROM level0 WHERE "LOT NO." = ?', (lot_no,))
        l0 = cursor.fetchone()
        if not l0:
            conn.close()
            return jsonify({"error": "Lot not found"}), 404
        l0 = dict(l0)
        cursor.execute('SELECT "포장시 요구량" FROM level1 WHERE "상위Lot" = ? AND "코드번호" LIKE "EMA015%"', (lot_no,))
        item = cursor.fetchone()
        ema015_qty = item[0] if item else 0
        conn.close()
        return jsonify({
            "A7": l0.get('제품명', ''), "I7": l0.get('제품코드', ''), "N7": l0.get('LOT NO.', ''),
            "T7": l0.get('제조일자', ''), "A9": l0.get('유효기간', ''), "I9": ema015_qty
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/api/product_management_download/<lot_no>')
def download_product_management(lot_no):
    template_path = os.path.join(ROOT_DIR, '25BCE01-완제품 관리.xlsx')
    if not os.path.exists(template_path):
        return jsonify({"error": "Template file not found"}), 404
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute('SELECT * FROM level0 WHERE "LOT NO." = ?', (lot_no,))
        l0 = cursor.fetchone()
        if not l0:
            conn.close()
            return jsonify({"error": "Lot not found"}), 404
        l0 = dict(l0)
        cursor.execute('SELECT "포장시 요구량" FROM level1 WHERE "상위Lot" = ? AND "코드번호" LIKE "EMA015%"', (lot_no,))
        item = cursor.fetchone()
        ema015_qty = item[0] if item else 0
        conn.close()
        
        wb = openpyxl.load_workbook(template_path)
        ws = wb.active
        ws['A7']=l0.get('제품명',''); ws['I7']=l0.get('제품코드',''); ws['N7']=l0.get('LOT NO.','')
        ws['T7']=l0.get('제조일자',''); ws['A9']=l0.get('유효기간',''); ws['I9']=ema015_qty
        tmp_fd, tmp_name = tempfile.mkstemp(suffix='.xlsx'); os.close(tmp_fd)
        wb.save(tmp_name); wb.close()
        return send_file(tmp_name, as_attachment=True, download_name=f'Product_Management_{lot_no}.xlsx', mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    print("--- Unified BOM System Server Starting ---")
    app.run(host='0.0.0.0', port=9000, debug=True)
