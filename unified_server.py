import sqlite3
import pandas as pd
from flask import Flask, jsonify, send_from_directory, request, send_file
import os
import math
import re
import glob
import io

app = Flask(__name__, static_folder='.')

# 디렉토리 설정
ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
MO_DIR = os.path.join(ROOT_DIR, 'MO')
DB_PATH = os.path.join(ROOT_DIR, 'bom.db')

def get_db_connection():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

# --- 공통 및 정적 파일 서버 ---
@app.route('/')
def index():
    return send_from_directory('.', 'index.html')

# --- BOM 조회 (Viewer) API ---
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

# --- BOM 생성 (Generator) API ---
@app.route('/api/lots')
def get_lots():
    files = []
    patterns = ["*_BOM_formula.csv", "BCE*_BOM_formula.csv", "BCQ*_BOM_formula.csv"]
    seen = set()
    for pattern in patterns:
        for filepath in glob.glob(os.path.join(MO_DIR, pattern)):
            filename = os.path.basename(filepath)
            if filename not in seen:
                files.append({"filename": filename})
                seen.add(filename)
    if not files:
        files = [{"filename": "BCE01_BOM_formula.csv"}]
    return jsonify(files)

@app.route('/api/generate_bom', methods=['POST'])
def generate_bom():
    data = request.json
    target_qty = float(data.get('target_qty', 0))
    requested_file = data.get('formula_file', 'BCE01_BOM_formula.csv')
    if target_qty <= 0:
        return jsonify({"error": "유효한 목표 생산량을 입력해주세요."}), 400
    try:
        safe_filename = os.path.basename(requested_file)
        target_path = os.path.join(MO_DIR, safe_filename)
        if not os.path.exists(target_path):
            return jsonify({"error": f"파일 '{safe_filename}'을(를) 찾을 수 없습니다."}), 404
        df = None
        ext = os.path.splitext(safe_filename)[1].lower()
        if ext == '.csv':
            try:
                df = pd.read_csv(target_path, encoding='utf-8')
            except:
                df = pd.read_csv(target_path, encoding='cp949')
        else:
            df = pd.read_excel(target_path)
        if 'Level' in df.columns:
            df['Level'] = df['Level'].ffill()
        def evaluate_formula(formula_str, target_val):
            def excel_if(cond, t, f): return t if cond else f
            def roundup(n, digits=0):
                factor = 10**digits
                return math.ceil(n * factor) / factor
            allowed_names = {
                "target_qty": float(target_val), "target_gty": float(target_val), "target": float(target_val),
                "round": round, "ROUND": round, "ROUNDUP": roundup, "ceil": math.ceil,
                "int": int, "float": float, "abs": abs, "ABS": abs, "IF": excel_if, "IFF": excel_if,
                "req": 1.0, "ratio": 1.0
            }
            try:
                f_str = str(formula_str).strip()
                if f_str.lower() == 'nan' or not f_str: return 0.0
                f_str = re.sub(r'\bif\s*\(', 'IFF(', f_str, flags=re.IGNORECASE)
                return float(eval(f_str, {"__builtins__": {}}, allowed_names))
            except Exception as e:
                return 0.0
        result = {
            "level0": {"제품명": "", "목표수량": target_qty, "생산수량": 0, "단위": "", "CodeNo": ""},
            "level1": [], "level2": [], "level3": []
        }
        for i, r in df.iterrows():
            lvl_raw = r.get('Level')
            if pd.isna(lvl_raw): continue
            lvl = str(int(lvl_raw)) if isinstance(lvl_raw, (int, float)) else str(lvl_raw).strip()
            cols = df.columns.tolist()
            parent = str(r[cols[1]]).strip() if len(cols) > 1 else ""
            name1 = str(r[cols[2]]).strip() if len(cols) > 2 else ""
            code_no = str(r[cols[3]]).strip() if len(cols) > 3 else ""
            formula = str(r[cols[5]]).strip() if len(cols) > 5 else ""
            unit = str(r[cols[6]]).strip() if len(cols) > 6 else ""
            name2 = str(r[cols[7]]).strip() if len(cols) > 7 else ""
            
            final_name = name2 if name2 and len(name2) > len(name1) and name2.lower() != 'nan' else name1
            if not final_name or final_name.lower() in ['nan', '']: continue
            
            calculated_val = evaluate_formula(formula, target_qty)
            item_dict = {
                "ID": f"gen_{lvl}_{i}",
                "상위연결": parent if parent.lower() != 'nan' else '',
                "구성품": final_name,
                "CodeNo": code_no if code_no.lower() != 'nan' else '',
                "LotNo": "", "계산된_소요량": round(calculated_val, 3),
                "단위": unit if unit.lower() != 'nan' else '', "레벨": lvl,
            }
            if lvl == '0':
                result["level0"] = {
                    "제품명": final_name, "구성품": final_name, "CodeNo": code_no,
                    "목표수량": target_qty, "생산수량": round(calculated_val, 3), "단위": unit
                }
            elif lvl == '1': result["level1"].append(item_dict)
            elif lvl == '2': result["level2"].append(item_dict)
            elif lvl == '3': result["level3"].append(item_dict)
        return jsonify(result)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/api/download_excel', methods=['POST'])
def download_excel():
    try:
        data = request.json
        l0 = data.get('level0', {})
        levels = [data.get('level1', []), data.get('level2', []), data.get('level3', [])]
        rows = []
        rows.append({
            "Level": "0", "상위 연결": "-", "구성품(제품명)": l0.get('제품명', '-'), 
            "Code No.": l0.get('제품코드', '-'), "Lot No.": l0.get('topLotNo', '-'), 
            "수량": l0.get('생산수량', 0), "단위": l0.get('단위', '')
        })
        rows.append({}) 
        for i, items in enumerate(levels):
            if items:
                for item in items:
                    rows.append({
                        "Level": str(i + 1), "상위 연결": item.get('상위연결', ''),
                        "구성품(제품명)": item.get('구성품', ''), "Code No.": item.get('CodeNo', ''),
                        "Lot No.": item.get('LotNo', ''), "수량": item.get('계산된_소요량', 0),
                        "단위": item.get('단위', '')
                    })
                rows.append({})
        df = pd.DataFrame(rows)
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='BOM_Result')
        output.seek(0)
        filename = f"BOM_Result_{l0.get('제품코드', 'Export')}.xlsx"
        return send_file(output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 
                         as_attachment=True, download_name=filename)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    print("--- Unified BOM System Server Starting ---")
    print("Accessible at: http://localhost:9000")
    app.run(host='0.0.0.0', port=9000, debug=False)
