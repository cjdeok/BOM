import sqlite3
import pandas as pd
from flask import Flask, render_template, request, jsonify, send_file
import sys
import os
import math
import re
import glob
import io

# 디렉토리 설정
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, '..', 'bom.db')

app = Flask(__name__)

def get_db_connection():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/lots')
def get_lots():
    # MO 디렉토리에서 수식 파일 리스트 검색 (CSV 파일만)
    files = []
    # 최신 CSV 파일 패턴 적용
    patterns = ["*_BOM_formula.csv", "BCE*_BOM_formula.csv", "BCQ*_BOM_formula.csv"]
    
    seen = set()
    for pattern in patterns:
        for filepath in glob.glob(os.path.join(BASE_DIR, pattern)):
            filename = os.path.basename(filepath)
            if filename not in seen:
                files.append({"filename": filename})
                seen.add(filename)
                
    # 파일이 하나도 없을 경우 기본값 반환
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
        target_path = os.path.join(BASE_DIR, safe_filename)
        
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
                "target_qty": float(target_val), 
                "target_gty": float(target_val),
                "target": float(target_val),
                "round": round, "ROUND": round,
                "ROUNDUP": roundup, "ceil": math.ceil,
                "int": int, "float": float, "abs": abs, "ABS": abs,
                "IF": excel_if, "IFF": excel_if,
                "req": 1.0, "ratio": 1.0
            }
            try:
                f_str = str(formula_str).strip()
                if f_str.lower() == 'nan' or not f_str: return 0.0
                f_str = re.sub(r'\bif\s*\(', 'IFF(', f_str, flags=re.IGNORECASE)
                return float(eval(f_str, {"__builtins__": {}}, allowed_names))
            except Exception as e:
                print(f"Eval Error on '{formula_str}': {e}")
                return 0.0

        result = {
            "level0": {"제품명": "", "목표수량": target_qty, "생산수량": 0, "단위": "", "CodeNo": ""},
            "level1": [], "level2": [], "level3": []
        }
        
        for _, r in df.iterrows():
            lvl_raw = r.get('Level')
            if pd.isna(lvl_raw): continue
            lvl = str(int(lvl_raw)) if isinstance(lvl_raw, (int, float)) else str(lvl_raw).strip()
            
            cols = df.columns.tolist()
            parent = str(r[cols[1]]).strip() if len(cols) > 1 else ""
            name1 = str(r[cols[2]]).strip() if len(cols) > 2 else ""
            name2 = str(r[cols[4]]).strip() if len(cols) > 4 else ""
            final_name = name2 if name2 and len(name2) > len(name1) and name2.lower() != 'nan' else name1
            
            code_no = str(r[cols[3]]).strip() if len(cols) > 3 else ""
            formula = str(r[cols[5]]).strip() if len(cols) > 5 else ""
            unit = str(r[cols[6]]).strip() if len(cols) > 6 else ""
            
            if not final_name or final_name.lower() in ['nan', '']: continue
            
            calculated_val = evaluate_formula(formula, target_qty)
            
            item_dict = {
                "상위연결": parent if parent.lower() != 'nan' else '',
                "구성품": final_name,
                "CodeNo": code_no if code_no.lower() != 'nan' else '',
                "LotNo": "",
                "계산된_소요량": round(calculated_val, 3),
                "단위": unit if unit.lower() != 'nan' else '',
                "레벨": lvl,
            }
            
            if lvl == '0':
                result["level0"] = {
                    "제품명": final_name,
                    "구성품": final_name,
                    "CodeNo": code_no,
                    "목표수량": target_qty,
                    "생산수량": round(calculated_val, 3),
                    "단위": unit
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
        # Level 0
        rows.append({
            "Level": "0", "상위 연결": "-", "구성품(제품명)": l0.get('제품명', '-'), 
            "Code No.": l0.get('제품코드', '-'), "Lot No.": l0.get('topLotNo', '-'), 
            "수량": l0.get('생산수량', 0), "단위": l0.get('단위', '')
        })
        rows.append({}) 
        
        # Levels 1, 2, 3
        for i, items in enumerate(levels):
            if items:
                for item in items:
                    rows.append({
                        "Level": str(i + 1),
                        "상위 연결": item.get('상위연결', ''),
                        "구성품(제품명)": item.get('구성품', ''),
                        "Code No.": item.get('CodeNo', ''),
                        "Lot No.": item.get('LotNo', ''),
                        "수량": item.get('계산된_소요량', 0),
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
    app.run(host='0.0.0.0', port=9001, debug=True)
