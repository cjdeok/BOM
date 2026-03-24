import sqlite3
import pandas as pd
from flask import Flask, render_template, request, jsonify
import sys
import os
import math
import re
import glob

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
    # MO 디렉토리에서 수식 파일 리스트 검색 (_BOM_fomula_ 또는 _수식마스터 패턴)
    files = []
    # CSV 및 XLSX 파일 검색
    patterns = ["*_BOM_fomula_*.csv", "*_수식마스터*.xlsx", "*.csv"] # 사용자 제공 파일명 패턴 고려
    
    seen = set()
    for pattern in patterns:
        for filepath in glob.glob(os.path.join(BASE_DIR, pattern)):
            filename = os.path.basename(filepath)
            if filename not in seen:
                files.append({"filename": filename})
                seen.add(filename)
                
    # 파일이 하나도 없을 경우 기본값 반환
    if not files:
        files = [{"filename": "BCE01_BOM_fomula_1.csv"}]
        
    return jsonify(files)

@app.route('/api/generate_bom', methods=['POST'])
def generate_bom():
    data = request.json
    target_qty = float(data.get('target_qty', 0))
    # 클라이언트가 요청한 수식 파일명 (기본값 설정)
    requested_file = data.get('formula_file', 'BCE01_BOM_fomula_1.csv')
    
    if target_qty <= 0:
        return jsonify({"error": "유효한 목표 생산량을 입력해주세요."}), 400
        
    try:
        # 파일 경로 보안 검증 (상위 디렉토리 접근 차단)
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
            
        # Level 컬럼 결측치 채우기
        if 'Level' in df.columns:
            df['Level'] = df['Level'].ffill()
            
        cols = df.columns.tolist()
        
        # 수식 평가기
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
            "level0": {"제품명": f"전개 소스: {safe_filename}", "목표수량": target_qty},
            "level1": [], "level2": [], "level3": []
        }
        
        for _, r in df.iterrows():
            lvl_raw = r.get('Level')
            if pd.isna(lvl_raw): continue
            lvl = str(int(lvl_raw)) if isinstance(lvl_raw, (int, float)) else str(lvl_raw).strip()
            
            name1 = str(r.get('명칭 / 구성품', '')).strip()
            name2 = str(r.get('생산LOT', '')).strip()
            final_name = name2 if name2 and len(name2) > len(name1) and name2.lower() != 'nan' else name1
            
            formula = str(r.get('수식 (Formula)', '')).strip()
            unit = str(r.get('단위', '')).strip()
            parent = str(r.get('상위 LOT 연결', '')).strip()
            
            if not final_name or final_name.lower() in ['nan', '']: continue
            
            calculated_val = evaluate_formula(formula, target_qty)
            
            item_dict = {
                "상위Lot": parent if parent.lower() != 'nan' else '',
                "명칭 / 구성품": final_name,
                "생산Lot": name2 if name2.lower() != 'nan' else '',
                "계산된_소요량": round(calculated_val, 3),
                "단위": unit if unit.lower() != 'nan' else '',
                "레벨": lvl,
            }
            
            if lvl == '1': result["level1"].append(item_dict)
            elif lvl == '2': result["level2"].append(item_dict)
            elif lvl == '3': result["level3"].append(item_dict)
            
        return jsonify(result)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=9001, debug=True)
