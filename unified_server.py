import calendar
import sqlite3
import pandas as pd
import re
from datetime import date, timedelta
from flask import Flask, jsonify, send_from_directory, request, send_file
import os
import io
import openpyxl
from openpyxl.cell.cell import MergedCell
from openpyxl.utils import get_column_letter
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
        cursor.execute(
            'SELECT "포장시 요구량" FROM level1 WHERE "상위Lot" = ? AND UPPER(TRIM("코드번호")) LIKE "EMA015%" LIMIT 1',
            (lot_no,),
        )
        ema_row = cursor.fetchone()
        conn.close()
        try:
            pack_qty = float(str(ema_row[0]).replace(',', '')) if ema_row and ema_row[0] not in (None, '') else None
        except (ValueError, TypeError):
            pack_qty = None
        kit_qty = l0.get('생산 수량(kit)') or 0
        try:
            kit_qty = float(str(kit_qty).replace(',', ''))
        except (ValueError, TypeError):
            kit_qty = 0.0
        total_qty = pack_qty if pack_qty is not None else kit_qty
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
        try:
            ema015 = next(
                (x for x in l1_items if str(x.get('코드번호') or '').strip().upper().startswith('EMA015')),
                None,
            )
            pq = ema015.get('포장시 요구량') if ema015 else None
            pack_qty = float(str(pq).replace(',', '')) if pq not in (None, '') else None
        except (ValueError, TypeError):
            pack_qty = None
        kq = l0.get('생산 수량(kit)') or 0
        try:
            kq = float(str(kq).replace(',', ''))
        except (ValueError, TypeError):
            kq = 0
        total_qty = pack_qty if pack_qty is not None else kq

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
            
        ws['L33'] = ''
        ws['S33'] = l0['제조일자']
        ws['X33'] = ''
        ws['AI33'] = total_qty
        ws['N7'] = total_qty

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

# --- 반제품 관리 API (25BCE01-반제품 관리.xlsx) ---
# B2·M7(item_master unit)·H6~H9·X6·X7, B12:AD30 비움
SEMI_PRODUCT_MGMT_TEMPLATE = '25BCE01-반제품 관리.xlsx'
FINISHED_PRODUCT_MGMT_TEMPLATE = '25BCE01-완제품 관리.xlsx'


def _semi_product_name_line_for_b2(cursor, a7_구성품명칭, code_i7):
    """B2 첫 줄: Level1 구성품 명칭 → 없으면 item_master.description → 없으면 코드."""
    n = str(a7_구성품명칭 or '').strip()
    if n:
        return n
    ku = str(code_i7 or '').strip().upper()
    if not ku:
        return ''
    cursor.execute(
        'SELECT description FROM item_master WHERE UPPER(TRIM(code_no)) = ? LIMIT 1',
        (ku,),
    )
    r = cursor.fetchone()
    if r and r[0] is not None and str(r[0]).strip():
        return str(r[0]).strip()
    return str(code_i7 or '').strip()


# B2 첫 줄: 약어(코드)별 고정 표기 (I7 정규화 키와 일치)
SEMI_B2_DISPLAY_NAME_BY_CODE_KEY = {
    'PB(01)': 'PBSA Buffer',
    'CB(01)': 'Coating Buffer',
    'WB(01)': 'Washing Buffer',
}


def _semi_b2_strip_plate_b_parenthetical(text):
    """문자열에서 (Plate-B) 괄호 구문 제거."""
    s = str(text or '')
    s = re.sub(r'\s*\(\s*Plate-B\s*\)\s*', ' ', s, flags=re.IGNORECASE)
    return re.sub(r'\s{2,}', ' ', s).strip()


def _semi_b2_first_line_display(cursor, a7_구성품명칭, code_i7):
    """B2 표시용 첫 줄: PB/CB/WB 고정명 → PL 계열은 (Plate-B) 제거 → 그 외는 기본 규칙."""
    base = _semi_product_name_line_for_b2(cursor, a7_구성품명칭, code_i7)
    ck = _instruction_code_key(code_i7)
    if ck in SEMI_B2_DISPLAY_NAME_BY_CODE_KEY:
        return SEMI_B2_DISPLAY_NAME_BY_CODE_KEY[ck]
    if ck.startswith('PL'):
        adj = _semi_b2_strip_plate_b_parenthetical(base)
        return adj if adj else base
    return base


def _semi_mgmt_b2_cell_text(cursor, a7_구성품명칭, code_i7):
    """「반제품 명\\n반제품 관리대장」형식."""
    first = _semi_b2_first_line_display(cursor, a7_구성품명칭, code_i7)
    if first:
        return f'{first}\n반제품 관리대장'
    return '반제품 관리대장'


def _item_master_unit(cursor, code_no):
    """item_master.unit (code_no 일치)."""
    ku = str(code_no or '').strip().upper()
    if not ku:
        return ''
    cursor.execute(
        'SELECT unit FROM item_master WHERE UPPER(TRIM(code_no)) = ? LIMIT 1',
        (ku,),
    )
    r = cursor.fetchone()
    if r and r[0] is not None and str(r[0]).strip():
        return str(r[0]).strip()
    return ''


def _semi_mgmt_clear_range_b12_ad30(ws):
    """템플릿 B12:AD30 내용 삭제."""
    for row in range(12, 31):
        for col in range(2, 31):
            _write_cell_safe(ws, f'{get_column_letter(col)}{row}', None)


def _semi_mgmt_h9_fridge(division_or_code):
    """H9: 약어/코드에 따른 냉장고 자산번호."""
    d = re.sub(r'\s+', '', str(division_or_code or '').strip().upper())
    if d.startswith('PB') or d.startswith('CB') or d.startswith('WB'):
        return '냉장고(ESH-GP-088)'
    for prefix in ('CR', 'PC', 'NC', 'DA', 'PL', 'RD', 'WS', 'TM', 'SS'):
        if d.startswith(prefix):
            return '냉장고(ESH-GP-089)'
    return ''


def _open_semi_mgmt_workbook():
    """반제품 템플릿 우선, 없으면 완제품 관리(동일 입력 셀), 둘 다 없으면 빈 워크북."""
    semi_path = os.path.join(ROOT_DIR, SEMI_PRODUCT_MGMT_TEMPLATE)
    if os.path.exists(semi_path):
        return openpyxl.load_workbook(semi_path)
    fp_path = os.path.join(ROOT_DIR, FINISHED_PRODUCT_MGMT_TEMPLATE)
    if os.path.exists(fp_path):
        return openpyxl.load_workbook(fp_path)
    return openpyxl.Workbook()


def _write_cell_safe(ws, coord, value):
    """병합 셀(MergedCell)이면 병합 범위의 좌상단에만 값을 씁니다."""
    cell = ws[coord]
    if isinstance(cell, MergedCell):
        for mr in ws.merged_cells.ranges:
            if coord in mr:
                ws.cell(row=mr.min_row, column=mr.min_col, value=value)
                return
        return
    ws[coord] = value


def _split_lot_tokens(s):
    if s is None or str(s).strip() == '':
        return []
    return [x.strip() for x in re.split(r'[\n,;]+', str(s)) if x.strip()]


def _build_semi_product_management_preview(cursor, parent_lot, semi_lot_raw, division):
    """DB에서 반제품 관리용 필드 조회. 성공 시 (dict, None), 실패 시 (None, error_msg)."""
    parent_lot = (parent_lot or '').strip()
    if not parent_lot:
        return None, 'parent_lot이 필요합니다.'
    semi_lot_raw = semi_lot_raw or ''
    division = (division or '').strip()
    div_u = division.upper()

    cursor.execute('SELECT * FROM level0 WHERE "LOT NO." = ?', (parent_lot,))
    l0r = cursor.fetchone()
    if not l0r:
        return None, '상위 Lot을 찾을 수 없습니다.'
    l0 = dict(l0r)

    cursor.execute('SELECT * FROM level1 WHERE "상위Lot" = ?', (parent_lot,))
    l1_all = [dict(r) for r in cursor.fetchall()]
    cursor.execute('SELECT * FROM instruction_summary WHERE "상위Lot" = ?', (parent_lot,))
    summ_all = [dict(r) for r in cursor.fetchall()]

    semi_tokens = _split_lot_tokens(semi_lot_raw)
    if not semi_tokens and not div_u:
        return None, 'semi_lot 또는 division이 필요합니다.'

    l1 = None

    def lot_matches_row(lot_str):
        if not semi_tokens:
            return False
        lot_str = str(lot_str or '').strip()
        if not lot_str:
            return False
        parts = _split_lot_tokens(lot_str)
        for t in semi_tokens:
            if t == lot_str or t in parts:
                return True
            if t in lot_str or lot_str in t:
                return True
        return False

    for r in l1_all:
        if lot_matches_row(r.get('Lot No.')):
            l1 = r
            break

    if l1 is None and div_u:
        for r in l1_all:
            code = str(r.get('코드번호') or '').strip().upper()
            if not code:
                continue
            if code == div_u or div_u in code or code in div_u:
                l1 = r
                break

    summ = None
    for s in summ_all:
        lot_dot = str(s.get('Lot. No.') or '').strip()
        abbrev = str(s.get('약어') or '').strip().upper()
        if semi_tokens:
            parts = _split_lot_tokens(lot_dot)
            for t in semi_tokens:
                if t == lot_dot or t in parts or (lot_dot and (t in lot_dot or lot_dot in t)):
                    summ = s
                    break
            if summ:
                break
        if div_u and abbrev:
            ak = _instruction_code_key(abbrev)
            dk = _instruction_code_key(div_u)
            if ak == dk or (dk and dk in ak) or (ak and ak in dk):
                summ = s
                break

    if l1 is None and summ is None:
        return None, '반제품에 해당하는 Level1 또는 지시 요약을 찾을 수 없습니다.'

    def _to_float(v):
        try:
            return float(str(v or 0).replace(',', ''))
        except (ValueError, TypeError):
            return 0.0

    i9 = _to_float((l1 or {}).get('포장시 요구량')) if l1 else 0.0
    if i9 == 0 and l1:
        i9 = _to_float(l1.get('할당수량'))

    n7 = str(semi_lot_raw).strip() if str(semi_lot_raw).strip() else ''
    if not n7 and l1:
        n7 = str(l1.get('Lot No.') or '')
    if not n7 and summ:
        n7 = str(summ.get('Lot. No.') or '')

    a7 = str((l1 or {}).get('구성품 명칭') or '')
    i7 = str((l1 or {}).get('코드번호') or '')
    if summ and not a7:
        a7 = str(summ.get('약어') or '')
    if summ and not i7:
        i7 = str(summ.get('약어') or '')
    t7 = ''
    if l1 and l1.get('제조일자'):
        t7 = str(l1.get('제조일자'))
    elif summ and summ.get('제조일자'):
        t7 = str(summ.get('제조일자'))
    else:
        t7 = str(l0.get('제조일자') or '')
    a9 = str((l1 or {}).get('유효기간') or '')

    instr_no = str((summ or {}).get('제조지침서 No.') or '')
    div_out = division
    if summ and summ.get('약어'):
        div_out = str(summ.get('약어'))
    elif not div_out and l1:
        div_out = str(l1.get('코드번호') or '')

    b2 = _semi_mgmt_b2_cell_text(cursor, a7, i7)
    # H6·X6·H7: instruction_summary만 (L1·N7·I9 폴백 없음)
    h6 = x6 = h7 = ''
    if summ:
        h6 = str(summ.get('Lot. No.') or '').strip()
        x6 = _fmt_date_yyyy_mm_dd(summ.get('제조일자'))
        h7 = str(summ.get('생산량') or '').strip()
    mfg_for_x7 = None
    if summ and summ.get('제조일자'):
        mfg_for_x7 = summ.get('제조일자')
    elif l1 and l1.get('제조일자'):
        mfg_for_x7 = l1.get('제조일자')
    else:
        mfg_for_x7 = l0.get('제조일자')
    x7 = _expiry_plus_13_months_minus_1_day(mfg_for_x7)
    if not x7:
        x7 = str(a9 or '').strip()
    h8 = '2 ~ 8℃'
    h9 = _semi_mgmt_h9_fridge(div_out or i7)
    m7 = _item_master_unit(cursor, i7)

    preview = {
        'A7': a7,
        'I7': i7,
        'N7': n7,
        'T7': t7,
        'A9': a9,
        'I9': i9,
        'B2': b2,
        'M7': m7,
        'H6': h6,
        'X6': x6,
        'H7': h7,
        'X7': x7,
        'H8': h8,
        'H9': h9,
        'division': div_out,
        'instructionNo': instr_no,
        'lotNo': n7,
        'productName': a7,
        'productCode': i7,
        'mfgDate': t7,
        'expiry': x7,
        'qty': i9,
    }
    return preview, None


@app.route('/api/semi_product_management_preview')
def semi_product_management_preview():
    parent_lot = request.args.get('parent_lot', '')
    semi_lot = request.args.get('semi_lot', '')
    division = request.args.get('division', '')
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        preview, err = _build_semi_product_management_preview(cursor, parent_lot, semi_lot, division)
        conn.close()
        if err:
            return jsonify({'error': err}), 404
        return jsonify(preview)
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/semi_product_management_download')
def semi_product_management_download():
    parent_lot = request.args.get('parent_lot', '')
    semi_lot = request.args.get('semi_lot', '')
    division = request.args.get('division', '')
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        preview, err = _build_semi_product_management_preview(cursor, parent_lot, semi_lot, division)
        conn.close()
        if err or not preview:
            return jsonify({'error': err or 'No data'}), 404

        wb = _open_semi_mgmt_workbook()
        ws = wb.active
        _semi_mgmt_clear_range_b12_ad30(ws)
        _write_cell_safe(ws, 'B2', preview.get('B2') or '')
        _write_cell_safe(ws, 'M7', preview.get('M7') or '')
        _write_cell_safe(ws, 'H6', preview.get('H6') or '')
        _write_cell_safe(ws, 'X6', preview.get('X6') or '')
        _write_cell_safe(ws, 'H7', preview.get('H7') or '')
        _write_cell_safe(ws, 'X7', preview.get('X7') or '')
        _write_cell_safe(ws, 'H8', preview.get('H8') or '')
        _write_cell_safe(ws, 'H9', preview.get('H9') or '')

        safe_name = re.sub(r'[^\w\-]+', '_', str(preview.get('N7') or 'semi'))[:80]
        tmp_fd, tmp_name = tempfile.mkstemp(suffix='.xlsx')
        os.close(tmp_fd)
        wb.save(tmp_name)
        wb.close()
        return send_file(
            tmp_name,
            as_attachment=True,
            download_name=f'Semi_Product_Management_{safe_name}.xlsx',
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        )
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# --- 마스터 조회 (제조지시 실행 탭) ---
def _row_get(row, *keys):
    if not row:
        return None
    for k in keys:
        if k in row and row[k] is not None and str(row[k]).strip() != '':
            return row[k]
    return None


@app.route('/api/item_master/<path:code_no>')
def get_item_master(code_no):
    code = urllib.parse.unquote(code_no or '').strip()
    if not code:
        return jsonify({})
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute(
            'SELECT code_no, description, detailed_description, version FROM item_master WHERE code_no = ? LIMIT 1',
            (code,),
        )
        r = cursor.fetchone()
        conn.close()
        if not r:
            return jsonify({})
        d = dict(r)
        return jsonify({
            'description': d.get('description') or '',
            'detailed_description': d.get('detailed_description') or '',
            'version': d.get('version') or '',
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/doc_master/<path:code_no>')
def get_doc_master(code_no):
    code = urllib.parse.unquote(code_no or '').strip()
    if not code:
        return jsonify([])
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute('SELECT * FROM instruction_doc_master WHERE code_no = ? ORDER BY id', (code,))
        rows = [dict(r) for r in cursor.fetchall()]
        conn.close()
        return jsonify(rows)
    except Exception as e:
        return jsonify([]), 500


def _fmt_date_yyyy_mm_dd(val):
    """제조일자를 YYYY-MM-DD 문자열로 통일 (저장 시). YYMMDD(6자리) → 20YY-MM-DD."""
    if val is None or val == '':
        return ''
    s = str(val).strip()
    if re.match(r'^\d{4}-\d{2}-\d{2}', s):
        return s[:10]
    digits = re.sub(r'\D', '', s)
    if len(digits) == 6:
        return f'20{digits[:2]}-{digits[2:4]}-{digits[4:6]}'
    if len(digits) >= 8:
        return f'{digits[:4]}-{digits[4:6]}-{digits[6:8]}'
    return s[:10] if len(s) >= 10 else s


def _parse_mfg_date_to_date(val):
    """app.js parseDateInput 대응: 숫자 8자리(YYYYMMDD), 6자리(YYMMDD), YYYY-MM-DD."""
    if val is None or str(val).strip() == '':
        return None
    s = str(val).strip()
    digits = re.sub(r'\D', '', s)
    if len(digits) >= 8:
        try:
            return date(int(digits[:4]), int(digits[4:6]), int(digits[6:8]))
        except ValueError:
            return None
    if len(digits) == 6:
        try:
            return date(2000 + int(digits[:2]), int(digits[2:4]), int(digits[4:6]))
        except ValueError:
            return None
    m = re.match(r'^(\d{4})-(\d{2})-(\d{2})', s)
    if m:
        try:
            return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            return None
    return None


def _expiry_plus_13_months_minus_1_day(mfg_val):
    """
    반제품 유효기간: 제조일 + 13개월(말일 보정) − 1일.
    app.js calcExpiryDate(setMonth+13), setDate(0) 보정, setDate(-1) 와 동치.
    """
    mfg = _parse_mfg_date_to_date(mfg_val)
    if not mfg:
        return ''
    y, mo, d = mfg.year, mfg.month, mfg.day
    total_m = y * 12 + (mo - 1) + 13
    ny = total_m // 12
    nm = total_m % 12 + 1
    last = calendar.monthrange(ny, nm)[1]
    nd = min(d, last)
    d2 = date(ny, nm, nd)
    d3 = d2 - timedelta(days=1)
    return d3.strftime('%Y-%m-%d')


def _lot_no_equiv_set(lot_str):
    """
    Lot 문자열 동치(날짜 접두만 다른 경우): 111127-01PB-01R3 ↔ 2011-11-27-01PB-01R3
    """
    s = str(lot_str or '').strip()
    if not s:
        return frozenset()
    out = {s}
    m = re.match(r'^(\d{2})(\d{2})(\d{2})(-.+)$', s)
    if m:
        yy, mm, dd, rest = m.groups()
        out.add(f'20{yy}-{mm}-{dd}{rest}')
    m = re.match(r'^(20\d{2})-(\d{2})-(\d{2})(-.+)$', s)
    if m:
        y, mm, dd, rest = m.group(1), m.group(2), m.group(3), m.group(4)
        out.add(f'{y[2:]}{mm}{dd}{rest}')
    return frozenset(out)


def _lot_refs_equal(a, b):
    """instruction_summary Lot. No. 와 level3 상위Lot이 같은 반제품을 가리키는지."""
    a, b = str(a or '').strip(), str(b or '').strip()
    if not a or not b:
        return False
    if a == b:
        return True
    return bool(_lot_no_equiv_set(a) & _lot_no_equiv_set(b))


def _gubun_from_item_master(cursor, code_no, memo):
    """item_master의 category로 구분(완제품·반제품·원재료·소모품 등) 조회. IFU 계열은 소모품."""
    k = str(code_no or '').strip()
    if not k:
        return ''
    ku = k.upper()
    if ku == 'IFU' or ku.startswith('IFU'):
        return '소모품'
    if ku in memo:
        return memo[ku]
    cursor.execute(
        'SELECT category FROM item_master WHERE UPPER(TRIM(code_no)) = ? LIMIT 1',
        (ku,),
    )
    row = cursor.fetchone()
    cat = ''
    if row and row[0] is not None:
        cat = str(row[0]).strip()
    memo[ku] = cat
    return cat


def _manufacturer_from_item_master(cursor, code_no, memo):
    """item_master의 manufacturer를 code_no로 조회."""
    k = str(code_no or '').strip()
    if not k:
        return ''
    ku = k.upper()
    if ku in memo:
        return memo[ku]
    cursor.execute(
        'SELECT manufacturer FROM item_master WHERE UPPER(TRIM(code_no)) = ? LIMIT 1',
        (ku,),
    )
    row = cursor.fetchone()
    mfr = ''
    if row and row[0] is not None:
        mfr = str(row[0]).strip()
    memo[ku] = mfr
    return mfr


def _instruction_code_key(s):
    """약어·코드번호 비교용(공백 제거, 대문자)."""
    return re.sub(r'\s+', '', str(s or '').strip().upper())


def _l1_row_for_instruction_summary(l1_raw, parent_lot, division, semi_lot):
    """
    Level1에서 상위Lot=생산Lot, 코드번호=약어(division),
    할당 Lot이 반제품 Lot(calcLot)과 맞는 첫 행(없으면 None).
    """
    parent_lot = (parent_lot or '').strip()
    div_key = _instruction_code_key(division)
    semi_s = str(semi_lot or '').strip()
    semi_tokens = _split_lot_tokens(semi_lot)
    if not parent_lot or not div_key or not semi_s:
        return None
    for r in l1_raw:
        pl = str(_row_get(r, '상위Lot', '상위 Lot', '상위 LOT') or '').strip()
        if pl != parent_lot:
            continue
        code = str(_row_get(r, '코드번호', 'Code No.', 'Code') or '').strip().upper()
        if _instruction_code_key(code) != div_key:
            continue
        lot_cell = str(_row_get(r, 'Lot No.', '할당 Lot', '할당Lot') or '').strip()
        lot_set = set(_split_lot_tokens(lot_cell))
        if lot_cell:
            lot_set.add(lot_cell)
        ok = False
        for t in semi_tokens:
            if not t:
                continue
            if t == lot_cell or t in lot_set:
                ok = True
                break
            if lot_cell and (t in lot_cell or lot_cell in t):
                ok = True
                break
        if not ok:
            continue
        return r
    return None


def _l1_packaging_qty_for_instruction_summary(l1_raw, parent_lot, division, semi_lot):
    """
    instruction_summary 생산량: Level1에서 상위Lot=생산Lot, 코드번호=약어(division),
    할당 Lot이 반제품 Lot(calcLot)과 맞는 행의 포장시 요구량(또는 할당수량 등).
    """
    r = _l1_row_for_instruction_summary(l1_raw, parent_lot, division, semi_lot)
    if not r:
        return ''
    q = _row_get(r, '포장시 요구량', '할당수량', '필요 수량', '제조량')
    if q is None or str(q).strip() == '':
        return ''
    return str(q).strip()


def _l3_cam006_alloc_for_instruction_lot(l3_raw, instruction_lot):
    """
    업로드 CSV level3: PB/CB/WB 반제품 Lot(calcLot)과 상위Lot이 동치(날짜 접두 포함)인 행 중
    코드번호 CAM006(또는 자료에 CMA006으로 적힌 경우)의 할당수량만 사용(첫 건, 합산 없음).
    """
    semi = str(instruction_lot or '').strip()
    if not semi:
        return ''
    cam_keys = frozenset({'CAM006', 'CMA006'})
    for r in l3_raw:
        code = str(_row_get(r, '코드번호', 'Code No.', 'Code') or '').strip().upper()
        if _instruction_code_key(code) not in cam_keys:
            continue
        pl3 = str(_row_get(r, '상위Lot', '상위 Lot', '상위 LOT') or '').strip()
        if not _lot_refs_equal(semi, pl3):
            continue
        q = _row_get(r, '할당수량')
        if q is not None and str(q).strip() != '':
            return str(q).strip()
    return ''


def _l1_packaging_qty_for_cr(l1_raw, parent_lot):
    """
    Level1에서 상위Lot=생산 Lot이고 코드번호가 CR로 시작하는 첫 행의 포장시 요구량.
    PI instruction_summary 생산량을 CR과 동일하게 맞출 때 사용(calcLot 없음).
    """
    parent_lot = (parent_lot or '').strip()
    if not parent_lot:
        return ''
    for r in l1_raw:
        pl = str(_row_get(r, '상위Lot', '상위 Lot', '상위 LOT') or '').strip()
        if pl != parent_lot:
            continue
        code = str(_row_get(r, '코드번호', 'Code No.', 'Code') or '').strip().upper()
        if not code.startswith('CR'):
            continue
        q = _row_get(r, '포장시 요구량', '할당수량', '필요 수량', '제조량')
        if q is None or str(q).strip() == '':
            continue
        return str(q).strip()
    return ''


@app.route('/api/save_instruction', methods=['POST'])
def save_instruction():
    try:
        data = request.get_json(force=True, silent=True) or {}
        l0_src = data.get('level0') or {}
        lot_no = (l0_src.get('lotNo') or l0_src.get('LOT NO.') or '').strip()
        if not lot_no:
            return jsonify({'status': 'error', 'error': 'LOT No.가 없습니다.'}), 400

        conn = get_db_connection()
        cursor = conn.cursor()
        cat_memo = {}
        mfr_memo = {}

        def gubun_for_code(code_no, fallback=''):
            g = _gubun_from_item_master(cursor, code_no, cat_memo)
            return g if g else (fallback or '')

        def manufacturer_for_code(code_no, fallback=''):
            m = _manufacturer_from_item_master(cursor, code_no, mfr_memo)
            return m if m else (fallback or '')

        l0_code = l0_src.get('modelName') or l0_src.get('제품코드') or ''
        l0 = {
            'Level': 0,
            '제품코드': l0_code,
            '구분': gubun_for_code(l0_code, l0_src.get('구분') or ''),
            '제품명': l0_src.get('productName') or l0_src.get('제품명') or '',
            'LOT NO.': lot_no,
            '제품버전': l0_src.get('version') or l0_src.get('제품버전') or '',
            '제조일자': _fmt_date_yyyy_mm_dd(l0_src.get('mfgDate') or l0_src.get('제조일자') or ''),
            '유효기간': l0_src.get('expiryDate') or l0_src.get('유효기간') or '',
            '생산 수량(kit)': l0_src.get('targetQty') or l0_src.get('생산 수량(kit)') or '',
            '생산의뢰서 번호': l0_src.get('생산의뢰서 번호') or '',
            '의뢰팀': l0_src.get('requestTeam') or l0_src.get('의뢰팀') or '',
            '생산목적': l0_src.get('purpose') or l0_src.get('생산목적') or '',
            '작업자': l0_src.get('작업자') or '',
            '검사자': l0_src.get('검사자') or '',
            '검사일': l0_src.get('검사일') or '',
            '완제품검사 문서번호': l0_src.get('완제품검사 문서번호') or '',
            '제품정보': l0_src.get('productInfo') or l0_src.get('제품정보') or '',
        }

        def row_mfg_date(row):
            raw = _row_get(row, '제조일자', '제조 일자', 'MfgDate', 'mfgDate', 'Manufacturing Date') or ''
            fmt = _fmt_date_yyyy_mm_dd(raw)
            return fmt if fmt else l0['제조일자']

        def norm_l1(row):
            code = _row_get(row, '코드번호', 'Code No.', 'Code') or ''
            return {
                'Level': int(_row_get(row, 'Level', 'level') or 1),
                '상위Lot': _row_get(row, '상위Lot', '상위 Lot', '상위 LOT') or '',
                '코드번호': code,
                '구분': gubun_for_code(code, _row_get(row, '구분') or ''),
                '구성품 명칭': _row_get(row, '구성품 명칭', '명칭 / 구성품', '명칭/구성품') or '',
                'Lot No.': _row_get(row, 'Lot No.', '할당 Lot', '할당Lot') or '',
                '제조일자': row_mfg_date(row),
                '유효기간': _row_get(row, '유효기간') or '',
                '포장 기준량': _row_get(row, '포장 기준량') or '',
                '포장시 요구량': _row_get(row, '포장시 요구량', '할당수량', '필요 수량', '제조량') or '',
                '단위': _row_get(row, '단위') or '',
            }

        def norm_l2(row):
            code = _row_get(row, '코드번호', 'Code No.', 'Code') or ''
            return {
                'Level': int(_row_get(row, 'Level', 'level') or 2),
                '상위Lot': _row_get(row, '상위Lot', '상위 Lot', '상위 LOT') or '',
                '코드번호': code,
                '구분': gubun_for_code(code, _row_get(row, '구분') or ''),
                '원재료명': _row_get(row, '원재료명', '명칭 / 구성품', '구성품 명칭') or '',
                '제조사': manufacturer_for_code(code, _row_get(row, '제조사', 'Maker', 'maker', 'Manufacturer') or ''),
                'Lot No.': _row_get(row, 'Lot No.', '할당 Lot', '할당Lot') or '',
                '제조일자': row_mfg_date(row),
                '유효기간': _row_get(row, '유효기간') or '',
                '제조량': _row_get(row, '제조량', '할당수량', '포장시 요구량', '필요 수량') or '',
                '단위': _row_get(row, '단위') or '',
            }

        l1_raw = data.get('level1') or []
        l2_raw = data.get('level2') or []
        l3_raw = data.get('level3') or []
        l1_rows = [norm_l1(r) for r in l1_raw]
        l2_rows = [norm_l2(r) for r in l2_raw]
        l3_rows = [norm_l2(r) for r in l3_raw]

        summary_in = data.get('instruction_summary') or []
        summary_rows = []
        for item in summary_in:
            div = item.get('division') or item.get('약어') or ''
            calc_raw = item.get('calcLot') or item.get('Lot. No.') or ''
            div_u = (div or '').strip().upper()
            is_pb_cb_wb = (
                div_u.startswith('PB') or div_u.startswith('CB') or div_u.startswith('WB')
            )
            if is_pb_cb_wb:
                prod_qty = _l3_cam006_alloc_for_instruction_lot(l3_raw, calc_raw)
            else:
                prod_qty = ''
                qty_l1 = _l1_packaging_qty_for_instruction_summary(l1_raw, lot_no, div, calc_raw)
                if not qty_l1 and div_u.startswith('PI'):
                    qty_l1 = _l1_packaging_qty_for_cr(l1_raw, lot_no)
                prod_qty = qty_l1 if qty_l1 else (item.get('생산량') or item.get('productionQty') or '')
            summary_rows.append({
                '상위Lot': lot_no,
                '약어': div,
                '제조지침서 No.': item.get('latest_doc_no') or item.get('제조지침서 No.') or '',
                'Lot. No.': str(calc_raw).strip(),
                '생산량': prod_qty,
                '제조일자': _fmt_date_yyyy_mm_dd(
                    item.get('mfgDate') or item.get('제조일자') or l0_src.get('mfgDate') or l0_src.get('제조일자')
                ),
            })

        def insert_table(cursor, table, rows):
            if not rows:
                return
            cols = list(rows[0].keys())
            qcols = ','.join(['"' + c.replace('"', '""') + '"' for c in cols])
            ph = ','.join(['?' for _ in cols])
            sql = f'INSERT INTO {table} ({qcols}) VALUES ({ph})'
            for r in rows:
                cursor.execute(sql, [r[c] for c in cols])

        cursor.execute(
            'DELETE FROM level3 WHERE "상위Lot" IN (SELECT "Lot No." FROM level2 WHERE "상위Lot" IN (SELECT "Lot No." FROM level1 WHERE "상위Lot" = ?))',
            (lot_no,),
        )
        cursor.execute(
            'DELETE FROM level2 WHERE "상위Lot" IN (SELECT "Lot No." FROM level1 WHERE "상위Lot" = ?)',
            (lot_no,),
        )
        cursor.execute('DELETE FROM level1 WHERE "상위Lot" = ?', (lot_no,))
        cursor.execute('DELETE FROM instruction_summary WHERE "상위Lot" = ?', (lot_no,))
        cursor.execute('DELETE FROM level0 WHERE "LOT NO." = ?', (lot_no,))

        insert_table(cursor, 'level0', [l0])
        insert_table(cursor, 'level1', l1_rows)
        insert_table(cursor, 'level2', l2_rows)
        insert_table(cursor, 'level3', l3_rows)
        insert_table(cursor, 'instruction_summary', summary_rows)

        conn.commit()
        conn.close()
        return jsonify({'status': 'success'})
    except Exception as e:
        return jsonify({'status': 'error', 'error': str(e)}), 500


if __name__ == '__main__':
    print("--- Unified BOM System Server Starting ---")
    app.run(host='0.0.0.0', port=9000, debug=True)
