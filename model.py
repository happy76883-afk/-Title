from io import BytesIO
import json
import openpyxl

STATE = {
    "sheets": [],
    "parsed": {},
}

def _clean_str(val):
    if val is None:
        return ""
    return str(val).strip()

def _detect_unit(ws):
    for r in range(1, min(10, ws.max_row + 1)):
        for c in range(1, min(10, ws.max_column + 1)):
            s = _clean_str(ws.cell(r, c).value)
            if "단위" in s or "백만원" in s or "억원" in s or "명" in s:
                return s
    return "원자료 표기 기준"

def _parse_sheet(ws):
    max_r = ws.max_row
    max_c = ws.max_column
    unit = _detect_unit(ws)

    # 표의 헤더 행/열 시작점 탐색: 문자열이 밀집한 상단 행을 열 헤더로 사용
    header_row = 1
    best_non_empty = -1
    for r in range(1, min(15, max_r + 1)):
        cnt = sum(1 for c in range(1, max_c + 1) if _clean_str(ws.cell(r, c).value) != "")
        if cnt > best_non_empty:
            best_non_empty = cnt
            header_row = r

    # 왼쪽 몇 열이 코드/부문명인지 판별 (기본 1~2열을 행 식별자로 취급)
    data_start_col = 3 if max_c >= 4 else 2

    cols = []
    for c in range(data_start_col, max_c + 1):
        raw_h = _clean_str(ws.cell(header_row, c).value)
        col_code = f"C{c - data_start_col + 1:03d}"
        col_name = raw_h if raw_h else f"열 {c}"
        cols.append({"col_idx": c, "code": col_code, "name": col_name})

    rows = []
    for r in range(header_row + 1, max_r + 1):
        v1 = _clean_str(ws.cell(r, 1).value)
        v2 = _clean_str(ws.cell(r, 2).value) if max_c >= 2 else ""
        if not v1 and not v2:
            continue
        code = v1 if v1 else f"R{r:03d}"
        name = v2 if v2 else v1
        cells = []
        for col in cols:
            raw_val = ws.cell(r, col["col_idx"]).value
            if isinstance(raw_val, (int, float)):
                val = raw_val
            else:
                s = _clean_str(raw_val)
                if s == "":
                    val = None
                else:
                    try:
                        val = float(s.replace(",", ""))
                    except ValueError:
                        val = s
            cells.append({
                "code": col["code"],
                "name": col["name"],
                "value": val,
            })
        rows.append({
            "code": code,
            "name": name,
            "cells": cells,
        })

    return {
        "unit": unit,
        "header_row": header_row,
        "rows": rows,
    }

def initialize(raw_bytes):
    wb = openpyxl.load_workbook(BytesIO(bytes(raw_bytes)), data_only=True, read_only=True)
    sheet_names = list(wb.sheetnames)
    parsed = {}
    summary = {}
    for name in sheet_names:
        ws = wb[name]
        info = _parse_sheet(ws)
        parsed[name] = info
        summary[name] = {
            "row_count": len(info["rows"]),
            "col_count": len(info["rows"][0]["cells"]) if info["rows"] else 0,
            "unit": info["unit"],
        }
    wb.close()

    STATE["sheets"] = sheet_names
    STATE["parsed"] = parsed

    return json.dumps({
        "sheets": sheet_names,
        "summary": summary,
    }, ensure_ascii=False)

def explore(sheet_name: str, row_code: str = "") -> str:
    if sheet_name not in STATE["parsed"]:
        raise ValueError(f"시트를 찾을 수 없습니다: {sheet_name}")

    info = STATE["parsed"][sheet_name]
    rows = info["rows"]
    if not rows:
        return json.dumps({
            "unit": info["unit"],
            "selected": "",
            "name": "데이터 없음",
            "rows": [],
            "cells": [],
        }, ensure_ascii=False)

    selected_row = rows[0]
    if row_code:
        for r in rows:
            if r["code"] == row_code:
                selected_row = r
                break

    row_choices = [{"code": r["code"], "name": r["name"]} for r in rows]
    return json.dumps({
        "unit": info["unit"],
        "selected": selected_row["code"],
        "name": selected_row["name"],
        "rows": row_choices,
        "cells": selected_row["cells"],
    }, ensure_ascii=False)
