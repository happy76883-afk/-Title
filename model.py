import io
import json
import pandas as pd
import numpy as np

EXCEL_BYTES = None

def load_basic_workbook(file_bytes) -> str:
    global EXCEL_BYTES
    EXCEL_BYTES = bytes(file_bytes)
    xl = pd.ExcelFile(io.BytesIO(EXCEL_BYTES))
    return json.dumps({"sheets": xl.sheet_names}, ensure_ascii=False)

def _parse_sheet(sheet_name: str):
    df = pd.read_excel(io.BytesIO(EXCEL_BYTES), sheet_name=sheet_name, header=None)
    rows = df.values.tolist()

    header_idx = 4
    for idx in range(min(10, len(rows))):
        r = rows[idx]
        if r and len(r) > 0 and pd.notna(r[0]) and "상품" in str(r[0]):
            header_idx = idx
            break

    code_row = rows[header_idx] if len(rows) > header_idx else []
    name_row = rows[header_idx + 1] if len(rows) > header_idx + 1 else []

    cols = []
    max_len = max(len(code_row), len(name_row))
    for c_idx in range(2, max_len):
        c_code = str(code_row[c_idx]).strip() if c_idx < len(code_row) and pd.notna(code_row[c_idx]) else ""
        c_name = str(name_row[c_idx]).strip() if c_idx < len(name_row) and pd.notna(name_row[c_idx]) else ""
        if c_code or c_name:
            cols.append({"col_idx": c_idx, "code": c_code or f"C{c_idx}", "name": c_name or c_code})

    parsed_rows = []
    for r in rows[header_idx + 2:]:
        if not r:
            continue
        r_code = str(r[0]).strip() if len(r) > 0 and pd.notna(r[0]) else ""
        r_name = str(r[1]).strip() if len(r) > 1 and pd.notna(r[1]) else ""
        if not r_code and not r_name:
            continue
        vals = []
        for col in cols:
            idx = col["col_idx"]
            v = r[idx] if idx < len(r) and pd.notna(r[idx]) else None
            if isinstance(v, (np.integer, np.floating)):
                v = float(v) if isinstance(v, np.floating) else int(v)
            vals.append(v)
        parsed_rows.append({
            "row_idx": len(parsed_rows),
            "code": r_code,
            "name": r_name or r_code,
            "values": vals,
        })

    return {"cols": cols, "rows": parsed_rows}

def get_sheet(sheet_name: str) -> str:
    return json.dumps(_parse_sheet(sheet_name), ensure_ascii=False, default=str)

def get_industries() -> str:
    parsed = _parse_sheet("생산유발계수")
    inds = [
        {"code": r["code"], "name": r["name"]}
        for r in parsed["rows"]
        if r["code"] and not any(k in r["name"] for k in ("계", "합계", "총투입", "부가가치"))
    ]
    return json.dumps(inds, ensure_ascii=False)

def _matrix_from_sheet(sheet_name: str, inds: list[dict]) -> np.ndarray:
    parsed = _parse_sheet(sheet_name)
    row_map = {r["code"]: r for r in parsed["rows"]}
    col_map = {c["code"]: idx for idx, c in enumerate(parsed["cols"])}
    n = len(inds)
    M = np.zeros((n, n), dtype=float)
    for i, ri in enumerate(inds):
        r = row_map.get(ri["code"])
        if not r:
            continue
        for j, cj in enumerate(inds):
            ci = col_map.get(cj["code"])
            if ci is not None and ci < len(r["values"]) and r["values"][ci] is not None:
                try:
                    M[i, j] = float(r["values"][ci])
                except (ValueError, TypeError):
                    pass
    return M

def simulate(shocks_json: str) -> str:
    shocks = json.loads(shocks_json)
    inds = json.loads(get_industries())
    n = len(inds)
    idx_map = {d["code"]: i for i, d in enumerate(inds)}

    dy = np.zeros(n, dtype=float)
    for s in shocks:
        if s["code"] in idx_map:
            dy[idx_map[s["code"]]] += float(s["amount"])

    out_eff = float((_matrix_from_sheet("생산유발계수", inds) @ dy).sum())
    va_eff = float((_matrix_from_sheet("부가가치유발계수", inds) @ dy).sum())
    imp_eff = float((_matrix_from_sheet("수입유발계수", inds) @ dy).sum())

    return json.dumps({
        "total_shock": round(float(dy.sum()), 2),
        "total_output": round(out_eff, 2),
        "total_va": round(va_eff, 2),
        "total_import": round(imp_eff, 2),
    }, ensure_ascii=False)
