import json, io, re
import numpy as np
import openpyxl

book = None
ad = None
am = None
a_mat = None
L = None
M = None
V = None
v_vec = None
codes = []
names = []
sheets = []
valid = False


def emit(x):
    return json.dumps(x, ensure_ascii=False, allow_nan=False)


def norm_code(val):
    if val is None:
        return ''
    if isinstance(val, float) and val.is_integer():
        return str(int(val))
    return str(val).strip()


def is_sector_code(code):
    if not code:
        return False
    if code.isdigit() and int(code) >= 9000:
        return False
    if re.fullmatch(r'9\d{3}', code):
        return False
    if any(k in code for k in ('합계', '계', '부가가치', '최종수요', '총투입', '총산출')):
        return False
    return bool(re.fullmatch(r'[0-9A-Za-z._-]+', code))


def build_row_index(rows):
    ri = {}
    for i, r in enumerate(rows):
        if r and r[0] is not None:
            k = norm_code(r[0])
            if k and k not in ri:
                ri[k] = i
    return ri


def build_col_index(header_row):
    ci = {}
    for i, c in enumerate(header_row):
        if c is not None:
            k = norm_code(c)
            if k and k not in ci:
                ci[k] = i
    return ci


def initialize(raw):
    global book, codes, names, L, M, V, v_vec, ad, am, a_mat, sheets, valid
    wb = openpyxl.load_workbook(io.BytesIO(bytes(raw)), read_only=True, data_only=True)
    book = {s: list(wb[s].values) for s in wb.sheetnames}
    wb.close()
    sheets = list(book)

    base = book['국산투입계수(Ad)']
    nm = {}
    for r in base[6:]:
        if r and r[0] is not None and len(r) > 1 and r[1] is not None:
            k = norm_code(r[0])
            v = str(r[1]).strip()
            if k and v and k not in nm:
                nm[k] = v

    raw_codes = [norm_code(c) for c in base[4][2:]]
    codes = [c for c in raw_codes if is_sector_code(c) and c in nm]

    if len(codes) < 5 or len(set(codes)) != len(codes):
        sample_rows = [(norm_code(r[0]), str(r[1]).strip() if len(r) > 1 and r[1] is not None else '') for r in base[6:12] if r and r[0] is not None]
        raise ValueError(f'부문 코드를 정상 인식하지 못했습니다. (인식: {len(codes)}, 열후보: {raw_codes[:8]}, 행샘플: {sample_rows})')

    names = [nm[c] for c in codes]

    def matrix(s):
        rows = book[s]
        ri = build_row_index(rows)
        ci = build_col_index(rows[4])
        a = np.array([[rows[ri[c]][ci[k]] for k in codes] for c in codes], dtype=float)
        if not np.isfinite(a).all():
            raise ValueError(s + ': 비수치 또는 결측 데이터')
        return a

    L = matrix('생산유발계수')
    M = matrix('수입유발계수')
    V = matrix('부가가치유발계수')
    ad = matrix('국산투입계수(Ad)')
    am = matrix('수입투입계수(Am)')
    a_mat = matrix('총투입계수(A)')

    rows = book['총투입계수(A)']
    ci = build_col_index(rows[4])
    vr = next(r for r in rows if r and norm_code(r[0]) == '9690')
    v_vec = np.array([vr[ci[c]] for c in codes], dtype=float)

    residuals = {
        '부문수': len(codes),
        '생산': float(np.max(np.abs((np.eye(len(codes)) - ad) @ L - np.eye(len(codes))))),
        '수입': float(np.max(np.abs(am @ L - M))),
        '부가가치': float(np.max(np.abs(v_vec[:, None] * L - V))),
    }
    valid = max(residuals['생산'], residuals['수입'], residuals['부가가치']) < 1e-7
    if not valid:
        raise ValueError('계수 검증 실패: ' + str(residuals))

    return emit({
        'sectors': [{'code': c, 'name': n} for c, n in zip(codes, names)],
        'sheets': sheets,
        'residuals': residuals,
    })


def explore(sheet, rowcode):
    rows = book[sheet]
    choices = [
        {'code': norm_code(r[0]), 'name': str(r[1]).strip()}
        for r in rows[6:]
        if len(r) > 1 and r[0] is not None and r[1] is not None and norm_code(r[0])
    ]
    chosen = next((r for r in rows[6:] if len(r) > 1 and norm_code(r[0]) == rowcode), None)
    if chosen is None:
        chosen = next(r for r in rows[6:] if len(r) > 1 and r[0] is not None and r[1] is not None)

    cells = []
    for i in range(2, len(rows[4])):
        if rows[4][i] is None:
            continue
        val = chosen[i] if i < len(chosen) else None
        if isinstance(val, (float, int)):
            val = float(val)
        elif val is not None:
            val = str(val)
        cells.append({'code': norm_code(rows[4][i]), 'name': str(rows[5][i] or '').strip(), 'value': val})

    return emit({
        'rows': choices,
        'selected': norm_code(chosen[0]),
        'name': str(chosen[1]).strip(),
        'unit': str(rows[3][0] or '계수').replace('係数', '계수 / 비율: 행·열 명칭 확인'),
        'cells': cells,
    })


def simulate(payload):
    if not valid:
        raise ValueError('검증된 데이터가 없습니다.')
    shocks = json.loads(payload)
    f = np.zeros(len(codes), dtype=float)
    for s in shocks:
        a = float(s['amount'])
        if not np.isfinite(a) or a < 0:
            raise ValueError('수요 증가액은 유한한 0 이상의 수여야 합니다.')
        f[codes.index(s['code'])] += a
    if not f.sum() > 0:
        raise ValueError('한 산업 이상의 양의 수요 증가액을 입력하세요.')

    x = L @ f
    m = M @ f
    v = V @ f
    result = [
        {
            'code': c,
            'name': n,
            'production': float(a),
            'imports': float(b),
            'valueAdded': float(d),
        }
        for c, n, a, b, d in zip(codes, names, x, m, v)
    ]
    result.sort(key=lambda r: r['production'], reverse=True)

    return emit({
        'demand': float(f.sum()),
        'production': float(x.sum()),
        'imports': float(m.sum()),
        'valueAdded': float(v.sum()),
        'multiplier': float(x.sum() / f.sum()),
        'rows': result,
    })


def linkage_and_shock(payload):
    if not valid:
        raise ValueError('검증된 데이터가 없습니다.')
    data = json.loads(payload)
    mode = data.get('mode', 'linkage')
    target_code = norm_code(data.get('code', ''))
    top_k = max(1, min(100, int(data.get('top_k', 15))))
    shock_pct = float(data.get('shock_pct', 10.0))

    if target_code not in codes:
        raise ValueError('선택한 산업코드를 찾을 수 없습니다: ' + target_code)

    idx = codes.index(target_code)
    n = len(codes)

    if mode == 'linkage':
        b_l = L[:, idx]
        b_ad = ad[:, idx]
        b_order = np.argsort(-b_l)[:top_k]

        f_ad = ad[idx, :]
        f_sum = float(np.sum(f_ad))
        f_share = (f_ad / f_sum * 100.0) if f_sum > 0 else np.zeros(n, dtype=float)
        f_order = np.argsort(-f_ad)[:top_k]

        backward_rows = [
            {
                'code': codes[i],
                'name': names[i],
                'productionInducement': float(b_l[i]),
                'domesticInputCoeff': float(b_ad[i]),
            }
            for i in b_order
        ]
        forward_rows = [
            {
                'code': codes[i],
                'name': names[i],
                'forwardInputCoeff': float(f_ad[i]),
                'forwardSharePct': float(f_share[i]),
            }
            for i in f_order
        ]

        return emit({
            'mode': 'linkage',
            'target': {'code': codes[idx], 'name': names[idx]},
            'summary': {
                'backwardMultiplier': float(np.sum(b_l)),
                'backwardTop1': backward_rows[0]['name'] if backward_rows else '-',
                'forwardCoeffSum': f_sum,
                'forwardTop1': forward_rows[0]['name'] if forward_rows else '-',
            },
            'backwardRows': backward_rows,
            'forwardRows': forward_rows,
        })

    else:
        if not np.isfinite(shock_pct):
            raise ValueError('공급충격 비율은 유한한 숫자여야 합니다.')
        shock_rate = shock_pct / 100.0
        direct_vec = np.zeros(n, dtype=float)
        direct_vec[idx] = shock_rate

        eye = np.eye(n, dtype=float)
        try:
            total_prop = np.linalg.solve(eye - ad.T, direct_vec)
        except np.linalg.LinAlgError:
            total_prop = direct_vec + ad.T @ direct_vec

        order = np.argsort(-total_prop)[:top_k]
        rows = []
        for i in order:
            d_pct = shock_pct if i == idx else float(ad[idx, i] * shock_pct)
            t_pct = float(total_prop[i] * 100.0)
            i_pct = max(0.0, t_pct - d_pct)
            rows.append({
                'code': codes[i],
                'name': names[i],
                'directImpactPct': d_pct,
                'indirectImpactPct': i_pct,
                'totalImpactPct': t_pct,
            })

        return emit({
            'mode': 'shock',
            'target': {'code': codes[idx], 'name': names[idx]},
            'shockPct': shock_pct,
            'summary': {
                'targetShockPct': shock_pct,
                'top1Sector': rows[0]['name'] if rows else '-',
                'top1ImpactPct': rows[0]['totalImpactPct'] if rows else 0.0,
                'affectedCount': int(np.sum(total_prop > 1e-6)),
            },
            'rows': rows,
        })
