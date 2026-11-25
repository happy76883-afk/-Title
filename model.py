import json, io
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


def initialize(raw):
    global book, codes, names, L, M, V, v_vec, ad, am, a_mat, sheets, valid
    wb = openpyxl.load_workbook(io.BytesIO(bytes(raw)), read_only=True, data_only=True)
    book = {s: list(wb[s].values) for s in wb.sheetnames}
    wb.close()
    sheets = list(book)

    base = book['국산투입계수(Ad)']
    codes = [str(c) for c in base[4][2:] if c is not None and str(c).isdigit() and int(str(c)) < 9000]
    if len(codes) != 380 or len(set(codes)) != len(codes):
        raise ValueError('이 버전은 첨부된 2023년 380부문 워크북 구조만 지원합니다.')

    nm = {str(r[0]): str(r[1]) for r in base[6:] if r and r[0] is not None and len(r) > 1 and r[1] is not None}
    names = [nm[c] for c in codes]

    def matrix(s):
        rows = book[s]
        ri = {str(r[0]): i for i, r in enumerate(rows) if r and r[0] is not None}
        ci = {str(c): i for i, c in enumerate(rows[4]) if c is not None}
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
    ci = {str(c): i for i, c in enumerate(rows[4]) if c is not None}
    vr = next(r for r in rows if r and str(r[0]) == '9690')
    v_vec = np.array([vr[ci[c]] for c in codes], dtype=float)

    residuals = {
        '생산': float(np.max(np.abs((np.eye(len(codes)) - ad) @ L - np.eye(len(codes))))),
        '수입': float(np.max(np.abs(am @ L - M))),
        '부가가치': float(np.max(np.abs(v_vec[:, None] * L - V))),
    }
    valid = max(residuals.values()) < 1e-8
    if not valid:
        raise ValueError('계수 검증 실패: ' + str(residuals))

    return emit({
        'sectors': [{'code': c, 'name': n} for c, n in zip(codes, names)],
        'sheets': sheets,
        'residuals': residuals,
    })


def explore(sheet, rowcode):
    rows = book[sheet]
    choices = [{'code': str(r[0]), 'name': str(r[1])} for r in rows[6:] if len(r) > 1 and r[0] is not None and r[1] is not None]
    chosen = next((r for r in rows[6:] if len(r) > 1 and str(r[0]) == rowcode), None)
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
        cells.append({'code': str(rows[4][i]), 'name': str(rows[5][i] or ''), 'value': val})

    return emit({
        'rows': choices,
        'selected': str(chosen[0]),
        'name': str(chosen[1]),
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
    mode = data.get('mode', 'linkage')  # 'linkage' | 'shock'
    target_code = str(data.get('code', ''))
    top_k = max(1, min(100, int(data.get('top_k', 15))))
    shock_pct = float(data.get('shock_pct', 10.0))

    if target_code not in codes:
        raise ValueError('선택한 산업코드를 찾을 수 없습니다: ' + target_code)

    idx = codes.index(target_code)
    n = len(codes)

    if mode == 'linkage':
        # 후방연쇄: 기준 산업(idx) 열 기준 생산유발계수(L[:, idx])와 국산투입계수(ad[:, idx])
        b_l = L[:, idx]
        b_ad = ad[:, idx]
        b_order = np.argsort(-b_l)[:top_k]

        # 전방연쇄: 기준 산업(idx) 행 기준 다른 산업으로 가는 국산투입계수(ad[idx, :])
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
        # 공급충격(기준 산업 투입원가 shock_pct% 상승 시 투입망 전파)
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
