const $ = id => document.getElementById(id);
const fmt = x => Number(x).toLocaleString('ko-KR', { maximumFractionDigits: 3 });

const DATASETS = [
  {
    key: 'basic',
    label: '기본부문',
    btnId: 'switch-basic',
    urls: [
      './data/(표)(2023연장)투입산출표_생산자가격_기본부문.xlsx',
      './data/' + encodeURIComponent('(표)(2023연장)투입산출표_생산자가격_기본부문.xlsx'),
      './data/io-2023.xlsx'
    ]
  },
  {
    key: 'major',
    label: '대분류',
    btnId: 'switch-major',
    urls: [
      './data/(표)(2023연장)투입산출표_생산자가격_통합대분류.xlsx',
      './data/' + encodeURIComponent('(표)(2023연장)투입산출표_생산자가격_통합대분류.xlsx')
    ]
  },
  {
    key: 'medium',
    label: '중분류',
    btnId: 'switch-medium',
    urls: [
      './data/(표)(2023연장)투입산출표_생산자가격_통합중분류.xlsx',
      './data/' + encodeURIComponent('(표)(2023연장)투입산출표_생산자가격_통합중분류.xlsx')
    ]
  },
  {
    key: 'minor',
    label: '소분류',
    btnId: 'switch-minor',
    urls: [
      './data/(표)(2023연장)투입산출표_생산자가격_통합소분류.xlsx',
      './data/' + encodeURIComponent('(표)(2023연장)투입산출표_생산자가격_통합소분류.xlsx')
    ]
  }
];

const cache = new Map();
const errors = new Map();
const allValidation = new Map();

let worker, nextId = 0, requests = new Map(), sectors = [], rowChoices = [], cells = [], basket = [], result = null, currentKey = 'basic';

function text(tag, value) {
  const e = document.createElement(tag);
  e.textContent = value;
  return e;
}

function call(type, payload) {
  return new Promise((resolve, reject) => {
    const id = ++nextId;
    requests.set(id, { resolve, reject });
    worker.postMessage({ id, type, payload });
  });
}

function startWorker() {
  worker = new Worker('./worker.js');
  worker.onmessage = ({ data: d }) => {
    if (d.type === 'status') {
      $('status').textContent = d.text;
      return;
    }
    const p = requests.get(d.id);
    if (!p) return;
    requests.delete(d.id);
    d.ok ? p.resolve(d.result) : p.reject(Error(d.error));
  };
  worker.onerror = e => {
    for (const p of requests.values()) p.reject(Error(e.message || 'Worker 실행 오류'));
    requests.clear();
  };
}

function options(el, list) {
  el.replaceChildren(...list.map(x => {
    const o = document.createElement('option');
    o.value = x.code;
    o.textContent = x.code + ' ' + x.name;
    return o;
  }));
}

function invalidate() {
  result = null;
  $('results').hidden = true;
  $('linkage-results').hidden = true;
  $('shock-results').hidden = true;
}

const CONTROL_IDS = [
  'sheet', 'row', 'sector', 'add', 'run',
  'row-search', 'col-search', 'sector-search', 'amount',
  'linkage-sector', 'linkage-topk', 'run-linkage',
  'shock-sector', 'shock-pct', 'run-shock'
];

function setControlsDisabled(disabled) {
  for (const id of CONTROL_IDS) {
    if ($(id)) $(id).disabled = disabled;
  }
}

function refreshSwitchButtons() {
  for (const d of DATASETS) {
    const btn = $(d.btnId);
    if (!btn) continue;
    btn.disabled = !cache.has(d.key);
    btn.style.outline = d.key === currentKey ? '2px solid #1f2328' : 'none';
  }
}

function renderValidationSummary() {
  const lines = [];
  for (const d of DATASETS) {
    if (allValidation.has(d.key)) {
      lines.push(`[${d.label} - OK] ` + JSON.stringify(allValidation.get(d.key)));
    } else if (errors.has(d.key)) {
      lines.push(`[${d.label} - FAIL] ${errors.get(d.key)}`);
    } else {
      lines.push(`[${d.label} - 대기]`);
    }
  }
  lines.push('최대 절대 잔차: 허용값 1e-7 미만');
  $('validation').textContent = lines.join('\n');
}

async function initInWorker(buf) {
  if (worker) worker.terminate();
  for (const p of requests.values()) p.reject(Error('데이터 다시 로드'));
  requests.clear();
  startWorker();
  return await call('init', buf.slice(0));
}

async function activateDataset(key) {
  const cfg = DATASETS.find(d => d.key === key);
  if (!cfg) return;
  const buf = cache.get(key);
  if (!buf) {
    $('status').textContent = `${cfg.label} 불러오기 실패: ${errors.get(key) || '파일 없음'}`;
    return;
  }

  currentKey = key;
  refreshSwitchButtons();
  invalidate();
  basket = [];
  renderBasket();
  setControlsDisabled(true);

  try {
    const info = await initInWorker(buf);
    allValidation.set(key, info.residuals);
    renderValidationSummary();

    sectors = info.sectors;
    options($('sector'), sectors);
    options($('linkage-sector'), sectors);
    options($('shock-sector'), sectors);

    $('sector-search').value = '';
    $('row-search').value = '';
    $('col-search').value = '';
    $('sheet').replaceChildren(...info.sheets.map(s => {
      const o = text('option', s);
      o.value = s;
      return o;
    }));

    setControlsDisabled(false);
    await explore();

    const okList = DATASETS.filter(d => cache.has(d.key)).map(d => d.label).join(', ');
    const failList = DATASETS.filter(d => errors.has(d.key)).map(d => `${d.label}(${errors.get(d.key)})`).join(', ');
    $('status').textContent = failList
      ? `현재: ${cfg.label} (${sectors.length}개 부문) · 성공: [${okList}] · 실패: [${failList}]`
      : `전체 로드 완료 · 현재: ${cfg.label} (${sectors.length}개 부문) · [${okList}]`;
    $('linkage-status').textContent = '기준 산업을 선택하고 연쇄효과 계산을 누르세요.';
    $('shock-status').textContent = '기준 산업과 충격률을 입력하고 공급충격 계산을 누르세요.';
  } catch (e) {
    errors.set(key, e.message);
    renderValidationSummary();
    $('status').textContent = `${cfg.label} 준비 실패: ${e.message}`;
  }
}

async function fetchOne(cfg) {
  let lastErr = '파일 없음';
  for (const url of cfg.urls) {
    try {
      const res = await fetch(url);
      if (!res.ok) {
        lastErr = `HTTP ${res.status}`;
        continue;
      }
      cache.set(cfg.key, await res.arrayBuffer());
      errors.delete(cfg.key);
      return;
    } catch (err) {
      lastErr = err.message;
    }
  }
  errors.set(cfg.key, lastErr);
}

async function loadAllAtOnce() {
  $('load-all').disabled = true;
  setControlsDisabled(true);
  cache.clear();
  errors.clear();
  allValidation.clear();
  refreshSwitchButtons();
  renderValidationSummary();
  $('status').textContent = '기본·대분류·중분류·소분류 4개 엑셀 한 번에 내려받는 중…';

  await Promise.all(DATASETS.map(d => fetchOne(d)));
  refreshSwitchButtons();
  renderValidationSummary();

  // 다운로드된 4개 파일을 순차 검증해 검증 통과된 항목만 활성화
  const verifiedKeys = [];
  for (const d of DATASETS) {
    if (!cache.has(d.key)) continue;
    try {
      $('status').textContent = `${d.label} 계수 검증 중…`;
      const info = await initInWorker(cache.get(d.key));
      allValidation.set(d.key, info.residuals);
      verifiedKeys.push(d.key);
    } catch (e) {
      cache.delete(d.key);
      errors.set(d.key, e.message);
    }
  }

  refreshSwitchButtons();
  renderValidationSummary();

  if (verifiedKeys.length > 0) {
    const startKey = verifiedKeys.includes('basic') ? 'basic' : verifiedKeys[0];
    await activateDataset(startKey);
  } else {
    $('status').textContent = '4개 엑셀 모두 로드에 실패했습니다. data 폴더 파일명과 형식을 확인하세요.';
  }
  $('load-all').disabled = false;
}

let exploreSeq = 0;
async function explore() {
  const seq = ++exploreSeq;
  try {
    const d = await call('explore', { sheet: $('sheet').value, row: $('row').value });
    if (seq !== exploreSeq) return;
    rowChoices = d.rows;
    filterRows(d.selected);
    cells = d.cells;
    $('unit').textContent = d.unit;
    $('explore-caption').textContent = $('sheet').value + ' / 행: ' + d.selected + ' ' + d.name;
    renderCells();
  } catch (e) {
    $('status').textContent = e.message;
  }
}

function filterRows(selected) {
  const q = $('row-search').value.trim();
  options($('row'), rowChoices.filter(r => (r.code + ' ' + r.name).includes(q)));
  if ([...$('row').options].some(o => o.value === selected)) $('row').value = selected;
}

function renderCells() {
  const q = $('col-search').value.trim();
  $('cells').replaceChildren(...cells.filter(c => (c.code + ' ' + c.name).includes(q)).map(c => {
    const tr = document.createElement('tr');
    for (const v of [c.code, c.name, c.value === null ? '빈 셀' : typeof c.value === 'number' ? fmt(c.value) : c.value]) {
      tr.append(text('td', v));
    }
    return tr;
  }));
}

$('sheet').onchange = () => { $('row-search').value = ''; $('row').replaceChildren(); explore(); };
$('row').onchange = explore;
$('row-search').oninput = () => {
  filterRows();
  if ($('row').value) explore();
  else { $('cells').replaceChildren(); $('explore-caption').textContent = '일치하는 행이 없습니다.'; }
};
$('col-search').oninput = renderCells;
$('sector-search').oninput = () => {
  const q = $('sector-search').value.trim();
  options($('sector'), sectors.filter(s => (s.code + ' ' + s.name).includes(q)));
};

function renderBasket() {
  $('basket').replaceChildren(...basket.map((s, i) => {
    const li = text('li', s.code + ' ' + s.name + ' / ' + fmt(s.amount) + '백만원 ');
    const b = text('button', '삭제');
    b.type = 'button';
    b.onclick = () => { basket.splice(i, 1); invalidate(); renderBasket(); };
    li.append(b);
    return li;
  }));
  $('sim-status').textContent = basket.length ? basket.length + '개 입력 준비. 계산 버튼을 누르세요.' : '시나리오를 추가하세요.';
}

$('scenario').onsubmit = e => {
  e.preventDefault();
  const amount = Number($('amount').value);
  const s = sectors.find(s => s.code === $('sector').value);
  if (!s || !Number.isFinite(amount) || amount <= 0) {
    $('sim-status').textContent = '산업과 양의 증가액을 확인하세요.';
    return;
  }
  basket.push({ ...s, amount });
  invalidate();
  renderBasket();
};

$('clear').onclick = () => { basket = []; invalidate(); renderBasket(); };

$('run').onclick = async () => {
  if (!basket.length) { $('sim-status').textContent = '시나리오를 먼저 추가하세요.'; return; }
  const snapshot = basket.map(s => ({ ...s }));
  $('run').disabled = true;
  $('sim-status').textContent = '계산 중…';
  try {
    const r = await call('simulate', snapshot);
    if (JSON.stringify(snapshot) !== JSON.stringify(basket)) {
      $('sim-status').textContent = '입력이 변경되었습니다. 다시 계산하세요.';
      return;
    }
    result = r;
    renderResult();
    $('sim-status').textContent = '계산 완료. 아래 결과는 현재 시나리오에 대응합니다.';
  } catch (e) {
    invalidate();
    $('sim-status').textContent = '계산 실패: ' + e.message;
  } finally {
    $('run').disabled = false;
  }
};

function renderResult() {
  $('results').hidden = false;
  $('cards').replaceChildren(...[['수요 증가액', result.demand], ['생산유발액', result.production], ['수입유발액', result.imports], ['부가가치유발액', result.valueAdded]].map(([title, v]) => {
    const d = text('div', title);
    d.className = 'card';
    d.append(text('strong', fmt(v)), text('span', '백만원'));
    return d;
  }));

  const max = Math.max(...result.rows.map(r => r.production), 1);
  $('bars').replaceChildren(...result.rows.slice(0, 15).map(r => {
    const d = text('div', r.name + ' · ' + fmt(r.production) + '백만원');
    d.className = 'bar';
    const t = document.createElement('div');
    t.className = 'track';
    const f = document.createElement('div');
    f.className = 'fill';
    f.style.width = Math.max(0, r.production / max * 100) + '%';
    t.append(f);
    d.append(t);
    return d;
  }));

  $('result-rows').replaceChildren(...result.rows.map(r => {
    const tr = document.createElement('tr');
    [r.code, r.name, fmt(r.production), fmt(r.imports), fmt(r.valueAdded)].forEach(v => tr.append(text('td', v)));
    return tr;
  }));
}

$('run-linkage').onclick = async () => {
  const code = $('linkage-sector').value;
  const top_k = Number($('linkage-topk').value) || 15;
  $('run-linkage').disabled = true;
  $('linkage-status').textContent = '연쇄효과 계산 중…';
  try {
    const r = await call('linkage_shock', { mode: 'linkage', code, top_k });
    $('linkage-results').hidden = false;
    $('linkage-cards').replaceChildren(...[
      ['후방연쇄 계수합', fmt(r.summary.backwardMultiplier), ''],
      ['후방연쇄 1위', r.summary.backwardTop1, ''],
      ['전방연쇄 계수합', fmt(r.summary.forwardCoeffSum), ''],
      ['전방연쇄 1위', r.summary.forwardTop1, '']
    ].map(([title, val, unit]) => {
      const d = text('div', title);
      d.className = 'card';
      d.append(text('strong', val), text('span', unit));
      return d;
    }));

    $('backward-rows').replaceChildren(...r.backwardRows.map(row => {
      const tr = document.createElement('tr');
      [row.code, row.name, fmt(row.productionInducement), fmt(row.domesticInputCoeff)].forEach(v => tr.append(text('td', v)));
      return tr;
    }));

    $('forward-rows').replaceChildren(...r.forwardRows.map(row => {
      const tr = document.createElement('tr');
      [row.code, row.name, fmt(row.forwardInputCoeff), fmt(row.forwardSharePct)].forEach(v => tr.append(text('td', v)));
      return tr;
    }));

    $('linkage-status').textContent = `${r.target.code} ${r.target.name} 연쇄효과 계산 완료`;
  } catch (e) {
    $('linkage-status').textContent = '연쇄효과 계산 실패: ' + e.message;
  } finally {
    $('run-linkage').disabled = false;
  }
};

$('run-shock').onclick = async () => {
  const code = $('shock-sector').value;
  const shock_pct = Number($('shock-pct').value);
  if (!Number.isFinite(shock_pct)) {
    $('shock-status').textContent = '유효한 공급충격 비율(%)을 입력하세요.';
    return;
  }
  $('run-shock').disabled = true;
  $('shock-status').textContent = '공급충격 계산 중…';
  try {
    const r = await call('linkage_shock', { mode: 'shock', code, shock_pct, top_k: 20 });
    $('shock-results').hidden = false;
    $('shock-cards').replaceChildren(...[
      ['기준 충격률', fmt(r.summary.targetShockPct), '%'],
      ['최대 파급 산업', r.summary.top1Sector, ''],
      ['최대 총파급률', fmt(r.summary.top1ImpactPct), '%'],
      ['영향 산업 수', String(r.summary.affectedCount), '개']
    ].map(([title, val, unit]) => {
      const d = text('div', title);
      d.className = 'card';
      d.append(text('strong', val), text('span', unit));
      return d;
    }));

    $('shock-rows').replaceChildren(...r.rows.map(row => {
      const tr = document.createElement('tr');
      [row.code, row.name, fmt(row.directImpactPct), fmt(row.indirectImpactPct), fmt(row.totalImpactPct)].forEach(v => tr.append(text('td', v)));
      return tr;
    }));

    $('shock-status').textContent = `${r.target.code} ${r.target.name} (+${shock_pct}%) 공급충격 계산 완료`;
  } catch (e) {
    $('shock-status').textContent = '공급충격 계산 실패: ' + e.message;
  } finally {
    $('run-shock').disabled = false;
  }
};

function csvCell(v) {
  let s = String(v);
  if (/^[=+@-]/.test(s) && typeof v !== 'number') s = "'" + s;
  return '"' + s.replaceAll('"', '""') + '"';
}

$('download').onclick = () => {
  if (!result) return;
  const rows = [
    ['산업코드', '산업명', '생산유발액_백만원', '수입유발액_백만원', '부가가치유발액_백만원'],
    ...result.rows.map(r => [r.code, r.name, r.production, r.imports, r.valueAdded])
  ];
  const url = URL.createObjectURL(new Blob(['\ufeff' + rows.map(r => r.map(csvCell).join(',')).join('\r\n')], { type: 'text/csv;charset=utf-8' }));
  const a = document.createElement('a');
  a.href = url;
  a.download = `io-2023-${currentKey}-results.csv`;
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
};

$('load-all').onclick = loadAllAtOnce;
$('switch-basic').onclick = () => activateDataset('basic');
$('switch-major').onclick = () => activateDataset('major');
$('switch-medium').onclick = () => activateDataset('medium');
$('switch-minor').onclick = () => activateDataset('minor');
window.addEventListener('DOMContentLoaded', loadAllAtOnce);
