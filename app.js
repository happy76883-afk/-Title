const $ = id => document.getElementById(id);
const fmt = x => Number(x).toLocaleString('ko-KR', { maximumFractionDigits: 3 });

const WORKBOOKS = [
  { key: 'basic', label: '기본부문', file: '(표)(2023연장)투입산출표_생산자가격_기본부문.xlsx' },
  { key: 'major', label: '통합대분류', file: '(표)(2023연장)투입산출표_생산자가격_통합대분류.xlsx' },
  { key: 'medium', label: '통합중분류', file: '(표)(2023연장)투입산출표_생산자가격_통합중분류.xlsx' },
  { key: 'minor', label: '통합소분류', file: '(표)(2023연장)투입산출표_생산자가격_통합소분류.xlsx' }
];

const buffers = new Map();
const fileStatuses = new Map();
let activeKey = 'basic';

let worker, nextId = 0, requests = new Map(), sectors = [], rowChoices = [], cells = [], basket = [], result = null;

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

function updateFileSummary() {
  $('file-summary').textContent = WORKBOOKS
    .map(w => `${w.label}: ${fileStatuses.get(w.key) || '대기'}`)
    .join(' | ');
}

async function activateDataset(key) {
  const buf = buffers.get(key);
  const meta = WORKBOOKS.find(w => w.key === key);
  if (!buf || !meta) {
    $('status').textContent = `${meta ? meta.label : key} 파일이 다운로드되지 않았습니다.`;
    return;
  }

  activeKey = key;
  $('dataset').value = key;
  invalidate();
  basket = [];
  renderBasket();
  setControlsDisabled(true);

  if (worker) worker.terminate();
  for (const p of requests.values()) p.reject(Error('데이터 전환'));
  requests.clear();
  startWorker();

  try {
    // ArrayBuffer가 worker로 transfer되어 비워지지 않도록 복사본 전달
    const copy = buf.slice(0);
    const info = await call('init', copy);
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

    $('validation').textContent = `[${meta.label}]\n` + JSON.stringify(info.residuals, null, 2) + '\n최대 절대 잔차: 허용값 1e-7 미만';
    setControlsDisabled(false);

    await explore();
    $('status').textContent = `${meta.label} 준비 완료 · ${sectors.length}개 부문 · ${info.sheets.length}개 시트 · 계수 검증 통과`;
    $('linkage-status').textContent = '기준 산업을 선택하고 연쇄효과 계산을 누르세요.';
    $('shock-status').textContent = '기준 산업과 충격률을 입력하고 공급충격 계산을 누르세요.';
  } catch (e) {
    $('status').textContent = `${meta.label} 준비 실패: ${e.message}`;
  }
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

$('dataset').onchange = () => activateDataset($('dataset').value);
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
  a.download = `io-2023-${activeKey}-results.csv`;
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
};

async function loadAllWorkbooks() {
  const button = $('load-basic');
  button.disabled = true;
  button.textContent = '4개 엑셀 내려받는 중…';
  $('status').textContent = '4개 엑셀 파일 다운로드 중…';

  WORKBOOKS.forEach(w => fileStatuses.set(w.key, '다운로드 중'));
  updateFileSummary();

  await Promise.all(WORKBOOKS.map(async w => {
    try {
      const res = await fetch('./data/' + encodeURIComponent(w.file));
      if (!res.ok) throw Error('HTTP ' + res.status);
      buffers.set(w.key, await res.arrayBuffer());
      fileStatuses.set(w.key, '완료');
    } catch (err) {
      fileStatuses.set(w.key, `실패(${err.message})`);
    }
    updateFileSummary();
  }));

  const firstReady = WORKBOOKS.find(w => buffers.has(w.key));
  if (firstReady) {
    await activateDataset(buffers.has('basic') ? 'basic' : firstReady.key);
  } else {
    $('status').textContent = '4개 엑셀 모두 다운로드에 실패했습니다. data 폴더 파일명을 확인하세요.';
  }

  button.disabled = false;
  button.textContent = '4개 엑셀 다시 불러오기';
}

$('load-basic').onclick = loadAllWorkbooks;
window.addEventListener('DOMContentLoaded', loadAllWorkbooks);
