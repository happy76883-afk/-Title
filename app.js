const $ = id => document.getElementById(id);
const fmt = x => Number(x).toLocaleString('ko-KR', { maximumFractionDigits: 3 });

let worker, nextId = 0, requests = new Map(), sectors = [], rowChoices = [], cells = [], basket = [], result = null, lsResult = null;
const controlsToLock = [
  'sheet', 'row', 'sector', 'add', 'run', 'row-search', 'col-search', 'sector-search', 'amount',
  'ls-sector', 'ls-sector-search', 'ls-mode', 'ls-topk', 'ls-shock', 'ls-run'
];

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
}

function invalidateLs() {
  lsResult = null;
  $('ls-results').hidden = true;
}

async function load(buffer) {
  invalidate();
  invalidateLs();
  basket = [];
  renderBasket();
  for (const id of controlsToLock) $(id).disabled = true;
  if (worker) worker.terminate();
  for (const p of requests.values()) p.reject(Error('데이터 다시 로드'));
  requests.clear();
  startWorker();

  try {
    const info = await call('init', buffer);
    sectors = info.sectors;
    options($('sector'), sectors);
    options($('ls-sector'), sectors);
    $('sector-search').value = '';
    $('ls-sector-search').value = '';
    $('row-search').value = '';
    $('col-search').value = '';
    $('sheet').replaceChildren(...info.sheets.map(s => {
      const o = text('option', s);
      o.value = s;
      return o;
    }));
    $('validation').textContent = JSON.stringify(info.residuals, null, 2) + '\n최대 절대 잔차: 허용값 1e-8 미만';
    for (const id of controlsToLock) $(id).disabled = false;
    $('ls-status').textContent = '기준 산업과 분석 모드를 선택한 뒤 실행하세요.';
    await explore();
    $('status').textContent = '준비 완료 · ' + sectors.length + '개 부문 · ' + info.sheets.length + '개 시트 · 계수 검증 통과';
  } catch (e) {
    $('status').textContent = '준비 실패: ' + e.message + ' — 새로고침 후 재시도하세요.';
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
    for (const v of [c.code, c.name, c.value === null ? '빈 셀' : typeof c.value === 'number' ? fmt(c.value) : c.value]) tr.append(text('td', v));
    return tr;
  }));
}

$('sheet').onchange = () => {
  $('row-search').value = '';
  $('row').replaceChildren();
  explore();
};
$('row').onchange = explore;
$('row-search').oninput = () => {
  filterRows();
  if ($('row').value) explore();
  else {
    $('cells').replaceChildren();
    $('explore-caption').textContent = '일치하는 행이 없습니다.';
  }
};
$('col-search').oninput = renderCells;
$('sector-search').oninput = () => {
  const q = $('sector-search').value.trim();
  options($('sector'), sectors.filter(s => (s.code + ' ' + s.name).includes(q)));
};
$('ls-sector-search').oninput = () => {
  const q = $('ls-sector-search').value.trim();
  options($('ls-sector'), sectors.filter(s => (s.code + ' ' + s.name).includes(q)));
};

function renderBasket() {
  $('basket').replaceChildren(...basket.map((s, i) => {
    const li = text('li', s.code + ' ' + s.name + ' / ' + fmt(s.amount) + '백만원 ');
    const b = text('button', '삭제');
    b.type = 'button';
    b.onclick = () => {
      basket.splice(i, 1);
      invalidate();
      renderBasket();
    };
    li.append(b);
    return li;
  }));
  $('sim-status').textContent = basket.length ? basket.length + '개 입력 준비. 계산 버튼을 누르세요.' : '시나리오를 추가하세요.';
}

$('scenario').onsubmit = e => {
  e.preventDefault();
  const amount = Number($('amount').value), s = sectors.find(s => s.code === $('sector').value);
  if (!s || !Number.isFinite(amount) || amount <= 0) {
    $('sim-status').textContent = '산업과 양의 증가액을 확인하세요.';
    return;
  }
  basket.push({ ...s, amount });
  invalidate();
  renderBasket();
};

$('clear').onclick = () => {
  basket = [];
  invalidate();
  renderBasket();
};

$('run').onclick = async () => {
  if (!basket.length) {
    $('sim-status').textContent = '시나리오를 먼저 추가하세요.';
    return;
  }
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

$('ls-form').onsubmit = async e => {
  e.preventDefault();
  const code = $('ls-sector').value;
  const mode = $('ls-mode').value;
  const top_k = Number($('ls-topk').value) || 15;
  const shock_pct = Number($('ls-shock').value);
  if (!code) {
    $('ls-status').textContent = '기준 산업을 선택하세요.';
    return;
  }
  $('ls-run').disabled = true;
  $('ls-status').textContent = '연관·충격 분석 계산 중…';
  try {
    const r = await call('linkage_shock', { code, mode, top_k, shock_pct });
    lsResult = r;
    renderLsResult();
    $('ls-status').textContent = '분석 완료 · 기준 산업: ' + r.target.code + ' ' + r.target.name;
  } catch (err) {
    invalidateLs();
    $('ls-status').textContent = '분석 실패: ' + err.message;
  } finally {
    $('ls-run').disabled = false;
  }
};

function renderLsResult() {
  if (!lsResult) return;
  $('ls-results').hidden = false;

  if (lsResult.mode === 'linkage') {
    $('ls-linkage-view').hidden = false;
    $('ls-shock-view').hidden = true;
    const s = lsResult.summary;
    $('ls-cards').replaceChildren(...[
      ['후방 생산유발계수 합계', fmt(s.backwardMultiplier), ''],
      ['후방연쇄 1위 산업', s.backwardTop1, ''],
      ['전방 국산투입계수 합계', fmt(s.forwardCoeffSum), ''],
      ['전방연쇄 1위 산업', s.forwardTop1, '']
    ].map(([title, val, unit]) => {
      const d = text('div', title);
      d.className = 'card';
      d.append(text('strong', val), text('span', unit));
      return d;
    }));

    $('ls-backward-rows').replaceChildren(...lsResult.backwardRows.map(r => {
      const tr = document.createElement('tr');
      [r.code, r.name, fmt(r.productionInducement), fmt(r.domesticInputCoeff)].forEach(v => tr.append(text('td', v)));
      return tr;
    }));

    $('ls-forward-rows').replaceChildren(...lsResult.forwardRows.map(r => {
      const tr = document.createElement('tr');
      [r.code, r.name, fmt(r.forwardInputCoeff), fmt(r.forwardSharePct)].forEach(v => tr.append(text('td', v)));
      return tr;
    }));
  } else {
    $('ls-linkage-view').hidden = true;
    $('ls-shock-view').hidden = false;
    const s = lsResult.summary;
    $('ls-cards').replaceChildren(...[
      ['기준 투입충격', fmt(s.targetShockPct), '%'],
      ['최대 파급 산업', s.top1Sector, ''],
      ['최대 총파급률', fmt(s.top1ImpactPct), '%'],
      ['영향 산업 수', String(s.affectedCount), '개']
    ].map(([title, val, unit]) => {
      const d = text('div', title);
      d.className = 'card';
      d.append(text('strong', val), text('span', unit));
      return d;
    }));

    const maxPct = Math.max(...lsResult.rows.map(r => r.totalImpactPct), 0.001);
    $('ls-bars').replaceChildren(...lsResult.rows.slice(0, 15).map(r => {
      const d = text('div', r.name + ' · ' + fmt(r.totalImpactPct) + '%');
      d.className = 'bar';
      const t = document.createElement('div');
      t.className = 'track';
      const f = document.createElement('div');
      f.className = 'fill';
      f.style.width = Math.max(0, Math.min(100, r.totalImpactPct / maxPct * 100)) + '%';
      t.append(f);
      d.append(t);
      return d;
    }));

    $('ls-shock-rows').replaceChildren(...lsResult.rows.map(r => {
      const tr = document.createElement('tr');
      [r.code, r.name, fmt(r.directImpactPct), fmt(r.indirectImpactPct), fmt(r.totalImpactPct)].forEach(v => tr.append(text('td', v)));
      return tr;
    }));
  }
}

function csvCell(v) {
  let s = String(v);
  if (/^[=+@-]/.test(s) && typeof v !== 'number') s = "'" + s;
  return '"' + s.replaceAll('"', '""') + '"';
}

$('download').onclick = () => {
  if (!result) return;
  const rows = [
    ['산업코드', '산업명', '생산유발액_백만원', '수입유발액_백만원', '부가가치유발액_백만원'],
    ...result.rows.map(r => [r.code, r.name, r.production, r.imports, r.valueAdded]),
  ];
  const url = URL.createObjectURL(new Blob(['\ufeff' + rows.map(r => r.map(csvCell).join(',')).join('\r\n')], { type: 'text/csv;charset=utf-8' }));
  const a = document.createElement('a');
  a.href = url;
  a.download = 'io-2023-results.csv';
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
};

async function loadDefaultWorkbook() {
  try {
    $('status').textContent = '기본 Excel 내려받는 중…';
    const response = await fetch('./data/io-2023.xlsx');
    if (!response.ok) throw Error('기본 Excel 다운로드 실패 (HTTP ' + response.status + ')');
    await load(await response.arrayBuffer());
  } catch (error) {
    $('status').textContent = '준비 실패: ' + error.message + ' — 연결 상태를 확인하고 새로고침하세요.';
  }
}

$('load-basic').onclick = async () => {
  const button = $('load-basic');
  button.disabled = true;
  button.textContent = '불러오는 중…';
  try {
    await loadDefaultWorkbook();
    const ready = !$('sheet').disabled;
    button.textContent = ready ? '기본부문 불러오기 완료' : '기본부문 다시 불러오기';
    button.disabled = ready;
  } finally {
    if ($('sheet').disabled) {
      button.disabled = false;
      button.textContent = '기본부문 다시 불러오기';
    }
  }
};
