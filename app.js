const $ = id => document.getElementById(id);
const fmt = x => Number(x).toLocaleString('ko-KR', { maximumFractionDigits: 4 });

// data/ 폴더 안에 넣은 실제 엑셀 파일명으로 수정하세요.
const DEFAULT_EXCEL_PATH = './data/2023년 산업연관표(연장표).xlsx';

let worker = null;
let nextId = 0;
const requests = new Map();
let rowChoices = [];
let cells = [];
let exploreSeq = 0;

const controls = ['sheet', 'row-search', 'row', 'col-search'];

function text(tag, value) {
  const el = document.createElement(tag);
  el.textContent = value;
  return el;
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
    const req = requests.get(d.id);
    if (!req) return;
    requests.delete(d.id);
    d.ok ? req.resolve(d.result) : req.reject(Error(d.error));
  };
  worker.onerror = e => {
    for (const req of requests.values()) req.reject(Error(e.message || 'Worker 실행 오류'));
    requests.clear();
  };
}

function setOptions(el, list) {
  el.replaceChildren(...list.map(item => {
    const opt = document.createElement('option');
    opt.value = item.code;
    opt.textContent = item.code + (item.name && item.name !== item.code ? ' · ' + item.name : '');
    return opt;
  }));
}

function filterRows(selectedCode) {
  const q = $('row-search').value.trim().toLowerCase();
  const filtered = rowChoices.filter(r => (r.code + ' ' + r.name).toLowerCase().includes(q));
  setOptions($('row'), filtered);
  if (selectedCode && [...$('row').options].some(o => o.value === selectedCode)) {
    $('row').value = selectedCode;
  }
}

function renderCells() {
  const q = $('col-search').value.trim().toLowerCase();
  const filtered = cells.filter(c => (c.code + ' ' + c.name).toLowerCase().includes(q));
  $('cells').replaceChildren(...filtered.map(c => {
    const tr = document.createElement('tr');
    const displayVal = c.value === null
      ? '빈 셀'
      : typeof c.value === 'number'
        ? fmt(c.value)
        : c.value;
    [c.code, c.name, displayVal].forEach(v => tr.append(text('td', v)));
    return tr;
  }));
}

async function explore() {
  const seq = ++exploreSeq;
  try {
    const d = await call('explore', {
      sheet: $('sheet').value,
      row: $('row').value,
    });
    if (seq !== exploreSeq) return;
    rowChoices = d.rows;
    filterRows(d.selected);
    cells = d.cells;
    $('unit').textContent = d.unit;
    $('explore-caption').textContent = `${$('sheet').value} / 행: ${d.selected} ${d.name} (${cells.length}개 열)`;
    renderCells();
  } catch (err) {
    $('status').textContent = '조회 오류: ' + err.message;
  }
}

async function loadWorkbook(arrayBuffer) {
  controls.forEach(id => { $(id).disabled = true; });
  if (worker) worker.terminate();
  for (const req of requests.values()) req.reject(Error('다시 불러오는 중'));
  requests.clear();
  startWorker();

  const info = await call('init', arrayBuffer);
  $('sheet').replaceChildren(...info.sheets.map(name => {
    const opt = text('option', name);
    opt.value = name;
    return opt;
  }));
  $('validation').textContent = JSON.stringify(info.summary, null, 2);

  controls.forEach(id => { $(id).disabled = false; });
  $('row-search').value = '';
  $('col-search').value = '';
  await explore();
  $('status').textContent = `준비 완료 · ${info.sheets.length}개 시트 불러옴`;
}

$('sheet').onchange = () => {
  $('row-search').value = '';
  $('row').replaceChildren();
  explore();
};

$('row').onchange = explore;

$('row-search').oninput = () => {
  filterRows();
  if ($('row').value) {
    explore();
  } else {
    $('cells').replaceChildren();
    $('explore-caption').textContent = '일치하는 행이 없습니다.';
  }
};

$('col-search').oninput = renderCells;

$('load-basic').onclick = async () => {
  const btn = $('load-basic');
  btn.disabled = true;
  btn.textContent = '불러오는 중…';
  try {
    $('status').textContent = '기본 Excel 파일 내려받는 중…';
    const res = await fetch(encodeURI(DEFAULT_EXCEL_PATH));
    if (!res.ok) throw Error('파일 다운로드 실패 (HTTP ' + res.status + ')');
    await loadWorkbook(await res.arrayBuffer());
    btn.textContent = '기본부문 불러오기 완료';
  } catch (err) {
    $('status').textContent = '불러오기 실패: ' + err.message;
    btn.disabled = false;
    btn.textContent = '기본부문 다시 불러오기';
  }
};
