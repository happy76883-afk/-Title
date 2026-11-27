window.addEventListener("error", (e) => {
  const el = document.getElementById("auto-status");
  if (el) el.innerHTML = `<span class="err">스크립트 오류: ${e.message}</span>`;
});

window.addEventListener("unhandledrejection", (e) => {
  const el = document.getElementById("auto-status");
  if (el) el.innerHTML = `<span class="err">로드 오류: ${e.reason?.message || e.reason}</span>`;
});

const FILES = [
  { key: "basic", label: "기본부문", name: "(표)(2023연장)투입산출표_생산자가격_기본부문.xlsx" },
  { key: "major", label: "통합대분류", name: "(표)(2023연장)투입산출표_생산자가격_통합대분류.xlsx" },
  { key: "medium", label: "통합중분류", name: "(표)(2023연장)투입산출표_생산자가격_통합중분류.xlsx" },
  { key: "minor", label: "통합소분류", name: "(표)(2023연장)투입산출표_생산자가격_통합소분류.xlsx" }
];

const state = {
  workbooks: {},
  statusText: {},
  currentDatasetKey: "basic",
  shockItems: []
};

function toNum(v) {
  if (typeof v === "number") return Number.isFinite(v) ? v : 0;
  if (v == null) return 0;
  const s = String(v).replace(/,/g, "").trim();
  if (!s || s === "-" || s === "·" || s === "…") return 0;
  const n = Number(s);
  return Number.isFinite(n) ? n : 0;
}

function isNumericCell(v) {
  if (typeof v === "number") return Number.isFinite(v);
  if (v == null) return false;
  const s = String(v).replace(/,/g, "").trim();
  return /^-?\d+(\.\d+)?\$/.test(s);
}

function isCodeToken(v) {
  const s = String(v ?? "").trim();
  return /^[A-Z]?\d{2,7}\$/i.test(s);
}

function isPurePunctOrNum(v) {
  const s = String(v ?? "").replace(/,/g, "").trim();
  return !s || /^-?[\d.]+$/.test(s) \vert{}\vert{} /^[-·…()]+$/.test(s);
}

function isLabelText(v) {
  const s = String(v ?? "").trim();
  return Boolean(s) && !isCodeToken(s) && !isPurePunctOrNum(s);
}

function isNonSectorRow(code, name) {
  const c = String(code ?? "").trim();
  const n = String(name ?? "").trim();
  if (/^(95|96|97|98|99)\d*\$/.test(c)) return true;
  return /(중간투입계|중간수요계|부가가치계|총투입계|총산출|총공급|피용자보수|영업잉여|고정자본소모|생산세|보조금|최종수요|민간소비|정부소비|총고정자본형성|재고증감|수출계|수입계|합계|소계|투입계수|유발계수)/.test(n);
}

function extractUnit(raw) {
  for (let r = 0; r < Math.min(raw.length, 10); r++) {
    for (const cell of raw[r] || []) {
      const s = String(cell ?? "").trim();
      if (/단위\s*[:：]/.test(s)) return s;
      if (/^(백만원|억원|천원|원|지수|비율)\$/.test(s)) return `단위 : ${s}`;
    }
  }
  return "단위 : 백만원";
}

function parseSheetData(ws) {
  const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (!Array.isArray(raw) || raw.length === 0) {
    return { unitText: "단위 : 백만원", allRows: [], sectors: [], matrix: [] };
  }

  const unitText = extractUnit(raw);
  const maxCols = Math.max(...raw.map((r) => (Array.isArray(r) ? r.length : 0)), 1);

  let codeRowIdx = -1;
  let maxCodeHits = 0;
  for (let r = 0; r < Math.min(raw.length, 25); r++) {
    const hits = (raw[r] || []).filter((c) => isCodeToken(c)).length;
    if (hits >= 3 && hits > maxCodeHits) {
      maxCodeHits = hits;
      codeRowIdx = r;
    }
  }

  let nameRowIdx = -1;
  if (codeRowIdx !== -1) {
    for (let r = codeRowIdx + 1; r <= Math.min(raw.length - 1, codeRowIdx + 3); r++) {
      const hits = (raw[r] || []).filter((c) => isLabelText(c)).length;
      if (hits >= 2) {
        nameRowIdx = r;
        break;
      }
    }
  }
  if (nameRowIdx === -1) nameRowIdx = codeRowIdx !== -1 ? codeRowIdx : 0;

  const codeRow = codeRowIdx !== -1 ? raw[codeRowIdx] || [] : [];
  const nameRow = raw[nameRowIdx] || [];

  const colCodes = Array.from({ length: maxCols }, (_, c) => {
    const cd = String(codeRow[c] ?? "").trim();
    return isCodeToken(cd) ? cd : "";
  });

  const colNames = Array.from({ length: maxCols }, (_, c) => {
    const nm = String(nameRow[c] ?? "").trim();
    if (isLabelText(nm)) return nm;
    const cd = String(codeRow[c] ?? "").trim();
    return isCodeToken(cd) ? cd : nm;
  });

  const dataStartRow = Math.max(codeRowIdx, nameRowIdx) + 1;
  const allRows = [];

  for (let r = dataStartRow; r < raw.length; r++) {
    const row = raw[r];
    if (!Array.isArray(row)) continue;

    const dupCount = row.filter((cell, c) => {
      const val = String(cell ?? "").trim();
      return val && (val === String(nameRow[c] ?? "").trim() || val === String(codeRow[c] ?? "").trim());
    }).length;
    if (dupCount >= 3) continue;

    let rowCode = "";
    let rowName = "";
    for (let c = 0; c < Math.min(4, maxCols); c++) {
      const val = String(row[c] ?? "").trim();
      if (!val) continue;
      if (!rowCode && isCodeToken(val)) rowCode = val;
      else if (!rowName && isLabelText(val)) rowName = val;
    }

    if (!rowName) continue;
    if (!row.some((cell) => isNumericCell(cell))) continue;

    const cells = Array.from({ length: maxCols }, (_, c) => ({
      colIdx: c,
      colCode: colCodes[c],
      colName: colNames[c],
      rawVal: row[c],
      numVal: toNum(row[c]),
      isNum: isNumericCell(row[c])
    }));

    allRows.push({ rowIdx: r, code: rowCode, name: rowName, cells });
  }

  const endogenousRows = allRows.filter(
    (item) => !isNonSectorRow(item.code, item.name) && item.cells.filter((c) => c.isNum).length >= 3
  );

  const endogenousCols = [];
  for (let c = 0; c < maxCols; c++) {
    const cd = colCodes[c];
    const nm = colNames[c];
    if (isNonSectorRow(cd, nm) || /^9\d{3}\$/.test(cd)) continue;
    const hasNum = endogenousRows.some((r) => r.cells[c]?.isNum);
    if (hasNum && (isCodeToken(cd) || endogenousCols.length > 0)) {
      endogenousCols.push(c);
    }
  }

  const fallbackStart =
    endogenousRows.length > 0
      ? Math.max(0, endogenousRows[0].cells.findIndex((cell) => cell.isNum))
      : 2;

  const n = endogenousRows.length;
  const colSpan = endogenousCols.length >= n && n > 0 ? endogenousCols : null;

  const sectors = endogenousRows.map((item, idx) => ({
    idx,
    code: item.code || String(idx + 1).padStart(3, "0"),
    name: item.name
  }));

  const matrix = Array.from({ length: n }, (_, i) =>
    Array.from({ length: n }, (_, j) => {
      const targetCol = colSpan ? colSpan[j] : fallbackStart + j;
      const cell = endogenousRows[i].cells[targetCol];
      return cell ? cell.numVal : 0;
    })
  );

  return { unitText, allRows, sectors, matrix };
}

function renderStatusBar() {
  const bar = document.getElementById("auto-status");
  if (!bar) return;
  bar.innerHTML = FILES.map((f) => {
    const st = state.statusText[f.key] || '<span class="err">대기</span>';
    return `${f.label}: ${st}`;
  }).join(" &nbsp;|&nbsp; ");
}

function getActiveSheet(key = state.currentDatasetKey) {
  const ds = state.workbooks[key];
  if (!ds || !ds.sheetNames.length) return null;
  const sheetName =
    key === state.currentDatasetKey
      ? document.getElementById("sheet-select")?.value || ds.sheetNames[0]
      : ds.sheetNames[0];
  return { ds, sheetName, data: ds.sheets[sheetName] || ds.sheets[ds.sheetNames[0]] };
}

function syncDatasets() {
  const avail = FILES.filter((f) => state.workbooks[f.key]);
  for (const id of ["dataset-select", "chain-dataset", "shock-dataset"]) {
    const sel = document.getElementById(id);
    if (!sel) continue;
    sel.innerHTML = "";
    for (const f of avail) {
      const firstSheet = state.workbooks[f.key].sheetNames[0];
      const cnt = state.workbooks[f.key].sheets[firstSheet]?.sectors.length || 0;
      const opt = document.createElement("option");
      opt.value = f.key;
      opt.textContent = `${f.label} (${cnt}개 산업)`;
      sel.appendChild(opt);
    }
    if (state.workbooks[state.currentDatasetKey]) sel.value = state.currentDatasetKey;
  }
}

function populateSheets() {
  const ds = state.workbooks[state.currentDatasetKey];
  const sel = document.getElementById("sheet-select");
  if (!sel) return;
  sel.innerHTML = "";
  if (!ds) return;

  for (const sName of ds.sheetNames) {
    const opt = document.createElement("option");
    opt.value = sName;
    opt.textContent = sName;
    sel.appendChild(opt);
  }

  populateRows();
  populateSimSectors();
  populateChainSectors();
  populateShockSectors();
}

function populateRows() {
  const active = getActiveSheet();
  const sel = document.getElementById("row-select");
  const unitEl = document.getElementById("raw-unit");
  if (!sel) return;
  sel.innerHTML = "";
  if (!active) return;

  if (unitEl) unitEl.textContent = active.data.unitText || "단위 : 백만원";

  const q = (document.getElementById("row-search")?.value || "").trim().toLowerCase();
  const list = active.data.allRows.filter(
    (r) => !q || r.name.toLowerCase().includes(q) || r.code.toLowerCase().includes(q)
  );

  list.forEach((r, idx) => {
    const opt = document.createElement("option");
    opt.value = String(r.rowIdx);
    opt.textContent = r.code ? `${r.code} ${r.name}` : r.name;
    if (idx === 0) opt.selected = true;
    sel.appendChild(opt);
  });

  renderRawTable();
}

function renderRawTable() {
  const active = getActiveSheet();
  const tbody = document.querySelector("#raw-table tbody");
  if (!tbody) return;
  tbody.innerHTML = "";
  if (!active) return;

  const rowIdx = Number(document.getElementById("row-select")?.value);
  const colQ = (document.getElementById("col-search")?.value || "").trim().toLowerCase();
  const targetRow = active.data.allRows.find((r) => r.rowIdx === rowIdx);
  if (!targetRow) return;

  const cols = targetRow.cells.filter((cell, idx) => {
    const codeStr = String(cell.colCode ?? "").trim();
    const nameStr = String(cell.colName ?? "").trim();
    const valStr = String(cell.rawVal ?? "").trim();

    if (!codeStr && !nameStr) return false;
    if (idx < 2 && (!cell.isNum || valStr === targetRow.code || valStr === targetRow.name)) {
      return false;
    }
    if (!isCodeToken(codeStr) && !isLabelText(nameStr)) return false;

    const combined = `${codeStr} ${nameStr}`.toLowerCase();
    return !colQ || combined.includes(colQ);
  });

  tbody.innerHTML = cols
    .slice(0, 250)
    .map((cell) => {
