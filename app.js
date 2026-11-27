const FILES = [
  { key: "basic", label: "기본부문", name: "(표)(2023연장)투입산출표_생산자가격_기본부문.xlsx" },
  { key: "major", label: "통합대분류", name: "(표)(2023연장)투입산출표_생산자가격_통합대분류.xlsx" },
  { key: "medium", label: "통합중분류", name: "(표)(2023연장)투입산출표_생산자가격_통합중분류.xlsx" },
  { key: "minor", label: "통합소분류", name: "(표)(2023연장)투입산출표_생산자가격_통합소분류.xlsx" }
];

const state = {
  workbooks: {},
  currentDatasetKey: "basic",
  shockItems: []
};

function setStatus(html) {
  const el = document.getElementById("auto-status");
  if (el) el.innerHTML = html;
}

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

function isSummaryOrValueAdded(code, name) {
  const c = String(code ?? "").trim();
  const n = String(name ?? "").trim();
  if (/^(95|96|97|98|99)\d*\$/.test(c)) return true;
  return /(중간투입계|중간수요계|부가가치계|총투입계|총산출|총공급|피용자보수|영업잉여|고정자본소모|생산세|보조금|최종수요|민간소비|정부소비|총고정자본형성|재고증감|수출계|수입계|합계|소계|투입계수|유발계수)/.test(n);
}

function parseSheetData(ws) {
  const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (!Array.isArray(raw) || raw.length === 0) {
    return { unitText: "단위 : 백만원", allRows: [], sectors: [], matrix: [] };
  }

  let unitText = "단위 : 백만원";
  for (let r = 0; r < Math.min(raw.length, 10); r++) {
    for (const c of raw[r] || []) {
      const s = String(c ?? "").trim();
      if (/단위\s*[:：]/.test(s)) unitText = s;
    }
  }

  const maxCols = Math.max(...raw.map((r) => (Array.isArray(r) ? r.length : 0)), 1);

  let codeRowIdx = -1;
  let maxHits = 0;
  for (let r = 0; r < Math.min(raw.length, 25); r++) {
    const hits = (raw[r] || []).filter((c) => isCodeToken(c)).length;
    if (hits >= 3 && hits > maxHits) {
      maxHits = hits;
      codeRowIdx = r;
    }
  }

  let nameRowIdx = -1;
  if (codeRowIdx !== -1) {
    for (let r = codeRowIdx + 1; r <= Math.min(raw.length - 1, codeRowIdx + 3); r++) {
      const labelHits = (raw[r] || []).filter((c) => isLabelText(c)).length;
      if (labelHits >= 2) {
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

  const startRow = Math.max(codeRowIdx, nameRowIdx) + 1;
  const allRows = [];

  for (let r = startRow; r < raw.length; r++) {
    const row = raw[r];
    if (!Array.isArray(row)) continue;

    const dup = row.filter((cell, c) => {
      const val = String(cell ?? "").trim();
      return val && (val === String(nameRow[c] ?? "").trim() || val === String(codeRow[c] ?? "").trim());
    }).length;
    if (dup >= 3) continue;

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
    (item) => !isSummaryOrValueAdded(item.code, item.name) && item.cells.filter((c) => c.isNum).length >= 3
  );

  const endogenousColIndices = [];
  for (let c = 0; c < maxCols; c++) {
    const cd = colCodes[c];
    const nm = colNames[c];
    if (isSummaryOrValueAdded(cd, nm) || /^9\d{3}\$/.test(cd)) continue;
    const hasNum = endogenousRows.some((r) => r.cells[c]?.isNum);
    if (hasNum && (isCodeToken(cd) || endogenousColIndices.length > 0)) {
      endogenousColIndices.push(c);
    }
  }

  const fallbackStart =
    endogenousRows.length > 0
      ? Math.max(0, endogenousRows[0].cells.findIndex((cell) => cell.isNum))
      : 2;

  const n = endogenousRows.length;
  const colSpan = endogenousColIndices.length >= n && n > 0 ? endogenousColIndices : null;

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
      const cnt = state.workbooks[f.key].sheets[state.workbooks[f.key].sheetNames[0]]?.sectors.length || 0;
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

  list.forEach((r, i) => {
    const opt = document.createElement("option");
    opt.value = String(r.rowIdx);
    opt.textContent = r.code ? `${r.code} ${r.name}` : r.name;
    if (i === 0) opt.selected = true;
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
    if (idx < 2 && (!cell.isNum || valStr === targetRow.code || valStr === targetRow.name)) return false;
    if (!isCodeToken(codeStr) && !isLabelText(nameStr)) return false;
    const combined = `${codeStr} ${nameStr}`.toLowerCase();
    return !colQ || combined.includes(colQ);
  });

  tbody.innerHTML = cols
    .slice(0, 250)
    .map((cell) => {
      const val = cell.isNum
        ? cell.numVal.toLocaleString("ko-KR", { maximumFractionDigits: 4 })
        : String(cell.rawVal ?? "");
      return `<tr><td>${cell.colCode || "-"}</td><td>${cell.colName || "-"}</td><td class="num">${val}</td></tr>`;
    })
    .join("");
}

function populateSimSectors() {
  const active = getActiveSheet();
  const sel = document.getElementById("sim-sector");
  if (!sel) return;
  sel.innerHTML = "";
  if (!active) return;

  const q = (document.getElementById("sim-search")?.value || "").trim().toLowerCase();
  const list = active.data.sectors.filter(
    (s) => !q || s.name.toLowerCase().includes(q) || s.code.toLowerCase().includes(q)
  );
  for (const s of list) {
    const opt = document.createElement("option");
    opt.value = String(s.idx);
    opt.textContent = `[${s.code}] ${s.name}`;
    sel.appendChild(opt);
  }
}

function renderBasket() {
  const box = document.getElementById("shock-basket");
  if (!box) return;
  if (!state.shockItems.length) {
    box.innerHTML = `<div style="font-size:13px; color:var(--muted);">추가된 산업 수요 증가 목록이 여기에 표시됩니다.</div>`;
    return;
  }
  const total = state.shockItems.reduce((sum, item) => sum + item.amount, 0);
  box.innerHTML =
    state.shockItems
      .map(
        (item, idx) =>
          `<div class="basket-item"><span><strong>[${item.sectorCode}] ${item.sectorName}</strong> (${item.datasetLabel}) — +${item.amount.toLocaleString("ko-KR")} 백만원</span><button class="del-btn" type="button" data-del="${idx}">삭제</button></div>`
      )
      .join("") +
    `<div style="margin-top:8px; font-weight:700; font-size:13.5px;">합산 수요 증가액: ${total.toLocaleString("ko-KR")} 백만원 (${(total / 100).toLocaleString("ko-KR")} 억원)</div>`;

  box.querySelectorAll("[data-del]").forEach((btn) => {
    btn.addEventListener("click", () => {
      state.shockItems.splice(Number(btn.getAttribute("data-del")), 1);
      renderBasket();
      runSimulation();
    });
  });
}

function addShockItem() {
  const active = getActiveSheet();
  if (!active) return;
  const sectorIdx = Number(document.getElementById("sim-sector")?.value || 0);
  const amount = Math.max(0, toNum(document.getElementById("sim-amount")?.value));
  const sec = active.data.sectors[sectorIdx];
  if (!sec || amount <= 0) return;

  state.shockItems.push({
    key: state.currentDatasetKey,
    datasetLabel: active.ds.label,
    sectorIdx,
    sectorCode: sec.code,
    sectorName: sec.name,
    amount
  });
  renderBasket();
  runSimulation();
}

function runSimulation() {
  const active = getActiveSheet();
  const kpiEl = document.getElementById("sim-kpi");
  const tbody = document.querySelector("#sim-table tbody");
  if (!kpiEl || !tbody) return;
  kpiEl.innerHTML = "";
  tbody.innerHTML = "";
  if (!active || !active.data.sectors.length) return;

  const { sectors, matrix } = active.data;
  const n = sectors.length;
  const colTotals = Array.from({ length: n }, (_, c) =>
    matrix.reduce((sum, r) => sum + Math.max(0, r[c] || 0), 0) || 1
  );

  const directVec = Array(n).fill(0);
  for (const item of state.shockItems) {
    if (item.key === state.currentDatasetKey && item.sectorIdx < n) {
      directVec[item.sectorIdx] += item.amount;
    }
  }

  const totalDirect = directVec.reduce((a, b) => a + b, 0);
  const items = sectors.map((sec, i) => {
    const indirect = directVec.reduce((acc, shockVal, j) => acc + shockVal * ((matrix[i]?.[j] || 0) / colTotals[j]), 0);
    return { name: sec.name, direct: directVec[i], indirect, totalImpact: directVec[i] + indirect };
  });

  const sumIndirect = items.reduce((a, x) => a + x.indirect, 0);
  const sumTotal = items.reduce((a, x) => a + x.totalImpact, 0);

  kpiEl.innerHTML = `
    <div class="kpi"><div class="k-lbl">직접 수요 증가 합계</div><div class="k-val">${totalDirect.toLocaleString("ko-KR")} 백만원</div></div>
    <div class="kpi"><div class="k-lbl">중간투입 간접 파급 합계</div><div class="k-val">${sumIndirect.toLocaleString("ko-KR", { maximumFractionDigits: 1 })} 백만원</div></div>
    <div class="kpi"><div class="k-lbl">총 수요유발 파급합</div><div class="k-val">${sumTotal.toLocaleString("ko-KR", { maximumFractionDigits: 1 })} 백만원</div></div>
  `;

  tbody.innerHTML = [...items]
    .sort((a, b) => b.totalImpact - a.totalImpact)
    .slice(0, 25)
    .map(
      (m, idx) =>
        `<tr><td>${idx + 1}</td><td>${m.name}</td><td class="num">${m.direct.toLocaleString("ko-KR", { maximumFractionDigits: 1 })}</td><td class="num">${m.indirect.toLocaleString("ko-KR", { maximumFractionDigits: 1 })}</td><td class="num">${m.totalImpact.toLocaleString("ko-KR", { maximumFractionDigits: 1 })}</td></tr>`
    )
    .join("");
}

function populateChainSectors() {
  const key = document.getElementById("chain-dataset")?.value || state.currentDatasetKey;
  const ds = state.workbooks[key];
  const sel = document.getElementById("chain-sector");
  if (!sel) return;
  sel.innerHTML = "";
  const sheet = ds?.sheets[ds.sheetNames[0]];
  if (!sheet) return;
  sheet.sectors.forEach((sec, idx) => {
    const opt = document.createElement("option");
    opt.value = String(idx);
    opt.textContent = `[${sec.code}] ${sec.name}`;
    sel.appendChild(opt);
  });
  runChainAnalysis();
}

function runChainAnalysis() {
  const key = document.getElementById("chain-dataset")?.value || state.currentDatasetKey;
  const ds = state.workbooks[key];
  const kpiEl = document.getElementById("chain-kpi");
  const tbody = document.querySelector("#chain-table tbody");
  if (!kpiEl || !tbody) return;
  kpiEl.innerHTML = "";
  tbody.innerHTML = "";
  const sheet = ds?.sheets[ds.sheetNames[0]];
  if (!sheet || !sheet.sectors.length) return;

  const targetIdx = Number(document.getElementById("chain-sector")?.value || 0);
  const n = sheet.sectors.length;
  const rowTotals = sheet.matrix.map((r) => r.reduce((a, b) => a + Math.max(0, b), 0) || 1);
  const colTotals = Array.from({ length: n }, (_, c) =>
    sheet.matrix.reduce((a, r) => a + Math.max(0, r[c] || 0), 0) || 1
  );

  const allBack = Array.from({ length: n }, (_, j) =>
    Array.from({ length: n }, (_, i) => (sheet.matrix[i]?.[j] || 0) / colTotals[j]).reduce((a, b) => a + b, 0)
  );
  const allFwd = Array.from({ length: n }, (_, i) =>
    Array.from({ length: n }, (_, j) => (sheet.matrix[i]?.[j] || 0) / rowTotals[i]).reduce((a, b) => a + b, 0)
  );

  const avgBack = allBack.reduce((a, b) => a + b, 0) / (n || 1) || 1;
  const avgFwd = allFwd.reduce((a, b) => a + b, 0) / (n || 1) || 1;

  const backSum = allBack[targetIdx] || 0;
  const fwdSum = allFwd[targetIdx] || 0;

  kpiEl.innerHTML = `
    <div class="kpi"><div class="k-lbl">후방연쇄 계수합 (투입계수)</div><div class="k-val">${backSum.toFixed(4)}</div></div>
    <div class="kpi"><div class="k-lbl">전방연쇄 계수합 (산출계수)</div><div class="k-val">${fwdSum.toFixed(4)}</div></div>
    <div class="kpi"><div class="k-lbl">영향력계수 (후방 상대지수)</div><div class="k-val">${(backSum / avgBack).toFixed(3)}</div></div>
    <div class="kpi"><div class="k-lbl">감응도계수 (전방 상대지수)</div><div class="k-val">${(fwdSum / avgFwd).toFixed(3)}</div></div>
  `;

  const rows = sheet.sectors.map((sec, idx) => {
    const aIj = (sheet.matrix[idx]?.[targetIdx] || 0) / colTotals[targetIdx];
    const bIj = (sheet.matrix[targetIdx]?.[idx] || 0) / rowTotals[targetIdx];
    return { name: sec.name, aIj, bIj, linkSum: aIj + bIj };
  });

  tbody.innerHTML = [...rows]
    .sort((a, b) => b.linkSum - a.linkSum)
    .slice(0, 20)
    .map(
      (r, idx) =>
        `<tr><td>${idx + 1}</td><td>${r.name}</td><td class="num">${r.aIj.toFixed(4)}</td><td class="num">${r.bIj.toFixed(4)}</td><td class="num">${r.linkSum.toFixed(4)}</td></tr>`
    )
    .join("");
}

function populateShockSectors() {
  const key = document.getElementById("shock-dataset")?.value || state.currentDatasetKey;
  const ds = state.workbooks[key];
  const sel = document.getElementById("shock-sector");
  if (!sel) return;
  sel.innerHTML = "";
  const sheet = ds?.sheets[ds.sheetNames[0]];
  if (!sheet) return;
  sheet.sectors.forEach((sec, idx) => {
    const opt = document.createElement("option");
    opt.value = String(idx);
    opt.textContent = `[${sec.code}] ${sec.name}`;
    sel.appendChild(opt);
  });
  runShockSimulation();
}

function runShockSimulation() {
  const key = document.getElementById("shock-dataset")?.value || state.currentDatasetKey;
  const ds = state.workbooks[key];
  const kpiEl = document.getElementById("shock-kpi");
  const tbody = document.querySelector("#shock-table tbody");
  if (!kpiEl || !tbody) return;
  kpiEl.innerHTML = "";
  tbody.innerHTML = "";
  const sheet = ds?.sheets[ds.
