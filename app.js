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
  return /^-?\d+(\.\d+)?$/.test(s);
}

function isCodeToken(v) {
  const s = String(v ?? "").trim();
  return /^[A-Z]?\d{2,7}$/i.test(s);
}

function isPurePunctOrNum(v) {
  const s = String(v ?? "").replace(/,/g, "").trim();
  return !s || /^-?[\d.]+$/.test(s) || /^[-·…()]+$/.test(s);
}

function isLabelText(v) {
  const s = String(v ?? "").trim();
  return Boolean(s) && !isCodeToken(s) && !isPurePunctOrNum(s);
}

function isNonSectorRow(code, name) {
  const c = String(code ?? "").trim();
  const n = String(name ?? "").trim();
  if (/^(95|96|97|98|99)\d*$/.test(c)) return true;
  return /(중간투입계|중간수요계|부가가치계|총투입계|총산출|총공급|피용자보수|영업잉여|고정자본소모|생산세|보조금|최종수요|민간소비|정부소비|총고정자본형성|재고증감|수출계|수입계|합계|소계|투입계수|유발계수)/.test(n);
}

function extractUnit(raw) {
  for (let r = 0; r < Math.min(raw.length, 10); r++) {
    for (const cell of raw[r] || []) {
      const s = String(cell ?? "").trim();
      if (/단위\s*[:：]/.test(s)) return s;
      if (/^(백만원|억원|천원|원|지수|비율)$/.test(s)) return `단위 : ${s}`;
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

  // 시뮬레이션·연쇄·충격용 내생 산업 행(중간투입계·부가가치계·총투입계 등 제외)
  const endogenousRows = allRows.filter(
    (item) => !isNonSectorRow(item.code, item.name) && item.cells.filter((c) => c.isNum).length >= 3
  );

  const endogenousCols = [];
  for (let c = 0; c < maxCols; c++) {
    const cd = colCodes[c];
    const nm = colNames[c];
    if (isNonSectorRow(cd, nm) || /^9\d{3}$/.test(cd)) continue;
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
      const displayVal = cell.isNum
        ? cell.numVal.toLocaleString("ko-KR", { maximumFractionDigits: 4 })
        : String(cell.rawVal ?? "");
      return `<tr><td>${cell.colCode || "-"}</td><td>${cell.colName || "-"}</td><td class="num">${displayVal}</td></tr>`;
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
    const indirect = directVec.reduce(
      (acc, shockVal, j) => acc + shockVal * ((matrix[i]?.[j] || 0) / colTotals[j]),
      0
    );
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
  const sheet = ds?.sheets[ds.sheetNames[0]];
  if (!sheet || !sheet.sectors.length) return;

  const targetIdx = Number(document.getElementById("shock-sector")?.value || 0);
  const dRate = toNum(document.getElementById("demand-shock-rate")?.value) / 100;
  const sRate = toNum(document.getElementById("supply-shock-rate")?.value) / 100;

  const n = sheet.sectors.length;
  const rowTotals = sheet.matrix.map((r) => r.reduce((a, b) => a + Math.max(0, b), 0) || 1);
  const colTotals = Array.from({ length: n }, (_, c) =>
    sheet.matrix.reduce((a, r) => a + Math.max(0, r[c] || 0), 0) || 1
  );

  const baseScale = rowTotals[targetIdx] || 1;
  const rows = sheet.sectors.map((sec, idx) => {
    const aIj = (sheet.matrix[idx]?.[targetIdx] || 0) / colTotals[targetIdx];
    const bIj = (sheet.matrix[targetIdx]?.[idx] || 0) / rowTotals[targetIdx];
    const dImpact = baseScale * dRate * aIj;
    const sImpact = baseScale * sRate * bIj;
    return { name: sec.name, dImpact, sImpact, total: dImpact + sImpact, absScore: Math.abs(dImpact) + Math.abs(sImpact) };
  });

  const sumD = rows.reduce((a, r) => a + r.dImpact, 0);
  const sumS = rows.reduce((a, r) => a + r.sImpact, 0);

  kpiEl.innerHTML = `
    <div class="kpi"><div class="k-lbl">수요충격 파급 추정합</div><div class="k-val">${sumD.toLocaleString("ko-KR", { maximumFractionDigits: 1 })}</div></div>
    <div class="kpi"><div class="k-lbl">공급충격 파급 추정합</div><div class="k-val">${sumS.toLocaleString("ko-KR", { maximumFractionDigits: 1 })}</div></div>
    <div class="kpi"><div class="k-lbl">합산 충격 파급 추정</div><div class="k-val">${(sumD + sumS).toLocaleString("ko-KR", { maximumFractionDigits: 1 })}</div></div>
  `;

  tbody.innerHTML = [...rows]
    .sort((a, b) => b.absScore - a.absScore)
    .slice(0, 20)
    .map(
      (r, idx) =>
        `<tr><td>${idx + 1}</td><td>${r.name}</td><td class="num">${r.dImpact.toLocaleString("ko-KR", { maximumFractionDigits: 1 })}</td><td class="num">${r.sImpact.toLocaleString("ko-KR", { maximumFractionDigits: 1 })}</td><td class="num">${r.total.toLocaleString("ko-KR", { maximumFractionDigits: 1 })}</td></tr>`
    )
    .join("");
}

function downloadCSV(tableId, fileName) {
  const table = document.getElementById(tableId);
  if (!table) return;
  const lines = [...table.querySelectorAll("tr")].map((tr) =>
    [...tr.querySelectorAll("th, td")].map((td) => `"${td.textContent.trim().replace(/"/g, '""')}"`).join(",")
  );
  const blob = new Blob(["\uFEFF" + lines.join("\n")], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = fileName;
  a.click();
  URL.revokeObjectURL(url);
}

async function loadOne(fileObj) {
  const res = await fetch(`./data/${encodeURIComponent(fileObj.name)}`);
  if (!res.ok) throw new Error(`HTTP ${res.status}`);
  const buf = await res.arrayBuffer();
  if (typeof XLSX === "undefined") throw new Error("XLSX 미로드");
  const wb = XLSX.read(buf, { type: "array" });
  const sheets = {};
  for (const sName of wb.SheetNames) {
    sheets[sName] = parseSheetData(wb.Sheets[sName]);
  }
  state.workbooks[fileObj.key] = { label: fileObj.label, sheetNames: wb.SheetNames, sheets };
}

async function autoLoadAll() {
  for (const f of FILES) {
    state.statusText[f.key] = "읽는 중...";
  }
  renderStatusBar();

  for (const f of FILES) {
    try {
      await loadOne(f);
      state.statusText[f.key] = '<span class="ok">로드 완료</span>';
    } catch (err) {
      state.statusText[f.key] = `<span class="err">${err.message}</span>`;
    }
    renderStatusBar();
  }

  const firstAvail = FILES.find((f) => state.workbooks[f.key]);
  if (firstAvail) {
    state.currentDatasetKey = firstAvail.key;
    syncDatasets();
    populateSheets();
    runSimulation();
  }
}

document.getElementById("btn-load-basic")?.addEventListener("click", async () => {
  const basic = FILES.find((f) => f.key === "basic");
  try {
    await loadOne(basic);
    state.statusText.basic = '<span class="ok">로드 완료</span>';
    state.currentDatasetKey = "basic";
    renderStatusBar();
    syncDatasets();
    populateSheets();
    runSimulation();
  } catch (err) {
    state.statusText.basic = `<span class="err">${err.message}</span>`;
    renderStatusBar();
  }
});

document.getElementById("btn-reload-all")?.addEventListener("click", autoLoadAll);

document.getElementById("dataset-select")?.addEventListener("change", (e) => {
  state.currentDatasetKey = e.target.value;
  const cSel = document.getElementById("chain-dataset");
  const sSel = document.getElementById("shock-dataset");
  if (cSel) cSel.value = e.target.value;
  if (sSel) sSel.value = e.target.value;
  populateSheets();
  runSimulation();
});

document.getElementById("sheet-select")?.addEventListener("change", populateRows);
document.getElementById("row-search")?.addEventListener("input", populateRows);
document.getElementById("row-select")?.addEventListener("change", renderRawTable);
document.getElementById("col-search")?.addEventListener("input", renderRawTable);

document.getElementById("sim-search")?.addEventListener("input", populateSimSectors);
document.getElementById("btn-add-shock")?.addEventListener("click", addShockItem);
document.getElementById("btn-run-sim")?.addEventListener("click", runSimulation);
document.getElementById("btn-clear-shock")?.addEventListener("click", () => {
  state.shockItems = [];
  renderBasket();
  runSimulation();
});

document.getElementById("chain-dataset")?.addEventListener("change", populateChainSectors);
document.getElementById("chain-sector")?.addEventListener("change", runChainAnalysis);
document.getElementById("btn-run-chain")?.addEventListener("click", runChainAnalysis);

document.getElementById("shock-dataset")?.addEventListener("change", populateShockSectors);
document.getElementById("shock-sector")?.addEventListener("change", runShockSimulation);
document.getElementById("btn-run-shock")?.addEventListener("click", runShockSimulation);

document.getElementById("btn-csv-sim")?.addEventListener("click", () => downloadCSV("sim-table", "최종수요_파급효과.csv"));
document.getElementById("btn-csv-chain")?.addEventListener("click", () => downloadCSV("chain-table", "전후방_연쇄효과.csv"));
document.getElementById("btn-csv-shock")?.addEventListener("click", () => downloadCSV("shock-table", "공급수요_충격파급.csv"));

window.addEventListener("DOMContentLoaded", autoLoadAll);
