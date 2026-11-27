window.addEventListener("error", (e) => {
  const b = document.getElementById("auto-status");
  if (b) b.textContent = "에러: " + e.message;
});
const FILES = [
  { key: "basic", label: "기본부문", name: "(표)(2023연장)투입산출표_생산자가격_기본부문.xlsx" },
  { key: "major", label: "통합대분류", name: "(표)(2023연장)투입산출표_생산자가격_통합대분류.xlsx" },
  { key: "medium", label: "통합중분류", name: "(표)(2023연장)투입산출표_생산자가격_통합중분류.xlsx" },
  { key: "minor", label: "통합소분류", name: "(표)(2023연장)투입산출표_생산자가격_통합소분류.xlsx" }
];

const state = {
  workbooks: {},
  loadStatusMap: {},
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

function isSummaryOrValueAddedCode(code) {
  const c = String(code ?? "").trim();
  return /^(95|96|97|98|99)\d*$/.test(c);
}

function isSummaryOrValueAddedName(name) {
  const n = String(name ?? "").trim();
  return /(중간투입계|중간수요계|부가가치계|총투입계|총산출|총공급|피용자보수|영업잉여|고정자본소모|생산세|보조금|최종수요|민간소비|정부소비|총고정자본형성|재고증감|수출계|수입계|합계|소계|투입계수|유발계수)/.test(
    n
  );
}

function extractUnitText(raw) {
  for (let r = 0; r < Math.min(raw.length, 10); r++) {
    const row = raw[r] || [];
    for (const cell of row) {
      const s = String(cell ?? "").trim();
      if (/단위\s*[:：]/.test(s)) return s;
      if (/^(백만원|억원|천원|원|지수|비율)$/.test(s)) return `단위 : ${s}`; } } return "단위 : 백만원"; }  function parseSheetData(ws) { const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" }); if (!Array.isArray(raw) \vert{}\vert{} raw.length === 0) { return { unitText: "단위 : 백만원", colCodes: [], colNames: [], allRows: [], sectors: [], matrix: [] }; }  const unitText = extractUnitText(raw); const maxCols = Math.max(...raw.map((r) => (Array.isArray(r) ? r.length : 0)), 1);  let codeRowIdx = -1; let maxCodeHits = 0;  for (let r = 0; r < Math.min(raw.length, 25); r++) { const row = raw[r] \vert{}\vert{} []; const hits = row.filter((c) => isCodeToken(c)).length; if (hits >= 3 && hits > maxCodeHits) { maxCodeHits = hits; codeRowIdx = r; } }  let nameRowIdx = -1; if (codeRowIdx !== -1) { for (let r = codeRowIdx + 1; r <= Math.min(raw.length - 1, codeRowIdx + 3); r++) { const row = raw[r] \vert{}\vert{} []; const labelHits = row.filter((c) => isLabelText(c)).length; if (labelHits >= 2) { nameRowIdx = r; break; } } }  if (nameRowIdx === -1) { let bestScore = -1; for (let r = 0; r < Math.min(raw.length, 25); r++) { if (r === codeRowIdx) continue; const row = raw[r] \vert{}\vert{} []; const labelHits = row.filter((c) => isLabelText(c)).length; const kwBonus = row.some((c) => /(부문\vert{}합계\vert{}중간투입\vert{}부가가치\vert{}최종수요\vert{}총산출\vert{}수입\vert{}수출\vert{}소비\vert{}투자)/.test(String(c ?? "")) ) ? 5 : 0; const score = labelHits + kwBonus; if (labelHits >= 2 && score > bestScore) { bestScore = score; nameRowIdx = r; } } }  if (nameRowIdx === -1) nameRowIdx = codeRowIdx !== -1 ? codeRowIdx : 0;  const codeRow = codeRowIdx !== -1 ? raw[codeRowIdx] \vert{}\vert{} [] : []; const nameRow = raw[nameRowIdx] \vert{}\vert{} [];  const colCodes = Array.from({ length: maxCols }, (_, c) => { const cd = String(codeRow[c] ?? "").trim(); return isCodeToken(cd) ? cd : ""; });  const colNames = Array.from({ length: maxCols }, (_, c) => { const nm = String(nameRow[c] ?? "").trim(); if (isLabelText(nm)) return nm; const cd = String(codeRow[c] ?? "").trim(); return isCodeToken(cd) ? cd : nm; });  const dataStartRow = Math.max(codeRowIdx, nameRowIdx) + 1; const allRows = [];  for (let r = dataStartRow; r < raw.length; r++) { const row = raw[r]; if (!Array.isArray(row)) continue;  const headerDupCount = row.filter((cell, c) => { const val = String(cell ?? "").trim(); return val && (val === String(nameRow[c] ?? "").trim() \vert{}\vert{} val === String(codeRow[c] ?? "").trim()); }).length; if (headerDupCount >= 3) continue;  let rowCode = ""; let rowName = "";  for (let c = 0; c < Math.min(4, maxCols); c++) { const val = String(row[c] ?? "").trim(); if (!val) continue; if (!rowCode && isCodeToken(val)) { rowCode = val; } else if (!rowName && isLabelText(val)) { rowName = val; } }  if (!rowName) continue; if (!row.some((cell) => isNumericCell(cell))) continue;  const cells = Array.from({ length: maxCols }, (_, c) => ({ colIdx: c, colCode: colCodes[c], colName: colNames[c], rawVal: row[c], numVal: toNum(row[c]), isNum: isNumericCell(row[c]) }));  allRows.push({ rowIdx: r, code: rowCode, name: rowName, cells }); }  const endogenousRows = allRows.filter((item) => { if (isSummaryOrValueAddedCode(item.code)) return false; if (isSummaryOrValueAddedName(item.name)) return false; return item.cells.filter((c) => c.isNum).length >= 3; });  const endogenousColIndices = []; for (let c = 0; c < maxCols; c++) { const cd = colCodes[c]; const nm = colNames[c]; if (isSummaryOrValueAddedCode(cd) \vert{}\vert{} /^9\d{3}$/.test(cd)) continue; if (isSummaryOrValueAddedName(nm)) continue; const hasNumInCol = endogenousRows.some((r) => r.cells[c]?.isNum); if (hasNumInCol && (isCodeToken(cd) \vert{}\vert{} endogenousColIndices.length > 0)) { endogenousColIndices.push(c); } }  const fallbackStart = endogenousRows.length > 0 ? Math.max(0, endogenousRows[0].cells.findIndex((cell) => cell.isNum)) : 2;  const n = endogenousRows.length; const colSpan = endogenousColIndices.length >= n && n > 0 ? endogenousColIndices : null;  const sectors = endogenousRows.map((item, idx) => ({ idx, code: item.code \vert{}\vert{} String(idx + 1).padStart(3, "0"), name: item.name }));  const matrix = Array.from({ length: n }, (_, i) => Array.from({ length: n }, (_, j) => { const targetCol = colSpan ? colSpan[j] : fallbackStart + j; const cell = endogenousRows[i].cells[targetCol]; return cell ? cell.numVal : 0; }) );  return { unitText, colCodes, colNames, allRows, sectors, matrix }; }  function getActiveSheet(datasetKey = state.currentDatasetKey) { const ds = state.workbooks[datasetKey]; if (!ds \vert{}\vert{} !ds.sheetNames.length) return null; const chosenSheet = datasetKey === state.currentDatasetKey ? document.getElementById("sheet-select")?.value \vert{}\vert{} ds.sheetNames[0] : ds.sheetNames[0]; return { ds, sheetName: chosenSheet, data: ds.sheets[chosenSheet] \vert{}\vert{} ds.sheets[ds.sheetNames[0]] }; }  function renderAutoStatus() { const bar = document.getElementById("auto-status"); if (!bar) return; const parts = FILES.map((f) => { const info = state.loadStatusMap[f.key]; if (!info) { return `${f.label}: <span class="err">대기 중</span>`;
    }
    if (info.ok) {
      return `\${f.label}: <span class="ok">로드 완료</span>`;
    }
    return `\${f.label}: <span class="err">\${info.msg}</span>`;
  });
  bar.innerHTML = parts.join(" &nbsp;|&nbsp; ");
}

function syncDatasetDropdowns() {
  const available = FILES.filter((f) => state.workbooks[f.key]);
  const ids = ["dataset-select", "chain-dataset", "shock-dataset"];

  for (const id of ids) {
    const sel = document.getElementById(id);
    sel.innerHTML = "";
    for (const f of available) {
      const firstSheetKey = state.workbooks[f.key].sheetNames[0];
      const cnt = state.workbooks[f.key].sheets[firstSheetKey]?.sectors.length || 0;
      const opt = document.createElement("option");
      opt.value = f.key;
      opt.textContent = `${f.label} (${cnt}개 산업)`;
      sel.appendChild(opt);
    }
    if (state.workbooks[state.currentDatasetKey]) {
      sel.value = state.currentDatasetKey;
    }
  }
}

function populateSheetDropdown(preferredSheet = null) {
  const ds = state.workbooks[state.currentDatasetKey];
  const sel = document.getElementById("sheet-select");
  sel.innerHTML = "";
  if (!ds) return;

  for (const sName of ds.sheetNames) {
    const opt = document.createElement("option");
    opt.value = sName;
    opt.textContent = sName;
    sel.appendChild(opt);
  }

  if (preferredSheet && ds.sheets[preferredSheet]) {
    sel.value = preferredSheet;
  }

  populateRowDropdown();
  populateSimDropdown();
  populateChainDropdown();
  populateShockDropdown();
}

function populateRowDropdown() {
  const active = getActiveSheet();
  const sel = document.getElementById("row-select");
  const unitEl = document.getElementById("raw-unit");
  sel.innerHTML = "";
  if (!active) return;

  unitEl.textContent = active.data.unitText || "단위 : 백만원";

  const q = document.getElementById("row-search").value.trim().toLowerCase();
  const list = active.data.allRows.filter(
    (r) => !q || r.name.toLowerCase().includes(q) || r.code.toLowerCase().includes(q)
  );

  list.forEach((r, idx) => {
    const opt = document.createElement("option");
    opt.value = String(r.rowIdx);
    opt.textContent = r.code ? `${r.code}${r.name}` : r.name;
    if (idx === 0) opt.selected = true;
    sel.appendChild(opt);
  });

  renderRawTable();
}

function renderRawTable() {
  const active = getActiveSheet();
  const tbody = document.querySelector("#raw-table tbody");
  tbody.innerHTML = "";
  if (!active) return;

  const rowIdx = Number(document.getElementById("row-select").value);
  const colQ = document.getElementById("col-search").value.trim().toLowerCase();
  const targetRow = active.data.allRows.find((r) => r.rowIdx === rowIdx);
  if (!targetRow) return;

  const filteredCols = targetRow.cells.filter((cell, idx) => {
    const codeStr = String(cell.colCode ?? "").trim();
    const nameStr = String(cell.colName ?? "").trim();
    const valStr = String(cell.rawVal ?? "").trim();

    if (!codeStr && !nameStr) return false;
    if (idx < 2 && (!cell.isNum || valStr === targetRow.code || valStr === targetRow.name)) {
      return false;
    }
    if (!isCodeToken(codeStr) && !isLabelText(nameStr)) return false;

    const combined = `${codeStr}${nameStr}`.toLowerCase();
    return !colQ || combined.includes(colQ);
  });

  tbody.innerHTML = filteredCols
    .slice(0, 250)
    .map((cell) => {
      const displayVal = cell.isNum
        ? cell.numVal.toLocaleString("ko-KR", { maximumFractionDigits: 4 })
        : String(cell.rawVal ?? "");
      return `
        <tr>
          <td>\${cell.colCode || "-"}</td>
          <td>\${cell.colName || "-"}</td>
          <td class="num">\${displayVal}</td>
        </tr>
      `;
    })
    .join("");
}

function populateSimDropdown() {
  const active = getActiveSheet();
  const sel = document.getElementById("sim-sector");
  sel.innerHTML = "";
  if (!active) return;

  const q = document.getElementById("sim-search").value.trim().toLowerCase();
  const list = active.data.sectors.filter(
    (s) => !q || s.name.toLowerCase().includes(q) || s.code.toLowerCase().includes(q)
  );

  for (const s of list) {
    const opt = document.createElement("option");
    opt.value = String(s.idx);
    opt.textContent = `[${s.code}]${s.name}`;
    sel.appendChild(opt);
  }
}

function renderBasket() {
  const box = document.getElementById("shock-basket");
  if (!state.shockItems.length) {
    box.innerHTML = `<div style="font-size:13px; color:var(--muted);">추가된 산업 수요 증가 목록이 여기에 표시됩니다.</div>`;
    return;
  }

  const total = state.shockItems.reduce((sum, item) => sum + item.amount, 0);
  box.innerHTML =
    state.shockItems
      .map(
        (item, idx) => `
      <div class="basket-item">
        <span><strong>[${item.sectorCode}]${item.sectorName}</strong> (${item.datasetLabel}) — +${item.amount.toLocaleString("ko-KR")} 백만원</span>
        <button class="del-btn" type="button" data-del="\${idx}">삭제</button>
      </div>
    `
      )
      .join("") +
    `<div style="margin-top:8px; font-weight:700; font-size:13.5px;">합산 수요 증가액: ${total.toLocaleString("ko-KR")} 백만원 (${(total / 100).toLocaleString("ko-KR")} 억원)</div>`;

  box.querySelectorAll("[data-del]").forEach((btn) => {
    btn.addEventListener("click", () => {
      const idx = Number(btn.getAttribute("data-del"));
      state.shockItems.splice(idx, 1);
      renderBasket();
      runSimulation();
    });
  });
}

function addShockItem() {
  const active = getActiveSheet();
  if (!active) return;

  const sectorIdx = Number(document.getElementById("sim-sector").value || 0);
  const amount = Math.max(0, toNum(document.getElementById("sim-amount").value));
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

  if (!active || !active.data.sectors.length) {
    kpiEl.innerHTML = "";
    tbody.innerHTML = "";
    return;
  }

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
    const indirect = directVec.reduce((acc, shockVal, j) => {
      const aIj = (matrix[i]?.[j] || 0) / colTotals[j];
      return acc + shockVal * aIj;
    }, 0);
    const totalImpact = directVec[i] + indirect;
    return {
      name: sec.name,
      direct: directVec[i],
      indirect,
      totalImpact
    };
  });

  const sumIndirect = items.reduce((a, x) => a + x.indirect, 0);
  const sumTotal = items.reduce((a, x) => a + x.totalImpact, 0);

  kpiEl.innerHTML = `
    <div class="kpi">
      <div class="k-lbl">직접 수요 증가 합계</div>
      <div class="k-val">\${totalDirect.toLocaleString("ko-KR")} 백만원</div>
    </div>
    <div class="kpi">
      <div class="k-lbl">중간투입 간접 파급 합계</div>
      <div class="k-val">\${sumIndirect.toLocaleString("ko-KR", { maximumFractionDigits: 1 })} 백만원</div>
    </div>
    <div class="kpi">
      <div class="k-lbl">총 수요유발 파급합</div>
      <div class="k-val">\${sumTotal.toLocaleString("ko-KR", { maximumFractionDigits: 1 })} 백만원</div>
    </div>
  `;

  const ranked = [...items].sort((a, b) => b.totalImpact - a.totalImpact).slice(0, 25);
  tbody.innerHTML = ranked
    .map(
      (m, idx) => `
      <tr>
        <td>\${idx + 1}</td>
        <td>\${m.name}</td>
        <td class="num">\${m.direct.toLocaleString("ko-KR", { maximumFractionDigits: 1 })}</td>
        <td class="num">\${m.indirect.toLocaleString("ko-KR", { maximumFractionDigits: 1 })}</td>
        <td class="num">\${m.totalImpact.toLocaleString("ko-KR", { maximumFractionDigits: 1 })}</td>
      </tr>
    `
    )
    .join("");
}

function populateChainDropdown() {
  const key = document.getElementById("chain-dataset").value || state.currentDatasetKey;
  const ds = state.workbooks[key];
  const sel = document.getElementById("chain-sector");
  sel.innerHTML = "";
  if (!ds) return;

  const sheet = ds.sheets[ds.sheetNames[0]];
  if (!sheet) return;

  sheet.sectors.forEach((sec, idx) => {
    const opt = document.createElement("option");
    opt.value = String(idx);
    opt.textContent = `[${sec.code}]${sec.name}`;
    sel.appendChild(opt);
  });

  runChainAnalysis();
}

function runChainAnalysis() {
  const key = document.getElementById("chain-dataset").value || state.currentDatasetKey;
  const ds = state.workbooks[key];
  const kpiEl = document.getElementById("chain-kpi");
  const tbody = document.querySelector("#chain-table tbody");

  if (!ds) {
    kpiEl.innerHTML = "";
    tbody.innerHTML = "";
    return;
  }

  const sheet = ds.sheets[ds.sheetNames[0]];
  if (!sheet || !sheet.sectors.length) return;

  const targetIdx = Number(document.getElementById("chain-sector").value || 0);
  const n = sheet.sectors.length;
  const rowTotals = sheet.matrix.map((r) => r.reduce((a, b) => a + Math.max(0, b), 0) || 1);
  const colTotals = Array.from({ length: n }, (_, c) =>
    sheet.matrix.reduce((a, r) => a + Math.max(0, r[c] || 0), 0) || 1
  );

  const allBackwardSums = Array.from({ length: n }, (_, j) =>
    Array.from({ length: n }, (_, i) => (sheet.matrix[i]?.[j] || 0) / colTotals[j]).reduce((a, b) => a + b, 0)
  );
  const allForwardSums = Array.from({ length: n }, (_, i) =>
    Array.from({ length: n }, (_, j) => (sheet.matrix[i]?.[j] || 0) / rowTotals[i]).reduce((a, b) => a + b, 0)
  );

  const avgBackward = allBackwardSums.reduce((a, b) => a + b, 0) / (n || 1) || 1;
  const avgForward = allForwardSums.reduce((a, b) => a + b, 0) / (n || 1) || 1;

  const backwardSum = allBackwardSums[targetIdx] || 0;
  const forwardSum = allForwardSums[targetIdx] || 0;
  const powerIdx = backwardSum / avgBackward;
  const sensIdx = forwardSum / avgForward;

  kpiEl.innerHTML = `
    <div class="kpi">
      <div class="k-lbl">후방연쇄 계수합 (투입계수)</div>
      <div class="k-val">\${backwardSum.toFixed(4)}</div>
    </div>
    <div class="kpi">
      <div class="k-lbl">전방연쇄 계수합 (산출계수)</div>
      <div class="k-val">\${forwardSum.toFixed(4)}</div>
    </div>
    <div class="kpi">
      <div class="k-lbl">영향력계수 (후방 상대지수)</div>
      <div class="k-val">\${powerIdx.toFixed(3)}</div>
    </div>
    <div class="kpi">
      <div class="k-lbl">감응도계수 (전방 상대지수)</div>
      <div class="k-val">\${sensIdx.toFixed(3)}</div>
    </div>
  `;

  const rows = sheet.sectors.map((sec, idx) => {
    const aIj = (sheet.matrix[idx]?.[targetIdx] || 0) / colTotals[targetIdx];
    const bIj = (sheet.matrix[targetIdx]?.[idx] || 0) / rowTotals[targetIdx];
    return {
      name: sec.name,
      aIj,
      bIj,
      linkSum: aIj + bIj
    };
  });

  const top = [...rows].sort((a, b) => b.linkSum - a.linkSum).slice(0, 20);
  tbody.innerHTML = top
    .map(
      (r, idx) => `
      <tr>
        <td>\${idx + 1}</td>
        <td>\${r.name}</td>
        <td class="num">\${r.aIj.toFixed(4)}</td>
        <td class="num">\${r.bIj.toFixed(4)}</td>
        <td class="num
