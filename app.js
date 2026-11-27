window.addEventListener("error", (e) => {
  const b = document.getElementById("auto-status");
  if (b) b.innerHTML = `<span class="err">오류: ${e.message}</span>`;
});

window.addEventListener("unhandledrejection", (e) => {
  const b = document.getElementById("auto-status");
  if (b) b.innerHTML = `<span class="err">비동기 오류: ${e.reason?.message || e.reason}</span>`;
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

function isSummaryOrValueAddedCode(code) {
  const c = String(code ?? "").trim();
  return /^(95|96|97|98|99)\d*\$/.test(c);
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
      if (/^(백만원|억원|천원|원|지수|비율)\$/.test(s)) return `단위 : ${s}`;
    }
  }
  return "단위 : 백만원";
}

function parseSheetData(ws) {
  const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (!Array.isArray(raw) || raw.length === 0) {
    return { unitText: "단위 : 백만원", colCodes: [], colNames: [], allRows: [], sectors: [], matrix: [] };
  }

  const unitText = extractUnitText(raw);
  const maxCols = Math.max(...raw.map((r) => (Array.isArray(r) ? r.length : 0)), 1);

  let codeRowIdx = -1;
  let maxCodeHits = 0;

  for (let r = 0; r < Math.min(raw.length, 25); r++) {
    const row = raw[r] || [];
    const hits = row.filter((c) => isCodeToken(c)).length;
    if (hits >= 3 && hits > maxCodeHits) {
      maxCodeHits = hits;
      codeRowIdx = r;
    }
  }

  let nameRowIdx = -1;
  if (codeRowIdx !== -1) {
    for (let r = codeRowIdx + 1; r <= Math.min(raw.length - 1, codeRowIdx + 3); r++) {
      const row = raw[r] || [];
      const labelHits = row.filter((c) => isLabelText(c)).length;
      if (labelHits >= 2) {
        nameRowIdx = r;
        break;
      }
    }
  }

  if (nameRowIdx === -1) {
    let bestScore = -1;
    for (let r = 0; r < Math.min(raw.length, 25); r++) {
      if (r === codeRowIdx) continue;
      const row = raw[r] || [];
      const labelHits = row.filter((c) => isLabelText(c)).length;
      const kwBonus = row.some((c) =>
        /(부문|합계|중간투입|부가가치|최종수요|총산출|수입|수출|소비|투자)/.test(String(c ?? ""))
      )
        ? 5
        : 0;
      const score = labelHits + kwBonus;
      if (labelHits >= 2 && score > bestScore) {
        bestScore = score;
        nameRowIdx = r;
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

    const headerDupCount = row.filter((cell, c) => {
      const val = String(cell ?? "").trim();
      return val && (val === String(nameRow[c] ?? "").trim() || val === String(codeRow[c] ?? "").trim());
    }).length;
    if (headerDupCount >= 3) continue;

    let rowCode = "";
    let rowName = "";

    for (let c = 0; c < Math.min(4, maxCols); c++) {
      const val = String(row[c] ?? "").trim();
      if (!val) continue;
      if (!rowCode && isCodeToken(val)) {
        rowCode = val;
      } else if (!rowName && isLabelText(val)) {
        rowName = val;
      }
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

    allRows.push({
      rowIdx: r,
      code: rowCode,
      name: rowName,
      cells
    });
  }

  const endogenousRows = allRows.filter((item) => {
    if (isSummaryOrValueAddedCode(item.code)) return false;
    if (isSummaryOrValueAddedName(item.name)) return false;
    return item.cells.filter((c) => c.isNum).length >= 3;
  });

  const endogenousColIndices = [];
  for (let c = 0; c < maxCols; c++) {
    const cd = colCodes[c];
    const nm = colNames[c];
    if (isSummaryOrValueAddedCode(cd) || /^9\d{3}\$/.test(cd)) continue;
    if (isSummaryOrValueAddedName(nm)) continue;
    const hasNumInCol = endogenousRows.some((r) => r.cells[c]?.isNum);
    if (hasNumInCol && (isCodeToken(cd) || endogenousColIndices.length > 0)) {
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

  return { unitText, colCodes, colNames, allRows, sectors, matrix };
}

function getActiveSheet(datasetKey = state.currentDatasetKey) {
  const ds = state.workbooks[datasetKey];
  if (!ds || !ds.sheetNames.length) return null;
  const chosenSheet =
    datasetKey === state.currentDatasetKey
      ? document.getElementById("sheet-select")?.value || ds.sheetNames[0]
      : ds.sheetNames[0];
  return {
    ds,
    sheetName: chosenSheet,
    data: ds.sheets[chosenSheet] || ds.sheets[ds.sheetNames[0]]
  };
}

function renderAutoStatus() {
  const bar = document.getElementById("auto-status");
  if (!bar) return;
  const parts = FILES.map((f) => {
    const info = state.loadStatusMap[f.key];
    if (!info) return `${f.label}: <span class="err">대기 중</span>`;
    if (info.ok) return `${f.label}: <span class="ok">로드 완료</span>`;
    return `${f.label}: <span class="err">${info.msg}</span>`;
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
    opt.textContent = r.code ? `${r.code} ${r.name}` : r.name;
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
  if
