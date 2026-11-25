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

function toNum(v) {
  if (typeof v === "number") return Number.isFinite(v) ? v : 0;
  if (v == null) return 0;
  const s = String(v).replace(/,/g, "").trim();
  if (!s || s === "-" || s === "·" || s === "…") return 0;
  const n = Number(s);
  return Number.isFinite(n) ? n : 0;
}

function isNumericToken(rawVal) {
  if (typeof rawVal === "number") return Number.isFinite(rawVal);
  if (rawVal == null) return false;
  const cleaned = String(rawVal).replace(/,/g, "").trim();
  return /^-?\d+(\.\d+)?$/.test(cleaned);
}

function parseSheetData(ws) {
  const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (!Array.isArray(raw) || raw.length === 0) {
    return { headers: [], allRows: [], sectors: [], matrix: [] };
  }

  let headerRowIdx = 0;
  let bestScore = -1;
  for (let r = 0; r < Math.min(raw.length, 15); r++) {
    const row = raw[r] || [];
    const nonEmpty = row.filter((c) => String(c).trim() !== "").length;
    if (nonEmpty > bestScore) {
      bestScore = nonEmpty;
      headerRowIdx = r;
    }
  }

  const headerRow = raw[headerRowIdx] || [];
  const maxCols = Math.max(...raw.map((r) => (Array.isArray(r) ? r.length : 0)), headerRow.length);

  const headers = Array.from({ length: maxCols }, (_, c) => {
    const h = String(headerRow[c] ?? "").trim();
    return h || `열 ${c + 1}`;
  });

  const allRows = [];
  for (let r = headerRowIdx + 1; r < raw.length; r++) {
    const row = raw[r];
    if (!Array.isArray(row)) continue;

    let nameIdx = -1;
    let rowName = "";
    for (let c = 0; c < Math.min(4, maxCols); c++) {
      const val = String(row[c] ?? "").trim();
      if (val && !/^-?\d[\d,.]*$/.test(val)) {
        nameIdx = c;
        rowName = val;
        break;
      }
    }
    if (!rowName) continue;

    const codeVal = nameIdx > 0 ? String(row[nameIdx - 1] ?? "").trim() : "";
    const values = Array.from({ length: maxCols }, (_, c) => ({
      colIdx: c,
      colCode: headers[c],
      colName: String(headerRow[c] ?? "").trim() || headers[c],
      rawVal: row[c],
      numVal: toNum(row[c])
    }));

    allRows.push({
      rowIdx: r,
      code: codeVal,
      name: rowName,
      values
    });
  }

  const numericRows = allRows.filter(
    (item) => item.values.filter((v) => isNumericToken(v.rawVal)).length >= 3
  );

  const n = numericRows.length;
  const sectors = numericRows.map((item, idx) => ({
    idx,
    code: item.code || String(idx + 1).padStart(3, "0"),
    name: item.name
  }));

  const matrix = Array.from({ length: n }, (_, i) =>
    Array.from({ length: n }, (_, j) => {
      const cell = numericRows[i].values[j + 1] || numericRows[i].values[j];
      return cell ? cell.numVal : 0;
    })
  );

  return { headers, allRows, sectors, matrix };
}

function renderAutoStatus() {
  const bar = document.getElementById("auto-status");
  const parts = FILES.map((f) => {
    const loaded = !!state.workbooks[f.key];
    return `${f.label}: <span class="${loaded ? "ok" : "err"}">${loaded ? "자동 로드 완료" : "미로드"}</span>`;
  });
  bar.innerHTML = parts.join(" &nbsp;|&nbsp; ");
}

function populateDatasetSelects() {
  const selects = [
    document.getElementById("dataset-select"),
    document.getElementById("chain-dataset"),
    document.getElementById("shock-dataset")
  ];

  const available = FILES.filter((f) => state.workbooks[f.key]);
  for (const sel of selects) {
    sel.innerHTML = "";
    for (const f of available) {
      const cnt = Object.values(state.workbooks[f.key].sheets)[0]?.sectors.length || 0;
      const opt = document.createElement("option");
      opt.value = f.key;
      opt.textContent = `${f.label} (${cnt}개 부문)`;
      sel.appendChild(opt);
    }
    if (state.workbooks[state.currentDatasetKey]) {
      sel.value = state.currentDatasetKey;
    }
  }
}

function initSheetSelect(preferredSheetName = null) {
  const ds = state.workbooks[state.currentDatasetKey];
  const sheetSel = document.getElementById("sheet-select");
  sheetSel.innerHTML = "";

  if (!ds) return;
  const sheetNames = Object.keys(ds.sheets);
  for (const name of sheetNames) {
    const opt = document.createElement("option");
    opt.value = name;
    opt.textContent = name;
    sheetSel.appendChild(opt);
  }

  if (preferredSheetName && ds.sheets[preferredSheetName]) {
    sheetSel.value = preferredSheetName;
  }

  populateRowSelect();
  populateSimSectors();
  populateChainSectors();
  populateShockSectors();
}

function populateRowSelect() {
  const ds = state.workbooks[state.currentDatasetKey];
  const sheetName = document.getElementById("sheet-select").value;
  const search = document.getElementById("row-search").value.trim().toLowerCase();
  const rowSel = document.getElementById("row-select");
  rowSel.innerHTML = "";

  if (!ds || !ds.sheets[sheetName]) return;
  const rows = ds.sheets[sheetName].allRows;
  const filtered = rows.filter(
    (r) => !search || r.name.toLowerCase().includes(search) || r.code.toLowerCase().includes(search)
  );

  filtered.forEach((r, idx) => {
    const opt = document.createElement("option");
    opt.value = String(r.rowIdx);
    opt.textContent = `${r.code ? `[${r.code}] ` : ""}${r.name}`;
    if (idx === 0) opt.selected = true;
    rowSel.appendChild(opt);
  });

  renderRawConnections();
}

function renderRawConnections() {
  const ds = state.workbooks[state.currentDatasetKey];
  const sheetName = document.getElementById("sheet-select").value;
  const rowIdx = Number(document.getElementById("row-select").value);
  const colQuery = document.getElementById("col-search").value.trim().toLowerCase();
  const tbody = document.querySelector("#raw-table tbody");
  tbody.innerHTML = "";

  if (!ds || !ds.sheets[sheetName]) return;
  const targetRow = ds.sheets[sheetName].allRows.find((r) => r.rowIdx === rowIdx);
  if (!targetRow) return;

  const matchedCols = targetRow.values.filter((v) => {
    const labelText = `${v.colCode} ${v.colName}`.toLowerCase();
    return !colQuery || labelText.includes(colQuery);
  });

  tbody.innerHTML = matchedCols
    .slice(0, 150)
    .map(
      (c, i) => `
      <tr>
        <td>${c.colCode || i + 1}</td>
        <td>${c.colName}</td>
        <td class="num">${typeof c.rawVal === "number" ? c.rawVal.toLocaleString("ko-KR") : String(c.rawVal ?? "")}</td>
      </tr>
    `
    )
    .join("");
}

function populateSimSectors() {
  const ds = state.workbooks[state.currentDatasetKey];
  const q = document.getElementById("sim-search").value.trim().toLowerCase();
  const sel = document.getElementById("sim-sector");
  sel.innerHTML = "";

  if (!ds) return;
  const firstSheet = Object.values(ds.sheets)[0];
  if (!firstSheet) return;

  const list = firstSheet.sectors.filter(
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
  if (!state.shockItems.length) {
    box.innerHTML = `<div style="font-size:13px; color:var(--muted);">추가된 산업 수요 증가 목록이 여기에 표시됩니다.</div>`;
    return;
  }

  const total = state.shockItems.reduce((acc, item) => acc + item.amount, 0);
  box.innerHTML =
    state.shockItems
      .map(
        (item, i) => `
      <div class="basket-item">
        <span><strong>${item.sectorName}</strong> (${item.datasetLabel}) — +${item.amount.toLocaleString("ko-KR")} 백만원</span>
        <button class="del-btn" type="button" data-del="${i}">삭제</button>
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
  const ds = state.workbooks[state.currentDatasetKey];
  if (!ds) return;
  const firstSheet = Object.values(ds.sheets)[0];
  const sectorIdx = Number(document.getElementById("sim-sector").value || 0);
  const amount = Math.max(0, toNum(document.getElementById("sim-amount").value));
  const sector = firstSheet.sectors[sectorIdx];
  if (!sector || amount <= 0) return;

  state.shockItems.push({
    key: state.currentDatasetKey,
    datasetLabel: ds.label,
    sectorIdx,
    sectorName: sector.name,
    amount
  });

  renderBasket();
  runSimulation();
}

function runSimulation() {
  const ds = state.workbooks[state.currentDatasetKey];
  const kpiEl = document.getElementById("sim-kpi");
  const tbody = document.querySelector("#sim-table tbody");

  if (!ds) {
    kpiEl.innerHTML = "";
    tbody.innerHTML = "";
    return;
  }

  const sheet = Object.values(ds.sheets)[0];
  if (!sheet || !sheet.sectors.length) return;

  const n = sheet.sectors.length;
  const colTotals = Array.from({ length: n }, (_, c) =>
    sheet.matrix.reduce((a, r) => a + Math.max(0, r[c] || 0), 0) || 1
  );

  const directShock = Array(n).fill(0);
  for (const item of state.shockItems) {
    if (item.key === state.currentDatasetKey && item.sectorIdx < n) {
      directShock[item.sectorIdx] += item.amount;
    }
  }

  const totalDirect = directShock.reduce((a, b) => a + b, 0);

  const items = sheet.sectors.map((sec, i) => {
    const indirect = directShock.reduce((acc, shockVal, j) => {
      const aIj = (sheet.matrix[i]?.[j] || 0) / colTotals[j];
      return acc + shockVal * aIj;
    }, 0);
    const totalImpact = directShock[i] + indirect;
    return { name: sec.name, indirect, totalImpact };
  });

  const sumIndirect = items.reduce((a, x) => a + x.indirect, 0);
  const sumTotal = items.reduce((a, x) => a + x.totalImpact, 0);

  kpiEl.innerHTML = `
    <div class="kpi">
      <div class="k-lbl">직접 최종수요 증가 합계</div>
      <div class="k-val">${totalDirect.toLocaleString("ko-KR")} 백만원</div>
    </div>
    <div class="kpi">
      <div class="k-lbl">중간투입 간접유발 합계</div>
      <div class="k-val">${sumIndirect.toLocaleString("ko-KR", { maximumFractionDigits: 1 })} 백만원</div>
    </div>
    <div class="kpi">
      <div class="k-lbl">총 수요유발 파급합</div>
      <div class="k-val">${sumTotal.toLocaleString("ko-KR", { maximumFractionDigits: 1 })} 백만원</div>
    </div>
  `;

  const ranked = [...items].sort((a, b) => b.totalImpact - a.totalImpact).slice(0, 25);
  tbody.innerHTML = ranked
    .map(
      (m, idx) => `
      <tr>
        <td>${idx + 1}</td>
        <td>${m.name}</td>
        <td class="num">${m.indirect.toLocaleString("ko-KR", { maximumFractionDigits: 1 })}</td>
        <td class="num">${m.totalImpact.toLocaleString("ko-KR", { maximumFractionDigits: 1 })}</td>
      </tr>
    `
    )
    .join("");
}

function populateChainSectors() {
  const key = document.getElementById("chain-dataset").value || state.currentDatasetKey;
  const ds = state.workbooks[key];
  const sel = document.getElementById("chain-sector");
  sel.innerHTML = "";
  if (!ds) return;
  const sheet = Object.values(ds.sheets)[0];
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
  const key = document.getElementById("chain-dataset").value || state.currentDatasetKey;
  const ds = state.workbooks[key];
  const kpiEl = document.getElementById("chain-kpi");
  const tbody = document.querySelector("#chain-table tbody");

  if (!ds) {
    kpiEl.innerHTML = "";
    tbody.innerHTML = "";
    return;
  }

  const sheet = Object.values(ds.sheets)[0];
  if (!sheet || !sheet.sectors.length) return;

  const targetIdx = Number(document.getElementById("chain-sector").value || 0);
  const n = sheet.sectors.length;
  const rowTotals = sheet.matrix.map((r) => r.reduce((a, b) => a + Math.max(0, b), 0) || 1);
  const colTotals = Array.from({ length: n }, (_, c) =>
    sheet.matrix.reduce((a, r) => a + Math.max(0, r[c] || 0), 0) || 1
  );

  const rows = sheet.sectors.map((sec, idx) => {
    const aIj = (sheet.matrix[idx]?.[targetIdx] || 0) / colTotals[targetIdx];
    const bIj = (sheet.matrix[targetIdx]?.[idx] || 0) / rowTotals[targetIdx];
    return { name: sec.name, aIj, bIj, linkSum: aIj + bIj };
  });

  const backwardSum = rows.reduce((a, r) => a + r.aIj, 0);
  const forwardSum = rows.reduce((a, r) => a + r.bIj, 0);

  kpiEl.innerHTML = `
    <div class="kpi">
      <div class="k-lbl">선택 기준 산업</div>
      <div class="k-val" style="font-size:15px;">${sheet.sectors[targetIdx]?.name || ""}</div>
    </div>
    <div class="kpi">
      <div class="k-lbl">후방연쇄 계수합 (투입계수)</div>
      <div class="k-val">${backwardSum.toFixed(4)}</div>
    </div>
    <div class="kpi">
      <div class="k-lbl">전방연쇄 계수합 (산출계수)</div>
      <div class="k-val">${forwardSum.toFixed(4)}</div>
    </div>
  `;

  const top = [...rows].sort((a, b) => b.linkSum - a.linkSum).slice(0, 20);
  tbody.innerHTML = top
    .map(
      (r, idx) => `
      <tr>
        <td>${idx + 1}</td>
        <td>${r.name}</td>
        <td class="num">${r.aIj.toFixed(4)}</td>
        <td class="num">${r.bIj.toFixed(4)}</td>
        <td class="num">${r.linkSum.toFixed(4)}</td>
      </tr>
    `
    )
    .join("");
}

function populateShockSectors() {
  const key = document.getElementById("shock-dataset").value || state.currentDatasetKey;
  const ds = state.workbooks[key];
  const sel = document.getElementById("shock-sector");
  sel.innerHTML = "";
  if (!ds) return;
  const sheet = Object.values(ds.sheets)[0];
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
  const key = document.getElementById("shock-dataset").value || state.currentDatasetKey;
  const ds = state.workbooks[key];
  const kpiEl = document.getElementById("shock-kpi");
  const tbody = document.querySelector("#shock-table tbody");

  if (!ds) {
    kpiEl.innerHTML = "";
    tbody.innerHTML = "";
    return;
  }

  const sheet = Object.values(ds.sheets)[0];
  if (!sheet || !sheet.sectors.length) return;

  const targetIdx = Number(document.getElementById("shock-sector").value || 0);
  const dRate = toNum(document.getElementById("demand-shock-rate").value) / 100;
  const sRate = toNum(document.getElementById("supply-shock-rate").value) / 100;

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

  const sumDemand = rows.reduce((a, r) => a + r.dImpact, 0);
  const sumSupply = rows.reduce((a, r) => a + r.sImpact, 0);

  kpiEl.innerHTML = `
    <div class="kpi">
      <div class="k-lbl">수요충격 파급 추정합</div>
      <div class="k-val">${sumDemand.toLocaleString("ko-KR", { maximumFractionDigits: 1 })}</div>
    </div>
    <div class="kpi">
      <div class="k-lbl">공급충격 파급 추정합</div>
      <div class="k-val">${sumSupply.toLocaleString("ko-KR", { maximumFractionDigits: 1 })}</div>
    </div>
    <div class="kpi">
      <div class="k-lbl">합산 충격 파급 추정</div>
      <div class="k-val">${(sumDemand + sumSupply).toLocaleString("ko-KR", { maximumFractionDigits: 1 })}</div>
    </div>
  `;

  const top = [...rows].sort((a, b) => b.absScore - a.absScore).slice(0, 20);
  tbody.innerHTML = top
    .map(
      (r, idx) => `
      <tr>
        <td>${idx + 1}</td>
        <td>${r.name}</td>
        <td class="num">${r.dImpact.toLocaleString("ko-KR", { maximumFractionDigits: 1 })}</td>
        <td class="num">${r.sImpact.toLocaleString("ko-KR", { maximumFractionDigits: 1 })}</td>
        <td class="num">${r.total.toLocaleString("ko-KR", { maximumFractionDigits: 1 })}</td>
      </tr>
    `
    )
    .join("");
}

async function loadDataset(fileObj) {
  const url = encodeURI(`./data/${fileObj.name}`);
  const res = await fetch(url);
  if (!res.ok) throw new Error(`HTTP ${res.status}`);
  const buf = await res.arrayBuffer();
  const wb = XLSX.read(buf, { type: "array" });

  const sheets = {};
  for (const sName of wb.SheetNames) {
    sheets[sName] = parseSheetData(wb.Sheets[sName]);
  }

  state.workbooks[fileObj.key] = {
    label: fileObj.label,
    wb,
    sheets
  };
}

async function autoLoadAll(focusKey = "basic") {
  for (const f of FILES) {
    try {
      await loadDataset(f);
    } catch (_) {
      // 개별 파일 실패 시 상태 표시만 갱신
    }
    renderAutoStatus();
  }

  if (state.workbooks[focusKey]) {
    state.currentDatasetKey = focusKey;
  } else {
    const firstAvailable = FILES.find((f) => state.workbooks[f.key]);
    if (firstAvailable) state.currentDatasetKey = firstAvailable.key;
  }

  populateDatasetSelects();
  initSheetSelect();
  runSimulation();
}

document.getElementById("btn-load-basic").addEventListener("click", async () => {
  const basicFile = FILES.find((f) => f.key === "basic");
  try {
    await loadDataset(basicFile);
    state.currentDatasetKey = "basic";
    renderAutoStatus();
    populateDatasetSelects();
    initSheetSelect();
    runSimulation();
  } catch (e) {
    alert(`기본부문 불러오기 실패: ${e.message}`);
  }
});

document.getElementById("btn-reload-all").addEventListener("click", () => autoLoadAll(state.currentDatasetKey));
document.getElementById("dataset-select").addEventListener("change", (e) => {
  state.currentDatasetKey = e.target.value;
  document.getElementById("chain-dataset").value = e.target.value;
  document.getElementById("shock-dataset").value = e.target.value;
  initSheetSelect();
  runSimulation();
});
document.getElementById("sheet-select").addEventListener("change", populateRowSelect);
document.getElementById("row-search").addEventListener("input", populateRowSelect);
document.getElementById("row-select").addEventListener("change", renderRawConnections);
document.getElementById("col-search").addEventListener("input", renderRawConnections);
document.getElementById("sim-search").addEventListener("input", populateSimSectors);
document.getElementById("btn-add-shock").addEventListener("click", addShockItem);
document.getElementById("btn-run-sim").addEventListener("click", runSimulation);
document.getElementById("btn-clear-shock").addEventListener("click", () => {
  state.shockItems = [];
  renderBasket();
  runSimulation();
});
document.getElementById("chain-dataset").addEventListener("change", populateChainSectors);
document.getElementById("chain-sector").addEventListener("change", runChainAnalysis);
document.getElementById("btn-run-chain").addEventListener("click", runChainAnalysis);
document.getElementById("shock-dataset").addEventListener("change", populateShockSectors);
document.getElementById("shock-sector").addEventListener("change", runShockSimulation);
document.getElementById("btn-run-shock").addEventListener("click", runShockSimulation);

window.addEventListener("DOMContentLoaded", () => autoLoadAll("basic"));
