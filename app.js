const FILES = [
  { key: "basic", label: "기본부문", name: "(표)(2023연장)투입산출표_생산자가격_기본부문.xlsx" },
  { key: "major", label: "통합대분류", name: "(표)(2023연장)투입산출표_생산자가격_통합대분류.xlsx" },
  { key: "medium", label: "통합중분류", name: "(표)(2023연장)투입산출표_생산자가격_통합중분류.xlsx" },
  { key: "minor", label: "통합소분류", name: "(표)(2023연장)투입산출표_생산자가격_통합소분류.xlsx" }
];

const state = {
  workbooks: {}, // key -> { label, wb, sheets: { sheetName: { headers, rows, sectors } } }
  currentDatasetKey: "basic",
  shockItems: [] // { key, datasetLabel, sectorIdx, sectorName, amount }
};

function toNum(v) {
  if (typeof v === "number") return Number.isFinite(v) ? v : 0;
  if (v == null) return 0;
  const s = String(v).replace(/,/g, "").trim();
  if (!s || s === "-" || s === "·" || s === "…") return 0;
  const n = Number(s);
  return Number.isFinite(n) ? n : 0;
}

function parseSheetData(ws) {
  const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (!Array.isArray(raw) || raw.length === 0) {
    return { headers: [], allRows: [], sectors: [], matrix: [] };
  }

  // 숫자 셀이 가장 많이 등장하는 행을 헤더 행으로 추정
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

  // 행 이름(문자열)이 있는 모든 행(산업 행 + 합계/부가가치/지표 행 포함) 수집
  const allRows = [];
  for (let r = headerRowIdx + 1; r < raw.length; r++) {
    const row = raw[r];
    if (!Array.isArray(row)) continue;

    // 앞쪽 열(0~3)에서 행 이름 후보 찾기
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

  // 시뮬레이션용 산업 행렬 추출(숫자 열이 3개 이상인 행들)
  const numericRows = allRows.filter(
    (item) => item.values.filter((v) => typeof v.rawVal === "number" || /^-?\d[\d,.]*$/.test(String(v.rawVal).replace(/,/g, "").trim())).length >= 3
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
    return `${f.label}: <span class="${loaded ? "ok" : "err"}">${loaded ? "완료" : "미로드"}</span>`;
  });
  bar.innerHTML = parts.join(" &nbsp;|&nbsp; ");
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
    `<div style="margin-top:8px; font-weight:700; font-size:13.5px;">총 추가 수요 증가액: ${total.toLocaleString("ko-KR")} 백만원 (${(total / 100).toLocaleString("ko-KR")} 억원)</div>`;

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
  const supplyPct = toNum(document.getElementById("sim-supply-pct").value);

  if (!ds) {
    kpiEl.innerHTML = "";
    tbody.innerHTML = "";
    return;
  }

  const sheet = Object.values(ds.sheets)[0];
  if (!sheet || !sheet.sectors.length) return;

  const n = sheet.sectors.length;
  const rowTotals = sheet.matrix.map((r) => r.reduce((a, b) => a + Math.max(0, b), 0) || 1);
  const colTotals = Array.from({ length: n }, (_, c) =>
    sheet.matrix.reduce((a, r) => a + Math.max(0, r[c] || 0), 0) || 1
  );

  // 선택된 여러 산업의 수요 증가액을 합산
  const directShock = Array(n).fill(0);
  for (const item of state.shockItems) {
    if (item.key === state.currentDatasetKey && item.sectorIdx < n) {
      directShock[item.sectorIdx] += item.amount;
    }
  }

  const totalInputShock = directShock.reduce((a, b) => a + b, 0);

  const sectorMetrics = sheet.sectors.map((sec, i) => {
    const backwardSum = Array.from({ length: n }, (_, r) => (sheet.matrix[r]?.[i] || 0) / colTotals[i]).reduce(
      (a, b) => a + b,
      0
    );
    const forwardSum = Array.from({ length: n }, (_, c) => (sheet.matrix[i]?.[c] || 0) / rowTotals[i]).reduce(
      (a, b) => a + b,
      0
    );

    // 직접 충격 + 투입계수 기반 간접 파급 근사
    const indirectDemand = directShock.reduce((acc, shockVal, j) => {
      const aIj = (sheet.matrix[i]?.[j] || 0) / colTotals[j];
      return acc + shockVal * aIj;
    }, 0);

    const demandImpact = directShock[i] + indirectDemand;
    const supplyImpact = totalInputShock * (supplyPct / 100) * forwardSum;

    return {
      name: sec.name,
      demandImpact,
      backwardSum,
      forwardSum,
      supplyImpact
    };
  });

  const totalDemandImpact = sectorMetrics.reduce((a, m) => a + m.demandImpact, 0);
  const totalSupplyImpact = sectorMetrics.reduce((a, m) => a + m.supplyImpact, 0);

  kpiEl.innerHTML = `
    <div class="kpi">
      <div class="k-lbl">합산 수요 증가액</div>
      <div class="k-val">${totalInputShock.toLocaleString("ko-KR")} 백만원</div>
    </div>
    <div class="kpi">
      <div class="k-lbl">총 수요유발 파급(추정)</div>
      <div class="k-val">${totalDemandImpact.toLocaleString("ko-KR", { maximumFractionDigits: 1 })} 백만원</div>
    </div>
    <div class="kpi">
      <div class="k-lbl">총 공급충격 파급(추정)</div>
      <div class="k-val">${totalSupplyImpact.toLocaleString("ko-KR", { maximumFractionDigits: 1 })} 백만원</div>
    </div>
  `;

  const ranked = [...sectorMetrics]
    .sort((a, b) => Math.abs(b.demandImpact) + Math.abs(b.supplyImpact) - (Math.abs(a.demandImpact) + Math.abs(a.supplyImpact)))
    .slice(0, 25);

  tbody.innerHTML = ranked
    .map(
      (m, idx) => `
      <tr>
        <td>${idx + 1}</td>
        <td>${m.name}</td>
        <td class="num">${m.demandImpact.toLocaleString("ko-KR", { maximumFractionDigits: 1 })}</td>
        <td class="num">${m.backwardSum.toFixed(4)}</td>
        <td class="num">${m.forwardSum.toFixed(4)}</td>
        <td class="num">${m.supplyImpact.toLocaleString("ko-KR", { maximumFractionDigits: 1 })}</td>
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

  initSheetSelect();
  runSimulation();
}

document.getElementById("btn-load-basic").addEventListener("click", async () => {
  const basicFile = FILES.find((f) => f.key === "basic");
  try {
    await loadDataset(basicFile);
    state.currentDatasetKey = "basic";
    renderAutoStatus();
    initSheetSelect();
    runSimulation();
  } catch (e) {
    alert(`기본부문 불러오기 실패: ${e.message}`);
  }
});

document.getElementById("btn-reload-all").addEventListener("click", () => autoLoadAll(state.currentDatasetKey));
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

window.addEventListener("DOMContentLoaded", () => autoLoadAll("basic"));
