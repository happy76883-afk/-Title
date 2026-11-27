const FILES = [
  { key: "basic", label: "기본부문", name: "(표)(2023연장)투입산출표_생산자가격_기본부문.xlsx" },
  { key: "major", label: "통합대분류", name: "(표)(2023연장)투입산출표_생산자가격_통합대분류.xlsx" },
  { key: "medium", label: "통합중분류", name: "(표)(2023연장)투입산출표_생산자가격_통합중분류.xlsx" },
  { key: "minor", label: "통합소분류", name: "(표)(2023연장)투입산출표_생산자가격_통합소분류.xlsx" }
];

const state = {
  workbooks: {}, // key -> { label, sheetNames, sheets: { sheetName: parsed } }
  currentDatasetKey: "basic",
  shockItems: [] // { key, datasetLabel, sectorIdx, sectorCode, sectorName, amount }
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

/**
 * 한국은행 투입산출표 시트 파서
 * - 코드 행과 부문명 행을 분리해 실제 한글/영문 부문명을 열 이름·행 이름으로 매핑합니다.
 */
function parseSheetData(ws) {
  const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (!Array.isArray(raw) || raw.length === 0) {
    return { sheetTitle: "", colCodes: [], colNames: [], allRows: [], sectors: [], matrix: [] };
  }

  const maxCols = Math.max(...raw.map((r) => (Array.isArray(r) ? r.length : 0)), 1);

  // 상단 25행에서 코드 헤더 행과 부문명/항목명 헤더 행 찾기
  let codeRowIdx = -1;
  let nameRowIdx = -1;
  let bestNameScore = -1;

  for (let r = 0; r < Math.min(raw.length, 25); r++) {
    const row = raw[r] || [];
    const codeCount = row.filter((c) => isCodeToken(c)).length;
    const labelCells = row.filter((c) => isLabelText(c));
    const kwBonus = row.some((c) =>
      /(부문|합계|중간투입|부가가치|최종수요|총산출|수입|수출|소비|투자|피용자)/.test(String(c ?? ""))
    )
      ? 6
      : 0;

    if (codeCount >= 3 && codeRowIdx === -1) {
      codeRowIdx = r;
    }

    const score = labelCells.length + kwBonus;
    if (labelCells.length >= 2 && score > bestNameScore) {
      bestNameScore = score;
      nameRowIdx = r;
    }
  }

  if (nameRowIdx === -1) nameRowIdx = codeRowIdx !== -1 ? codeRowIdx : 0;

  const codeRow = codeRowIdx !== -1 ? raw[codeRowIdx] || [] : [];
  const nameRow = raw[nameRowIdx] || [];
  const subNameRow =
    nameRowIdx + 1 < raw.length &&
    (raw[nameRowIdx + 1] || []).filter((c) => isLabelText(c)).length >= 3
      ? raw[nameRowIdx + 1]
      : [];

  const colCodes = Array.from({ length: maxCols }, (_, c) => {
    const cd = String(codeRow[c] ?? "").trim();
    return cd || `열${c + 1}`;
  });

  const colNames = Array.from({ length: maxCols }, (_, c) => {
    const n1 = String(nameRow[c] ?? "").trim();
    const n2 = String(subNameRow[c] ?? "").trim();
    if (isLabelText(n1)) return n1;
    if (isLabelText(n2)) return n2;
    return n1 || n2 || colCodes[c];
  });

  const dataStartRow = Math.max(codeRowIdx, nameRowIdx, subNameRow.length ? nameRowIdx + 1 : 0) + 1;
  const allRows = [];

  for (let r = dataStartRow; r < raw.length; r++) {
    const row = raw[r];
    if (!Array.isArray(row)) continue;

    // 헤더 행 자체가 반복되는 경우 제외
    const sameHeader = row.filter((cell, c) => {
      const val = String(cell ?? "").trim();
      return val && (val === String(nameRow[c] ?? "").trim() || val === String(subNameRow[c] ?? "").trim());
    }).length;
    if (sameHeader >= 3) continue;

    // 앞쪽 5열에서 행 코드와 실제 행 이름 추출
    let rowCode = "";
    let rowName = "";

    for (let c = 0; c < Math.min(5, maxCols); c++) {
      const val = String(row[c] ?? "").trim();
      if (!val) continue;
      if (!rowCode && isCodeToken(val)) {
        rowCode = val;
      } else if (!rowName && isLabelText(val)) {
        rowName = val;
      }
    }

    if (!rowName) continue;

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

  // 시뮬레이션·연쇄효과용 내생 산업 행 추출(숫자 열이 5개 이상인 행)
  const sectorRows = allRows.filter((item) => item.cells.filter((cell) => cell.isNum).length >= 5);

  // 첫 번째 숫자 열 위치 기준으로 정방 행렬 추출
  const firstNumColIdx =
    sectorRows.length > 0 ? Math.max(0, sectorRows[0].cells.findIndex((cell) => cell.isNum)) : 2;

  const n = sectorRows.length;
  const sectors = sectorRows.map((item, idx) => ({
    idx,
    code: item.code || String(idx + 1).padStart(3, "0"),
    name: item.name
  }));

  const matrix = Array.from({ length: n }, (_, i) =>
    Array.from({ length: n }, (_, j) => {
      const cell = sectorRows[i].cells[firstNumColIdx + j];
      return cell ? cell.numVal : 0;
    })
  );

  return { colCodes, colNames, allRows, sectors, matrix };
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
  const parts = FILES.map((f) => {
    const ready = !!state.workbooks[f.key];
    return `${f.label}: <span class="${ready ? "ok" : "err"}">${ready ? "자동 로드 완료" : "미로드"}</span>`;
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
  sel.innerHTML = "";
  if (!active) return;

  const q = document.getElementById("row-search").value.trim().toLowerCase();
  const list = active.data.allRows.filter(
    (r) => !q || r.name.toLowerCase().includes(q) || r.code.toLowerCase().includes(q)
  );

  list.forEach((r, idx) => {
    const opt = document.createElement("option");
    opt.value = String(r.rowIdx);
    opt.textContent = `${r.code ? `[${r.code}] ` : ""}${r.name}`;
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

  const filteredCols = targetRow.cells.filter((cell) => {
    const combined = `${cell.colCode} ${cell.colName}`.toLowerCase();
    return !colQ || combined.includes(colQ);
  });

  tbody.innerHTML = filteredCols
    .slice(0, 200)
    .map((cell) => {
      const displayVal = cell.isNum
        ? cell.numVal.toLocaleString("ko-KR", { maximumFractionDigits: 4 })
        : String(cell.rawVal ?? "");
      return `
        <tr>
          <td>${cell.colCode}</td>
          <td>${cell.colName}</td>
          <td class="num">${displayVal}</td>
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

  const total = state.shockItems.reduce((sum, item) => sum + item.amount, 0);
  box.innerHTML =
    state.shockItems
      .map(
        (item, idx) => `
      <div class="basket-item">
        <span><strong>[${item.sectorCode}] ${item.sectorName}</strong> (${item.datasetLabel}) — +${item.amount.toLocaleString("ko-KR")} 백만원</span>
        <button class="del-btn" type="button" data-del="${idx}">삭제</button>
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
      <div class="k-val">${totalDirect.toLocaleString("ko-KR")} 백만원</div>
    </div>
    <div class="kpi">
      <div class="k-lbl">중간투입 간접 파급 합계</div>
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
        <td class="num">${m.direct.toLocaleString("ko-KR", { maximumFractionDigits: 1 })}</td>
        <td class="num">${m.indirect.toLocaleString("ko-KR", { maximumFractionDigits: 1 })}</td>
        <td class="num">${m.totalImpact.toLocaleString("ko-KR", { maximumFractionDigits: 1 })}</td>
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

  const sheet = ds.sheets[ds.sheetNames[0]];
  if (!sheet || !sheet.sectors.length) return;

  const targetIdx = Number(document.getElementById("chain-sector").value || 0);
  const n = sheet.sectors.length;
  const rowTotals = sheet.matrix.map((r) => r.reduce((a, b) => a + Math.max(0, b), 0) || 1);
  const colTotals = Array.from({ length: n }, (_, c) =>
    sheet.matrix.reduce((a, r) => a + Math.max(0, r[c] || 0), 0) || 1
  );

  // 전체 부문의 후방/전방 계수합 평균(영향력계수·감응도계수 산출용)
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
      <div class="k-val">${backwardSum.toFixed(4)}</div>
    </div>
    <div class="kpi">
      <div class="k-lbl">전방연쇄 계수합 (산출계수)</div>
      <div class="k-val">${forwardSum.toFixed(4)}</div>
    </div>
    <div class="kpi">
      <div class="k-lbl">영향력계수 (후방 상대지수)</div>
      <div class="k-val">${powerIdx.toFixed(3)}</div>
    </div>
    <div class="kpi">
      <div class="k-lbl">감응도계수 (전방 상대지수)</div>
      <div class="k-val">${sensIdx.toFixed(3)}</div>
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

function populateShockDropdown() {
  const key = document.getElementById("shock-dataset").value || state.currentDatasetKey;
  const ds = state.workbooks[key];
  const sel = document.getElementById("shock-sector");
  sel.innerHTML = "";
  if (!ds) return;

  const sheet = ds.sheets[ds.sheetNames[0]];
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

  const sheet = ds.sheets[ds.sheetNames[0]];
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
    return {
      name: sec.name,
      dImpact,
      sImpact,
      total: dImpact + sImpact,
      absScore: Math.abs(dImpact) + Math.abs(sImpact)
    };
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

function downloadTableCSV(tableId, fileName) {
  const table = document.getElementById(tableId);
  if (!table) return;
  const rows = [...table.querySelectorAll("tr")].map((tr) =>
    [...tr.querySelectorAll("th, td")].map((cell) => {
      const txt = cell.textContent.trim().replace(/"/g, '""');
      return `"${txt}"`;
    })
  );
  const csv = "\uFEFF" + rows.map((r) => r.join(",")).join("\n");
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = fileName;
  a.click();
  URL.revokeObjectURL(url);
}

async function loadOneFile(fileObj) {
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
    sheetNames: wb.SheetNames,
    sheets
  };
}

async function autoLoadAll(preferredKey = "basic") {
  for (const f of FILES) {
    try {
      await loadOneFile(f);
    } catch (_) {
      // 개별 파일 실패 시 상태 표시만 갱신
    }
    renderAutoStatus();
  }

  if (state.workbooks[preferredKey]) {
    state.currentDatasetKey = preferredKey;
  } else {
    const first = FILES.find((f) => state.workbooks[f.key]);
    if (first) state.currentDatasetKey = first.key;
  }

  syncDatasetDropdowns();
  populateSheetDropdown();
  runSimulation();
}

document.getElementById("btn-load-basic").addEventListener("click", async () => {
  const basic = FILES.find((f) => f.key === "basic");
  try {
    await loadOneFile(basic);
    state.currentDatasetKey = "basic";
    renderAutoStatus();
    syncDatasetDropdowns();
    populateSheetDropdown();
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
  populateSheetDropdown();
  runSimulation();
});

document.getElementById("sheet-select").addEventListener("change", populateRowDropdown);
document.getElementById("row-search").addEventListener("input", populateRowDropdown);
document.getElementById("row-select").addEventListener("change", renderRawTable);
document.getElementById("col-search").addEventListener("input", renderRawTable);

document.getElementById("sim-search").addEventListener("input", populateSimDropdown);
document.getElementById("btn-add-shock").addEventListener("click", addShockItem);
document.getElementById("btn-run-sim").addEventListener("click", runSimulation);
document.getElementById("btn-clear-shock").addEventListener("click", () => {
  state.shockItems = [];
  renderBasket();
  runSimulation();
});

document.getElementById("chain-dataset").addEventListener("change", populateChainDropdown);
document.getElementById("chain-sector").addEventListener("change", runChainAnalysis);
document.getElementById("btn-run-chain").addEventListener("click", runChainAnalysis);

document.getElementById("shock-dataset").addEventListener("change", populateShockDropdown);
document.getElementById("shock-sector").addEventListener("change", runShockSimulation);
document.getElementById("btn-run-shock").addEventListener("click", runShockSimulation);

document.getElementById("btn-csv-sim").addEventListener("click", () => downloadTableCSV("sim-table", "최종수요_파급효과.csv"));
document.getElementById("btn-csv-chain").addEventListener("click", () => downloadTableCSV("chain-table", "전후방_연쇄효과.csv"));
document.getElementById("btn-csv-shock").addEventListener("click", () => downloadTableCSV("shock-table", "공급수요_충격파급.csv"));

window.addEventListener("DOMContentLoaded", () => autoLoadAll("basic"));
