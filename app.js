const FILES = [
  { key: "basic", label: "기본부문", name: "(표)(2023연장)투입산출표_생산자가격_기본부문.xlsx" },
  { key: "major", label: "통합대분류", name: "(표)(2023연장)투입산출표_생산자가격_통합대분류.xlsx" },
  { key: "medium", label: "통합중분류", name: "(표)(2023연장)투입산출표_생산자가격_통합중분류.xlsx" },
  { key: "minor", label: "통합소분류", name: "(표)(2023연장)투입산출표_생산자가격_통합소분류.xlsx" }
];

const store = {
  loaded: {},
  logs: []
};

function addLog(text) {
  const now = new Date().toLocaleTimeString("ko-KR");
  store.logs.push(`[${now}] ${text}`);
  const box = document.getElementById("log-box");
  if (box) {
    box.textContent = store.logs.join("\n");
    box.scrollTop = box.scrollHeight;
  }
}

function toNum(v) {
  if (typeof v === "number") return Number.isFinite(v) ? v : 0;
  if (v == null) return 0;
  const s = String(v).replace(/,/g, "").trim();
  if (!s || s === "-" || s === "·") return 0;
  const n = Number(s);
  return Number.isFinite(n) ? n : 0;
}

function parseWorkbook(wb, label) {
  let best = null;

  for (const sheetName of wb.SheetNames) {
    const ws = wb.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
    if (!Array.isArray(rows) || rows.length < 5) continue;

    const candidates = [];
    for (let r = 0; r < rows.length; r++) {
      const row = rows[r];
      if (!Array.isArray(row) || row.length < 4) continue;

      const sectorName = row.find(
        (cell, cIdx) =>
          cIdx < 4 &&
          typeof cell === "string" &&
          cell.trim().length > 0 &&
          !/^-?\d+(\.\d+)?$/.test(cell.trim())
      );

      const nums = [];
      for (let c = 1; c < row.length; c++) {
        const cell = row[c];
        if (typeof cell === "number" && Number.isFinite(cell)) {
          nums.push(cell);
        } else if (typeof cell === "string") {
          const cleaned = cell.replace(/,/g, "").trim();
          if (/^-?\d+(\.\d+)?$/.test(cleaned)) {
            nums.push(Number(cleaned));
          }
        }
      }

      if (sectorName && nums.length >= 3) {
        candidates.push({ name: String(sectorName).trim(), nums });
      }
    }

    if (candidates.length >= 2) {
      const n = candidates.length;
      const sectors = candidates.map((d) => d.name);
      const matrix = Array.from({ length: n }, (_, i) =>
        Array.from({ length: n }, (_, j) => (candidates[i].nums[j] !== undefined ? candidates[i].nums[j] : 0))
      );
      if (!best || n > best.sectors.length) {
        best = { sheetName, sectors, matrix };
      }
    }
  }

  if (!best) {
    throw new Error(`${label} 시트에서 숫자 행렬을 찾지 못했습니다.`);
  }

  addLog(`${label}: "${best.sheetName}" 시트 자동 인식 (${best.sectors.length}개 부문)`);
  return best;
}

function renderStatus() {
  const grid = document.getElementById("status-grid");
  grid.innerHTML = "";

  for (const f of FILES) {
    const item = store.loaded[f.key];
    const card = document.createElement("div");
    card.className = "status-card";
    card.innerHTML = `
      <strong>${f.label} — ${item ? `<span class="ok">자동 로드 완료 (${item.sectors.length})</span>` : `<span class="err">로드 실패</span>`}</strong>
      <span class="fname">data/${f.name}</span>
    `;
    grid.appendChild(card);
  }
}

function refreshDropdowns() {
  const levelSel = document.getElementById("level-select");
  const sectorSel = document.getElementById("sector-select");
  levelSel.innerHTML = "";

  const readyFiles = FILES.filter((f) => store.loaded[f.key]);
  for (const f of readyFiles) {
    const opt = document.createElement("option");
    opt.value = f.key;
    opt.textContent = `${f.label} (${store.loaded[f.key].sectors.length}개)`;
    levelSel.appendChild(opt);
  }

  const fillSectors = () => {
    sectorSel.innerHTML = "";
    const cur = store.loaded[levelSel.value];
    if (!cur) return;
    cur.sectors.forEach((name, idx) => {
      const opt = document.createElement("option");
      opt.value = String(idx);
      opt.textContent = `${idx + 1}. ${name}`;
      sectorSel.appendChild(opt);
    });
  };

  levelSel.onchange = () => {
    fillSectors();
    calculate();
  };
  sectorSel.onchange = calculate;

  fillSectors();
}

function calculate() {
  const key = document.getElementById("level-select").value;
  const idx = Number(document.getElementById("sector-select").value || 0);
  const dPct = toNum(document.getElementById("demand-shock").value);
  const sPct = toNum(document.getElementById("supply-shock").value);

  const cur = store.loaded[key];
  const kpiGrid = document.getElementById("kpi-grid");
  const tbody = document.querySelector("#result-table tbody");

  if (!cur || !cur.sectors.length) {
    kpiGrid.innerHTML = `<div class="kpi"><div class="label">상태</div><div class="val">데이터 없음</div></div>`;
    tbody.innerHTML = "";
    return;
  }

  const n = cur.sectors.length;
  const rowSum = cur.matrix.map((r) => r.reduce((a, b) => a + Math.max(0, b), 0) || 1);
  const colSum = Array.from({ length: n }, (_, c) =>
    cur.matrix.reduce((a, r) => a + Math.max(0, r[c] || 0), 0) || 1
  );

  const aCol = Array.from({ length: n }, (_, i) => (cur.matrix[i]?.[idx] || 0) / colSum[idx]);
  const bRow = Array.from({ length: n }, (_, j) => (cur.matrix[idx]?.[j] || 0) / rowSum[idx]);

  const base = rowSum[idx];
  const backwardSum = aCol.reduce((a, b) => a + b, 0);
  const forwardSum = bRow.reduce((a, b) => a + b, 0);

  const demandTotal = base * (dPct / 100) * (1 + backwardSum);
  const supplyTotal = base * (sPct / 100) * (1 + forwardSum);

  kpiGrid.innerHTML = `
    <div class="kpi">
      <div class="label">선택 부문</div>
      <div class="val" style="font-size:15px;">${cur.sectors[idx]}</div>
    </div>
    <div class="kpi">
      <div class="label">후방연쇄 계수합</div>
      <div class="val">${backwardSum.toFixed(4)}</div>
    </div>
    <div class="kpi">
      <div class="label">수요충격 총파급</div>
      <div class="val">${demandTotal.toLocaleString("ko-KR", { maximumFractionDigits: 1 })}</div>
    </div>
    <div class="kpi">
      <div class="label">공급충격 총파급</div>
      <div class="val">${supplyTotal.toLocaleString("ko-KR", { maximumFractionDigits: 1 })}</div>
    </div>
  `;

  const ranked = cur.sectors
    .map((name, i) => {
      const aIj = aCol[i] || 0;
      const bIj = bRow[i] || 0;
      const dVal = base * (dPct / 100) * aIj;
      const sVal = base * (sPct / 100) * bIj;
      return { name, aIj, bIj, dVal, sVal, mag: Math.abs(dVal) + Math.abs(sVal) };
    })
    .sort((a, b) => b.mag - a.mag)
    .slice(0, 20);

  tbody.innerHTML = ranked
    .map(
      (r, rank) => `
      <tr>
        <td>${rank + 1}</td>
        <td>${r.name}</td>
        <td class="num">${r.aIj.toFixed(4)}</td>
        <td class="num">${r.bIj.toFixed(4)}</td>
        <td class="num">${r.dVal.toLocaleString("ko-KR", { maximumFractionDigits: 1 })}</td>
        <td class="num">${r.sVal.toLocaleString("ko-KR", { maximumFractionDigits: 1 })}</td>
      </tr>
    `
    )
    .join("");
}

async function autoLoadAll() {
  store.loaded = {};
  store.logs = [];
  addLog("data/ 폴더 자동 읽기 시작...");

  for (const f of FILES) {
    const url = encodeURI(`./data/${f.name}`);
    try {
      const res = await fetch(url);
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const buf = await res.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });
      store.loaded[f.key] = parseWorkbook(wb, f.label);
    } catch (err) {
      addLog(`[실패] ${f.label} (${f.name}): ${err.message}`);
    }
  }

  renderStatus();
  refreshDropdowns();
  calculate();
}

document.getElementById("reload-btn").addEventListener("click", autoLoadAll);
document.getElementById("calc-btn").addEventListener("click", calculate);
window.addEventListener("DOMContentLoaded", autoLoadAll);
