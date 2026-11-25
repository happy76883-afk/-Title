const DATA_FILES = [
  {
    key: "basic",
    label: "기본부문",
    fileName: "(표)(2023연장)투입산출표_생산자가격_기본부문.xlsx"
  },
  {
    key: "major",
    label: "통합대분류",
    fileName: "(표)(2023연장)투입산출표_생산자가격_통합대분류.xlsx"
  },
  {
    key: "medium",
    label: "통합중분류",
    fileName: "(표)(2023연장)투입산출표_생산자가격_통합중분류.xlsx"
  },
  {
    key: "minor",
    label: "통합소분류",
    fileName: "(표)(2023연장)투입산출표_생산자가격_통합소분류.xlsx"
  }
];

const state = {
  datasets: {},
  logs: []
};

function logMsg(msg) {
  const time = new Date().toLocaleTimeString("ko-KR");
  state.logs.push(`[${time}] ${msg}`);
  const el = document.getElementById("debug-log");
  if (el) {
    el.textContent = state.logs.join("\n");
    el.scrollTop = el.scrollHeight;
  }
}

function parseNumber(val) {
  if (typeof val === "number") return Number.isFinite(val) ? val : 0;
  if (val == null) return 0;
  const cleaned = String(val).replace(/,/g, "").trim();
  if (!cleaned || cleaned === "-" || cleaned === "·") return 0;
  const num = Number(cleaned);
  return Number.isFinite(num) ? num : 0;
}

/**
 * 한국은행 투입산출표 엑셀에서 정방형 산업 매트릭스를 자동 추출합니다.
 * 상단 제목/단위 행을 건너뛰고 숫자 밀도가 높은 구간을 탐색합니다.
 */
function extractIOMatrix(workbook, fileLabel) {
  let bestResult = null;

  for (const sheetName of workbook.SheetNames) {
    const ws = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
    if (!rows || rows.length < 10) continue;

    // 각 행에서 숫자로 변환 가능한 셀 개수를 세어 데이터 행 후보를 찾음
    const numericRows = [];
    for (let r = 0; r < rows.length; r++) {
      const row = rows[r];
      if (!Array.isArray(row) || row.length < 5) continue;

      // 산업명 후보(문자열 셀) 찾기
      const nameCell = row.find((cell, idx) => idx < 4 && typeof cell === "string" && cell.trim().length > 0 && !/^\d+$/.test(cell.trim()));
      const nums = [];
      for (let c = 1; c < row.length; c++) {
        const v = row[c];
        if (typeof v === "number" && Number.isFinite(v)) {
          nums.push({ col: c, val: v });
        } else if (typeof v === "string") {
          const t = v.replace(/,/g, "").trim();
          if (/^-?\d+(\.\d+)?$/.test(t)) {
            nums.push({ col: c, val: Number(t) });
          }
        }
      }

      if (nameCell && nums.length >= 4) {
        numericRows.push({
          rowIdx: r,
          name: String(nameCell).trim(),
          nums
        });
      }
    }

    if (numericRows.length >= 3) {
      // 연속된 블록 중 가장 큰 블록 선택
      const n = numericRows.length;
      const sectors = numericRows.slice(0, n).map((item) => item.name);
      const matrix = Array.from({ length: n }, () => Array(n).fill(0));

      for (let i = 0; i < n; i++) {
        const rowNums = numericRows[i].nums;
        for (let j = 0; j < Math.min(n, rowNums.length); j++) {
          matrix[i][j] = rowNums[j].val;
        }
      }

      if (!bestResult || n > bestResult.sectors.length) {
        bestResult = { sheetName, sectors, matrix };
      }
    }
  }

  if (!bestResult) {
    throw new Error(`${fileLabel}: 유효한 숫자 행렬 시트를 찾지 못했습니다.`);
  }

  logMsg(`${fileLabel} -> 시트 "${bestResult.sheetName}"에서 ${bestResult.sectors.length}개 부문 추출`);
  return bestResult;
}

function renderFileStatus() {
  const container = document.getElementById("file-status");
  container.innerHTML = "";

  for (const item of DATA_FILES) {
    const ds = state.datasets[item.key];
    const div = document.createElement("div");
    div.className = "status-item";
    const statusHtml = ds
      ? `<span class="badge-ok">로드 완료 (${ds.sectors.length}개 부문)</span>`
      : `<span class="badge-err">로드 실패 / 대기</span>`;
    div.innerHTML = `
      <span class="label">${item.label} ${statusHtml}</span>
      <span class="path">data/${item.fileName}</span>
    `;
    container.appendChild(div);
  }
}

function populateSelectors() {
  const dsSelect = document.getElementById("dataset-select");
  const sectorSelect = document.getElementById("sector-select");
  dsSelect.innerHTML = "";

  const available = DATA_FILES.filter((f) => state.datasets[f.key]);
  for (const f of available) {
    const opt = document.createElement("option");
    opt.value = f.key;
    opt.textContent = `${f.label} (${state.datasets[f.key].sectors.length}개)`;
    dsSelect.appendChild(opt);
  }

  const updateSectors = () => {
    sectorSelect.innerHTML = "";
    const currentKey = dsSelect.value;
    const ds = state.datasets[currentKey];
    if (!ds) return;
    ds.sectors.forEach((name, idx) => {
      const opt = document.createElement("option");
      opt.value = String(idx);
      opt.textContent = `${idx + 1}. ${name}`;
      sectorSelect.appendChild(opt);
    });
  };

  dsSelect.onchange = updateSectors;
  updateSectors();
}

async function loadAllData() {
  state.datasets = {};
  state.logs = [];
  logMsg("data/ 폴더의 4개 엑셀 파일 로드를 시작합니다.");

  for (const item of DATA_FILES) {
    const url = encodeURI(`./data/${item.fileName}`);
    try {
      const res = await fetch(url);
      if (!res.ok) {
        throw new Error(`HTTP ${res.status}`);
      }
      const buf = await res.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });
      state.datasets[item.key] = extractIOMatrix(wb, item.label);
    } catch (err) {
      logMsg(`[오류] ${item.label} (${item.fileName}): ${err.message}`);
    }
  }

  renderFileStatus();
  populateSelectors();
  runAnalysis();
}

function runAnalysis() {
  const dsKey = document.getElementById("dataset-select").value;
  const targetIdx = Number(document.getElementById("sector-select").value || 0);
  const demandShockPct = parseNumber(document.getElementById("shock-pct").value);
  const supplyShockPct = parseNumber(document.getElementById("supply-shock-pct").value);

  const ds = state.datasets[dsKey];
  const kpiArea = document.getElementById("kpi-area");
  const tbody = document.querySelector("#result-table tbody");

  if (!ds || !ds.sectors.length) {
    kpiArea.innerHTML = `<div class="kpi"><div class="k-title">상태</div><div class="k-val">데이터 없음</div></div>`;
    tbody.innerHTML = "";
    return;
  }

  const n = ds.sectors.length;
  // 행 i의 총투입/총산출 근사치
  const rowTotals = ds.matrix.map((r) => r.reduce((acc, v) => acc + Math.max(0, v), 0) || 1);
  const colTotals = Array.from({ length: n }, (_, c) =>
    ds.matrix.reduce((acc, r) => acc + Math.max(0, r[c] || 0), 0) || 1
  );

  // 전방연쇄(산출계수 b_ij = z_ij / X_i) 및 후방연쇄(투입계수 a_ij = z_ij / X_j)
  const forwardLinkage = Array.from({ length: n }, (_, j) =>
    (ds.matrix[targetIdx]?.[j] || 0) / rowTotals[targetIdx]
  );
  const backwardLinkage = Array.from({ length: n }, (_, i) =>
    (ds.matrix[i]?.[targetIdx] || 0) / colTotals[targetIdx]
  );

  const baseOutput = rowTotals[targetIdx];
  const demandImpactTotal = baseOutput * (demandShockPct / 100) * (1 + backwardLinkage.reduce((a, b) => a + b, 0));
  const supplyImpactTotal = baseOutput * (supplyShockPct / 100) * (1 + forwardLinkage.reduce((a, b) => a + b, 0));

  kpiArea.innerHTML = `
    <div class="kpi">
      <div class="k-title">선택 부문</div>
      <div class="k-val" style="font-size:15px;">${ds.sectors[targetIdx]}</div>
    </div>
    <div class="kpi">
      <div class="k-title">후방연쇄 계수합</div>
      <div class="k-val">${backwardLinkage.reduce((a, b) => a + b, 0).toFixed(4)}</div>
    </div>
    <div class="kpi">
      <div class="k-title">수요충격 총파급 추정</div>
      <div class="k-val">${demandImpactTotal.toLocaleString("ko-KR", { maximumFractionDigits: 1 })}</div>
    </div>
    <div class="kpi">
      <div class="k-title">공급충격 총파급 추정</div>
      <div class="k-val">${supplyImpactTotal.toLocaleString("ko-KR", { maximumFractionDigits: 1 })}</div>
    </div>
  `;

  const rowsRanked = ds.sectors
    .map((name, idx) => {
      const aIj = backwardLinkage[idx] || 0;
      const bIj = forwardLinkage[idx] || 0;
      const dShock = baseOutput * (demandShockPct / 100) * aIj;
      const sShock = baseOutput * (supplyShockPct / 100) * bIj;
      return { idx, name, aIj, bIj, dShock, sShock, score: Math.abs(dShock) + Math.abs(sShock) };
    })
    .sort((a, b) => b.score - a.score)
    .slice(0, 20);

  tbody.innerHTML = rowsRanked
    .map(
      (r, i) => `
      <tr>
        <td>${i + 1}</td>
        <td>${r.name}</td>
        <td class="num">${r.aIj.toFixed(4)}</td>
        <td class="num">${r.bIj.toFixed(4)}</td>
        <td class="num">${r.dShock.toLocaleString("ko-KR", { maximumFractionDigits: 1 })}</td>
        <td class="num">${r.sShock.toLocaleString("ko-KR", { maximumFractionDigits: 1 })}</td>
      </tr>
    `
    )
    .join("");
}

document.getElementById("btn-reload").addEventListener("click", loadAllData);
document.getElementById("btn-run").addEventListener("click", runAnalysis);
window.addEventListener("DOMContentLoaded", loadAllData);
