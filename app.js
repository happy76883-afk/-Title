const CANDIDATE_FILES = [
  "(표)(2023연장)투입산출표_생산자가격_기본부문.xlsx",
  "basic.xlsx"
];

let pyodide = null;
let sheetData = null;
let industries = [];
let shocks = [];

const $ = (id) => document.getElementById(id);

function setStatus(msg, isError = false) {
  const el = $("statusText");
  el.textContent = msg;
  el.style.color = isError ? "#b91c1c" : "#0b4a82";
}

async function boot() {
  try {
    setStatus("1/3 Pyodide 엔진 다운로드 중... (10~20초 소요)");
    if (typeof loadPyodide !== "function") {
      throw new Error("Pyodide 스크립트를 불러오지 못했습니다. 네트워크를 확인하세요.");
    }
    pyodide = await loadPyodide();

    setStatus("2/3 openpyxl, numpy 패키지 설치 중...");
    await pyodide.loadPackage(["numpy", "pandas"]);

    setStatus("3/3 model.py 불러오는 중...");
    const pyResp = await fetch("./model.py");
    if (!pyResp.ok) {
      throw new Error(`model.py 파일을 찾을 수 없습니다 (HTTP ${pyResp.status})`);
    }
    const code = await pyResp.text();
    await pyodide.runPythonAsync(code);

    setStatus("준비 완료! 오른쪽 위 '기본부문 불러오기' 버튼을 누르세요.");
    $("loadBasicBtn").disabled = false;
  } catch (err) {
    setStatus("초기화 오류: " + (err.message || err), true);
  }
}

async function fetchExcelBytes() {
  const tried = [];
  for (const name of CANDIDATE_FILES) {
    for (const url of ["./" + encodeURI(name), "./" + name]) {
      try {
        const res = await fetch(url);
        if (res.ok) {
          const buf = await res.arrayBuffer();
          return { name, bytes: new Uint8Array(buf) };
        }
        tried.push(`${name} (${res.status})`);
      } catch (e) {
        tried.push(`${name} (네트워크 오류)`);
      }
    }
  }
  throw new Error("엑셀 파일을 찾지 못했습니다: " + tried.join(", "));
}

async function loadBasic() {
  $("loadBasicBtn").disabled = true;
  try {
    setStatus("기본부문 엑셀 파일 불러오는 중...");
    const { name, bytes } = await fetchExcelBytes();

    setStatus(`'${name}' 분석 중...`);
    pyodide.globals.set("js_bytes", bytes);
    const meta = JSON.parse(await pyodide.runPythonAsync("load_basic_workbook(js_bytes)"));

    $("sheetSelect").innerHTML = meta.sheets.map((s) => `<option value="${s}">${s}</option>`).join("");
    ["sheetSelect", "rowSearch", "rowSelect", "colSearch", "simSearch", "simIndustrySelect", "simAmount", "addShockBtn", "resetShockBtn"]
      .forEach((id) => { $(id).disabled = false; });

    industries = JSON.parse(await pyodide.runPythonAsync("get_industries()"));
    renderIndustries("");

    await onSheetChange();
    setStatus(`연동 완료! '${name}' (시트 ${meta.sheets.length}개)`);
  } catch (err) {
    setStatus("엑셀 연동 실패: " + (err.message || err), true);
  } finally {
    $("loadBasicBtn").disabled = false;
  }
}

async function onSheetChange() {
  pyodide.globals.set("js_sheet", $("sheetSelect").value);
  sheetData = JSON.parse(await pyodide.runPythonAsync("get_sheet(js_sheet)"));
  $("rowSearch").value = "";
  $("colSearch").value = "";
  renderRows("");
  renderRawTable();
}

function renderRows(kw) {
  const q = kw.trim().toLowerCase();
  const list = sheetData.rows.filter((r) => !q || r.name.toLowerCase().includes(q) || r.code.toLowerCase().includes(q));
  $("rowSelect").innerHTML = list.map((r) => `<option value="${r.row_idx}">[${r.code}] ${r.name}</option>`).join("");
}

function renderRawTable() {
  const rIdx = Number($("rowSelect").value);
  const row = sheetData.rows.find((r) => r.row_idx === rIdx);
  if (!row) {
    $("emptyMsg").hidden = false;
    $("rawTableWrap").hidden = true;
    return;
  }
  const q = $("colSearch").value.trim().toLowerCase();
  const html = sheetData.cols
    .map((c, i) => ({ c, v: row.values[i] }))
    .filter(({ c }) => !q || c.name.toLowerCase().includes(q) || c.code.toLowerCase().includes(q))
    .map(({ c, v }) => {
      const valText = typeof v === "number"
        ? (Number.isInteger(v) ? v.toLocaleString("ko-KR") : v.toLocaleString("ko-KR", { maximumFractionDigits: 6 }))
        : (v ?? "-");
      return `<tr><td>${c.code}</td><td>${c.name}</td><td class="num">${valText}</td></tr>`;
    })
    .join("");

  $("emptyMsg").hidden = true;
  $("rawTableWrap").hidden = false;
  $("rawTableBody").innerHTML = html || `<tr><td colspan="3" class="num">일치하는 열이 없습니다.</td></tr>`;
}

function renderIndustries(kw) {
  const q = kw.trim().toLowerCase();
  const list = industries.filter((d) => !q || d.name.toLowerCase().includes(q) || d.code.toLowerCase().includes(q));
  $("simIndustrySelect").innerHTML = list.map((d) => `<option value="${d.code}">[${d.code}] ${d.name}</option>`).join("");
}

async function applyShock() {
  try {
    const code = $("simIndustrySelect").value;
    const amount = Number($("simAmount").value);
    if (!code || !Number.isFinite(amount) || amount === 0) return;

    const ind = industries.find((d) => d.code === code);
    const found = shocks.find((s) => s.code === code);
    if (found) found.amount += amount;
    else shocks.push({ code, name: ind ? ind.name : code, amount });

    $("shockChips").innerHTML = shocks
      .map((s) => `<span class="chip">[${s.code}] ${s.name} +${s.amount.toLocaleString("ko-KR")}백만원</span>`)
      .join("");

    pyodide.globals.set("js_shocks", JSON.stringify(shocks));
    const r = JSON.parse(await pyodide.runPythonAsync("simulate(js_shocks)"));
    $("simResultBox").hidden = false;
    $("simResultBox").innerHTML =
      `최종수요 증가: <b>${r.total_shock.toLocaleString("ko-KR")}</b> 백만원 · ` +
      `생산유발액: <b>${r.total_output.toLocaleString("ko-KR")}</b> 백만원 · ` +
      `부가가치유발액: <b>${r.total_va.toLocaleString("ko-KR")}</b> 백만원 · ` +
      `수입유발액: <b>${r.total_import.toLocaleString("ko-KR")}</b> 백만원`;
  } catch (err) {
    setStatus("시뮬레이션 오류: " + (err.message || err), true);
  }
}

$("loadBasicBtn").disabled = true;
$("loadBasicBtn").addEventListener("click", loadBasic);
$("sheetSelect").addEventListener("change", onSheetChange);
$("rowSearch").addEventListener("input", (e) => { renderRows(e.target.value); renderRawTable(); });
$("rowSelect").addEventListener("change", renderRawTable);
$("colSearch").addEventListener("input", renderRawTable);
$("simSearch").addEventListener("input", (e) => renderIndustries(e.target.value));
$("addShockBtn").addEventListener("click", applyShock);
$("resetShockBtn").addEventListener("click", () => {
  shocks = [];
  $("shockChips").innerHTML = "";
  $("simResultBox").hidden = true;
});

boot();
