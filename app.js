const BASIC_FILE = "(표)(2023연장)투입산출표_생산자가격_기본부문.xlsx";

let pyodide = null;
let sheetData = null;
let industries = [];
let shocks = [];

const $ = (id) => document.getElementById(id);

async function boot() {
  pyodide = await loadPyodide();
  await pyodide.loadPackage(["openpyxl", "numpy"]);
  const code = await (await fetch("model.py")).text();
  await pyodide.runPythonAsync(code);
  $("statusText").textContent = "준비 완료. '기본부문 불러오기'를 누르세요.";
  $("loadBasicBtn").disabled = false;
}

async function loadBasic() {
  $("loadBasicBtn").disabled = true;
  $("statusText").textContent = "기본부문 엑셀 불러오는 중...";

  const res = await fetch(encodeURI(BASIC_FILE));
  const bytes = new Uint8Array(await res.arrayBuffer());
  pyodide.globals.set("js_bytes", bytes);
  const meta = JSON.parse(await pyodide.runPythonAsync("load_basic_workbook(js_bytes)"));

  $("sheetSelect").innerHTML = meta.sheets.map((s) => `<option value="${s}">${s}</option>`).join("");
  ["sheetSelect", "rowSearch", "rowSelect", "colSearch", "simSearch", "simIndustrySelect", "simAmount", "addShockBtn", "resetShockBtn"]
    .forEach((id) => { $(id).disabled = false; });

  industries = JSON.parse(await pyodide.runPythonAsync("get_industries()"));
  renderIndustries("");

  await onSheetChange();
  $("statusText").textContent = `기본부문 로딩 완료 (${meta.sheets.length}개 시트)`;
  $("loadBasicBtn").disabled = false;
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
