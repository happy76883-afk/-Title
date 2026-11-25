const WORKER_MSG = {
  LOAD: "load_excel",
  RUN: "run_analysis",
  LOADED: ["loaded", "excel_loaded", "ready"],
  VALIDATED: ["validated", "validation_result"],
  ANALYZED: ["analyzed", "analysis_result", "result"],
  ERROR: ["error", "load_error", "validation_error", "analysis_error"],
};

const els = {
  fileInput: document.getElementById("excelFileInput"),
  loadBtn: document.getElementById("loadBtn"),
  runBtn: document.getElementById("runBtn"),
  sheetSelect: document.getElementById("sheetSelect"),
  sectorSelect: document.getElementById("sectorSelect"),
  shockInput: document.getElementById("shockInput"),
  status: document.getElementById("status"),
  validation: document.getElementById("validation"),
  result: document.getElementById("result"),
};

let selectedFile = null;
let isDataReady = false;

const worker = new Worker("./worker.js", { type: "module" });

function setStatus(text) {
  els.status.textContent = text;
}

function setValidation(text) {
  els.validation.textContent = text;
}

function setResult(content) {
  if (typeof content === "string") {
    els.result.textContent = content;
  } else {
    els.result.textContent = JSON.stringify(content, null, 2);
  }
}

function populateSelect(selectEl, items = []) {
  selectEl.innerHTML = "";
  items.forEach((item, idx) => {
    const opt = document.createElement("option");
    if (typeof item === "object" && item !== null) {
      opt.value = item.value ?? item.code ?? item.id ?? idx;
      opt.textContent = item.label ?? item.name ?? String(opt.value);
    } else {
      opt.value = item;
      opt.textContent = String(item);
    }
    selectEl.appendChild(opt);
  });
  selectEl.disabled = items.length === 0;
}

function setAnalysisEnabled(enabled) {
  isDataReady = enabled;
  els.sheetSelect.disabled = !enabled || els.sheetSelect.options.length === 0;
  els.sectorSelect.disabled = !enabled || els.sectorSelect.options.length === 0;
  els.shockInput.disabled = !enabled;
  els.runBtn.disabled = !enabled;
}

els.fileInput.addEventListener("change", (event) => {
  const file = event.target.files?.[0] ?? null;
  selectedFile = file;
  setAnalysisEnabled(false);
  setValidation("");
  setResult("");

  if (!file) {
    setStatus("대기 중 · .xlsx 파일을 선택해 주세요.");
    return;
  }

  const sizeMB = (file.size / (1024 * 1024)).toFixed(2);
  setStatus(`선택됨: ${file.name} (${sizeMB} MB) · [불러오기 및 검증]을 눌러주세요.`);
});

async function handleLoadExcel() {
  const file = selectedFile || els.fileInput.files?.[0];
  if (!file) {
    setStatus("먼저 로컬에서 기본부문 .xlsx 파일을 선택해 주세요.");
    return;
  }

  try {
    setAnalysisEnabled(false);
    setValidation("");
    setResult("");
    els.loadBtn.disabled = true;
    setStatus("로컬 파일 읽는 중…");

    const buffer = await file.arrayBuffer();
    if (!buffer || buffer.byteLength < 100) {
      throw new Error(`파일 크기가 너무 작습니다 (${buffer?.byteLength ?? 0} bytes). 정상 .xlsx 파일인지 확인해 주세요.`);
    }

    setStatus("Excel 읽기와 계수 검증 중…");

    // worker.js가 buffer 또는 bytes 중 어느 키를 읽어도 대응할 수 있게 전달
    worker.postMessage(
      {
        type: WORKER_MSG.LOAD,
        fileName: file.name,
        buffer,
        bytes: new Uint8Array(buffer.slice(0)),
      },
      [buffer]
    );
  } catch (err) {
    els.loadBtn.disabled = false;
    setStatus(`불러오기 실패: ${err?.message || err}`);
  }
}

function handleRunAnalysis() {
  if (!isDataReady) {
    setStatus("먼저 기본부문 엑셀 파일을 불러와 주세요.");
    return;
  }

  setStatus("분석 계산 중…");
  worker.postMessage({
    type: WORKER_MSG.RUN,
    sheet: els.sheetSelect.value,
    sector: els.sectorSelect.value,
    shock: Number(els.shockInput.value),
  });
}

worker.addEventListener("message", (event) => {
  const msg = event.data || {};
  const type = msg.type || msg.event || "";

  if (WORKER_MSG.ERROR.includes(type) || msg.error) {
    els.loadBtn.disabled = false;
    setAnalysisEnabled(false);
    const errText = msg.error?.message || msg.error || msg.message || "알 수 없는 워커 오류";
    setStatus(`오류 발생: ${errText}`);
    if (msg.validation) {
      setValidation(typeof msg.validation === "string" ? msg.validation : JSON.stringify(msg.validation, null, 2));
    }
    return;
  }

  if (msg.status) {
    setStatus(msg.status);
  }

  if (msg.validation) {
    setValidation(typeof msg.validation === "string" ? msg.validation : JSON.stringify(msg.validation, null, 2));
  }

  if (WORKER_MSG.LOADED.includes(type) || WORKER_MSG.VALIDATED.includes(type) || msg.sectors || msg.sheets) {
    els.loadBtn.disabled = false;

    if (Array.isArray(msg.sheets) && msg.sheets.length > 0) {
      populateSelect(els.sheetSelect, msg.sheets);
    }
    if (Array.isArray(msg.sectors) && msg.sectors.length > 0) {
      populateSelect(els.sectorSelect, msg.sectors);
    }

    const sectorCount = msg.sectorCount ?? msg.sectors?.length ?? "";
    setStatus(sectorCount ? `준비 완료 · ${sectorCount}개 부문` : "준비 완료");
    setAnalysisEnabled(true);
    return;
  }

  if (WORKER_MSG.ANALYZED.includes(type) || msg.result !== undefined || msg.data !== undefined) {
    setStatus("분석 완료");
    setResult(msg.result ?? msg.data ?? msg);
  }
});

worker.addEventListener("error", (err) => {
  els.loadBtn.disabled = false;
  setAnalysisEnabled(false);
  setStatus(`워커 실행 오류: ${err.message || "worker.js를 확인해 주세요."}`);
});

els.loadBtn.addEventListener("click", handleLoadExcel);
els.runBtn.addEventListener("click", handleRunAnalysis);
