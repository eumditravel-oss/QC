// app.js
let workbook = null;
let currentSheetName = "";
let currentRows = [];
let currentHeader = [];
let currentHeaderRowIndex = -1;
let currentFormulaColIndex = -1;
let currentResultColIndex = -1;
let lastFilteredRows = [];
let lastLimit = null;

const excelFile = document.getElementById("excelFile");
const sheetButtons = document.getElementById("sheetButtons");
const fileName = document.getElementById("fileName");
const selectedSheet = document.getElementById("selectedSheet");
const totalRows = document.getElementById("totalRows");
const warningRows = document.getElementById("warningRows");
const limitValue = document.getElementById("limitValue");
const manualCheckBtn = document.getElementById("manualCheckBtn");
const manualInputArea = document.getElementById("manualInputArea");
const manualLimit = document.getElementById("manualLimit");
const runManualCheck = document.getElementById("runManualCheck");
const resultMessage = document.getElementById("resultMessage");
const resultTable = document.getElementById("resultTable");
const showAllBtn = document.getElementById("showAllBtn");
const showWarningBtn = document.getElementById("showWarningBtn");

const targetSheetNames = ["부재별산출서", "아파트옹벽 Unit별산출서"];
const PAGE_SIZE_WARNING = 500;
const PAGE_SIZE_ALL = 500;

excelFile.addEventListener("change", handleFileUpload);

manualCheckBtn.addEventListener("click", () => {
  manualInputArea.classList.toggle("hidden");
});

runManualCheck.addEventListener("click", runManualErrorCheck);

showAllBtn.addEventListener("click", () => {
  renderTable(currentRows.slice(0, PAGE_SIZE_ALL), null, false);
  resultMessage.textContent = `전체 행 중 상위 ${PAGE_SIZE_ALL}건만 표시 중입니다. 속도 저하 방지를 위한 제한입니다.`;
});

showWarningBtn.addEventListener("click", () => {
  if (!lastFilteredRows.length) {
    resultMessage.textContent = "표시할 경고 행이 없습니다.";
    renderTable([], lastLimit, true);
    return;
  }

  renderTable(lastFilteredRows.slice(0, PAGE_SIZE_WARNING), lastLimit, true);

  resultMessage.textContent =
    lastFilteredRows.length > PAGE_SIZE_WARNING
      ? `경고 행 ${lastFilteredRows.length}건 중 상위 ${PAGE_SIZE_WARNING}건만 표시 중입니다.`
      : `경고 행 ${lastFilteredRows.length}건을 표시 중입니다.`;
});

function handleFileUpload(e) {
  const file = e.target.files[0];
  if (!file) return;

  fileName.textContent = file.name;

  const reader = new FileReader();

  reader.onload = event => {
    const data = new Uint8Array(event.target.result);

    workbook = XLSX.read(data, {
      type: "array",
      cellFormula: false,
      cellHTML: false,
      cellNF: false,
      cellStyles: false
    });

    renderSheetButtons(workbook.SheetNames);
    resetView();

    resultMessage.textContent =
      "시트를 선택하세요. 부재별산출서 또는 아파트옹벽 Unit별산출서 시트를 우선 검토할 수 있습니다.";
  };

  reader.readAsArrayBuffer(file);
}

function renderSheetButtons(sheetNames) {
  sheetButtons.innerHTML = "";

  sheetNames.forEach(name => {
    const btn = document.createElement("button");
    btn.className = "sheet-btn";

    if (isTargetSheet(name)) {
      btn.classList.add("target");
    }

    btn.textContent = name;

    btn.addEventListener("click", () => selectSheet(name, btn));

    sheetButtons.appendChild(btn);
  });
}

function selectSheet(sheetName, button) {
  document.querySelectorAll(".sheet-btn").forEach(btn => btn.classList.remove("active"));
  button.classList.add("active");

  currentSheetName = sheetName;
  selectedSheet.textContent = sheetName;

  const sheet = workbook.Sheets[sheetName];

  const rows = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: "",
    raw: false
  });

  const parsed = parseSheetRows(rows);

  currentRows = parsed.dataRows;
  currentHeader = parsed.header;
  currentHeaderRowIndex = parsed.headerRowIndex;
  currentFormulaColIndex = parsed.formulaColIndex;
  currentResultColIndex = parsed.resultColIndex;

  lastFilteredRows = [];
  lastLimit = null;

  totalRows.textContent = currentRows.length;
  warningRows.textContent = "0";
  limitValue.textContent = "-";

  manualCheckBtn.disabled = currentFormulaColIndex === -1;
  showAllBtn.disabled = currentRows.length === 0;
  showWarningBtn.disabled = true;

  if (currentFormulaColIndex === -1) {
    resultMessage.textContent =
      "이 시트에서 '산출식' 열을 찾지 못했습니다. 헤더명에 산출식이 포함되어야 합니다.";
  } else {
    resultMessage.textContent =
      `'${sheetName}' 시트에서 산출식 열을 찾았습니다. 수기입력 오류 확인을 실행할 수 있습니다.`;
  }

  renderTable(currentRows.slice(0, PAGE_SIZE_ALL), null, false);
}

function parseSheetRows(rows) {
  let headerRowIndex = -1;
  let formulaColIndex = -1;

  for (let r = 0; r < rows.length; r++) {
    const row = rows[r];

    for (let c = 0; c < row.length; c++) {
      const value = String(row[c] || "").replace(/\s/g, "");

      if (value.includes("산출식")) {
        headerRowIndex = r;
        formulaColIndex = c;
        break;
      }
    }

    if (headerRowIndex !== -1) break;
  }

  if (headerRowIndex === -1) {
    return {
      header: [],
      dataRows: [],
      headerRowIndex: -1,
      formulaColIndex: -1,
      resultColIndex: -1
    };
  }

  const header = rows[headerRowIndex].map((cell, index) => {
    const text = String(cell || "").trim();
    return text || `열${index + 1}`;
  });

  const resultColIndex = findResultColumnIndex(header);

  const dataRows = rows
    .slice(headerRowIndex + 1)
    .filter(row => row.some(cell => String(cell || "").trim() !== ""))
    .map((row, index) => ({
      originalRowNumber: headerRowIndex + index + 2,
      values: normalizeRow(row, header.length),
      qcWarning: false,
      overNumbers: []
    }));

  return {
    header,
    dataRows,
    headerRowIndex,
    formulaColIndex,
    resultColIndex
  };
}

function findResultColumnIndex(header) {
  return header.findIndex(head => {
    const value = String(head || "").replace(/\s/g, "");
    return value.includes("결과값") || value.includes("결과");
  });
}

function normalizeRow(row, length) {
  const normalized = [];

  for (let i = 0; i < length; i++) {
    normalized.push(row[i] ?? "");
  }

  return normalized;
}

function runManualErrorCheck() {
  const limit = Number(manualLimit.value);

  if (Number.isNaN(limit) || limit < 0 || limit > 9999) {
    alert("0~9999 사이의 숫자를 입력하세요.");
    return;
  }

  if (currentFormulaColIndex === -1) {
    alert("'산출식' 열을 찾을 수 없습니다.");
    return;
  }

  lastLimit = limit;
  const checkedRows = [];

  for (const row of currentRows) {
    const formulaText = String(row.values[currentFormulaColIndex] || "").trim();

    if (!formulaText) continue;

    // 결과값이 0이거나 비어 있는 행 제외
    if (isZeroOrEmptyResult(row)) continue;

    const targetNumbers = extractManualInputNumbersOnly(formulaText);
    const overNumbers = targetNumbers.filter(num => num > limit);

    if (overNumbers.length > 0) {
      checkedRows.push({
        ...row,
        qcWarning: true,
        overNumbers
      });
    }
  }

  lastFilteredRows = checkedRows;

  warningRows.textContent = lastFilteredRows.length;
  limitValue.textContent = limit;
  showWarningBtn.disabled = lastFilteredRows.length === 0;

  renderTable(lastFilteredRows.slice(0, PAGE_SIZE_WARNING), limit, true);

  if (lastFilteredRows.length > PAGE_SIZE_WARNING) {
    resultMessage.textContent =
      `산출식 열의 수기입력 영역에서 ${limit}보다 큰 숫자가 포함된 행 ${lastFilteredRows.length}건을 찾았습니다. ` +
      `단, 결과값이 0이거나 빈 행은 제외했습니다. 속도 개선을 위해 상위 ${PAGE_SIZE_WARNING}건만 먼저 표시합니다.`;
  } else if (lastFilteredRows.length > 0) {
    resultMessage.textContent =
      `산출식 열의 수기입력 영역에서 ${limit}보다 큰 숫자가 포함된 행 ${lastFilteredRows.length}건을 찾았습니다. ` +
      `단, 결과값이 0이거나 빈 행은 제외했습니다.`;
  } else {
    resultMessage.textContent =
      `산출식 열의 수기입력 영역에서 ${limit}보다 큰 숫자가 포함된 행이 없습니다. ` +
      `결과값이 0이거나 빈 행은 제외했습니다.`;
  }
}

function isZeroOrEmptyResult(row) {
  if (currentResultColIndex === -1) return false;

  const rawValue = String(row.values[currentResultColIndex] || "").trim();

  if (!rawValue) return true;

  const cleanedValue = rawValue
    .replace(/,/g, "")
    .replace(/\s/g, "");

  const numericValue = Number(cleanedValue);

  if (Number.isNaN(numericValue)) return false;

  return numericValue === 0;
}

function extractManualInputNumbersOnly(text) {
  if (!text) return [];

  const source = String(text);

  /*
    검사 기준:
    1. = 뒤 계산 결과값은 전부 제외
    2. /1000, /1000/1000 같은 단위 변환 숫자 제외
    3. 0보다 작거나 9999보다 큰 숫자 제외
    4. 1 미만 소수값은 계수 가능성이 높아 제외
  */
  const beforeEqual = source.split("=")[0];

  const matches = beforeEqual.match(/\d+(\.\d+)?/g);

  if (!matches) return [];

  const numbers = [];

  matches.forEach(value => {
    const num = Number(value);

    if (Number.isNaN(num)) return;
    if (num === 1000) return;
    if (num < 1) return;
    if (num < 0 || num > 9999) return;

    numbers.push(num);
  });

  return numbers;
}

function renderTable(rows, limit, warningOnly) {
  const thead = resultTable.querySelector("thead");
  const tbody = resultTable.querySelector("tbody");

  thead.innerHTML = "";
  tbody.innerHTML = "";

  if (!currentHeader.length) {
    return;
  }

  const headerRow = document.createElement("tr");

  const qcTh = document.createElement("th");
  qcTh.textContent = "QC";
  headerRow.appendChild(qcTh);

  const rowNumTh = document.createElement("th");
  rowNumTh.textContent = "엑셀 행";
  headerRow.appendChild(rowNumTh);

  currentHeader.forEach((head, index) => {
    const th = document.createElement("th");

    if (index === currentFormulaColIndex) {
      th.textContent = `${head} ★`;
    } else if (index === currentResultColIndex) {
      th.textContent = `${head} ※`;
    } else {
      th.textContent = head;
    }

    headerRow.appendChild(th);
  });

  thead.appendChild(headerRow);

  if (warningOnly && rows.length === 0) {
    const tr = document.createElement("tr");
    const td = document.createElement("td");
    td.colSpan = currentHeader.length + 2;
    td.textContent = "경고 항목이 없습니다.";
    tr.appendChild(td);
    tbody.appendChild(tr);
    return;
  }

  const fragment = document.createDocumentFragment();

  rows.forEach(row => {
    const tr = document.createElement("tr");

    if (row.qcWarning) {
      tr.classList.add("warning-row");
    }

    const qcTd = document.createElement("td");
    qcTd.innerHTML = row.qcWarning
      ? `<span class="qc-badge warning">경고</span>`
      : `<span class="qc-badge normal">정상</span>`;
    tr.appendChild(qcTd);

    const rowNumTd = document.createElement("td");
    rowNumTd.textContent = row.originalRowNumber;
    tr.appendChild(rowNumTd);

    row.values.forEach((cell, index) => {
      const td = document.createElement("td");

      if (index === currentFormulaColIndex && limit !== null && limit !== undefined) {
        td.innerHTML = highlightOverNumbers(String(cell), limit);
      } else {
        td.textContent = cell;
      }

      tr.appendChild(td);
    });

    fragment.appendChild(tr);
  });

  tbody.appendChild(fragment);
}

function highlightOverNumbers(text, limit) {
  const escaped = escapeHtml(text);

  const parts = escaped.split("=");

  const beforeEqual = parts[0];
  const afterEqual = parts.length > 1 ? "=" + parts.slice(1).join("=") : "";

  const highlightedBeforeEqual = beforeEqual.replace(/\d+(\.\d+)?/g, match => {
    const num = Number(match);

    if (
      !Number.isNaN(num) &&
      num > limit &&
      num !== 1000 &&
      num >= 1 &&
      num <= 9999
    ) {
      return `<span class="match-number">${match}</span>`;
    }

    return match;
  });

  return highlightedBeforeEqual + afterEqual;
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function isTargetSheet(name) {
  return targetSheetNames.some(target => name.includes(target));
}

function resetView() {
  currentSheetName = "";
  currentRows = [];
  currentHeader = [];
  currentHeaderRowIndex = -1;
  currentFormulaColIndex = -1;
  currentResultColIndex = -1;
  lastFilteredRows = [];
  lastLimit = null;

  selectedSheet.textContent = "-";
  totalRows.textContent = "0";
  warningRows.textContent = "0";
  limitValue.textContent = "-";

  manualCheckBtn.disabled = true;
  showAllBtn.disabled = true;
  showWarningBtn.disabled = true;
  manualInputArea.classList.add("hidden");

  resultTable.querySelector("thead").innerHTML = "";
  resultTable.querySelector("tbody").innerHTML = "";
}
