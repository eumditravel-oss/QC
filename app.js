// app.js
let workbook = null;
let currentSheetName = "";
let currentRows = [];
let currentHeader = [];
let currentHeaderRowIndex = -1;
let currentFormulaColIndex = -1;
let lastFilteredRows = [];

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

const targetSheetNames = ["부재별산출서", "아파트옹벽 Unit별산출서"];

excelFile.addEventListener("change", handleFileUpload);
manualCheckBtn.addEventListener("click", () => {
  manualInputArea.classList.toggle("hidden");
});
runManualCheck.addEventListener("click", runManualErrorCheck);
showAllBtn.addEventListener("click", () => {
  renderTable(currentRows, null, false);
  resultMessage.textContent = "전체 행을 표시 중입니다.";
});

function handleFileUpload(e) {
  const file = e.target.files[0];
  if (!file) return;

  fileName.textContent = file.name;

  const reader = new FileReader();
  reader.onload = event => {
    const data = new Uint8Array(event.target.result);
    workbook = XLSX.read(data, { type: "array" });

    renderSheetButtons(workbook.SheetNames);
    resetView();

    resultMessage.textContent = "시트를 선택하세요. 부재별산출서 또는 아파트옹벽 Unit별산출서 시트를 우선 검토할 수 있습니다.";
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

  totalRows.textContent = currentRows.length;
  warningRows.textContent = "0";
  limitValue.textContent = "-";

  manualCheckBtn.disabled = currentFormulaColIndex === -1;
  showAllBtn.disabled = false;

  if (currentFormulaColIndex === -1) {
    resultMessage.textContent = "이 시트에서 '산출식' 열을 찾지 못했습니다. 헤더명에 산출식이 포함되어야 합니다.";
  } else {
    resultMessage.textContent = `'${sheetName}' 시트에서 산출식 열을 찾았습니다. 수기입력 오류 확인을 실행할 수 있습니다.`;
  }

  renderTable(currentRows, null, false);
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
      formulaColIndex: -1
    };
  }

  const header = rows[headerRowIndex].map((cell, index) => {
    const text = String(cell || "").trim();
    return text || `열${index + 1}`;
  });

  const dataRows = rows
    .slice(headerRowIndex + 1)
    .filter(row => row.some(cell => String(cell || "").trim() !== ""))
    .map((row, index) => ({
      originalRowNumber: headerRowIndex + index + 2,
      values: normalizeRow(row, header.length)
    }));

  return {
    header,
    dataRows,
    headerRowIndex,
    formulaColIndex
  };
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

  const checkedRows = currentRows.map(row => {
    const formulaText = String(row.values[currentFormulaColIndex] || "");
    const numbers = extractNumbers(formulaText);
    const overNumbers = numbers.filter(num => num > limit);

    return {
      ...row,
      qcWarning: overNumbers.length > 0,
      overNumbers
    };
  });

  lastFilteredRows = checkedRows.filter(row => row.qcWarning);

  warningRows.textContent = lastFilteredRows.length;
  limitValue.textContent = limit;

  renderTable(lastFilteredRows, limit, true);

  resultMessage.textContent =
    lastFilteredRows.length > 0
      ? `산출식 열에서 ${limit}보다 큰 숫자가 포함된 행 ${lastFilteredRows.length}건을 찾았습니다.`
      : `산출식 열에서 ${limit}보다 큰 숫자가 포함된 행이 없습니다.`;
}

function extractNumbers(text) {
  const matches = String(text).match(/\d+(\.\d+)?/g);
  if (!matches) return [];

  return matches
    .map(Number)
    .filter(num => !Number.isNaN(num) && num >= 0 && num <= 999999);
}

function renderTable(rows, limit, warningOnly) {
  const thead = resultTable.querySelector("thead");
  const tbody = resultTable.querySelector("tbody");

  thead.innerHTML = "";
  tbody.innerHTML = "";

  if (!currentHeader.length) {
    thead.innerHTML = "";
    tbody.innerHTML = "";
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
    th.textContent = index === currentFormulaColIndex ? `${head} ★` : head;
    headerRow.appendChild(th);
  });

  thead.appendChild(headerRow);

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

    tbody.appendChild(tr);
  });

  if (warningOnly && rows.length === 0) {
    const tr = document.createElement("tr");
    const td = document.createElement("td");
    td.colSpan = currentHeader.length + 2;
    td.textContent = "경고 항목이 없습니다.";
    tr.appendChild(td);
    tbody.appendChild(tr);
  }
}

function highlightOverNumbers(text, limit) {
  return escapeHtml(text).replace(/\d+(\.\d+)?/g, match => {
    const num = Number(match);

    if (num > limit) {
      return `<span class="match-number">${match}</span>`;
    }

    return match;
  });
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
  lastFilteredRows = [];

  selectedSheet.textContent = "-";
  totalRows.textContent = "0";
  warningRows.textContent = "0";
  limitValue.textContent = "-";
  manualCheckBtn.disabled = true;
  showAllBtn.disabled = true;
  manualInputArea.classList.add("hidden");

  resultTable.querySelector("thead").innerHTML = "";
  resultTable.querySelector("tbody").innerHTML = "";
}
