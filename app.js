// app.js
"use strict";

/* =========================
   공통 상태
========================= */

let workbook = null;
let uploadedFiles = [];
let currentSheetName = "";
let currentRows = [];
let currentHeader = [];
let currentFormulaColIndex = -1;
let currentResultColIndex = -1;
let lastFilteredRows = [];
let lastLimit = null;

const CATEGORIES = ["콘크리트", "거푸집", "철근", "잡/기타"];

const floorState = {
  rawItems: [],
  dongs: [],
  floors: [],
  data: {},
  mappings: [],
  areas: {},
  ready: false
};

const $ = id => document.getElementById(id);

const excelFile = $("excelFile");
const sheetButtons = $("sheetButtons");
const fileName = $("fileName");
const selectedSheet = $("selectedSheet");
const totalRows = $("totalRows");
const warningRows = $("warningRows");
const limitValue = $("limitValue");
const manualCheckBtn = $("manualCheckBtn");
const manualInputArea = $("manualInputArea");
const manualLimit = $("manualLimit");
const runManualCheck = $("runManualCheck");
const resultMessage = $("resultMessage");
const resultTable = $("resultTable");
const showAllBtn = $("showAllBtn");
const showWarningBtn = $("showWarningBtn");
const normalQcPanel = $("normalQcPanel");
const floorSummaryPanel = $("floorSummaryPanel");
const floorSubView = $("floorSubView");

const targetSheetNames = ["부재별산출서", "아파트옹벽 Unit별산출서"];
const floorSheetKeyword = "층별총집계표";
const PAGE_SIZE_WARNING = 500;
const PAGE_SIZE_ALL = 500;

/* =========================
   이벤트
========================= */

excelFile.addEventListener("change", handleFileUpload);

manualCheckBtn.addEventListener("click", () => {
  manualInputArea.classList.toggle("hidden");
});

runManualCheck.addEventListener("click", runManualErrorCheck);

showAllBtn.addEventListener("click", () => {
  renderTable(currentRows.slice(0, PAGE_SIZE_ALL), null, false);
  resultMessage.textContent = `전체 행 중 상위 ${PAGE_SIZE_ALL}건만 표시 중입니다.`;
});

showWarningBtn.addEventListener("click", () => {
  renderTable(lastFilteredRows.slice(0, PAGE_SIZE_WARNING), lastLimit, true);
  resultMessage.textContent = `경고 행 ${lastFilteredRows.length}건 중 상위 ${Math.min(PAGE_SIZE_WARNING, lastFilteredRows.length)}건을 표시 중입니다.`;
});

$("btnFloorParse").addEventListener("click", parseFloorSummary);
$("btnMapping").addEventListener("click", renderMappingView);
$("btnArea").addEventListener("click", renderAreaView);
$("btnFloorView").addEventListener("click", renderFloorView);
$("btnFloorExcel").addEventListener("click", exportFloorExcel);

/* =========================
   파일 업로드 / 시트 선택
========================= */

function handleFileUpload(e) {
  uploadedFiles = Array.from(e.target.files);
  if (!uploadedFiles.length) return;

  fileName.textContent = uploadedFiles.length === 1
    ? uploadedFiles[0].name
    : `${uploadedFiles.length}개 파일 업로드됨`;

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
    resetNormalQcView();

    resultMessage.textContent = "시트를 선택하세요.";
  };

  reader.readAsArrayBuffer(uploadedFiles[0]);
}

function renderSheetButtons(sheetNames) {
  sheetButtons.innerHTML = "";

  sheetNames.forEach(name => {
    const btn = document.createElement("button");
    btn.className = "sheet-btn";

    if (isTargetSheet(name)) btn.classList.add("target");
    if (isFloorSheet(name)) btn.classList.add("floor-target");

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

  if (isFloorSheet(sheetName)) {
  normalQcPanel.classList.add("hidden");
  floorSummaryPanel.classList.remove("hidden");

  document.querySelector(".cards").classList.add("hidden");
     normalQcPanel.classList.remove("hidden");
floorSummaryPanel.classList.add("hidden");
document.querySelector(".cards").classList.remove("hidden");
  resetNormalQcTableOnly();

  resultMessage.textContent = "층별총집계표 분석 기능을 사용할 수 있습니다.";
  return;
}

  normalQcPanel.classList.remove("hidden");
  floorSummaryPanel.classList.add("hidden");

  const sheet = workbook.Sheets[sheetName];

  const rows = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: "",
    raw: false
  });

  const parsed = parseSheetRows(rows);

  currentRows = parsed.dataRows;
  currentHeader = parsed.header;
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

  resultMessage.textContent = currentFormulaColIndex === -1
    ? "이 시트에서 '산출식' 열을 찾지 못했습니다."
    : `'${sheetName}' 시트에서 산출식 열을 찾았습니다.`;

  renderTable(currentRows.slice(0, PAGE_SIZE_ALL), null, false);
}

/* =========================
   산출서 수기입력 QC
========================= */

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
    return { header: [], dataRows: [], formulaColIndex: -1, resultColIndex: -1 };
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

  return { header, dataRows, formulaColIndex, resultColIndex };
}

function findResultColumnIndex(header) {
  return header.findIndex(head => {
    const value = String(head || "").replace(/\s/g, "");
    return value.includes("결과값") || value.includes("결과");
  });
}

function normalizeRow(row, length) {
  return Array.from({ length }, (_, i) => row[i] ?? "");
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
    if (isZeroOrEmptyResult(row)) continue;

    const targetNumbers = extractManualInputNumbersOnly(formulaText);
    const overNumbers = targetNumbers.filter(num => num > limit);

    if (overNumbers.length > 0) {
      checkedRows.push({ ...row, qcWarning: true, overNumbers });
    }
  }

  lastFilteredRows = checkedRows;

  warningRows.textContent = lastFilteredRows.length;
  limitValue.textContent = limit;
  showWarningBtn.disabled = lastFilteredRows.length === 0;

  renderTable(lastFilteredRows.slice(0, PAGE_SIZE_WARNING), limit, true);

  resultMessage.textContent =
    `산출식 열의 수기입력 영역에서 ${limit}보다 큰 숫자가 포함된 행 ${lastFilteredRows.length}건을 찾았습니다. ` +
    `결과값이 0이거나 빈 행은 제외했습니다.`;
}

function isZeroOrEmptyResult(row) {
  if (currentResultColIndex === -1) return false;

  const rawValue = String(row.values[currentResultColIndex] || "").trim();
  if (!rawValue) return true;

  const numericValue = Number(rawValue.replace(/,/g, "").replace(/\s/g, ""));
  if (Number.isNaN(numericValue)) return false;

  return numericValue === 0;
}

function extractManualInputNumbersOnly(text) {
  const beforeEqual = String(text).split("=")[0];
  const matches = beforeEqual.match(/\d+(\.\d+)?/g);
  if (!matches) return [];

  return matches
    .map(Number)
    .filter(num =>
      !Number.isNaN(num) &&
      num !== 1000 &&
      num >= 1 &&
      num <= 9999
    );
}

function renderTable(rows, limit, warningOnly) {
  const thead = resultTable.querySelector("thead");
  const tbody = resultTable.querySelector("tbody");

  thead.innerHTML = "";
  tbody.innerHTML = "";

  if (!currentHeader.length) return;

  const headerRow = document.createElement("tr");

  ["QC", "엑셀 행"].forEach(text => {
    const th = document.createElement("th");
    th.textContent = text;
    headerRow.appendChild(th);
  });

  currentHeader.forEach((head, index) => {
    const th = document.createElement("th");
    th.textContent = index === currentFormulaColIndex ? `${head} ★` : head;
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
    if (row.qcWarning) tr.classList.add("warning-row");

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
      td.innerHTML =
        index === currentFormulaColIndex && limit !== null && limit !== undefined
          ? highlightOverNumbers(String(cell), limit)
          : escapeHtml(cell);
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

    if (!Number.isNaN(num) && num > limit && num !== 1000 && num >= 1 && num <= 9999) {
      return `<span class="match-number">${match}</span>`;
    }

    return match;
  });

  return highlightedBeforeEqual + afterEqual;
}

/* =========================
   층별총집계표 분석
========================= */

async function parseFloorSummary() {
  const files = uploadedFiles.length ? uploadedFiles : [];
  if (!files.length) {
    alert("엑셀 파일을 먼저 업로드하세요.");
    return;
  }

  resetFloorState();

  for (const file of files) {
    const buffer = await file.arrayBuffer();
    const wb = XLSX.read(buffer, { type: "array" });

    for (const sheetName of wb.SheetNames) {
      if (!isFloorSheet(sheetName)) continue;

      const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], {
        header: 1,
        defval: ""
      });

      parseFloorRows(rows);
    }
  }

  buildMapping();

  $("btnMapping").disabled = false;
  $("btnArea").disabled = false;
  $("btnFloorView").disabled = false;
  $("btnFloorExcel").disabled = false;

  totalRows.textContent = floorState.floors.length;
  warningRows.textContent = floorState.rawItems.length;
  limitValue.textContent = "-";

  resultMessage.textContent =
    `층별총집계표 분석 완료: 동 ${floorState.dongs.length}개, 층 ${floorState.floors.length}개, 아이템 ${floorState.rawItems.length}개를 인식했습니다.`;

  renderMappingView();
}

function resetFloorState() {
  floorState.rawItems = [];
  floorState.dongs = [];
  floorState.floors = [];
  floorState.data = {};
  floorState.mappings = [];
  floorState.areas = {};
  floorState.ready = false;
}

function parseFloorRows(rows) {
  let curDong = "";
  let lastF = "";
  const r3 = rows[2] || [];
  const r4 = rows[3] || [];

  for (let r = 4; r < rows.length; r++) {
    const row = rows[r];
    if (!row || row.length === 0) continue;

    const txt = row.join("|");
    const m = txt.match(/동\s*명\s*:\s*\[([^\]]+)\]/);

    if (m) {
      const raw = m[1].trim();

      if (raw) {
        curDong = raw;
        if (!floorState.dongs.includes(curDong)) floorState.dongs.push(curDong);
        floorState.data[curDong] = floorState.data[curDong] || {};
      }

      lastF = "";
      continue;
    }

    if (!curDong) continue;

    const fRaw = String(row[0]).trim();

    if (fRaw === "층" || fRaw.includes("계") || fRaw.includes("합") || fRaw.includes("공사명")) {
      lastF = "";
      continue;
    }

    if (fRaw !== "") {
      lastF = /^\d+$/.test(fRaw) ? `${fRaw}F` : fRaw;
      if (!floorState.floors.includes(lastF)) floorState.floors.push(lastF);
    }

    if (!lastF) continue;

    for (let c = 1; c < row.length; c++) {
      const val = parseFloat(String(row[c]).replace(/,/g, ""));
      if (Number.isNaN(val) || val === 0) continue;

      let name = fRaw !== "" ? String(r3[c] || "").trim() : String(r4[c] || "").trim();
      if (!name) name = String(r3[c] || r4[c] || "").trim();
      if (!name) continue;

      if (!floorState.rawItems.includes(name)) floorState.rawItems.push(name);

      floorState.data[curDong][name] = floorState.data[curDong][name] || {};
      floorState.data[curDong][name][lastF] =
        (floorState.data[curDong][name][lastF] || 0) + val;
    }
  }
}

function buildMapping() {
  floorState.mappings = floorState.rawItems.map((item, idx) => ({
    id: idx,
    original: item,
    canonical: item,
    category: predictCategory(item)
  }));
}

function predictCategory(name) {
  const s = String(name).toUpperCase().replace(/\s+/g, "");

  if (/(H|D|HD|SD)\d+/.test(s) || s.includes("철근")) return "철근";
  if (s.includes("MPA") || /\d+-\d+-\d+/.test(s) || (/^\d+$/.test(s) && parseInt(s) >= 150)) return "콘크리트";
  if (["폼", "FORM", "회", "알폼", "갱폼", "합벽"].some(k => s.includes(k)) || /[가-힣]/.test(s)) return "거푸집";

  return "잡/기타";
}

function renderMappingView() {
  floorSubView.innerHTML = `
    <div class="toolbar">
      <h3>아이템 분류 설정</h3>
      <button class="primary" onclick="renderAreaView()">설정 적용 및 면적 입력으로 이동</button>
    </div>
    <div class="mapping-container">
      <div class="list-header">
        <div>번호</div>
        <div>원본 아이템 명칭</div>
        <div>표준화 명칭</div>
        <div>중분류</div>
      </div>
      <div class="mapping-list">
        ${floorState.mappings.map(m => {
          const catClass = m.category === "잡/기타" ? "etc" : m.category;
          return `
            <div class="item-row cat-${catClass}">
              <div>${m.id + 1}</div>
              <div>${escapeHtml(m.original)}</div>
              <div><input class="input" value="${escapeHtml(m.canonical)}" oninput="updateMapping(${m.id}, 'canonical', this.value)" /></div>
              <div>
                <select class="input" onchange="updateMapping(${m.id}, 'category', this.value)">
                  ${CATEGORIES.map(c => `<option value="${c}" ${m.category === c ? "selected" : ""}>${c}</option>`).join("")}
                </select>
              </div>
            </div>
          `;
        }).join("")}
      </div>
    </div>
  `;
}

window.updateMapping = function(id, field, value) {
  floorState.mappings[id][field] = value;
};

function renderAreaView() {
  const dongs = [...floorState.dongs].sort();
  const floors = [...floorState.floors].sort(floorSorter);

  floorSubView.innerHTML = `
    <div class="toolbar">
      <h3>층별 면적(m²) 입력</h3>
      <button class="sub" onclick="downloadAreaTemplate()">면적 양식 내보내기</button>
      <label class="sub" style="cursor:pointer;">
        면적 양식 불러오기
        <input type="file" accept=".xlsx,.xls" onchange="uploadAreaTemplate(event)" style="display:none;" />
      </label>
      <button class="primary" onclick="renderFloorView()">입력 완료 및 수량표 확인</button>
    </div>

    <div class="excel-view-container">
      <table>
        <thead>
          <tr>
            <th>층 명칭</th>
            ${dongs.map(d => `<th>${escapeHtml(d)}</th>`).join("")}
          </tr>
        </thead>
        <tbody>
          ${floors.map((f, rIdx) => `
            <tr>
              <td><b>${escapeHtml(f)}</b></td>
              ${dongs.map((d, cIdx) => `
                <td>
                  <input
                    type="number"
                    class="area-input"
                    data-r="${rIdx}"
                    data-c="${cIdx}"
                    value="${floorState.areas[d]?.[f] || ""}"
                    oninput="updateArea('${escapeAttr(d)}','${escapeAttr(f)}',this.value)"
                    onkeydown="handleAreaNav(event, ${rIdx}, ${cIdx}, ${floors.length}, ${dongs.length})"
                    placeholder="-"
                  />
                </td>
              `).join("")}
            </tr>
          `).join("")}
        </tbody>
      </table>
    </div>
  `;
}

window.updateArea = function(dong, floor, val) {
  if (!floorState.areas[dong]) floorState.areas[dong] = {};
  floorState.areas[dong][floor] = parseFloat(val) || 0;
};

window.handleAreaNav = function(e, r, c, maxR, maxC) {
  let nr = r;
  let nc = c;

  if (e.key === "ArrowUp") nr = Math.max(0, r - 1);
  else if (e.key === "ArrowDown" || e.key === "Enter") {
    nr = Math.min(maxR - 1, r + 1);
    e.preventDefault();
  } else if (e.key === "ArrowLeft") nc = Math.max(0, c - 1);
  else if (e.key === "ArrowRight") nc = Math.min(maxC - 1, c + 1);
  else return;

  const input = document.querySelector(`.area-input[data-r="${nr}"][data-c="${nc}"]`);
  if (input) {
    input.focus();
    input.select();
  }
};

function downloadAreaTemplate() {
  const dongs = [...floorState.dongs].sort();
  const floors = [...floorState.floors].sort(floorSorter);
  const aoa = [["층 명칭", ...dongs]];

  floors.forEach(f => {
    const row = [f];
    dongs.forEach(d => row.push(floorState.areas[d]?.[f] || ""));
    aoa.push(row);
  });

  const ws = XLSX.utils.aoa_to_sheet(aoa);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "면적데이터");
  XLSX.writeFile(wb, "QS_면적입력양식.xlsx");
}

window.uploadAreaTemplate = async function(e) {
  const file = e.target.files[0];
  if (!file) return;

  const buffer = await file.arrayBuffer();
  const wb = XLSX.read(buffer, { type: "array" });
  const data = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: "" });

  const headers = data[0] || [];

  for (let i = 1; i < data.length; i++) {
    const floor = String(data[i][0]).trim();
    if (!floor) continue;

    for (let c = 1; c < headers.length; c++) {
      const dong = String(headers[c]).trim();
      const val = parseFloat(data[i][c]);

      if (!Number.isNaN(val)) {
        if (!floorState.areas[dong]) floorState.areas[dong] = {};
        floorState.areas[dong][floor] = val;
      }
    }
  }

  renderAreaView();
};

function renderFloorView() {
  floorState.ready = true;

  const dongs = [...floorState.dongs].sort();

  floorSubView.innerHTML = `
    <div class="toolbar">
      <label>조회 동 선택</label>
      <select id="filterDong" class="input" style="max-width:220px;" onchange="renderSelectedDongView()">
        ${dongs.map(d => `<option value="${escapeAttr(d)}">${escapeHtml(d)}</option>`).join("")}
      </select>
      <button class="success" onclick="exportFloorExcel()">템플릿 엑셀 다운로드</button>
    </div>

    <div class="excel-view-container">
      <table>
        <thead id="floorTableHead"></thead>
        <tbody id="floorTableBody"></tbody>
      </table>
    </div>
  `;

  renderSelectedDongView();
}

window.renderSelectedDongView = function() {
  const dong = $("filterDong").value;
  const floors = [...floorState.floors].sort(floorSorter);
  const dongData = floorState.data[dong] || {};
  const grouped = buildGroupedData(dongData, floors);

  $("floorTableHead").innerHTML = `
    <tr>
      <th rowspan="2">동</th>
      <th rowspan="2">아이템</th>
      <th rowspan="2">구분</th>
      <th rowspan="2">단위</th>
      <th colspan="${floors.length}">현재 프로젝트 수량</th>
      <th rowspan="2">합계</th>
    </tr>
    <tr>
      ${floors.map(f => `<th>${escapeHtml(f)}</th>`).join("")}
    </tr>
  `;

  let bodyHtml = "";

  ["콘크리트", "철근", "거푸집", "잡/기타"].forEach(cat => {
    const items = Object.keys(grouped).filter(n => grouped[n].category === cat).sort();
    if (!items.length) return;

    const catClass = cat === "잡/기타" ? "etc" : cat;
    let catSum = 0;

    items.forEach(name => {
      const item = grouped[name];
      const total = floors.reduce((s, f) => s + (item.floors[f] || 0), 0);
      catSum += total;

      bodyHtml += `
        <tr class="row-cat-${catClass}">
          <td>${escapeHtml(dong)}</td>
          <td>${cat === "콘크리트" ? "레미콘" : escapeHtml(cat)}</td>
          <td>${escapeHtml(name)}</td>
          <td>${cat === "철근" ? "TON" : cat === "콘크리트" ? "M3" : "M2"}</td>
          ${floors.map(f => `<td>${fmt(item.floors[f], 3)}</td>`).join("")}
          <td class="col-total">${fmt(total, 3)}</td>
        </tr>
      `;
    });

    bodyHtml += `
      <tr class="row-subtotal">
        <td colspan="3" style="text-align:right;">합계</td>
        <td>${cat === "철근" ? "TON" : cat === "콘크리트" ? "M3" : "M2"}</td>
        ${floors.map(f => {
          const s = items.reduce((sum, n) => sum + (grouped[n].floors[f] || 0), 0);
          return `<td>${fmt(s, 3)}</td>`;
        }).join("")}
        <td class="col-total">${fmt(catSum, 3)}</td>
      </tr>
    `;

    if (cat === "철근") {
      const numFn = f => sumCategory(grouped, "철근", f);
      bodyHtml += renderRatioRow("레미콘/철근", "Ton/m³", numFn, f => sumCategory(grouped, "콘크리트", f), "#C00000", "#FFFFFF", floors);
      bodyHtml += renderRatioRow("면적/철근", "Ton/m²", numFn, f => floorState.areas[dong]?.[f] || 0, "#FFC000", "#000000", floors);
      bodyHtml += renderRatioRow("평수/철근", "Ton/Py", numFn, f => (floorState.areas[dong]?.[f] || 0) * 0.3025, "#FFFF00", "#000000", floors);
    }

    if (cat === "거푸집") {
      const numFn = f => sumCategory(grouped, "거푸집", f);
      bodyHtml += renderRatioRow("거푸집/면적", "m²/m²", numFn, f => floorState.areas[dong]?.[f] || 0, "#00B050", "#FFFFFF", floors);
      bodyHtml += renderRatioRow("거푸집/평수", "m²/Py", numFn, f => (floorState.areas[dong]?.[f] || 0) * 0.3025, "#92D050", "#000000", floors);
    }
  });

  $("floorTableBody").innerHTML = bodyHtml;
};

function buildGroupedData(dongData, floors) {
  const grouped = {};

  floorState.mappings.forEach(m => {
    const qByF = dongData[m.original] || {};
    if (!Object.keys(qByF).length) return;

    if (!grouped[m.canonical]) {
      grouped[m.canonical] = { category: m.category, floors: {} };
    }

    floors.forEach(f => {
      grouped[m.canonical].floors[f] =
        (grouped[m.canonical].floors[f] || 0) + (qByF[f] || 0);
    });
  });

  return grouped;
}

function renderRatioRow(title, unit, numFn, divFn, bg, text, floors) {
  let totalNum = 0;
  let totalDiv = 0;

  let html = `
    <tr class="row-ratio">
      <td colspan="3" style="text-align:right;background:${bg};color:${text};">${title}</td>
      <td style="background:${bg};color:${text};">${unit}</td>
  `;

  floors.forEach(f => {
    const nVal = numFn(f);
    const dVal = divFn(f);
    totalNum += nVal;
    totalDiv += dVal;
    html += `<td style="background:${bg};color:${text};">${fmt(dVal > 0 ? nVal / dVal : 0, 4)}</td>`;
  });

  html += `<td class="col-total" style="background:${bg};color:${text};">${fmt(totalDiv > 0 ? totalNum / totalDiv : 0, 4)}</td></tr>`;
  return html;
}

function sumCategory(grouped, category, floor) {
  return Object.keys(grouped)
    .filter(n => grouped[n].category === category)
    .reduce((s, n) => s + (grouped[n].floors[floor] || 0), 0);
}

/* =========================
   층별총집계표 Excel 내보내기
========================= */

async function exportFloorExcel() {
  if (!floorState.rawItems.length) {
    alert("먼저 층별총집계표 분석을 실행하세요.");
    return;
  }

  if (typeof ExcelJS === "undefined") {
    alert("ExcelJS 라이브러리를 불러오지 못했습니다.");
    return;
  }

  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("비교양식", {
    views: [{ state: "frozen", ySplit: 4, xSplit: 4, showZeros: false }]
  });

  const floors = [...floorState.floors].sort(floorSorter);
  const endCol = 4 + floors.length + 1;
  const maxCol = endCol + 1;

  ws.columns = [
    { width: 10 },
    { width: 15 },
    { width: 18 },
    { width: 10 },
    ...floors.map(() => ({ width: 9 })),
    { width: 13 },
    { width: 12 }
  ];

  const titleRow = ws.addRow(["QS 분석용 프로젝트 통합 템플릿"]);
  titleRow.height = 25;
  ws.mergeCells(1, 1, 2, maxCol);

  ws.getCell(1, 1).font = { size: 16, bold: true, name: "맑은 고딕" };
  ws.getCell(1, 1).alignment = { vertical: "middle", horizontal: "center" };

  const r3Data = ["동", "아이템", "구분", "단위", "현재 프로젝트 수량"];
  for (let i = 0; i < floors.length - 1; i++) r3Data.push("");
  r3Data.push("합계", "비고");

  const r4Data = ["", "", "", "", ...floors, "", ""];

  ws.addRow(r3Data);
  ws.addRow(r4Data);

  ws.mergeCells(3, 1, 4, 1);
  ws.mergeCells(3, 2, 4, 2);
  ws.mergeCells(3, 3, 4, 3);
  ws.mergeCells(3, 4, 4, 4);
  ws.mergeCells(3, 5, 3, endCol - 1);
  ws.mergeCells(3, endCol, 4, endCol);
  ws.mergeCells(3, maxCol, 4, maxCol);

  const borderAll = {
    top: { style: "thin" },
    left: { style: "thin" },
    bottom: { style: "thin" },
    right: { style: "thin" }
  };

  for (let r = 3; r <= 4; r++) {
    for (let c = 1; c <= maxCol; c++) {
      const cell = ws.getCell(r, c);
      cell.font = { bold: true, size: 10, name: "맑은 고딕", color: { argb: "FFFFFFFF" } };
      cell.alignment = { vertical: "middle", horizontal: "center" };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1F4E78" } };
      cell.border = borderAll;
    }
  }

  const dataBorder = {
    top: { style: "thin", color: { argb: "FFBFBFBF" } },
    left: { style: "thin", color: { argb: "FFBFBFBF" } },
    bottom: { style: "thin", color: { argb: "FFBFBFBF" } },
    right: { style: "thin", color: { argb: "FFBFBFBF" } }
  };

  floorState.dongs.sort().forEach(dong => {
    const dongData = floorState.data[dong] || {};
    const grouped = buildGroupedData(dongData, floors);
    const startRow = ws.rowCount + 1;

    ["콘크리트", "철근", "거푸집"].forEach(cat => {
      const items = Object.keys(grouped).filter(n => grouped[n].category === cat).sort();
      if (!items.length) return;

      let rowFill = "FFFFFFFF";
      if (cat === "콘크리트") rowFill = "FFEEF4FF";
      if (cat === "철근") rowFill = "FFF0FCF4";
      if (cat === "거푸집") rowFill = "FFFFF9EC";

      const catSum = {};
      floors.forEach(f => catSum[f] = 0);
      let totalSum = 0;

      items.forEach(name => {
        const item = grouped[name];
        const rowData = [
          dong,
          cat === "콘크리트" ? "레미콘" : cat,
          name,
          cat === "철근" ? "TON" : cat === "콘크리트" ? "M3" : "M2"
        ];

        let rowTotal = 0;

        floors.forEach(f => {
          const val = item.floors[f] || 0;
          rowData.push(val);
          catSum[f] += val;
          rowTotal += val;
        });

        rowData.push(rowTotal, "");
        totalSum += rowTotal;

        const row = ws.addRow(rowData);
        row.outlineLevel = 1;

        for (let c = 1; c <= maxCol; c++) {
          const cell = row.getCell(c);
          cell.border = dataBorder;
          cell.font = { name: "맑은 고딕", size: 10 };
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: rowFill } };
          cell.alignment = { vertical: "middle", horizontal: c <= 4 ? "center" : "right" };
          if (c > 4) cell.numFmt = "#,##0.000";
        }
      });

      const sumRowData = [
        dong,
        cat === "콘크리트" ? "레미콘" : cat,
        "합계",
        cat === "철근" ? "TON" : cat === "콘크리트" ? "M3" : "M2"
      ];

      floors.forEach(f => sumRowData.push(catSum[f]));
      sumRowData.push(totalSum, "");

      const sumRow = ws.addRow(sumRowData);

      for (let c = 1; c <= maxCol; c++) {
        const cell = sumRow.getCell(c);
        cell.font = { name: "맑은 고딕", size: 10, bold: true };
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF2F2F2" } };
        cell.border = dataBorder;
        cell.alignment = { vertical: "middle", horizontal: c <= 4 ? "center" : "right" };
        if (c > 4) cell.numFmt = "#,##0.000";
      }
    });

    const endRow = ws.rowCount;
    if (startRow < endRow) {
      ws.mergeCells(startRow, 1, endRow, 1);
      ws.getCell(startRow, 1).alignment = { vertical: "middle", horizontal: "center" };
    }
  });

  const buffer = await wb.xlsx.writeBuffer();
  saveAs(
    new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }),
    "QS_통합템플릿_리포트.xlsx"
  );
}

/* =========================
   유틸
========================= */

function floorSorter(a, b) {
  const getRank = name => {
    const s = String(name).toUpperCase().trim();
    if (s.startsWith("B")) return 1000 - (parseInt(s.replace("B", "")) || 0);
    if (s === "FT") return 2000;
    if (s.endsWith("F") || /^\d+$/.test(s)) return 3000 + (parseInt(s.replace("F", "")) || 0);
    if (s.startsWith("PH")) return 4000 + (parseInt(s.replace("PH", "")) || 0);
    return 5000;
  };

  return getRank(a) - getRank(b);
}

function fmt(val, digits = 3) {
  if (val === 0 || !val || Number.isNaN(val)) return "-";
  return Number(val).toLocaleString(undefined, { maximumFractionDigits: digits });
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function escapeAttr(value) {
  return String(value ?? "")
    .replaceAll("\\", "\\\\")
    .replaceAll("'", "\\'")
    .replaceAll('"', "&quot;");
}

function isTargetSheet(name) {
  return targetSheetNames.some(target => name.includes(target));
}

function isFloorSheet(name) {
  return String(name || "").includes(floorSheetKeyword);
}

function resetNormalQcView() {
  currentSheetName = "";
  currentRows = [];
  currentHeader = [];
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

  normalQcPanel.classList.remove("hidden");
  floorSummaryPanel.classList.add("hidden");

  resetNormalQcTableOnly();
}

function resetNormalQcTableOnly() {
  resultTable.querySelector("thead").innerHTML = "";
  resultTable.querySelector("tbody").innerHTML = "";
}
