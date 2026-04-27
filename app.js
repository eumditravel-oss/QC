// app.js
const qcData = [
  {
    risk: "danger",
    status: "위험",
    member: "B1",
    item: "보 주근 길이 검토",
    drawing: "보 길이 6,000 / 4-D25",
    estimate: "철근 길이 11,800",
    diff: "+28%",
    action: "정착/이음 길이 재검토"
  },
  {
    risk: "warning",
    status: "주의",
    member: "S3",
    item: "슬래브 배근 간격",
    drawing: "D13@200",
    estimate: "D13@250",
    diff: "-20%",
    action: "도면 배근 기준 확인"
  },
  {
    risk: "danger",
    status: "위험",
    member: "W12",
    item: "벽체 중복 산출",
    drawing: "1개소",
    estimate: "2개소",
    diff: "+100%",
    action: "동일 좌표 중복 여부 확인"
  },
  {
    risk: "warning",
    status: "주의",
    member: "C5",
    item: "기둥 철근량 편차",
    drawing: "동일 타입 C5",
    estimate: "평균 대비 증가",
    diff: "+17%",
    action: "동일 타입 기둥 비교"
  },
  {
    risk: "safe",
    status: "정상",
    member: "B7",
    item: "스터럽 개수",
    drawing: "D10@150",
    estimate: "기준 범위 내",
    diff: "±2%",
    action: "확인 완료"
  }
];

const tableBody = document.getElementById("qcTable");
const riskFilter = document.getElementById("riskFilter");

function renderTable(filter = "all") {
  tableBody.innerHTML = "";

  const filteredData =
    filter === "all" ? qcData : qcData.filter(item => item.risk === filter);

  filteredData.forEach(item => {
    const tr = document.createElement("tr");

    tr.innerHTML = `
      <td><span class="badge ${item.risk}">${item.status}</span></td>
      <td>${item.member}</td>
      <td>${item.item}</td>
      <td>${item.drawing}</td>
      <td>${item.estimate}</td>
      <td>${item.diff}</td>
      <td>${item.action}</td>
    `;

    tableBody.appendChild(tr);
  });
}

riskFilter.addEventListener("change", e => {
  renderTable(e.target.value);
});

renderTable();
