/*
===========================================
대시보드 JavaScript 애플리케이션 (dashboard.js)
===========================================

판매 데이터 대시보드의 클라이언트 측 로직을 담당합니다.

주요 기능:
1. API 통신 (데이터 조회, 분석, 예측)
2. UI 렌더링 (테이블, 차트, 카드)
3. 이벤트 처리 (버튼 클릭, 탭 전환)
4. 페이지네이션
5. 로딩 상태 관리

작성자: AI Assistant
작성일: 2026-02-09
*/

// ===========================================
// 전역 상태 관리
// ===========================================
const state = {
  currentTab: "overview", // 현재 활성 탭
  currentPage: {
    // 각 테이블의 현재 페이지
    "product-daily": 1,
    "product-mall": 1,
    "daily-mall": 1,
    "yearly-mall": 1,
  },
  analysisData: null, // 최근 분석 결과
  predictionData: null, // 최근 예측 결과
  charts: {}, // Chart.js 인스턴스 저장
};

// API 엔드포인트 베이스 URL
const API_BASE = ""; // 같은 도메인이므로 빈 문자열

// ===========================================
// 유틸리티 함수
// ===========================================

/**
 * 날짜를 YYYY/MM/DD 형식으로 포맷팅합니다.
 */
function formatDate(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}/${month}/${day}`;
}

/**
 * N일 전 날짜를 반환합니다.
 */
function getDaysAgo(days) {
  const date = new Date();
  date.setDate(date.getDate() - days + 1); // +1은 오늘 포함
  return date;
}

/**
 * 오늘 날짜를 반환합니다.
 */
function getToday() {
  return new Date();
}

/**
 * 숫자를 천 단위 콤마로 포맷팅합니다.
 */
function formatNumber(num) {
  if (num === null || num === undefined) return "0";
  return num.toLocaleString("ko-KR");
}

/**
 * 로딩 오버레이를 표시/숨김합니다.
 */
function showLoading(show = true) {
  const overlay = document.getElementById("loadingOverlay");
  if (show) {
    overlay.classList.add("active");
  } else {
    overlay.classList.remove("active");
  }
}

/**
 * 상태 메시지를 업데이트합니다.
 */
function setStatusMessage(message) {
  const statusEl = document.getElementById("statusMessage");
  if (statusEl) {
    statusEl.textContent = message;
  }
}

// ===========================================
// API 통신 함수
// ===========================================

/**
 * 데이터 정보를 조회합니다.
 */
async function fetchDataInfo() {
  try {
    const response = await fetch(`${API_BASE}/api/info`);
    if (!response.ok) throw new Error("정보 조회 실패");
    return await response.json();
  } catch (error) {
    console.error("데이터 정보 조회 오류:", error);
    return null;
  }
}

/**
 * 쇼핑몰 목록을 조회합니다.
 */
async function fetchMalls() {
  try {
    const response = await fetch(`${API_BASE}/api/malls`);
    if (!response.ok) throw new Error("쇼핑몰 목록 조회 실패");
    const data = await response.json();
    return data.malls || [];
  } catch (error) {
    console.error("쇼핑몰 목록 조회 오류:", error);
    return [];
  }
}

/**
 * 분석을 실행합니다.
 */
async function runAnalysis(params) {
  try {
    const response = await fetch(`${API_BASE}/api/analyze`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify(params),
    });

    if (!response.ok) {
      const error = await response.json();
      throw new Error(error.error || "분석 실패");
    }

    return await response.json();
  } catch (error) {
    console.error("분석 실행 오류:", error);
    throw error;
  }
}

/**
 * 쇼핑몰별 예측을 실행합니다.
 */
async function predictMall(params) {
  try {
    const response = await fetch(`${API_BASE}/api/predict-mall`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify(params),
    });

    if (!response.ok) {
      const error = await response.json();
      throw new Error(error.error || "예측 실패");
    }

    return await response.json();
  } catch (error) {
    console.error("쇼핑몰별 예측 오류:", error);
    throw error;
  }
}

/**
 * 상품별 예측을 실행합니다.
 */
async function predictProduct(params) {
  try {
    const response = await fetch(`${API_BASE}/api/predict-product`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify(params),
    });

    if (!response.ok) {
      const error = await response.json();
      throw new Error(error.error || "예측 실패");
    }

    return await response.json();
  } catch (error) {
    console.error("상품별 예측 오류:", error);
    throw error;
  }
}

/**
 * Excel 파일을 다운로드합니다.
 */
async function downloadExcel(params) {
  try {
    const response = await fetch(`${API_BASE}/api/export`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify(params),
    });

    if (!response.ok) {
      throw new Error("Excel 다운로드 실패");
    }

    // Blob으로 변환하여 다운로드
    const blob = await response.blob();
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `판매분석_${new Date().getTime()}.xlsx`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    window.URL.revokeObjectURL(url);

    return true;
  } catch (error) {
    console.error("Excel 다운로드 오류:", error);
    throw error;
  }
}

// ===========================================
// UI 렌더링 함수
// ===========================================

/**
 * 쇼핑몰 필터 리스트를 렌더링합니다.
 */
function renderMallList(malls) {
  const mallList = document.getElementById("mallList");
  if (!mallList) return;

  mallList.innerHTML = "";

  malls.forEach((mall, index) => {
    const div = document.createElement("div");
    div.className = "mall-item";

    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.id = `mall-${index}`;
    checkbox.value = mall;
    checkbox.checked = true;

    const label = document.createElement("label");
    label.htmlFor = `mall-${index}`;
    label.textContent = mall;

    div.appendChild(checkbox);
    div.appendChild(label);
    mallList.appendChild(div);
  });
}

/**
 * 요약 카드를 렌더링합니다.
 */
function renderSummaryCards(summary) {
  const container = document.querySelector(".summary-cards");
  if (!container) return;

  container.innerHTML = `
        <div class="summary-card">
            <div class="summary-card-title">총 판매 수량</div>
            <div class="summary-card-value">${formatNumber(summary.total_quantity)}</div>
            <div class="summary-card-subtitle">${summary.period.days}일간</div>
        </div>
        
        <div class="summary-card secondary">
            <div class="summary-card-title">총 상품 수</div>
            <div class="summary-card-value">${formatNumber(summary.total_products)}</div>
            <div class="summary-card-subtitle">개</div>
        </div>
        
        <div class="summary-card accent">
            <div class="summary-card-title">총 주문 건수</div>
            <div class="summary-card-value">${formatNumber(summary.total_orders)}</div>
            <div class="summary-card-subtitle">건</div>
        </div>
        
        <div class="summary-card">
            <div class="summary-card-title">상위 상품</div>
            <div class="summary-card-value">${summary.top_product ? summary.top_product.name : "-"}</div>
            <div class="summary-card-subtitle">${summary.top_product ? formatNumber(summary.top_product.quantity) + "개" : ""}</div>
        </div>
        
        <div class="summary-card secondary">
            <div class="summary-card-title">상위 쇼핑몰</div>
            <div class="summary-card-value">${summary.top_mall ? summary.top_mall.name : "-"}</div>
            <div class="summary-card-subtitle">${summary.top_mall ? formatNumber(summary.top_mall.quantity) + "개" : ""}</div>
        </div>
    `;
}

/**
 * 데이터 테이블을 렌더링합니다.
 */
function renderTable(tableId, pivotData) {
  const table = document.getElementById(tableId);
  if (!table) return;

  if (!pivotData || !pivotData.data || pivotData.data.length === 0) {
    table.innerHTML =
      '<tr><td colspan="100" class="text-center">데이터가 없습니다.</td></tr>';
    return;
  }

  // 헤더 생성
  const thead = document.createElement("thead");
  const headerRow = document.createElement("tr");

  pivotData.columns.forEach((col) => {
    const th = document.createElement("th");
    th.textContent = col;
    headerRow.appendChild(th);
  });

  thead.appendChild(headerRow);

  // 바디 생성
  const tbody = document.createElement("tbody");

  pivotData.data.forEach((row, index) => {
    const tr = document.createElement("tr");

    // 첫 번째 행이 "합계"면 total-row 클래스 추가
    if (index === 0 && Object.values(row)[0] === "합계") {
      tr.className = "total-row";
    }

    pivotData.columns.forEach((col) => {
      const td = document.createElement("td");
      const value = row[col];

      // 숫자면 포맷팅
      if (typeof value === "number") {
        td.textContent = formatNumber(value);
      } else {
        td.textContent = value !== null && value !== undefined ? value : "";
      }

      tr.appendChild(td);
    });

    tbody.appendChild(tr);
  });

  table.innerHTML = "";
  table.appendChild(thead);
  table.appendChild(tbody);
}

/**
 * 페이지네이션을 렌더링합니다.
 */
function renderPagination(paginationId, pivotData, tabKey) {
  const container = document.getElementById(paginationId);
  if (!container) return;

  if (!pivotData || pivotData.total === 0) {
    container.innerHTML = "";
    return;
  }

  const { page, per_page, total, total_pages } = pivotData;

  // per_page가 0이면 페이지네이션 불필요
  if (per_page === 0) {
    container.innerHTML = `<span class="pagination-info">전체 ${total}행</span>`;
    return;
  }

  let html = "";

  // 이전 버튼
  html += `<button class="pagination-button" ${page <= 1 ? "disabled" : ""} 
             onclick="changePage('${tabKey}', ${page - 1})">◀</button>`;

  // 페이지 번호
  const maxButtons = 5;
  let startPage = Math.max(1, page - Math.floor(maxButtons / 2));
  let endPage = Math.min(total_pages, startPage + maxButtons - 1);

  if (endPage - startPage < maxButtons - 1) {
    startPage = Math.max(1, endPage - maxButtons + 1);
  }

  for (let i = startPage; i <= endPage; i++) {
    html += `<button class="pagination-button ${i === page ? "active" : ""}" 
                 onclick="changePage('${tabKey}', ${i})">${i}</button>`;
  }

  // 다음 버튼
  html += `<button class="pagination-button" ${page >= total_pages ? "disabled" : ""} 
             onclick="changePage('${tabKey}', ${page + 1})">▶</button>`;

  // 정보 표시
  html += `<span class="pagination-info">${(page - 1) * per_page + 1}-${Math.min(page * per_page, total)} / ${total}행</span>`;

  container.innerHTML = html;
}

/**
 * 일자별 추이 차트를 렌더링합니다.
 */
function renderDailyTrendChart(pivotData) {
  const canvas = document.getElementById("chartDailyTrend");
  if (!canvas) return;

  // 기존 차트 제거
  if (state.charts["dailyTrend"]) {
    state.charts["dailyTrend"].destroy();
  }

  if (!pivotData || !pivotData.data || pivotData.data.length <= 1) {
    return;
  }

  // 합계 행 제외 (첫 번째 행)
  const data = pivotData.data.slice(1);

  // 레이블 (날짜) - YYYY/MM/DD 형식 보장
  const dateCol = pivotData.columns[0];
  const labels = data.map((row) => {
    const dateValue = row[dateCol];
    // 이미 YYYY/MM/DD 형식이면 그대로, 아니면 변환
    if (typeof dateValue === 'string' && dateValue.includes('/')) {
      return dateValue;
    }
    // Date 객체인 경우 변환
    if (dateValue instanceof Date) {
      return formatDate(dateValue);
    }
    // ISO 형식 문자열인 경우 변환
    if (typeof dateValue === 'string' && dateValue.includes('-')) {
      const parts = dateValue.split('-');
      return `${parts[0]}/${parts[1]}/${parts[2]}`;
    }
    return dateValue;
  });

  // 데이터셋 (쇼핑몰별)
  const datasets = [];
  const colors = [
    "#1976D2",
    "#43A047",
    "#F57C00",
    "#D32F2F",
    "#7B1FA2",
    "#00796B",
    "#C2185B",
    "#0288D1",
    "#689F38",
    "#FFA000",
  ];

  pivotData.columns.slice(1).forEach((mall, index) => {
    if (mall === "합계") return; // 합계 컬럼 제외

    datasets.push({
      label: mall,
      data: data.map((row) => row[mall] || 0),
      borderColor: colors[index % colors.length],
      backgroundColor: colors[index % colors.length] + "20",
      tension: 0.3,
    });
  });

  const ctx = canvas.getContext("2d");
  state.charts["dailyTrend"] = new Chart(ctx, {
    type: "line",
    data: { labels, datasets },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: {
          position: "bottom",
        },
      },
      scales: {
        y: {
          beginAtZero: true,
        },
      },
    },
  });
}

/**
 * 쇼핑몰별 비교 차트를 렌더링합니다.
 */
function renderMallComparisonChart(pivotData) {
  const canvas = document.getElementById("chartMallComparison");
  if (!canvas) return;

  // 기존 차트 제거
  if (state.charts["mallComparison"]) {
    state.charts["mallComparison"].destroy();
  }

  if (!pivotData || !pivotData.data || pivotData.data.length === 0) {
    return;
  }

  // 합계 행 (첫 번째 행) 사용
  const totalRow = pivotData.data[0];

  // 레이블 및 데이터
  const labels = [];
  const data = [];
  const colors = [
    "#1976D2",
    "#43A047",
    "#F57C00",
    "#D32F2F",
    "#7B1FA2",
    "#00796B",
    "#C2185B",
    "#0288D1",
    "#689F38",
    "#FFA000",
  ];

  pivotData.columns.slice(1).forEach((mall, index) => {
    if (mall === "합계") return; // 합계 컬럼 제외

    labels.push(mall);
    data.push(totalRow[mall] || 0);
  });

  const ctx = canvas.getContext("2d");
  state.charts["mallComparison"] = new Chart(ctx, {
    type: "bar",
    data: {
      labels,
      datasets: [
        {
          label: "판매량",
          data,
          backgroundColor: colors.slice(0, labels.length),
        },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: {
          display: false,
        },
      },
      scales: {
        y: {
          beginAtZero: true,
        },
      },
    },
  });
}

/**
 * 예측 결과를 렌더링합니다.
 */
function renderPredictionResults(mallData, productData) {
  const container = document.getElementById("predictionResults");
  if (!container) return;

  let html = "";

  // 쇼핑몰별 예측
  if (mallData && mallData.by_mall) {
    html += `
            <div class="prediction-card">
                <h4 class="card-title">쇼핑몰별 예측 (향후 2주)</h4>
                <div class="table-container">
                    <table class="data-table">
                        <thead>
                            <tr>
                                <th>쇼핑몰</th>
                                <th>예측 판매량</th>
                            </tr>
                        </thead>
                        <tbody>
        `;

    for (const [mall, data] of Object.entries(mallData.by_mall)) {
      html += `
                <tr>
                    <td>${mall}</td>
                    <td>${formatNumber(Math.round(data.total_predicted))}</td>
                </tr>
            `;
    }

    html += `
                            <tr class="total-row">
                                <td>합계</td>
                                <td>${formatNumber(Math.round(mallData.total_sum))}</td>
                            </tr>
                        </tbody>
                    </table>
                </div>
            </div>
        `;
  }

  // 상품별 예측
  if (productData && productData.by_product) {
    html += `
            <div class="prediction-card">
                <h4 class="card-title">상위 상품 예측 (향후 2주)</h4>
                <div class="table-container">
                    <table class="data-table">
                        <thead>
                            <tr>
                                <th>상품코드</th>
                                <th>예측 판매량</th>
                                <th>과거 판매량</th>
                            </tr>
                        </thead>
                        <tbody>
        `;

    for (const [product, data] of Object.entries(productData.by_product)) {
      html += `
                <tr>
                    <td>${product}</td>
                    <td>${formatNumber(Math.round(data.total_predicted))}</td>
                    <td>${formatNumber(data.historical_total)}</td>
                </tr>
            `;
    }

    html += `
                        </tbody>
                    </table>
                </div>
            </div>
        `;
  }

  if (!html) {
    html =
      '<p class="text-center">예측 결과가 없습니다. "예측 실행" 버튼을 클릭하세요.</p>';
  }

  container.innerHTML = html;
}

// ===========================================
// 이벤트 핸들러
// ===========================================

/**
 * 분석을 실행합니다.
 */
async function handleAnalyze() {
  try {
    showLoading(true);
    setStatusMessage("분석 중...");

    // 입력값 수집
    const startDate = document.getElementById("startDate").value;
    const endDate = document.getElementById("endDate").value;
    const perPage = parseInt(document.getElementById("perPage").value);

    // 선택된 쇼핑몰
    const mallCheckboxes = document.querySelectorAll(
      '.mall-item input[type="checkbox"]:checked',
    );
    const malls = Array.from(mallCheckboxes).map((cb) => cb.value);

    // 검증
    if (!startDate || !endDate) {
      alert("시작일과 종료일을 입력하세요.");
      return;
    }

    // API 호출
    const result = await runAnalysis({
      start_date: startDate,
      end_date: endDate,
      malls: malls.length > 0 ? malls : null,
      page: 1,
      per_page: perPage,
    });

    // 상태 저장
    state.analysisData = result;

    // 페이지 초기화
    state.currentPage = {
      "product-daily": 1,
      "product-mall": 1,
      "daily-mall": 1,
      "yearly-mall": 1,
    };

    // UI 렌더링
    renderSummaryCards(result.summary);
    renderTable("table-product-daily", result.pivots.product_daily);
    renderTable("table-product-mall", result.pivots.product_mall);
    renderTable("table-daily-mall", result.pivots.daily_mall);
    renderTable("table-yearly-mall", result.pivots.yearly_mall);

    renderPagination(
      "pagination-product-daily",
      result.pivots.product_daily,
      "product-daily",
    );
    renderPagination(
      "pagination-product-mall",
      result.pivots.product_mall,
      "product-mall",
    );
    renderPagination(
      "pagination-daily-mall",
      result.pivots.daily_mall,
      "daily-mall",
    );
    renderPagination(
      "pagination-yearly-mall",
      result.pivots.yearly_mall,
      "yearly-mall",
    );

    renderDailyTrendChart(result.pivots.daily_mall);
    renderMallComparisonChart(result.pivots.daily_mall);

    setStatusMessage(
      `분석 완료 (${formatNumber(result.summary.total_orders)}건 처리)`,
    );
  } catch (error) {
    alert(`분석 실패: ${error.message}`);
    setStatusMessage("분석 실패");
  } finally {
    showLoading(false);
  }
}

/**
 * 페이지를 변경합니다.
 */
async function changePage(tabKey, page) {
  try {
    if (!state.analysisData) return;

    showLoading(true);

    // 현재 페이지 업데이트
    state.currentPage[tabKey] = page;

    // 입력값 수집
    const startDate = document.getElementById("startDate").value;
    const endDate = document.getElementById("endDate").value;
    const perPage = parseInt(document.getElementById("perPage").value);

    const mallCheckboxes = document.querySelectorAll(
      '.mall-item input[type="checkbox"]:checked',
    );
    const malls = Array.from(mallCheckboxes).map((cb) => cb.value);

    // API 호출
    const result = await runAnalysis({
      start_date: startDate,
      end_date: endDate,
      malls: malls.length > 0 ? malls : null,
      page: page,
      per_page: perPage,
    });

    // 해당 탭만 업데이트
    const tableMap = {
      "product-daily": "table-product-daily",
      "product-mall": "table-product-mall",
      "daily-mall": "table-daily-mall",
      "yearly-mall": "table-yearly-mall",
    };

    const paginationMap = {
      "product-daily": "pagination-product-daily",
      "product-mall": "pagination-product-mall",
      "daily-mall": "pagination-daily-mall",
      "yearly-mall": "pagination-yearly-mall",
    };

    const pivotKey = tabKey.replace("-", "_");

    renderTable(tableMap[tabKey], result.pivots[pivotKey]);
    renderPagination(paginationMap[tabKey], result.pivots[pivotKey], tabKey);
  } catch (error) {
    alert(`페이지 변경 실패: ${error.message}`);
  } finally {
    showLoading(false);
  }
}

/**
 * 예측을 실행합니다.
 */
async function handlePredict() {
  try {
    showLoading(true);
    setStatusMessage("예측 중...");

    // 입력값 수집
    const startDate = document.getElementById("startDate").value;
    const endDate = document.getElementById("endDate").value;

    if (!startDate || !endDate) {
      alert("시작일과 종료일을 입력하세요.");
      return;
    }

    // 쇼핑몰별 예측
    const mallPrediction = await predictMall({
      start_date: startDate,
      end_date: endDate,
      days: 14,
    });

    // 상품별 예측
    const productPrediction = await predictProduct({
      start_date: startDate,
      end_date: endDate,
      days: 14,
      top_n: 10,
    });

    // 상태 저장
    state.predictionData = {
      mall: mallPrediction,
      product: productPrediction,
    };

    // UI 렌더링
    renderPredictionResults(mallPrediction, productPrediction);

    setStatusMessage("예측 완료");
  } catch (error) {
    alert(`예측 실패: ${error.message}`);
    setStatusMessage("예측 실패");
  } finally {
    showLoading(false);
  }
}

/**
 * Excel 파일을 다운로드합니다.
 */
async function handleExport() {
  try {
    if (!state.analysisData) {
      alert("먼저 분석을 실행하세요.");
      return;
    }

    showLoading(true);
    setStatusMessage("Excel 생성 중...");

    // 입력값 수집
    const startDate = document.getElementById("startDate").value;
    const endDate = document.getElementById("endDate").value;

    const mallCheckboxes = document.querySelectorAll(
      '.mall-item input[type="checkbox"]:checked',
    );
    const malls = Array.from(mallCheckboxes).map((cb) => cb.value);

    // Excel 다운로드
    await downloadExcel({
      start_date: startDate,
      end_date: endDate,
      malls: malls.length > 0 ? malls : null,
    });

    setStatusMessage("Excel 다운로드 완료");
  } catch (error) {
    alert(`Excel 다운로드 실패: ${error.message}`);
    setStatusMessage("다운로드 실패");
  } finally {
    showLoading(false);
  }
}

/**
 * 초기화합니다.
 */
function handleReset() {
  // 날짜 초기화 (최근 1주일)
  setDateRange(7);

  // 쇼핑몰 전체 선택
  document.getElementById("btnSelectAll").click();

  // 페이지 설정 초기화
  document.getElementById("perPage").value = "20";

  setStatusMessage("초기화 완료");
}

/**
 * 날짜 범위를 설정합니다.
 */
function setDateRange(days) {
  const startDate = getDaysAgo(days);
  const endDate = getToday();

  document.getElementById("startDate").value = formatDate(startDate);
  document.getElementById("endDate").value = formatDate(endDate);
}

/**
 * 탭을 전환합니다.
 */
function switchTab(tabName) {
  // 탭 버튼 활성화
  document.querySelectorAll(".tab-button").forEach((btn) => {
    btn.classList.remove("active");
  });
  document.querySelector(`[data-tab="${tabName}"]`).classList.add("active");

  // 탭 콘텐츠 표시
  document.querySelectorAll(".tab-pane").forEach((pane) => {
    pane.classList.remove("active");
  });
  document.getElementById(`tab-${tabName}`).classList.add("active");

  state.currentTab = tabName;
}

// ===========================================
// 초기화
// ===========================================

/**
 * 애플리케이션을 초기화합니다.
 */
async function initApp() {
  try {
    // 데이터 정보 조회
    const info = await fetchDataInfo();

    if (info && info.loaded) {
      const statusEl = document.getElementById("dataStatus");
      const indicator = statusEl.querySelector(".status-indicator");
      const text = statusEl.querySelector(".status-text");

      if (info.is_dummy) {
        indicator.style.background = "#F57C00";
        text.textContent = "더미 데이터";
      } else {
        indicator.style.background = "#4CAF50";
        text.textContent = `데이터 로드됨 (${formatNumber(info.row_count)}행)`;
      }
    }

    // 쇼핑몰 목록 조회 및 렌더링
    const malls = await fetchMalls();
    renderMallList(malls);

    // 날짜 초기화 (최근 1주일)
    setDateRange(7);

    // 이벤트 리스너 등록
    document
      .getElementById("btnAnalyze")
      .addEventListener("click", handleAnalyze);
    document
      .getElementById("btnPredict")
      .addEventListener("click", handlePredict);
    document
      .getElementById("btnExport")
      .addEventListener("click", handleExport);
    document.getElementById("btnReset").addEventListener("click", handleReset);

    // 쇼핑몰 선택 버튼
    document.getElementById("btnSelectAll").addEventListener("click", () => {
      document
        .querySelectorAll('.mall-item input[type="checkbox"]')
        .forEach((cb) => {
          cb.checked = true;
        });
    });

    document.getElementById("btnDeselectAll").addEventListener("click", () => {
      document
        .querySelectorAll('.mall-item input[type="checkbox"]')
        .forEach((cb) => {
          cb.checked = false;
        });
    });

    // 빠른 기간 선택 버튼
    document.querySelectorAll(".btn-period").forEach((btn) => {
      btn.addEventListener("click", () => {
        const days = parseInt(btn.dataset.days);
        setDateRange(days);
      });
    });

    // 탭 전환
    document.querySelectorAll(".tab-button").forEach((btn) => {
      btn.addEventListener("click", () => {
        switchTab(btn.dataset.tab);
      });
    });

    // 초기 분석 실행
    await handleAnalyze();

    setStatusMessage("준비됨");
  } catch (error) {
    console.error("초기화 오류:", error);
    setStatusMessage("초기화 실패");
  }
}

// DOM 로드 완료 후 초기화
document.addEventListener("DOMContentLoaded", initApp);
