const TRACKS = ["이과", "문과"];
const MAX_RENDERED_ROWS = 300;

const STATUS_PRIORITY = {
  "적정점수 이상": 4,
  "예상점수 이상": 3,
  "소신점수 이상": 2,
  "소신점수 미만": 1,
  "제외(수탐결격)": 0,
  "오류(영어국사)": -1
};

const SHEET_NAMES = {
  exam: "수능입력",
  school: "내신입력",
  science: "이과계열분석결과",
  humanities: "문과계열분석결과"
};

const numberFormatter = new Intl.NumberFormat("ko-KR", {
  maximumFractionDigits: 3
});

const elements = {
  uploadStatus: document.getElementById("upload-status"),
  fileInput: document.getElementById("workbook-input"),
  dropzone: document.getElementById("dropzone"),
  clearButton: document.getElementById("clear-button"),
  fileList: document.getElementById("file-list"),
  activeYearBadge: document.getElementById("active-year-badge"),
  summaryGrid: document.getElementById("summary-grid"),
  examInputs: document.getElementById("exam-inputs"),
  schoolInputs: document.getElementById("school-inputs"),
  comparePanel: document.getElementById("compare-panel"),
  compareBadge: document.getElementById("compare-badge"),
  compareGrid: document.getElementById("compare-grid"),
  compareList: document.getElementById("compare-list"),
  trackTabs: document.getElementById("track-tabs"),
  yearFilter: document.getElementById("year-filter"),
  statusFilter: document.getElementById("status-filter"),
  groupFilter: document.getElementById("group-filter"),
  regionFilter: document.getElementById("region-filter"),
  sortFilter: document.getElementById("sort-filter"),
  searchInput: document.getElementById("search-input"),
  resultCount: document.getElementById("result-count"),
  resultBody: document.getElementById("result-body"),
  resultFootnote: document.getElementById("result-footnote"),
  detailBody: document.getElementById("detail-body")
};

const state = {
  analyses: new Map(),
  activeYearKey: null,
  track: TRACKS[0],
  status: "all",
  group: "all",
  region: "all",
  sort: "status",
  search: "",
  filteredRows: [],
  selectedProgramKey: null
};

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

function toNumber(value) {
  if (typeof value === "number") {
    return Number.isFinite(value) ? value : null;
  }

  if (typeof value !== "string") {
    return null;
  }

  const normalized = value.replaceAll(",", "").trim();
  if (!normalized || normalized === "-") {
    return null;
  }

  const parsed = Number(normalized);
  return Number.isFinite(parsed) ? parsed : null;
}

function formatNumber(value) {
  if (value === null || value === undefined || Number.isNaN(value)) {
    return "-";
  }

  return numberFormatter.format(value);
}

function formatDelta(value) {
  if (value === null || value === undefined || Number.isNaN(value)) {
    return { text: "-", tone: "neutral" };
  }

  if (value > 0) {
    return { text: `+${formatNumber(value)}`, tone: "up" };
  }

  if (value < 0) {
    return { text: formatNumber(value), tone: "down" };
  }

  return { text: "0", tone: "neutral" };
}

function calcDelta(left, right) {
  if (left === null || right === null) {
    return null;
  }

  return left - right;
}

function normalizeText(value) {
  return String(value || "").trim().toLowerCase();
}

function cellAddress(row, column) {
  return XLSX.utils.encode_cell({ r: row, c: column });
}

function getCell(sheet, row, column) {
  return sheet?.[cellAddress(row, column)] || null;
}

function getCellValue(sheet, address) {
  return sheet?.[address]?.v ?? null;
}

function getCellText(sheet, address) {
  const value = getCellValue(sheet, address);
  return value === null || value === undefined ? "" : String(value);
}

function requireSheet(workbook, name) {
  const sheet = workbook.Sheets[name];

  if (!sheet) {
    throw new Error(`필수 시트가 없습니다: ${name}`);
  }

  return sheet;
}

function detectAnalyzerYear(fileName, workbook) {
  const clues = [
    fileName,
    getCellText(workbook.Sheets.INFO, "A1"),
    getCellText(workbook.Sheets[SHEET_NAMES.exam], "A1")
  ].join(" ");

  if (clues.includes("202511") || clues.includes("2026학년도")) {
    return { key: "26", label: "26수능" };
  }

  if (clues.includes("202411") || clues.includes("2025학년도")) {
    return { key: "25", label: "25수능" };
  }

  return fileName.toLowerCase().endsWith(".xlsb")
    ? { key: "26", label: "26수능" }
    : { key: "25", label: "25수능" };
}

function getAnalyzerStamp(workbook) {
  const infoSheet = workbook.Sheets.INFO;
  if (infoSheet && getCellText(infoSheet, "A1")) {
    return getCellText(infoSheet, "A1");
  }

  return getCellText(workbook.Sheets[SHEET_NAMES.exam], "A1");
}

function buildHeaderIndex(headerRow) {
  const index = {};

  headerRow.forEach((label, position) => {
    if (label === null || label === undefined) {
      return;
    }

    const key = String(label).trim();
    if (!key || key in index) {
      return;
    }

    index[key] = position;
  });

  return index;
}

function readByHeader(row, headerIndex, label) {
  const position = headerIndex[label];
  return position === undefined ? null : row[position];
}

function firstHeader(headerIndex, predicate) {
  return Object.keys(headerIndex).find(predicate);
}

function extractExamInputs(workbook) {
  const sheet = requireSheet(workbook, SHEET_NAMES.exam);
  const range = XLSX.utils.decode_range(sheet["!ref"]);
  const items = [];
  let currentArea = "";

  for (let row = 0; row <= Math.min(range.e.r, 120); row += 1) {
    const areaCell = getCell(sheet, row, 0);
    const subjectCell = getCell(sheet, row, 1);
    const inputCell = getCell(sheet, row, 2);

    if (areaCell && !areaCell.f && typeof areaCell.v === "string") {
      currentArea = areaCell.v.trim();
    }

    if (!subjectCell || !inputCell || inputCell.f) {
      continue;
    }

    if (inputCell.v === "" || inputCell.v === null || inputCell.v === undefined) {
      continue;
    }

    if (currentArea === "영역" || currentArea.includes("수능 점수 입력")) {
      continue;
    }

    items.push({
      area: currentArea || "기타",
      subject: String(subjectCell.v).trim(),
      value: inputCell.v
    });
  }

  return items;
}

function extractSchoolInputs(workbook) {
  const sheet = requireSheet(workbook, SHEET_NAMES.school);
  const range = XLSX.utils.decode_range(sheet["!ref"]);
  const items = [];

  for (let row = 0; row <= Math.min(range.e.r, 100); row += 1) {
    const universityCell = getCell(sheet, row, 0);
    const formulaCell = getCell(sheet, row, 1);
    const directCell = getCell(sheet, row, 3);

    if (!formulaCell || !directCell || directCell.f) {
      continue;
    }

    if (directCell.v === "" || directCell.v === null || directCell.v === undefined) {
      continue;
    }

    items.push({
      university: universityCell?.v ? String(universityCell.v).trim() : "",
      formula: String(formulaCell.v).trim(),
      value: directCell.v
    });
  }

  return items;
}

function buildProgramKey(record) {
  return [
    record.track,
    record.university,
    record.major,
    record.selectionType,
    record.admissionGroup
  ].map(normalizeText).join("||");
}

function extractResults(workbook, sheetName) {
  const sheet = requireSheet(workbook, sheetName);
  const matrix = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: null,
    blankrows: false,
    raw: true
  });

  const headerRow = matrix[4] || [];
  const headerIndex = buildHeaderIndex(headerRow);
  const nationalRankHeader = firstHeader(headerIndex, (label) => label.startsWith("전국등수"));
  const rows = [];

  for (let index = 5; index < matrix.length; index += 1) {
    const row = matrix[index];
    const university = readByHeader(row, headerIndex, "대학교");
    const major = readByHeader(row, headerIndex, "전공");

    if (!university || !major) {
      continue;
    }

    const totalScore = toNumber(readByHeader(row, headerIndex, "수능+내신"));
    const expectedScore = toNumber(readByHeader(row, headerIndex, "예상점수"));
    const safeScore = toNumber(readByHeader(row, headerIndex, "적정점수"));
    const reachScore = toNumber(readByHeader(row, headerIndex, "소신점수"));

    const record = {
      rowNumber: index + 1,
      track: String(readByHeader(row, headerIndex, "계열") || "").trim() || (sheetName.includes("이과") ? "이과" : "문과"),
      university: String(university).trim(),
      major: String(major).trim(),
      status: String(readByHeader(row, headerIndex, "합격가능성") || "").trim(),
      examScore: toNumber(readByHeader(row, headerIndex, "수능점수")),
      schoolScore: toNumber(readByHeader(row, headerIndex, "내신점수")),
      totalScore,
      percentile: toNumber(readByHeader(row, headerIndex, "누백")),
      nationalRank: nationalRankHeader ? String(readByHeader(row, headerIndex, nationalRankHeader) || "").trim() : "",
      safeScore,
      expectedScore,
      reachScore,
      safePercentile: toNumber(readByHeader(row, headerIndex, "적정누백")),
      expectedPercentile: toNumber(readByHeader(row, headerIndex, "예상누백")),
      reachPercentile: toNumber(readByHeader(row, headerIndex, "소신누백")),
      universityType: String(readByHeader(row, headerIndex, "대학구분") || "").trim(),
      admissionGroup: String(readByHeader(row, headerIndex, "모집군") || "").trim(),
      capacity: readByHeader(row, headerIndex, "정원"),
      region: String(readByHeader(row, headerIndex, "시도") || "").trim(),
      city: String(readByHeader(row, headerIndex, "시군") || "").trim(),
      category: String(readByHeader(row, headerIndex, "학과특성") || "").trim(),
      universityShort: String(readByHeader(row, headerIndex, "대학교(약칭)") || "").trim(),
      majorShort: String(readByHeader(row, headerIndex, "전공(약칭)") || "").trim(),
      selectionType: String(readByHeader(row, headerIndex, "선발유형") || "").trim(),
      scoreMethod: String(readByHeader(row, headerIndex, "점수환산") || "").trim(),
      subjectRule: String(readByHeader(row, headerIndex, "수탐선택") || "").trim(),
      examElements: String(readByHeader(row, headerIndex, "수능요소") || "").trim(),
      examCombo: String(readByHeader(row, headerIndex, "수능조합") || "").trim(),
      requiredSubjects: String(readByHeader(row, headerIndex, "필수") || "").trim(),
      optionalSubjects: String(readByHeader(row, headerIndex, "선택") || "").trim(),
      weightedOption: String(readByHeader(row, headerIndex, "가중택") || "").trim(),
      inquiryCount: readByHeader(row, headerIndex, "탐구과목수"),
      koreanWeight: toNumber(readByHeader(row, headerIndex, "국어배점")),
      mathWeight: toNumber(readByHeader(row, headerIndex, "수학배점")),
      inquiryWeight: toNumber(readByHeader(row, headerIndex, "탐구배점")),
      koreanRatio: toNumber(readByHeader(row, headerIndex, "국어구성비")),
      mathRatio: toNumber(readByHeader(row, headerIndex, "수학구성비")),
      inquiryRatio: toNumber(readByHeader(row, headerIndex, "탐구구성비")),
      englishAdjustments: Array.from({ length: 9 }, (_, offset) => {
        const label = `영어${offset + 1}환점`;
        return {
          grade: offset + 1,
          value: toNumber(readByHeader(row, headerIndex, label))
        };
      })
    };

    record.deltaSafe = calcDelta(record.totalScore, record.safeScore);
    record.deltaExpected = calcDelta(record.totalScore, record.expectedScore);
    record.deltaReach = calcDelta(record.totalScore, record.reachScore);
    record.statusPriority = STATUS_PRIORITY[record.status] ?? -2;
    record.programKey = buildProgramKey(record);

    rows.push(record);
  }

  return rows;
}

function countByStatus(rows) {
  const counts = {};

  for (const row of rows) {
    counts[row.status] = (counts[row.status] || 0) + 1;
  }

  return counts;
}

function extractAnalysis(file, workbook) {
  const year = detectAnalyzerYear(file.name, workbook);
  const scienceRows = extractResults(workbook, SHEET_NAMES.science);
  const humanitiesRows = extractResults(workbook, SHEET_NAMES.humanities);
  const allRows = [...scienceRows, ...humanitiesRows];

  return {
    yearKey: year.key,
    yearLabel: year.label,
    fileName: file.name,
    fileSize: file.size,
    analyzerStamp: getAnalyzerStamp(workbook),
    examInputs: extractExamInputs(workbook),
    schoolInputs: extractSchoolInputs(workbook),
    results: {
      이과: scienceRows,
      문과: humanitiesRows
    },
    allRows,
    counts: countByStatus(allRows)
  };
}

function getSortedAnalyses() {
  return [...state.analyses.values()].sort((left, right) => Number(right.yearKey) - Number(left.yearKey));
}

function getActiveAnalysis() {
  if (!state.analyses.size) {
    return null;
  }

  if (state.activeYearKey && state.analyses.has(state.activeYearKey)) {
    return state.analyses.get(state.activeYearKey);
  }

  const first = getSortedAnalyses()[0];
  state.activeYearKey = first.yearKey;
  return first;
}

function setUploadStatus(text) {
  elements.uploadStatus.textContent = text;
}

function createEmptyState(text) {
  const div = document.createElement("div");
  div.className = "empty-state";
  div.textContent = text;
  return div;
}

function renderFileList() {
  elements.fileList.replaceChildren();

  for (const analysis of getSortedAnalyses()) {
    const chip = document.createElement("div");
    chip.className = "file-chip";
    chip.innerHTML = `
      <strong>${escapeHtml(analysis.yearLabel)}</strong>
      <div class="muted">${escapeHtml(analysis.fileName)}</div>
      <div class="muted">이과 ${formatNumber(analysis.results["이과"].length)}개 · 문과 ${formatNumber(analysis.results["문과"].length)}개</div>
    `;
    elements.fileList.append(chip);
  }
}

function renderTrackTabs() {
  elements.trackTabs.replaceChildren();

  for (const track of TRACKS) {
    const button = document.createElement("button");
    button.type = "button";
    button.textContent = track;
    button.className = track === state.track ? "active" : "";
    button.addEventListener("click", () => {
      state.track = track;
      state.selectedProgramKey = null;
      renderAll();
    });
    elements.trackTabs.append(button);
  }
}

function fillSelect(select, options, currentValue, allLabel) {
  const nextValue = options.includes(currentValue) ? currentValue : "all";

  select.replaceChildren();

  const allOption = document.createElement("option");
  allOption.value = "all";
  allOption.textContent = allLabel;
  select.append(allOption);

  for (const optionValue of options) {
    const option = document.createElement("option");
    option.value = optionValue;
    option.textContent = optionValue;
    select.append(option);
  }

  select.value = nextValue;
}

function renderFilters() {
  const analyses = getSortedAnalyses();
  const active = getActiveAnalysis();
  const rows = active ? active.results[state.track] : [];

  fillSelect(
    elements.yearFilter,
    analyses.map((analysis) => analysis.yearKey),
    state.activeYearKey,
    "선택된 최신 연도"
  );

  const statuses = [...new Set(rows.map((row) => row.status))].sort(
    (left, right) => (STATUS_PRIORITY[right] ?? -2) - (STATUS_PRIORITY[left] ?? -2)
  );
  const groups = [...new Set(rows.map((row) => row.admissionGroup).filter(Boolean))].sort();
  const regions = [...new Set(rows.map((row) => row.region).filter(Boolean))].sort();

  fillSelect(elements.statusFilter, statuses, state.status, "전체 상태");
  fillSelect(elements.groupFilter, groups, state.group, "전체 모집군");
  fillSelect(elements.regionFilter, regions, state.region, "전체 지역");
  elements.sortFilter.value = state.sort;
  elements.searchInput.value = state.search;
}

function renderSummary() {
  const analysis = getActiveAnalysis();
  elements.summaryGrid.replaceChildren();
  elements.examInputs.replaceChildren();
  elements.schoolInputs.replaceChildren();

  if (!analysis) {
    elements.activeYearBadge.textContent = "연도 미선택";
    elements.summaryGrid.append(createEmptyState("분석기 파일을 올리면 업로드 요약이 여기에 표시됩니다."));
    elements.examInputs.append(createEmptyState("수능 입력값이 여기에 표시됩니다."));
    elements.schoolInputs.append(createEmptyState("내신 입력값이 여기에 표시됩니다."));
    return;
  }

  elements.activeYearBadge.textContent = `${analysis.yearLabel} · ${analysis.fileName}`;

  const summaryItems = [
    { label: "전체 모집단위", value: analysis.allRows.length, note: "업로드된 결과 행 수" },
    { label: "이과 결과", value: analysis.results["이과"].length, note: "이과계열분석결과" },
    { label: "문과 결과", value: analysis.results["문과"].length, note: "문과계열분석결과" },
    { label: "적정점수 이상", value: analysis.counts["적정점수 이상"] || 0, note: "업로드 파일 기준" },
    { label: "예상점수 이상", value: analysis.counts["예상점수 이상"] || 0, note: "업로드 파일 기준" },
    { label: "소신점수 이상", value: analysis.counts["소신점수 이상"] || 0, note: "업로드 파일 기준" }
  ];

  for (const item of summaryItems) {
    const card = document.createElement("article");
    card.className = "metric-card";
    card.innerHTML = `
      <strong>${escapeHtml(item.label)}</strong>
      <div class="value">${formatNumber(item.value)}</div>
      <small>${escapeHtml(item.note)}</small>
    `;
    elements.summaryGrid.append(card);
  }

  if (analysis.examInputs.length) {
    for (const item of analysis.examInputs) {
      const pill = document.createElement("div");
      pill.className = "pill";
      pill.textContent = `${item.subject} ${formatNumber(item.value)}`;
      elements.examInputs.append(pill);
    }
  } else {
    elements.examInputs.append(createEmptyState("엑셀 파일에서 직접 입력된 수능 점수를 찾지 못했습니다."));
  }

  if (analysis.schoolInputs.length) {
    for (const item of analysis.schoolInputs) {
      const pill = document.createElement("div");
      pill.className = "pill";
      pill.textContent = `${item.formula} ${item.value}`;
      elements.schoolInputs.append(pill);
    }
  } else {
    elements.schoolInputs.append(createEmptyState("직접 입력된 내신값이 없거나 기본값만 사용된 상태입니다."));
  }
}

function getFilteredRows() {
  const analysis = getActiveAnalysis();
  if (!analysis) {
    return [];
  }

  const query = normalizeText(state.search);
  const rows = analysis.results[state.track] || [];

  const filtered = rows.filter((row) => {
    if (state.status !== "all" && row.status !== state.status) {
      return false;
    }

    if (state.group !== "all" && row.admissionGroup !== state.group) {
      return false;
    }

    if (state.region !== "all" && row.region !== state.region) {
      return false;
    }

    if (!query) {
      return true;
    }

    const haystack = [
      row.university,
      row.major,
      row.selectionType,
      row.scoreMethod,
      row.universityType,
      row.subjectRule
    ].join(" ").toLowerCase();

    return haystack.includes(query);
  });

  filtered.sort((left, right) => {
    if (state.sort === "delta-desc") {
      return (right.deltaExpected ?? -Infinity) - (left.deltaExpected ?? -Infinity);
    }

    if (state.sort === "delta-asc") {
      return (left.deltaExpected ?? Infinity) - (right.deltaExpected ?? Infinity);
    }

    if (state.sort === "expected-desc") {
      return (right.expectedScore ?? -Infinity) - (left.expectedScore ?? -Infinity);
    }

    if (state.sort === "expected-asc") {
      return (left.expectedScore ?? Infinity) - (right.expectedScore ?? Infinity);
    }

    if (state.sort === "name") {
      return left.university.localeCompare(right.university, "ko");
    }

    if (right.statusPriority !== left.statusPriority) {
      return right.statusPriority - left.statusPriority;
    }

    return (right.deltaExpected ?? -Infinity) - (left.deltaExpected ?? -Infinity);
  });

  return filtered;
}

function statusTone(status) {
  if (status === "적정점수 이상") return "good";
  if (status === "예상점수 이상") return "mid";
  if (status === "소신점수 이상") return "warn";
  if (status === "소신점수 미만" || status.startsWith("오류")) return "bad";
  return "neutral";
}

function renderResults() {
  const filtered = getFilteredRows();
  state.filteredRows = filtered;
  elements.resultBody.replaceChildren();

  if (!filtered.length) {
    const row = document.createElement("tr");
    row.innerHTML = `<td colspan="10">조건에 맞는 결과가 없습니다.</td>`;
    elements.resultBody.append(row);
    elements.resultCount.textContent = "0개 결과";
    elements.resultFootnote.textContent = "";
    state.selectedProgramKey = null;
    renderDetails();
    return;
  }

  if (!filtered.some((row) => row.programKey === state.selectedProgramKey)) {
    state.selectedProgramKey = filtered[0].programKey;
  }

  const visible = filtered.slice(0, MAX_RENDERED_ROWS);

  for (const row of visible) {
    const tr = document.createElement("tr");
    tr.dataset.key = row.programKey;

    if (row.programKey === state.selectedProgramKey) {
      tr.classList.add("selected");
    }

    const delta = formatDelta(row.deltaExpected);

    tr.innerHTML = `
      <td>
        <strong>${escapeHtml(row.university)}</strong>
        <div class="muted">${escapeHtml(row.universityType || "-")}</div>
      </td>
      <td>
        <strong>${escapeHtml(row.major)}</strong>
        <div class="muted">${escapeHtml(row.category || "-")}</div>
      </td>
      <td><span class="status-tag ${statusTone(row.status)}">${escapeHtml(row.status || "-")}</span></td>
      <td>${formatNumber(row.totalScore)}</td>
      <td>${formatNumber(row.expectedScore)}</td>
      <td><span class="delta ${delta.tone}">${delta.text}</span></td>
      <td>${escapeHtml(row.admissionGroup || "-")}</td>
      <td>${escapeHtml(row.region || "-")}</td>
      <td>${escapeHtml(row.selectionType || "-")}</td>
      <td>${escapeHtml(row.scoreMethod || "-")}</td>
    `;

    elements.resultBody.append(tr);
  }

  elements.resultCount.textContent = `${formatNumber(filtered.length)}개 결과`;
  elements.resultFootnote.textContent = filtered.length > visible.length
    ? `필터된 ${formatNumber(filtered.length)}개 중 상위 ${formatNumber(visible.length)}개만 표시했습니다. 검색어나 필터를 더 좁히면 전체를 더 쉽게 볼 수 있습니다.`
    : `${state.track} 결과 ${formatNumber(filtered.length)}개를 표시 중입니다.`;

  renderDetails();
}

function findSelectedRow() {
  return state.filteredRows.find((row) => row.programKey === state.selectedProgramKey) || null;
}

function renderDetailMetric(label, value, note = "") {
  return `
    <article class="detail-card">
      <strong>${escapeHtml(label)}</strong>
      <div class="value">${escapeHtml(value)}</div>
      ${note ? `<div class="muted">${escapeHtml(note)}</div>` : ""}
    </article>
  `;
}

function renderDetails() {
  const row = findSelectedRow();
  elements.detailBody.replaceChildren();

  if (!row) {
    elements.detailBody.append(createEmptyState("업로드 후 결과 행을 선택하면 상세 정보가 여기에 표시됩니다."));
    return;
  }

  const deltaSafe = formatDelta(row.deltaSafe);
  const deltaExpected = formatDelta(row.deltaExpected);
  const deltaReach = formatDelta(row.deltaReach);

  const englishRows = row.englishAdjustments
    .map((item) => `<div class="pill">영어 ${item.grade}등급 ${formatNumber(item.value)}</div>`)
    .join("");

  const wrapper = document.createElement("div");
  wrapper.innerHTML = `
    <div class="detail-card">
      <strong>${escapeHtml(row.university)} ${escapeHtml(row.major)}</strong>
      <div class="muted">${escapeHtml(row.track)} · ${escapeHtml(row.selectionType || "-")} · ${escapeHtml(row.admissionGroup || "-")} · ${escapeHtml(row.region || "-")} ${escapeHtml(row.city || "")}</div>
    </div>

    <div class="detail-grid">
      ${renderDetailMetric("내 점수", formatNumber(row.totalScore), "수능+내신")}
      ${renderDetailMetric("적정점수", formatNumber(row.safeScore), deltaSafe.text)}
      ${renderDetailMetric("예상점수", formatNumber(row.expectedScore), deltaExpected.text)}
      ${renderDetailMetric("소신점수", formatNumber(row.reachScore), deltaReach.text)}
      ${renderDetailMetric("누백", formatNumber(row.percentile), row.nationalRank || "전국등수 정보 없음")}
      ${renderDetailMetric("정원", row.capacity === null ? "-" : formatNumber(row.capacity), row.universityType || "대학구분 없음")}
    </div>

    <div class="detail-card">
      <strong>반영 규칙</strong>
      <div class="muted">점수환산: ${escapeHtml(row.scoreMethod || "-")}</div>
      <div class="muted">수탐선택: ${escapeHtml(row.subjectRule || "-")}</div>
      <div class="muted">수능요소: ${escapeHtml(row.examElements || "-")}</div>
      <div class="muted">수능조합: ${escapeHtml(row.examCombo || "-")}</div>
      <div class="muted">필수: ${escapeHtml(row.requiredSubjects || "-")}</div>
      <div class="muted">선택: ${escapeHtml(row.optionalSubjects || "-")}</div>
      <div class="muted">가중택: ${escapeHtml(row.weightedOption || "-")}</div>
      <div class="muted">탐구과목수: ${escapeHtml(row.inquiryCount ?? "-")}</div>
    </div>

    <div class="detail-card">
      <strong>배점 및 구성비</strong>
      <div class="muted">국어 ${formatNumber(row.koreanWeight)} / 수학 ${formatNumber(row.mathWeight)} / 탐구 ${formatNumber(row.inquiryWeight)}</div>
      <div class="muted">국어 ${formatNumber(row.koreanRatio)} / 수학 ${formatNumber(row.mathRatio)} / 탐구 ${formatNumber(row.inquiryRatio)}</div>
    </div>

    <div class="detail-card">
      <strong>영어 등급별 환점</strong>
      <div class="pill-list">${englishRows}</div>
    </div>
  `;

  elements.detailBody.append(wrapper);
}

function buildComparisonRows(track) {
  const analyzer25 = state.analyses.get("25");
  const analyzer26 = state.analyses.get("26");

  if (!analyzer25 || !analyzer26) {
    return [];
  }

  const older = new Map(analyzer25.results[track].map((row) => [row.programKey, row]));
  const compared = [];

  for (const newerRow of analyzer26.results[track]) {
    const olderRow = older.get(newerRow.programKey);
    if (!olderRow) {
      continue;
    }

    compared.push({
      programKey: newerRow.programKey,
      label: `${newerRow.university} ${newerRow.major}`,
      selectionType: newerRow.selectionType,
      expectedDelta: calcDelta(newerRow.expectedScore, olderRow.expectedScore),
      safeDelta: calcDelta(newerRow.safeScore, olderRow.safeScore),
      reachDelta: calcDelta(newerRow.reachScore, olderRow.reachScore)
    });
  }

  return compared;
}

function renderCompare() {
  const comparisons = buildComparisonRows(state.track);

  if (!comparisons.length) {
    elements.comparePanel.hidden = true;
    return;
  }

  elements.comparePanel.hidden = false;
  elements.compareGrid.replaceChildren();
  elements.compareList.replaceChildren();
  elements.compareBadge.textContent = `${state.track} 기준 ${formatNumber(comparisons.length)}개 모집단위 비교`;

  const harder = comparisons.filter((item) => (item.expectedDelta ?? 0) > 0);
  const easier = comparisons.filter((item) => (item.expectedDelta ?? 0) < 0);
  const largest = [...comparisons].sort(
    (left, right) => Math.abs(right.expectedDelta ?? 0) - Math.abs(left.expectedDelta ?? 0)
  )[0];

  const cards = [
    { label: "비교 가능", value: comparisons.length, note: "25와 26 모두 존재" },
    { label: "26이 더 높음", value: harder.length, note: "예상점수 상승" },
    { label: "26이 더 낮음", value: easier.length, note: "예상점수 하락" },
    { label: "최대 변동", value: largest ? formatDelta(largest.expectedDelta).text : "-", note: largest ? largest.label : "변동 없음" }
  ];

  for (const item of cards) {
    const card = document.createElement("article");
    card.className = "compare-card";
    card.innerHTML = `
      <strong>${escapeHtml(item.label)}</strong>
      <div class="value">${escapeHtml(String(item.value))}</div>
      <small>${escapeHtml(item.note)}</small>
    `;
    elements.compareGrid.append(card);
  }

  const topMovers = [...comparisons]
    .sort((left, right) => Math.abs(right.expectedDelta ?? 0) - Math.abs(left.expectedDelta ?? 0))
    .slice(0, 8);

  for (const item of topMovers) {
    const delta = formatDelta(item.expectedDelta);
    const wrapper = document.createElement("div");
    wrapper.className = "compare-item";
    wrapper.innerHTML = `
      <strong>${escapeHtml(item.label)}</strong>
      <div class="muted">${escapeHtml(item.selectionType || "-")}</div>
      <div>예상점수 변화 <span class="delta ${delta.tone}">${delta.text}</span></div>
    `;
    elements.compareList.append(wrapper);
  }
}

function renderAll() {
  renderFileList();
  renderTrackTabs();
  renderSummary();
  renderFilters();
  renderCompare();
  renderResults();

  if (!state.analyses.size) {
    setUploadStatus("파일을 불러오지 않았습니다.");
  } else {
    const active = getActiveAnalysis();
    setUploadStatus(`${getSortedAnalyses().length}개 파일 로드됨 · 현재 ${active.yearLabel}`);
  }
}

async function readWorkbook(file) {
  const buffer = await file.arrayBuffer();
  return XLSX.read(buffer, {
    type: "array",
    cellFormula: true,
    dense: false
  });
}

async function handleFiles(fileList) {
  const files = [...fileList];
  if (!files.length) {
    return;
  }

  setUploadStatus("분석기 파일을 읽는 중입니다...");

  for (const file of files) {
    try {
      const workbook = await readWorkbook(file);
      const analysis = extractAnalysis(file, workbook);
      state.analyses.set(analysis.yearKey, analysis);
      state.activeYearKey = analysis.yearKey;
    } catch (error) {
      setUploadStatus(`파일 읽기 실패: ${file.name}`);
      console.error(error);
    }
  }

  state.selectedProgramKey = null;
  renderAll();
}

function clearAll() {
  state.analyses.clear();
  state.activeYearKey = null;
  state.track = TRACKS[0];
  state.status = "all";
  state.group = "all";
  state.region = "all";
  state.sort = "status";
  state.search = "";
  state.filteredRows = [];
  state.selectedProgramKey = null;
  elements.fileInput.value = "";
  renderAll();
}

elements.fileInput.addEventListener("change", (event) => {
  void handleFiles(event.target.files);
});

elements.clearButton.addEventListener("click", clearAll);

elements.yearFilter.addEventListener("change", (event) => {
  state.activeYearKey = event.target.value === "all" ? getSortedAnalyses()[0]?.yearKey || null : event.target.value;
  state.selectedProgramKey = null;
  renderAll();
});

elements.statusFilter.addEventListener("change", (event) => {
  state.status = event.target.value;
  state.selectedProgramKey = null;
  renderAll();
});

elements.groupFilter.addEventListener("change", (event) => {
  state.group = event.target.value;
  state.selectedProgramKey = null;
  renderAll();
});

elements.regionFilter.addEventListener("change", (event) => {
  state.region = event.target.value;
  state.selectedProgramKey = null;
  renderAll();
});

elements.sortFilter.addEventListener("change", (event) => {
  state.sort = event.target.value;
  state.selectedProgramKey = null;
  renderAll();
});

elements.searchInput.addEventListener("input", (event) => {
  state.search = event.target.value;
  state.selectedProgramKey = null;
  renderAll();
});

elements.resultBody.addEventListener("click", (event) => {
  const row = event.target.closest("tr[data-key]");
  if (!row) {
    return;
  }

  state.selectedProgramKey = row.dataset.key;
  renderResults();
});

["dragenter", "dragover"].forEach((eventName) => {
  elements.dropzone.addEventListener(eventName, (event) => {
    event.preventDefault();
    elements.dropzone.classList.add("is-dragging");
  });
});

["dragleave", "drop"].forEach((eventName) => {
  elements.dropzone.addEventListener(eventName, (event) => {
    event.preventDefault();
    elements.dropzone.classList.remove("is-dragging");
  });
});

elements.dropzone.addEventListener("drop", (event) => {
  const files = event.dataTransfer?.files;
  if (files?.length) {
    void handleFiles(files);
  }
});

renderAll();
