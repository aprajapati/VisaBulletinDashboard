const palette = {
  final: "#2a9d8f",
  filing: "#e76f51",
  muted: "rgba(38, 70, 83, 0.25)",
  grid: "rgba(17, 24, 39, 0.08)",
  bar: "rgba(38, 70, 83, 0.2)",
  highlight: "rgba(231, 111, 81, 0.8)"
};

const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

const monthLookup = {
  january: 1,
  february: 2,
  march: 3,
  april: 4,
  may: 5,
  june: 6,
  july: 7,
  august: 8,
  september: 9,
  october: 10,
  november: 11,
  december: 12
};

const SYSTEM_LABELS = {
  FAMILY: "Family",
  EMPLOYMENT: "Employment"
};

const FAMILY_ORDER = ["F1", "F2A", "F2B", "F3", "F4"];
const EMPLOYMENT_ORDER = [
  "EB1",
  "EB2",
  "EB3",
  "EB3_OTHER",
  "EB4",
  "EB4_RELIGIOUS",
  "EB5",
  "EB5_TARGETED"
];
const COLUMN_ORDER = [
  "ALL_CHARGEABILITY",
  "CHINA_MAINLAND",
  "CHINA_HONG_KONG",
  "INDIA",
  "MEXICO",
  "PHILIPPINES"
];

const state = {
  system: "FAMILY",
  preferenceKey: null,
  regionKey: null,
  window: 12,
  metric: "final"
};

const dataState = {
  bulletins: [],
  systems: {
    FAMILY: { rows: new Map(), columns: new Map() },
    EMPLOYMENT: { rows: new Map(), columns: new Map() }
  },
  metricAvailability: {
    FAMILY: { final: false, filing: false },
    EMPLOYMENT: { final: false, filing: false }
  }
};

const regionSelect = document.getElementById("regionSelect");
const categorySelect = document.getElementById("categorySelect");
const preferenceSelect = document.getElementById("preferenceSelect");
const windowRange = document.getElementById("windowRange");
const windowValue = document.getElementById("windowValue");
const metricButtons = document.querySelectorAll(".segmented-btn");
const statusValue = document.getElementById("statusValue");
const statusDate = document.getElementById("statusDate");
const statusPill = document.getElementById("statusPill");
const movementChip = document.getElementById("movementChip");
const backlogChip = document.getElementById("backlogChip");
const snapshotChip = document.getElementById("snapshotChip");

const kpiCurrentLabel = document.getElementById("kpiCurrentLabel");
const kpiCurrent = document.getElementById("kpiCurrent");
const kpiCurrentNote = document.getElementById("kpiCurrentNote");
const kpiQuarter = document.getElementById("kpiQuarter");
const kpiYtd = document.getElementById("kpiYtd");
const kpiGap = document.getElementById("kpiGap");
const snapshotTable = document.getElementById("snapshotTable");


function round1(value) {
  return Math.round(value * 10) / 10;
}

function normalizeText(value) {
  if (value === null || value === undefined) return "";
  return String(value).replace(/\s+/g, " ").trim();
}

function cleanKey(value) {
  return normalizeText(value).toUpperCase().replace(/[^A-Z0-9]+/g, "");
}

function parsePublicationFromUrl(url) {
  if (!url) return null;
  const match = url.match(/visa-bulletin-for-([a-z]+)-(\d{4})/i);
  if (!match) return null;
  const monthName = match[1].toLowerCase();
  const month = monthLookup[monthName];
  if (!month) return null;
  const year = Number(match[2]);
  const date = new Date(Date.UTC(year, month - 1, 1));
  const key = `${year}-${String(month).padStart(2, "0")}`;
  const label = `${monthNames[month - 1]} ${year}`;
  return { date, key, label, year, month };
}

function parsePublication(bulletin) {
  const publication = bulletin.publication || {};
  if (Number.isInteger(publication.month) && Number.isInteger(publication.year)) {
    const month = publication.month;
    const year = publication.year;
    const date = new Date(Date.UTC(year, month - 1, 1));
    const key = `${year}-${String(month).padStart(2, "0")}`;
    const label = `${monthNames[month - 1]} ${year}`;
    return { date, key, label, year, month };
  }
  return parsePublicationFromUrl(bulletin.sources && bulletin.sources.htmlUrl);
}

function detectSystem(rows) {
  const labels = (rows || []).map((row) => cleanKey(row.label));
  if (labels.some((label) => label.includes("FAMILY"))) return "FAMILY";
  if (labels.some((label) => label.includes("EMPLOYMENT"))) return "EMPLOYMENT";
  if (labels.some((label) => label.includes("AFRICA")) && labels.some((label) => label.includes("EUROPE"))) {
    return "DIVERSITY";
  }
  if (labels.some((label) => label.startsWith("F1") || label.startsWith("F2"))) return "FAMILY";
  if (labels.some((label) => label.includes("OTHERWORKERS") || label.includes("RELIGIOUS"))) return "EMPLOYMENT";
  return "OTHER";
}

function extractColumns(chart) {
  const columns = (chart.columns || []).map((col, index) => ({
    id: col.colId !== undefined && col.colId !== null ? String(col.colId) : String(index + 1),
    label: normalizeText(col.label || "")
  }));

  const usesFallbackLabels =
    columns.length > 0 &&
    columns.every((col) => {
      const label = col.label.toLowerCase();
      return label === "" || label.startsWith("unnamed") || /^\d+$/.test(col.label);
    });

  if (!usesFallbackLabels || !chart.rows || chart.rows.length === 0) {
    return { columns, headerRowId: null };
  }

  const headerRowId = chart.rows[0].rowId !== undefined ? String(chart.rows[0].rowId) : null;
  const headerLabels = {};
  (chart.cells || []).forEach((cell) => {
    if (String(cell.rowId) !== headerRowId) return;
    const raw = normalizeText(cell.rawText || (cell.value && cell.value.status) || "");
    if (raw) headerLabels[String(cell.colId)] = raw;
  });

  const derived = columns.map((col) => ({
    id: col.id,
    label: headerLabels[col.id] || col.label
  }));

  return { columns: derived, headerRowId };
}

function normalizeColumnKey(label) {
  const cleaned = normalizeText(label).toUpperCase();
  if (!cleaned) return null;
  if (cleaned.includes("WORLDWIDE") || cleaned.includes("ALL CHARGEABILITY")) {
    return "ALL_CHARGEABILITY";
  }
  if (cleaned.includes("CHINA") && cleaned.includes("MAINLAND")) return "CHINA_MAINLAND";
  if (cleaned.includes("CHINA") && cleaned.includes("BORN")) return "CHINA_MAINLAND";
  if (cleaned.includes("HONG KONG")) return "CHINA_HONG_KONG";
  if (cleaned.includes("INDIA")) return "INDIA";
  if (cleaned.includes("MEXICO")) return "MEXICO";
  if (cleaned.includes("PHILIPPINES")) return "PHILIPPINES";
  return cleaned.replace(/[^A-Z0-9]+/g, "_").replace(/^_|_$/g, "");
}

function mapFamilyRow(label) {
  const cleaned = cleanKey(label);
  if (!cleaned) return null;
  if (cleaned === "FAMILY" || cleaned.includes("FAMILYSPONSORED")) return null;
  if (cleaned.startsWith("F1") || cleaned === "1ST") return { key: "F1", label: "F1" };
  if (cleaned.startsWith("F2A") || cleaned === "2A") return { key: "F2A", label: "F2A" };
  if (cleaned.startsWith("F2B") || cleaned === "2B") return { key: "F2B", label: "F2B" };
  if (cleaned.startsWith("F3") || cleaned === "3RD") return { key: "F3", label: "F3" };
  if (cleaned.startsWith("F4") || cleaned === "4TH") return { key: "F4", label: "F4" };
  return { key: cleaned, label: normalizeText(label) };
}

function mapEmploymentRow(label) {
  const cleaned = cleanKey(label);
  if (!cleaned) return null;
  if (cleaned.includes("EMPLOYMENTBASED")) return null;
  if (cleaned.startsWith("EB1") || cleaned === "1ST") return { key: "EB1", label: "EB-1" };
  if (cleaned.startsWith("EB2") || cleaned === "2ND") return { key: "EB2", label: "EB-2" };
  if (cleaned.startsWith("EB3") || cleaned === "3RD") return { key: "EB3", label: "EB-3" };
  if (cleaned.includes("OTHERWORKERS")) return { key: "EB3_OTHER", label: "EB-3 Other Workers" };
  if (cleaned.startsWith("EB4") || cleaned === "4TH") return { key: "EB4", label: "EB-4" };
  if (cleaned.includes("RELIGIOUS")) return { key: "EB4_RELIGIOUS", label: "EB-4 Religious" };
  if (cleaned.startsWith("EB5") || cleaned === "5TH") return { key: "EB5", label: "EB-5" };
  if (cleaned.includes("TARGETED") || cleaned.includes("REGIONALCENTERS")) {
    return { key: "EB5_TARGETED", label: "EB-5 Targeted/Regional" };
  }
  return { key: cleaned, label: normalizeText(label) };
}

function buildChartModel(chart, system) {
  const columnInfo = extractColumns(chart);
  const rows = [];
  const columns = [];
  const rowIdToKey = new Map();
  const colIdToKey = new Map();

  (chart.rows || []).forEach((row) => {
    const rowId = String(row.rowId);
    if (columnInfo.headerRowId && rowId === columnInfo.headerRowId) return;
    const mapped = system === "FAMILY" ? mapFamilyRow(row.label) : mapEmploymentRow(row.label);
    if (!mapped) return;
    rowIdToKey.set(rowId, mapped.key);
    if (!rows.some((entry) => entry.key === mapped.key)) {
      rows.push(mapped);
    }
  });

  columnInfo.columns.forEach((col) => {
    const key = normalizeColumnKey(col.label);
    if (!key) return;
    const colId = String(col.id);
    colIdToKey.set(colId, key);
    if (!columns.some((entry) => entry.key === key)) {
      columns.push({ key, label: normalizeText(col.label) || key });
    }
  });

  const cellMap = new Map();
  (chart.cells || []).forEach((cell) => {
    const rowId = String(cell.rowId);
    if (columnInfo.headerRowId && rowId === columnInfo.headerRowId) return;
    const rowKey = rowIdToKey.get(rowId);
    const colKey = colIdToKey.get(String(cell.colId));
    if (!rowKey || !colKey) return;
    cellMap.set(`${rowKey}|${colKey}`, cell);
  });

  return { rows, columns, cellMap };
}

function parseIsoDate(value) {
  if (!value) return null;
  const match = value.match(/(\d{4})-(\d{2})-(\d{2})/);
  if (!match) return null;
  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  return new Date(Date.UTC(year, month - 1, day));
}

function monthsBetween(bulletinDate, cutoffDate) {
  const yearDiff = bulletinDate.getUTCFullYear() - cutoffDate.getUTCFullYear();
  const monthDiff = bulletinDate.getUTCMonth() - cutoffDate.getUTCMonth();
  const dayDiff = bulletinDate.getUTCDate() - cutoffDate.getUTCDate();
  return Math.max(0, round1(yearDiff * 12 + monthDiff + dayDiff / 30));
}

function cellToDate(cell, bulletinDate) {
  if (!cell || !cell.value) return null;
  if (cell.value.kind === "DATE" && cell.value.date) {
    return parseIsoDate(cell.value.date);
  }
  if (cell.value.kind === "STATUS" && cell.value.status === "C") {
    return bulletinDate;
  }
  return null;
}

function cellToValue(cell, bulletinDate) {
  if (!cell || !cell.value) {
    return { months: null, status: "MISSING", raw: null, date: null };
  }
  if (cell.value.kind === "DATE" && cell.value.date) {
    const cutoffDate = parseIsoDate(cell.value.date);
    if (!cutoffDate) {
      return { months: null, status: "DATE", raw: cell.rawText || null, date: null };
    }
    return {
      months: monthsBetween(bulletinDate, cutoffDate),
      status: "DATE",
      raw: cell.rawText || null,
      date: cutoffDate
    };
  }
  if (cell.value.kind === "STATUS") {
    const status = cell.value.status || "UNKNOWN";
    if (status === "C") {
      return { months: 0, status: "C", raw: cell.rawText || "C", date: null };
    }
    if (status === "U") {
      return { months: null, status: "U", raw: cell.rawText || "U", date: null };
    }
    if (status === "NA") {
      return { months: null, status: "NA", raw: cell.rawText || "NA", date: null };
    }
    return { months: null, status: status, raw: cell.rawText || status, date: null };
  }
  return { months: null, status: "UNKNOWN", raw: cell.rawText || null, date: null };
}

function averageChartDate(chart, bulletinDate) {
  const timestamps = [];
  chart.cellMap.forEach((cell) => {
    const date = cellToDate(cell, bulletinDate);
    if (date) timestamps.push(date.getTime());
  });
  if (!timestamps.length) return null;
  const sum = timestamps.reduce((acc, value) => acc + value, 0);
  return sum / timestamps.length;
}

function assignMetrics(charts, bulletinDate) {
  if (charts.length === 1) {
    return { final: charts[0], filing: null };
  }
  const scored = charts.map((chart) => ({
    chart,
    score: averageChartDate(chart, bulletinDate)
  }));
  scored.sort((a, b) => {
    const left = a.score === null ? Number.POSITIVE_INFINITY : a.score;
    const right = b.score === null ? Number.POSITIVE_INFINITY : b.score;
    return left - right;
  });
  return {
    final: scored[0] ? scored[0].chart : null,
    filing: scored[1] ? scored[1].chart : null
  };
}

function scoreBulletin(bulletin) {
  let score = 0;
  ["FAMILY", "EMPLOYMENT"].forEach((system) => {
    const sys = bulletin.systems[system];
    if (!sys) return;
    ["final", "filing"].forEach((metric) => {
      const chart = sys[metric];
      if (!chart) return;
      score += chart.rows.length * chart.columns.length;
    });
  });
  return score;
}

function prepareData(raw) {
  const byKey = new Map();
  const systemsMeta = {
    FAMILY: { rows: new Map(), columns: new Map() },
    EMPLOYMENT: { rows: new Map(), columns: new Map() }
  };
  const metricAvailability = {
    FAMILY: { final: false, filing: false },
    EMPLOYMENT: { final: false, filing: false }
  };

  (raw.bulletins || []).forEach((bulletin) => {
    const publication = parsePublication(bulletin);
    if (!publication) return;

    const familyCharts = [];
    const employmentCharts = [];

    (bulletin.charts || []).forEach((chart) => {
      const system = detectSystem(chart.rows || []);
      if (system !== "FAMILY" && system !== "EMPLOYMENT") return;
      const model = buildChartModel(chart, system);
      if (!model.rows.length || !model.columns.length) return;
      if (system === "FAMILY") familyCharts.push(model);
      if (system === "EMPLOYMENT") employmentCharts.push(model);
    });

    if (!familyCharts.length && !employmentCharts.length) return;

    const systems = {};
    if (familyCharts.length) systems.FAMILY = assignMetrics(familyCharts, publication.date);
    if (employmentCharts.length) systems.EMPLOYMENT = assignMetrics(employmentCharts, publication.date);

    const item = {
      key: publication.key,
      date: publication.date,
      label: publication.label,
      sourceUrl: bulletin.sources ? bulletin.sources.htmlUrl : "",
      systems
    };

    const existing = byKey.get(publication.key);
    if (!existing || scoreBulletin(item) > scoreBulletin(existing)) {
      byKey.set(publication.key, item);
    }
  });

  const bulletins = Array.from(byKey.values()).sort((a, b) => a.date - b.date);

  bulletins.forEach((bulletin) => {
    ["FAMILY", "EMPLOYMENT"].forEach((system) => {
      const sys = bulletin.systems[system];
      if (!sys) return;
      ["final", "filing"].forEach((metric) => {
        const chart = sys[metric];
        if (!chart) return;
        metricAvailability[system][metric] = true;
        chart.rows.forEach((row) => systemsMeta[system].rows.set(row.key, row.label));
        chart.columns.forEach((col) => systemsMeta[system].columns.set(col.key, col.label));
      });
    });
  });

  return { bulletins, systemsMeta, metricAvailability };
}

function getRowOptions(system) {
  const meta = dataState.systems[system];
  if (!meta) return [];
  const rows = Array.from(meta.rows, ([key, label]) => ({ key, label }));
  const order = system === "FAMILY" ? FAMILY_ORDER : EMPLOYMENT_ORDER;
  rows.sort((a, b) => {
    const aIndex = order.indexOf(a.key);
    const bIndex = order.indexOf(b.key);
    if (aIndex === -1 && bIndex === -1) return a.label.localeCompare(b.label);
    if (aIndex === -1) return 1;
    if (bIndex === -1) return -1;
    return aIndex - bIndex;
  });
  return rows;
}

function getColumnOptions(system) {
  const meta = dataState.systems[system];
  if (!meta) return [];
  const columns = Array.from(meta.columns, ([key, label]) => ({ key, label }));
  columns.sort((a, b) => {
    const aIndex = COLUMN_ORDER.indexOf(a.key);
    const bIndex = COLUMN_ORDER.indexOf(b.key);
    if (aIndex === -1 && bIndex === -1) return a.label.localeCompare(b.label);
    if (aIndex === -1) return 1;
    if (bIndex === -1) return -1;
    return aIndex - bIndex;
  });
  return columns;
}

function getRowLabel(system, key) {
  const meta = dataState.systems[system];
  if (!meta || !key) return "";
  return meta.rows.get(key) || key;
}

function getColumnLabel(system, key) {
  const meta = dataState.systems[system];
  if (!meta || !key) return "";
  return meta.columns.get(key) || key;
}

function formatIsoDate(date) {
  return date ? date.toISOString().slice(0, 10) : "";
}

function formatShortDisplay(value) {
  if (!value) return "n/a";
  if (value.status === "DATE") return value.raw || formatIsoDate(value.date);
  if (value.status === "C") return "C";
  if (value.status === "U") return "U";
  if (value.status === "NA") return "NA";
  if (value.status === "MISSING") return "n/a";
  return value.raw || "n/a";
}

function formatLongDisplay(value) {
  if (!value) return "n/a";
  if (value.status === "DATE") return value.raw || formatIsoDate(value.date);
  if (value.status === "C") return "Current";
  if (value.status === "U") return "Unavailable";
  if (value.status === "NA") return "NA";
  if (value.status === "MISSING") return "n/a";
  return value.raw || "n/a";
}

function formatValueWithMonths(value) {
  if (!value) return "n/a";
  if (value.months === null || value.months === undefined) return formatLongDisplay(value);
  const monthsText = `${value.months.toFixed(1)} mo`;
  return `${formatLongDisplay(value)} (${monthsText})`;
}

function formatTooltipValue(value) {
  if (!value) return "n/a";
  if (value.months === null || value.months === undefined) return formatLongDisplay(value);
  const monthsText = `${value.months.toFixed(1)} mo`;
  const detail = formatLongDisplay(value);
  return detail ? `${monthsText} (${detail})` : monthsText;
}

function computeScaleBounds(values, minRange) {
  const nums = values.filter((value) => Number.isFinite(value));
  if (!nums.length) {
    return { min: 0, max: minRange };
  }
  let min = Math.min(...nums);
  let max = Math.max(...nums);
  if (max - min < minRange) {
    max = Math.max(max, minRange);
    min = Math.max(0, max - minRange);
    return { min: round1(min), max: round1(max) };
  }
  const pad = (max - min) * 0.1;
  min = Math.max(0, min - pad);
  max += pad;
  return { min: round1(min), max: round1(max) };
}

function buildSeries(system, metric, rowKey, colKey) {
  return dataState.bulletins.map((bulletin) => {
    const chart = bulletin.systems[system] ? bulletin.systems[system][metric] : null;
    if (!chart) {
      return { label: bulletin.label, value: null };
    }
    const cell = chart.cellMap.get(`${rowKey}|${colKey}`);
    return { label: bulletin.label, value: cellToValue(cell, bulletin.date) };
  });
}

function findLatestIndex(series) {
  for (let i = series.length - 1; i >= 0; i--) {
    const value = series[i].value;
    if (value && value.months !== null && value.months !== undefined) return i;
  }
  return -1;
}

function getDelta(series, monthsBack) {
  const latestIndex = findLatestIndex(series);
  if (latestIndex < 0) return null;
  const previousIndex = latestIndex - monthsBack;
  if (previousIndex < 0) return null;
  const latest = series[latestIndex].value ? series[latestIndex].value.months : null;
  const previous = series[previousIndex].value ? series[previousIndex].value.months : null;
  if (latest === null || latest === undefined || previous === null || previous === undefined) return null;
  return round1(latest - previous);
}

function getYtdDelta(series) {
  const latestIndex = findLatestIndex(series);
  if (latestIndex < 0) return null;
  const latestYear = dataState.bulletins[latestIndex].date.getUTCFullYear();
  let startIndex = -1;
  for (let i = 0; i <= latestIndex; i++) {
    const year = dataState.bulletins[i].date.getUTCFullYear();
    const value = series[i].value;
    if (year === latestYear && value && value.months !== null && value.months !== undefined) {
      startIndex = i;
      break;
    }
  }
  if (startIndex < 0) return null;
  const latest = series[latestIndex].value.months;
  const start = series[startIndex].value.months;
  return round1(latest - start);
}

function formatDelta(value) {
  if (value === null || value === undefined) return "n/a";
  const sign = value > 0 ? "+" : value < 0 ? "-" : "";
  return `${sign}${Math.abs(value).toFixed(1)} mo`;
}

function getLatestBulletin(system) {
  for (let i = dataState.bulletins.length - 1; i >= 0; i--) {
    if (dataState.bulletins[i].systems[system]) return dataState.bulletins[i];
  }
  return null;
}

function getLatestChart(system, metric) {
  for (let i = dataState.bulletins.length - 1; i >= 0; i--) {
    const chart = dataState.bulletins[i].systems[system]
      ? dataState.bulletins[i].systems[system][metric]
      : null;
    if (chart) return { chart, bulletin: dataState.bulletins[i] };
  }
  return null;
}

function getLatestValue(system, metric, rowKey, colKey) {
  for (let i = dataState.bulletins.length - 1; i >= 0; i--) {
    const chart = dataState.bulletins[i].systems[system]
      ? dataState.bulletins[i].systems[system][metric]
      : null;
    if (!chart) continue;
    const cell = chart.cellMap.get(`${rowKey}|${colKey}`);
    if (!cell) continue;
    return { value: cellToValue(cell, dataState.bulletins[i].date), bulletin: dataState.bulletins[i] };
  }
  return null;
}

function updateStatus() {
  const regionLabel = getColumnLabel(state.system, state.regionKey) || "Region";
  const rowLabel = getRowLabel(state.system, state.preferenceKey) || "Preference";
  const systemLabel = SYSTEM_LABELS[state.system] || state.system;
  const latest = getLatestBulletin(state.system);

  statusValue.textContent = `${regionLabel} - ${rowLabel} (${systemLabel})`;
  statusDate.textContent = latest ? `Updated: ${latest.label}` : "No data available";
  statusPill.textContent = state.metric === "final" ? "Final Action" : "Dates for Filing";
}

function updateKpis() {
  if (!state.preferenceKey || !state.regionKey) {
    kpiCurrent.textContent = "n/a";
    kpiQuarter.textContent = "n/a";
    kpiYtd.textContent = "n/a";
    kpiGap.textContent = "n/a";
    return;
  }

  const series = buildSeries(state.system, state.metric, state.preferenceKey, state.regionKey);
  const latestIndex = findLatestIndex(series);
  const latestValue = latestIndex >= 0 ? series[latestIndex].value : null;

  kpiCurrentLabel.textContent = state.metric === "final" ? "Current Final Action" : "Current Filing Date";
  kpiCurrent.textContent = latestValue ? formatValueWithMonths(latestValue) : "n/a";

  const latestBulletin = getLatestBulletin(state.system);
  kpiCurrentNote.textContent = latestBulletin ? `Latest bulletin: ${latestBulletin.label}` : "No bulletin data";

  kpiQuarter.textContent = formatDelta(getDelta(series, 3));
  kpiYtd.textContent = formatDelta(getYtdDelta(series));

  const finalValue = getLatestValue(state.system, "final", state.preferenceKey, state.regionKey);
  const filingValue = getLatestValue(state.system, "filing", state.preferenceKey, state.regionKey);
  if (finalValue && filingValue && finalValue.value.months !== null && filingValue.value.months !== null) {
    kpiGap.textContent = formatDelta(round1(filingValue.value.months - finalValue.value.months));
  } else {
    kpiGap.textContent = "n/a";
  }
}

function toPercentages(counts) {
  const total = counts.reduce((sum, value) => sum + value, 0);
  if (!total) return counts.map(() => 0);
  return counts.map((value) => round1((value / total) * 100));
}

function updateSnapshot() {
  snapshotTable.innerHTML = "";
  const rows = getRowOptions(state.system);

  const latestFinal = getLatestChart(state.system, "final");
  const latestFiling = getLatestChart(state.system, "filing");

  rows.forEach((row) => {
    const series = buildSeries(state.system, state.metric, row.key, state.regionKey);
    const delta = getDelta(series, 3);
    const trendClass = delta > 0 ? "trend-up" : delta < 0 ? "trend-down" : "trend-flat";

    let finalValue = null;
    let filingValue = null;

    if (latestFinal) {
      const cell = latestFinal.chart.cellMap.get(`${row.key}|${state.regionKey}`);
      finalValue = cellToValue(cell, latestFinal.bulletin.date);
    }
    if (latestFiling) {
      const cell = latestFiling.chart.cellMap.get(`${row.key}|${state.regionKey}`);
      filingValue = cellToValue(cell, latestFiling.bulletin.date);
    }

    const rowElement = document.createElement("tr");
    if (row.key === state.preferenceKey) {
      rowElement.classList.add("active-row");
    }

    rowElement.innerHTML = `
      <td>${row.label}</td>
      <td>${formatShortDisplay(finalValue)}</td>
      <td>${formatShortDisplay(filingValue)}</td>
      <td class="${trendClass}">${formatDelta(delta)}</td>
    `;
    snapshotTable.appendChild(rowElement);
  });

  snapshotChip.textContent = `${getColumnLabel(state.system, state.regionKey)} snapshot`;
}

function updateAll() {
  if (!dataState.bulletins.length) return;
  updateStatus();
  updateKpis();
  updateMovementChart();
  updateCutoffChart();
  updateBacklogChart();
  updateSnapshot();
}

function populateSystemOptions() {
  const systems = Object.keys(dataState.systems).filter(
    (system) => dataState.systems[system].rows.size > 0
  );
  categorySelect.innerHTML = "";
  systems.forEach((system) => {
    const option = document.createElement("option");
    option.value = system;
    option.textContent = SYSTEM_LABELS[system] || system;
    categorySelect.appendChild(option);
  });
  if (!systems.includes(state.system)) {
    state.system = systems[0];
  }
  categorySelect.value = state.system || "";
}

function populatePreferenceOptions() {
  const rows = getRowOptions(state.system);
  preferenceSelect.innerHTML = "";
  rows.forEach((row) => {
    const option = document.createElement("option");
    option.value = row.key;
    option.textContent = row.label;
    preferenceSelect.appendChild(option);
  });
  if (!rows.find((row) => row.key === state.preferenceKey)) {
    state.preferenceKey = rows.length ? rows[0].key : null;
  }
  preferenceSelect.value = state.preferenceKey || "";
}

function populateRegionOptions() {
  const columns = getColumnOptions(state.system);
  regionSelect.innerHTML = "";
  columns.forEach((column) => {
    const option = document.createElement("option");
    option.value = column.key;
    option.textContent = column.label;
    regionSelect.appendChild(option);
  });
  if (!columns.find((column) => column.key === state.regionKey)) {
    state.regionKey = columns.length ? columns[0].key : null;
  }
  regionSelect.value = state.regionKey || "";
}

function updateMetricButtons() {
  const availability = dataState.metricAvailability[state.system] || { final: true, filing: false };
  metricButtons.forEach((button) => {
    const metric = button.dataset.metric;
    const enabled = Boolean(availability[metric]);
    button.disabled = !enabled;
    if (!enabled && state.metric === metric) {
      state.metric = "final";
    }
  });

  metricButtons.forEach((button) => {
    const metric = button.dataset.metric;
    const isActive = metric === state.metric;
    button.classList.toggle("is-active", isActive);
    button.setAttribute("aria-pressed", isActive ? "true" : "false");
  });
}

function updateWindowRange() {
  const total = dataState.bulletins.length;
  const minWindow = Math.min(6, total);
  const maxWindow = Math.max(minWindow, total);
  windowRange.min = String(minWindow);
  windowRange.max = String(maxWindow);
  if (state.window < minWindow) state.window = minWindow;
  if (state.window > maxWindow) state.window = maxWindow;
  windowRange.value = String(state.window);
  windowValue.textContent = `${state.window} months`;
}

function enableControls() {
  regionSelect.disabled = false;
  categorySelect.disabled = false;
  preferenceSelect.disabled = false;
}

function setLoadingState() {
  statusValue.textContent = "Loading data...";
  statusDate.textContent = "Parsing visa_bulletins.all.json";
  regionSelect.disabled = true;
  categorySelect.disabled = true;
  preferenceSelect.disabled = true;
}

function showError(error) {
  console.error(error);
  statusValue.textContent = "Unable to load visa bulletin data";
  statusDate.textContent = "Run a local server so fetch can access visa_bulletins.all.json";
}

regionSelect.addEventListener("change", (event) => {
  state.regionKey = event.target.value;
  updateAll();
});

categorySelect.addEventListener("change", (event) => {
  state.system = event.target.value;
  updateMetricButtons();
  populatePreferenceOptions();
  populateRegionOptions();
  updateAll();
});

preferenceSelect.addEventListener("change", (event) => {
  state.preferenceKey = event.target.value;
  updateAll();
});

windowRange.addEventListener("input", (event) => {
  state.window = Number(event.target.value);
  windowValue.textContent = `${state.window} months`;
  updateAll();
});

metricButtons.forEach((button) => {
  button.addEventListener("click", () => {
    if (button.disabled) return;
    state.metric = button.dataset.metric;
    updateMetricButtons();
    updateAll();
  });
});

const movementChart = echarts.init(document.getElementById("movementChart"));
const cutoffChart = echarts.init(document.getElementById("cutoffChart"));
const backlogChart = echarts.init(document.getElementById("backlogChart"));

window.addEventListener("resize", () => {
  movementChart.resize();
  cutoffChart.resize();
  backlogChart.resize();
});

function updateMovementChart() {
  if (!state.preferenceKey || !state.regionKey) return;
  const finalSeries = buildSeries(state.system, "final", state.preferenceKey, state.regionKey);
  const filingSeries = buildSeries(state.system, "filing", state.preferenceKey, state.regionKey);

  const total = dataState.bulletins.length;
  const windowSize = Math.min(state.window, total);
  const start = Math.max(0, total - windowSize);

  const labels = dataState.bulletins.slice(start).map((bulletin) => bulletin.label);
  const finalData = finalSeries.slice(start).map((item) => ({
    value: item.value ? item.value.months : null,
    detail: item.value ? formatTooltipValue(item.value) : "n/a"
  }));
  const filingData = filingSeries.slice(start).map((item) => ({
    value: item.value ? item.value.months : null,
    detail: item.value ? formatTooltipValue(item.value) : "n/a"
  }));

  const finalActive = state.metric === "final";
  const finalColor = finalActive ? palette.final : palette.muted;
  const filingColor = finalActive ? palette.muted : palette.filing;

  const filingAvailable = dataState.metricAvailability[state.system]
    ? dataState.metricAvailability[state.system].filing
    : false;
  const scaleValues = filingAvailable
    ? finalData.concat(filingData).map((item) => item.value)
    : finalData.map((item) => item.value);
  const bounds = computeScaleBounds(scaleValues, 24);

  const series = [
    {
      name: "Final Action",
      type: "line",
      smooth: true,
      showSymbol: false,
      data: finalData,
      lineStyle: { color: finalColor, width: 3 },
      itemStyle: { color: finalColor },
      areaStyle: { color: finalActive ? "rgba(42, 157, 143, 0.18)" : "rgba(38, 70, 83, 0.08)" }
    }
  ];

  if (filingAvailable) {
    series.push({
      name: "Dates for Filing",
      type: "line",
      smooth: true,
      showSymbol: false,
      data: filingData,
      lineStyle: { color: filingColor, width: 3 },
      itemStyle: { color: filingColor },
      areaStyle: { color: finalActive ? "rgba(38, 70, 83, 0.08)" : "rgba(231, 111, 81, 0.16)" }
    });
  }

  movementChart.setOption(
    {
      animationDuration: 500,
      textStyle: { fontFamily: "Space Grotesk, sans-serif", color: "#101312" },
      grid: { left: 48, right: 24, top: 24, bottom: 40 },
      legend: { data: series.map((item) => item.name), top: 0, right: 0, textStyle: { color: "#4c5a55" } },
      tooltip: {
        trigger: "axis",
        formatter: (params) => {
          const list = Array.isArray(params) ? params : [params];
          const title = list[0] ? list[0].axisValue : "";
          const lines = list.map((item) => {
            const detail = item.data && item.data.detail ? item.data.detail : "n/a";
            return `${item.marker}${item.seriesName}: ${detail}`;
          });
          return [title].concat(lines).join("<br/>");
        }
      },
      xAxis: {
        type: "category",
        data: labels,
        axisLabel: { color: "#4c5a55" },
        axisTick: { alignWithLabel: true }
      },
      yAxis: {
        type: "value",
        min: bounds.min,
        max: bounds.max,
        axisLabel: { color: "#4c5a55", formatter: "{value} mo" },
        splitLine: { lineStyle: { color: "rgba(17, 24, 39, 0.08)" } }
      },
      series
    },
    true
  );

  movementChip.textContent = `${getRowLabel(state.system, state.preferenceKey)} - ${windowSize} months`;
}

function updateCutoffChart() {
  const latest = getLatestChart(state.system, state.metric) || getLatestChart(state.system, "final");
  if (!latest || !state.regionKey) return;

  const rows = getRowOptions(state.system);
  const labels = [];
  const data = [];

  rows.forEach((row) => {
    const cell = latest.chart.cellMap.get(`${row.key}|${state.regionKey}`);
    const value = cellToValue(cell, latest.bulletin.date);
    labels.push(row.label);
    data.push({
      value: value.months,
      detail: formatTooltipValue(value),
      itemStyle: { color: row.key === state.preferenceKey ? palette.highlight : palette.bar }
    });
  });

  const bounds = computeScaleBounds(
    data.map((item) => item.value),
    24
  );
  const rotate = labels.length > 6 ? 30 : 0;

  cutoffChart.setOption(
    {
      animationDuration: 500,
      textStyle: { fontFamily: "Space Grotesk, sans-serif", color: "#101312" },
      grid: { left: 48, right: 16, top: 24, bottom: 60 },
      tooltip: {
        trigger: "axis",
        axisPointer: { type: "shadow" },
        formatter: (params) => {
          const item = Array.isArray(params) ? params[0] : params;
          const detail = item.data && item.data.detail ? item.data.detail : "n/a";
          return `${item.name}: ${detail}`;
        }
      },
      xAxis: {
        type: "category",
        data: labels,
        axisLabel: { color: "#4c5a55", rotate, interval: 0 }
      },
      yAxis: {
        type: "value",
        min: bounds.min,
        max: bounds.max,
        axisLabel: { color: "#4c5a55", formatter: "{value} mo" },
        splitLine: { lineStyle: { color: "rgba(17, 24, 39, 0.08)" } }
      },
      series: [
        {
          type: "bar",
          data,
          barWidth: "60%",
          itemStyle: { borderRadius: [8, 8, 0, 0] }
        }
      ]
    },
    true
  );
}

function updateBacklogChart() {
  const latest = getLatestChart(state.system, state.metric) || getLatestChart(state.system, "final");
  if (!latest || !state.regionKey) return;

  const rows = getRowOptions(state.system);
  const counts = {
    current: 0,
    short: 0,
    long: 0,
    unavailable: 0
  };

  rows.forEach((row) => {
    const cell = latest.chart.cellMap.get(`${row.key}|${state.regionKey}`);
    const value = cellToValue(cell, latest.bulletin.date);
    if (!value) {
      counts.unavailable += 1;
      return;
    }
    if (value.status === "C") {
      counts.current += 1;
    } else if (value.months === null || value.months === undefined) {
      counts.unavailable += 1;
    } else if (value.months <= 12) {
      counts.short += 1;
    } else {
      counts.long += 1;
    }
  });

  const labels = ["Current", "0-12 mo", "12+ mo"];
  const data = [counts.current, counts.short, counts.long];
  const colors = ["rgba(42, 157, 143, 0.8)", "rgba(231, 111, 81, 0.8)", "rgba(38, 70, 83, 0.8)"];

  if (counts.unavailable > 0) {
    labels.push("Unavailable");
    data.push(counts.unavailable);
    colors.push("rgba(17, 24, 39, 0.2)");
  }

  const percentages = toPercentages(data);
  const seriesData = labels.map((label, index) => ({
    name: label,
    value: percentages[index],
    itemStyle: { color: colors[index] }
  }));

  backlogChart.setOption(
    {
      animationDuration: 500,
      textStyle: { fontFamily: "Space Grotesk, sans-serif", color: "#101312" },
      tooltip: {
        trigger: "item",
        formatter: (params) => `${params.name}: ${params.value}%`
      },
      legend: {
        bottom: 0,
        left: "center",
        textStyle: { color: "#4c5a55" }
      },
      series: [
        {
          type: "pie",
          radius: ["45%", "70%"],
          avoidLabelOverlap: true,
          label: { show: false },
          emphasis: { label: { show: true, fontWeight: 600 } },
          data: seriesData
        }
      ]
    },
    true
  );

  backlogChip.textContent = `${SYSTEM_LABELS[state.system]} - ${state.metric === "final" ? "Final" : "Filing"}`;
}

async function loadData() {
  setLoadingState();
  try {
    const response = await fetch("visa_bulletins.all.json");
    if (!response.ok) {
      throw new Error(`Data load failed: ${response.status}`);
    }
    const raw = await response.json();
    const prepared = prepareData(raw);
    dataState.bulletins = prepared.bulletins;
    dataState.systems = prepared.systemsMeta;
    dataState.metricAvailability = prepared.metricAvailability;

    if (!dataState.bulletins.length) {
      throw new Error("No bulletin entries found.");
    }

    populateSystemOptions();
    populatePreferenceOptions();
    populateRegionOptions();
    updateMetricButtons();
    updateWindowRange();
    enableControls();
    updateAll();
  } catch (error) {
    showError(error);
  }
}

loadData();
