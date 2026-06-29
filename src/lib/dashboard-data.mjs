const METRIC_ALIASES = {
  sales: ["sales", "sale", "penjualan", "omzet", "revenue", "grosssales"],
  quantity: ["quantity", "qty", "kuantitas", "jumlah", "units", "unit"],
  discount: ["discount", "diskon", "potongan", "rebate"],
  profit: ["profit", "laba", "keuntungan", "netprofit", "income"],
};

const METRIC_LABELS = {
  sales: "Sales",
  quantity: "Quantity",
  discount: "Discount",
  profit: "Profit",
};

function normalizeText(value) {
  return String(value ?? "")
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "");
}

export function normalizeSheetRows(rows = []) {
  const cleanedRows = (rows ?? [])
    .filter(Array.isArray)
    .map((row) => (Array.isArray(row) ? row.map((cell) => cell ?? "") : [row ?? ""]))
    .filter((row) => row.some((cell) => String(cell ?? "").trim() !== ""));

  if (!cleanedRows.length) {
    return { header: [], rows: [] };
  }

  const headerRowIndex = cleanedRows.findIndex((row) => {
    const nonEmptyCells = row.filter((cell) => String(cell ?? "").trim() !== "");
    if (nonEmptyCells.length < 2) return false;

    const headerKeywords = ["date", "tanggal", "year", "tahun", "category", "kategori", "sales", "profit", "quantity", "qty", "discount", "diskon", "revenue", "omzet", "total", "value", "nilai", "product", "produk", "customer", "pelanggan"];
    const hasHeaderKeyword = nonEmptyCells.some((cell) => headerKeywords.some((keyword) => normalizeText(cell).includes(keyword)));
    const textCells = nonEmptyCells.filter((cell) => {
      const text = String(cell ?? "").trim();
      return text.length >= 2 && /[A-Za-z]/.test(text);
    });

    return hasHeaderKeyword || textCells.length >= Math.max(2, Math.ceil(nonEmptyCells.length / 2));
  });

  const resolvedHeaderIndex = headerRowIndex >= 0 ? headerRowIndex : 0;
  const header = cleanedRows[resolvedHeaderIndex].map((cell) => String(cell ?? "").trim());
  const dataRows = cleanedRows
    .slice(resolvedHeaderIndex + 1)
    .filter((row) => row.some((cell) => String(cell ?? "").trim() !== ""));

  return { header, rows: dataRows };
}

function toNumber(value) {
  if (typeof value === "number") return Number.isFinite(value) ? value : null;
  if (typeof value === "string") {
    const trimmed = value.trim();
    if (!trimmed) return null;

    const cleaned = trimmed.replace(/[^0-9,.-]/g, "").replace(/,/g, "");
    if (!cleaned || cleaned === "." || cleaned === "-" || cleaned === "+") return null;

    const parsed = Number(cleaned);
    return Number.isNaN(parsed) ? null : parsed;
  }
  return null;
}

function isIgnoredColumnName(columnName) {
  const normalized = normalizeText(columnName);
  return [
    "year",
    "tahun",
    "thn",
    "date",
    "tanggal",
    "tanggaltransaksi",
    "month",
    "bulan",
    "day",
    "hari",
    "id",
    "code",
    "kode",
    "phone",
    "postal",
    "zip",
    "email",
    "category",
    "kategori",
    "segment",
    "class",
    "department",
    "productcategory",
    "productcategory",
    "jenis",
  ].some((token) => normalized.includes(token));
}

export function findHeaderIndex(header, aliases) {
  const normalizedAliases = aliases.map((alias) => normalizeText(alias));
  return header.findIndex((column) => {
    const normalized = normalizeText(column);
    return normalizedAliases.some((alias) => normalized === alias || normalized.includes(alias));
  });
}

export function getMetricColumnIndex(header, metricKey) {
  if (!metricKey || metricKey === "all") return -1;
  const aliases = METRIC_ALIASES[metricKey] ?? [];
  return findHeaderIndex(header, aliases);
}

export function getMetricOptions(header) {
  const options = [
    {
      value: "all",
      label: "Semua Statistik",
      description: "Tampilkan ringkasan dari semua metrik utama",
    },
  ];

  Object.entries(METRIC_ALIASES).forEach(([metricKey, aliases]) => {
    const matched = header.some((column) => {
      const normalized = normalizeText(column);
      return aliases.some((alias) => normalized === normalizeText(alias) || normalized.includes(normalizeText(alias)));
    });

    if (matched) {
      options.push({
        value: metricKey,
        label: METRIC_LABELS[metricKey],
        description: `Gunakan ${METRIC_LABELS[metricKey]} sebagai metrik utama`,
      });
    }
  });

  return options;
}

export function getYearValue(row, yearIndex, dateIndex) {
  if (yearIndex >= 0) {
    const value = row[yearIndex];
    if (value !== undefined && value !== null && value !== "") {
      const parsed = Number(value);
      if (!Number.isNaN(parsed)) return String(Math.floor(parsed));
      return String(value).trim();
    }
  }

  if (dateIndex >= 0) {
    const value = row[dateIndex];
    if (typeof value === "number") {
      const date = new Date((value - 25569) * 86400 * 1000);
      const year = date.getFullYear();
      if (!Number.isNaN(year) && year > 1980 && year < 2100) return String(year);
    } else if (typeof value === "string") {
      const match = value.match(/\b(19|20)\d{2}\b/);
      if (match) return match[0];
    }
  }

  return null;
}

export function getCategoryValue(row, categoryIndex) {
  if (categoryIndex < 0) return "all";
  const value = row[categoryIndex];
  if (value === undefined || value === null || value === "") return "Tidak Ada Kategori";
  return String(value).trim() || "Tidak Ada Kategori";
}

function getMonthKey(row, dateIndex) {
  if (dateIndex < 0) return null;
  const rawValue = row[dateIndex];
  let date;

  if (typeof rawValue === "number") {
    date = new Date((rawValue - 25569) * 86400 * 1000);
  } else if (typeof rawValue === "string") {
    const trimmed = rawValue.trim();
    const isoMatch = trimmed.match(/(\d{4})[/-](\d{1,2})[/-](\d{1,2})/);
    if (isoMatch) {
      date = new Date(Number(isoMatch[1]), Number(isoMatch[2]) - 1, Number(isoMatch[3]));
    } else {
      const compactMatch = trimmed.match(/(\d{4})(\d{2})(\d{2})/);
      if (compactMatch) {
        date = new Date(Number(compactMatch[1]), Number(compactMatch[2]) - 1, Number(compactMatch[3]));
      } else {
        const slashMatch = trimmed.match(/(\d{1,2})[/-](\d{1,2})[/-](\d{4})/);
        if (slashMatch) {
          date = new Date(Number(slashMatch[3]), Number(slashMatch[1]) - 1, Number(slashMatch[2]));
        }
      }
    }
  }

  if (!date || Number.isNaN(date.getTime())) return null;
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
}

function getMonthLabel(monthKey) {
  if (!monthKey) return "Tidak ada data";
  const [year, month] = monthKey.split("-");
  const monthNames = ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agu", "Sep", "Okt", "Nov", "Des"];
  return `${monthNames[Number(month) - 1] || month} ${year}`;
}

function createSummaryRow(columnName, values) {
  if (!values.length) return null;
  const sum = values.reduce((acc, value) => acc + value, 0);
  const avg = sum / values.length;
  const sorted = [...values].sort((a, b) => a - b);
  const mid = Math.floor(sorted.length / 2);
  const median = sorted.length % 2 === 0 ? (sorted[mid - 1] + sorted[mid]) / 2 : sorted[mid];
  const variance = values.reduce((acc, value) => acc + Math.pow(value - avg, 2), 0) / values.length;
  return {
    Column: columnName,
    Sum: Number(sum.toFixed(4)),
    Avg: Number(avg.toFixed(4)),
    Median: Number(median.toFixed(4)),
    StdDev: Number(Math.sqrt(variance).toFixed(4)),
    Min: Math.min(...values),
    Max: Math.max(...values),
    Count: values.length,
  };
}

export function buildChartPayload(rows, header, snapshot, selectedYear = "all", selectedCategory = "all", selectedMetric = "all") {
  const { filteredRows, yearValues, categoryValues, yearIndex, dateIndex, categoryIndex, selectedMetric: snapshotMetric } = snapshot;
  const activeMetric = selectedMetric === "all" ? snapshotMetric : selectedMetric;
  const metricIndex = getMetricColumnIndex(header, activeMetric);
  const fallbackMetricIndex = header.findIndex((column, index) => {
    const columnName = String(column ?? "").toLowerCase();
    if (isIgnoredColumnName(column)) return false;

    const values = rows
      .map((row) => toNumber(row[index]))
      .filter((value) => value !== null && Number.isFinite(value));

    return values.length > 0 && !columnName.includes("year") && !columnName.includes("tahun") && !columnName.includes("date") && !columnName.includes("tanggal") && !columnName.includes("category") && !columnName.includes("kategori") && !columnName.includes("segment") && !columnName.includes("class") && !columnName.includes("department");
  });
  const resolvedMetricIndex = metricIndex >= 0 ? metricIndex : fallbackMetricIndex;

  if (!filteredRows.length || resolvedMetricIndex < 0) {
    return null;
  }

  const metricLabel = header[resolvedMetricIndex] ?? activeMetric ?? "Nilai";
  const metricValues = filteredRows
    .map((row) => toNumber(row[resolvedMetricIndex]))
    .filter((value) => value !== null && Number.isFinite(value));
  const metricTotal = metricValues.reduce((sum, value) => sum + value, 0);
  const monthTotals = filteredRows.reduce((acc, row) => {
    const monthKey = getMonthKey(row, dateIndex);
    const currentMetric = toNumber(row[resolvedMetricIndex]);
    if (!monthKey || currentMetric === null || !Number.isFinite(currentMetric)) {
      return acc;
    }
    acc[monthKey] = (acc[monthKey] ?? 0) + currentMetric;
    return acc;
  }, {});
  const monthKeys = Object.keys(monthTotals).sort();
  const labels = monthKeys.length ? monthKeys.map(getMonthLabel) : ["Tidak ada data"];
  const data = monthKeys.length ? monthKeys.map((monthKey) => monthTotals[monthKey] ?? 0) : [0];

  return {
    labels,
    datasets: [{ label: `${metricLabel} (${metricTotal.toLocaleString()})`, data: data.length ? data : [0] }],
    title: `Total ${metricLabel} per Bulan`,
  };
}

export function buildDashboardSnapshot(rows, header, selectedYear = "all", selectedCategory = "all", selectedMetric = "all") {
  const yearIndex = findHeaderIndex(header, ["year", "tahun", "thn"]);
  const dateIndex = findHeaderIndex(header, ["date", "tanggal", "tanggaltransaksi", "orderdate", "order date", "createdate", "created at"]);
  const categoryIndex = findHeaderIndex(header, ["category", "kategori", "segment", "class", "department", "productcategory", "product category", "jenis"]);

  const normalizedRows = rows.filter((row) => {
    const rowYear = getYearValue(row, yearIndex, dateIndex);
    const rowCategory = getCategoryValue(row, categoryIndex);

    const yearMatch = selectedYear === "all" || rowYear === selectedYear;
    const categoryMatch = selectedCategory === "all" || rowCategory === selectedCategory;
    return yearMatch && categoryMatch;
  });

  const summaryRows = header.reduce((acc, columnName, index) => {
    if (!columnName || isIgnoredColumnName(columnName)) return acc;
    const values = normalizedRows
      .map((row) => toNumber(row[index]))
      .filter((value) => value !== null && Number.isFinite(value));

    if (!values.length) return acc;
    const summaryRow = createSummaryRow(columnName, values);
    if (summaryRow) acc.push(summaryRow);
    return acc;
  }, []);

  const allYears = Array.from(
    new Set(rows.map((row) => getYearValue(row, yearIndex, dateIndex)).filter((value) => value !== null))
  ).sort();

  const allCategories = Array.from(
    new Set(rows.map((row) => getCategoryValue(row, categoryIndex)).filter((value) => value !== null && value !== ""))
  ).sort();

  const metricOptions = getMetricOptions(header);
  const matchedMetricKey = metricOptions.find((option) => option.value === selectedMetric)
    ? selectedMetric
    : metricOptions.find((option) => option.value !== "all")?.value ?? "all";

  return {
    filteredRows: normalizedRows,
    summaryRows,
    yearValues: allYears,
    categoryValues: allCategories,
    metricOptions,
    selectedMetric: matchedMetricKey,
    yearIndex,
    dateIndex,
    categoryIndex,
  };
}

export function buildExcelExportData({ analysisData = [], chartPayload = null, header = [], sourceRows = [], selectedYear = "all", selectedCategory = "all", selectedMetric = "all" }) {
  /** @type {(string|number)[][]} */
  const sheet1Rows = [];

  sheet1Rows.push(["Hasil Analisis Excel Analyzer"]);
  sheet1Rows.push(["Filter Tahun", selectedYear === "all" ? "Semua Tahun" : selectedYear]);
  sheet1Rows.push(["Filter Kategori", selectedCategory === "all" ? "Semua Kategori" : selectedCategory]);
  sheet1Rows.push(["Filter Statistik", selectedMetric === "all" ? "Semua Statistik" : selectedMetric]);
  sheet1Rows.push([]);
  sheet1Rows.push(["Ringkasan Statistik"]);
  sheet1Rows.push(["Kolom", "Sum", "Rata-rata", "Median", "Std Dev", "Min", "Max", "Jumlah Data"]);

  analysisData.forEach((row) => {
    sheet1Rows.push([row.Column, row.Sum, row.Avg, row.Median, row.StdDev, row.Min, row.Max, row.Count]);
  });

  sheet1Rows.push([]);
  sheet1Rows.push(["Diagram Batang"]);
  sheet1Rows.push(["Label", "Nilai"]);
  if (chartPayload?.labels?.length) {
    chartPayload.labels.forEach((label, index) => {
      sheet1Rows.push([label, chartPayload.datasets?.[0]?.data?.[index] ?? 0]);
    });
  } else {
    sheet1Rows.push(["Tidak ada data", 0]);
  }

  sheet1Rows.push([]);
  sheet1Rows.push(["Diagram Lingkaran"]);
  sheet1Rows.push(["Label", "Nilai"]);
  if (chartPayload?.labels?.length) {
    chartPayload.labels.forEach((label, index) => {
      sheet1Rows.push([label, chartPayload.datasets?.[0]?.data?.[index] ?? 0]);
    });
  } else {
    sheet1Rows.push(["Tidak ada data", 0]);
  }

  sheet1Rows.push([]);
  sheet1Rows.push(["Diagram Garis"]);
  sheet1Rows.push(["Label", "Nilai"]);
  if (chartPayload?.labels?.length) {
    chartPayload.labels.forEach((label, index) => {
      sheet1Rows.push([label, chartPayload.datasets?.[0]?.data?.[index] ?? 0]);
    });
  } else {
    sheet1Rows.push(["Tidak ada data", 0]);
  }

  const sheet2Rows = header.length ? [header, ...sourceRows] : sourceRows;

  return { sheet1Rows, sheet2Rows };
}
