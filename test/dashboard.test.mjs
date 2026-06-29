import test from "node:test";
import assert from "node:assert/strict";
import { buildChartPayload, buildDashboardSnapshot, getMetricColumnIndex, normalizeSheetRows } from "../src/lib/dashboard-data.mjs";

test("buildDashboardSnapshot filters rows by year and category", () => {
  const header = ["Year", "Category", "Sales", "Quantity", "Discount", "Profit"];
  const rows = [
    [2023, "A", 100, 10, 5, 20],
    [2023, "B", 200, 20, 10, 40],
    [2024, "A", 300, 30, 15, 60],
  ];

  const snapshot = buildDashboardSnapshot(rows, header, "2023", "A", "sales");

  assert.equal(snapshot.filteredRows.length, 1);
  assert.equal(snapshot.filteredRows[0][2], 100);
  assert.equal(snapshot.summaryRows.find((item) => item.Column === "Sales")?.Sum, 100);
  assert.deepEqual(snapshot.yearValues, ["2023", "2024"]);
  assert.deepEqual(snapshot.categoryValues, ["A", "B"]);
  assert.equal(snapshot.selectedMetric, "sales");
});

test("getMetricColumnIndex detects sales and profit aliases", () => {
  const header = ["Revenue", "Qty", "Net Profit"];

  assert.equal(getMetricColumnIndex(header, "sales"), 0);
  assert.equal(getMetricColumnIndex(header, "quantity"), 1);
  assert.equal(getMetricColumnIndex(header, "profit"), 2);
});

test("normalizeSheetRows detects the real header row when a title row precedes it", () => {
  const rawRows = [
    ["Laporan Penjualan 2024"],
    ["Tanggal", "Kategori", "Sales", "Profit"],
    ["2024-01-15", "A", 100, 20],
    ["2024-02-20", "B", 200, 30],
  ];

  const normalized = normalizeSheetRows(rawRows);

  assert.deepEqual(normalized.header, ["Tanggal", "Kategori", "Sales", "Profit"]);
  assert.equal(normalized.rows.length, 2);
  assert.deepEqual(normalized.rows[0], ["2024-01-15", "A", 100, 20]);
});

test("buildChartPayload uses the date column even when a separate year column exists", () => {
  const header = ["Order ID", "Order Date", "Tahun", "Customer Name", "Kategori", "Sales"];
  const rows = [
    ["CA-1", "2014-03-15", 2014, "Astra", "Consumer", 220.5],
    ["CA-2", "2014-06-22", 2014, "Indomobil", "Corporate", 170],
    ["CA-3", "2015-01-20", 2015, "Grab", "Corporate", 311.6],
  ];
  const snapshot = buildDashboardSnapshot(rows, header, "all", "all", "all");

  const payload = buildChartPayload(rows, header, snapshot, "all", "all", "all");

  assert.equal(snapshot.dateIndex, 1);
  assert.deepEqual(payload.labels, ["Mar 2014", "Jun 2014", "Jan 2015"]);
  assert.deepEqual(payload.datasets[0].data, [220.5, 170, 311.6]);
});

test("buildChartPayload groups values by month using the date column", () => {
  const header = ["ID", "Tanggal", "Kategori", "Sales"];
  const rows = [
    [1, "2023-01-01", "A", 100],
    [2, "2023-02-01", "B", 200],
    [3, "2024-01-01", "A", 300],
  ];
  const snapshot = buildDashboardSnapshot(rows, header, "all", "all", "all");

  const payload = buildChartPayload(rows, header, snapshot, "all", "all", "all");

  assert.deepEqual(payload.labels, ["Jan 2023", "Feb 2023", "Jan 2024"]);
  assert.deepEqual(payload.datasets[0].data, [100, 200, 300]);
  assert.equal(payload.title, "Total Sales per Bulan");
});
