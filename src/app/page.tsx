"use client";

export const dynamic = "force-static";
import { useMemo, useRef, useState, type RefObject } from "react";
import * as XLSX from "xlsx";
import { buildChartPayload, buildDashboardSnapshot, buildExcelExportData, getMetricColumnIndex, normalizeSheetRows } from "@/lib/dashboard-data.mjs";

const PALETTE = [
  "#6366F1",
  "#10B981",
  "#F59E0B",
  "#EF4444",
  "#0EA5E9",
  "#8B5CF6",
  "#EC4899",
  "#14B8A6",
  "#F97316",
  "#84CC16",
];

type RowData = Array<string | number | null | undefined>;

interface SummaryRow {
  Column: string;
  Sum: number;
  Avg: number;
  Median: number;
  StdDev: number;
  Min: number;
  Max: number;
  Count: number;
}

interface MetricCard {
  key: string;
  label: string;
  value: number;
  accent: string;
  background: string;
}

interface SimpleChartProps {
  labels: string[];
  data: number[];
  color: string;
}

function SimpleBarChart({ labels, data, color }: SimpleChartProps) {
  const maxValue = Math.max(...data, 1);
  const labelCount = labels.length;
  const width = Math.max(360, labelCount * 54 + 70);
  const height = 240;
  const padding = 32;
  const innerHeight = height - padding * 2;
  const barGap = Math.max(10, Math.min(18, labelCount > 8 ? 8 : 14));
  const barWidth = Math.max(20, Math.min(42, (width - padding * 2 - barGap * (labelCount - 1)) / labelCount));
  const labelSize = labelCount > 8 ? 8 : 10;

  return (
    <div style={{ width: "100%", height: "100%", overflowX: "auto", paddingBottom: 6 }}>
      <svg viewBox={`0 0 ${width} ${height}`} width={width} height={height} role="img" aria-label="Diagram batang">
        <rect x="0" y="0" width={width} height={height} rx="16" fill="#fff" />
        {[0, 1, 2].map((line) => (
          <line key={line} x1={padding} y1={padding + (innerHeight / 3) * line} x2={width - padding} y2={padding + (innerHeight / 3) * line} stroke="#E5E7EB" strokeDasharray="4 4" />
        ))}
        {data.map((value, index) => {
          const barHeight = Math.max(18, (value / maxValue) * (innerHeight - 10));
          const x = padding + index * (barWidth + barGap);
          const y = height - padding - barHeight;
          return (
            <g key={`${labels[index]}-${index}`}>
              <rect x={x} y={y} width={barWidth} height={barHeight} rx="10" fill={color} opacity="0.95" />
              <rect x={x} y={y} width={barWidth} height={barHeight} rx="10" fill="url(#barGradient)" opacity="0.22" />
              <text x={x + barWidth / 2} y={y - 8} textAnchor="middle" fontSize="11" fontWeight="700" fill="#111827">{value.toLocaleString()}</text>
              <text x={x + barWidth / 2} y={height - 10} textAnchor="middle" fontSize={labelSize} fill="#6B7280">{labels[index]}</text>
            </g>
          );
        })}
        <defs>
          <linearGradient id="barGradient" x1="0%" y1="0%" x2="0%" y2="100%">
            <stop offset="0%" stopColor="#ffffff" />
            <stop offset="100%" stopColor="#ffffff" stopOpacity="0" />
          </linearGradient>
        </defs>
      </svg>
    </div>
  );
}

function SimpleLineChart({ labels, data, color }: SimpleChartProps) {
  const maxValue = Math.max(...data, 1);
  const labelCount = labels.length;
  const width = Math.max(360, labelCount * 56 + 70);
  const height = 240;
  const padding = 32;
  const innerWidth = width - padding * 2;
  const innerHeight = height - padding * 2;
  const points = data.map((value, index) => ({
    x: padding + (innerWidth / Math.max(data.length - 1, 1)) * index,
    y: height - padding - (value / maxValue) * innerHeight,
    value,
  }));
  const path = points.map((point, index) => `${index === 0 ? "M" : "L"} ${point.x} ${point.y}`).join(" ");
  const areaPath = `${path} L ${points[points.length - 1].x} ${height - padding} L ${points[0].x} ${height - padding} Z`;

  return (
    <div style={{ width: "100%", height: "100%", overflowX: "auto", paddingBottom: 6 }}>
      <svg viewBox={`0 0 ${width} ${height}`} width={width} height={height} role="img" aria-label="Diagram garis">
        <rect x="0" y="0" width={width} height={height} rx="16" fill="#fff" />
        {[0, 1, 2].map((line) => (
          <line key={line} x1={padding} y1={padding + (innerHeight / 3) * line} x2={width - padding} y2={padding + (innerHeight / 3) * line} stroke="#E5E7EB" strokeDasharray="4 4" />
        ))}
        <path d={areaPath} fill={`url(#lineFill)`} opacity="0.16" />
        <path d={path} fill="none" stroke={color} strokeWidth="3.2" strokeLinecap="round" strokeLinejoin="round" />
        {points.map((point, index) => (
          <g key={`${labels[index]}-${index}`}>
            <circle cx={point.x} cy={point.y} r="5.2" fill="#fff" stroke={color} strokeWidth="3" />
            <circle cx={point.x} cy={point.y} r="9.6" fill="transparent" stroke={color} strokeOpacity="0.16" />
            <text x={point.x} y={height - 10} textAnchor="middle" fontSize={labelCount > 8 ? 8 : 10} fill="#6B7280">{labels[index]}</text>
            <text x={point.x} y={point.y - 10} textAnchor="middle" fontSize="10" fontWeight="700" fill="#111827">{point.value.toLocaleString()}</text>
          </g>
        ))}
        <defs>
          <linearGradient id="lineFill" x1="0%" y1="0%" x2="0%" y2="100%">
            <stop offset="0%" stopColor={color} />
            <stop offset="100%" stopColor={color} stopOpacity="0" />
          </linearGradient>
        </defs>
      </svg>
    </div>
  );
}

function SimplePieChart({ labels, data, color }: SimpleChartProps) {
  const total = Math.max(data.reduce((sum, value) => sum + value, 0), 1);
  const segments = data.map((value) => (value / total) * 360);
  const sliceColors = [color, "#10B981", "#0EA5E9", "#F59E0B", "#8B5CF6"];
  const pathData = segments.map((slice, index) => {
    const prevSum = segments.slice(0, index).reduce((acc, v) => acc + v, 0);
    const startAngle = -90 + prevSum;
    const endAngle = startAngle + slice;
    const largeArcFlag = slice > 180 ? 1 : 0;
    const start = {
      x: 88 + 60 * Math.cos((startAngle * Math.PI) / 180),
      y: 88 + 60 * Math.sin((startAngle * Math.PI) / 180),
    };
    const end = {
      x: 88 + 60 * Math.cos((endAngle * Math.PI) / 180),
      y: 88 + 60 * Math.sin((endAngle * Math.PI) / 180),
    };
    return { d: `M 88 88 L ${start.x} ${start.y} A 60 60 0 ${largeArcFlag} 1 ${end.x} ${end.y} Z`, color: sliceColors[index % sliceColors.length] };
  });


  return (
    <div style={{ width: "100%", height: "100%", display: "flex", alignItems: "center", justifyContent: "center", gap: 16, flexWrap: "wrap" }}>
      <div style={{ width: 176, height: 176, borderRadius: "50%", background: "#fff", padding: 10, boxShadow: "0 10px 30px rgba(15, 23, 42, 0.08)", flexShrink: 0 }}>
        <svg viewBox="0 0 176 176" width="100%" height="100%" role="img" aria-label="Diagram lingkaran">
          {pathData.map((slice, index) => (
            <path key={`${labels[index]}-${index}`} d={slice.d} fill={slice.color} opacity="0.94" />
          ))}
          <circle cx="88" cy="88" r="42" fill="#fff" />
          <text x="88" y="82" textAnchor="middle" fontSize="14" fontWeight="800" fill="#111827">{total.toLocaleString()}</text>
          <text x="88" y="100" textAnchor="middle" fontSize="11" fill="#6B7280">Total</text>
        </svg>
      </div>
      <div style={{ display: "flex", flexDirection: "column", gap: 8, minWidth: 150, maxWidth: 260, maxHeight: 220, overflowY: "auto", paddingRight: 4 }}>
        {labels.map((label, index) => (
          <div key={`${label}-${index}`} style={{ display: "flex", alignItems: "center", gap: 8, fontSize: labels.length > 8 ? 10 : 12, color: "#374151", whiteSpace: "nowrap" }}>
            <span style={{ width: 10, height: 10, borderRadius: "50%", background: sliceColors[index % sliceColors.length], flexShrink: 0 }} />
            <span style={{ flex: 1, minWidth: 0, overflow: "hidden", textOverflow: "ellipsis" }}>{label}</span>
            <strong>{data[index].toLocaleString()}</strong>
          </div>
        ))}
      </div>
    </div>
  );
}

export default function Home() {
  const [file1, setFile1] = useState<File | null>(null);
  const [file2, setFile2] = useState<File | null>(null);
  const [resultMessage, setResultMessage] = useState("");
  const [isSuccess, setIsSuccess] = useState(false);
  const [loading, setLoading] = useState(false);
  const [mergedData, setMergedData] = useState<RowData[][]>([]);
  const [header, setHeader] = useState<string[]>([]);
  const [rows, setRows] = useState<RowData[][]>([]);
  const [selectedYear, setSelectedYear] = useState<string>("all");
  const [selectedCategory, setSelectedCategory] = useState<string>("all");
  const [selectedMetric, setSelectedMetric] = useState<string>("all");

  const barRef = useRef<HTMLDivElement | null>(null);
  const pieRef = useRef<HTMLDivElement | null>(null);
  const lineRef = useRef<HTMLDivElement | null>(null);

  const handleFileChange =
    (setter: (f: File | null) => void) =>
    (e: React.ChangeEvent<HTMLInputElement>) => {
      if (e.target.files?.[0]) setter(e.target.files[0]);
    };

  const dashboard = useMemo(() => {
    if (!rows.length || !header.length) {
      return {
        analysisData: [] as SummaryRow[],
        availableYears: [] as string[],
        availableCategories: [] as string[],
        metricOptions: [] as { value: string; label: string; description: string }[],
        filteredRowsCount: 0,
        chartPayload: null as { labels: string[]; datasets: { label: string; data: number[] }[]; title: string } | null,
      };
    }

    const snapshot = buildDashboardSnapshot(rows, header, selectedYear, selectedCategory, selectedMetric);
    const chartPayload = buildChartPayload(rows, header, snapshot, selectedYear, selectedCategory, selectedMetric);

    return {
      analysisData: snapshot.summaryRows,
      availableYears: snapshot.yearValues,
      availableCategories: snapshot.categoryValues,
      metricOptions: snapshot.metricOptions,
      filteredRowsCount: snapshot.filteredRows.length,
      filteredRows: snapshot.filteredRows,
      chartPayload,
    };
  }, [header, rows, selectedCategory, selectedMetric, selectedYear]);

  const analysisData = dashboard.analysisData;
  const availableYears = dashboard.availableYears;
  const availableCategories = dashboard.availableCategories;
  const metricOptions = dashboard.metricOptions;
  const filteredRowsCount = dashboard.filteredRowsCount;
  const filteredRows = dashboard.filteredRows;
  const chartPayload = dashboard.chartPayload;


  const hasAnalysisData = header.length > 0 && rows.length > 0;

  const metricCards = useMemo<MetricCard[]>(() => {
    const metricMap = [
      { key: "sales", label: "Sales", accent: "#6366F1", background: "#EEF2FF" },
      { key: "quantity", label: "Quantity", accent: "#0EA5E9", background: "#E0F2FE" },
      { key: "discount", label: "Discount", accent: "#F59E0B", background: "#FEF3C7" },
      { key: "profit", label: "Profit", accent: "#10B981", background: "#ECFDF5" },
    ];

    return metricMap
      .map((metric) => {
        const metricIndex = getMetricColumnIndex(header, metric.key);
        if (metricIndex < 0) return null;

        const values = filteredRows
          .map((row: RowData) => Number(row[metricIndex]))
          .filter((value: number) => Number.isFinite(value));

        return {
          key: metric.key,
          label: metric.label,
          value: values.reduce((sum: number, value: number) => sum + value, 0),
          accent: metric.accent,
          background: metric.background,
        };
      })
      .filter((card): card is MetricCard => Boolean(card));
  }, [filteredRows, header]);


  const analyze = async () => {
    if (!file1) {
      setIsSuccess(false);
      setResultMessage("Upload minimal satu file Excel terlebih dahulu.");
      return;
    }

    setLoading(true);
    setResultMessage("");
    try {
      const readWorkbook = (file: File) =>
        new Promise<XLSX.WorkBook>((resolve, reject) => {
          const reader = new FileReader();
          reader.onload = (event) => {
            const data = event.target?.result;
            if (data) resolve(XLSX.read(data, { type: "array" }));
            else reject(new Error("Read failed"));
          };
          reader.onerror = reject;
          reader.readAsArrayBuffer(file);
        });

      const workbook1 = await readWorkbook(file1);
      const workbook2 = file2 ? await readWorkbook(file2) : null;
      const sheet1 = workbook1.Sheets[workbook1.SheetNames[0]];
      const sheet2 = workbook2 ? workbook2.Sheets[workbook2.SheetNames[0]] : null;
      const rows1: RowData[][] = sheet1 ? (XLSX.utils.sheet_to_json(sheet1, { header: 1 }) as RowData[][]) : [];
      const rows2: RowData[][] = sheet2 ? (XLSX.utils.sheet_to_json(sheet2, { header: 1 }) as RowData[][]).slice(1) : [];
      const merged: RowData[][] = [...rows1, ...rows2];

      const normalizedSheet = normalizeSheetRows(merged);
      const nextHeader = normalizedSheet.header.map((cell) => String(cell ?? ""));
      const nextRows = normalizedSheet.rows;
      const normalizedMerged = [nextHeader, ...nextRows];

      setHeader(nextHeader);
      setRows(nextRows);
      setMergedData(normalizedMerged);
      setSelectedYear("all");
      setSelectedCategory("all");
      setSelectedMetric("all");
      setIsSuccess(true);
      setResultMessage("Analisis selesai. Filter tahun, kategori, dan metrik siap dipakai.");
    } catch (error) {
      console.error(error);
      setIsSuccess(false);
      setResultMessage("Terjadi kesalahan saat analisis. Periksa konsol untuk detail.");
    } finally {
      setLoading(false);
    }
  };

  const saveChart = (ref: React.RefObject<HTMLDivElement | null>, name: string) => {
    if (!ref.current) return;
    const svgMarkup = `<svg xmlns="http://www.w3.org/2000/svg" width="1200" height="800"><rect width="100%" height="100%" fill="white"/><text x="40" y="60" font-size="28" font-family="Segoe UI">${name}</text></svg>`;
    const blob = new Blob([svgMarkup], { type: "image/svg+xml;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const link = Object.assign(document.createElement("a"), { href: url, download: name });
    link.click();
    URL.revokeObjectURL(url);
  };

  const downloadExcel = async () => {
    if (!analysisData.length) {
      setResultMessage("Belum ada hasil untuk diunduh.");
      return;
    }
    setLoading(true);
    setResultMessage("Menyiapkan berkas Excel...");
    try {
      const ExcelJS = (await import("exceljs")).default;
      const workbook = new ExcelJS.Workbook();
      const exportData = (buildExcelExportData as unknown as (args: any) => any)({
        analysisData,
        chartPayload,
        header: header as any,
        sourceRows: mergedData as any,

        selectedYear,
        selectedCategory,
        selectedMetric,
      });








      const sheet = workbook.addWorksheet("Sheet 1");
      exportData.sheet1Rows.forEach((row: any[]) => sheet.addRow(row));


      if (exportData.sheet2Rows.length) {
        const sourceSheet = workbook.addWorksheet("Sheet 2");
        exportData.sheet2Rows.forEach((row: any[]) => sourceSheet.addRow(row));
      }



      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      const url = URL.createObjectURL(blob);
      const link = Object.assign(document.createElement("a"), { href: url, download: "laporan-analisis.xlsx" });
      link.click();
      URL.revokeObjectURL(url);
      setResultMessage("File Excel berhasil diunduh dengan hasil analisis dan diagram.");
    } catch (error) {
      console.error(error);
      setResultMessage("Gagal membuat dan mengunduh Excel. Silakan coba lagi.");
    } finally {
      setLoading(false);
    }
  };

  const primaryBtn = (active: boolean): React.CSSProperties => ({
    flex: "1 1 160px",
    minHeight: 46,
    padding: "13px 20px",
    background: active ? "linear-gradient(135deg,#6366F1,#8B5CF6)" : "#E5E7EB",
    color: active ? "#fff" : "#9CA3AF",
    border: "none",
    borderRadius: 12,
    fontWeight: 700,
    fontSize: 15,
    cursor: active ? "pointer" : "not-allowed",
    boxShadow: active ? "0 4px 14px rgba(99,102,241,0.38)" : "none",
    transition: "all .2s",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    gap: 8,
  });

  const greenBtn = (active: boolean): React.CSSProperties => ({
    ...primaryBtn(active),
    background: active ? "linear-gradient(135deg,#10B981,#059669)" : "#E5E7EB",
    boxShadow: active ? "0 4px 14px rgba(16,185,129,0.35)" : "none",
  });

  const chartDlBtn: React.CSSProperties = {
    padding: "6px 13px",
    border: "none",
    borderRadius: 8,
    fontWeight: 600,
    fontSize: 12,
    cursor: "pointer",
    color: "#fff",
    background: "linear-gradient(135deg,#6366F1,#0EA5E9)",
    boxShadow: "0 3px 8px rgba(99,102,241,0.3)",
  };

  const card: React.CSSProperties = {
    background: "#fff",
    borderRadius: 16,
    boxShadow: "0 2px 16px rgba(0,0,0,0.07)",
    border: "1px solid #E9ECF5",
    overflow: "hidden",
  };

  return (
    <div style={{ minHeight: "100vh", background: "linear-gradient(145deg,#EEF2FF 0%,#F0F9FF 45%,#ECFDF5 100%)", fontFamily: "'Segoe UI',system-ui,-apple-system,sans-serif" }}>
      <nav style={{ position: "sticky", top: 0, zIndex: 20, minHeight: 58, display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10, padding: "0 clamp(12px, 3vw, 20px)", background: "rgba(255,255,255,0.88)", backdropFilter: "blur(14px)", borderBottom: "1px solid rgba(99,102,241,0.1)", flexWrap: "wrap" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 9 }}>
          <div style={{ width: 32, height: 32, borderRadius: 9, background: "linear-gradient(135deg,#6366F1,#8B5CF6)", display: "flex", alignItems: "center", justifyContent: "center" }}>
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="#fff" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
              <path d="M9 17v-2m3 2v-4m3 4v-6M5 20h14a2 2 0 002-2V6a2 2 0 00-2-2H5a2 2 0 00-2 2v12a2 2 0 002 2z" />
            </svg>
          </div>
          <span style={{ fontWeight: 800, fontSize: 16, color: "#1E1B4B" }}>Excel Analyzer</span>
        </div>
        <span style={{ fontSize: 11, color: "#6B7280", background: "#F3F4F6", padding: "3px 10px", borderRadius: 20, fontWeight: 600 }}>v2.1</span>
      </nav>

      <main style={{ maxWidth: 1100, margin: "0 auto", padding: "24px clamp(12px, 3vw, 24px) 72px" }}>
        <div style={{ textAlign: "center", marginBottom: 32 }}>
          <h1 style={{ fontSize: "clamp(24px,5vw,38px)", fontWeight: 800, color: "#1E1B4B", lineHeight: 1.25, margin: "0 0 10px", letterSpacing: "-0.4px" }}>
            Analisis Data Excel <span style={{ background: "linear-gradient(90deg,#6366F1,#06B6D4)", WebkitBackgroundClip: "text", WebkitTextFillColor: "transparent" }}>Otomatis</span>
          </h1>
          <p style={{ color: "#6B7280", fontSize: 15, maxWidth: 620, margin: "0 auto", lineHeight: 1.65 }}>
            Upload file Excel, lalu filter berdasarkan tahun, kategori, dan metrik utama seperti sales, quantity, discount, dan profit — hasil statistik dan diagram akan berubah otomatis.
          </p>
        </div>

        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(min(100%, 260px), 1fr))", gap: 14, marginBottom: 16 }}>
          {[
            { id: "file1", label: "File 1", req: true, file: file1, set: setFile1, accent: "#6366F1", dash: "#C7D2FE", glow: "#EEF2FF", iconStroke: "#A78BFA" },
            { id: "file2", label: "File 2", req: false, file: file2, set: setFile2, accent: "#10B981", dash: "#A7F3D0", glow: "#ECFDF5", iconStroke: "#6EE7B7" },
          ].map(({ id, label, req, file, set, accent, dash, glow, iconStroke }) => (
            <label key={id} htmlFor={id} style={{ display: "block", background: "#fff", border: `2px dashed ${file ? accent : dash}`, borderRadius: 14, padding: 20, cursor: "pointer", boxShadow: file ? `0 0 0 3px ${glow}` : "0 2px 8px rgba(0,0,0,0.04)", transition: "all .2s" }}>
              <div style={{ display: "flex", alignItems: "center", gap: 14 }}>
                <div style={{ width: 42, height: 42, borderRadius: 11, background: glow, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
                  <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke={file ? accent : iconStroke} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                    <path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z" /><polyline points="14 2 14 8 20 8" />
                  </svg>
                </div>
                <div style={{ flex: 1, minWidth: 0 }}>
                  <div style={{ fontWeight: 700, fontSize: 13, color: "#1E1B4B", marginBottom: 3 }}>
                    {label} {req ? <span style={{ color: "#EF4444" }}>*</span> : <span style={{ color: "#9CA3AF", fontWeight: 400 }}>(opsional)</span>}
                  </div>
                  <div style={{ fontSize: 12, color: file ? accent : "#9CA3AF", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{file ? `✓ ${file.name}` : "Klik untuk pilih .xlsx / .xls"}</div>
                </div>
              </div>
              <input id={id} type="file" accept=".xlsx,.xls" onChange={handleFileChange(set)} style={{ display: "none" }} />
            </label>
          ))}
        </div>

        <div style={{ display: "flex", gap: 12, marginBottom: 18, flexWrap: "wrap" }}>
          <button id="analyze-btn" onClick={analyze} disabled={loading} style={primaryBtn(!loading)}>
            {loading ? (
              <><svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="#fff" strokeWidth="2.5" style={{ animation: "spin 1s linear infinite" }}><path d="M21 12a9 9 0 11-18 0 9 9 0 0118 0z" strokeOpacity="0.25" /><path d="M21 12a9 9 0 00-9-9" /></svg>Menganalisis…</>
            ) : (
              <><svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="#fff" strokeWidth="2.5" strokeLinecap="round"><circle cx="11" cy="11" r="8" /><line x1="21" y1="21" x2="16.65" y2="16.65" /></svg>Analisis Data</>
            )}
          </button>
          <button id="download-excel-btn" onClick={downloadExcel} disabled={!analysisData.length || loading} style={greenBtn(analysisData.length > 0 && !loading)}>
            <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round">
              <path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4" /><polyline points="7 10 12 15 17 10" /><line x1="12" y1="15" x2="12" y2="3" />
            </svg>
            Download Excel
          </button>
        </div>

        {resultMessage && (
          <div style={{ padding: "11px 16px", borderRadius: 10, marginBottom: 18, fontSize: 13, fontWeight: 500, display: "flex", alignItems: "center", gap: 8, background: isSuccess ? "#ECFDF5" : "#FEF2F2", color: isSuccess ? "#065F46" : "#991B1B", border: `1px solid ${isSuccess ? "#A7F3D0" : "#FECACA"}` }}>
            <span>{isSuccess ? "✅" : "❌"}</span>{resultMessage}
          </div>
        )}

        {hasAnalysisData && (
          <>
            <div className="filter-card" style={{ ...card, marginBottom: 18, padding: "18px 20px" }}>
              <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(min(100%, 220px), 1fr))", gap: 12 }}>
                <label style={{ display: "flex", flexDirection: "column", gap: 7 }}>
                  <span style={{ fontSize: 12, fontWeight: 700, color: "#4B5563", textTransform: "uppercase", letterSpacing: "0.4px" }}>Filter Tahun</span>
                  <select value={selectedYear} onChange={(event) => setSelectedYear(event.target.value)} style={{ width: "100%", padding: "10px 12px", borderRadius: 10, border: "1px solid #D1D5DB", background: "#F9FAFB", fontWeight: 600, color: "#374151" }}>
                    <option value="all">Semua Tahun</option>
                    {availableYears.map((year) => (
                      <option key={year} value={year}>Tahun {year}</option>
                    ))}
                  </select>
                </label>

                <label style={{ display: "flex", flexDirection: "column", gap: 7 }}>
                  <span style={{ fontSize: 12, fontWeight: 700, color: "#4B5563", textTransform: "uppercase", letterSpacing: "0.4px" }}>Filter Kategori</span>
                  <select value={selectedCategory} onChange={(event) => setSelectedCategory(event.target.value)} style={{ width: "100%", padding: "10px 12px", borderRadius: 10, border: "1px solid #D1D5DB", background: "#F9FAFB", fontWeight: 600, color: "#374151" }}>
                    <option value="all">Semua Kategori</option>
                    {availableCategories.map((category) => (
                      <option key={category} value={category}>{category}</option>
                    ))}
                  </select>
                </label>

                <label style={{ display: "flex", flexDirection: "column", gap: 7 }}>
                  <span style={{ fontSize: 12, fontWeight: 700, color: "#4B5563", textTransform: "uppercase", letterSpacing: "0.4px" }}>Filter Statistik</span>
                  <select value={selectedMetric} onChange={(event) => setSelectedMetric(event.target.value)} style={{ width: "100%", padding: "10px 12px", borderRadius: 10, border: "1px solid #D1D5DB", background: "#F9FAFB", fontWeight: 600, color: "#374151" }}>
                    {metricOptions.map((option) => (
                      <option key={option.value} value={option.value}>{option.label}</option>
                    ))}
                  </select>
                </label>
              </div>

              <div style={{ marginTop: 12, fontSize: 13, color: "#6B7280" }}>
                Menampilkan <strong style={{ color: "#111827" }}>{filteredRowsCount}</strong> baris data untuk <strong style={{ color: "#111827" }}>{selectedYear === "all" ? "semua tahun" : `tahun ${selectedYear}`}</strong> dan <strong style={{ color: "#111827" }}>{selectedCategory === "all" ? "semua kategori" : selectedCategory}</strong>.
              </div>
            </div>

            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(min(100%, 180px), 1fr))", gap: 12, marginBottom: 20 }}>
              {metricCards.map((metric) => (
                <div key={metric.key} style={{ background: metric.background, borderRadius: 14, padding: "16px 16px", border: `1px solid ${metric.accent}22` }}>
                  <div style={{ fontSize: 11, fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.5px", color: metric.accent, marginBottom: 6 }}>{metric.label}</div>
                  <div style={{ fontSize: 24, fontWeight: 800, color: "#111827" }}>{metric.value.toLocaleString()}</div>
                </div>
              ))}
            </div>

            <div style={{ ...card, marginBottom: 20 }}>
              <div style={{ padding: "16px 22px", borderBottom: "1px solid #F3F4F6", display: "flex", alignItems: "center", gap: 9 }}>
                <div style={{ width: 8, height: 8, borderRadius: "50%", background: "#6366F1" }} />
                <h2 style={{ margin: 0, fontSize: 15, fontWeight: 700, color: "#1E1B4B" }}>Ringkasan Statistik</h2>
              </div>
              <div style={{ overflowX: "auto" }}>
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
                  <thead>
                    <tr style={{ background: "#F8F9FF" }}>
                      {['Kolom', 'Sum', 'Rata-rata', 'Median', 'Std Dev', 'Min', 'Max', 'Data'].map((headerText, index) => (
                        <th key={headerText} style={{ padding: "11px 15px", textAlign: index === 0 ? "left" : "right", fontWeight: 700, fontSize: 10.5, color: "#6B7280", textTransform: "uppercase", letterSpacing: "0.4px", borderBottom: "2px solid #E9ECF5", whiteSpace: "nowrap" }}>{headerText}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {analysisData.map((row: SummaryRow, index: number) => (
                      <tr key={`${row.Column}-${index}`} style={{ background: index % 2 === 0 ? "#fff" : "#FAFBFF", borderBottom: "1px solid #F3F4F6" }}>
                        <td style={{ padding: "12px 15px", fontWeight: 600, color: "#6366F1", whiteSpace: "nowrap" }}>{row.Column}</td>
                        <td style={{ padding: "12px 15px", textAlign: "right", color: "#374151" }}>{row.Sum.toLocaleString()}</td>
                        <td style={{ padding: "12px 15px", textAlign: "right", color: "#374151" }}>{row.Avg.toFixed(2)}</td>
                        <td style={{ padding: "12px 15px", textAlign: "right", color: "#374151" }}>{row.Median.toFixed(2)}</td>
                        <td style={{ padding: "12px 15px", textAlign: "right", color: "#374151" }}>{row.StdDev.toFixed(2)}</td>
                        <td style={{ padding: "12px 15px", textAlign: "right", color: "#10B981", fontWeight: 700 }}>{row.Min}</td>
                        <td style={{ padding: "12px 15px", textAlign: "right", color: "#EF4444", fontWeight: 700 }}>{row.Max}</td>
                        <td style={{ padding: "12px 15px", textAlign: "right", color: "#9CA3AF" }}>{row.Count}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>

            <div style={{ display: "flex", flexDirection: "column", gap: 16, marginBottom: 16 }}>
              <div className="chart-card" style={card}>
                <div style={{ padding: "16px 20px", borderBottom: "1px solid #F3F4F6", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                    <div style={{ width: 10, height: 10, borderRadius: 3, background: "#6366F1" }} />
                    <span style={{ fontWeight: 700, fontSize: 14, color: "#1E1B4B" }}>Diagram Batang</span>
                  </div>
                  <button id="dl-bar" onClick={() => saveChart(barRef, "bar-chart.png")} style={chartDlBtn}>↓ Unduh</button>
                </div>
                <div ref={barRef} style={{ padding: "clamp(12px, 2.2vw, 20px)", height: "clamp(250px, 36vw, 320px)", minHeight: 250, position: "relative", width: "100%", background: "linear-gradient(135deg,#F9FAFB,#EEF2FF)" }}>
                  {chartPayload ? (
                    <SimpleBarChart labels={chartPayload.labels} data={chartPayload.datasets[0].data} color="#6366F1" />
                  ) : (
                    <div style={{ display: "flex", alignItems: "center", justifyContent: "center", height: "100%", color: "#9CA3AF", textAlign: "center", padding: 16 }}>
                      Belum ada metrik numerik yang bisa divisualisasikan dari data terfilter.
                    </div>
                  )}
                </div>
              </div>

              <div className="chart-card" style={card}>
                <div style={{ padding: "16px 20px", borderBottom: "1px solid #F3F4F6", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                    <div style={{ width: 10, height: 10, borderRadius: "50%", background: "#F59E0B" }} />
                    <span style={{ fontWeight: 700, fontSize: 14, color: "#1E1B4B" }}>Diagram Lingkaran</span>
                  </div>
                  <button id="dl-pie" onClick={() => saveChart(pieRef, "pie-chart.png")} style={{ ...chartDlBtn, background: "linear-gradient(135deg,#F59E0B,#F97316)" }}>↓ Unduh</button>
                </div>
                <div ref={pieRef} style={{ padding: "clamp(12px, 2.2vw, 20px)", height: "clamp(250px, 36vw, 320px)", minHeight: 250, position: "relative", width: "100%", background: "linear-gradient(135deg,#FFF7ED,#FEF3C7)" }}>
                  {chartPayload ? (
                    <SimplePieChart labels={chartPayload.labels} data={chartPayload.datasets[0].data} color="#F59E0B" />
                  ) : (
                    <div style={{ display: "flex", alignItems: "center", justifyContent: "center", height: "100%", color: "#9CA3AF", textAlign: "center", padding: 16 }}>
                      Belum ada metrik numerik yang bisa divisualisasikan dari data terfilter.
                    </div>
                  )}
                </div>
              </div>
            </div>

            <div className="chart-card" style={card}>
              <div style={{ padding: "16px 20px", borderBottom: "1px solid #F3F4F6", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
                <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                  <div style={{ width: 10, height: 10, borderRadius: 3, background: "#EF4444" }} />
                  <span style={{ fontWeight: 700, fontSize: 14, color: "#1E1B4B" }}>Diagram Garis</span>
                </div>
                <button id="dl-line" onClick={() => saveChart(lineRef, "line-chart.png")} style={{ ...chartDlBtn, background: "linear-gradient(135deg,#EF4444,#F97316)" }}>↓ Unduh</button>
              </div>
              <div ref={lineRef} style={{ padding: "clamp(12px, 2.2vw, 20px)", height: "clamp(260px, 38vw, 330px)", minHeight: 260, position: "relative", width: "100%", background: "linear-gradient(135deg,#FEF2F2,#FEE2E2)" }}>
                {chartPayload ? (
                  <SimpleLineChart labels={chartPayload.labels} data={chartPayload.datasets[0].data} color="#EF4444" />
                ) : (
                  <div style={{ display: "flex", alignItems: "center", justifyContent: "center", height: "100%", color: "#9CA3AF", textAlign: "center", padding: 16 }}>
                    Belum ada metrik numerik yang bisa divisualisasikan dari data terfilter.
                  </div>
                )}
              </div>
            </div>
          </>
        )}

        {!hasAnalysisData && !loading && (
          <div style={{ textAlign: "center", padding: "52px 24px", background: "#fff", borderRadius: 16, boxShadow: "0 2px 16px rgba(0,0,0,0.06)", border: "1px solid #E9ECF5" }}>
            <div style={{ width: 70, height: 70, borderRadius: 20, background: "#EEF2FF", display: "flex", alignItems: "center", justifyContent: "center", margin: "0 auto 16px" }}>
              <svg width="34" height="34" viewBox="0 0 24 24" fill="none" stroke="#6366F1" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round">
                <path d="M9 17v-2m3 2v-4m3 4v-6M5 20h14a2 2 0 002-2V6a2 2 0 00-2-2H5a2 2 0 00-2 2v12a2 2 0 002 2z" />
              </svg>
            </div>
            <h3 style={{ margin: "0 0 8px", fontSize: 17, fontWeight: 700, color: "#1E1B4B" }}>Belum ada data</h3>
            <p style={{ margin: 0, color: "#9CA3AF", fontSize: 14 }}>Upload file Excel lalu klik <strong style={{ color: "#6366F1" }}>Analisis Data</strong></p>
          </div>
        )}
      </main>

      <style>{`
        @keyframes spin { to { transform: rotate(360deg); } }
        * { box-sizing: border-box; }
        body { margin: 0; }
        .chart-canvas { max-width: 100% !important; max-height: 100% !important; }
        button:hover { opacity: 0.92; transform: translateY(-1px); }
        button:active { transform: scale(0.97) !important; opacity: 1; }
        label:hover { border-opacity: 0.8; }
        @media (max-width: 768px) {
          main { padding: 20px 10px 56px; }
          nav { padding: 0 12px; min-height: 64px; }
          .filter-card { padding: 14px 14px !important; }
          .chart-card > div:first-child { padding: 12px 14px !important; }
          .chart-card > div:last-child { padding: 12px !important; height: clamp(240px, 58vw, 280px) !important; min-height: 240px !important; }
          button, select { width: 100%; }
        }
        @media (max-width: 480px) {
          main { padding: 16px 8px 56px; }
          nav { padding: 0 10px; }
          .filter-card { padding: 12px !important; }
          .chart-card > div:last-child { padding: 10px !important; height: 230px !important; min-height: 230px !important; }
          h1 { font-size: 24px !important; }
          p { font-size: 14px !important; }
        }
      `}</style>
    </div>
  );
}
