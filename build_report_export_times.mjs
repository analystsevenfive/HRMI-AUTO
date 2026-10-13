import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const [, , jsonPath, outputPath, mode] = process.argv;
if (!jsonPath || !outputPath) throw new Error("usage: node build_report_export_times.mjs <input.json> <output.xlsx>");

let rows = [];
try { rows = JSON.parse(await fs.readFile(jsonPath, "utf8")); }
catch (error) { if (error.code !== "ENOENT") throw error; }

const REQUIRED = ["[BB]", "[BK]", "[NTS]", "[PK]", "[SM]"];
const TYPES = ["ขาด", "สาย"];
const displayRows = [];
const runOrder = [...new Set(rows.map((row) => row.run_id).filter(Boolean))];
for (const runId of runOrder) {
  const runRows = rows.filter((row) => row.run_id === runId);
  displayRows.push(...runRows.map((row) => ({ ...row, is_summary: false })));
  const successful = new Set(runRows.filter((row) => row.status === "สำเร็จ")
    .map((row) => `${row.organization}|${row.report_type}`));
  if (!REQUIRED.every((org) => TYPES.every((type) => successful.has(`${org}|${type}`)))) continue;
  const first = runRows.reduce((value, row) => new Date(row.started_at) < new Date(value.started_at) ? row : value);
  const last = runRows.reduce((value, row) => new Date(row.finished_at) > new Date(value.finished_at) ? row : value);
  displayRows.push({
    run_id: runId, organization: "สรุปรอบ", report_type: "ครบ 10 รายงาน",
    target_start_date: runRows[0].target_start_date, target_end_date: runRows[0].target_end_date,
    started_at: first.started_at, finished_at: last.finished_at,
    duration_seconds: (new Date(last.finished_at) - new Date(first.started_at)) / 1000,
    status: "สำเร็จครบ 5 สาขา", path: "", details: "", is_summary: true,
  });
}

const excelDate = (value) => value ? new Date(value.replace(/(?:Z|[+-]\d{2}:\d{2})$/, "") + "Z") : null;
const rangeText = (row) => {
  const format = (value) => { const [y, m, d] = value.split("-"); return `${d}/${m}/${y}`; };
  return row.target_start_date === row.target_end_date ? format(row.target_start_date)
    : `${format(row.target_start_date)} - ${format(row.target_end_date)}`;
};

const workbook = Workbook.create();
const sheet = workbook.worksheets.add("เวลา Export รายงาน");
sheet.showGridLines = false;
sheet.freezePanes.freezeRows(4);
sheet.tabColor = "#1F4E78";
sheet.getRange("A2:J2").merge();
sheet.getRange("A2").values = [["ระยะเวลา Export รายงานขาดและสาย รายสาขา"]];
sheet.getRange("A2").format = { font: { name: "Arial", size: 14, bold: true, color: "#1F1F1F" } };
sheet.getRange("A4:J4").values = [["รอบการทำงาน", "สาขา", "ประเภทรายงาน", "ช่วงวันที่", "เวลาเริ่ม", "เวลาสิ้นสุด", "ระยะเวลา", "สถานะ", "ไฟล์", "รายละเอียด"]];
sheet.getRange("A4:J4").format = { fill: "#1F4E78", font: { name: "Arial", size: 10, bold: true, color: "#FFFFFF" }, horizontalAlignment: "center", verticalAlignment: "center" };
if (displayRows.length) {
  const values = displayRows.map((row) => [row.run_id, row.organization, row.report_type, rangeText(row),
    excelDate(row.started_at), excelDate(row.finished_at), Number(row.duration_seconds || 0) / 86400,
    row.status, row.path || "", row.details || ""]);
  const lastRow = 4 + values.length;
  sheet.getRange(`A5:J${lastRow}`).values = values;
  sheet.getRange(`A5:J${lastRow}`).format = { font: { name: "Arial", size: 10, color: "#1F1F1F" }, verticalAlignment: "center" };
  sheet.getRange(`E5:F${lastRow}`).format.numberFormat = "dd/mm/yyyy hh:mm:ss";
  sheet.getRange(`G5:G${lastRow}`).format.numberFormat = "[h]:mm:ss";
  sheet.getRange(`H5:H${lastRow}`).conditionalFormats.add("containsText", { text: "สำเร็จ", format: { fill: "#E2F0D9", font: { color: "#375623", bold: true } } });
  sheet.getRange(`H5:H${lastRow}`).conditionalFormats.add("containsText", { text: "ผิดพลาด", format: { fill: "#FCE4D6", font: { color: "#C00000", bold: true } } });
  displayRows.forEach((row, index) => {
    const target = sheet.getRange(`A${index + 5}:J${index + 5}`);
    if (row.is_summary) target.format = { fill: "#D9EAF7", font: { name: "Arial", size: 10, bold: true, color: "#17365D" }, borders: { top: { style: "medium", color: "#1F4E78" }, bottom: { style: "medium", color: "#1F4E78" } } };
    else if (index % 2 === 1) target.format.fill = "#F2F7FB";
  });
}
sheet.getRange("A:A").format.columnWidth = 23;
sheet.getRange("B:B").format.columnWidth = 12;
sheet.getRange("C:C").format.columnWidth = 19;
sheet.getRange("D:D").format.columnWidth = 25;
sheet.getRange("E:F").format.columnWidth = 23;
sheet.getRange("G:G").format.columnWidth = 14;
sheet.getRange("H:H").format.columnWidth = 20;
sheet.getRange("I:I").format.columnWidth = 52;
sheet.getRange("J:J").format.columnWidth = 36;
workbook.recalculate();
await fs.mkdir(path.dirname(outputPath), { recursive: true });
const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);
const lastRow = Math.max(7, 4 + displayRows.length);
const check = await workbook.inspect({ kind: "table", range: `เวลา Export รายงาน!A2:J${lastRow}`, include: "values,formulas", tableMaxRows: 15, tableMaxCols: 10 });
const errors = await workbook.inspect({ kind: "match", searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!", options: { useRegex: true, maxResults: 100 } });
process.stdout.write(`${check.ndjson}\n${errors.ndjson}`);
if (mode === "--preview") {
  const preview = await workbook.render({ sheetName: "เวลา Export รายงาน", range: `A1:J${lastRow}`, scale: 1.4, format: "png" });
  await fs.writeFile(path.join(path.dirname(jsonPath), "HRMI-report-export-times-preview.png"), new Uint8Array(await preview.arrayBuffer()));
}
