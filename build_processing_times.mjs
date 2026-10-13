import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const [, , jsonPath, outputPath, mode] = process.argv;
if (!jsonPath || !outputPath) throw new Error("usage: node build_processing_times.mjs <input.json> <output.xlsx>");

let rows = [];
try { rows = JSON.parse(await fs.readFile(jsonPath, "utf8")); }
catch (error) { if (error.code !== "ENOENT") throw error; }

function excelBangkokDateTime(value) {
  if (!value) return null;
  const localPart = value.replace(/(?:Z|[+-]\d{2}:\d{2})$/, "");
  return new Date(`${localPart}Z`);
}

function processingRange(row) {
  const start = row.target_start_date || row.target_date;
  const end = row.target_end_date || row.target_date;
  if (!start) return "";
  const format = (value) => {
    const [year, month, day] = value.split("-");
    return `${day}/${month}/${year}`;
  };
  return start === end ? format(start) : `${format(start)} - ${format(end)}`;
}

const REQUIRED_ORGANIZATIONS = ["[BB]", "[BK]", "[NTS]", "[PK]", "[SM]"];

function buildDisplayRows(sourceRows) {
  const runOrder = [];
  const runs = new Map();
  for (const row of sourceRows) {
    if (!runs.has(row.run_id)) {
      runs.set(row.run_id, []);
      runOrder.push(row.run_id);
    }
    runs.get(row.run_id).push(row);
  }

  const displayRows = [];
  for (const runId of runOrder) {
    const runRows = runs.get(runId);
    displayRows.push(...runRows.map((row) => ({ ...row, is_summary: false })));
    const successful = new Set(runRows.filter((row) => row.status === "สำเร็จ").map((row) => row.organization));
    if (!REQUIRED_ORGANIZATIONS.every((org) => successful.has(org))) continue;

    const timedRows = runRows.filter((row) =>
      !Number.isNaN(new Date(row.started_at).valueOf()) && !Number.isNaN(new Date(row.finished_at).valueOf()),
    );
    if (!timedRows.length) continue;
    const firstRow = timedRows.reduce((first, row) =>
      new Date(row.started_at) < new Date(first.started_at) ? row : first,
    );
    const lastRow = timedRows.reduce((last, row) =>
      new Date(row.finished_at) > new Date(last.finished_at) ? row : last,
    );
    const startedAt = new Date(firstRow.started_at);
    const finishedAt = new Date(lastRow.finished_at);
    displayRows.push({
      run_id: runId,
      organization: "สรุปรอบ",
      target_date: runRows[0]?.target_date || null,
      target_start_date: runRows[0]?.target_start_date || runRows[0]?.target_date || null,
      target_end_date: runRows[0]?.target_end_date || runRows[0]?.target_date || null,
      started_at: firstRow.started_at,
      finished_at: lastRow.finished_at,
      duration_seconds: (finishedAt - startedAt) / 1000,
      status: "สำเร็จครบ 5 สาขา",
      message: "รวมตั้งแต่เริ่ม [BB] ถึง Finish [SM]",
      is_summary: true,
    });
  }
  return displayRows;
}

const displayRows = buildDisplayRows(rows);
const workbook = Workbook.create();
const sheet = workbook.worksheets.add("เวลาประมวลผล");
sheet.showGridLines = false;
sheet.freezePanes.freezeRows(4);
sheet.tabColor = "#1F4E78";

sheet.getRange("A2:H2").merge();
sheet.getRange("A2").values = [["ระยะเวลาประมวลผล HRMI รายสาขา"]];
sheet.getRange("A2").format = { font: { name: "Arial", size: 14, bold: true, color: "#1F1F1F" }, verticalAlignment: "center" };
sheet.getRange("A3:H3").format.borders = { bottom: { style: "thin", color: "#7F8C8D" } };

sheet.getRange("A4:H4").values = [["รอบการทำงาน", "สาขา", "ช่วงวันที่ประมวลผล", "เวลาเริ่ม", "เวลาสิ้นสุด", "ระยะเวลา", "สถานะ", "รายละเอียด"]];
sheet.getRange("A4:H4").format = {
  fill: "#1F4E78",
  font: { name: "Arial", size: 10, bold: true, color: "#FFFFFF" },
  horizontalAlignment: "center", verticalAlignment: "center",
  borders: { preset: "inside", style: "thin", color: "#FFFFFF" },
};

if (displayRows.length) {
  const values = displayRows.map((row) => [
    row.run_id, row.organization,
    processingRange(row),
    excelBangkokDateTime(row.started_at), excelBangkokDateTime(row.finished_at),
    Number(row.duration_seconds || 0) / 86400, row.status, row.message || "",
  ]);
  const last = 4 + values.length;
  sheet.getRange(`A5:H${last}`).values = values;
  sheet.getRange(`A5:H${last}`).format = { font: { name: "Arial", size: 10, color: "#1F1F1F" }, verticalAlignment: "center" };
  sheet.getRange(`D5:E${last}`).format.numberFormat = "dd/mm/yyyy hh:mm:ss";
  sheet.getRange(`F5:F${last}`).format.numberFormat = "[h]:mm:ss";
  sheet.getRange(`G5:G${last}`).conditionalFormats.add("containsText", { text: "สำเร็จ", format: { fill: "#E2F0D9", font: { color: "#375623", bold: true } } });
  sheet.getRange(`G5:G${last}`).conditionalFormats.add("containsText", { text: "ผิดพลาด", format: { fill: "#FCE4D6", font: { color: "#C00000", bold: true } } });

  for (let index = 0; index < displayRows.length; index += 1) {
    const rowRange = sheet.getRange(`A${index + 5}:H${index + 5}`);
    if (displayRows[index].is_summary) {
      rowRange.format = {
        fill: "#D9EAF7",
        font: { name: "Arial", size: 10, bold: true, color: "#17365D" },
        verticalAlignment: "center",
        borders: { top: { style: "medium", color: "#1F4E78" }, bottom: { style: "medium", color: "#1F4E78" } },
      };
    } else if (index % 2 === 1) rowRange.format.fill = "#F2F7FB";
  }
}

sheet.getRange("A:H").format.font = { name: "Arial", size: 10 };
sheet.getRange("A:A").format.columnWidth = 23;
sheet.getRange("B:B").format.columnWidth = 12;
sheet.getRange("C:C").format.columnWidth = 26;
sheet.getRange("D:E").format.columnWidth = 23;
sheet.getRange("F:F").format.columnWidth = 14;
sheet.getRange("G:G").format.columnWidth = 20;
sheet.getRange("H:H").format.columnWidth = 48;
sheet.getRange("H:H").format.wrapText = false;

workbook.recalculate();
await fs.mkdir(path.dirname(outputPath), { recursive: true });
const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);

const lastDisplayRow = Math.max(5, 4 + displayRows.length);
const check = await workbook.inspect({ kind: "table", range: `เวลาประมวลผล!A2:H${lastDisplayRow}`, include: "values,formulas", tableMaxRows: 12, tableMaxCols: 8 });
const errors = await workbook.inspect({ kind: "match", searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!", options: { useRegex: true, maxResults: 100 }, summary: "final formula error scan" });
process.stdout.write(`${check.ndjson}\n${errors.ndjson}`);

if (mode === "--preview") {
  const preview = await workbook.render({ sheetName: "เวลาประมวลผล", range: `A1:H${Math.max(7, 4 + displayRows.length)}`, scale: 1.5, format: "png" });
  await fs.writeFile(path.join(path.dirname(jsonPath), "HRMI-processing-times-preview.png"), new Uint8Array(await preview.arrayBuffer()));
}
