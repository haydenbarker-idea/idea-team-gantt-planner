// Add Section: adds a section header plus one task, after the section the selection
// is in, or at the end of the list when nothing in the list is selected.
//
// Works on the active project tab (any sheet with "TASK" in B6). A section header row is bold in column B with
// no start date (column E); a task row has a name or a start date. The list runs from
// row 8 down to the "Insert new rows ABOVE this one" row.

/**
 * Adds a section with one task after the selected section, or at the end of the list.
 * @param sectionName Example: Rough In
 */
function main(workbook: ExcelScript.Workbook, sectionName: string) {
  const name = (sectionName || "").trim();
  if (!name) throw new Error("Enter a name for the new section.");
  const sheet = projectSheet(workbook);
  const list = readList(sheet);

  const h = sectionHeader(list, workbook.getActiveCell().getRowIndex() + 1);
  const at = Math.max(FIRST_ROW, h ? sectionEnd(list, h) + 1 : lastContentRow(list) + 1);

  // Copy the look of this tab's sections; a tab with none left borrows from the Template tab.
  let src = sheet;
  let srcList = list;
  if (!findRow(list, true) || !findRow(list, false)) {
    const tmpl = workbook.getWorksheet(TEMPLATE_SHEET);
    if (tmpl) { src = tmpl; srcList = readList(tmpl); }
  }
  let headerTmpl = findRow(srcList, true);
  let taskTmpl = findRow(srcList, false);
  if (!headerTmpl || !taskTmpl) throw new Error("Couldn't find an existing section and task to copy the formatting from.");

  const idx = pickColor(list, at);
  const start = nextStartDate(list, at, sheet);

  unprotected(sheet, () => {
    sheet.getRange(`${at}:${at + 1}`).insert(ExcelScript.InsertShiftDirection.down);
    if (src === sheet && headerTmpl >= at) headerTmpl += 2;
    if (src === sheet && taskTmpl >= at) taskTmpl += 2;

    copyLook(src, headerTmpl, sheet, at);
    recolor(sheet, at, HEADER_COLORS[idx]);
    sheet.getRange(`B${at}`).setValue(name);

    copyLook(src, taskTmpl, sheet, at + 1);
    recolor(sheet, at + 1, TASK_COLORS[idx]);
    fillTask(sheet, at + 1, "New Task", start);
  });
  sheet.getRange(`B${at + 1}`).select();
}

// ------------------------------------------------------------------ shared with Add Task.ts

const TEMPLATE_SHEET = "Template";
const FIRST_ROW = 8;
const END_MARKER = "Insert new rows ABOVE this one";
const DAYS_FORMULA = '=IF(OR(ISBLANK(task_start),ISBLANK(task_end)),"",task_end-task_start+1)';
// Header / task fill colours, cycled for new sections (same as the existing sections).
const HEADER_COLORS = ["#B8CCE4", "#D8E4BC", "#E6B8B7", "#CCC0DA", "#B7DEE8"];
const TASK_COLORS = ["#DCE6F1", "#EBF1DE", "#F2DCDB", "#E4DFEC", "#DAEEF3"];

// The active sheet, if it's a project tab.
function projectSheet(workbook: ExcelScript.Workbook): ExcelScript.Worksheet {
  const sheet = workbook.getActiveWorksheet();
  if (String(sheet.getRange("B6").getValue()) !== "TASK") throw new Error("Open a project tab first.");
  return sheet;
}

interface Row { name: string; hasStart: boolean; end: number; bold: boolean; fill: string; }
interface List { marker: number; rows: { [row: number]: Row }; }

// Snapshot of the task list, read once to keep the script fast.
function readList(sheet: ExcelScript.Worksheet): List {
  const found = sheet.getRange("B:B").find(END_MARKER, { completeMatch: true, matchCase: false });
  const marker = found ? found.getRowIndex() + 1 : (sheet.getUsedRange()?.getLastRow().getRowIndex() ?? FIRST_ROW) + 2;
  const rows: { [row: number]: Row } = {};
  if (marker <= FIRST_ROW) return { marker, rows };
  const values = sheet.getRange(`B${FIRST_ROW}:F${marker - 1}`).getValues();
  for (let i = 0; i < values.length; i++) {
    const r = FIRST_ROW + i;
    const b = sheet.getRange(`B${r}`).getFormat();
    rows[r] = {
      name: String(values[i][0] ?? ""),
      hasStart: values[i][3] !== "" && values[i][3] !== null,
      end: typeof values[i][4] === "number" ? (values[i][4] as number) : 0,
      bold: b.getFont().getBold() === true,
      fill: (b.getFill().getColor() || "").toUpperCase(),
    };
  }
  return { marker, rows };
}

function isHeader(list: List, r: number): boolean {
  const row = list.rows[r];
  return !!row && row.bold && row.name.length > 0 && !row.hasStart;
}

function isTask(list: List, r: number): boolean {
  const row = list.rows[r];
  return !!row && !isHeader(list, r) && (row.name.length > 0 || row.hasStart);
}

// Header row of the section row r belongs to, or 0.
function sectionHeader(list: List, r: number): number {
  if (r < FIRST_ROW || r >= list.marker) return 0;
  for (; r >= FIRST_ROW; r--) {
    if (isHeader(list, r)) return r;
    if (!isTask(list, r)) return 0;
  }
  return 0;
}

// Last row of the section whose header is on row h (h itself if it has no tasks).
function sectionEnd(list: List, h: number): number {
  let r = h + 1;
  while (r < list.marker && isTask(list, r)) r++;
  return r - 1;
}

function lastContentRow(list: List): number {
  for (let r = list.marker - 1; r >= FIRST_ROW; r--) {
    if (isHeader(list, r) || isTask(list, r)) return r;
  }
  return FIRST_ROW - 1;
}

// First header (wantHeader) or first task row, used as a formatting template.
function findRow(list: List, wantHeader: boolean): number {
  for (let r = FIRST_ROW; r < list.marker; r++) {
    if (wantHeader ? isHeader(list, r) : isTask(list, r)) return r;
  }
  return 0;
}

// Palette index for a section inserted at row `at`: the next colour in the cycle,
// skipped forward if it matches the section above or below.
function pickColor(list: List, at: number): number {
  let prevFill = "";
  let nextFill = "";
  let n = 0;
  for (let r = FIRST_ROW; r < at; r++) {
    if (isHeader(list, r)) { n++; prevFill = list.rows[r].fill; }
  }
  for (let r = at; r < list.marker; r++) {
    if (isHeader(list, r)) { nextFill = list.rows[r].fill; break; }
  }
  let idx = n % HEADER_COLORS.length;
  for (let i = 0; i < HEADER_COLORS.length; i++) {
    if (HEADER_COLORS[idx] !== prevFill && HEADER_COLORS[idx] !== nextFill) break;
    idx = (idx + 1) % HEADER_COLORS.length;
  }
  return idx;
}

// Day after the task just above row `at`; failing that the latest end date above; failing that the project start.
function nextStartDate(list: List, at: number, sheet: ExcelScript.Worksheet): number {
  if (isTask(list, at - 1) && list.rows[at - 1].end > 0) return list.rows[at - 1].end + 1;
  let best = 0;
  for (let r = FIRST_ROW; r < at; r++) {
    if (isTask(list, r)) best = Math.max(best, list.rows[r].end);
  }
  return best > 0 ? best + 1 : (sheet.getRange("E2").getValue() as number);
}

// Copies the look of columns A:H from one row to another. Column D is set by hand so the
// progress data bar rule isn't duplicated (the row insert already extended it).
function copyLook(from: ExcelScript.Worksheet, fromRow: number, sheet: ExcelScript.Worksheet, toRow: number) {
  sheet.getRange(`A${toRow}:C${toRow}`).copyFrom(from.getRange(`A${fromRow}:C${fromRow}`), ExcelScript.RangeCopyType.formats);
  sheet.getRange(`E${toRow}:H${toRow}`).copyFrom(from.getRange(`E${fromRow}:H${fromRow}`), ExcelScript.RangeCopyType.formats);
  const src = from.getRange(`D${fromRow}`);
  const dst = sheet.getRange(`D${toRow}`);
  dst.getFormat().getFill().setColor(src.getFormat().getFill().getColor());
  dst.getFormat().getFont().setBold(src.getFormat().getFont().getBold());
  dst.getFormat().getFont().setColor(src.getFormat().getFont().getColor());
  dst.getFormat().setHorizontalAlignment(src.getFormat().getHorizontalAlignment());
  dst.setNumberFormat(src.getNumberFormat());
  sheet.getRange(`${toRow}:${toRow}`).getFormat().setRowHeight(from.getRange(`${fromRow}:${fromRow}`).getFormat().getRowHeight());
  sheet.getRange(`H${toRow}`).setFormula(DAYS_FORMULA);
}

// Re-fills the coloured cells of columns B:H.
function recolor(sheet: ExcelScript.Worksheet, r: number, color: string) {
  for (const col of ["B", "C", "D", "E", "F", "G", "H"]) {
    const fill = sheet.getRange(`${col}${r}`).getFormat().getFill();
    const current = (fill.getColor() || "").toUpperCase();
    if (current && current !== "#FFFFFF") fill.setColor(color);
  }
}

function fillTask(sheet: ExcelScript.Worksheet, r: number, name: string, start: number) {
  sheet.getRange(`B${r}`).setValue(name);
  sheet.getRange(`D${r}`).setValue(0);
  sheet.getRange(`E${r}:F${r}`).setValues([[start, start]]);
  sheet.getRange(`H${r}`).setFormula(DAYS_FORMULA);
}

// The sheets are protected so headings and formulas can't be overwritten by accident.
// Scripts lift the protection while they change a sheet, then put the same protection back.
function unprotected<T>(ws: ExcelScript.Worksheet, work: () => T): T {
  const p = ws.getProtection();
  if (!p.getProtected()) return work();
  const options = p.getOptions();
  p.unprotect();
  try { return work(); } finally { p.protect(options); }
}
