// Prepare Snapshot: run by the "Project Schedules - email snapshot" flow on a TEMPORARY COPY
// of the workbook, never on the real one. It deletes every tab except the requested project
// and turns that tab into the customer view (same as the desktop SnapshotTools macro):
// values only, no buttons, links, invoice markers or Display Week, the chart cropped to the
// project's weeks, and a page setup that prints one page wide. The flow then converts the
// copy to PDF and attaches both files to the email.

interface Snapshot {
  projectName: string;
  client: string;
  clientEmail: string;
  fileBase: string;
  asOf: string;
}

/** Used by the email-snapshot flow on a temporary copy: keeps one project tab and makes it the customer view. */
function main(workbook: ExcelScript.Workbook, tabName: string): Snapshot {
  // It deletes tabs, so only ever on the flow's temporary copies.
  if (!workbook.getName().toLowerCase().startsWith(TEMP_PREFIX)) {
    throw new Error(`Prepare Snapshot only runs on the flow's temporary copies (files named ${TEMP_PREFIX}...).`);
  }
  const src = workbook.getWorksheet(tabName);
  if (!src) throw new Error(`No tab called '${tabName}'.`);

  // Client details from the Projects list (C client, J tab, K client email) before it goes.
  let client = "";
  let clientEmail = "";
  const list = workbook.getWorksheet(PROJECTS_SHEET);
  if (list) {
    const rows = list.getRange("B6:K2000").getValues();
    for (const row of rows) {
      if (String(row[8]).toLowerCase() === tabName.toLowerCase()) {
        client = String(row[1]);
        clientEmail = String(row[9]).trim();
        break;
      }
    }
  }

  const projectName = String(src.getRange("B1").getValue());
  const asOfSerial = src.getRange("E3").getValue() as number;
  const asOf = formatDate(asOfSerial);

  for (const ws of workbook.getWorksheets()) {
    if (ws.getName() !== tabName) ws.delete();
  }
  src.setName("Schedule");
  // The live tabs are protected against accidental edits; this copy is reworked, then frozen as values.
  if (src.getProtection().getProtected()) src.getProtection().unprotect();
  makeCustomerView(workbook, src, asOf);

  return { projectName, client, clientEmail, fileBase: `${safeFileName(projectName)} - Schedule ${isoDate(asOfSerial)}`, asOf };
}

const TEMP_PREFIX = "snapshot-tmp-";
const PROJECTS_SHEET = "Projects";
const FIRST_ROW = 8;
const END_MARKER = "Insert new rows ABOVE this one";
const CHART_FIRST_COL = 9;   // column I
const CHART_LAST_COL = 134;  // column ED

function makeCustomerView(workbook: ExcelScript.Workbook, ws: ExcelScript.Worksheet, asOf: string) {
  ws.getShapes().forEach(s => s.delete());
  ws.getComments().forEach(c => c.delete());
  ws.getRange("I1:N1").clear(ExcelScript.ClearApplyTo.contents);   // "All projects" link
  ws.getRange("I3:ED3").clear(ExcelScript.ClearApplyTo.contents);  // invoice markers (internal)
  ws.getRange("E4").setValue(1);                                    // chart starts at the first week
  workbook.getApplication().calculate(ExcelScript.CalculationType.full);

  const used = ws.getUsedRange();
  used.copyFrom(used, ExcelScript.RangeCopyType.values);            // freeze as of today, no formulas
  ws.getRange("D3").setValue("As of:");
  ws.getRange("D4:E4").clear(ExcelScript.ClearApplyTo.all);        // Display Week is internal
  ws.getRange("I5:ED5").getFormat().getFill().setColor(ws.getRange("I4").getFormat().getFill().getColor());

  // Last task row and last end date.
  const found = ws.getRange("B:B").find(END_MARKER, { completeMatch: true, matchCase: false });
  const marker = found ? found.getRowIndex() + 1 : used.getRowIndex() + used.getRowCount() + 1;
  const values = ws.getRange(`B${FIRST_ROW}:F${marker - 1}`).getValues();
  let lastRow = FIRST_ROW;
  let lastEnd = 0;
  values.forEach((row, i) => {
    if (String(row[0]) !== "") lastRow = FIRST_ROW + i;
    if (typeof row[4] === "number" && row[4] > lastEnd) lastEnd = row[4] as number;
  });

  // Whole weeks up to the last end date, within the chart's 18 weeks.
  const startDay = ws.getCell(4, CHART_FIRST_COL - 1).getValue() as number;
  const nDays = lastEnd >= startDay ? Math.ceil((lastEnd - startDay + 1) / 7) * 7 : 7;
  const lastCol = Math.min(CHART_FIRST_COL + nDays - 1, CHART_LAST_COL);
  const lastColLetter = columnLetter(lastCol);

  if (lastCol < CHART_LAST_COL) {
    ws.getRange(`${columnLetter(lastCol + 1)}:ED`).delete(ExcelScript.DeleteShiftDirection.left);
  }
  const usedNow = ws.getUsedRange();
  const usedLastRow = usedNow.getRowIndex() + usedNow.getRowCount();
  if (usedLastRow > lastRow) ws.getRange(`${lastRow + 1}:${usedLastRow}`).delete(ExcelScript.DeleteShiftDirection.up);

  // Shorter dates so the task columns take less of the page.
  ws.getRange(`E${FIRST_ROW}:F${lastRow}`).setNumberFormat("ddd mmm d, yyyy");
  ws.getRange("E:F").getFormat().setColumnWidth(95);

  let width = 0;
  for (const col of ["A", "B", "C", "D", "E", "F", "G", "H"]) width += ws.getRange(`${col}1`).getFormat().getColumnWidth();
  width += 13.5 * (lastCol - CHART_FIRST_COL + 1);

  const page = ws.getPageLayout();
  page.setPrintArea(`A1:${lastColLetter}${lastRow}`);
  page.setPrintTitleRows("$4:$6");
  page.setOrientation(ExcelScript.PageOrientation.landscape);
  page.setPaperSize(width > 1100 ? ExcelScript.PaperType.tabloid : ExcelScript.PaperType.letter);
  page.setZoom({ horizontalFitToPages: 1, verticalFitToPages: 0 });
  page.setCenterHorizontally(true);
  page.setLeftMargin(28.8);
  page.setRightMargin(28.8);
  page.setTopMargin(36);
  page.setBottomMargin(43.2);
  page.setPrintGridlines(false);
  const footer = page.getHeadersFooters().getDefaultForAllPages();
  footer.setLeftFooter("Idea Networks");
  footer.setCenterFooter("Page &P of &N");
  footer.setRightFooter(`Schedule as of ${asOf}`);
  ws.setShowGridlines(false);
}

function columnLetter(n: number): string {
  let s = "";
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function serialToDate(serial: number): Date {
  return new Date(Date.UTC(1899, 11, 30) + Math.round(serial) * 86400000);
}

function formatDate(serial: number): string {
  const d = serialToDate(serial);
  const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  return `${months[d.getUTCMonth()]} ${d.getUTCDate()}, ${d.getUTCFullYear()}`;
}

function isoDate(serial: number): string {
  return serialToDate(serial).toISOString().slice(0, 10);
}

function safeFileName(s: string): string {
  return s.replace(/[\\\/:*?"<>|]/g, " ").replace(/\s+/g, " ").trim() || "Project";
}
