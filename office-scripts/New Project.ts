// New Project: copies the hidden Template tab into a new project tab and adds it to
// the Projects list.
//
// The Projects tab lists projects from row 6 down; its columns are
//   B Project (link)  C Client  D Project manager  E Start  F Finish
//   G Progress  H Status  I Archived ("Yes" or blank)  J Tab name  K Client email
// Everything but Client, Archived, Tab name and Client email is a formula reading the project's tab;
// its ranges start on the blank row 7 so rows added at the top of the task list are included.

// The @param lines are the hints shown under each field when the button runs. The email
// example uses a full-width at sign: a plain one would be read as the start of a new tag.
/**
 * Creates a new project tab from the Template and adds it to the Projects list.
 * @param projectName Example: Atlantic Packaging - Line 3
 * @param startDate Example: 2026-11-02 (or 11/2/2026). Leave blank for today.
 * @param client Example: Atlantic Packaging (optional)
 * @param projectManager Example: Vinny Singh (optional)
 * @param clientEmail Example: jane.doe＠example.com (optional, used by Email Snapshot to Client)
 */
function main(workbook: ExcelScript.Workbook, projectName: string, startDate?: string, client?: string, projectManager?: string, clientEmail?: string) {
  const name = (projectName || "").trim();
  if (!name) throw new Error("Enter a project name.");
  let start = parseDate(startDate);
  if (start === null && (startDate || "").trim()) throw new Error(`'${startDate}' isn't a date. Use a date like 2026-11-02, or leave it blank for today.`);

  const tmpl = workbook.getWorksheet(TEMPLATE_SHEET);
  const list = workbook.getWorksheet(PROJECTS_SHEET);
  if (!tmpl || !list) throw new Error(`This workbook needs a ${TEMPLATE_SHEET} tab and a ${PROJECTS_SHEET} tab.`);

  const wasVisible = tmpl.getVisibility();
  tmpl.setVisibility(ExcelScript.SheetVisibility.visible);
  const ws = tmpl.copy(ExcelScript.WorksheetPositionType.before, tmpl);
  tmpl.setVisibility(wasVisible);
  ws.setName(uniqueSheetName(workbook, name));
  // Blank start date: today, from the tab's own =TODAY() cell (the workbook's clock, not the script's).
  if (start === null) start = Math.floor(ws.getRange("E3").getValue() as number);

  ws.getRange("B1").setValue(name);
  ws.getRange("B3").setValue((projectManager || "").trim());
  ws.getRange("E2").setValue(start);
  ws.getRange("E4").setValue(1);

  // The template's tasks all start on the project start date.
  const found = ws.getRange("B:B").find(END_MARKER, { completeMatch: true, matchCase: false });
  const marker = found ? found.getRowIndex() + 1 : 60;
  const starts = ws.getRange(`E8:E${marker - 1}`).getValues();
  for (let i = 0; i < starts.length; i++) {
    if (starts[i][0] !== "" && starts[i][0] !== null) ws.getRange(`E${8 + i}:F${8 + i}`).setValues([[start, start]]);
  }

  unprotected(list, () => addToList(list, ws.getName(), (client || "").trim(), (clientEmail || "").trim()));
  ws.activate();
  ws.getRange("B8").select();
}

const PROJECTS_SHEET = "Projects";
const TEMPLATE_SHEET = "Template";
const LIST_FIRST_ROW = 6;
const END_MARKER = "Insert new rows ABOVE this one";

// yyyy-mm-dd or m/d/yyyy, as an Excel date serial number.
function parseDate(text?: string): number | null {
  const t = (text || "").trim();
  let m = t.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  let y = 0, mo = 0, d = 0;
  if (m) { y = +m[1]; mo = +m[2]; d = +m[3]; }
  else {
    m = t.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (!m) return null;
    mo = +m[1]; d = +m[2]; y = +m[3];
  }
  const ms = Date.UTC(y, mo - 1, d);
  const check = new Date(ms);
  if (check.getUTCMonth() !== mo - 1 || check.getUTCDate() !== d) return null;
  return Math.round((ms - Date.UTC(1899, 11, 30)) / 86400000);
}

// Tab names: max 31 characters, none of : \ / ? * [ ], and unique in the workbook.
function uniqueSheetName(workbook: ExcelScript.Workbook, projectName: string): string {
  const taken = workbook.getWorksheets().map(s => s.getName().toLowerCase());
  const base = (projectName.replace(/[:\\\/?*\[\]]/g, " ").replace(/\s+/g, " ").trim() || "Project").slice(0, 31).trim();
  let name = base;
  for (let n = 2; taken.indexOf(name.toLowerCase()) >= 0; n++) {
    const suffix = ` (${n})`;
    name = base.slice(0, 31 - suffix.length) + suffix;
  }
  return name;
}

function addToList(list: ExcelScript.Worksheet, tabName: string, client: string, clientEmail: string) {
  const tabs = list.getRange(`J${LIST_FIRST_ROW}:J2000`).getValues();
  let r = LIST_FIRST_ROW;
  for (let i = 0; i < tabs.length; i++) if (tabs[i][0] !== "") r = LIST_FIRST_ROW + i + 1;
  if (r > LIST_FIRST_ROW) list.getRange(`${r}:${r}`).copyFrom(`${r - 1}:${r - 1}`, ExcelScript.RangeCopyType.formats);

  const q = `'${tabName.replace(/'/g, "''")}'!`;
  list.getRange(`J${r}`).setValue(tabName);
  list.getRange(`B${r}`).setFormula(`=HYPERLINK("#'"&SUBSTITUTE(J${r},"'","''")&"'!A1",${q}$B$1)`);
  list.getRange(`C${r}`).setValue(client);
  list.getRange(`K${r}`).setValue(clientEmail);
  list.getRange(`D${r}`).setFormula(`=${q}$B$3&""`);
  list.getRange(`E${r}`).setFormula(`=IF(COUNT(${q}$E$7:$E$1000)=0,"",MIN(${q}$E$7:$E$1000))`);
  list.getRange(`F${r}`).setFormula(`=IF(COUNT(${q}$F$7:$F$1000)=0,"",MAX(${q}$F$7:$F$1000))`);
  list.getRange(`G${r}`).setFormula(`=IFERROR(SUMPRODUCT(${q}$D$7:$D$1000,${q}$H$7:$H$1000)/SUM(${q}$H$7:$H$1000),"")`);
  list.getRange(`H${r}`).setFormula(`=IF(I${r}="Yes","Archived",IF(E${r}="","Not scheduled",IF(N(G${r})>=1,"Complete",IF(TODAY()<E${r},"Not started",IF(TODAY()>F${r},"Overdue","In progress")))))`);
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
