// Dates and Protection (one-off): makes the Start/End date picker easier to find and
// protects the sheets against accidental edits. Run once in Excel for the web, then
// delete the script. Safe to run again.
//  - Template and project tabs: Start/End (E7:F down to the "Insert new rows ABOVE this one"
//    row) only take dates, and selecting one shows "Double-click for a calendar".
//    The range includes the marker row so rows inserted above it inherit the rule.
//    Headers E6/F6 get a calendar icon.
//  - Protection (no password; it guards against accidents, not people):
//      project tabs + Template: everything locked except B1 name, B3 manager, E2 start,
//        E4 Display Week, row 3 chart labels and the task rows (whole rows 7 down to just
//        above the marker, so they can be deleted). Rows can be inserted/deleted, cells/rows/columns formatted; columns can't
//        be inserted or deleted.
//      Projects: everything locked except Client (C), Tab (J) and Client Email (K).
//      Instructions: fully locked.
//    Snapshot Requests stays unprotected: the email flow writes to it.
//  - Instructions tab: calendar hint, blank-means-today start date, what's locked.

/** One-off: date validation and calendar hint on Start/End, sheet protection, Instructions wording. */
function main(workbook: ExcelScript.Workbook) {
  const projectTabs = workbook.getWorksheets().filter(ws => String(ws.getRange("B6").getValue()) === "TASK");
  const list = workbook.getWorksheet("Projects");
  const ins = workbook.getWorksheet("Instructions");
  for (const ws of [...projectTabs, list, ins]) {
    if (ws && ws.getProtection().getProtected()) ws.getProtection().unprotect();
  }

  for (const ws of projectTabs) {
    // Range.find doesn't search hidden tabs (the Template), so scan column B instead.
    const marker = ws.getRange("B1:B2000").getValues().findIndex(r => String(r[0]).toLowerCase() === END_MARKER.toLowerCase()) + 1;
    if (marker < 1) { console.log(`${ws.getName()}: no end marker, skipped`); continue; }

    const dv = ws.getRange(`E7:F${marker}`).getDataValidation();
    dv.clear();
    dv.setRule({ date: { formula1: "2000-01-01", formula2: "2099-12-31", operator: ExcelScript.DataValidationOperator.between } });
    dv.setPrompt({ showPrompt: true, title: "Pick a date", message: "Double-click for a calendar, or type a date like 2026-11-02." });
    dv.setErrorAlert({ showAlert: true, style: ExcelScript.DataValidationAlertStyle.stop, title: "Not a date",
      message: "Enter a date like 2026-11-02, or double-click the cell for a calendar." });
    dv.setIgnoreBlanks(true);

    if (String(ws.getRange("E6").getValue()) === "START") ws.getRange("E6").setValue(`START ${CAL}`);
    if (String(ws.getRange("F6").getValue()) === "END") ws.getRange("F6").setValue(`END ${CAL}`);

    // Whole task rows stay unlocked: a protected sheet won't delete a row that has any locked cell.
    lock(ws, ["B1", "B3", "E2", "E4", "I3:ED3", `7:${marker - 1}`]);
    ws.getProtection().protect({ ...EDITABLE, allowInsertRows: true, allowDeleteRows: true });
    console.log(`${ws.getName()}: dates E7:F${marker}, protected`);
  }

  if (list) {
    lock(list, ["C6:C2000", "J6:J2000", "K6:K2000"]);
    list.getProtection().protect({ ...EDITABLE, allowInsertRows: false, allowDeleteRows: false });
    console.log("Projects: protected");
  }

  if (!ins) return;
  const fixes: [string, string][] = [
    ["2.  Enter the project name and start date (for example 2026-11-02); client, project manager and client email are optional.",
      "2.  Enter the project name. Start date (for example 2026-11-02), client, project manager and client email are optional; a blank start date means today."],
  ];
  const col = () => ins.getRange("C1:C80").getValues().map(r => String(r[0]));
  col().forEach((text, i) => {
    for (const [from, to] of fixes) {
      if (text.indexOf(from) >= 0) {
        ins.getRange(`C${i + 1}`).setValue(text.replace(from, to));
        ins.getRange(`C${i + 1}`).getEntireRow().getFormat().autofitRows();
        console.log(`Row ${i + 1} reworded`);
      }
    }
  });
  addBullet(ins, col(), "Type task names", "for a calendar",
    `Start and End (${CAL}): double-click a date cell for a calendar, or type a date like 2026-11-02. Selecting the cell shows the same hint.`);
  addBullet(ins, col(), "Everyone should work in this one shared file", "are locked",
    "Headings, formulas and the Projects tab's calculated columns are locked so they can't be broken by accident. " +
    "You can type in the task rows, the project name, manager, start date and Display Week, and in Client, Tab and Client Email on the Projects tab. " +
    "The buttons unlock what they need by themselves.");
  ins.getRange("A:XFD").getFormat().getProtection().setLocked(true);
  ins.getProtection().protect({ allowFormatCells: false, allowFormatColumns: false, allowFormatRows: false, allowInsertRows: false,
    allowDeleteRows: false, allowInsertColumns: false, allowDeleteColumns: false, allowInsertHyperlinks: false, allowSort: false,
    allowAutoFilter: false, allowPivotTables: false, allowEditObjects: false, allowEditScenarios: false });
  console.log("Instructions: protected");
}

const END_MARKER = "Insert new rows ABOVE this one";
const CAL = "\u{1F4C5}";
const EDITABLE: ExcelScript.WorksheetProtectionOptions = {
  allowFormatCells: true, allowFormatColumns: true, allowFormatRows: true,
  allowInsertColumns: false, allowDeleteColumns: false, allowInsertHyperlinks: false,
  // Script buttons don't respond on a protected sheet unless objects are editable.
  allowSort: false, allowAutoFilter: false, allowPivotTables: false, allowEditObjects: true, allowEditScenarios: false,
};

// Locks the whole sheet, then unlocks the given ranges.
function lock(ws: ExcelScript.Worksheet, editable: string[]) {
  ws.getRange("A:XFD").getFormat().getProtection().setLocked(true);
  for (const a of editable) ws.getRange(a).getFormat().getProtection().setLocked(false);
}

// Adds a bullet (formatted like the row it follows) below the row starting with `after`,
// unless a row already contains `marker`.
function addBullet(ins: ExcelScript.Worksheet, col: string[], after: string, marker: string, text: string) {
  if (col.some(t => t.indexOf(marker) >= 0)) return;
  const row = col.findIndex(t => t.startsWith(after)) + 1;
  if (row < 1) return;
  const at = row + 1;
  ins.getRange(`${at}:${at}`).insert(ExcelScript.InsertShiftDirection.down);
  ins.getRange(`B${at}:C${at}`).copyFrom(ins.getRange(`B${row}:C${row}`), ExcelScript.RangeCopyType.formats);
  ins.getRange(`B${at}`).setValue(String(ins.getRange(`B${row}`).getValue()));
  ins.getRange(`C${at}`).setValue(text);
  ins.getRange(`${at}:${at}`).getFormat().autofitRows();
  console.log(`Bullet added at row ${at}`);
}
