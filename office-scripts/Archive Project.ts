// Archive / Restore: archives the selected project (on the Projects tab) or the project
// tab you're on: marks it Archived in the list and hides its tab. Run it on an archived
// project in the list to restore it.
//
// Projects tab columns: I Archived ("Yes" or blank), J Tab name; the list starts on row 6.

/** Archives the selected project (hides its tab), or restores it if it's already archived. */
function main(workbook: ExcelScript.Workbook) {
  const list = workbook.getWorksheet(PROJECTS_SHEET);
  if (!list) throw new Error(`This workbook needs a ${PROJECTS_SHEET} tab.`);

  const active = workbook.getActiveWorksheet();
  let r = 0;
  if (active.getName() === list.getName()) {
    r = workbook.getActiveCell().getRowIndex() + 1;
    if (r < LIST_FIRST_ROW || String(list.getRange(`J${r}`).getValue()) === "") r = 0;
  } else {
    const tabs = list.getRange(`J${LIST_FIRST_ROW}:J2000`).getValues();
    for (let i = 0; i < tabs.length; i++) {
      if (String(tabs[i][0]).toLowerCase() === active.getName().toLowerCase()) { r = LIST_FIRST_ROW + i; break; }
    }
  }
  if (!r) throw new Error("Select a project in the Projects list first.");

  const tabName = String(list.getRange(`J${r}`).getValue());
  const ws = workbook.getWorksheet(tabName);
  if (!ws) throw new Error(`Couldn't find the tab '${tabName}'. If it was renamed, put the new name in the Tab column.`);

  if (String(list.getRange(`I${r}`).getValue()) !== "Yes") {
    list.activate();
    list.getRange(`I${r}`).setValue("Yes");
    ws.setVisibility(ExcelScript.SheetVisibility.hidden);
  } else {
    list.getRange(`I${r}`).setValue("");
    ws.setVisibility(ExcelScript.SheetVisibility.visible);
    ws.activate();
  }
}

const PROJECTS_SHEET = "Projects";
const LIST_FIRST_ROW = 6;
