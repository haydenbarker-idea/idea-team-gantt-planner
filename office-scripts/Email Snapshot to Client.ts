// Email Snapshot to Client: queues a snapshot request for the project tab you're on. The
// Power Automate flow "Project Schedules - email snapshot" checks for requests every couple
// of minutes, makes a PDF and an Excel copy of just this project (see Prepare Snapshot.ts)
// and emails both to whoever clicked, to forward to the client.
//
// Requests go into the SnapshotRequests table on the hidden "Snapshot Requests" tab:
//   Id | Requested (UTC) | Tab | Requester | Status | Note
// The flow finds and updates rows by Id.
// The requester's email is taken from a comment the script adds and removes, since
// Office Scripts has no "current user" API.

/** Emails you a PDF and Excel snapshot of this project within a few minutes, to forward to the client. */
function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();
  if (String(sheet.getRange("B6").getValue()) !== "TASK" || sheet.getName() === TEMPLATE_SHEET) {
    throw new Error("Open the project's tab first.");
  }

  const requests = requestsSheet(workbook);
  const requester = currentUserEmail(workbook, requests);
  const table = requestsTable(workbook, requests);
  table.addRow(-1, [newId(), new Date().toISOString(), sheet.getName(), requester, "Pending", ""]);
  console.log(`Requested a snapshot of ${sheet.getName()}; it will be emailed to ${requester}.`);
}

const TEMPLATE_SHEET = "Template";
const REQUESTS_SHEET = "Snapshot Requests";
const REQUESTS_TABLE = "SnapshotRequests";

function requestsSheet(workbook: ExcelScript.Workbook): ExcelScript.Worksheet {
  let ws = workbook.getWorksheet(REQUESTS_SHEET);
  if (!ws) {
    ws = workbook.addWorksheet(REQUESTS_SHEET);
    ws.setVisibility(ExcelScript.SheetVisibility.hidden);
  }
  return ws;
}

function requestsTable(workbook: ExcelScript.Workbook, ws: ExcelScript.Worksheet): ExcelScript.Table {
  let table = workbook.getTable(REQUESTS_TABLE);
  if (!table) {
    ws.getRange("A1:F1").setValues([["Id", "Requested (UTC)", "Tab", "Requester", "Status", "Note"]]);
    table = ws.addTable("A1:F1", true);
    table.setName(REQUESTS_TABLE);
  } else if (!table.getColumnByName("Id")) {
    // Tables made before the Id column existed: add it and give old rows an Id.
    const col = table.addColumn(0, null, "Id");
    const body = col.getRangeBetweenHeaderAndTotal();
    const ids = body.getValues().map(() => [newId()]);
    if (ids.length) body.setValues(ids);
  }
  return table;
}

function newId(): string {
  return "R" + Date.now().toString(36) + Math.floor(Math.random() * 1e6).toString(36);
}

// The signed-in user's email, read from the author of a throwaway comment.
function currentUserEmail(workbook: ExcelScript.Workbook, ws: ExcelScript.Worksheet): string {
  const comment = workbook.addComment(ws.getRange("H1"), "snapshot request");
  const email = comment.getAuthorEmail();
  comment.delete();
  if (!email) throw new Error("Couldn't tell who you are to send you the snapshot. Try again, or ask Hayden.");
  return email;
}
