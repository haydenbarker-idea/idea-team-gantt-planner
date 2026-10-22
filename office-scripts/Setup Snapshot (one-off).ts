// Setup Snapshot (one-off): prepares an existing Project Schedules workbook for the
// Email Snapshot to Client button. Run once in Excel for the web, then delete the script.
//  - Projects tab: adds the "Client Email" heading in K5 (formatted like J5).
//  - Template tab: clears hand-coloured day markers in row 5 of the chart.
//  - Instructions tab: adds a "Send the customer a snapshot" section after
//    "Finish or put away a project".
//  - Project tabs and Template: lines up the Run Add Section / Run Add Task /
//    Run Email Snapshot to Client buttons in row 1 (same size, even gaps).

/** One-off setup for Email Snapshot to Client: Client Email column, Template markers, Instructions. */
function main(workbook: ExcelScript.Workbook) {
  const list = workbook.getWorksheet("Projects");
  if (list && String(list.getRange("K5").getValue()) !== "Client Email") {
    list.getRange("K5").copyFrom(list.getRange("J5"), ExcelScript.RangeCopyType.formats);
    list.getRange("K5").setValue("Client Email");
    list.getRange("K:K").getFormat().setColumnWidth(170);
  }

  const tmpl = workbook.getWorksheet("Template");
  if (tmpl) {
    tmpl.getRange("I5:ED5").getFormat().getFill().setColor(tmpl.getRange("I4").getFormat().getFill().getColor());
  }

  // Row 1 buttons, left to right after the "All projects" link.
  const order = ["Add Section", "Add Task", "Email Snapshot"];
  const widths = [110, 110, 170];
  for (const ws of workbook.getWorksheets()) {
    if (String(ws.getRange("B6").getValue()) !== "TASK") continue;
    let left = ws.getRange("I1").getLeft() + 90;
    order.forEach((label, i) => {
      const shape = ws.getShapes().find(s => {
        try { return s.getTextFrame().getTextRange().getText().indexOf(label) >= 0; } catch { return false; }
      });
      if (shape) {
        shape.setLeft(left);
        shape.setTop(3.25);
        shape.setWidth(widths[i]);
        shape.setHeight(22);
      }
      left += widths[i] + 8;
    });
  }

  const ins = workbook.getWorksheet("Instructions");
  if (!ins) return;
  const col = ins.getRange("B1:B80").getValues().map(r => String(r[0]));
  if (col.indexOf("Send the customer a snapshot") >= 0) return;
  const finish = col.indexOf("Finish or put away a project");
  if (finish < 0) return;
  // The section's bullets end at the next blank row.
  let at = finish + 1;
  while (at < 80 && String(ins.getRange(`C${at + 1}`).getValue()) !== "") at++;
  const insertRow = at + 2;  // after the blank spacer row

  const lines = [
    ["h", "Send the customer a snapshot"],
    ["b", "On the project's tab, click Run Email Snapshot to Client. Within about two minutes you'll get an email with a PDF and an Excel copy of just that project, ready to forward to the client."],
    ["b", "The snapshot is a frozen copy as of today: no other projects, no formulas, no invoice markers or internal controls."],
    ["b", "The client's address comes from the Client Email column on the Projects tab; fill it in so the email tells you where to send it."],
    ["x", ""],
  ];
  ins.getRange(`${insertRow}:${insertRow + lines.length - 1}`).insert(ExcelScript.InsertShiftDirection.down);
  const headingFmt = ins.getRange(`B${finish + 1}`);
  const bulletRow = finish + 2;
  lines.forEach(([kind, text], i) => {
    const r = insertRow + i;
    if (kind === "h") {
      ins.getRange(`B${r}`).copyFrom(headingFmt, ExcelScript.RangeCopyType.formats);
      ins.getRange(`B${r}`).setValue(text);
    } else if (kind === "b") {
      ins.getRange(`B${r}:C${r}`).copyFrom(ins.getRange(`B${bulletRow}:C${bulletRow}`), ExcelScript.RangeCopyType.formats);
      ins.getRange(`B${r}`).setValue(String(ins.getRange(`B${bulletRow}`).getValue()));
      ins.getRange(`C${r}`).setValue(text);
    }
  });
  ins.getRange(`${insertRow}:${insertRow + lines.length - 1}`).getFormat().autofitRows();
}
