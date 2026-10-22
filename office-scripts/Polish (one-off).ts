// Polish (one-off): QA fixes for the Client Email column and the Instructions wording.

/** One-off QA polish: Client Email header/border and Instructions wording. */
function main(workbook: ExcelScript.Workbook) {
  const list = workbook.getWorksheet("Projects");
  if (list) {
    list.getRange("K5").getFormat().setHorizontalAlignment(list.getRange("J5").getFormat().getHorizontalAlignment());
    const src = list.getRange("J6").getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeBottom);
    const dst = list.getRange("K6").getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeBottom);
    dst.setStyle(src.getStyle());
    dst.setWeight(src.getWeight());
    dst.setColor(src.getColor());
  }

  const ins = workbook.getWorksheet("Instructions");
  if (!ins) return;
  const fixes: [string, string][] = [
    ["Send the customer a snapshot", "Send the client a snapshot"],
    ["Within about two minutes you'll get an email", "Within about two minutes the person who clicked gets an email"],
    ["client and project manager are optional", "client, project manager and client email are optional"],
    ["progress and status. Click", "progress, status and client email. Click"],
  ];
  const range = ins.getRange("B1:C80");
  const values = range.getValues();
  values.forEach((row, r) => row.forEach((v, c) => {
    let text = String(v);
    let changed = false;
    for (const [from, to] of fixes) {
      if (text.indexOf(from) >= 0) { text = text.replace(from, to); changed = true; }
    }
    if (changed) {
      const cell = range.getCell(r, c);
      cell.setValue(text);
      if (c === 1) cell.getEntireRow().getFormat().autofitRows();
      console.log(`Row ${r + 1}: ${text}`);
    }
  }));
}
