function generateMonthlyReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const templateSheet = ss.getSheetByName('template');
  const listSheet = ss.getSheetByName('devotees list');
  if (!templateSheet || !listSheet) throw new Error('Template or devotees list sheet not found');

  const dateRangeStr = templateSheet.getRange('A14').getValue(); // e.g., "30/06/2025 - 06/07/2025"
  const devoteeNames = listSheet.getRange('A2:A').getValues().flat().filter(n => n && n.toString().trim() !== '');

  // Create or clear the Monthly Reports sheet
  let reportSheet = ss.getSheetByName('Monthly Reports');
  if (!reportSheet) {
    reportSheet = ss.insertSheet('Monthly Reports');
  } else {
    reportSheet.clear();
  }

  // Set the date range string in A1 and format it
  reportSheet.getRange('A1').setValue(dateRangeStr);
  templateSheet.getRange('A14').copyFormatToRange(reportSheet, 1, 1, 1, 1); // Copy format to A1

  // Copy metric labels (A15:A24) to A2:A11 with formatting
  const labelRange = templateSheet.getRange('A15:A24');
  labelRange.copyTo(reportSheet.getRange('A2:A11'), { contentsOnly: false });

  // Copy devotee names to row 1 starting from B1, with formatting
  for (let i = 0; i < devoteeNames.length; i++) {
    const name = devoteeNames[i];
    const cell = reportSheet.getRange(1, i + 2);
    cell.setValue(name);
    templateSheet.getRange(14, i + 2).copyFormatToRange(reportSheet, i + 2, i + 2, 1, 1); // B14 → Bi1
  }

  // Fill metric data for each devotee from their M19:M28 block
  for (let i = 0; i < devoteeNames.length; i++) {
    const name = devoteeNames[i];
    const sheet = ss.getSheetByName(name);
    if (!sheet) continue;

    const sourceRange = sheet.getRange('M19:M28'); // 10 metrics
    const destRange = reportSheet.getRange(2, i + 2, 10, 1);
    sourceRange.copyTo(destRange, { contentsOnly: false }); // copy values + formatting
  }
}
