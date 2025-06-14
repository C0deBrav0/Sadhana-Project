/**
 * Creates individual devotee sheets from the template's sadhana card section.
 * Only creates sheets for devotees listed in 'devotees list' sheet (column A, starting from row 2).
 */
function createDevoteeSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var templateSheet = ss.getSheetByName('template');
  var listSheet = ss.getSheetByName('devotees list');
  
  if (!templateSheet) throw new Error("Sheet named 'template' not found.");
  if (!listSheet) throw new Error("Sheet named 'devotees list' not found.");

  var names = listSheet.getRange('A2:A').getValues().flat().filter(name => name && name.toString().trim() !== '');

  names.forEach(name => {
    var sheetName = name.toString().trim();

    if (!ss.getSheetByName(sheetName)) {
      var newSheet = ss.insertSheet(sheetName);
      var sourceRange = templateSheet.getRange('A3:L12');
      sourceRange.copyTo(newSheet.getRange(1, 1), { contentsOnly: false });
    }
  });
}