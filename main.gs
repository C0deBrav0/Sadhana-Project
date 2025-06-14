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
      var sourceRange = templateSheet.getRange('A1:L10');
      sourceRange.copyTo(newSheet.getRange(1, 1), { contentsOnly: false });
    }
  });
}

function updateTemplate(sheet) {
  var today = new Date();
  var dayOfWeek = today.getDay();

  // Monday of previous week
  var lastMonday = new Date(today);
  lastMonday.setDate(today.getDate() - 7 - ((dayOfWeek + 6) % 7));

  // Sunday of previous week
  var nextSunday = new Date(lastMonday);
  nextSunday.setDate(lastMonday.getDate() + 6);

  var formattedLastMonday = Utilities.formatDate(lastMonday, Session.getScriptTimeZone(), "d MMM");
  var formattedNextSunday = Utilities.formatDate(nextSunday, Session.getScriptTimeZone(), "d MMM");
  var dateRange = formattedLastMonday + " - " + formattedNextSunday;

  // Set weekly date range in B10 (or any suitable cell near totals)
  sheet.getRange('B10').setValue(dateRange);

  // Set daily dates Mon-Sun in B2:B8
  for (var i = 0; i < 7; i++) {
    var day = new Date(lastMonday);
    day.setDate(lastMonday.getDate() + i);
    var formattedDay = Utilities.formatDate(day, Session.getScriptTimeZone(), "d MMM");
    sheet.getRange('B' + (2 + i)).setValue(formattedDay);
  }
}
