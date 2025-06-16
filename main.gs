// @ts-nocheck
/**
 * Creates individual devotee sheets from the template's sadhana card section.
 * Only creates sheets for devotees listed in 'devotees list' sheet (column A, starting from row 2).
 * Copies the sadhana card range (A1:L10) from the 'template' sheet to the new devotee sheets.
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

      // Fill current week's dates and days in the new sheet
      var today = new Date();
      today.setHours(0, 0, 0, 0);
      var day = today.getDay();
      var monday = new Date(today);
      monday.setDate(today.getDate() - ((day === 0) ? 6 : day - 1)); // Get Monday

      for (var i = 0; i < 7; i++) {
        var date = new Date(monday.getTime() + i * 86400000);
        var formattedDate = `${date.getDate()} ${date.toLocaleString('en-US', { month: 'short' })}`;
        var dayName = date.toLocaleDateString('en-US', { weekday: 'short' });

        newSheet.getRange(2 + i, 1).setValue(formattedDate); // Column A (Dates)
        newSheet.getRange(2 + i, 2).setValue(dayName);        // Column B (Day Names)
      }
    }
  });
}


function appendNewWeekToDevoteeSheetsMerged() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var templateSheet = ss.getSheetByName('template');
  var listSheet = ss.getSheetByName('devotees list');

  if (!templateSheet || !listSheet) throw new Error("Missing 'template' or 'devotees list' sheet.");

  var devoteeNames = listSheet.getRange('A2:A').getValues().flat().filter(n => n && n.toString().trim() !== '');
  var templateRange = templateSheet.getRange('A1:L10');
  var blockHeight = templateRange.getNumRows();
  var gapRows = 2;

  devoteeNames.forEach(name => {
    var sheet = ss.getSheetByName(name.toString().trim());
    if (!sheet) return;

    // Parse Monday from week range
    var topWeekRange = sheet.getRange(1, 3).getValue();
    var monday = (function(str) {
      if (!str) return null;
      var parts = str.split(' - ');
      if (parts.length < 1) return null;

      var [dayStr, monthStr] = parts[0].trim().split(' ');
      var day = parseInt(dayStr);
      var months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
      var month = months.indexOf(monthStr);
      if (month === -1 || isNaN(day)) return null;

      var today = new Date();
      var year = today.getFullYear();
      var date = new Date(year, month, day);

      if ((date - today) > 3 * 86400000) date.setFullYear(year - 1);

      var d = new Date(date);
      d.setHours(0, 0, 0, 0);
      var weekday = d.getDay();
      var diff = (weekday === 0) ? -6 : 1 - weekday;
      d.setDate(d.getDate() + diff);
      return d;
    })(topWeekRange) || (function(date) {
      var d = new Date(date);
      d.setHours(0, 0, 0, 0);
      var weekday = d.getDay();
      var diff = (weekday === 0) ? -6 : 1 - weekday;
      d.setDate(d.getDate() + diff);
      return d;
    })(new Date());

    var nextMonday = new Date(monday.getTime() + 7 * 86400000);

    // Insert new block
    sheet.insertRows(1, blockHeight + gapRows);
    templateRange.copyTo(sheet.getRange(1, 1), { contentsOnly: false });
    sheet.getRange(blockHeight + 1, 1, gapRows, sheet.getMaxColumns()).clearContent();

    // Fill only dates and days (no header update)
    for (var i = 0; i < 7; i++) {
      var d = new Date(nextMonday.getTime() + i * 86400000);
      var formattedDate = `${d.getDate()} ${d.toLocaleString('en-US', { month: 'short' })}`;
      var dayName = d.toLocaleDateString('en-US', { weekday: 'short' });
      sheet.getRange(1 + 1 + i, 1).setValue(formattedDate); // Column A
      sheet.getRange(1 + 1 + i, 2).setValue(dayName);       // Column B
    }
  });
}
function syncDevoteeSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var templateSheet = ss.getSheetByName('template');
  var listSheet = ss.getSheetByName('devotees list');

  if (!templateSheet) throw new Error("Sheet named 'template' not found.");
  if (!listSheet) throw new Error("Sheet named 'devotees list' not found.");

  var names = listSheet.getRange('A2:A').getValues()
    .flat()
    .filter(name => name && name.toString().trim() !== '');

  var sourceRange = templateSheet.getRange('A1:L10');

  // Calculate current week's Monday date
  var today = new Date();
  var day = today.getDay();
  var diff = (day === 0 ? -6 : 1) - day; // Sunday(0) => last Monday
  var currentMonday = new Date(today);
  currentMonday.setDate(today.getDate() + diff);
  currentMonday.setHours(0, 0, 0, 0);

  names.forEach(function(name) {
    var sheetName = name.toString().trim();
    var sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sourceRange.copyTo(sheet.getRange(1, 1), { contentsOnly: false });
      ss.setActiveSheet(sheet);
      ss.moveActiveSheet(ss.getSheets().length);
    } else {
      sourceRange.copyTo(sheet.getRange(1, 1), { contentsOnly: false });
    }

    for (var i = 0; i < 7; i++) {
      var date = new Date(currentMonday.getTime() + i * 86400000);
      var formattedDate = date.getDate() + ' ' + date.toLocaleString('en-US', { month: 'short' });
      var dayName = date.toLocaleDateString('en-US', { weekday: 'short' });
      sheet.getRange(2 + i, 1).setValue(formattedDate);
      sheet.getRange(2 + i, 2).setValue(dayName);
    }
  });
}
