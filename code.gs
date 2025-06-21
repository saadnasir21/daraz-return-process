// Adds a custom menu to easily open the sidebar when the spreadsheet is opened
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Returns')
    .addItem('Open Sidebar', 'openSidebar')
    .addToUi();
}

// Function to show the sidebar
function openSidebar() {
  const htmlOutput = HtmlService.createHtmlOutputFromFile('sidebar')
    .setWidth(300)
    .setHeight(300);

  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}

// Function to process TNs from the sidebar input
function processTrackingNumbers(inputTNs) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Failed Delivery');

  if (!sheet) {
    SpreadsheetApp.getUi().alert("Sheet 'Failed Delivery' not found!");
    return;
  }

  // Build a lookup table of tracking number -> row for faster searches
  const tnValues = sheet.getRange('B2:B' + sheet.getLastRow()).getValues();
  const tnMap = {};
  for (let i = 0; i < tnValues.length; i++) {
    const value = tnValues[i][0];
    if (value) {
      tnMap[value.toString().trim()] = i + 2; // row index in sheet
    }
  }

  // Split the input string by newline or comma
  const emailTNs = inputTNs.split(/\r?\n|\s*,\s*/);

  let updated = 0;
  emailTNs.forEach((raw) => {
    const tn = raw.trim();
    if (tn === '') return;
    const row = tnMap[tn];
    if (row) {
      sheet.getRange(row, 1).setValue('Return Received');
      updated++;
    }
  });

  return 'Updated ' + updated + ' tracking numbers.';
}
