// Function to show the sidebar
function openSidebar() {
  const htmlOutput = HtmlService.createHtmlOutputFromFile('sidebar')
      .setWidth(300)
      .setHeight(300);
  
  // Add the sidebar next to the "Help" menu
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}

// Function to process TNs from the sidebar input
function processTrackingNumbers(inputTNs) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Failed Delivery');  // Get the 'Failed Delivery' sheet
  
  if (!sheet) {
    SpreadsheetApp.getUi().alert("Sheet 'Failed Delivery' not found!");
    return;
  }
  
  const tnColumn = sheet.getRange("B2:B" + sheet.getLastRow()).getValues();  // Get TNs from column B in the 'Failed Delivery' sheet
  
  // Split the input string by newline or comma
  const emailTNs = inputTNs.split(/\r?\n|\s*,\s*/);
  
  // Loop through each TN in the email list and check for matches in the sheet
  for (let i = 0; i < emailTNs.length; i++) {
    const emailTN = emailTNs[i].trim();
    if (emailTN === "") continue;  // Skip empty inputs
    
    // Search for the matching TN in the sheet
    for (let j = 0; j < tnColumn.length; j++) {
      if (tnColumn[j][0] == emailTN) {
        // If a match is found, set "Return Received" in the first column (A)
        sheet.getRange(j + 2, 1).setValue("Return Received");
        break;  // Exit inner loop once a match is found
      }
    }
  }
}
