function moveDataToMasterSheet() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // Source sheets
  var sourceSheet1 = spreadsheet.getSheetByName("LineSettingandOilStatusForm");
  var sourceSheet2 = spreadsheet.getSheetByName("ShiftConsumablesForm");
  var sourceSheet3 = spreadsheet.getSheetByName("finishedProductForm");

  // Master sheet
  var masterSheet = spreadsheet.getSheetByName("MasterProductionLogSheet");

  // Get the data from each source sheet (including empty cells)
  var dataFromSheet1 = getAllDataFromSheet(sourceSheet1);
  var dataFromSheet2 = getAllDataFromSheet(sourceSheet2);
  var dataFromSheet3 = getAllDataFromSheet(sourceSheet3);

  // Transpose the data to combine rows
  var combinedData = transposeRows(
    dataFromSheet1,
    dataFromSheet2,
    dataFromSheet3
  );

  // Append the combined data as rows in the master sheet
  var numRows = combinedData.length;
  var numCols = combinedData[0].length;
  var targetRange = masterSheet.getRange(
    masterSheet.getLastRow() + 1,
    1,
    numRows,
    numCols
  );
  targetRange.setValues(combinedData);

  // Clear data from source sheets
  // sourceSheet1.getRange(2, 1, sourceSheet1.getLastRow() - 1, sourceSheet1.getLastColumn()).clearContent();
  // sourceSheet2.getRange(2, 1, sourceSheet2.getLastRow() - 1, sourceSheet2.getLastColumn()).clearContent();
  // sourceSheet3.getRange(2, 1, sourceSheet3.getLastRow() - 1, sourceSheet3.getLastColumn()).clearContent();

  // Call the function to create and send the PDF
  // createAndSendPDF();
}

// Function to get all data from a source sheet (including empty cells)
function getAllDataFromSheet(sheet) {
  var dataRange = sheet.getDataRange();
  var numRows = dataRange.getNumRows();
  var numCols = dataRange.getNumColumns();
  var data = dataRange.getValues();
  return data.slice(1); // Exclude the first row (header)
}

// Function to transpose rows from multiple sources into a single set of rows
function transposeRows() {
  var transposedData = [];
  var maxRows = 0;
  for (var i = 0; i < arguments.length; i++) {
    maxRows = Math.max(maxRows, arguments[i].length);
  }
  for (var i = 0; i < maxRows; i++) {
    var rowData = [];
    for (var j = 0; j < arguments.length; j++) {
      if (i < arguments[j].length) {
        rowData = rowData.concat(arguments[j][i]);
      } else {
        rowData = rowData.concat(Array(arguments[j][0].length).fill("")); // Fill with empty cells
      }
    }
    transposedData.push(rowData);
  }
  return transposedData;
}

// Creating PDF and sending in email
function createandSendPDF() {
  // Get the active spreadsheet
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // Specify the sheet name you want to convert to PDF
  var sheetName = "ProductionLogSheet"; // Corrected the sheet name to match your sheet

  // Get the sheet by name
  var sheet = spreadsheet.getSheetByName(sheetName);

  // Create a temporary folder in Google Drive
  var folder = DriveApp.createFolder("TempFolder");

  // Create a new spreadsheet and copy the desired sheet to it
  var newSpreadsheet = SpreadsheetApp.create("TempSpreadsheet");
  var newSheet = newSpreadsheet.getSheetByName("Sheet1"); // Rename 'Sheet1' to your desired sheet name
  sheet.copyTo(newSpreadsheet);
  newSpreadsheet.deleteSheet(newSheet);

  // Convert the copied sheet to PDF
  var pdf = DriveApp.getFileById(newSpreadsheet.getId()).getAs(
    "application/pdf"
  );
  pdf.setName(sheetName + ".pdf");
  folder.createFile(pdf);

  // Get the email address to send the PDF to
  var recipientEmail = "craftbentsolutionsdev@gmail.com"; // Replace with the recipient's email address
  Logger.log("Recipient Email: " + recipientEmail);
  // Send the email with the PDF attachment
  var subject = "Your Approved PDF Report";
  var body =
    "Please find the PDF APPROVED report for ProductionLogSheet attached.";
  MailApp.sendEmail(recipientEmail, subject, body, {
    attachments: [pdf],
    name: "Craft Bent Solutions", // Replace with your name or organization name
  });

  // Delete the temporary folder and the temporary spreadsheet
  folder.setTrashed(true);
  DriveApp.getFileById(newSpreadsheet.getId()).setTrashed(true);
  // Logger.log("Recipient Email: " + recipientEmail);
}

// custom menu
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Approved Data")
    .addItem("Click To Approve this Shift", "moveDataToMasterSheet")
    .addToUi();
}

// automatically add fields in finishedForm
function addSpecificCellsToFinishedProductForm() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // Source sheet (LineSettingandOilStatusForm)
  var sourceSheet1 = spreadsheet.getSheetByName("LineSettingandOilStatusForm");

  // Target sheet (finishedProductForm)
  var targetSheet = spreadsheet.getSheetByName("finishedProductForm");

  // Get values from specific cells in LineSettingandOilStatusForm

  var specificCellB2 = sourceSheet1.getRange("B2").getValue();
  var specificCellC2 = sourceSheet1.getRange("C2").getValue();

  // Find the next available row in the target sheet
  var lastRowInTarget = targetSheet.getLastRow() + 1;

  // Add specific cells (A2, B2, and C2) to the next available row in finishedProductForm

  targetSheet.getRange(lastRowInTarget, 2).setValue(specificCellB2);
  targetSheet.getRange(lastRowInTarget, 3).setValue(specificCellC2);
}
