function onFormSubmitTrigger(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Orders");
  var responses = e.namedValues; // {questionTitle: [answer]}

  // Find first empty template row (column E = Site Name blank), rows 4-303
  var data = sheet.getRange("E4:E303").getValues();
  var targetRow = -1;
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === "") { targetRow = i + 4; break; }
  }
  if (targetRow === -1) {
    MailApp.sendEmail("YOUR_EMAIL@example.com", "Orders sheet full",
      "All 300 template rows are used. Ask admin to add more rows with formulas copied down.");
    return;
  }

  // EDIT the question-title strings below to match your Google Form questions EXACTLY
  sheet.getRange(targetRow, 2).setValue(new Date());                              // B Date Posted
  sheet.getRange(targetRow, 3).setValue(responses["Market Person"][0]);           // C
  sheet.getRange(targetRow, 5).setValue(responses["Site Name"][0]);               // E
  sheet.getRange(targetRow, 8).setValue(responses["Grade"][0]);                   // H
  sheet.getRange(targetRow, 9).setValue(responses["Quantity (Cum)"][0]);          // I
  sheet.getRange(targetRow, 10).setValue(responses["Slump (mm)"][0]);             // J
  sheet.getRange(targetRow, 11).setValue(responses["Pump Required (Y/N)"][0]);    // K
  sheet.getRange(targetRow, 12).setValue(responses["Date of Requirement"][0]);    // L
  sheet.getRange(targetRow, 13).setValue(responses["Time at Site"][0]);           // M
}
