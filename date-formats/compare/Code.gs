function applyConditionalFormatting() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Define the range where you want to apply the formatting (only column H)
  var rangeH = sheet.getRange('H2:H1000');

  // Clear all existing conditional formatting rules
  sheet.clearConditionalFormatRules();

  // Rule 1: Red background if I has a date but J is empty and current date is greater than 7 days from I
  var redRule1 = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND(ISDATE(I2), ISBLANK(J2), TODAY() > (I2 + 7))')
    .setBackground('#FF0000')
    .setRanges([rangeH])
    .build();

  // Rule 2: Green background if I has a date but J is empty and current date is less than or equal to 7 days from I
  var greenRule1 = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND(ISDATE(I2), ISBLANK(J2), TODAY() <= (I2 + 7))')
    .setBackground('#00FF00')
    .setRanges([rangeH])
    .build();

  // Rule 3: Red background if I has a date and J has a date and current date is greater than 7 days from J
  var redRule2 = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND(ISDATE(I2), ISDATE(J2), TODAY() > (J2 + 7))')
    .setBackground('#FF0000')
    .setRanges([rangeH])
    .build();

  // Rule 4: Green background if I has a date and J has a date and current date is less than or equal to 7 days from J
  var greenRule2 = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND(ISDATE(I2), ISDATE(J2), TODAY() <= (J2 + 7))')
    .setBackground('#00FF00')
    .setRanges([rangeH])
    .build();

  // Apply the new rules to the sheet
  var rules = [redRule1, greenRule1, redRule2, greenRule2];
  sheet.setConditionalFormatRules(rules);

  Logger.log("Conditional formatting rules applied.");
}
