function applyConditionalFormatting() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Define the range where you want to apply the formatting
  var range = sheet.getRange('J2:J10');
  
  // Remove existing conditional formatting rules for the specific range
  var rules = sheet.getConditionalFormatRules();
  var newRules = [];
  
  for (var i = 0; i < rules.length; i++) {
    var ruleRanges = rules[i].getRanges();
    var keepRule = true;
    
    for (var j = 0; j < ruleRanges.length; j++) {
      if (ruleRanges[j].getA1Notation() === range.getA1Notation()) {
        keepRule = false;
        break;
      }
    }
    
    if (keepRule) {
      newRules.push(rules[i]);
    }
  }
  
  // Rule 1: Red background if J2 is empty or more than seven days after I2 and H2 is filled, or if J2 is more than seven days from today
  var redRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND(NOT(ISBLANK(H2)), OR(ISBLANK(J2), AND(ISDATE(I2), ISDATE(J2), J2 > I2 + 7), AND(ISDATE(J2), J2 > TODAY() + 7)))')
    .setBackground('#FF0000')
    .setRanges([range])
    .build();
  
  // Rule 2: Green background if J2 is empty and less than or equal to seven days after I2, or if J2 is less than or equal to seven days after I2 and H2 is filled, and if J2 is less than or equal to seven days from today
  var greenRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND(NOT(ISBLANK(H2)), ISDATE(I2), OR(ISBLANK(J2), AND(ISDATE(J2), J2 <= I2 + 7), AND(ISDATE(J2), J2 <= TODAY() + 7)))')
    .setBackground('#00FF00')
    .setRanges([range])
    .build();
  
  // Add new rules to the sheet
  newRules.push(redRule);
  newRules.push(greenRule);
  sheet.setConditionalFormatRules(newRules);
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Scripts')
    .addItem('Apply Conditional Formatting', 'applyConditionalFormatting')
    .addToUi();
}
