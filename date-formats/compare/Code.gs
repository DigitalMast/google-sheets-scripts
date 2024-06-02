function applyConditionalFormatting() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Define the range where you want to apply the formatting
  var range = sheet.getRange('H2:H1000');

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

  // Rule 1: Red background if I has a date but J is empty and current date is greater than 7 days from I
  var redRule1 = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND(ISDATE(I2), ISBLANK(J2), TODAY() > I2 + 7)')
    .setBackground('#FF0000')
    .setRanges([range])
    .build();

  Logger.log("Red rule 1 created: %s", redRule1.getRanges().map(function(r) { return r.getA1Notation(); }).join(", "));

  // Rule 2: Green background if I has a date but J is empty and current date is less than or equal to 7 days from I
  var greenRule1 = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND(ISDATE(I2), ISBLANK(J2), TODAY() <= I2 + 7)')
    .setBackground('#00FF00')
    .setRanges([range])
    .build();

  Logger.log("Green rule 1 created: %s", greenRule1.getRanges().map(function(r) { return r.getA1Notation(); }).join(", "));

  // Rule 3: Red background if I has a date and J has a date and current date is greater than 7 days from J
  var redRule2 = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND(ISDATE(I2), ISDATE(J2), TODAY() > J2 + 7)')
    .setBackground('#FF0000')
    .setRanges([range])
    .build();

  Logger.log("Red rule 2 created: %s", redRule2.getRanges().map(function(r) { return r.getA1Notation(); }).join(", "));

  // Rule 4: Green background if I has a date and J has a date and current date is less than or equal to 7 days from J
  var greenRule2 = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND(ISDATE(I2), ISDATE(J2), TODAY() <= J2 + 7)')
    .setBackground('#00FF00')
    .setRanges([range])
    .build();

  Logger.log("Green rule 2 created: %s", greenRule2.getRanges().map(function(r) { return r.getA1Notation(); }).join(", "));

  // Add new rules to the sheet
  newRules.push(redRule1);
  newRules.push(greenRule1);
  newRules.push(redRule2);
  newRules.push(greenRule2);
  sheet.setConditionalFormatRules(newRules);

  Logger.log("Conditional formatting rules applied.");
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Scripts')
    .addItem('Apply Conditional Formatting', 'applyConditionalFormatting')
    .addToUi();
}
