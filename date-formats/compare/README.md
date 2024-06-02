# Google Sheets Conditional Formatting Script

## Overview

This Google Apps Script helps to automatically apply conditional formatting to a range of cells in a Google Sheet. The script changes the background color of cells based on the staleness of dates:
- **Red**: If the date is stale (i.e., empty or more than seven days past another date) and a specific cell has a dropdown value.
- **Green**: If the date is not stale (i.e., less than or equal to seven days past another date) or the specific cell is empty.

## Problem Solved

The main problem addressed by this script is to visually highlight dates that are considered stale (older than seven days) or non-stale (up to seven days old) based on a reference date. This is particularly useful for tracking deadlines, follow-ups, or any time-sensitive tasks.

## Features

- Automatically applies conditional formatting to a specified range.
- Highlights cells with stale dates in red.
- Highlights cells with non-stale dates in green.
- Considers a specific cell (e.g., a dropdown value) to determine if formatting should be applied.

## Installation

1. **Open Your Google Sheet**:
   - Open the Google Sheet where you want to apply the script.

2. **Open the Script Editor**:
   - Click on **Extensions** in the menu.
   - Select **Apps Script**.

3. **Create the Script**:
   - Delete any existing code in the script editor.
   - Copy and paste the following script:

   ```javascript
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
     
     // Rule 1: Red background if J2 is empty or more than seven days after I2 and H2 is filled
     var redRule = SpreadsheetApp.newConditionalFormatRule()
       .whenFormulaSatisfied('=AND(NOT(ISBLANK(H2)), OR(ISBLANK(J2), AND(ISDATE(I2), J2 > I2 + 7)))')
       .setBackground('#FF0000')
       .setRanges([range])
       .build();
     
     // Rule 2: Green background if J2 is empty and less than or equal to seven days after I2, or if J2 is less than or equal to seven days after I2 and H2 is filled
     var greenRule = SpreadsheetApp.newConditionalFormatRule()
       .whenFormulaSatisfied('=AND(NOT(ISBLANK(H2)), ISDATE(I2), OR(ISBLANK(J2), J2 <= I2 + 7))')
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
