# Phone Number Formatter

This Google Sheets script automatically formats phone numbers to the standard format `(123) 456-7890` and validates the input to ensure it is a 10-digit number. It strips out any non-numeric characters and provides immediate feedback if the input is invalid.

## Features

- **Automatic Formatting**: Converts any 10-digit phone number to the format `(123) 456-7890`.
- **Validation**: Ensures the input is exactly 10 digits long and strips out any non-numeric characters.
- **Error Handling**: Reverts to the previous value and displays an error message if the input is invalid.

## Installation

1. **Open Google Sheets**.
2. Go to **Extensions > Apps Script**.
3. Delete any code in the script editor and copy-paste the following code:

```javascript
function onEdit(e) {
  var sheet = e.source.getActiveSheet();
  var range = e.range;
  var oldValue = e.oldValue;
  
  // Define the column you want to apply the script to (e.g., column G)
  var column = 7; // Change this to the column number you want
  
  // Check if the edited cell is in the specified column
  if (range.getColumn() == column) {
    var value = range.getValue();
    
    // Handle empty input
    if (value === '') {
      sheet.getRange(range.getRow(), range.getColumn() + 1).setValue('');
      return;
    }
    
    // Ensure the value is a string
    value = value.toString();
    
    // Strip out non-numeric characters
    var cleanedValue = value.replace(/\D/g, '');
    
    // Check the length of the cleaned value
    if (cleanedValue.length == 10) {
      // Format the cleaned value as (123) 456-7890
      var formattedValue = '(' + cleanedValue.slice(0, 3) + ') ' + cleanedValue.slice(3, 6) + '-' + cleanedValue.slice(6);
      range.setValue(formattedValue);
      // Clear any previous error message
      sheet.getRange(range.getRow(), range.getColumn() + 1).setValue('');
    } else {
      // If the input is invalid, revert to the old value and show an error message
      range.setValue(oldValue);
      sheet.getRange(range.getRow(), range.getColumn() + 1).setValue('Error: Please enter a 10-digit phone number.');
      
      // Optionally, clear the error message after a delay
      SpreadsheetApp.flush();
      Utilities.sleep(3000);
      sheet.getRange(range.getRow(), range.getColumn() + 1).setValue('');
    }
  }
}
