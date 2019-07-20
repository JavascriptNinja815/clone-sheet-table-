// make menu item
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Add Property")
    .addItem("Add Your Next Property", "addProperty")
    .addToUi();
}

function addProperty() {
  // get sheet by sheet name
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Dev for Property Information");
  var propertyNums = sheet.getDataRange().getValues().splice(1);
  if (propertyNums.length == 0) {
    var startRowNum = 1;
    var propertyTitle = "Property #1";
  } else {
    var startRowNum = propertyNums.length + 3;
    var propertyTitle = "Property #" + ((propertyNums.length + 2) / 8 + 1);
  }

  // define first column
  var col = [propertyTitle, "Property Name / Nickname", "Property Description", "Property Address", "Number of Bathrooms", "Number of Bedrooms", "iCal URL"];
  var range = sheet.getRange(startRowNum, 1, 7, 2);
  // set borders
  for (var i = 1; i < 8; i++) {
    for (var j = 1; j < 3; j++) {
      range.getCell(i, j).setBorder(true, true, true, true, false, false);
    }
  }

  // decoration table and set initial values
  range.getCell(1, 1).setValue(col[0]);
  range.getCell(1, 1).setBackground('#2bc2b0');

  range.getCell(2, 1).setValue(col[1]);
  range.getCell(3, 1).setValue(col[2]);
  range.getCell(4, 1).setValue(col[3]);

  range.getCell(5, 1).setValue(col[4]);
  range.getCell(5, 1).setFontColor('blue');
  range.getCell(5, 2).setValue(1);

  range.getCell(6, 1).setValue(col[5]);
  range.getCell(6, 1).setFontColor('blue');
  range.getCell(6, 2).setValue(1);
  
  range.getCell(7, 1).setValue(col[6]);

  // make dropdown inside cell
  var arrayValues = ['0.5', '1', '1.5', '2', '2.5', '3', '3.5', '4', '4.5', '5'];
  var rule = SpreadsheetApp.newDataValidation().requireValueInList(arrayValues).build();
  var string = 'Dev for Property Information!B'+ (startRowNum + 4) + ':B' + (startRowNum + 5);
  sheet.getRange(string).setDataValidation(rule);
}
