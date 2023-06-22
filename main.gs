// Reusable function to include an HTML file
function include(filename) {
  return HtmlService.createTemplateFromFile(filename).evaluate().getContent();
}

// Reusable function to get the week number of a date
function getWeekNumber(date) {
  var onejan = new Date(date.getFullYear(), 0, 1);
  return Math.ceil((((date - onejan) / 86400000) + onejan.getDay() + 1) / 7);
}
// Reusable function to open the spreadsheet
function openSpreadsheet() {
  var sheetId = openSpreadsheet();
  return SpreadsheetApp.openById(sheetId);
}

// Handler for HTTP GET requests
function doGet() {
  var template = HtmlService.createTemplateFromFile("index");
  var html = template.evaluate()
    .setTitle("Lab Schedule")
    .addMetaTag("viewport", "width=device-width");

  return html;
}

// Format a date with Daylight Saving Time (DST) adjustment
function formatDateWithDST(date) {
  var timeZone = Session.getScriptTimeZone();
  var offset = Utilities.formatDate(date, timeZone, 'ZZ');
  var offsetNum = parseInt(offset.slice(0, 3), 10);
  var dst = offsetNum === 2;
  var timeZoneWithDST = 'GMT+' + (offsetNum - (dst ? 0 : 1)) + ':00';
  var formattedDate = Utilities.formatDate(date, timeZoneWithDST, 'yyyy-MM-dd HH:mm:ss.SSS\'Z\'');
  return formattedDate;
}

// Retrieve data from a spreadsheet
function getData() {
  var sheetId = openSpreadsheet();
  var data = {};
  var ss = SpreadsheetApp.openById(sheetId);
  var dataRange = ss.getDataRange();
  var values = dataRange.getValues();
  values.splice(0, 1);
  
  var newData = values.map(function(row) {
    var dateIndexes = [3, 13, 14, 15];
    
    dateIndexes.forEach(function(index) {
      var date = new Date(row[index]);
      row[index] = JSON.stringify(formatDateWithDST(date));
    });
    
    return row;
  });
  
  return {
    data: JSON.stringify(newData)
  };
}


// Get the values from the "weeknumber" range in "Sheet2"
function getSelectList() {
  var ss = openSpreadsheet();
  var prodSheet = ss.getSheetByName("Sheet2");
  return prodSheet.getRange("weeknumber").getValues();
}

// Get the values from the "gmao" range in "Sheet2"
function getSelectList1() {
  var ss = openSpreadsheet();
  var prodSheet = ss.getSheetByName("Sheet2");
  Logger.log(prodSheet.getRange("gmao").getValues());
  return prodSheet.getRange("gmao").getValues();
}

// Get the values from the "people" range in "Sheet2"
function getSelectList2() {
  var ss = openSpreadsheet();
  var prodSheet = ss.getSheetByName("Sheet2");
  return prodSheet.getRange("people").getValues();
}

// Get the values from the "t_oven" range in "Etuve"
function getSelectList3() {
  var ss = openSpreadsheet();
  var prodSheet = ss.getSheetByName("Etuve");
  var data = prodSheet.getRange("t_oven").getValues();
  Logger.log(data);
  return {
    data: JSON.stringify(data)
  };
}

// Get the values from the "gmaocode" range in "Etuve"
function getSelectList4() {
  var ss = openSpreadsheet();
  var prodSheet = ss.getSheetByName("Etuve");
  var data = prodSheet.getRange("gmaocode").getValues();
  Logger.log(data);
  return data;
}

// Append data to the "DATA" sheet in the spreadsheet
function send(data) {
  var ss = openSpreadsheet();
  var lock = LockService.getScriptLock();
  
  try {
    lock.waitLock(5000);
    Logger.log(data);
    var sheet = ss.getSheetByName("DATA");
    sheet.appendRow(data);
    setFormulas();
    return true;
  } catch (e) {
    return false;
  } finally {
    lock.releaseLock();
  }
}

// Set formulas in specific cells of the "DATA" sheet
function setFormulas() {
  var ss = openSpreadsheet();
  var sheet = ss.getSheetByName("DATA");
  var lastRow = sheet.getLastRow();
  var datarange = sheet.getDataRange();
  Logger.log(lastRow);
  var cell = datarange.getCell(lastRow, 17); // startweekday
  var cell2 = datarange.getCell(lastRow, 18); // finishweekday
  var cell3 = datarange.getCell(lastRow, 19); // startweeknum
  var cell1 = datarange.getCell(lastRow, 14); // finishdate
  var cell4 = datarange.getCell(lastRow, 21); // rowid
  var cell5 = datarange.getCell(lastRow, 20); // finishweeknum
  var cell6 = datarange.getCell(lastRow, 16); // startendphasedate
  var cell7 = datarange.getCell(lastRow, 15); // endprepdate

  cell.setFormulaR1C1('=WEEKDAY(R[0]C[-13])-1'); // startweekday
  cell1.setFormulaR1C1("=IF(WEEKDAY(R[0]C[-10]+SUM(R[0]C[-9]:R[0]C[-4])/24)-1=0;R[0]C[-10]+SUM(R[0]C[-9]:R[0]C[-4])/24+1;R[0]C[-10]+SUM(R[0]C[-9]:R[0]C[-4])/24)"); // finishdate
  cell2.setFormulaR1C1('=WEEKDAY(R[0]C[-4])-1'); // finishweekday
  cell3.setFormulaR1C1("=WEEKNUM(R[0]C[-15]-1)&\"-\"&YEAR(R[0]C[-15])"); // startweeknum
  cell4.setFormula("=ROW()"); // rowid
  cell5.setFormulaR1C1("=WEEKNUM(R[0]C[-6]-1)&\"-\"&YEAR(R[0]C[-6])"); // finishweeknum
  cell6.setFormulaR1C1("=IF(R[0]C[-6]<>0;R[0]C[-2]-((R[0]C[-6])/24);\"\")"); // startendphasedate
  cell7.setFormulaR1C1("=IF(R[0]C[-8]<>0;R[0]C[-11]+((R[0]C[-8])/24);\"\")"); // endprepdate
}
// Removes a row from the sheet based on the provided row index
function remove(liID) {
  var sheetId = openSpreadsheet();// Spreadsheet ID
  var ss = SpreadsheetApp.openById(sheetId); // Open the spreadsheet
  var sheet = ss.getSheetByName("DATA"); // Get the sheet by name
  sheet.deleteRow(liID); // Delete the specified row
}

// Retrieves row data for a given row index and formats date values
function getrowdata(rowID) {
  var rangeArray = getRangeValues(rowID, 1, 1, 21); // Get range values using helper function
  var newData = rangeArray.map(function(row) {
    row[3] = JSON.stringify(formatDateWithDST(new Date(row[3]))); // Format date value at index 3
    row[13] = JSON.stringify(formatDateWithDST(new Date(row[13]))); // Format date value at index 13
    row[14] = JSON.stringify(formatDateWithDST(new Date(row[14]))); // Format date value at index 14
    row[15] = JSON.stringify(formatDateWithDST(new Date(row[15]))); // Format date value at index 15
    return row;
  });

  return {
    rangeArray: JSON.stringify(newData) // Return formatted data as a JSON string
  };
}

// Retrieves the last row data from the sheet and formats date values
function getlastrowdata() {
  var sheetId = openSpreadsheet(); // Spreadsheet ID
  var ss = SpreadsheetApp.openById(sheetId); // Open the spreadsheet
  var sheet = ss.getSheetByName("DATA"); // Get the sheet by name
  var lastRow = sheet.getLastRow(); // Get the last row index
  var rangeArray = getRangeValues(lastRow, 1, 1, 21); // Get range values using helper function
  var newData = rangeArray.map(function(row) {
    row[3] = JSON.stringify(formatDateWithDST(new Date(row[3]))); // Format date value at index 3
    row[13] = JSON.stringify(formatDateWithDST(new Date(row[13]))); // Format date value at index 13
    row[14] = JSON.stringify(formatDateWithDST(new Date(row[14]))); // Format date value at index 14
    row[15] = JSON.stringify(formatDateWithDST(new Date(row[15]))); // Format date value at index 15
    return row;
  });

  return {
    rangeArray: JSON.stringify(newData) // Return formatted data as a JSON string
  };
}

// Retrieves range values from the sheet based on the provided parameters
function getRangeValues(row, column, numRows, numColumns) {
  var sheetId = openSpreadsheet(); // Spreadsheet ID
  var ss = SpreadsheetApp.openById(sheetId); // Open the spreadsheet
  var sheet = ss.getSheetByName("DATA"); // Get the sheet by name
  var rng = sheet.getRange(row, column, numRows, numColumns); // Get the range based on parameters
 return rng.getValues(); // Return the range values
}
// Retrieves milestone data for the specified number of rows
function getMilestoneData(numRows) {
  Utilities.sleep(500);
  Logger.log(numRows);

  var sheetId = openSpreadsheet();
  var sheet = SpreadsheetApp.openById(sheetId).getSheetByName("DATA");
  var lastRow = sheet.getLastRow();
  var rng = sheet.getRange(lastRow - numRows + 1, 1, numRows, 21);
  var rangeArray = rng.getValues();
  var newData = rangeArray.map(function(row) {
    row[3] = JSON.stringify(formatDateWithDST(new Date(row[3])));
    row[13] = JSON.stringify(formatDateWithDST(new Date(row[13])));
    row[14] = JSON.stringify(formatDateWithDST(new Date(row[14])));
    row[15] = JSON.stringify(formatDateWithDST(new Date(row[15])));
    return row;
  });

  Logger.log(rangeArray);
  return {
    rangeArray: JSON.stringify(newData)
  };
}

// Sets data for a specific row by deleting the row and appending the new data
function setData(dataEdit, rowID) {
  var sheetId = openSpreadsheet();
  var sheet = SpreadsheetApp.openById(sheetId).getSheetByName("DATA");

  sheet.deleteRow(rowID);
  sheet.appendRow(dataEdit);
  setFormulas();
}

// Sets the temperature for a given machine in the "Etuve" sheet
function setTemp(machine, temp) {
  var sheetId = openSpreadsheet();
  var prodSheet = SpreadsheetApp.openById(sheetId).getSheetByName("Etuve");
  var values = prodSheet.getRange("t_oven").getValues();

  for (let i = 0; i < values.length; i++) {
    if (values[i][0] === machine) {
      values[i][1] = temp; // Update the temperature value
      break;
    }
  }

  prodSheet.getRange("t_oven").setValues(values);
}

// Toggles the "received" flag for a given row in the "DATA" sheet
function toggleReceivedFlag(rowID, column) {
  var sheetId = openSpreadsheet();
  var sheet = SpreadsheetApp.openById(sheetId).getSheetByName("DATA");
  var range = sheet.getRange(rowID, column);
  var flag = range.getValue();

  range.setValue(flag === 0 ? 1 : 0); // Toggle the flag value
}

// Toggle received flag for column 12
function received(rowID) {
  toggleReceivedFlag(rowID, 12);
}

// Toggle received flag for column 13
function received1(rowID) {
  toggleReceivedFlag(rowID, 13);
}

// Checks if there is an ongoing task for the specified machine
function checkOngoingTask(machine) {
  var sheetId = openSpreadsheet();
  var ss = SpreadsheetApp.openById(sheetId);
  var sheet = ss.getSheetByName('BACKUP');
  var today = new Date();
  Logger.log(today);
  var dataRange = ss.getDataRange();
  var data = dataRange.getValues().slice(1);
  var isOngoingTask = false;

  data.forEach(row => {
    const gmao = row[0];
    const taskDateRange = [row[3], row[13]];
    const formattedStartDate = new Date(Utilities.formatDate(taskDateRange[0], "GMT+1", "yyyy-MM-dd"));
    const formattedEndDate = new Date(Utilities.formatDate(taskDateRange[1], "GMT+1", "yyyy-MM-dd"));

    if (gmao.includes(machine) && today >= formattedStartDate && today <= formattedEndDate) {
      isOngoingTask = true;
    }
  });
  return isOngoingTask;
}

// Searches for a matching task based on the input string
function searchMatch(inputString) {
  var sheetId = openSpreadsheet();
  var ss = SpreadsheetApp.openById(sheetId);
  var sheet = ss.getSheetByName('BACKUP');
  var dataRange = sheet.getDataRange();
  var data = dataRange.getValues().slice(1);
  var results = [];
  var threshold = 1; // Set threshold for similarity

  for (let i = 0; i < data.length; i++) {
    var task = data[i][2].toString();
    var distance = levenshteinDistance(inputString.toLowerCase(), task);

    if (
      /*distance <= threshold ||*/
      task.toLowerCase().includes(inputString.toLowerCase()) ||
      inputString.toLowerCase().includes(task.toLowerCase())
    ) {
      var date = new Date(data[i][3]);
      date = Utilities.formatDate(date, "GMT+1", "yyyy-MM-dd HH:mm'T'ss.SSS'Z'");
      var gmao = data[i][0];
      var resp = data[i][1];
      var week = data[i][18];
      var endDate = new Date(data[i][13]);
      endDate = Utilities.formatDate(endDate, "GMT+1", "yyyy-MM-dd HH:mm'T'ss.SSS'Z'");

      results.push([task, gmao, resp, date, week, endDate]);
    }
  }

  Logger.log(results);
  return {
    data: JSON.stringify(results)
  };
}
// Calculates the Levenshtein distance between two strings
function levenshteinDistance(a, b) {
  Logger.log(a);
  Logger.log(b);
  if (a.length === 0) return b.length;
  if (b.length === 0) return a.length;

  var matrix = [];

  // Increment along the first column of each row
  for (var i = 0; i <= b.length; i++) {
    matrix[i] = [i];
  }

  // Increment each column in the first row
  for (var j = 0; j <= a.length; j++) {
    matrix[0][j] = j;
  }

  // Fill in the rest of the matrix
  for (var i = 1; i <= b.length; i++) {
    for (var j = 1; j <= a.length; j++) {
      if (b.charAt(i - 1) == a.charAt(j - 1)) {
        matrix[i][j] = matrix[i - 1][j - 1];
      } else {
        matrix[i][j] = Math.min(
          matrix[i - 1][j - 1] + 1, // substitution
          Math.min(matrix[i][j - 1] + 1, matrix[i - 1][j] + 1) // insertion, deletion
        );
      }
    }
  }

  return matrix[b.length][a.length];
}

// Backs up data from the "DATA" sheet to the "BACKUP" sheet
function backup() {
  var sheetId = openSpreadsheet();
  var ss = SpreadsheetApp.openById(sheetId);
  var sheet = ss.getSheetByName('DATA');
  var backupSheet = ss.getSheetByName("BACKUP");
  
  var values = sheet.getDataRange().getValues();
  var found = false;
  var rowValues = backupSheet.getRange("A:T").getValues();
  
  for (var i = 0; i < values.length; i++) {
    for (var j = 0; j < rowValues.length; j++) {
      if (JSON.stringify(rowValues[j].slice(0, 20)) === JSON.stringify(values[i].slice(0, 20))) {
        found = true;
        break;
      }
    }
    if (!found) {
      backupSheet.appendRow(values[i]);
    }
    found = false;
  }
}

// Backs up data from the "DATA" sheet to the "BACKUP" sheet (optimized)
function backup2() {
  var sheetId = openSpreadsheet();
  var ss = SpreadsheetApp.openById(sheetId);
  var sheet = ss.getSheetByName('DATA');
  var backupSheet = ss.getSheetByName("BACKUP");

  // Load data into arrays
  var sheetData = sheet.getDataRange().getValues();
  var backupData = backupSheet.getDataRange().getValues();

  // Create a hash table to store the backup data for faster searching
  var backupHashTable = {};
  for (var i = 0; i < backupData.length; i++) {
    backupHashTable[JSON.stringify(backupData[i])] = true;
  }

  // Append new rows to the backup sheet
  var rowsToAppend = [];
  for (var i = 0; i < sheetData.length; i++) {
    var key = JSON.stringify(sheetData[i]);
    if (!backupHashTable[key]) {
      rowsToAppend.push(sheetData[i]);
      backupHashTable[key] = true;
    }
  }

  if (rowsToAppend.length > 0) {
    backupSheet.getRange(backupSheet.getLastRow() + 1, 1, rowsToAppend.length, rowsToAppend[0].length).setValues(rowsToAppend);
  }
}

// Deletes rows from the "DATA" sheet that are older than the specified threshold
function deleteoldrows() {
  var sheetId = openSpreadsheet();
  var ss = SpreadsheetApp.openById(sheetId);
  var sheet = ss.getSheetByName('DATA');
  var values = sheet.getDataRange().getValues();
  var daysThreshold = 30; // Number of days to consider data old
  var oneMonthAgo = new Date();

  oneMonthAgo.setDate(oneMonthAgo.getDate() - daysThreshold);

  for (var i = values.length - 1; i > 0; i--) {
    var rowDate = new Date(values[i][13]); // Assuming date is in column N
    var id = i + 1; // Add 1 to get the correct row number

    if (rowDate.getTime() < oneMonthAgo.getTime()) {
      // Delete row from main sheet
      sheet.deleteRow(id);
      Logger.log("deletedrow:" + id);
    }
  }
}

// Deletes rows from the "DATA" sheet that are older than the specified threshold (optimized)
function deleteOldRows2() {
  var sheetId = openSpreadsheet();
  var ss = SpreadsheetApp.openById(sheetId);
  var sheet = ss.getSheetByName('DATA');
  var daysThreshold = 30; // Number of days to consider data old
  var oneMonthAgo = new Date();
  oneMonthAgo.setDate(oneMonthAgo.getDate() - daysThreshold);
  var values = sheet.getDataRange().getValues();
  values.splice(0, 1);
  var rowsToDelete = [];

  values.forEach((row, rowIndex) => {
    var rowDate = new Date(row[13]);
    Logger.log(rowDate);

    if (rowDate.getTime() < oneMonthAgo.getTime()) {
      rowsToDelete.push(rowIndex + 2);
    }
  });

  Logger.log(rowsToDelete);
  deleteRows_(sheet, rowsToDelete);
}

// Deletes rows from the specified sheet
function deleteRows_(sheet, rowsToDelete) {
  const rowNumbers = rowsToDelete.filter((value, index, array) => array.indexOf(value) === index);
  const runLengths = getRunLengths_(rowNumbers.sort((a, b) => a - b));

  for (let i = runLengths.length - 1; i >= 0; i--) {
    sheet.deleteRows(runLengths[i][0], runLengths[i][1]);
  }

  return runLengths.length;
}

// Calculates the run lengths in the provided array of numbers
function getRunLengths_(numbers) {
  if (!numbers.length) {
    return [];
  }

  return numbers.reduce((accumulator, value, index) => {
    if (!index || value !== 1 + numbers[index - 1]) {
      accumulator.push([value]);
    }
    const lastIndex = accumulator.length - 1;
    accumulator[lastIndex][1] = (accumulator[lastIndex][1] || 0) + 1;
    return accumulator;
  }, []);
}

// Retrieves public holidays for the specified year and country code
function getPublicHolidays(year) {
  const countryCode = 'FR';
  const url = `https://date.nager.at/api/v3/PublicHolidays/${year}/${countryCode}`;
  const response = UrlFetchApp.fetch(url);
  const holidays = JSON.parse(response.getContentText());
  var dates = holidays.map(function(holiday) {
    return holiday.date;
  });
  Logger.log(dates);
  return dates;
}





