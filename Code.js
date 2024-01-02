function mergeCsvFiles() {
  // Get the current spreadsheet
  var currentSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get the parent folder of the current spreadsheet
  var parentFolder = DriveApp.getFileById(currentSpreadsheet.getId()).getParents().next();
  
  // Get the 'CSV' folder in the parent folder
  var csvFolder = parentFolder.getFoldersByName('CSV').next();
  
  // Get all the files in the 'CSV' folder
  var files = csvFolder.getFiles();
  
  // Get the first sheet in the current spreadsheet
  var sheet = currentSpreadsheet.getSheets()[0];
  
  // Clear the contents of the sheet
  sheet.clearContents();
  
  // Create an object to hold all the data
  var allData = {};
  
  // Loop through all the files in the 'CSV' folder
  while (files.hasNext()) {
    var file = files.next();
    
    // Check if the file is a CSV file
    if (file.getMimeType() === 'text/csv') {
      // Get the data from the CSV file
      var dataString = file.getBlob().getDataAsString('utf-16le');
      
      // Parse the CSV data using a tab as the delimiter
      var csvData = Utilities.parseCsv(dataString, '\t');
      
      // Get the headings from the first row of the CSV data
      var headings = csvData[0];
      
      // Loop through the rest of the rows in the CSV data
      for (var i = 1; i < csvData.length; i++) {
        var row = csvData[i];
        
        // Loop through each cell in the row
        for (var j = 0; j < row.length; j++) {
          // Get the heading for this cell
          var heading = headings[j];
          
          // If this heading doesn't exist in allData, add it
          if (!allData[heading]) {
            allData[heading] = [];
          }
          
          // Add the cell data to the corresponding heading in allData
          allData[heading].push(row[j]);
        }
      }
    }
  }
  
  // Convert allData to a 2D array
  var data = Object.keys(allData).map(function(heading) {
    return [heading].concat(allData[heading]);
  });
  
  // Transpose the data so that each heading is a column
  var transposedData = data[0].map(function(col, i) {
    return data.map(function(row) {
      return row[i];
    });
  });
  
    
  // Get the maximum number of columns in the data
  var maxColumns = Math.max.apply(null, transposedData.map(function(row) { return row.length; }));

  // Write all the data to the sheet in one operation
  var range = sheet.getRange(1, 1, transposedData.length, maxColumns);
  range.setValues(transposedData);
  
  // Remove existing banding
  var bandings = range.getBandings();
  for (var i = 0; i < bandings.length; i++) {
    bandings[i].remove();
  }
  
  // Define the banding themes
  var bandingThemes = [
    SpreadsheetApp.BandingTheme.BLUE,
    SpreadsheetApp.BandingTheme.GREEN,
    SpreadsheetApp.BandingTheme.GREY,
    SpreadsheetApp.BandingTheme.LIGHT_GREY,
    SpreadsheetApp.BandingTheme.ORANGE,
    SpreadsheetApp.BandingTheme.YELLOW
  ];
  
  // Select a random banding theme
  var randomTheme = bandingThemes[Math.floor(Math.random() * bandingThemes.length)];
  
  // Apply the random banding theme to the range
  range.applyRowBanding(randomTheme);
  
  // Apply borders to the range
  var borderStyle = SpreadsheetApp.BorderStyle.SOLID;
  range.setBorder(true, true, true, true, true, true, null, borderStyle);
  
   // Auto-resize the columns and wrap text
  for (var i = 0; i < maxColumns; i++) {
    sheet.autoResizeColumn(i + 1);
    sheet.getRange(1, i + 1, transposedData.length).setWrap(true);
  }

  // Log the URL of the spreadsheet
  Logger.log('Merged CSV Files: ' + currentSpreadsheet.getUrl());
}
