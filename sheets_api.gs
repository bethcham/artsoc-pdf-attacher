function fetchAndInsertCSPData() {
  // The API endpoint
  var centreCode = '2xx';  // ArtSoc code - find this on eactivities, it's a 3 digit number
  let productIDs = ['5xxxx']; // change depending on show - also on eactivities, it's a 5 digit number
  let sheetNames = ['Les Mis Shop 2'];  // change depending on show - name whatever you want
  
  var apiKey = 'xxxxxx-xxxxxx-xxxxxxx'; // long api code from eactivities
  
  // full API URL
  let urls = []
  for (let productID of productIDs) {
    urls.push('https://eactivities.union.ic.ac.uk/API/CSP/' + centreCode + '/products/' + productID + '/sales');
  }
  
  var options = {
    'method': 'GET',
    'headers': {
      'X-API-Key': apiKey,
      'Accept': 'application/json'
    }
  };
  
  for (var sheetNo = 0; sheetNo < urls.length; sheetNo++) {
    var response = UrlFetchApp.fetch(urls[sheetNo], options);
    
    var data = JSON.parse(response.getContentText());
    
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetNames[sheetNo]);
    
    // Check if the sheet exists
    if (!sheet) {
      // If the sheet does not exist, create it
      sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetNames[sheetNo]);
    }
    
    // Get all existing rows and store the emails to avoid duplicates
    var existingRows = sheet.getDataRange().getValues();
    var existingOrders = new Set(existingRows.slice(1).map(function(row) {
      return String(row[2]).trim(); // Assuming order is in the 3rd column (index 2)
    }));
    var existingLogins = new Set(existingRows.slice(1).map(function(row) {
      return String(row[3]).trim(); // Assuming logins is in the 4th column (index 3)
    }));
    
    // Insert headers if the sheet is empty
    if (existingRows.length === 1) {
      var headers = ["Time", "Product Line ID", "Order Number", "Login", "Gross Price", "Quantity", "First Name", "Surname", "Email", "Seat 1", "Seat 2"];
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
      existingRows = [headers];  // Update existingRows to include headers
    }
    
    for (var i = 0; i < data.length; i++) {
      var sale = data[i];
      var customer = sale.Customer;
      
      // Check if the sale's order no is already in the sheet
      var orderID = String(sale.OrderNumber).trim();
      var orderLogin = String(customer.Login).trim();
      if (!existingOrders.has(orderID)) {
        var row = [
          sale.SaleDateTime,
          sale.ProductLineID,       // ProductLineID
          sale.OrderNumber,           // order ID
          customer.Login, //short code
          sale.Price,               // Price
          sale.Quantity,            // Quantity
          customer.FirstName,       // First name
          customer.Surname,         // Surname
          customer.Email             // Email
        ];
        
        sheet.appendRow(row);

        if(existingLogins.has(orderLogin)){
          for (var r = 1; r < existingRows.length; r++) { // skip headers
            if (String(existingRows[r][3]).trim() === orderLogin) {
              sheet.getRange(r + 1, 4).setBackground('#FFFF00'); // highlight in yellow if there is an exisiting purchase by same customer (eg if tickets were bought in 2 separate transactions)
              var newRow = sheet.getLastRow();
              sheet.getRange(newRow, 4).setBackground('#FFFF00');
            }
          }
        }
      }
    }

    SpreadsheetApp.getUi().alert('New sale data fetched and inserted into: ' + sheetNames[sheetNo]);
  }
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Eactivities API')
      .addItem('Refresh Sale List', 'fetchAndInsertCSPData')
      .addToUi();
}
