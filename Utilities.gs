function createSheet(sheetName, sheetData) {
    var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheetToCreate = activeSpreadsheet.getSheetByName(sheetName);

    if (sheetToCreate == null) {
        sheetToCreate = activeSpreadsheet.insertSheet(sheetName);
    } else {
        sheetToCreate.clear();
    }

    if (sheetData.length != 0) {
      sheetToCreate.getRange(1, 1, sheetData.length, sheetData[0].length).setValues(sheetData);
      sheetToCreate.getRange(1, 2, sheetToCreate.getLastRow(), 1).setNumberFormat('@STRING@');
    }

    var defaultSheet = activeSpreadsheet.getSheetByName("Sheet1");
    if (defaultSheet != null) {
      activeSpreadsheet.deleteSheet(defaultSheet);
    }

    return sheetToCreate;
}

function deleteSheet(sheetName) {
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheetToDelete = activeSpreadsheet.getSheetByName(sheetName);

  if (sheetToDelete != null) {
    if (activeSpreadsheet.getSheets().length < 2) {
      activeSpreadsheet.insertSheet("Sheet1");
    }
    activeSpreadsheet.deleteSheet(sheetToDelete);
    if (sheetName.includes("Data")) {
      SpreadsheetApp.getUi()
        .alert(sheetName + " has been successfully cleared!");
    }
  }
}

function importData(importType, dataType) {
  var htmlTemplate = HtmlService.createTemplateFromFile("Form");
  htmlTemplate.importType = importType;
  htmlTemplate.dataType = dataType;
  
  var htmlOutput = htmlTemplate.evaluate()
    .setWidth(360)
    .setHeight(360);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput,
                                        (importType == "file" ? "Update " : importType == "API" ? "Updating " : "") + dataType
);
}

function processData(csvData, dataType) {
  var originalData = Utilities.newBlob(csvData).getDataAsString().split("\n").map(row => row.split(";"));

  var extractedData = (dataType == "Bingo Data") ? extractBingoDataFromFile(originalData) :
                    (dataType == "Lotto Data") ? extractLottoDataFromFile(originalData) :
                    [];

  var dataSheet = createSheet(dataType, extractedData);
  formatDataSheet(dataSheet);

  if (dataType == "Bingo Data") {
    generateBingoList();
  }

  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(dataType).activate();
}

function getAndProcessData(dataType) {
  var url = "https://app.ecwid.com/api/v3/" + getDocumentProperty("storeID") + "/orders?limit=100";

  var options = {
    method: "get",
    headers: {
      "Authorization": "Bearer " + getDocumentProperty("secretToken")
    }
  };
  
  var response = UrlFetchApp.fetch(url, options); 
  var json = response.getContentText();

  var originalData = JSON.parse(json);

  var extractedData = (dataType == "Bingo Data") ? extractBingoDataFromAPI(originalData) :
                      (dataType == "Lotto Data") ? extractLottoDataFromAPI(originalData) :
                      [];
  
  var dataSheet = createSheet(dataType, extractedData);
  formatDataSheet(dataSheet);

  if (dataType == "Bingo Data") {
    generateBingoList();
  }

  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(dataType).activate();
}

function normalizeDataArrays(...arrays) {
  var normalizedDataArray = [];

  arrays.forEach((arr, index) => {
    if (arr.length > 2) {
      var maxLength = Math.max(...arrays.flat().map(row => row.length));

      if (index > 0) {
        normalizedDataArray.push(new Array(maxLength).fill(""));
      }

      arr.forEach(row => {
        while (row.length < maxLength) {
          row.push("");
        }
        normalizedDataArray.push(row);
      });
    }
  });

  return normalizedDataArray;
}

function formatDataSheet(dataSheet) {
  var data = dataSheet.getDataRange().getValues();
  if (data != "") {
    for (var r = 0; r < data.length; r++) {
      if (r - 1 < 0 || data[r - 1].join("") == "") {
        var rowLength = data[r + 1].filter(cell => cell != "").length;
        dataSheet.getRange(r + 1, 1, 1, rowLength)
          .merge(); // Merge header row
        dataSheet.getRange(r + 1, 1, 1, rowLength)
          .setFontSize(12); // Set 12 font size for header
        dataSheet.getRange(r + 1, 1, 1, rowLength)
          .setFontStyle("italic"); // Set italic font for header
        dataSheet.getRange(r + 1, 1, 1, rowLength)
          .setHorizontalAlignment("center"); // Set centered alignment for header

        dataSheet.getRange(r + 2, 1, 1, rowLength)
          .setFontWeight("bold"); // Set bold font for column names
        dataSheet.getRange(r + 2, 1, 1, rowLength)
          .setHorizontalAlignment("center"); // Set centered alignment for column names

        dataSheet.getRange(r + 2, 1, 1, rowLength)
          .setBorder(false, false, true, false, false, false, '#000000', SpreadsheetApp.BorderStyle.DOTTED); // Set dotted border for Bingo data part
        dataSheet.getRange(r + 1, 1, 1, rowLength)
          .setBorder(true, true, true, true, false, false, '#000000', SpreadsheetApp.BorderStyle.SOLID_THICK); // Set thick border for header
      }
    }

    resizeColumns(dataSheet, data); // Set column size fit to data
  }
}

function resizeColumns(sheet, data) {
  sheet.autoResizeColumns(1, data[0].length);
  sheet.setColumnWidth(1, 100);
  sheet.setColumnWidth(2, 200);
}

function formatDate(date) {
  var day = date.getDate().toString().padStart(2, '0');
  var month = (date.getMonth() + 1).toString().padStart(2, '0');
  var year = date.getFullYear().toString().slice(-2);

  return `${day}.${month}.${year}`;
}

function formatDateTime(dateTime) {
  let date = new Date(dateTime);
  
  let options = { 
    year: 'numeric', 
    month: 'short', 
    day: 'numeric', 
    hour: 'numeric', 
    minute: '2-digit', 
    hour12: true 
  };

  return date.toLocaleString('en-US', options);
}

function extractNameAndSurname(comment) {
    var name = "";
    var surname = "";

    var commentParts = comment.split(/\s+/);
    switch (commentParts.length) {
        case 1:
            name = capitalizeFirstLetter(commentParts[0]);
            break;
        case 2:
            name = capitalizeFirstLetter(commentParts[0]);
            surname = capitalizeFirstLetter(commentParts[1]);
            break;
        case 3:
            name = capitalizeFirstLetter(commentParts[0]);
            surname = capitalizeFirstLetter(commentParts[1]) + " " + capitalizeFirstLetter(commentParts[2]);
            break;
    }

    return [name, surname];
}

function capitalizeFirstLetter(str) {
  return str.charAt(0).toUpperCase() + str.slice(1);
}

function checkEmailDomain(email) {
  const domains = ['@gmail.com', '@outlook.com', '@outlook.ie', '@hotmail.com',
                   '@yahoo.com', '@yahoo.ie', '@yahoo.co.uk', '@icloud.com'];
  for (var domain of domains) {
    if (email.endsWith(domain)) {
      return true;
    }
  }
  return false;
}

function setDocumentProperty(key, value) {
  var properties = PropertiesService.getDocumentProperties();
  properties.setProperty(key, value);
}

function getDocumentProperty(key) {
  var properties = PropertiesService.getDocumentProperties();
  return properties.getProperty(key);
}

function deleteDocumentProperty(key) {
  var properties = PropertiesService.getDocumentProperties();
  properties.deleteProperty(key);
}