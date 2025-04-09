const COLUMNS = {
  ORDER_NUMBER: 1,
  ORDER_COMMENTS: 60,
  ORDER_EMAIL: 2,
  ORDER_PLAN: 4,
  ORDER_TOTAL: 7,
  ORDER_QUANTITY: 8,
  ORDER_STATUS: 40,
  ORDER_TIME: 14
};

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Import')
      .addSubMenu
      (
          ui.createMenu('Bingo Data')
          .addItem('From file', 'importBingoDataFromFile')
          .addItem('From API', 'importBingoDataFromAPI')
      )
      .addSubMenu
      (
          ui.createMenu('Lotto Data')
          .addItem('From file', 'importLottoDataFromFile')
          .addItem('From API', 'importLottoDataFromAPI')
      )
      .addToUi();
  ui.createMenu('Clear')
      .addItem('Bingo Data', 'clearBingoData')
      .addItem('Lotto Data', 'clearLottoData')
      .addToUi();
}

function onEdit(e) {
  var listSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Bingo List");

  if (e.source.getActiveSheet().getName() != listSheet.getName()) {
    return;
  }

  var listData = listSheet.getDataRange().getValues();
  var range = e.range;

  if (range.getNumRows() == 1 && range.getNumColumns() == 1) {
    var col = range.getColumn();
    var row = range.getRow();
    
    if ((col == 3 || col == 4) && row != 1) {
      var newValue = listData[row - 1][col - 1]
      var orderNumber = listData[row - 1][0];

      for (var r = row; listData[r - 1][0] == orderNumber; r++) {
        var updatedCell = listSheet.getRange(r, col);
        updatedCell.setValue(newValue);
        updatedCell.setFontStyle("normal");
        updatedCell.setFontColor(null);
      }

      resizeColumns(listSheet, listData);
    }
  }
}

function importBingoDataFromFile() {
  importData("file", "Bingo Data")
}

function importLottoDataFromFile() {
  importData("file", "Lotto Data")
}

function importBingoDataFromAPI() {
  importData("API", "Bingo Data")
}

function importLottoDataFromAPI() {
  importData("API", "Lotto Data")
}

function extractBingoDataFromFile(originalData) {
  var bingoBooksData = [
    ["Bingo Books"],
    ["NUMBER", "COMMENT", "EMAIL", "REGISTRATION PLAN", "VALUE", "QUANTITY"]
  ];
  var bingoVouchersData = [
    ["Bingo Vouchers"],
    ["TOTAL VALUE", "TO", "FROM"]
  ];
  var totalRows = 0;
  var newRows = 0;
  var bingoBooksRows = 0;
  var bingoVouchersRows = 0;
  var bingoRows = 0;
  var refundedRows = 0;
  var lottoRows = 0;
  var unknownRows = 0;

  var prevBingoOrderNumber;
  var setupBingo = true;

  var setupBingoBooks = true;
  var firstBingoBooksTime = "?";
  var lastBingoBooksTime = "?";

  var setupBingoVouchers = true;
  var firstBingoVouchersTime = "?";
  var lastBingoVouchersTime = "?";

  for (var r = 1; r < originalData.length - 1; r++) {
    totalRows++;

    if (setupBingo) {
      setupBingo = false;
      prevBingoOrderNumber = parseInt(getDocumentProperty("prevBingoOrderNumber"), 10) || Number.MIN_VALUE;
      setDocumentProperty("prevBingoOrderNumber", originalData[r][COLUMNS.ORDER_NUMBER - 1]);
    }

    var orderNumber = parseInt(originalData[r][COLUMNS.ORDER_NUMBER - 1], 10);

    if (orderNumber > prevBingoOrderNumber) {
      newRows++;

      if (originalData[r][COLUMNS.ORDER_PLAN - 1].includes("Bingo")) {
        if (originalData[r][COLUMNS.ORDER_STATUS - 1] == "Paid") {
          bingoRows++;

          var newRow = [];
          var bingoVoucherIndex = originalData[r][COLUMNS.ORDER_PLAN - 1].indexOf("Voucher");

          if (bingoVoucherIndex == -1) {
            bingoBooksRows++;

            var booksTime = originalData[r][COLUMNS.ORDER_TIME - 1];
            if (setupBingoBooks) {
              setupBingoBooks = false;
              lastBingoBooksTime = booksTime;
            }
            firstBingoBooksTime = booksTime;

            newRow.push(originalData[r][COLUMNS.ORDER_NUMBER - 1]);
            newRow.push(originalData[r][COLUMNS.ORDER_COMMENTS - 1]);
            newRow.push(originalData[r][COLUMNS.ORDER_EMAIL - 1]);
            newRow.push(originalData[r][COLUMNS.ORDER_PLAN - 1]);
            newRow.push(originalData[r][COLUMNS.ORDER_TOTAL - 1]);
            newRow.push(originalData[r][COLUMNS.ORDER_QUANTITY - 1]);

            bingoBooksData.push(newRow);
          }

          else {
            bingoVouchersRows++;

            var vouchersTime = originalData[r][COLUMNS.ORDER_TIME - 1];
            if (setupBingoVouchers) {
              setupBingoVouchers = false;
              lastBingoVouchersTime = vouchersTime;
            }
            firstBingoVouchersTime = vouchersTime;

            var voucherQuantity = originalData[r][COLUMNS.ORDER_QUANTITY - 1];
            var voucherValue = originalData[r][COLUMNS.ORDER_PLAN - 1].substring(bingoVoucherIndex + "Voucher".length + 1);

            if (voucherQuantity > 1) {
              newRow.push(voucherQuantity + " x " + voucherValue);
            }
            else {
              newRow.push(voucherValue);
            }

            newRow.push(originalData[r][COLUMNS.ORDER_COMMENTS - 1]);
            newRow.push(originalData[r][COLUMNS.ORDER_EMAIL - 1]);

            bingoVouchersData.push(newRow);
          }
        }

        else {
          refundedRows++;
        }
      }

      else {
        if (originalData[r][COLUMNS.ORDER_PLAN - 1].includes("Community lotto")) {
          lottoRows++;
        }

        else {
          unknownRows++;
        }
      }
    }
  }

  bingoBooksData[0][0] += " (" + firstBingoBooksTime + " - " + lastBingoBooksTime + ")"
  bingoVouchersData[0][0] += " (" + firstBingoVouchersTime + " - " + lastBingoVouchersTime + ")"

  setDocumentProperty("totalRows", totalRows);
  setDocumentProperty("newRows", newRows);
  setDocumentProperty("bingoBooksRows", bingoBooksRows);
  setDocumentProperty("bingoVouchersRows", bingoVouchersRows);
  setDocumentProperty("bingoRows", bingoRows);
  setDocumentProperty("refundedRows", refundedRows);
  setDocumentProperty("lottoRows", lottoRows);
  setDocumentProperty("unknownRows", unknownRows);

  return normalizeDataArrays(bingoBooksData, bingoVouchersData);
}

function extractBingoDataFromAPI(originalData) {
  var bingoBooksData = [
    ["Bingo Books"],
    ["NUMBER", "COMMENT", "EMAIL", "REGISTRATION PLAN", "VALUE", "QUANTITY"]
  ];
  var bingoVouchersData = [
    ["Bingo Vouchers"],
    ["TOTAL VALUE", "TO", "FROM"]
  ];
  var totalRows = 0;
  var newRows = 0;
  var bingoBooksRows = 0;
  var bingoVouchersRows = 0;
  var bingoRows = 0;
  var refundedRows = 0;
  var lottoRows = 0;
  var unknownRows = 0;

  var prevBingoOrderNumber;
  var setupBingo = true;

  var setupBingoBooks = true;
  var firstBingoBooksTime = "?";
  var lastBingoBooksTime = "?";

  var setupBingoVouchers = true;
  var firstBingoVouchersTime = "?";
  var lastBingoVouchersTime = "?";

  for (var or = 0; or < originalData.count; or++) {
    for (var r = 0; r < originalData.items[or].items.length; r++) {
      totalRows++;

      if (setupBingo) {
        setupBingo = false;
        prevBingoOrderNumber = parseInt(getDocumentProperty("prevBingoOrderNumber"), 10) || Number.MIN_VALUE;
        setDocumentProperty("prevBingoOrderNumber", originalData.items[or].orderNumber);
      }

      var orderNumber = parseInt(originalData.items[or].orderNumber, 10);

      if (orderNumber > prevBingoOrderNumber) {
        newRows++;

        if (originalData.items[or].items[r].name.includes("Bingo")) {
          if (originalData.items[or].paymentStatus == "PAID") {
            bingoRows++;

            var newRow = [];
            var bingoVoucherIndex = originalData.items[or].items[r].name.indexOf("Voucher");

            if (bingoVoucherIndex == -1) {
              bingoBooksRows++;

              var booksTime = formatDateTime(originalData.items[or].createDate);
              if (setupBingoBooks) {
                setupBingoBooks = false;
                lastBingoBooksTime = booksTime;
              }
              firstBingoBooksTime = booksTime;

              newRow.push(originalData.items[or].orderNumber);
              newRow.push(originalData.items[or].orderComments);
              newRow.push(originalData.items[or].email);
              newRow.push(originalData.items[or].items[r].name);
              newRow.push(originalData.items[or].items[r].price * originalData.items[or].items[r].quantity);
              newRow.push(originalData.items[or].items[r].quantity);

              bingoBooksData.push(newRow);
            }

            else {
              bingoVouchersRows++;

              var vouchersTime = formatDateTime(originalData.items[or].createDate);
              if (setupBingoVouchers) {
                setupBingoVouchers = false;
                lastBingoVouchersTime = vouchersTime;
              }
              firstBingoVouchersTime = vouchersTime;

              var voucherQuantity = originalData.items[or].items[r].quantity;
              var voucherValue = originalData.items[or].items[r].name.substring(bingoVoucherIndex + "Voucher".length + 1);

              if (voucherQuantity > 1) {
                newRow.push(voucherQuantity + " x " + voucherValue);
              }
              else {
                newRow.push(voucherValue);
              }

              newRow.push(originalData.items[or].orderComments);
              newRow.push(originalData.items[or].email);

              bingoVouchersData.push(newRow);
            }
          }

          else {
            refundedRows++;
          }
        }

        else {
          if (originalData.items[or].items[r].name.includes("Community lotto")) {
            lottoRows++;
          }

          else {
            unknownRows++;
          }
        }
      }
    }
  }

  bingoBooksData[0][0] += " (" + firstBingoBooksTime + " - " + lastBingoBooksTime + ")"
  bingoVouchersData[0][0] += " (" + firstBingoVouchersTime + " - " + lastBingoVouchersTime + ")"

  setDocumentProperty("totalRows", totalRows);
  setDocumentProperty("newRows", newRows);
  setDocumentProperty("bingoBooksRows", bingoBooksRows);
  setDocumentProperty("bingoVouchersRows", bingoVouchersRows);
  setDocumentProperty("bingoRows", bingoRows);
  setDocumentProperty("refundedRows", refundedRows);
  setDocumentProperty("lottoRows", lottoRows);
  setDocumentProperty("unknownRows", unknownRows);

  return normalizeDataArrays(bingoBooksData, bingoVouchersData);
}

function extractLottoDataFromFile(originalData) {
  var lottoData = [
    ["Lotto Tickets"],
    ["TOTAL VALUE", "NUMBERS, NAME", "ADDRESS", "SELLER", "DATE"]
  ];
  var totalRows = 0;
  var newRows = 0;
  var lottoRows = 0;
  var refundedRows = 0;
  var bingoRows = 0;
  var unknownRows = 0;

  var seller = "LF";
  var date = formatDate(new Date());

  var prevLottoOrderNumber;
  var setupLotto = true;

  var setupLottoTickets = true;
  var firstLottoTicketsTime = "?";
  var lastLottoTicketsTime = "?";

  for (var r = 1; r < originalData.length - 1; r++) {
    totalRows++;

    if (setupLotto) {
      setupLotto = false;
      prevLottoOrderNumber = parseInt(getDocumentProperty("prevLottoOrderNumber"), 10) || Number.MIN_VALUE;
      setDocumentProperty("prevLottoOrderNumber", originalData[r][COLUMNS.ORDER_NUMBER - 1]);
    }

    var orderNumber = parseInt(originalData[r][COLUMNS.ORDER_NUMBER - 1], 10);

    if (orderNumber > prevLottoOrderNumber) {
      newRows++;

      if (originalData[r][COLUMNS.ORDER_PLAN - 1].includes("Community lotto")) {
        if (originalData[r][COLUMNS.ORDER_STATUS - 1] == "Paid") {
          lottoRows++;

          var ticketTime = originalData[r][COLUMNS.ORDER_TIME - 1];
          if (setupLottoTickets) {
            setupLottoTickets = false;
            lastLottoTicketsTime = ticketTime;
          }
          firstLottoTicketsTime = ticketTime;

          var newRow = [];

          var ticketQuantity = originalData[r][COLUMNS.ORDER_QUANTITY - 1];
          var ticketValue = originalData[r][COLUMNS.ORDER_PLAN - 1].split(" - ")[1];

          if (ticketQuantity > 1) {
            newRow.push(ticketQuantity + " x " + ticketValue);
          }
          else {
            newRow.push(ticketValue);
          }

          newRow.push(originalData[r][COLUMNS.ORDER_COMMENTS - 1]);
          newRow.push(originalData[r][COLUMNS.ORDER_EMAIL - 1]);
          newRow.push(seller);
          newRow.push(date);

          lottoData.push(newRow);
        }

        else {
          refundedRows++;
        }
      }

      else {
        if (originalData[r][COLUMNS.ORDER_PLAN - 1].includes("Bingo")) {
          bingoRows++;
        }

        else {
          unknownRows++;
        }
      }
    }
  }

  lottoData[0][0] += " (" + firstLottoTicketsTime + " - " + lastLottoTicketsTime + ")"

  setDocumentProperty("totalRows", totalRows);
  setDocumentProperty("newRows", newRows);
  setDocumentProperty("lottoRows", lottoRows);
  setDocumentProperty("refundedRows", refundedRows);
  setDocumentProperty("bingoRows", bingoRows);
  setDocumentProperty("unknownRows", unknownRows);

  return normalizeDataArrays(lottoData);
}

function extractLottoDataFromAPI(originalData) {
  var lottoData = [
    ["Lotto Tickets"],
    ["TOTAL VALUE", "NUMBERS, NAME", "ADDRESS", "SELLER", "DATE"]
  ];
  var totalRows = 0;
  var newRows = 0;
  var lottoRows = 0;
  var refundedRows = 0;
  var bingoRows = 0;
  var unknownRows = 0;

  var seller = "LF";
  var date = formatDate(new Date());

  var prevLottoOrderNumber;
  var setupLotto = true;

  var setupLottoTickets = true;
  var firstLottoTicketsTime = "?";
  var lastLottoTicketsTime = "?";

  for (var or = 0; or < originalData.count; or++) {
    for (var r = 0; r < originalData.items[or].items.length; r++) {
      totalRows++;

      if (setupLotto) {
        setupLotto = false;
        prevLottoOrderNumber = parseInt(getDocumentProperty("prevLottoOrderNumber"), 10) || Number.MIN_VALUE;
        setDocumentProperty("prevLottoOrderNumber", originalData.items[or].orderNumber);
      }

      var orderNumber = parseInt(originalData.items[or].orderNumber, 10);

      if (orderNumber > prevLottoOrderNumber) {
        newRows++;

        if (originalData.items[or].items[r].name.includes("Community lotto")) {
          if (originalData.items[or].paymentStatus == "PAID") {
            lottoRows++;

            var ticketTime = formatDateTime(originalData.items[or].createDate);
            if (setupLottoTickets) {
              setupLottoTickets = false;
              lastLottoTicketsTime = ticketTime;
            }
            firstLottoTicketsTime = ticketTime;

            var newRow = [];

            var ticketQuantity = originalData.items[or].items[r].quantity;
            var ticketValue = originalData.items[or].items[r].name.split(" - ")[1];

            if (ticketQuantity > 1) {
              newRow.push(ticketQuantity + " x " + ticketValue);
            }
            else {
              newRow.push(ticketValue);
            }

            newRow.push(originalData.items[or].orderComments);
            newRow.push(originalData.items[or].email);
            newRow.push(seller);
            newRow.push(date);

            lottoData.push(newRow);
          }

          else {
            refundedRows++;
          }
        }

        else {
          if (originalData.items[or].items[r].name.includes("Bingo")) {
            bingoRows++;
          }

          else {
            unknownRows++;
          }
        }
      }
    }
  }

  lottoData[0][0] += " (" + firstLottoTicketsTime + " - " + lastLottoTicketsTime + ")"

  setDocumentProperty("totalRows", totalRows);
  setDocumentProperty("newRows", newRows);
  setDocumentProperty("lottoRows", lottoRows);
  setDocumentProperty("refundedRows", refundedRows);
  setDocumentProperty("bingoRows", bingoRows);
  setDocumentProperty("unknownRows", unknownRows);

  return normalizeDataArrays(lottoData);
}

function alertDataImportSuccess(dataType) {
  var totalRows = parseInt(getDocumentProperty("totalRows"), 10);
  var newRows = parseInt(getDocumentProperty("newRows"), 10);
  var bingoBooksRows = parseInt(getDocumentProperty("bingoBooksRows"), 10);
  var bingoVouchersRows = parseInt(getDocumentProperty("bingoVouchersRows"), 10);
  var bingoRows = parseInt(getDocumentProperty("bingoRows"), 10);
  var refundedRows = parseInt(getDocumentProperty("refundedRows"), 10);
  var lottoRows = parseInt(getDocumentProperty("lottoRows"), 10);
  var unknownRows = parseInt(getDocumentProperty("unknownRows"), 10);
  
  var alertMessage = dataType + ' has been successfully imported!\n\nTotal found entries: ' + totalRows +
                     '\nNew: ' + newRows + '\n\n';
  if (dataType == "Bingo Data") {
    alertMessage += 'Bingo orders:\nRelevant: ' + bingoRows + ' (' + bingoBooksRows + ' Bingo Books, ' +
                    bingoVouchersRows + ' Bingo Vouchers)\nRefunded: ' + refundedRows +
                    '\n\nOther orders:\nLotto Tickets: ' + lottoRows + '\nUnidentified: ' + unknownRows;
  }
  else if (dataType == "Lotto Data") {
    alertMessage += 'Lotto orders:\nRelevant: ' + lottoRows + '\nRefunded: ' + refundedRows +
                    '\n\nOther orders:\nBingo: ' + bingoRows + '\nUnidentified: ' + unknownRows;
  }

  SpreadsheetApp.getUi()
    .alert(alertMessage);
}

function generateBingoList() {
  var dataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Bingo Data");
  var bingoData = dataSheet.getDataRange().getValues();
  var listData = [];

  if (bingoData != "") {
    var bingoBooksData = [];
    for (var i = 2; i < bingoData.length; i++) {
      if (bingoData[i].join("") == "") {
        break;
      }
      bingoBooksData.push(bingoData[i]);
    }

    listData.push(["NUMBER", "COMMENT", "NAME", "SURNAME", "EMAIL", "REGISTRATION PLAN"]);
    
    for (var r = 0; r < bingoBooksData.length; r++) {
      var newRow = [];

      newRow.push(bingoBooksData[r][0]); // Number
      newRow.push(bingoBooksData[r][1]); // Comment

      var [name, surname] = extractNameAndSurname(bingoBooksData[r][1]);
      newRow.push(name, surname); // Name and Surname

      newRow.push(bingoBooksData[r][2]); // Email
      newRow.push(bingoBooksData[r][3]); // Registration plan

      for(var n = 1 ; n <= Number(bingoBooksData[r][5]); n++){
        listData.push(newRow);
      }
    }
  }

  var listSheet = createSheet("Bingo List", listData);
  listSheet.protect();
  formatListSheet(listSheet);
}

function formatListSheet(listSheet) {
  var listData = listSheet.getDataRange().getValues();
  if (listData != "") {
    listSheet.getRange(1, 1, 1, listData[0].length)
      .setFontWeight("bold"); // Set bold font for column names
    listSheet.getRange(1, 1, 1, listData[0].length)
      .setHorizontalAlignment("center"); // Set centered alignment for column names
    listSheet.getRange(1, 3, 1, listData[0].length - 2)
      .setBorder(false, false, true, false, false, false, '#000000', SpreadsheetApp.BorderStyle.DOTTED); // Set dotted border for Bingo list part
    listSheet.getRange(2, 1, listData.length - 1, 2)
      .setBorder(false, false, false, true, false, false, '#000000', SpreadsheetApp.BorderStyle.DOTTED); // Set dotted border for Bingo list part
    resizeColumns(listSheet, listData); // Set column size fit to data

    for (var i = 1; i < listData.length; i++) {
      var name = listData[i][2];
      if (!name) {
        var nameCell = listSheet.getRange(i + 1, 3);
        nameCell.setValue("name"); // Set default name for missing values
        nameCell.setFontStyle("italic"); // Set italic foont foor missing values
        nameCell.setFontColor("red"); // Set red font color for missing values

        var surnameCell = listSheet.getRange(i + 1, 4);
        surnameCell.setValue("surname"); // Set default name for missing values
        surnameCell.setFontStyle("italic"); // Set italic foont foor missing values
        surnameCell.setFontColor("red"); // Set red font color for missing values

      }
      var email = listData[i][4];
      if (!checkEmailDomain(email)) {
        var emailCell = listSheet.getRange(i + 1, 5);
        emailCell.setBackground("yellow") // Set yellow background for unknown emails domain addresses
      }
    }
  }
}

function clearBingoData() {
  deleteDocumentProperty("prevBingoOrderNumber");

  deleteSheet("Bingo Data");
  deleteSheet("Bingo List");
}

function clearLottoData() {
  deleteDocumentProperty("prevLottoOrderNumber");

  deleteSheet("Lotto Data");
}