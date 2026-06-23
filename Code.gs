// ═══════════════════════════════════════════════════════════════════════════
// AJIO • MYNTRA • FLIPKART — Sale & Purchase Processor
// Google Apps Script (Server-side)
// ═══════════════════════════════════════════════════════════════════════════

// ─── CONFIGURATION ──────────────────────────────────────────────────────────
// PASTE YOUR GOOGLE SHEET ID HERE (from the URL: docs.google.com/spreadsheets/d/THIS_PART/edit)
const SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID_HERE';

// ─── Web App Entry Point ────────────────────────────────────────────────────
function doGet(e) {
  if (e && e.parameter && e.parameter.action) {
    var action = e.parameter.action;
    var result;
    try {
      if (action === 'getSheetNames') {
        result = getSheetNames();
      } else if (action === 'getSheetData') {
        var startRow = parseInt(e.parameter.startRow || 3);
        var numRows = parseInt(e.parameter.numRows || 100);
        result = getSheetData(e.parameter.sheetName, startRow, numRows);
      } else if (action === 'searchSheetData') {
        var maxResults = parseInt(e.parameter.maxResults || 500);
        result = searchSheetData(e.parameter.sheetName, e.parameter.query || '', maxResults);
      } else if (action === 'clearSheetData') {
        result = clearSheetData(e.parameter.sheetName);
      } else if (action === 'getExistingUniqueIds') {
        result = getExistingUniqueIds(e.parameter.sheetName);
      } else {
        result = { success: false, error: 'Unknown action: ' + action };
      }
    } catch(err) {
      result = { success: false, error: err.message };
    }
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('AJIO • MYNTRA • FLIPKART — Sale & Purchase Processor')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function doPost(e) {
  var result;
  try {
    var contents = e.postData.contents;
    var params = JSON.parse(contents);
    var action = params.action;
    
    if (action === 'writeData') {
      var sheetName = params.sheetName;
      var saleRows = params.saleRows;
      var purchaseRows = params.purchaseRows;
      var isAppend = params.isAppend === undefined ? true : params.isAppend;
      result = writeDataToSheet(sheetName, saleRows, purchaseRows, isAppend);
    } else if (action === 'appendBatch') {
      result = appendBatch(params.sheetName, params.saleRows, params.purchaseRows, params.batchIndex, params.totalBatches);
    } else {
      result = { success: false, error: 'Unknown action: ' + action };
    }
  } catch(err) {
    result = { success: false, error: err.message };
  }
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─── Helper: Get Spreadsheet ────────────────────────────────────────────────
function getSpreadsheet() {
  if (!SPREADSHEET_ID || SPREADSHEET_ID === 'YOUR_SPREADSHEET_ID_HERE' || SPREADSHEET_ID.trim() === '') {
    return SpreadsheetApp.getActiveSpreadsheet();
  }
  try {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  } catch (e) {
    return SpreadsheetApp.getActiveSpreadsheet();
  }
}

// ─── Helper: Get Financial Year ─────────────────────────────────────────────
function getFinancialYear() {
  var now = new Date();
  if (now.getMonth() >= 3) { // April onwards (0-indexed, so 3 = April)
    return now.getFullYear() + '-' + String(now.getFullYear() + 1).slice(2);
  } else {
    return (now.getFullYear() - 1) + '-' + String(now.getFullYear()).slice(2);
  }
}

// ─── Get Sheet Names ────────────────────────────────────────────────────────
function getSheetNames() {
  try {
    var ss = getSpreadsheet();
    var sheets = ss.getSheets();
    var names = [];
    for (var i = 0; i < sheets.length; i++) {
      names.push({
        name: sheets[i].getName(),
        rows: sheets[i].getLastRow(),
        cols: sheets[i].getLastColumn()
      });
    }
    return { success: true, sheets: names, fy: getFinancialYear() };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ─── Get Sheet Data (Paginated) ─────────────────────────────────────────────
function getSheetData(sheetName, startRow, numRows) {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return { success: false, error: 'Sheet not found: ' + sheetName };
    
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    
    if (lastRow === 0 || lastCol === 0) {
      return { success: true, data: [], headers: [], totalRows: 0 };
    }
    
    // Get headers (row 1 and 2)
    var headerRange = sheet.getRange(1, 1, Math.min(2, lastRow), lastCol);
    var headers = headerRange.getValues();
    
    // Get data rows
    var dataStart = Math.max(startRow || 3, 3); // Data starts at row 3
    var rowCount = Math.min(numRows || 100, lastRow - dataStart + 1);
    
    if (dataStart > lastRow) {
      return { success: true, data: [], headers: headers, totalRows: lastRow - 2 };
    }
    
    var dataRange = sheet.getRange(dataStart, 1, rowCount, lastCol);
    var data = dataRange.getValues();
    
    return {
      success: true,
      headers: headers,
      data: data,
      totalRows: Math.max(0, lastRow - 2), // Minus 2 header rows
      startRow: dataStart,
      sheetName: sheetName
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ─── Get Existing UNIQUE IDs for Deduplication ──────────────────────────────
// Search across all data rows in the selected sheet, not only the current page.
function searchSheetData(sheetName, query, maxResults) {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return { success: false, error: 'Sheet not found: ' + sheetName };

    var q = String(query || '').trim();
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();

    if (lastRow === 0 || lastCol === 0) {
      return { success: true, data: [], headers: [], totalRows: 0, totalMatches: 0 };
    }

    var headers = sheet.getRange(1, 1, Math.min(2, lastRow), lastCol).getValues();
    if (!q) {
      return { success: true, data: [], headers: headers, totalRows: Math.max(0, lastRow - 2), totalMatches: 0 };
    }

    var limit = Math.min(Math.max(parseInt(maxResults || 500), 1), 2000);
    var dataRowCount = Math.max(0, lastRow - 2);
    if (dataRowCount === 0) {
      return { success: true, data: [], headers: headers, totalRows: 0, totalMatches: 0 };
    }

    var dataRange = sheet.getRange(3, 1, dataRowCount, lastCol);
    var matches = dataRange.createTextFinder(q).matchCase(false).findAll();
    var rowMap = {};

    for (var i = 0; i < matches.length; i++) {
      var rowNumber = matches[i].getRow();
      if (rowNumber >= 3) rowMap[rowNumber] = true;
    }

    var rowNumbers = Object.keys(rowMap).map(function(n) { return parseInt(n, 10); }).sort(function(a, b) { return a - b; });
    var data = [];
    var capped = rowNumbers.length > limit;

    for (var r = 0; r < rowNumbers.length && r < limit; r++) {
      data.push(sheet.getRange(rowNumbers[r], 1, 1, lastCol).getValues()[0]);
    }

    return {
      success: true,
      headers: headers,
      data: data,
      totalRows: dataRowCount,
      totalMatches: rowNumbers.length,
      returnedRows: data.length,
      limited: capped,
      sheetName: sheetName
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function getExistingUniqueIds(sheetName) {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return { success: true, saleIds: [], purchaseIds: [] };
    
    var lastRow = sheet.getLastRow();
    if (lastRow <= 2) return { success: true, saleIds: [], purchaseIds: [] };
    
    // SALE UNIQUE ID is in Column 27 (AA), PURCHASE UNIQUE ID is in Column 60 (BH)
    var saleIdCol = 27;   // Column AA
    var purchaseIdCol = 60; // Column BH
    var lastCol = sheet.getLastColumn();
    
    var saleIds = [];
    var purchaseIds = [];
    
    // Read in batches to avoid timeout
    var batchSize = 10000;
    for (var start = 3; start <= lastRow; start += batchSize) {
      var rows = Math.min(batchSize, lastRow - start + 1);
      
      if (saleIdCol <= lastCol) {
        var saleRange = sheet.getRange(start, saleIdCol, rows, 1);
        var saleValues = saleRange.getValues();
        for (var i = 0; i < saleValues.length; i++) {
          if (saleValues[i][0] !== '' && saleValues[i][0] !== null) {
            saleIds.push(String(saleValues[i][0]));
          }
        }
      }
      
      if (purchaseIdCol <= lastCol) {
        var purchaseRange = sheet.getRange(start, purchaseIdCol, rows, 1);
        var purchaseValues = purchaseRange.getValues();
        for (var i = 0; i < purchaseValues.length; i++) {
          if (purchaseValues[i][0] !== '' && purchaseValues[i][0] !== null) {
            purchaseIds.push(String(purchaseValues[i][0]));
          }
        }
      }
    }
    
    return { success: true, saleIds: saleIds, purchaseIds: purchaseIds };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ─── Write Data to Sheet (Batched) ──────────────────────────────────────────
function writeDataToSheet(sheetName, saleRows, purchaseRows, isAppend) {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName(sheetName);
    
    // Create sheet if not exists
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }
    
    var startRow;
    
    if (!isAppend || sheet.getLastRow() <= 2) {
      // Fresh write: Clear and write headers first
      sheet.clear();
      
      // Auto-clean: Delete extra columns beyond 70 to save cell limits
      var maxCols = sheet.getMaxColumns();
      if (maxCols > 70) {
        try {
          sheet.deleteColumns(71, maxCols - 70);
        } catch(e) {}
      }
      
      // Row 1: Merged headers
      sheet.getRange('A1').setValue('SALE');
      sheet.getRange('A1').setFontWeight('bold').setFontSize(12)
        .setBackground('#1B5E20').setFontColor('#FFFFFF')
        .setHorizontalAlignment('center');
      sheet.getRange('A1:AL1').merge();
      
      sheet.getRange('AN1').setValue('PURCHASE');
      sheet.getRange('AN1').setFontWeight('bold').setFontSize(12)
        .setBackground('#0D47A1').setFontColor('#FFFFFF')
        .setHorizontalAlignment('center');
      sheet.getRange('AN1:BR1').merge();
      
      // Row 2: Column headers
      var saleHeaders = [
        'No.', 'Invoice No', 'TYPE', 'Invoice Date', 'Warehouse Name',
        'Warehouse Code', 'GST No', 'Order ID', 'Item Asin', 'Item SKU',
        'Item Name', 'HSN Number', 'Quantity', 'Item Cost', 'Gross',
        'IGST', 'CGST', 'SGST', 'IGST Amt', 'CGST Amt', 'SGST Amt',
        'Invoice', 'Reason', 'Zoho Status', 'invoice id', 'State Code',
        'SALE UNIQUE ID', 'Quantity', 'Item Cost', 'Gross',
        'IGST', 'CGST', 'SGST', 'IGST Amt', 'CGST Amt', 'SGST Amt',
        'Invoice', 'STATE CODE'
      ];
      
      var purchaseHeaders = [
        'No.', 'Invoice No', 'Warehouse Name', 'Warehouse Code', 'GST No',
        'Order ID', 'Item Asin', 'Item SKU', 'Item Name', 'HSN Number',
        'Quantity', 'Item cost', 'Gross', 'IGST', 'CGST', 'SGST',
        'IGST Amt', 'CGST Amt', 'SGST Amt', 'Invoice',
        'PURCHASE UNIQUE ID', 'Quantity', 'Item cost', 'Gross',
        'IGST', 'CGST', 'SGST', 'IGST Amt', 'CGST Amt', 'SGST Amt',
        'Total Amt'
      ];
      
      // Write SALE headers (Col 1-38)
      if (saleHeaders.length > 0) {
        sheet.getRange(2, 1, 1, saleHeaders.length).setValues([saleHeaders]);
        sheet.getRange(2, 1, 1, saleHeaders.length)
          .setFontWeight('bold')
          .setBackground('#C8E6C9')
          .setHorizontalAlignment('center');
      }
      
      // Write PURCHASE headers (Col 40-70)
      if (purchaseHeaders.length > 0) {
        sheet.getRange(2, 40, 1, purchaseHeaders.length).setValues([purchaseHeaders]);
        sheet.getRange(2, 40, 1, purchaseHeaders.length)
          .setFontWeight('bold')
          .setBackground('#BBDEFB')
          .setHorizontalAlignment('center');
      }
      
      startRow = 3;
    } else {
      // Append mode: Find next empty row
      startRow = sheet.getLastRow() + 1;
    }
    
    // Write SALE data
    if (saleRows && saleRows.length > 0) {
      var batchSize = 5000;
      for (var i = 0; i < saleRows.length; i += batchSize) {
        var batch = saleRows.slice(i, i + batchSize);
        var numCols = batch[0].length;
        sheet.getRange(startRow + i, 1, batch.length, numCols).setValues(batch);
      }
    }
    
    // Write PURCHASE data
    if (purchaseRows && purchaseRows.length > 0) {
      var batchSize = 5000;
      for (var i = 0; i < purchaseRows.length; i += batchSize) {
        var batch = purchaseRows.slice(i, i + batchSize);
        var numCols = batch[0].length;
        // Purchase starts at column 40
        sheet.getRange(startRow + i, 40, batch.length, numCols).setValues(batch);
      }
    }
    
    // Trim extra blank rows at the bottom to prevent hitting 10M cell limit
    var lastRow = sheet.getLastRow();
    var maxRows = sheet.getMaxRows();
    if (maxRows > lastRow && lastRow > 2) {
      try {
        sheet.deleteRows(lastRow + 1, maxRows - lastRow);
      } catch(e) {}
    }
    
    // Auto-resize some columns
    try {
      sheet.autoResizeColumn(2);  // Invoice No
      sheet.autoResizeColumn(11); // Item Name
      sheet.autoResizeColumn(48); // Purchase Item Name
    } catch(e) {}
    
    return {
      success: true,
      message: 'Data written successfully!',
      saleRows: saleRows ? saleRows.length : 0,
      purchaseRows: purchaseRows ? purchaseRows.length : 0,
      sheetName: sheetName
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function appendBatch(sheetName, saleRows, purchaseRows, batchIndex, totalBatches) {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return { success: false, error: 'Sheet not found' };
    
    var lastRow = sheet.getLastRow();
    
    // If sheet is empty or only has headers (or no headers), write headers first
    if (lastRow <= 2) {
      sheet.clear();
      
      // Auto-clean columns
      var maxCols = sheet.getMaxColumns();
      if (maxCols > 70) {
        try { sheet.deleteColumns(71, maxCols - 70); } catch(e) {}
      }
      
      // Row 1: Merged headers
      sheet.getRange('A1').setValue('SALE');
      sheet.getRange('A1').setFontWeight('bold').setFontSize(12)
        .setBackground('#1B5E20').setFontColor('#FFFFFF')
        .setHorizontalAlignment('center');
      sheet.getRange('A1:AL1').merge();
      
      sheet.getRange('AN1').setValue('PURCHASE');
      sheet.getRange('AN1').setFontWeight('bold').setFontSize(12)
        .setBackground('#0D47A1').setFontColor('#FFFFFF')
        .setHorizontalAlignment('center');
      sheet.getRange('AN1:BR1').merge();
      
      // Row 2: Column headers
      var saleHeaders = [
        'No.', 'Invoice No', 'TYPE', 'Invoice Date', 'Warehouse Name',
        'Warehouse Code', 'GST No', 'Order ID', 'Item Asin', 'Item SKU',
        'Item Name', 'HSN Number', 'Quantity', 'Item Cost', 'Gross',
        'IGST', 'CGST', 'SGST', 'IGST Amt', 'CGST Amt', 'SGST Amt',
        'Invoice', 'Reason', 'Zoho Status', 'invoice id', 'State Code',
        'SALE UNIQUE ID', 'Quantity', 'Item Cost', 'Gross',
        'IGST', 'CGST', 'SGST', 'IGST Amt', 'CGST Amt', 'SGST Amt',
        'Invoice', 'STATE CODE'
      ];
      
      var purchaseHeaders = [
        'No.', 'Invoice No', 'Warehouse Name', 'Warehouse Code', 'GST No',
        'Order ID', 'Item Asin', 'Item SKU', 'Item Name', 'HSN Number',
        'Quantity', 'Item cost', 'Gross', 'IGST', 'CGST', 'SGST',
        'IGST Amt', 'CGST Amt', 'SGST Amt', 'Invoice',
        'PURCHASE UNIQUE ID', 'Quantity', 'Item cost', 'Gross',
        'IGST', 'CGST', 'SGST', 'IGST Amt', 'CGST Amt', 'SGST Amt',
        'Total Amt'
      ];
      
      if (saleHeaders.length > 0) {
        sheet.getRange(2, 1, 1, saleHeaders.length).setValues([saleHeaders]);
        sheet.getRange(2, 1, 1, saleHeaders.length)
          .setFontWeight('bold')
          .setBackground('#C8E6C9')
          .setHorizontalAlignment('center');
      }
      
      if (purchaseHeaders.length > 0) {
        sheet.getRange(2, 40, 1, purchaseHeaders.length).setValues([purchaseHeaders]);
        sheet.getRange(2, 40, 1, purchaseHeaders.length)
          .setFontWeight('bold')
          .setBackground('#BBDEFB')
          .setHorizontalAlignment('center');
      }
      
      lastRow = 2;
    }
    
    var startRow = lastRow + 1;
    
    if (saleRows && saleRows.length > 0) {
      sheet.getRange(startRow, 1, saleRows.length, saleRows[0].length).setValues(saleRows);
    }
    
    if (purchaseRows && purchaseRows.length > 0) {
      // Ensure we write at the same row position for purchase data
      var purchaseStartRow = startRow;
      if (saleRows && saleRows.length > 0) {
        purchaseStartRow = startRow; // Same row, different column
      }
      sheet.getRange(purchaseStartRow, 40, purchaseRows.length, purchaseRows[0].length).setValues(purchaseRows);
    }
    
    if (batchIndex === totalBatches) {
      var lastRow = sheet.getLastRow();
      var maxRows = sheet.getMaxRows();
      if (maxRows > lastRow && lastRow > 2) {
        try {
          sheet.deleteRows(lastRow + 1, maxRows - lastRow);
        } catch(e) {}
      }
    }
    
    return {
      success: true,
      batchIndex: batchIndex,
      totalBatches: totalBatches,
      rowsWritten: Math.max(saleRows ? saleRows.length : 0, purchaseRows ? purchaseRows.length : 0)
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ─── Clear Sheet Data (Keep Headers) ────────────────────────────────────────
function clearSheetData(sheetName) {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return { success: true, message: 'Sheet not found, nothing to clear' };
    
    var lastRow = sheet.getLastRow();
    if (lastRow > 2) {
      sheet.getRange(3, 1, lastRow - 2, sheet.getMaxColumns()).clear();
    }
    
    return { success: true, message: 'Sheet data cleared! Headers kept.' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ─── Get Spreadsheet URL ───────────────────────────────────────────────────
function getSpreadsheetUrl() {
  try {
    var ss = getSpreadsheet();
    return { success: true, url: ss.getUrl(), name: ss.getName() };
  } catch (e) {
    return { success: false, error: e.message };
  }
}
