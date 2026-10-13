/**
 * =========================================================================
 * MK EYE CENTER — PRODUCTION GOOGLE APPS SCRIPT DATABASE INTEGRATION
 * Primary Persistent Google Sheets Data Store Integration
 * =========================================================================
 * 
 * Authorized Spreadsheet IDs & Workspace Mapping:
 * - Reception Workspace: 11rBhezcklOndIPa7YAHr3zIBh6WCXkp2235slmirnlI
 * - Doctor Workspace:    1kdrkBEP2Iok8iKBRvBQdHalIcK5bl3uMrKgB46iWh94
 * 
 * Security: HMAC-SHA256 cryptographic signature validation is strictly 
 * required for every request. No unauthenticated endpoints.
 * Non-Destructive Saves: Strict validation, no silent truncation, no wholesale 
 * clearing of sheets, explicit cell clearing support, and exact coordinate validation.
 */

var ALLOWED_SPREADSHEETS = {
  "reception": "11rBhezcklOndIPa7YAHr3zIBh6WCXkp2235slmirnlI",
  "doctor": "1kdrkBEP2Iok8iKBRvBQdHalIcK5bl3uMrKgB46iWh94"
};

var MAX_REQUEST_AGE_MS = 5 * 60 * 1000; // 5 minutes drift protection

/**
 * Enforces strict workspace-to-spreadsheet mapping.
 * Reception must map only to Reception spreadsheet ID; Doctor only to Doctor ID.
 * Rejects missing or invalid workspace values explicitly.
 */
function validateWorkspaceMapping(workspace, spreadsheetId) {
  var wsKey = String(workspace || '').trim().toLowerCase();
  var targetId = String(spreadsheetId || '').trim();

  if (wsKey !== 'reception' && wsKey !== 'doctor') {
    throw new Error("Unauthorized: Invalid or missing workspace specified ('" + workspace + "').");
  }

  var expectedId = ALLOWED_SPREADSHEETS[wsKey];
  if (!targetId || targetId !== expectedId) {
    throw new Error("Unauthorized: Spreadsheet ID (" + targetId + ") does not match permitted ID for workspace '" + wsKey + "'.");
  }

  return SpreadsheetApp.openById(targetId);
}

function getApiSecret() {
  var props = PropertiesService.getScriptProperties();
  return props.getProperty('API_SECRET') || '';
}

/**
 * HMAC-SHA256 SIGNATURE VERIFICATION
 * 
 * Comparison with Node.js signing (/lib/sheets.ts):
 * Node.js:
 *   const bodyData = { action, timestamp, workspace, spreadsheetId, ...extraData };
 *   const messageToSign = `${timestamp}:${JSON.stringify(bodyData)}`;
 *   const signature = crypto.createHmac('sha256', secret).update(messageToSign).digest('hex');
 * 
 * Apps Script:
 *   - Copies payload omitting 'signature'.
 *   - messageToVerify = payload.timestamp + ':' + JSON.stringify(payloadCopy);
 *   - rawSignature = Utilities.computeHmacSha256Signature(messageToVerify, secret);
 *   - Converted to lowercase hex string and compared with payload.signature.
 */
function verifyRequestSecurity(payload) {
  var secret = getApiSecret();
  
  if (!secret) {
    return {
      valid: false,
      error: 'Security alert: API_SECRET is not configured in Google Apps Script Script Properties.'
    };
  }

  if (!payload || !payload.timestamp || !payload.signature) {
    return {
      valid: false,
      error: 'Unauthorized: Missing request timestamp or cryptographic signature.'
    };
  }

  var now = new Date().getTime();
  var reqTime = Number(payload.timestamp);
  if (isNaN(reqTime) || Math.abs(now - reqTime) > MAX_REQUEST_AGE_MS) {
    return {
      valid: false,
      error: 'Unauthorized: Request timestamp is invalid or expired (> 5 minutes drift).'
    };
  }

  var payloadCopy = {};
  for (var key in payload) {
    if (key !== 'signature') {
      payloadCopy[key] = payload[key];
    }
  }
  var messageToVerify = payload.timestamp + ':' + JSON.stringify(payloadCopy);

  var rawSignature = Utilities.computeHmacSha256Signature(messageToVerify, secret);
  var hexSignature = rawSignature.map(function(byte) {
    var v = (byte < 0 ? byte + 256 : byte).toString(16);
    return v.length === 1 ? '0' + v : v;
  }).join('');

  if (hexSignature.toLowerCase() !== String(payload.signature).toLowerCase()) {
    return {
      valid: false,
      error: 'Unauthorized: Cryptographic HMAC-SHA256 signature mismatch.'
    };
  }

  return { valid: true };
}

function createJsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function formatCellValue(val) {
  if (val === null || val === undefined) {
    return '';
  }
  if (val instanceof Date) {
    var year = val.getFullYear();
    var month = ('0' + (val.getMonth() + 1)).slice(-2);
    var day = ('0' + val.getDate()).slice(-2);
    var hours = val.getHours();
    var minutes = val.getMinutes();
    var seconds = val.getSeconds();
    
    if (hours === 0 && minutes === 0 && seconds === 0) {
      return year + '-' + month + '-' + day;
    }
    return val.toISOString();
  }
  return val;
}

function readWorksheetData(sheet) {
  var name = sheet.getName();
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();

  if (lastRow < 1 || lastColumn < 1) {
    return { name: name, headers: [], rows: [] };
  }

  var dataRange = sheet.getRange(1, 1, lastRow, lastColumn);
  var rawValues = dataRange.getValues();

  var headers = [];
  if (rawValues.length > 0) {
    headers = rawValues[0].map(function(cell) {
      return cell !== null && cell !== undefined ? String(cell).trim() : '';
    });
  }

  var rows = [];
  for (var r = 1; r < rawValues.length; r++) {
    rows.push(rawValues[r].map(formatCellValue));
  }

  return { name: name, headers: headers, rows: rows };
}

function doGet(e) {
  return createJsonResponse({
    status: 'success',
    service: 'MK Eye Center Secure Apps Script API',
    message: 'GET requests are active. All authenticated data queries and mutations must be sent via POST with valid HMAC-SHA256 signatures.',
    timestamp: new Date().toISOString()
  });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  var acquiredLock = false;

  try {
    if (!e || !e.postData || !e.postData.contents) {
      return createJsonResponse({
        status: 'error',
        message: 'No POST data payload received.'
      });
    }

    var payload = JSON.parse(e.postData.contents);

    var secCheck = verifyRequestSecurity(payload);
    if (!secCheck.valid) {
      return createJsonResponse({
        status: 'error',
        code: 'UNAUTHORIZED',
        message: secCheck.error
      });
    }

    var workspace = payload.workspace;
    var ss = validateWorkspaceMapping(workspace, payload.spreadsheetId);
    var action = payload.action;

    // 1. PING
    if (action === 'PING') {
      var sheetNames = ss.getSheets().map(function(s) { return s.getName(); });
      return createJsonResponse({
        status: 'success',
        service: 'MK Eye Center API',
        spreadsheetId: ss.getId(),
        sheets: sheetNames,
        timestamp: new Date().toISOString()
      });
    }

    // 2. GET_WORKBOOK
    if (action === 'GET_WORKBOOK') {
      var allSheets = ss.getSheets();
      var workbookData = {};

      for (var i = 0; i < allSheets.length; i++) {
        var currentSheet = allSheets[i];
        var sheetObj = readWorksheetData(currentSheet);
        workbookData[sheetObj.name] = sheetObj;
      }

      return createJsonResponse({
        status: 'success',
        sheets: workbookData,
        sheetNames: Object.keys(workbookData),
        timestamp: new Date().toISOString()
      });
    }

    // 3. GET_SHEET
    if (action === 'GET_SHEET') {
      var requestedSheetName = payload.sheet;
      if (!requestedSheetName) {
        return createJsonResponse({
          status: 'error',
          message: 'Parameter "sheet" is required.'
        });
      }

      var targetSheet = ss.getSheetByName(requestedSheetName);
      if (!targetSheet) {
        return createJsonResponse({
          status: 'error',
          message: 'Worksheet "' + requestedSheetName + '" does not exist.'
        });
      }

      var singleSheetData = readWorksheetData(targetSheet);
      return createJsonResponse({
        status: 'success',
        sheet: singleSheetData,
        timestamp: new Date().toISOString()
      });
    }

    acquiredLock = lock.tryLock(30000);
    if (!acquiredLock) {
      return createJsonResponse({
        status: 'error',
        message: 'Could not acquire database lock. Another transaction is currently writing. Please retry.'
      });
    }

    // 4. SAVE_WORKBOOK
    if (action === 'SAVE_WORKBOOK') {
      var sheetsToSave = payload.sheets;
      if (!sheetsToSave || typeof sheetsToSave !== 'object') {
        return createJsonResponse({
          status: 'error',
          message: 'Parameter "sheets" must be a valid mapping of sheet names to sheet definitions.'
        });
      }

      for (var sName in sheetsToSave) {
        if (!sheetsToSave.hasOwnProperty(sName)) continue;
        
        var sheetInput = sheetsToSave[sName];
        var sObj = ss.getSheetByName(sName);
        if (!sObj) {
          sObj = ss.insertSheet(sName);
        }

        saveWorksheetDataNonDestructive(sObj, sheetInput);
      }

      SpreadsheetApp.flush();

      return createJsonResponse({
        status: 'success',
        message: 'Workbook saved successfully to Google Sheets (non-destructive overlay).',
        timestamp: new Date().toISOString()
      });
    }

    // 5. UPDATE_SHEET
    if (action === 'UPDATE_SHEET') {
      var sheetName = payload.sheet;
      var sheetInput = {
        headers: payload.headers,
        rows: payload.rows
      };

      if (!sheetName || (!payload.rows && !payload.headers)) {
        return createJsonResponse({
          status: 'error',
          message: 'Parameters "sheet" and "rows" or "headers" are required.'
        });
      }

      var ws = ss.getSheetByName(sheetName);
      if (!ws) {
        ws = ss.insertSheet(sheetName);
      }

      saveWorksheetDataNonDestructive(ws, sheetInput);
      SpreadsheetApp.flush();

      return createJsonResponse({
        status: 'success',
        sheet: sheetName,
        message: 'Worksheet updated successfully (non-destructive overlay).',
        timestamp: new Date().toISOString()
      });
    }

    // 6. UPDATE_CELLS (Strict Coordinate & Integer Validation)
    if (action === 'UPDATE_CELLS') {
      var targetWsName = payload.sheet;
      var cellUpdates = payload.updates;

      if (!targetWsName || !Array.isArray(cellUpdates)) {
        return createJsonResponse({
          status: 'error',
          message: 'Parameters "sheet" and "updates" array are required.'
        });
      }

      var targetWs = ss.getSheetByName(targetWsName);
      if (!targetWs) {
        return createJsonResponse({
          status: 'error',
          message: 'Worksheet "' + targetWsName + '" not found.'
        });
      }

      // Validate every coordinate before writing (Number.isInteger and positive bounds)
      for (var v = 0; v < cellUpdates.length; v++) {
        var uItem = cellUpdates[v];
        var rCheck = uItem.row;
        var cCheck = uItem.col;
        if (typeof rCheck !== 'number' || !Number.isInteger(rCheck) || rCheck < 1 ||
            typeof cCheck !== 'number' || !Number.isInteger(cCheck) || cCheck < 1) {
          return createJsonResponse({
            status: 'error',
            message: 'Invalid cell coordinate at index ' + v + ': row=' + rCheck + ', col=' + cCheck + '. Must be positive integers >= 1.'
          });
        }
      }

      // Execute updates after full validation
      for (var u = 0; u < cellUpdates.length; u++) {
        var cell = cellUpdates[u];
        targetWs.getRange(cell.row, cell.col).setValue(cell.value !== undefined && cell.value !== null ? cell.value : '');
      }

      SpreadsheetApp.flush();

      return createJsonResponse({
        status: 'success',
        updatedCount: cellUpdates.length,
        message: 'Cell updates committed successfully after strict coordinate validation.',
        timestamp: new Date().toISOString()
      });
    }

    // 7. APPEND_ROWS
    if (action === 'APPEND_ROWS') {
      var appendSheetName = payload.sheet;
      var appendRows = payload.rows;

      if (!appendSheetName || !Array.isArray(appendRows) || appendRows.length === 0) {
        return createJsonResponse({
          status: 'error',
          message: 'Parameters "sheet" and non-empty "rows" array are required.'
        });
      }

      var appendWs = ss.getSheetByName(appendSheetName);
      if (!appendWs) {
        return createJsonResponse({
          status: 'error',
          message: 'Worksheet "' + appendSheetName + '" not found.'
        });
      }

      var startAppendRow = appendWs.getLastRow() + 1;
      var numAppendCols = appendRows[0].length;
      
      for (var ar = 0; ar < appendRows.length; ar++) {
        while (appendRows[ar].length < numAppendCols) {
          appendRows[ar].push('');
        }
      }

      appendWs.getRange(startAppendRow, 1, appendRows.length, numAppendCols).setValues(appendRows);
      SpreadsheetApp.flush();

      return createJsonResponse({
        status: 'success',
        appendedCount: appendRows.length,
        message: 'Appended rows committed successfully.',
        timestamp: new Date().toISOString()
      });
    }

    // 8. CREATE_WORKSHEET
    if (action === 'CREATE_WORKSHEET') {
      var newSheetName = payload.sheetName;
      if (!newSheetName || String(newSheetName).trim() === '') {
        return createJsonResponse({
          status: 'error',
          message: 'Worksheet name cannot be empty.'
        });
      }

      newSheetName = String(newSheetName).trim();

      if (ss.getSheetByName(newSheetName)) {
        return createJsonResponse({
          status: 'error',
          message: 'Worksheet "' + newSheetName + '" already exists in Google Sheets.'
        });
      }

      var createdWs = ss.insertSheet(newSheetName);

      if (payload.headers && Array.isArray(payload.headers) && payload.headers.length > 0) {
        createdWs.getRange(1, 1, 1, payload.headers.length).setValues([payload.headers]);
      }

      SpreadsheetApp.flush();

      return createJsonResponse({
        status: 'success',
        sheetName: newSheetName,
        message: 'Worksheet created successfully in Google Sheets.',
        timestamp: new Date().toISOString()
      });
    }

    return createJsonResponse({
      status: 'error',
      message: 'Unknown action requested: ' + String(action)
    });

  } catch (err) {
    return createJsonResponse({
      status: 'error',
      message: 'Server error processing request: ' + err.toString()
    });
  } finally {
    if (acquiredLock) {
      try {
        lock.releaseLock();
      } catch (lockErr) {}
    }
  }
}

/**
 * NON-DESTRUCTIVE CELL OVERLAY SAVE HELPER (Allows Intentional Cell Clearing)
 * - Distinguishes between "cell provided as empty string" (clears cell) and 
 *   "cell omitted/undefined" (preserves existing cell value).
 * - Preserves trailing rows not present in incoming payload.
 */
function saveWorksheetDataNonDestructive(sheet, sheetInput) {
  if (!sheetInput) return;

  var headers = null;
  var rows = null;

  if (Array.isArray(sheetInput)) {
    rows = sheetInput;
  } else if (typeof sheetInput === 'object') {
    headers = Array.isArray(sheetInput.headers) ? sheetInput.headers : null;
    rows = Array.isArray(sheetInput.rows) ? sheetInput.rows : null;
  }

  if (headers && headers.length > 0) {
    var maxHeaderCol = headers.length;
    var existingHeaderRange = sheet.getRange(1, 1, 1, maxHeaderCol);
    var existingHeaders = existingHeaderRange.getValues()[0];
    var mergedHeaders = [];
    for (var h = 0; h < maxHeaderCol; h++) {
      var incomingH = headers[h];
      var currentH = existingHeaders[h];
      mergedHeaders.push((incomingH !== undefined && incomingH !== null && String(incomingH).trim() !== '') ? incomingH : (currentH || ''));
    }
    existingHeaderRange.setValues([mergedHeaders]);
  }

  if (rows && Array.isArray(rows) && rows.length > 0) {
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();

    var maxCols = 1;
    for (var i = 0; i < rows.length; i++) {
      if (Array.isArray(rows[i]) && rows[i].length > maxCols) {
        maxCols = rows[i].length;
      }
    }
    maxCols = Math.max(maxCols, lastCol, 1);

    var existingValues = [];
    if (lastRow >= 2 && lastCol >= 1) {
      existingValues = sheet.getRange(2, 1, lastRow - 1, maxCols).getValues();
    }

    var finalMatrix = [];
    for (var r = 0; r < rows.length; r++) {
      var sourceRow = Array.isArray(rows[r]) ? rows[r] : [];
      var existingRow = (existingValues[r] && Array.isArray(existingValues[r])) ? existingValues[r] : [];
      var mergedRow = [];

      for (var c = 0; c < maxCols; c++) {
        var incomingVal = sourceRow[c];
        var existingVal = existingRow[c] !== undefined ? existingRow[c] : '';

        // Allow intentional cell clearing if incomingVal is explicitly "" (string) or null/empty string,
        // while preserving existingVal only if incomingVal is undefined.
        if (incomingVal !== undefined) {
          mergedRow.push(incomingVal);
        } else {
          mergedRow.push(existingVal);
        }
      }
      finalMatrix.push(mergedRow);
    }

    for (var tr = rows.length; tr < existingValues.length; tr++) {
      finalMatrix.push(existingValues[tr]);
    }

    if (finalMatrix.length > 0) {
      sheet.getRange(2, 1, finalMatrix.length, maxCols).setValues(finalMatrix);
    }
  }
}
