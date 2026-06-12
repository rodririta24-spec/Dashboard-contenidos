// ============================================================
// DASHBOARD DE CONTENIDOS — Google Apps Script Backend
// ============================================================
//
// CÓMO INSTALARLO (hacerlo una sola vez):
//
//  1. Abrí tu Google Sheet
//  2. Menú: Extensiones → Apps Script
//  3. Borrá todo el contenido de Code.gs y pegá ESTE código
//  4. Guardá (Ctrl+S)
//  5. Clic en "Implementar" → "Nueva implementación"
//     · Tipo:         Aplicación web
//     · Ejecutar como: Yo (tu cuenta de Google)
//     · Acceso:        Cualquier persona
//  6. Autorizá los permisos cuando los pida Google
//  7. Copiá la URL del Web App que aparece al final
//  8. En el dashboard, hacé clic en el ícono ⚙️ y pegá esa URL
//
// Cada vez que modifiques este script, hacé una NUEVA implementación
// (no actualices la existente) para que los cambios tomen efecto.
// ============================================================

const SHEET_NAME = 'Reporte';

// ── GET: devuelve todos los datos como JSON ─────────────────
function doGet(e) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet     = spreadsheet.getSheetByName(SHEET_NAME);
  const all       = sheet.getDataRange().getValues();
  const rawHdrs   = all[0];
  const headers   = rawHdrs.map(h => String(h).trim());

  // ── Hyperlink map via Sheets API v4 (reads Smart Chip URLs too) ──
  // Requires: Services → Google Sheets API (enabled in Apps Script)
  var hyperlinkMap = {}; // hyperlinkMap[sheetRowIndex0][colIndex] = url
  try {
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    var range   = SHEET_NAME + '!A1:' + colLetter(lastCol) + lastRow;
    var resp    = Sheets.Spreadsheets.get(spreadsheet.getId(), {
      ranges: [range],
      fields: 'sheets.data.rowData.values.hyperlink'
    });
    var rowData = resp.sheets[0].data[0].rowData || [];
    rowData.forEach(function(row, ri) {
      if (!row || !row.values) return;
      row.values.forEach(function(cell, ci) {
        if (cell && cell.hyperlink) {
          if (!hyperlinkMap[ri]) hyperlinkMap[ri] = {};
          hyperlinkMap[ri][ci] = cell.hyperlink;
        }
      });
    });
  } catch (ex) {
    Logger.log('Sheets API error (enable Google Sheets API in Services): ' + ex);
  }

  const rows = all.slice(1)
    .map((row, idx) => {
      if (!row[0]) return null;
      const obj = { _row: idx + 2 }; // número real de fila en el Sheet (1-based, +1 por encabezado)
      headers.forEach((h, i) => {
        let v = row[i];
        if (v instanceof Date) {
          v = `${v.getDate()}/${v.getMonth() + 1}/${v.getFullYear()}`;
        } else {
          v = (v !== null && v !== undefined) ? String(v) : '';
        }
        obj[h] = v;
        // Sheets API row index: idx+1 because rowData[0] = header row
        var apiRow = hyperlinkMap[idx + 1];
        if (apiRow && apiRow[i]) obj[h + '_url'] = apiRow[i];
      });
      return obj;
    })
    .filter(Boolean);

  return ContentService
    .createTextOutput(JSON.stringify({ data: rows, headers: headers }))
    .setMimeType(ContentService.MimeType.JSON);
}

function colLetter(n) {
  var s = '';
  while (n > 0) { s = String.fromCharCode(64 + (n - 1) % 26 + 1) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

// ── POST: agrega o edita una fila ───────────────────────────
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const sheet   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    const rawHdrs = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const headers = rawHdrs.map(h => String(h).trim());

    if (payload.action === 'add') {
      const row = rawHdrs.map((_, i) => {
        const val = (payload.data[headers[i]] || '').toString().trim();
        return toCellValue(headers[i], val);
      });
      sheet.appendRow(row);
      return ok('add');
    }

    if (payload.action === 'update') {
      const rowNum = Number(payload.row);
      if (!rowNum || rowNum < 2) return fail('Número de fila inválido: ' + payload.row);
      Object.entries(payload.data).forEach(([key, val]) => {
        const col = headers.indexOf(key);
        if (col === -1) return;
        sheet.getRange(rowNum, col + 1).setValue(toCellValue(key, String(val)));
      });
      return ok('update');
    }

    if (payload.action === 'delete') {
      const rowNum = Number(payload.row);
      if (!rowNum || rowNum < 2) return fail('Número de fila inválido: ' + payload.row);
      sheet.deleteRow(rowNum);
      return ok('delete');
    }

    if (payload.action === 'upload') {
      var fileName = payload.fileName || 'material';
      var mimeType = payload.mimeType || 'application/octet-stream';
      var folderId = payload.folderId || '';
      var b64data  = payload.data    || '';

      var folder;
      if (folderId) {
        folder = DriveApp.getFolderById(folderId);
      } else {
        var folderName = 'Dashboard Materials';
        var found = DriveApp.getFoldersByName(folderName);
        folder = found.hasNext() ? found.next() : DriveApp.createFolder(folderName);
      }

      var bytes     = Utilities.base64Decode(b64data);
      var blob      = Utilities.newBlob(bytes, mimeType, fileName);
      var driveFile = folder.createFile(blob);
      driveFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

      return ContentService
        .createTextOutput(JSON.stringify({ success: true, action: 'upload', url: driveFile.getUrl() }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    return fail('Acción desconocida: ' + payload.action);

  } catch (err) {
    return fail(err.toString());
  }
}

// ── TEST: corré esta función desde el editor para diagnosticar ─
function testLinks() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet       = spreadsheet.getSheetByName(SHEET_NAME);
  var lastRow     = sheet.getLastRow();
  var lastCol     = sheet.getLastColumn();
  var headers     = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h){ return String(h).trim(); });

  // Find Guion column
  var guionIdx = -1;
  headers.forEach(function(h, i) {
    if (h.toLowerCase().replace('ó','o') === 'guion') guionIdx = i;
  });
  Logger.log('Guion col: ' + guionIdx + ' (' + (headers[guionIdx]||'NOT FOUND') + ')');

  // Find first 3 data rows
  var titleVals = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  var dataRows = [];
  for (var i = 0; i < titleVals.length && dataRows.length < 3; i++) {
    if (titleVals[i][0]) dataRows.push(i + 2);
  }
  Logger.log('First data rows: ' + dataRows.join(', '));

  // ── Test 1: Sheets API v4 ──────────────────────────────────
  Logger.log('--- Testing Sheets API v4 ---');
  if (typeof Sheets === 'undefined') {
    Logger.log('ERROR: Google Sheets API NOT enabled! Go to: Services (+) → Google Sheets API → Add');
    Logger.log('Without it, Smart Chip URLs cannot be read.');
  } else {
    try {
      var range = SHEET_NAME + '!A1:' + colLetter(lastCol) + lastRow;
      var resp  = Sheets.Spreadsheets.get(spreadsheet.getId(), {
        ranges: [range],
        fields: 'sheets.data.rowData.values.hyperlink'
      });
      var rowData = resp.sheets[0].data[0].rowData || [];
      Logger.log('Sheets API OK — total rowData entries: ' + rowData.length);
      dataRows.forEach(function(rowNum) {
        var apiRow = rowData[rowNum - 1]; // 0-indexed
        var cell   = apiRow && apiRow.values ? apiRow.values[guionIdx] : null;
        Logger.log('Row ' + rowNum + ' Guion hyperlink: ' + (cell ? cell.hyperlink : 'null/no cell'));
      });
    } catch(ex) {
      Logger.log('Sheets API error: ' + ex.toString());
    }
  }
}

// ── Helpers ─────────────────────────────────────────────────

// Convierte string a Date si el campo es de fecha
function toCellValue(key, val) {
  if (!val) return '';
  if (key.toLowerCase().indexOf('fecha') !== -1) {
    var parts = val.split('/');
    if (parts.length === 3 && parts[2].length === 4) {
      var d = Number(parts[0]), m = Number(parts[1]), y = Number(parts[2]);
      if (d && m && y) return new Date(y, m - 1, d);
    }
  }
  return val;
}

function ok(action) {
  return ContentService
    .createTextOutput(JSON.stringify({ success: true, action: action }))
    .setMimeType(ContentService.MimeType.JSON);
}

function fail(msg) {
  return ContentService
    .createTextOutput(JSON.stringify({ success: false, error: msg }))
    .setMimeType(ContentService.MimeType.JSON);
}
