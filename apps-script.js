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
//  5. Habilitá el servicio: Services (+) → Google Sheets API → Add
//  6. Clic en "Implementar" → "Nueva implementación"
//     · Tipo:         Aplicación web
//     · Ejecutar como: Yo (tu cuenta de Google)
//     · Acceso:        Cualquier persona
//  7. Autorizá los permisos cuando los pida Google
//  8. Copiá la URL del Web App que aparece al final
//  9. En el dashboard, hacé clic en el ícono ⚙️ y pegá esa URL
//
// Cada vez que modifiques este script, hacé una NUEVA implementación
// ============================================================

const SHEET_NAME = 'Reporte';

// ── GET: devuelve todos los datos como JSON ─────────────────
function doGet(e) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet     = spreadsheet.getSheetByName(SHEET_NAME);
  const all       = sheet.getDataRange().getValues();
  const rawHdrs   = all[0];
  const headers   = rawHdrs.map(h => String(h).trim());

  // Build a map of hyperlink URLs using Sheets API v4.
  // This covers: regular hyperlinks, Smart Chips, and text-run links.
  // Requires: Services → Google Sheets API (Advanced Service)
  var hyperlinkMap = {}; // hyperlinkMap[rowData0Index][colIndex] = url
  try {
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    var range   = SHEET_NAME + '!A1:' + colLetter(lastCol) + lastRow;
    var resp    = Sheets.Spreadsheets.get(spreadsheet.getId(), {
      ranges: [range],
      fields: [
        'sheets.data.rowData.values.hyperlink',
        'sheets.data.rowData.values.effectiveFormat.textFormat.link',
        'sheets.data.rowData.values.textFormatRuns'
      ].join(',')
    });
    var rowData = resp.sheets[0].data[0].rowData || [];
    rowData.forEach(function(row, ri) {
      if (!row || !row.values) return;
      row.values.forEach(function(cell, ci) {
        var url = null;

        // 1. Top-level hyperlink field (standard hyperlinks)
        if (cell.hyperlink) {
          url = cell.hyperlink;
        }

        // 2. Whole-cell text format link (some chip styles)
        if (!url && cell.effectiveFormat &&
            cell.effectiveFormat.textFormat &&
            cell.effectiveFormat.textFormat.link) {
          url = cell.effectiveFormat.textFormat.link.uri;
        }

        // 3. Text format runs (partial-text links, Smart Chips)
        if (!url && cell.textFormatRuns) {
          for (var r = 0; r < cell.textFormatRuns.length; r++) {
            var run = cell.textFormatRuns[r];
            if (run.format && run.format.link && run.format.link.uri) {
              url = run.format.link.uri;
              break;
            }
          }
        }

        if (url) {
          if (!hyperlinkMap[ri]) hyperlinkMap[ri] = {};
          hyperlinkMap[ri][ci] = url;
        }
      });
    });
  } catch (ex) {
    Logger.log('Sheets API error: ' + ex);
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
        // rowData is 0-indexed from row 1; idx is 0-indexed from row 2, so +1
        var apiRow = hyperlinkMap[idx + 1];
        if (apiRow && apiRow[i]) obj[h + '_url'] = apiRow[i];
      });
      return obj;
    })
    .filter(Boolean);

  // Read Estado dropdown options from data validation on the sheet
  var statusOptions = [];
  try {
    var estadoIdx = headers.indexOf('Estado');
    if (estadoIdx >= 0) {
      var rule = sheet.getRange(2, estadoIdx + 1).getDataValidation();
      if (rule && rule.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
        statusOptions = rule.getCriteriaValues()[0];
      }
    }
  } catch(ex) { Logger.log('DataValidation error: ' + ex); }

  return ContentService
    .createTextOutput(JSON.stringify({ data: rows, headers: headers, statuses: statusOptions }))
    .setMimeType(ContentService.MimeType.JSON);
}

function colLetter(n) {
  var s = '';
  while (n > 0) {
    s = String.fromCharCode(64 + (n - 1) % 26 + 1) + s;
    n = Math.floor((n - 1) / 26);
  }
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

    return fail('Acción desconocida: ' + payload.action);

  } catch (err) {
    return fail(err.toString());
  }
}

// ── Helpers ─────────────────────────────────────────────────
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
