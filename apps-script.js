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

  // Build comments map from "Comentarios" sheet (keyed by ProjectRow)
  var commentsMap = {};
  try {
    var commSheet = spreadsheet.getSheetByName('Comentarios');
    if (commSheet) {
      var cAll = commSheet.getDataRange().getValues();
      if (cAll.length > 1) {
        var cHdrs = cAll[0].map(function(h){ return String(h).trim(); });
        cAll.slice(1).forEach(function(row) {
          var obj = {};
          cHdrs.forEach(function(h, i) {
            obj[h] = row[i] instanceof Date ? row[i].toISOString() : String(row[i] || '');
          });
          var pRow = obj['ProjectRow'];
          if (!commentsMap[pRow]) commentsMap[pRow] = [];
          commentsMap[pRow].push(obj);
        });
      }
    }
  } catch(ex) {
    Logger.log('Comments read error: ' + ex);
  }

  // Build history from "Historial" sheet (if it exists)
  var historyRows = [];
  try {
    var histSheet = spreadsheet.getSheetByName('Historial');
    if (histSheet) {
      var hAll = histSheet.getDataRange().getValues();
      if (hAll.length > 1) {
        var hHdrs = hAll[0].map(h => String(h).trim());
        historyRows = hAll.slice(1).reverse().slice(0, 50).map(function(row) {
          var obj = {};
          hHdrs.forEach(function(h, i) {
            obj[h] = row[i] instanceof Date ? row[i].toISOString() : String(row[i] || '');
          });
          return obj;
        });
      }
    }
  } catch(ex) {
    Logger.log('History read error: ' + ex);
  }

  return ContentService
    .createTextOutput(JSON.stringify({ data: rows, headers: headers, history: historyRows, comments: commentsMap }))
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
    const payload      = JSON.parse(e.postData.contents);
    const spreadsheet  = SpreadsheetApp.getActiveSpreadsheet();
    const sheet        = spreadsheet.getSheetByName(SHEET_NAME);
    const rawHdrs      = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const headers = rawHdrs.map(h => String(h).trim());

    if (payload.action === 'add') {
      const row = rawHdrs.map((_, i) => {
        const val = (payload.data[headers[i]] || '').toString().trim();
        return toCellValue(headers[i], val);
      });
      sheet.appendRow(row);
      var tituloAdd = payload.data['Titulo'] || payload.data[headers[0]] || '';
      logChange(spreadsheet, 'added', tituloAdd, '', '', sheet.getLastRow());
      return ok('add');
    }

    if (payload.action === 'update') {
      const rowNum = Number(payload.row);
      if (!rowNum || rowNum < 2) return fail('Número de fila inválido: ' + payload.row);
      var tituloCol = headers.indexOf('Titulo');
      var tituloUpd = tituloCol >= 0 ? String(sheet.getRange(rowNum, tituloCol + 1).getValue()) : '';
      Object.entries(payload.data).forEach(([key, val]) => {
        const col = headers.indexOf(key);
        if (col === -1) return;
        sheet.getRange(rowNum, col + 1).setValue(toCellValue(key, String(val)));
        logChange(spreadsheet, 'updated', tituloUpd, key, String(val), rowNum);
      });
      return ok('update');
    }

    if (payload.action === 'delete') {
      const rowNum = Number(payload.row);
      if (!rowNum || rowNum < 2) return fail('Número de fila inválido: ' + payload.row);
      var tituloDelCol = headers.indexOf('Titulo');
      var tituloDel = tituloDelCol >= 0 ? String(sheet.getRange(rowNum, tituloDelCol + 1).getValue()) : '';
      sheet.deleteRow(rowNum);
      logChange(spreadsheet, 'deleted', tituloDel, '', '', rowNum);
      return ok('delete');
    }

    if (payload.action === 'comment') {
      var cs = getOrCreateCommentsSheet(spreadsheet);
      cs.appendRow([new Date(), String(payload.row||''), String(payload.title||''), String(payload.author||''), String(payload.comment||'')]);
      return ok('comment');
    }

    return fail('Acción desconocida: ' + payload.action);

  } catch (err) {
    return fail(err.toString());
  }
}

// ── Helpers ─────────────────────────────────────────────────
function getOrCreateCommentsSheet(spreadsheet) {
  var s = spreadsheet.getSheetByName('Comentarios');
  if (!s) {
    s = spreadsheet.insertSheet('Comentarios');
    s.appendRow(['Timestamp','ProjectRow','ProjectTitle','Author','Comment']);
    s.setFrozenRows(1);
  }
  return s;
}

function getOrCreateHistSheet(spreadsheet) {
  var s = spreadsheet.getSheetByName('Historial');
  if (!s) {
    s = spreadsheet.insertSheet('Historial');
    s.appendRow(['Timestamp','Action','Titulo','Field','NewValue','RowNum']);
    s.setFrozenRows(1);
  }
  return s;
}

function logChange(spreadsheet, action, titulo, field, newVal, rowNum) {
  try {
    var s = getOrCreateHistSheet(spreadsheet);
    s.appendRow([new Date(), action, titulo || '', field || '', newVal || '', rowNum || '']);
  } catch(ex) {
    Logger.log('History write error: ' + ex);
  }
}

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
