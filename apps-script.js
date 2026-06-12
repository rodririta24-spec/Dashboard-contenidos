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
  const sheet   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const all     = sheet.getDataRange().getValues();
  const rawHdrs = all[0];
  const headers = rawHdrs.map(h => String(h).trim());

  var richText = [];
  var formulas = [];
  var dataRows = all.length - 1;
  if (dataRows > 0) {
    var range = sheet.getRange(2, 1, dataRows, headers.length);
    richText = range.getRichTextValues();
    formulas = range.getFormulas();
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

        var url = null;

        // 1. Rich text cell-level link (Insert > Link on whole cell)
        if (richText[idx] && richText[idx][i]) {
          url = richText[idx][i].getLinkUrl();
          // 2. Link applied to a text run (part of the cell text)
          if (!url) {
            var runs = richText[idx][i].getRuns();
            for (var ri = 0; ri < runs.length; ri++) {
              url = runs[ri].getLinkUrl();
              if (url) break;
            }
          }
        }

        // 3. =HYPERLINK("url","text") formula
        if (!url && formulas[idx] && formulas[idx][i]) {
          var m = formulas[idx][i].match(/=HYPERLINK\s*\(\s*"([^"]+)"/i);
          if (m) url = m[1];
        }

        if (url) obj[h + '_url'] = url;
      });
      return obj;
    })
    .filter(Boolean);

  return ContentService
    .createTextOutput(JSON.stringify({ data: rows, headers: headers }))
    .setMimeType(ContentService.MimeType.JSON);
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
  var sheet   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  var lastRow = sheet.getLastRow();
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function(h){ return String(h).trim(); });
  Logger.log('Headers: ' + headers.join(' | '));
  Logger.log('Last row: ' + lastRow);

  var guionIdx = -1;
  headers.forEach(function(h, i) {
    var n = h.toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g,'');
    if (n === 'guion') guionIdx = i;
  });
  Logger.log('Guion col index: ' + guionIdx + ' (' + (headers[guionIdx]||'not found') + ')');
  if (guionIdx === -1) return;

  // Find first 5 rows that have a non-empty Titulo (actual data rows)
  var titleCol  = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  var dataRowNums = [];
  for (var i = 0; i < titleCol.length && dataRowNums.length < 5; i++) {
    if (titleCol[i][0]) dataRowNums.push(i + 2); // sheet row number
  }
  Logger.log('First data rows: ' + dataRowNums.join(', '));

  dataRowNums.forEach(function(rowNum) {
    var cell      = sheet.getRange(rowNum, guionIdx + 1);
    var value     = cell.getValue();
    var formula   = cell.getFormula();
    var rt        = cell.getRichTextValue();
    var cellUrl   = rt ? rt.getLinkUrl() : null;
    var runUrl    = null;
    if (rt && !cellUrl) {
      var runs = rt.getRuns();
      for (var r = 0; r < runs.length; r++) {
        runUrl = runs[r].getLinkUrl();
        if (runUrl) break;
      }
    }
    var fmMatch = formula ? formula.match(/=HYPERLINK\s*\(\s*"([^"]+)"/i) : null;
    // Check if it might be a smart chip (value empty but cell not blank)
    var displayValue = cell.getDisplayValue();
    Logger.log('Row ' + rowNum + ': value="' + value + '" display="' + displayValue + '" | cellUrl=' + cellUrl + ' | runUrl=' + runUrl + ' | formula="' + formula + '" | fmUrl=' + (fmMatch ? fmMatch[1] : null));
  });
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
