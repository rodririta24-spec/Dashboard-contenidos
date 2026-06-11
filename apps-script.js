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

    return fail('Acción desconocida: ' + payload.action);

  } catch (err) {
    return fail(err.toString());
  }
}

// ── Helpers ─────────────────────────────────────────────────

// Convierte string a Date si el campo es de fecha
function toCellValue(key, val) {
  if (!val) return '';
  if (key.toLowerCase().includes('fecha') && /\d{1,2}\/\d{1,2}\/\d{4}/.test(val)) {
    const [d, m, y] = val.split('/').map(Number);
    return new Date(y, m - 1, d);
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
