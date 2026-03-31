/**
 * Google Apps Script – Seguimiento HO Plan 2026
 * Spreadsheet: 1cPeFZorUwiO3xXQmUwPhlV4Wy48Xg0xLOPV6xxoipBg  (mismo que TMERT)
 *
 * INSTRUCCIONES DE DESPLIEGUE:
 * 1. Abrir el spreadsheet en Google Sheets
 * 2. Extensions → Apps Script → pegar este código (puede estar en el mismo
 *    proyecto que SeguimientoTMERT_AppScript.gs, como archivo separado)
 * 3. Deploy → New deployment → Web app
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 4. Copiar la URL generada y pegarla en secrets.toml como ho_api.url
 *
 * NOTAS:
 * - Acepta parámetro "sheet" en el body del POST para escribir en hojas distintas.
 *   Si no se indica, usa SHEET_SEGUIMIENTO_HO por defecto.
 * - Hojas soportadas: "Seguimiento HO", "EP Detalle"
 */

const SPREADSHEET_ID_HO  = '1cPeFZorUwiO3xXQmUwPhlV4Wy48Xg0xLOPV6xxoipBg';
const SHEET_SEGUIMIENTO_HO = 'Seguimiento HO';
const API_KEY_HO           = 'ho_seguimiento_IST2026';

// ── GET ──────────────────────────────────────────────────────────────────────
function doGet(e) {
  try {
    const action = e.parameter.action;
    const key    = e.parameter.key;

    if (action === 'ping') {
      return json_ho({ success: true, message: 'HO API OK', ts: new Date().toString() });
    }

    if (key !== API_KEY_HO) {
      return json_ho({ success: false, error: 'Clave API inválida' });
    }

    if (action === 'getInfo') {
      const ss    = SpreadsheetApp.openById(SPREADSHEET_ID_HO);
      const sheet = getOrCreateSheet_HO(ss, SHEET_SEGUIMIENTO_HO);
      return json_ho({
        success:  true,
        rows:     Math.max(0, sheet.getLastRow() - 1),
        updated:  sheet.getRange(2, 1).getValue() || 'sin datos'
      });
    }

    return json_ho({ success: false, error: 'Acción no válida: ' + action });

  } catch (err) {
    return json_ho({ success: false, error: err.toString() });
  }
}

// ── POST ─────────────────────────────────────────────────────────────────────
function doPost(e) {
  try {
    const action = e.parameter.action;
    const key    = e.parameter.key;

    if (key !== API_KEY_HO) {
      return json_ho({ success: false, error: 'Clave API inválida' });
    }

    if (!e.postData || !e.postData.contents) {
      return json_ho({ success: false, error: 'Sin datos POST' });
    }

    const data = JSON.parse(e.postData.contents);

    switch (action) {
      case 'writeHeaders':
        return writeHeaders_HO(data);
      case 'appendRows':
        return appendRows_HO(data);
      default:
        return json_ho({ success: false, error: 'Acción no válida: ' + action });
    }

  } catch (err) {
    return json_ho({ success: false, error: err.toString(), stack: err.stack });
  }
}

// ── FUNCIONES INTERNAS ───────────────────────────────────────────────────────

function getOrCreateSheet_HO(ss, sheetName) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  return sheet;
}

function writeHeaders_HO(data) {
  const sheetName = data.sheet || SHEET_SEGUIMIENTO_HO;
  const headers   = data.headers;
  const ss        = SpreadsheetApp.openById(SPREADSHEET_ID_HO);
  const sheet     = getOrCreateSheet_HO(ss, sheetName);
  sheet.clearContents();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#4F0B7B');
  headerRange.setFontColor('#FFFFFF');
  return json_ho({ success: true, message: 'Cabecera escrita: ' + headers.length + ' columnas en ' + sheetName });
}

function appendRows_HO(data) {
  const sheetName = data.sheet || SHEET_SEGUIMIENTO_HO;
  const rows      = data.rows;
  if (!rows || rows.length === 0) {
    return json_ho({ success: true, message: 'Sin filas para agregar' });
  }
  const ss      = SpreadsheetApp.openById(SPREADSHEET_ID_HO);
  const sheet   = getOrCreateSheet_HO(ss, sheetName);
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1, rows.length, rows[0].length).setValues(rows);
  return json_ho({ success: true, rows_written: rows.length, sheet: sheetName });
}

function json_ho(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
