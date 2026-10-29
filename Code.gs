/**
 * Backend de la cata de turrones: guarda las calificaciones en un Google Sheet.
 *
 * Uso:
 *  1. Crea un Google Sheet nuevo > Extensiones > Apps Script y pega este archivo.
 *  2. Implementar > Nueva implementación > Aplicación web
 *     (Ejecutar como: Yo; Acceso: Cualquier persona).
 *  3. Copia la URL /exec en la constante API_URL de index.html.
 */
var HOJA = 'Calificaciones';
var COLS = ['fecha', 'nombre', 'turron', 'masa', 'miel', 'textura', 'sabor', 'general'];

function hoja_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(HOJA);
  if (!sh) {
    sh = ss.insertSheet(HOJA);
    sh.appendRow(COLS);
    sh.setFrozenRows(1);
  }
  return sh;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  var vals = hoja_().getDataRange().getValues();
  var filas = vals.slice(1).map(function (r) {
    var o = {};
    COLS.forEach(function (c, i) { o[c] = r[i]; });
    return o;
  });
  return json_({ ok: true, filas: filas });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var filas = JSON.parse(e.postData.contents).filas || [];
    var sh = hoja_();
    var vals = sh.getDataRange().getValues();
    filas.forEach(function (f) {
      var fila = COLS.map(function (c) { return f[c]; });
      var key = String(f.nombre).trim().toLowerCase() + '|' + f.turron;
      var idx = -1;
      for (var i = 1; i < vals.length; i++) {
        if (String(vals[i][1]).trim().toLowerCase() + '|' + vals[i][2] === key) { idx = i; break; }
      }
      if (idx >= 0) { sh.getRange(idx + 1, 1, 1, COLS.length).setValues([fila]); vals[idx] = fila; }
      else { sh.appendRow(fila); vals.push(fila); }
    });
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}
