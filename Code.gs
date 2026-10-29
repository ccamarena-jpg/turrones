/**
 * Backend de Halloween TT Audit '26: guarda las calificaciones en un Google Sheet.
 * Hojas: "Calificaciones" (turrones) y "Disfraces" (concurso de disfraces).
 *
 * Uso:
 *  1. En el Sheet > Extensiones > Apps Script, pega este archivo.
 *  2. Implementar > Administrar implementaciones > editar > Nueva versión > Implementar
 *     (Ejecutar como: Yo; Acceso: Cualquier persona). La URL /exec no cambia.
 */
var HOJAS = {
  turrones:  { nombre: 'Calificaciones', clave: 'turron',       cols: ['fecha', 'nombre', 'turron', 'masa', 'miel', 'textura', 'sabor', 'general'] },
  disfraces: { nombre: 'Disfraces',      clave: 'participante', cols: ['fecha', 'nombre', 'participante', 'mejor', 'creativo', 'divertido'] }
};

function cfg_(hoja) { return HOJAS[hoja] || HOJAS.turrones; }

function hoja_(cfg) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(cfg.nombre);
  if (!sh) {
    sh = ss.insertSheet(cfg.nombre);
    sh.appendRow(cfg.cols);
    sh.setFrozenRows(1);
  }
  return sh;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// Los resultados solo se entregan con el código guardado en Propiedades del script
// (Configuración del proyecto > Propiedades de la secuencia de comandos > CODIGO_RESULTADOS).
function doGet(e) {
  var p = (e && e.parameter) || {};
  var esperado = PropertiesService.getScriptProperties().getProperty('CODIGO_RESULTADOS');
  if (!esperado || String(p.codigo || '') !== esperado) return json_({ ok: false, error: 'Código incorrecto' });
  var cfg = cfg_(p.hoja);
  var vals = hoja_(cfg).getDataRange().getValues();
  var filas = vals.slice(1).map(function (r) {
    var o = {};
    cfg.cols.forEach(function (c, i) { o[c] = r[i]; });
    return o;
  });
  return json_({ ok: true, filas: filas });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var body = JSON.parse(e.postData.contents);
    var cfg = cfg_(body.hoja);
    var filas = body.filas || [];
    var sh = hoja_(cfg);
    var vals = sh.getDataRange().getValues();
    var iNom = cfg.cols.indexOf('nombre'), iClave = cfg.cols.indexOf(cfg.clave);
    filas.forEach(function (f) {
      var fila = cfg.cols.map(function (c) { return f[c]; });
      var key = String(f.nombre).trim().toLowerCase() + '|' + f[cfg.clave];
      var idx = -1;
      for (var i = 1; i < vals.length; i++) {
        if (String(vals[i][iNom]).trim().toLowerCase() + '|' + vals[i][iClave] === key) { idx = i; break; }
      }
      if (idx >= 0) { sh.getRange(idx + 1, 1, 1, cfg.cols.length).setValues([fila]); vals[idx] = fila; }
      else { sh.appendRow(fila); vals.push(fila); }
    });
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}
