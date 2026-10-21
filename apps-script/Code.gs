/**
 * Backend de Google Sheets para "Registro de Cilindros" (SGS) — versión 2.
 *
 * QUÉ CAMBIA FRENTE A LA VERSIÓN 1
 *  - Clave por persona: solo quien tenga una clave válida puede leer o
 *    guardar. La app sabe quién es quién por la clave (no por lo que cada
 *    uno escriba) y el Excel sale "Solicita <nombre>".
 *  - Guardar ya no borra y reescribe la pestaña: une los cambios muestra por
 *    muestra (gana la versión más nueva). Así dos celulares a la vez no se
 *    pisan los datos.
 *  - Deja constancia de quién registró, quién modificó por última vez y
 *    quién marcó como enviada cada muestra.
 *  - Guarda también la fecha de recolección de cada muestra (antes no se
 *    guardaba) y escribe todo como texto para que Sheets no convierta
 *    fechas y horas.
 *
 * VERSIÓN 3 (resistencias)
 *  - Pestaña nueva "Resistencias": la llena el Gestor de Supervisión
 *    (módulo Trazabilidad de Concretos) con los resultados de laboratorio
 *    de cada muestra, una fila por muestra y edad. El Excel de resistencias
 *    es la fuente: cada envío reemplaza la pestaña completa.
 *  - La app la descarga junto con las muestras y el buscador muestra los
 *    resultados (no los edita).
 *  - Pestaña "Historico SGS": las muestras importadas del export de SGS
 *    (las de antes de la app, para el buscador) van aparte. En "Muestras"
 *    quedan solo las registradas en la app para despachar cilindros. La app
 *    recibe las dos juntas, así que no cambia nada para ella. Al primer
 *    guardado con esta versión, las del export se pasan solas de pestaña.
 *
 * ⚠ LAS CLAVES SE ESCRIBEN SOLO AQUÍ, EN TU APPS SCRIPT.
 *   El repositorio de la app es público: nunca pongas las claves reales en
 *   GitHub. Este archivo del repositorio trae claves de ejemplo.
 *
 * INSTALACIÓN / ACTUALIZACIÓN
 * 1. En tu Sheet: "Extensiones" > "Apps Script".
 * 2. Borra todo el contenido de Code.gs y pega TODO este archivo.
 * 3. Cambia las claves de ejemplo de CLAVES (abajo) por claves tuyas
 *    (inventadas, difíciles de adivinar) y el nombre de cada persona.
 * 4. Guarda (ícono de disco o Ctrl+S).
 * 5. "Implementar" > "Gestionar implementaciones" > lápiz (editar) >
 *    Versión: "Nueva versión" > "Implementar". La URL /exec sigue igual.
 *    (Si es la primera vez: "Nueva implementación", tipo "Aplicación web",
 *    ejecutar como "Yo", acceso "Cualquier usuario".)
 * 6. En cada celular: Configuración > "Clave de acceso" > pega la clave de
 *    esa persona > Guardar configuración.
 */

// Clave -> nombre de quien solicita. Agrega o quita personas aquí.
const CLAVES = {
  'CAMBIA-ESTA-CLAVE-1': 'Sebastián Narváez',
  'CAMBIA-ESTA-CLAVE-2': 'José Manuel Urbina'
};

const SHEET_MUESTRAS = 'Muestras';
const SHEET_HISTORICO = 'Historico SGS';
const LOTE_HISTORICO = 'historico-sgs'; // loteId que pone la app a las muestras importadas del export de SGS
const SHEET_CONFIG = 'Config';
const SHEET_RESISTENCIAS = 'Resistencias';

// Una fila por muestra y edad de ensayo (la escribe el Gestor de Supervisión).
const RES_COLS = [
  'muestra','fcMpa','fechaVaciado','localizacion','edad','bloque','ensayos',
  'promedio','porcentaje','criterio','estadoFinal','actualizado','actualizadoPor'
];

const MUESTRA_COLS = [
  'id','numeroMuestra','fechaFundida','horaToma','tipoMaterial','codigoDimension',
  'numeroRemision','concretera','codigoMezcla','asentamientoMm','resistenciaNominalMpa','localizacion',
  'dia1','dia3','dia7','dia14','dia28','dia56','testigo28','testigo56','otroDia','cantidadOtroDia',
  'observacionesMuestra','updatedAt','eliminado','archivado','loteId','fechaRecoleccionLote','fechaExportLote',
  'fotoRemisionUrl','fechaRecoleccion',
  'registradoPor','modificadoPor','enviadoPor','fechaEnviado'
];
// Columnas que solo escribe el servidor (la app no las puede falsear).
const COLS_SERVIDOR = ['registradoPor','modificadoPor','enviadoPor','fechaEnviado'];
const CONFIG_KEYS = ['codigoObra','servicioEnvio','email','telefono','contacto','observacionesSolicitud'];
const COLS_FECHA = ['fechaFundida','fechaRecoleccion','fechaRecoleccionLote','fechaExportLote','fechaEnviado'];

const PHOTOS_FOLDER_NAME = 'Registro de Cilindros - Fotos de remisión';

function doGet(e) {
  let result;
  try {
    const usuario = usuarioDe_(e && e.parameter && e.parameter.clave);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    result = { ok: true, usuario: usuario, config: readConfig_(ss), muestras: readMuestras_(ss),
               resistencias: ss.getSheetByName(SHEET_RESISTENCIAS) ? readRows_(ss, SHEET_RESISTENCIAS) : [] };
  } catch (err) {
    result = { ok: false, error: String(err.message || err) };
  }
  if (e && e.parameter && e.parameter.callback) {
    return jsonpOut_(e.parameter.callback, result);
  }
  return jsonOut_(result);
}

function doPost(e) {
  let result;
  try {
    const body = parseBody_(e);
    const usuario = usuarioDe_(body.clave);
    if (body.action === 'uploadPhoto') {
      result = { ok: true, url: uploadPhoto_(body.fileData, body.fileName, body.mimeType) };
    } else if (body.action === 'resistencias') {
      const lock = LockService.getScriptLock();
      lock.waitLock(20000);
      try {
        const n = writeResistencias_(SpreadsheetApp.getActiveSpreadsheet(), body.resistencias, usuario);
        result = { ok: true, usuario: usuario, filas: n, updatedAt: new Date().toISOString() };
      } finally {
        lock.releaseLock();
      }
    } else {
      const lock = LockService.getScriptLock();
      lock.waitLock(20000); // un guardado a la vez
      try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        if (Array.isArray(body.muestras)) mergeMuestras_(ss, body.muestras, usuario);
        if (body.config) writeConfig_(ss, body.config);
      } finally {
        lock.releaseLock();
      }
      result = { ok: true, usuario: usuario, updatedAt: new Date().toISOString() };
    }
  } catch (err) {
    result = { ok: false, error: String(err.message || err) };
  }
  if (e && e.parameter && e.parameter.reqId) result.reqId = e.parameter.reqId;
  if (e && e.parameter && e.parameter.viaForm) return postMessageOut_(result);
  return jsonOut_(result);
}

function usuarioDe_(clave) {
  const nombre = CLAVES[String(clave || '').trim()];
  if (!nombre) throw new Error('Clave inválida o falta la clave de acceso.');
  return nombre;
}

/* ---------- muestras: unir en vez de reescribir ---------- */

/* Las dos pestañas de muestras juntas (como las ve la app). */
function readMuestras_(ss) {
  const historico = ss.getSheetByName(SHEET_HISTORICO) ? readRows_(ss, SHEET_HISTORICO) : [];
  return readRows_(ss, SHEET_MUESTRAS).concat(historico);
}

function mergeMuestras_(ss, entrantes, usuario) {
  const actuales = readMuestras_(ss);
  const porId = {};
  const orden = [];
  actuales.forEach(function (m) { if (m.id !== '') { porId[m.id] = m; orden.push(m.id); } });
  const ahora = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');

  entrantes.forEach(function (n) {
    const id = String(n.id || '');
    if (!id) return;
    const viejo = porId[id];
    if (viejo && Number(viejo.updatedAt || 0) >= Number(n.updatedAt || 0)) return; // ya hay una versión igual o más nueva
    const nuevo = {};
    MUESTRA_COLS.forEach(function (c) {
      if (COLS_SERVIDOR.indexOf(c) === -1) nuevo[c] = (n[c] === undefined || n[c] === null) ? '' : n[c];
    });
    nuevo.registradoPor = viejo ? (viejo.registradoPor || usuario) : usuario;
    nuevo.modificadoPor = usuario;
    const yaEnviada = viejo && esVerdadero_(viejo.archivado);
    if (esVerdadero_(nuevo.archivado) && !yaEnviada) {
      nuevo.enviadoPor = usuario;
      nuevo.fechaEnviado = ahora;
    } else {
      nuevo.enviadoPor = viejo ? (viejo.enviadoPor || '') : '';
      nuevo.fechaEnviado = viejo ? (viejo.fechaEnviado || '') : '';
    }
    if (!viejo) orden.push(id);
    porId[id] = nuevo;
  });

  // Las del export de SGS a su pestaña; las de la app en "Muestras".
  const deApp = orden.filter(function (id) { return porId[id].loteId !== LOTE_HISTORICO; });
  const deSgs = orden.filter(function (id) { return porId[id].loteId === LOTE_HISTORICO; });
  escribirMuestras_(getSheet_(ss, SHEET_MUESTRAS), deApp, porId);
  if (deSgs.length || ss.getSheetByName(SHEET_HISTORICO)) {
    escribirMuestras_(getSheet_(ss, SHEET_HISTORICO), deSgs, porId);
  }
}

function escribirMuestras_(sh, ids, porId) {
  const data = [MUESTRA_COLS].concat(ids.map(function (id) {
    const m = porId[id];
    return MUESTRA_COLS.map(function (c) {
      const v = m[c];
      if (v === undefined || v === null) return '';
      return typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : String(v);
    });
  }));
  sh.clearContents();
  const rango = sh.getRange(1, 1, data.length, MUESTRA_COLS.length);
  rango.setNumberFormat('@'); // texto: que Sheets no convierta fechas ni horas
  rango.setValues(data);
}

/* ---------- resistencias (desde el Gestor de Supervisión) ---------- */

function writeResistencias_(ss, filas, usuario) {
  if (!Array.isArray(filas)) throw new Error('Faltan las resistencias.');
  const ahora = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  const data = [RES_COLS].concat(filas.map(function (f) {
    return RES_COLS.map(function (c) {
      if (c === 'actualizado') return ahora;
      if (c === 'actualizadoPor') return usuario;
      const v = f[c];
      return (v === undefined || v === null) ? '' : String(v);
    });
  }));
  const sh = getSheet_(ss, SHEET_RESISTENCIAS);
  sh.clearContents();
  const rango = sh.getRange(1, 1, data.length, RES_COLS.length);
  rango.setNumberFormat('@');
  rango.setValues(data);
  return filas.length;
}

function esVerdadero_(v) { return v === true || v === 'TRUE' || v === 'true' || v === 1 || v === '1'; }

/* ---------- fotos ---------- */

function uploadPhoto_(base64Data, fileName, mimeType) {
  if (!base64Data) throw new Error('Falta la imagen.');
  const bytes = Utilities.base64Decode(base64Data);
  const blob = Utilities.newBlob(bytes, mimeType || 'image/jpeg', fileName || ('remision_' + Date.now() + '.jpg'));
  const file = getOrCreatePhotosFolder_().createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return 'https://drive.google.com/uc?id=' + file.getId();
}

function getOrCreatePhotosFolder_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let parent;
  try {
    const parents = DriveApp.getFileById(ss.getId()).getParents();
    parent = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  } catch (e) {
    parent = DriveApp.getRootFolder();
  }
  const existing = parent.getFoldersByName(PHOTOS_FOLDER_NAME);
  return existing.hasNext() ? existing.next() : parent.createFolder(PHOTOS_FOLDER_NAME);
}

function parseBody_(e) {
  // Formulario oculto (viaForm): el JSON viaja como campo de formulario "payload".
  if (e && e.parameter && e.parameter.payload) return JSON.parse(e.parameter.payload);
  // JSON crudo en el cuerpo del POST (respaldo con fetch o pruebas manuales).
  if (e && e.postData && e.postData.contents) return JSON.parse(e.postData.contents);
  return {};
}

/* ---------- helpers de datos ---------- */

function getSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

/* Lee las filas como objetos. Las fechas y horas que Sheets haya convertido
   (datos guardados con la versión 1) se devuelven otra vez como texto. */
function readRows_(ss, sheetName) {
  const values = getSheet_(ss, sheetName).getDataRange().getValues();
  if (values.length < 2) return [];
  const header = values[0];
  const tz = Session.getScriptTimeZone();
  return values.slice(1)
    .filter(function (r) { return r.some(function (c) { return c !== ''; }); })
    .map(function (r) {
      const obj = {};
      header.forEach(function (h, i) {
        let v = r[i];
        if (Object.prototype.toString.call(v) === '[object Date]') {
          v = h === 'horaToma' ? Utilities.formatDate(v, tz, 'HH:mm')
            : (COLS_FECHA.indexOf(h) > -1 ? Utilities.formatDate(v, tz, 'yyyy-MM-dd') : v.toISOString());
        }
        obj[h] = v;
      });
      return obj;
    });
}

function readConfig_(ss) {
  const cfg = {};
  getSheet_(ss, SHEET_CONFIG).getDataRange().getValues().forEach(function (r) { if (r[0]) cfg[r[0]] = r[1]; });
  return cfg;
}

function writeConfig_(ss, config) {
  const sh = getSheet_(ss, SHEET_CONFIG);
  const rows = CONFIG_KEYS.filter(function (k) { return config[k] !== undefined; })
    .map(function (k) { return [k, String(config[k] == null ? '' : config[k])]; });
  sh.clearContents();
  if (rows.length) {
    const rango = sh.getRange(1, 1, rows.length, 2);
    rango.setNumberFormat('@');
    rango.setValues(rows);
  }
}

/* ---------- helpers de respuesta ---------- */

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function jsonpOut_(callbackName, obj) {
  // Nombre de callback saneado por seguridad (solo letras, números, _ y $).
  const safeName = String(callbackName).replace(/[^a-zA-Z0-9_$]/g, '');
  return ContentService.createTextOutput(safeName + '(' + JSON.stringify(obj) + ');')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function postMessageOut_(obj) {
  const json = JSON.stringify(obj).replace(/</g, '\\u003c');
  return HtmlService.createHtmlOutput('<!DOCTYPE html><html><body><script>' +
    'window.parent.postMessage(' + json + ', "*");</script></body></html>');
}
