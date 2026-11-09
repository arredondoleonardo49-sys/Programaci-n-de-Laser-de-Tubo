/**
 * IGS Láser Tubo · Múltiples de Tubular · backend (Google Apps Script)
 * - Lee (solo lectura) la biblioteca de dibujos de Ingeniería en Drive: archivos .igs/.iges(.gz)
 *   dentro de las subcarpetas de tubular de cada carpeta de modelo (Nº ruta). El recorrido corre
 *   en segundo plano (activador + pasos cortos en cada consulta) y la app muestra el avance.
 * - Guarda programas, tramos (seguimiento de corte) y el inventario de recortes en esta hoja.
 * - Todo, salvo "ping", pide la clave de acceso (propiedad del script CLAVE; ver generarClave).
 * - "archivo" solo entrega IGS que están en la biblioteca leída; las escrituras solo aceptan
 *   las tablas y columnas de TABLAS.
 *
 * Instalación y actualización: ver COMO-IMPLEMENTAR.md (Ejecutar como: yo · Acceso: cualquier usuario).
 */
const VERSION = '2026-11-06a';
const ROOT_ID = '1c-agCWlM9NI_ki2-tIe8SFwYr5aj5rns';   // biblioteca de dibujos de Ingeniería (la misma de nidos)
const CARPETA_TUBULAR = /TUBULAR|IGES|IGS/i;          // nombre de la subcarpeta de tubular dentro de cada modelo
const SALTAR = /L[AÁ]MINA|DXF|MELAMINA|PANEL|AC[UÚ]STICO|PROTOTIPO/i;   // carpetas que no tienen tubular (no se recorren)
const EXT = /\.(igs|iges)(\.gz)*$/i;
const ATAJO = 'application/vnd.google-apps.shortcut', CARPETA = 'application/vnd.google-apps.folder';
const TABLAS = {
  programas: ['id', 'fecha', 'archivo', 'tramos', 'piezas', 'ops', 'resumen', 'estado'],
  tramos: ['id', 'programa', 'tipo', 'tipoLabel', 'orden', 'largo', 'origen', 'recorteUsado', 'recorteNuevo', 'largoRecorte', 'consumo', 'aprov', 'estado', 'inicio', 'fin', 'piezas', 'maquina'],
  recortes: ['id', 'tipo', 'tipoLabel', 'tipoCorto', 'largo', 'estado', 'origen', 'programa', 'tramo', 'usadoEn', 'ubicacion', 'alta', 'disponibleDesde', 'baja']
};
const BIB = ['id', 'nombre', 'ruta', 'tam', 'modificado'], COLA = ['id', 'ruta', 'enTubular', 'nivel', 'hecho'];
const HOJA = n => 'Tub_' + n.charAt(0).toUpperCase() + n.slice(1);
const MAX_FILAS = 200, MAX_CELDA = 49000, MAX_NIVEL = 15;
const PASO_WEB = 20000, PASO_ACTIVADOR = 270000;     // ms de recorrido por consulta y por activador
const props_ = () => PropertiesService.getScriptProperties();

/* ---------------- entrada ---------------- */
function doGet(e) {
  const p = (e && e.parameter) || {};
  const cb = /^[\w$.]{1,80}$/.test(p.callback || '') ? p.callback : '';
  let out;
  try {
    if (p.action === 'ping') out = { ok: true, mensaje: 'Conectado · ' + SpreadsheetApp.getActive().getName(), version: VERSION, clave: !!props_().getProperty('CLAVE') };
    else {
      claveOk_(p.clave);
      switch (p.action) {
        case 'estado': out = { programas: leer_('programas'), tramos: leer_('tramos'), recortes: leer_('recortes') }; break;
        case 'biblioteca': out = biblioteca_(p.refrescar == '1'); break;
        case 'archivo': out = archivo_(p.id); break;
        case 'diagnostico': out = { raiz: DriveApp.getFolderById(ROOT_ID).getName(), estado: estadoBib_().estado || '', version: VERSION }; break;
        default: out = { error: 'Acción desconocida' };
      }
    }
  } catch (err) { out = error_(err); }
  return salida_(out, cb);
}

function doPost(e) {
  let out;
  try {
    const b = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    claveOk_(b.clave);
    if (b.op !== 'upsert') throw new Error('Operación no válida');
    const filas = validar_(b.tabla, b.filas);
    const lock = LockService.getScriptLock(); lock.waitLock(30000);
    try { upsert_(b.tabla, filas); } finally { lock.releaseLock(); }
    out = { ok: true, n: filas.length };
  } catch (err) { out = error_(err); }
  return salida_(out, '');
}

function salida_(out, cb) {
  const txt = JSON.stringify(out);
  if (cb) return ContentService.createTextOutput(cb + '(' + txt + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  return ContentService.createTextOutput(txt).setMimeType(ContentService.MimeType.JSON);
}
function error_(err) { const o = { error: String((err && err.message) || err) }; if (err && err.auth) o.auth = err.auth; return o; }

/* ---------------- clave de acceso ---------------- */
function claveOk_(c) {
  const k = props_().getProperty('CLAVE');
  if (!k) throw auth_('Falta configurar la clave en el Apps Script: corre generarClave() en el editor.');
  if (String(c || '') !== k) throw auth_('Clave de acceso incorrecta o vacía.');
}
function auth_(m) { const e = new Error(m); e.auth = 'clave'; return e; }
/** Correr una vez desde el editor: crea la clave, la deja en el registro de ejecución y autoriza Drive, Sheets y activadores. */
function generarClave() {
  const k = Utilities.getUuid().replace(/-/g, '').slice(0, 12).toUpperCase();
  props_().setProperty('CLAVE', k);
  DriveApp.getFolderById(ROOT_ID).getName(); SpreadsheetApp.getActive().getName(); borrarActivadores_();   // también quita activadores de la versión anterior
  Logger.log('Clave de acceso de IGS Láser Tubo: ' + k + ' (escríbela en Configuración de cada equipo)');
  return k;
}

/* ---------------- tablas ---------------- */
function hoja_(tabla) {
  const ss = SpreadsheetApp.getActive(); let sh = ss.getSheetByName(HOJA(tabla));
  if (!sh) { sh = ss.insertSheet(HOJA(tabla)); sh.getRange(1, 1, 1, TABLAS[tabla].length).setValues([TABLAS[tabla]]).setFontWeight('bold'); sh.setFrozenRows(1); }
  return sh;
}
function leer_(tabla) {
  const sh = hoja_(tabla), v = sh.getDataRange().getValues(); if (v.length < 2) return [];
  const H = v[0];
  return v.slice(1).filter(r => r[0] !== '').map(r => { const o = {}; H.forEach((h, i) => o[h] = r[i] instanceof Date ? r[i].getTime() : r[i]); return o; });
}
function validar_(tabla, filas) {
  const cols = TABLAS[tabla]; if (!cols) throw new Error('Tabla no válida: ' + tabla);
  if (!Array.isArray(filas) || !filas.length) throw new Error('El envío no trae filas');
  if (filas.length > MAX_FILAS) throw new Error('Demasiadas filas en un envío (' + filas.length + ')');
  filas.forEach(f => {
    if (!f || typeof f !== 'object' || Array.isArray(f)) throw new Error('Fila no válida');
    const id = f.id == null ? '' : String(f.id); if (!id || id.length > 120) throw new Error('Fila sin id válido');
    Object.keys(f).forEach(k => {
      if (cols.indexOf(k) < 0) throw new Error('Columna no permitida en ' + tabla + ': ' + k);
      const v = f[k], s = typeof v === 'object' && v !== null ? JSON.stringify(v) : String(v == null ? '' : v);
      if (s.length > MAX_CELDA) throw new Error('Celda demasiado grande en ' + k);
    });
  });
  return filas;
}
/* upsert por id: se escribe un renglón por fila existente y las altas en un solo bloque */
function upsert_(tabla, filas) {
  const sh = hoja_(tabla); let H = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const nuevas = TABLAS[tabla].filter(k => H.indexOf(k) < 0 && filas.some(f => k in f));
  if (nuevas.length) { sh.getRange(1, H.length + 1, 1, nuevas.length).setValues([nuevas]).setFontWeight('bold'); H = H.concat(nuevas); }
  const n = sh.getLastRow(), datos = n > 1 ? sh.getRange(2, 1, n - 1, H.length).getValues() : [];
  const idx = {}; datos.forEach((r, i) => { if (r[0] !== '') idx[String(r[0])] = i; });
  const ahora = Date.now(), cambiadas = new Map(), altas = new Map();
  const valor = v => v === '__NOW__' ? ahora : typeof v === 'object' && v !== null ? JSON.stringify(v) : v;
  filas.forEach(f => {
    const id = String(f.id); let r;
    if (idx[id] != null) { r = datos[idx[id]]; cambiadas.set(idx[id], r); }
    else { r = altas.get(id); if (!r) { r = H.map(() => ''); altas.set(id, r); } }
    Object.keys(f).forEach(k => { r[H.indexOf(k)] = valor(f[k]); });
  });
  cambiadas.forEach((r, i) => sh.getRange(i + 2, 1, 1, H.length).setValues([r]));
  if (altas.size) { const v = Array.from(altas.values()); sh.getRange(sh.getLastRow() + 1, 1, v.length, H.length).setValues(v); }
}

/* ---------------- archivos IGS (solo los de la biblioteca) ---------------- */
function archivo_(id) {
  id = String(id || '');
  if (!/^[\w-]{10,}$/.test(id)) throw new Error('ID de archivo no válido');
  if (!idsBiblioteca_().has(id)) throw new Error('Ese archivo no está en la biblioteca de tubular. Si es nuevo, recorre de nuevo la biblioteca.');
  const f = DriveApp.getFileById(id);
  return { nombre: f.getName(), b64: Utilities.base64Encode(f.getBlob().getBytes()) };
}
function idsBiblioteca_() {
  const ids = new Set(), ss = SpreadsheetApp.getActive();
  ['Tub_Biblioteca', H_NUEVA].forEach(n => { const sh = ss.getSheetByName(n); if (sh && sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().forEach(r => ids.add(String(r[0]))); });
  return ids;
}

/* ---------------- biblioteca de Drive en segundo plano ---------------- */
/* hojas de trabajo del recorrido (mismos nombres que la versión anterior) y estado en una sola propiedad JSON */
const H_COLA = 'Tub_Cola', H_NUEVA = 'Tub_BibTmp';
function estadoBib_() { try { return JSON.parse(props_().getProperty('bibEstadoJson') || '{}'); } catch (e) { return {}; } }
function guardarEstadoBib_(st) { props_().setProperty('bibEstadoJson', JSON.stringify(st)); }
function hojaAux_(nombre, encabezado) {
  const ss = SpreadsheetApp.getActive(); let sh = ss.getSheetByName(nombre);
  if (!sh) { sh = ss.insertSheet(nombre); sh.getRange(1, 1, 1, encabezado.length).setValues([encabezado]).setFontWeight('bold'); if (nombre !== 'Tub_Biblioteca') sh.hideSheet(); }
  return sh;
}
function limpiar_(sh) { if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).clearContent(); }
function listaBib_() {
  const sh = SpreadsheetApp.getActive().getSheetByName('Tub_Biblioteca');
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues().filter(r => r[0] !== '')
    .map(r => ({ id: String(r[0]), nombre: r[1], ruta: r[2], tam: r[3], modificado: r[4] instanceof Date ? r[4].getTime() : r[4] }));
}
/* refrescar = volver a recorrer Drive; sin lectura previa también arranca. Cada consulta avanza un paso corto. */
function biblioteca_(refrescar) {
  let st = estadoBib_();
  const atorado = st.estado === 'corriendo' && Date.now() - (st.latido || st.inicio || 0) > 15 * 60000;
  const sh = SpreadsheetApp.getActive().getSheetByName('Tub_Biblioteca'), hay = !!sh && sh.getLastRow() > 1;
  if (((refrescar || !hay) && st.estado !== 'corriendo') || atorado) st = iniciarBib_();
  if (st.estado === 'corriendo') { pasoBib_(PASO_WEB); st = estadoBib_(); }
  return {
    estado: st.estado || 'listo',
    progreso: { carpetas: st.carpetas || 0, pendientes: st.pendientes || 0, archivos: st.archivos || 0, error: st.error || '' },
    actualizado: Number(props_().getProperty('bibliotecaFecha')) || null,
    archivos: listaBib_()
  };
}
function iniciarBib_() {
  const cola = hojaAux_(H_COLA, COLA), nueva = hojaAux_(H_NUEVA, BIB);
  [[cola, COLA], [nueva, BIB]].forEach(([sh, enc]) => { sh.clear(); sh.getRange(1, 1, 1, enc.length).setValues([enc]).setFontWeight('bold'); sh.hideSheet(); });   // la versión anterior las dejaba sin encabezado
  cola.getRange(2, 1, 1, 5).setValues([[ROOT_ID, '', 0, 0, '']]);
  const st = { estado: 'corriendo', inicio: Date.now(), latido: Date.now(), carpetas: 0, pendientes: 1, archivos: 0, error: '', avisos: 0 };
  guardarEstadoBib_(st); programarBib_();
  return st;
}
function programarBib_() { borrarActivadores_(); ScriptApp.newTrigger('continuarBiblioteca').timeBased().after(5000).create(); }
function borrarActivadores_() { ScriptApp.getProjectTriggers().forEach(t => { if (['continuarBiblioteca', 'actualizarBiblioteca'].indexOf(t.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(t); }); }
/** Lo llama el activador: avanza el recorrido hasta ~4.5 min y, si falta, se vuelve a programar. */
function continuarBiblioteca() {
  borrarActivadores_();
  pasoBib_(PASO_ACTIVADOR);
  if (estadoBib_().estado === 'corriendo') programarBib_();
}
/** Se puede ejecutar a mano desde el editor: recorre la biblioteca completa (sigue sola si no le alcanza el tiempo).
    Mismo nombre que en la versión anterior, así un activador pendiente de esa versión no falla. */
function actualizarBiblioteca() {
  if (estadoBib_().estado !== 'corriendo') iniciarBib_();
  continuarBiblioteca();
}
function pasoBib_(presupuesto) {
  const lock = LockService.getScriptLock(); if (!lock.tryLock(500)) return;   // otro paso ya está corriendo
  const t0 = Date.now();
  try {
    let st = estadoBib_(); if (st.estado !== 'corriendo') return;
    const cola = hojaAux_(H_COLA, COLA), nueva = hojaAux_(H_NUEVA, BIB);
    const n = cola.getLastRow(), Q = n > 1 ? cola.getRange(2, 1, n - 1, 5).getValues().filter(r => r[0] !== '') : [];
    const vistos = new Set(Q.map(r => String(r[0]))), hallados = [];
    for (let i = 0; i < Q.length && Date.now() - t0 < presupuesto; i++) {
      if (Q[i][4]) continue;
      const ruta = String(Q[i][1]), enT = !!Number(Q[i][2]), nivel = Number(Q[i][3]) || 0;
      try {
        recorrerCarpeta_(DriveApp.getFolderById(String(Q[i][0])), ruta, enT, nivel,
          s => { if (!vistos.has(s.id) && s.nivel <= MAX_NIVEL) { vistos.add(s.id); Q.push([s.id, s.ruta, s.enT ? 1 : 0, s.nivel, '']); } },
          f => hallados.push(f));
        Q[i][4] = 1;
      } catch (err) { Q[i][4] = 'error: ' + String((err && err.message) || err).slice(0, 200); st.avisos = (st.avisos || 0) + 1; }
    }
    if (hallados.length) nueva.getRange(nueva.getLastRow() + 1, 1, hallados.length, 5).setValues(hallados.map(a => [a.id, a.nombre, a.ruta, a.tam, new Date(a.modificado)]));
    if (Q.length) cola.getRange(2, 1, Q.length, 5).setValues(Q);
    const pendientes = Q.filter(r => !r[4]).length;
    st = Object.assign(st, { latido: Date.now(), carpetas: Q.length - pendientes, pendientes: pendientes, archivos: (st.archivos || 0) + hallados.length });
    if (!pendientes) terminarBib_(st, cola, nueva); else guardarEstadoBib_(st);
  } catch (err) {
    guardarEstadoBib_(Object.assign(estadoBib_(), { estado: 'error', error: String((err && err.message) || err) })); borrarActivadores_();
  } finally { lock.releaseLock(); }
}
function recorrerCarpeta_(carpeta, ruta, enTubular, nivel, alSub, alArchivo) {
  const files = carpeta.getFiles();
  while (files.hasNext()) {
    const f = files.next();
    if (f.getMimeType() === ATAJO) {   // sigue accesos directos (base compartida entre departamentos)
      const tid = f.getTargetId(), tm = f.getTargetMimeType();
      if (tm === CARPETA) { const sub = DriveApp.getFolderById(tid); if (!SALTAR.test(sub.getName())) alSub({ id: tid, ruta: ruta + '/' + sub.getName(), enT: enTubular || CARPETA_TUBULAR.test(sub.getName()), nivel: nivel + 1 }); }
      else { const tf = DriveApp.getFileById(tid); if (EXT.test(tf.getName()) && (enTubular || /TUBULAR/i.test(tf.getName()))) alArchivo(info_(tf, ruta)); }
      continue;
    }
    if (EXT.test(f.getName()) && (enTubular || /TUBULAR/i.test(f.getName()))) alArchivo(info_(f, ruta));
  }
  const subs = carpeta.getFolders();
  while (subs.hasNext()) { const s = subs.next(); if (SALTAR.test(s.getName())) continue; alSub({ id: s.getId(), ruta: ruta + '/' + s.getName(), enT: enTubular || CARPETA_TUBULAR.test(s.getName()), nivel: nivel + 1 }); }
}
function terminarBib_(st, cola, nueva) {
  const n = nueva.getLastRow(), v = n > 1 ? nueva.getRange(2, 1, n - 1, 5).getValues().filter(r => r[0] !== '') : [];
  const vistos = new Set(), unicos = v.filter(r => { const k = String(r[0]); if (vistos.has(k)) return false; vistos.add(k); return true; });
  const bib = hojaAux_('Tub_Biblioteca', BIB);
  if (!unicos.length && bib.getLastRow() > 1) {   // un recorrido vacío no borra la lectura anterior
    guardarEstadoBib_(Object.assign(st, { estado: 'error', error: 'El recorrido no encontró IGS de tubular; se conserva la lectura anterior.' }));
  } else {
    bib.clearContents(); bib.getRange(1, 1, 1, 5).setValues([BIB]).setFontWeight('bold');
    if (unicos.length) bib.getRange(2, 1, unicos.length, 5).setValues(unicos);
    const fecha = Date.now(); props_().setProperty('bibliotecaFecha', String(fecha));
    guardarEstadoBib_(Object.assign(st, { estado: 'listo', fin: fecha, archivos: unicos.length, pendientes: 0 }));
  }
  limpiar_(nueva); limpiar_(cola); borrarActivadores_();
}
function info_(f, ruta) { return { id: f.getId(), nombre: f.getName(), ruta: ruta, tam: f.getSize(), modificado: f.getLastUpdated().getTime() }; }
