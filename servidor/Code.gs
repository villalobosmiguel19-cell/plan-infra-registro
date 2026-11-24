/**
 * PLAN INFRA – REGISTRO OPERACIONAL · Servidor (Google Apps Script)
 *
 * Independiente de Plan Infra. Guarda los registros en la hoja de cálculo
 * PLAN_INFRA_REGISTRO_OPERACIONAL del dueño del script (no se comparte).
 *
 * Puesta en marcha (una vez): ejecutar configurar() desde el editor.
 * Códigos de acceso: Configuración del proyecto → Propiedades de script →
 *   CODIGO_INICIAL_<usuario>  (ej. CODIGO_INICIAL_supervisor). En el primer ingreso
 *   el servidor lo convierte en huella (hash con sal) y borra el texto original.
 */

const ARCHIVO = 'PLAN_INFRA_REGISTRO_OPERACIONAL';
const H = {
  USUARIOS: 'Usuarios', CONFIG: 'Config', RANGOS: 'Rangos',
  CALDERAS: 'Calderas', CLORACION: 'Cloracion', AUDITORIA: 'Auditoria',
};
const SESION_SEG = 21600;            // 6 h; se renueva con cada uso
const MAX_FALLOS = 8;                // intentos fallidos por usuario…
const VENTANA_FALLOS = 900;          // …cada 15 minutos
const MIN_CODIGO = 6;
const FMT = 'dd-MM-yyyy HH:mm:ss';

// Formularios: [clave, etiqueta, tipo, obligatorio]. tipo: 'num' o 'lista:<LISTA>'.
const FORMS = {
  calderas: {
    titulo: 'Control de calderas', hoja: H.CALDERAS, equipo: 'caldera',
    campos: [
      ['caldera', 'Caldera', 'lista:CALDERA', true],
      ['ablandador', 'Ablandador', 'lista:ESTADO_EQUIPO'],
      ['alimentacion', 'Alimentación', 'lista:ESTADO_EQUIPO'],
      ['filtro', 'Filtro', 'lista:ESTADO_EQUIPO'],
      ['pozo', 'Pozo', 'lista:ESTADO_EQUIPO'],
      ['ph', 'pH', 'num'],
      ['dureza', 'Dureza (ppm)', 'num'],
      ['conductividad', 'Conductividad (µS/cm)', 'num'],
      ['sulfito', 'Sulfito (ppm)', 'num'],
      ['consumoAgua', 'Consumo de agua (m³)', 'num'],
    ],
  },
  cloracion: {
    titulo: 'Control de cloración', hoja: H.CLORACION, equipo: 'punto',
    campos: [
      ['punto', 'Punto de medición', 'lista:PUNTO_CLORACION', true],
      ['cloroLibre', 'Cloro residual libre (mg/L)', 'num'],
      ['ph', 'pH del agua', 'num'],
      ['concentracion', 'Concentración de solución (%)', 'num'],
      ['nivelEstanque', 'Nivel del estanque (%)', 'num'],
      ['estadoBomba', 'Estado bomba dosificadora', 'lista:ESTADO_BOMBA'],
      ['consumoHipoclorito', 'Consumo de hipoclorito (L)', 'num'],
    ],
  },
};
const PRE = ['ID', 'Fecha del registro', 'Recepción servidor', 'Usuario', 'Operador'];
const POST = ['Observaciones', 'Dispositivo', 'Versión', 'Última corrección'];
const C_OBS = 'observaciones';

// ───────────────────────── Puesta en marcha ─────────────────────────

function configurar() {
  const props = PropertiesService.getScriptProperties();
  let ss = null;
  const id = props.getProperty('SHEET_ID');
  if (id) { try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; } }
  if (!ss) { ss = SpreadsheetApp.create(ARCHIVO); props.setProperty('SHEET_ID', ss.getId()); }
  const inicial = ss.getSheets()[0];

  asegurar_(ss, H.USUARIOS, ['Usuario', 'Nombre', 'Rol', 'Activo'], [
    ['supervisor', 'Supervisor', 'supervisor', 'SI'],
    ['operador1', 'Operador 1', 'operador', 'SI'],
    ['operador2', 'Operador 2', 'operador', 'SI'],
  ]);
  asegurar_(ss, H.CONFIG, ['Lista', 'Valor'], [
    ['CALDERA', 'Caldera 1'], ['CALDERA', 'Caldera 2'],
    ['ESTADO_EQUIPO', 'Normal'], ['ESTADO_EQUIPO', 'Con observación'], ['ESTADO_EQUIPO', 'Fuera de servicio'],
    ['PUNTO_CLORACION', 'Punto 1 (renombrar)'], ['PUNTO_CLORACION', 'Punto 2 (renombrar)'],
    ['ESTADO_BOMBA', 'Operativa'], ['ESTADO_BOMBA', 'Detenida'], ['ESTADO_BOMBA', 'Falla'],
  ]);
  asegurar_(ss, H.RANGOS,
    ['Formulario', 'Campo', 'Mínimo', 'Máximo', 'Equipo o punto (vacío = todos)', 'Validado por'], []);
  Object.keys(FORMS).forEach(function (k) {
    const sh = asegurar_(ss, FORMS[k].hoja, encabezados_(FORMS[k]), []);
    sh.getRange('A:A').setNumberFormat('@');
    sh.getRange('B:C').setNumberFormat('dd-mm-yyyy hh:mm:ss');
  });
  const au = asegurar_(ss, H.AUDITORIA,
    ['Fecha', 'Usuario', 'Acción', 'Formulario', 'ID registro', 'Campo', 'Valor anterior', 'Valor nuevo', 'Motivo'], []);
  au.getRange('A:A').setNumberFormat('dd-mm-yyyy hh:mm:ss');

  if (Object.values(H).indexOf(inicial.getName()) === -1 && ss.getSheets().length > 1) ss.deleteSheet(inicial);
  Logger.log('Hoja lista: ' + ss.getUrl());
  Logger.log('Define CODIGO_INICIAL_<usuario> en Propiedades de script para cada usuario de la hoja Usuarios.');
}

function asegurar_(ss, nombre, enc, filas) {
  let sh = ss.getSheetByName(nombre);
  if (!sh) sh = ss.insertSheet(nombre);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, enc.length).setValues([enc])
      .setFontWeight('bold').setBackground('#0B2545').setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
    if (filas.length) sh.getRange(2, 1, filas.length, enc.length).setValues(filas);
  }
  return sh;
}

function encabezados_(f) {
  return PRE.concat(f.campos.map(function (c) { return c[1]; }), POST);
}

// ───────────────────────── Entrada HTTP ─────────────────────────

function doGet(e) {
  const api = e && e.parameter && e.parameter.api;
  if (api === 'estado') return json_({ ok: true, servicio: 'PLAN INFRA – REGISTRO OPERACIONAL', hora: ahora_() });
  return json_({ ok: false, error: 'Usa la aplicación móvil.' });
}

function doPost(e) {
  let p;
  try { p = JSON.parse((e && e.postData && e.postData.contents) || ''); }
  catch (x) { return json_({ ok: false, error: 'Solicitud inválida.', codigo: 'VALIDACION' }); }
  try {
    return json_(despachar_(p || {}));
  } catch (x) {
    if (x && x.codigo) return json_({ ok: false, error: x.message, codigo: x.codigo });
    console.error(x);
    return json_({ ok: false, error: 'Error interno del servidor.', codigo: 'SERVIDOR' });
  }
}

function despachar_(p) {
  if (p.accion === 'usuarios') return { ok: true, usuarios: listaUsuarios_() };
  if (p.accion === 'login') return login_(p);
  const s = sesion_(p.token);
  switch (p.accion) {
    case 'catalogo': return { ok: true, catalogo: catalogo_(s) };
    case 'guardar': return guardar_(s, p);
    case 'historial': return historial_(s, p);
    case 'corregir': return corregir_(s, p);
    case 'auditoria': return auditoria_(s, p);
    case 'cambiarCodigo': return cambiarCodigo_(s, p);
    case 'resetCodigo': return resetCodigo_(s, p);
    case 'salir':
      CacheService.getScriptCache().remove('s_' + p.token);
      return { ok: true };
  }
  throw falla_('Acción desconocida.');
}

// ───────────────────────── Usuarios y sesiones ─────────────────────────

function usuarios_() {
  const v = hoja_(H.USUARIOS).getDataRange().getValues();
  const m = {};
  for (let i = 1; i < v.length; i++) {
    const u = String(v[i][0]).trim().toLowerCase();
    if (!u) continue;
    m[u] = {
      usuario: u, nombre: String(v[i][1]).trim() || u,
      rol: String(v[i][2]).trim().toLowerCase() === 'supervisor' ? 'supervisor' : 'operador',
      activo: String(v[i][3]).trim().toUpperCase() === 'SI',
    };
  }
  return m;
}

function listaUsuarios_() {
  const m = usuarios_();
  return Object.keys(m).filter(function (u) { return m[u].activo; })
    .map(function (u) { return { usuario: u, nombre: m[u].nombre }; });
}

function login_(p) {
  const u = String(p.usuario || '').trim().toLowerCase();
  const codigo = String(p.codigo || '');
  const cache = CacheService.getScriptCache();
  const claveFallos = 'f_' + u;
  const fallos = Number(cache.get(claveFallos) || 0);
  if (fallos >= MAX_FALLOS) throw falla_('Demasiados intentos fallidos. Espera 15 minutos.', 'BLOQUEO');

  const usr = usuarios_()[u];
  const props = PropertiesService.getScriptProperties();
  let ok = false;
  if (usr && usr.activo && codigo) {
    const guardado = props.getProperty('HASH_' + u);
    if (guardado) {
      const partes = guardado.split(':');
      ok = hash_(codigo, partes[0]) === partes[1];
    } else {
      const inicial = props.getProperty('CODIGO_INICIAL_' + u);
      ok = !!inicial && codigo === inicial;
      if (ok) {
        if (codigo.length < MIN_CODIGO) throw falla_('El código inicial debe tener al menos ' + MIN_CODIGO + ' caracteres.');
        guardarHash_(u, codigo);
        props.deleteProperty('CODIGO_INICIAL_' + u);
      }
    }
  }
  if (!ok) {
    cache.put(claveFallos, String(fallos + 1), VENTANA_FALLOS);
    throw falla_('Usuario o código incorrecto.', 'CREDENCIALES');
  }
  cache.remove(claveFallos);
  const token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  const s = { u: u, nombre: usr.nombre, rol: usr.rol, disp: texto_(p.dispositivo, 40, false) };
  cache.put('s_' + token, JSON.stringify(s), SESION_SEG);
  auditar_(u, 'Inicio de sesión', '', '', '', '', '', 'Dispositivo ' + s.disp);
  return { ok: true, token: token, usuario: u, nombre: usr.nombre, rol: usr.rol, catalogo: catalogo_(s) };
}

function sesion_(token) {
  if (!token || !/^[a-f0-9]{64}$/.test(String(token))) throw falla_('Sesión no válida. Ingresa de nuevo.', 'SESION');
  const cache = CacheService.getScriptCache();
  const v = cache.get('s_' + token);
  if (!v) throw falla_('Tu sesión venció. Ingresa de nuevo.', 'SESION');
  const s = JSON.parse(v);
  const usr = usuarios_()[s.u];
  if (!usr || !usr.activo) { cache.remove('s_' + token); throw falla_('Usuario desactivado.', 'SESION'); }
  s.nombre = usr.nombre; s.rol = usr.rol; s.token = token;
  cache.put('s_' + token, JSON.stringify({ u: s.u, nombre: s.nombre, rol: s.rol, disp: s.disp }), SESION_SEG);
  return s;
}

function hash_(codigo, sal) {
  let d = sal + '|' + codigo, b = null;
  for (let i = 0; i < 300; i++) {
    b = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, d, Utilities.Charset.UTF_8);
    d = Utilities.base64Encode(b) + sal;
  }
  return Utilities.base64Encode(b);
}

function guardarHash_(u, codigo) {
  const sal = Utilities.getUuid().replace(/-/g, '');
  PropertiesService.getScriptProperties().setProperty('HASH_' + u, sal + ':' + hash_(codigo, sal));
}

function cambiarCodigo_(s, p) {
  const guardado = PropertiesService.getScriptProperties().getProperty('HASH_' + s.u);
  if (!guardado) throw falla_('No hay código registrado.');
  const partes = guardado.split(':');
  if (hash_(String(p.actual || ''), partes[0]) !== partes[1]) throw falla_('El código actual no es correcto.');
  validarCodigoNuevo_(p.nuevo);
  guardarHash_(s.u, String(p.nuevo));
  auditar_(s.u, 'Cambio de código', '', '', '', '', '', '');
  return { ok: true };
}

function resetCodigo_(s, p) {
  soloSupervisor_(s);
  const u = String(p.usuario || '').trim().toLowerCase();
  if (!usuarios_()[u]) throw falla_('Usuario no existe.');
  validarCodigoNuevo_(p.nuevo);
  guardarHash_(u, String(p.nuevo));
  PropertiesService.getScriptProperties().deleteProperty('CODIGO_INICIAL_' + u);
  auditar_(s.u, 'Restablecer código', '', '', 'Código de ' + u, '', '', '');
  return { ok: true };
}

function validarCodigoNuevo_(c) {
  c = String(c || '');
  if (c.length < MIN_CODIGO || c.length > 64) throw falla_('El código debe tener entre ' + MIN_CODIGO + ' y 64 caracteres.');
}

function soloSupervisor_(s) {
  if (s.rol !== 'supervisor') throw falla_('Solo el supervisor puede hacer esto.', 'PERMISO');
}

// ───────────────────────── Catálogo ─────────────────────────

function listas_() {
  const v = hoja_(H.CONFIG).getDataRange().getValues();
  const m = {};
  for (let i = 1; i < v.length; i++) {
    const l = String(v[i][0]).trim().toUpperCase(), val = String(v[i][1]).trim();
    if (!l || !val) continue;
    (m[l] = m[l] || []).push(val);
  }
  return m;
}

function rangos_() {
  const v = hoja_(H.RANGOS).getDataRange().getValues();
  const r = [];
  for (let i = 1; i < v.length; i++) {
    const form = norm_(v[i][0]);
    const f = FORMS[form];
    if (!f) continue;
    const campo = f.campos.filter(function (c) { return norm_(c[0]) === norm_(v[i][1]) || norm_(c[1]) === norm_(v[i][1]); })[0];
    if (!campo) continue;
    const min = v[i][2] === '' ? null : Number(v[i][2]);
    const max = v[i][3] === '' ? null : Number(v[i][3]);
    if (min === null && max === null) continue;
    r.push({ form: form, campo: campo[0], min: min, max: max, equipo: String(v[i][4]).trim() });
  }
  return r;
}

function catalogo_(s) {
  const formularios = {};
  Object.keys(FORMS).forEach(function (k) {
    const f = FORMS[k];
    formularios[k] = {
      titulo: f.titulo, equipo: f.equipo,
      campos: f.campos.map(function (c) {
        return { k: c[0], l: c[1], t: c[2].split(':')[0], lista: c[2].split(':')[1] || '', req: !!c[3] };
      }),
    };
  });
  const cat = { formularios: formularios, listas: listas_(), rangos: rangos_() };
  if (s.rol === 'supervisor') {
    cat.usuarios = listaUsuarios_();
    cat.hojaUrl = SpreadsheetApp.openById(sheetId_()).getUrl();
  }
  return cat;
}

// ───────────────────────── Registros ─────────────────────────

function guardar_(s, p) {
  const f = FORMS[p.form];
  if (!f) throw falla_('Formulario desconocido.');
  const r = p.registro || {};
  const id = String(r.id || '').trim();
  if (!/^[A-Za-z0-9-]{8,64}$/.test(id)) throw falla_('Identificador de registro inválido.');
  const fecha = fecha_(r.fechaRegistro);
  if (!fecha) throw falla_('Fecha y hora del registro inválidas.');
  if (fecha.getTime() > Date.now() + 3600 * 1000) throw falla_('La fecha del registro está en el futuro.');
  const listas = listas_();
  const valores = f.campos.map(function (c) { return valorCampo_(c, r[c[0]], listas); });
  const obs = texto_(r.observaciones, 1000, true);
  const disp = texto_(r.dispositivo || s.disp, 40, false);

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw falla_('Servidor ocupado. Se reintentará.', 'OCUPADO');
  try {
    const sh = hoja_(f.hoja);
    const enc = sh.getRange('A:A').createTextFinder(id).matchEntireCell(true).findNext();
    if (enc) return { ok: true, duplicado: true, id: id, fila: enc.getRow(), recibido: ahora_() };
    const recepcion = new Date();
    sh.appendRow([id, fecha, recepcion, txt_(s.u), txt_(s.nombre)].concat(valores.map(celda_), [txt_(obs), txt_(disp), 1, '']));
    SpreadsheetApp.flush();
    return { ok: true, duplicado: false, id: id, fila: sh.getLastRow(), recibido: fmt_(recepcion) };
  } finally {
    lock.releaseLock();
  }
}

function valorCampo_(c, v, listas) {
  const tipo = c[2].split(':')[0];
  if (tipo === 'num') {
    const t = String(v == null ? '' : v).trim().replace(',', '.');
    if (t === '') return '';
    if (!/^-?\d+(\.\d+)?$/.test(t) || Math.abs(Number(t)) > 1e9) throw falla_('Valor no numérico en «' + c[1] + '».');
    return Number(t);
  }
  const val = String(v == null ? '' : v).trim();
  if (!val) { if (c[3]) throw falla_('Falta «' + c[1] + '».'); return ''; }
  const opciones = listas[c[2].split(':')[1]] || [];
  if (opciones.indexOf(val) === -1) throw falla_('Valor no válido en «' + c[1] + '».');
  return val;
}

function historial_(s, p) {
  const f = FORMS[p.form];
  if (!f) throw falla_('Formulario desconocido.');
  const desde = p.desde ? fecha_(p.desde + 'T00:00') : null;
  const hasta = p.hasta ? fecha_(p.hasta + 'T23:59') : null;
  const operador = s.rol === 'supervisor' ? String(p.operador || '').toLowerCase() : s.u;
  const equipo = String(p.equipo || '');
  const v = hoja_(f.hoja).getDataRange().getValues();
  const n = PRE.length, out = [];
  for (let i = 1; i < v.length; i++) {
    const row = v[i];
    if (!row[0]) continue;
    const fr = row[1] instanceof Date ? row[1] : null;
    if (desde && (!fr || fr < desde)) continue;
    if (hasta && (!fr || fr > hasta)) continue;
    if (operador && String(row[3]).toLowerCase() !== operador) continue;
    if (equipo && String(row[n]) !== equipo) continue;
    out.push(fila_(f, row, i + 1));
  }
  out.sort(function (a, b) { return a.orden < b.orden ? 1 : -1; });
  return { ok: true, total: out.length, registros: out.slice(0, 500) };
}

function fila_(f, row, nFila) {
  const n = PRE.length, m = f.campos.length;
  const valores = {};
  f.campos.forEach(function (c, j) { valores[c[0]] = row[n + j]; });
  return {
    id: String(row[0]), fila: nFila,
    fecha: row[1] instanceof Date ? fmt_(row[1]) : String(row[1]),
    orden: row[1] instanceof Date ? row[1].getTime() : 0,
    recepcion: row[2] instanceof Date ? fmt_(row[2]) : String(row[2]),
    usuario: String(row[3]), operador: String(row[4]), valores: valores,
    observaciones: String(row[n + m]), dispositivo: String(row[n + m + 1]),
    version: Number(row[n + m + 2]) || 1, ultimaCorreccion: String(row[n + m + 3]),
  };
}

function corregir_(s, p) {
  soloSupervisor_(s);
  const f = FORMS[p.form];
  if (!f) throw falla_('Formulario desconocido.');
  const motivo = texto_(p.motivo, 300, true);
  if (motivo.length < 5) throw falla_('Indica el motivo de la corrección (mínimo 5 caracteres).');
  const n = PRE.length, m = f.campos.length;
  let col, etiqueta, nuevo;
  if (p.campo === C_OBS) {
    col = n + m + 1; etiqueta = 'Observaciones'; nuevo = texto_(p.valor, 1000, true);
  } else if (p.campo === 'fechaRegistro') {
    col = 2; etiqueta = 'Fecha del registro'; nuevo = fecha_(p.valor);
    if (!nuevo) throw falla_('Fecha inválida.');
  } else {
    const j = f.campos.map(function (c) { return c[0]; }).indexOf(p.campo);
    if (j === -1) throw falla_('Campo no válido.');
    col = n + j + 1; etiqueta = f.campos[j][1]; nuevo = valorCampo_(f.campos[j], p.valor, listas_());
  }
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw falla_('Servidor ocupado. Reintenta.', 'OCUPADO');
  try {
    const sh = hoja_(f.hoja);
    const enc = sh.getRange('A:A').createTextFinder(String(p.id || '')).matchEntireCell(true).findNext();
    if (!enc || enc.getRow() === 1) throw falla_('Registro no encontrado.');
    const r = enc.getRow();
    const celda = sh.getRange(r, col);
    const anterior = celda.getValue();
    celda.setValue(celda_(nuevo));
    const cVer = sh.getRange(r, n + m + 3);
    cVer.setValue((Number(cVer.getValue()) || 1) + 1);
    sh.getRange(r, n + m + 4).setValue(ahora_() + ' por ' + s.u);
    auditar_(s.u, 'Corrección', p.form, String(p.id), etiqueta, legible_(anterior), legible_(nuevo), motivo);
    SpreadsheetApp.flush();
    return { ok: true, registro: fila_(f, sh.getRange(r, 1, 1, sh.getLastColumn()).getValues()[0], r) };
  } finally {
    lock.releaseLock();
  }
}

function auditoria_(s, p) {
  soloSupervisor_(s);
  const v = hoja_(H.AUDITORIA).getDataRange().getValues();
  const id = String(p.id || '');
  const out = [];
  for (let i = 1; i < v.length; i++) {
    if (id && String(v[i][4]) !== id) continue;
    out.push({ fecha: v[i][0] instanceof Date ? fmt_(v[i][0]) : String(v[i][0]), usuario: String(v[i][1]),
      accion: String(v[i][2]), campo: String(v[i][5]), anterior: String(v[i][6]),
      nuevo: String(v[i][7]), motivo: String(v[i][8]) });
  }
  return { ok: true, registros: out.reverse().slice(0, 200) };
}

function auditar_(u, accion, form, id, campo, anterior, nuevo, motivo) {
  hoja_(H.AUDITORIA).appendRow([new Date(), txt_(u), txt_(accion), txt_(form), txt_(id), txt_(campo),
    txt_(texto_(anterior, 500, true)), txt_(texto_(nuevo, 500, true)), txt_(texto_(motivo, 300, true))]);
}

/** Fuerza texto en la hoja: sin esto, Sheets (configuración regional de Chile) convierte
 *  textos como "7.5" en fechas. El apóstrofo inicial no se muestra ni se lee de vuelta. */
function txt_(v) {
  const s = String(v == null ? '' : v);
  return s === '' ? '' : "'" + s.replace(/^'/, '');
}
function celda_(v) { return typeof v === 'string' ? txt_(v) : v; }

// ───────────────────────── Utilidades ─────────────────────────

function sheetId_() {
  const id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  if (!id) throw falla_('Falta ejecutar configurar().', 'SERVIDOR');
  return id;
}

function hoja_(nombre) {
  const sh = SpreadsheetApp.openById(sheetId_()).getSheetByName(nombre);
  if (!sh) throw falla_('Falta la hoja «' + nombre + '». Ejecuta configurar().', 'SERVIDOR');
  return sh;
}

/** "AAAA-MM-DDTHH:MM" → Date en la zona horaria del script. */
function fecha_(v) {
  const m = /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})/.exec(String(v || ''));
  if (!m) return null;
  const f = new Date(+m[1], +m[2] - 1, +m[3], +m[4], +m[5], 0);
  if (isNaN(f.getTime()) || f.getMonth() !== +m[2] - 1) return null;
  return f;
}

function texto_(v, max, saltos) {
  let s = String(v == null ? '' : v);
  s = saltos ? s.replace(/[\u0000-\u0009\u000B-\u001F\u007F]/g, ' ') : s.replace(/[\u0000-\u001F\u007F]/g, ' ');
  s = s.trim().slice(0, max);
  if (/^[=+\-@]/.test(s)) s = "'" + s;   // evita que la hoja lo interprete como fórmula
  return s;
}

function legible_(v) { return v instanceof Date ? fmt_(v) : String(v); }
function norm_(s) { return String(s || '').trim().toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, ''); }
function fmt_(d) { return Utilities.formatDate(d, Session.getScriptTimeZone(), FMT); }
function ahora_() { return fmt_(new Date()); }
function falla_(msg, codigo) { const e = new Error(msg); e.codigo = codigo || 'VALIDACION'; return e; }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
