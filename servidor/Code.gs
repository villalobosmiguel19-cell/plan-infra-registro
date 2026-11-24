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
  CALDERAS: 'Calderas', CLORACION: 'Cloracion', AUDITORIA: 'Auditoria', CALDERAS_V1: 'Calderas_v1',
};
const SESION_SEG = 21600;            // 6 h; se renueva con cada uso
const MAX_FALLOS = 8;                // intentos fallidos por usuario…
const VENTANA_FALLOS = 900;          // …cada 15 minutos
const MIN_CODIGO = 6;
const FMT = 'dd-MM-yyyy HH:mm:ss';

// Control de calderas: mismos puntos y parámetros que «Control de Calderas» de Plan Infra.
const PUNTOS_CALDERA = ['Pozo', 'Filtro', 'Ablandador', 'Alimentación', 'Condensado', 'Caldera 1', 'Caldera 2'];
const PARAMS_CALDERA = [
  ['ph', 'pH', ''], ['dureza', 'Dureza', 'ppm'], ['conductividad', 'Conductividad', 'µS/cm'],
  ['sulfito', 'Sulfito', 'ppm'],
];
// Consumo de agua: ítem aparte, un solo punto, ~1 vez al día. No tiene rango: se suma.
const CONSUMO = { clave: 'consumoAgua', etiqueta: 'Consumo de agua (m³)', grupo: 'Consumo de agua' };

function claveCaldera_(punto, param) {
  return norm_(punto).replace(/\s+/g, '') + '_' + param;
}

function camposCalderas_() {
  const c = [['turno', 'Turno', 'lista:TURNO', true]];
  PUNTOS_CALDERA.forEach(function (p) {
    PARAMS_CALDERA.forEach(function (q) {
      const corto = q[1] + (q[2] ? ' (' + q[2] + ')' : '');
      // extra: g = grupo (punto), lc = etiqueta corta, s = se suma (no tiene rango)
      c.push([claveCaldera_(p, q[0]), p + ' · ' + corto, 'num', false, { g: p, lc: corto, s: false }]);
    });
  });
  c.push([CONSUMO.clave, CONSUMO.etiqueta, 'num', false, { g: CONSUMO.grupo, lc: CONSUMO.etiqueta, s: true }]);
  return c;
}

// Formularios: [clave, etiqueta, tipo, obligatorio, extra]. tipo: 'num' o 'lista:<LISTA>'.
const FORMS = {
  calderas: {
    titulo: 'Control de calderas', hoja: H.CALDERAS, equipo: 'turno',
    campos: camposCalderas_(),
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
  if (p.accion === 'consulta') return consulta_(p);
  if (p.accion === 'corregirPlanInfra') return corregirPlanInfra_(p);
  if (p.accion === 'importar') return importar_(p);
  const s = sesion_(p.token);
  switch (p.accion) {
    case 'catalogo': return { ok: true, catalogo: catalogo_(s) };
    case 'tendencia': return tendencia_(s, p);
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

// ───────────────────────── Consulta para Plan Infra (solo lectura) ─────────────────────────

/** Plan Infra consulta registros con una clave propia (PLANINFRA_CLAVE, mínimo 16 caracteres).
 *  Solo lee: no guarda, no corrige y no toca códigos de usuarios. */
function validarClavePlanInfra_(recibida) {
  const clave = PropertiesService.getScriptProperties().getProperty('PLANINFRA_CLAVE');
  const cache = CacheService.getScriptCache();
  const fallos = Number(cache.get('f_planinfra') || 0);
  if (fallos >= MAX_FALLOS) throw falla_('Demasiados intentos fallidos. Espera 15 minutos.', 'BLOQUEO');
  if (!clave || clave.length < 16 || String(recibida || '') !== clave) {
    cache.put('f_planinfra', String(fallos + 1), VENTANA_FALLOS);
    throw falla_('Clave de consulta inválida.', 'CREDENCIALES');
  }
}

/** Corrección de un valor desde Plan Infra (supervisor del PC). Mismas reglas que en la app:
 *  motivo obligatorio y auditoría con valor anterior. Requiere PLANINFRA_PUEDE_CORREGIR = SI. */
function corregirPlanInfra_(p) {
  if (PropertiesService.getScriptProperties().getProperty('PLANINFRA_PUEDE_CORREGIR') !== 'SI') {
    throw falla_('Las correcciones desde Plan Infra no están habilitadas.', 'PERMISO');
  }
  validarClavePlanInfra_(p.clave);
  const usuario = String(p.usuario || '').replace(/[^A-Za-z0-9._@-]/g, '').slice(0, 40);
  if (!usuario) throw falla_('Falta el usuario de Plan Infra.');
  return corregir_({ u: 'planinfra:' + usuario, rol: 'supervisor', nombre: usuario }, p);
}

function consulta_(p) {
  validarClavePlanInfra_(p.clave);
  const h = historial_({ u: 'planinfra', rol: 'supervisor', nombre: 'Plan Infra' }, p);
  const cat = catalogo_({ rol: 'operador' });
  return {
    ok: true, total: h.total, registros: h.registros, formularios: cat.formularios,
    listas: cat.listas, rangos: cat.rangos, usuarios: listaUsuarios_(), consultado: ahora_(),
  };
}

// ───────────────────────── Fusión con «Control de Calderas» (una vez) ─────────────────────────

/** Ejecutar una vez desde el editor. Pasa la hoja Calderas al formato de matriz (7 puntos ×
 *  5 parámetros + turno) sin perder datos: la hoja anterior queda como «Calderas_v1» y sus
 *  registros se copian al formato nuevo. Agrega la lista TURNO y las filas de Rangos vacías. */
function actualizarEstructura() {
  const ss = SpreadsheetApp.openById(sheetId_());
  const f = FORMS.calderas;
  const enc = encabezados_(f);
  let sh = ss.getSheetByName(H.CALDERAS);
  if (sh && sh.getLastColumn() > 0) {
    const actual = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
    if (actual.join('|') !== enc.join('|')) {
      if (ss.getSheetByName(H.CALDERAS_V1)) throw new Error('Ya existe la hoja ' + H.CALDERAS_V1 + '. Revisar a mano.');
      sh.setName(H.CALDERAS_V1);
      sh = null;
    }
  }
  if (!sh) {
    sh = asegurar_(ss, H.CALDERAS, enc, []);
    sh.getRange('A:A').setNumberFormat('@');
    sh.getRange('B:C').setNumberFormat('dd-mm-yyyy hh:mm:ss');
    Logger.log('Registros del formato anterior copiados: ' + convertirV1_(ss, sh, f));
  }
  if (!listas_().TURNO) {
    const cfg = ss.getSheetByName(H.CONFIG);
    cfg.getRange(cfg.getLastRow() + 1, 1, 3, 2).setValues([['TURNO', 'Mañana'], ['TURNO', 'Tarde'], ['TURNO', 'Noche']]);
  }
  const rg = ss.getSheetByName(H.RANGOS);
  const existentes = rg.getDataRange().getValues().slice(1).map(function (r) { return norm_(r[0]) + '|' + norm_(r[1]); });
  const filas = f.campos
    .filter(function (c) { return c[2] === 'num' && !(c[4] && c[4].s); })
    .map(function (c) { return ['Calderas', c[1], '', '', '', '']; })
    .filter(function (r) { return existentes.indexOf(norm_(r[0]) + '|' + norm_(r[1])) === -1; });
  if (filas.length) rg.getRange(rg.getLastRow() + 1, 1, filas.length, 6).setValues(filas);
  Logger.log('Estructura lista. Filas de Rangos agregadas: ' + filas.length);
}

function convertirV1_(ss, nueva, f) {
  const v1 = ss.getSheetByName(H.CALDERAS_V1);
  if (!v1 || v1.getLastRow() < 2) return 0;
  const datos = v1.getDataRange().getValues();
  const h = datos[0].map(String);
  const col = function (nombre) { return h.indexOf(nombre); };
  const mapa = { 'pH': 'ph', 'Dureza (ppm)': 'dureza', 'Conductividad (µS/cm)': 'conductividad',
    'Sulfito (ppm)': 'sulfito', 'Consumo de agua (m³)': 'consumoAgua' };
  const filas = [];
  for (let i = 1; i < datos.length; i++) {
    const r = datos[i];
    if (!r[0]) continue;
    const caldera = String(r[col('Caldera')] || '');
    const valores = {};
    Object.keys(mapa).forEach(function (et) {
      const v = r[col(et)];
      if (v === '' || !caldera) return;
      valores[mapa[et] === CONSUMO.clave ? CONSUMO.clave : claveCaldera_(caldera, mapa[et])] = v;
    });
    const estados = ['Ablandador', 'Alimentación', 'Filtro', 'Pozo']
      .filter(function (e) { return r[col(e)] !== ''; })
      .map(function (e) { return e + ': ' + r[col(e)]; }).join('; ');
    const obs = (estados ? '[Formato anterior] ' + estados + '. ' : '') + String(r[col('Observaciones')] || '');
    filas.push([r[0], r[1], r[2], txt_(r[3]), txt_(r[4])]
      .concat(f.campos.map(function (c) { return c[0] === 'turno' ? '' : (valores[c[0]] !== undefined ? valores[c[0]] : ''); }),
        [txt_(obs.trim()), txt_(r[col('Dispositivo')]), r[col('Versión')] || 1, txt_(r[col('Última corrección')])]));
  }
  if (filas.length) nueva.getRange(2, 1, filas.length, filas[0].length).setValues(filas);
  auditar_('sistema', 'Cambio de formato', 'calderas', '', 'Registros copiados desde Calderas_v1', '', String(filas.length), '');
  return filas.length;
}

/** Ejecutar una vez desde el editor. Saca el consumo de agua de los 7 puntos y lo deja como un
 *  solo ítem. Sin perder datos: la hoja actual queda como «Calderas_v2» y se copia fila por fila
 *  al formato nuevo (el consumo se toma del punto que lo tenga; si hay más de uno, se avisa). */
function separarConsumo() {
  const ss = SpreadsheetApp.openById(sheetId_());
  const f = FORMS.calderas;
  const enc = encabezados_(f);
  const actual = ss.getSheetByName(H.CALDERAS);
  const viejo = actual.getRange(1, 1, 1, actual.getLastColumn()).getValues()[0].map(String);
  if (viejo.join('|') === enc.join('|')) { Logger.log('La hoja ya tiene el formato nuevo.'); return; }
  if (ss.getSheetByName('Calderas_v2')) throw new Error('Ya existe la hoja Calderas_v2. Revisar a mano.');
  const datos = actual.getDataRange().getValues();
  actual.setName('Calderas_v2');
  const nueva = asegurar_(ss, H.CALDERAS, enc, []);
  nueva.getRange('A:A').setNumberFormat('@');
  nueva.getRange('B:C').setNumberFormat('dd-mm-yyyy hh:mm:ss');
  const colConsumo = viejo.map(function (h, i) { return /Consumo de agua/.test(h) ? i : -1; }).filter(function (i) { return i >= 0; });
  let conflictos = 0;
  const filas = datos.slice(1).filter(function (r) { return r[0]; }).map(function (r) {
    return enc.map(function (h, j) {
      let v;
      if (h === CONSUMO.etiqueta) {
        const con = colConsumo.map(function (i) { return r[i]; }).filter(function (x) { return x !== ''; });
        if (con.length > 1) conflictos++;
        v = con.length ? con[0] : '';
      } else {
        const i = viejo.indexOf(h);
        v = i >= 0 ? r[i] : '';
      }
      // Textos (usuario, operador, turno, observaciones, dispositivo, corrección) como texto.
      return typeof v === 'string' && v !== '' && j !== 0 ? txt_(v) : v;
    });
  });
  if (filas.length) nueva.getRange(2, 1, filas.length, enc.length).setValues(filas);
  auditar_('sistema', 'Cambio de formato', 'calderas', '', 'Consumo de agua como ítem aparte (respaldo en Calderas_v2)',
    '', String(filas.length), conflictos ? conflictos + ' registro(s) con consumo en más de un punto: se tomó el primero' : '');
  Logger.log('Copiados ' + filas.length + ' registros. Registros con consumo en más de un punto: ' + conflictos);
}

/** Importación única del histórico de Plan Infra. Requiere PLANINFRA_CLAVE y la propiedad
 *  IMPORTACION_HABILITADA = SI (se borra al terminar). IDs fijos «PI-<id>»: nunca duplica. */
function importar_(p) {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('IMPORTACION_HABILITADA') !== 'SI') throw falla_('La importación no está habilitada.', 'PERMISO');
  const clave = props.getProperty('PLANINFRA_CLAVE');
  if (!clave || String(p.clave || '') !== clave) throw falla_('Clave de consulta inválida.', 'CREDENCIALES');
  if (p.form !== 'calderas') throw falla_('Solo se importa el formulario de calderas.');
  const f = FORMS.calderas;
  const lista = Array.isArray(p.registros) ? p.registros.slice(0, 200) : [];
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw falla_('Servidor ocupado. Reintenta.', 'OCUPADO');
  try {
    const sh = hoja_(f.hoja);
    const ids = {};
    if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().forEach(function (r) { ids[String(r[0])] = true; });
    const filas = [];
    let omitidos = 0;
    lista.forEach(function (r) {
      const id = String(r.id || '');
      if (!/^PI-\d+$/.test(id)) throw falla_('ID inválido: ' + id);
      if (ids[id]) { omitidos++; return; }
      const fecha = fecha_(r.fechaRegistro);
      if (!fecha) throw falla_('Fecha inválida en ' + id);
      const vals = f.campos.map(function (c) {
        if (c[0] === 'turno') return txt_(texto_(r.turno, 20, false));
        const v = r.valores ? r.valores[c[0]] : undefined;
        return typeof v === 'number' && isFinite(v) ? v : '';
      });
      filas.push([id, fecha, new Date(), txt_('historico'), txt_(texto_(r.operador, 60, false))]
        .concat(vals, [txt_(texto_(r.observaciones, 1000, true)), txt_('Migración Plan Infra'), 1, '']));
      ids[id] = true;
    });
    if (filas.length) sh.getRange(sh.getLastRow() + 1, 1, filas.length, filas[0].length).setValues(filas);
    auditar_('plan_infra', 'Importación histórico', 'calderas', '', 'Registros importados', '', String(filas.length),
      'Desde Control de Calderas de Plan Infra');
    return { ok: true, importados: filas.length, omitidos: omitidos };
  } finally {
    lock.releaseLock();
  }
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
        const x = c[4] || {};
        return { k: c[0], l: c[1], t: c[2].split(':')[0], lista: c[2].split(':')[1] || '', req: !!c[3],
          g: x.g || '', lc: x.lc || '', s: !!x.s };
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
  const operador = s.rol === 'supervisor' ? String(p.operador || '').toLowerCase() : s.u;
  const out = filtrar_(f, p.desde, p.hasta, operador, String(p.equipo || ''));
  return { ok: true, total: out.length, registros: out.slice(0, 500) };
}

/** Gráficos de tendencia: todos los registros de la planta en el período (orden cronológico). */
function tendencia_(s, p) {
  const f = FORMS[p.form];
  if (!f) throw falla_('Formulario desconocido.');
  const out = filtrar_(f, p.desde, p.hasta, '', '').slice(0, 3000).reverse();
  return {
    ok: true, total: out.length, rangos: rangos_(),
    registros: out.map(function (r) { return { fecha: r.fecha, orden: r.orden, valores: r.valores }; }),
  };
}

function filtrar_(f, desdeTxt, hastaTxt, operador, equipo) {
  const desde = desdeTxt ? fecha_(desdeTxt + 'T00:00') : null;
  const hasta = hastaTxt ? fecha_(hastaTxt + 'T23:59') : null;
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
  return out;
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
