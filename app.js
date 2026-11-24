/* Plan Infra – Registro Operacional · Aplicación móvil (PWA)
 *
 * Reglas:
 *  - Un registro solo se muestra como GUARDADO cuando el servidor confirma su recepción.
 *  - Todo registro nuevo entra primero a una cola local con un ID único; si no hay
 *    conexión queda PENDIENTE y se reintenta solo. El servidor descarta IDs repetidos.
 *  - Los pendientes de un usuario solo se envían con la sesión de ese mismo usuario.
 */
(function () {
  'use strict';

  const API = (window.RO_CONFIG || {}).API_URL || '';
  const K = { ses: 'ro_sesion', cat: 'ro_catalogo', cola: 'ro_cola', usr: 'ro_usuarios', disp: 'ro_dispositivo', bor: 'ro_borrador_' };
  const $ = (id) => document.getElementById(id);
  const st = { sesion: null, cat: null, form: null, histForm: 'calderas', hist: [], detalle: null, ultimoId: null, sinc: false, errorSync: '' };

  // ── Almacenamiento local (puede fallar en modo privado) ──
  const leer = (k, def) => { try { const v = localStorage.getItem(k); return v ? JSON.parse(v) : def; } catch (e) { return def; } };
  const guardar = (k, v) => { try { localStorage.setItem(k, JSON.stringify(v)); return true; } catch (e) { return false; } };
  const borrar = (k) => { try { localStorage.removeItem(k); } catch (e) { /* nada */ } };

  function uuid() {
    if (crypto.randomUUID) return crypto.randomUUID();
    const b = crypto.getRandomValues(new Uint8Array(16));
    b[6] = (b[6] & 15) | 64; b[8] = (b[8] & 63) | 128;
    const h = Array.from(b, (x) => x.toString(16).padStart(2, '0')).join('');
    return h.slice(0, 8) + '-' + h.slice(8, 12) + '-' + h.slice(12, 16) + '-' + h.slice(16, 20) + '-' + h.slice(20);
  }
  function dispositivo() {
    let d = leer(K.disp, '');
    if (!d) { d = 'D-' + uuid().slice(0, 8).toUpperCase(); guardar(K.disp, d); }
    return d;
  }
  function ahoraLocal() {
    const d = new Date(); d.setMinutes(d.getMinutes() - d.getTimezoneOffset());
    return d.toISOString().slice(0, 16);
  }
  const hoyMenos = (dias) => { const d = new Date(Date.now() - dias * 864e5); d.setMinutes(d.getMinutes() - d.getTimezoneOffset()); return d.toISOString().slice(0, 10); };
  const esc = (s) => String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

  // ── Comunicación con el servidor ──
  async function api(accion, datos) {
    if (!/^https:\/\/script\.google\.com\/macros\/s\/.+\/exec$/.test(API)) throw new Error('Falta configurar la URL del servidor (config.js).');
    const cuerpo = Object.assign({ accion: accion }, datos || {});
    if (st.sesion) cuerpo.token = st.sesion.token;
    const ctrl = new AbortController();
    const t = setTimeout(() => ctrl.abort(), 30000);
    try {
      // Sin cabeceras propias: viaja como text/plain y evita la consulta previa CORS.
      const r = await fetch(API, { method: 'POST', body: JSON.stringify(cuerpo), redirect: 'follow', signal: ctrl.signal });
      if (!r.ok) throw new Error('El servidor respondió HTTP ' + r.status + '.');
      const j = await r.json();
      if (j && j.codigo === 'SESION' && accion !== 'login') sesionVencida(j.error);
      return j;
    } catch (e) {
      if (e.name === 'AbortError') throw new Error('El servidor no respondió a tiempo.');
      if (!navigator.onLine) throw new Error('Sin conexión a Internet.');
      if (e instanceof TypeError) throw new Error('No se pudo conectar con el servidor.');
      throw e;
    } finally { clearTimeout(t); }
  }

  // ── Estado visual ──
  function estado(id, tipo, titulo, detalle) {
    const el = $(id);
    if (!tipo) { el.className = 'estado'; el.textContent = ''; return; }
    el.className = 'estado visible ' + tipo;
    el.textContent = titulo;
    if (detalle) { const s = document.createElement('span'); s.className = 'detalle'; s.textContent = detalle; el.appendChild(s); }
    el.setAttribute('role', tipo === 'error' ? 'alert' : 'status');
  }

  function ir(vista) {
    if (!st.sesion && vista !== 'login') vista = 'login';
    if (vista === 'calderas' || vista === 'cloracion') { abrirFormulario(vista); vista = 'form'; }
    if (vista === 'historial') prepararHistorial();
    if (vista === 'pendientes') pintarPendientes();
    if (vista === 'cuenta') prepararCuenta();
    if (vista === 'inicio') pintarInicio();
    document.querySelectorAll('.vista').forEach((v) => { v.hidden = v.id !== 'v-' + vista; });
    window.scrollTo(0, 0);
  }

  function conexion() {
    $('conexion').textContent = navigator.onLine ? 'EN LÍNEA' : 'SIN CONEXIÓN';
    $('conexion').classList.toggle('off', !navigator.onLine);
    pintarSync();
  }

  // ── Sesión ──
  async function cargarUsuarios() {
    const sel = $('l-usuario');
    let lista = leer(K.usr, []);
    try { const r = await api('usuarios'); if (r.ok) { lista = r.usuarios; guardar(K.usr, lista); } } catch (e) { /* usa la lista guardada */ }
    sel.innerHTML = '<option value="">Selecciona tu nombre…</option>' +
      lista.map((u) => '<option value="' + esc(u.usuario) + '">' + esc(u.nombre) + '</option>').join('');
    const ultimo = leer('ro_ultimo_usuario', '');
    if (ultimo) sel.value = ultimo;
  }

  async function ingresar(ev) {
    ev.preventDefault();
    const usuario = $('l-usuario').value, codigo = $('l-codigo').value;
    if (!usuario || !codigo) { estado('l-estado', 'error', 'Selecciona tu nombre e ingresa tu código.'); return; }
    estado('l-estado', 'enviando', 'Verificando…');
    try {
      const r = await api('login', { usuario: usuario, codigo: codigo, dispositivo: dispositivo() });
      if (!r.ok) { estado('l-estado', 'error', r.error || 'No se pudo ingresar.'); return; }
      st.sesion = { token: r.token, usuario: r.usuario, nombre: r.nombre, rol: r.rol };
      st.cat = r.catalogo;
      guardar(K.ses, st.sesion); guardar(K.cat, st.cat); guardar('ro_ultimo_usuario', r.usuario);
      $('l-codigo').value = ''; estado('l-estado');
      ir('inicio');
      sincronizar();
    } catch (e) { estado('l-estado', 'error', 'No se pudo ingresar.', e.message); }
  }

  function sesionVencida(msg) {
    st.sesion = null; borrar(K.ses);
    ir('login');
    estado('l-estado', 'pendiente', msg || 'Tu sesión venció.', 'Ingresa de nuevo. Los registros pendientes se conservan en este teléfono.');
  }

  async function salir() {
    const n = colaDe().length;
    if (n && !confirm('Tienes ' + n + ' registro(s) sin sincronizar. Se conservarán en este teléfono y se enviarán cuando vuelvas a ingresar. ¿Cerrar sesión?')) return;
    try { await api('salir'); } catch (e) { /* se cierra igual */ }
    st.sesion = null; borrar(K.ses); ir('login');
  }

  function pintarInicio() {
    $('i-nombre').textContent = st.sesion.nombre;
    $('i-rol').textContent = st.sesion.rol === 'supervisor' ? 'Supervisor' : 'Operador';
    const n = colaDe().length;
    $('i-pend').textContent = n ? String(n) : '';
    const hoja = st.cat && st.cat.hojaUrl;
    $('i-hoja').classList.toggle('oculto', !hoja);
    if (hoja) $('i-hoja').href = hoja;
  }

  // ── Formularios de registro ──
  function rangoDe(form, campo, equipo) {
    const rs = ((st.cat && st.cat.rangos) || []).filter((r) => r.form === form && r.campo === campo);
    return rs.find((r) => r.equipo && r.equipo === equipo) || rs.find((r) => !r.equipo) || null;
  }
  const textoRango = (r) => r.min != null && r.max != null ? r.min + ' a ' + r.max : r.min != null ? 'mínimo ' + r.min : 'máximo ' + r.max;

  function abrirFormulario(form) {
    const f = st.cat.formularios[form];
    st.form = form;
    $('r-titulo').textContent = f.titulo;
    $('r-operador').value = st.sesion.nombre;
    $('r-campos').innerHTML = f.campos.map((c) => {
      const id = 'c_' + c.k;
      const lbl = '<label for="' + id + '">' + esc(c.l) + (c.req ? ' <span class="req">*</span>' : '') + '</label>';
      if (c.t === 'lista') {
        const ops = (st.cat.listas[c.lista] || []).map((o) => '<option>' + esc(o) + '</option>').join('');
        return '<div class="campo">' + lbl + '<select id="' + id + '" data-k="' + c.k + '"><option value="">Selecciona…</option>' + ops + '</select></div>';
      }
      return '<div class="campo">' + lbl + '<input id="' + id + '" data-k="' + c.k + '" type="text" inputmode="decimal" maxlength="20" autocomplete="off">' +
        '<span class="ayuda" id="a_' + c.k + '"></span><span class="aviso oculto" id="w_' + c.k + '"></span></div>';
    }).join('');
    // Recuperar borrador si la app se cerró a mitad de un registro.
    const b = leer(K.bor + form, null);
    $('r-fecha').value = (b && b.fechaRegistro) || ahoraLocal();
    $('r-obs').value = (b && b.observaciones) || '';
    f.campos.forEach((c) => { if (b && b[c.k] != null) $('c_' + c.k).value = b[c.k]; });
    estado('r-estado', b ? 'pendiente' : '', b ? 'Se recuperó un borrador sin enviar.' : '');
    revisarRangos();
  }

  function leerFormulario() {
    const f = st.cat.formularios[st.form];
    const reg = { fechaRegistro: $('r-fecha').value, observaciones: $('r-obs').value.trim() };
    f.campos.forEach((c) => { reg[c.k] = $('c_' + c.k).value.trim().replace(',', '.'); });
    return reg;
  }

  function revisarRangos() {
    if (!st.form) return;
    const f = st.cat.formularios[st.form];
    const equipo = ($('c_' + f.equipo) || {}).value || '';
    f.campos.filter((c) => c.t === 'num').forEach((c) => {
      const r = rangoDe(st.form, c.k, equipo);
      const v = $('c_' + c.k).value.trim().replace(',', '.');
      $('a_' + c.k).textContent = r ? 'Rango configurado: ' + textoRango(r) : '';
      const n = Number(v);
      const fuera = r && v !== '' && !isNaN(n) && ((r.min != null && n < r.min) || (r.max != null && n > r.max));
      $('w_' + c.k).classList.toggle('oculto', !fuera);
      $('w_' + c.k).textContent = fuera ? 'Fuera del rango configurado. Revisa el valor y anótalo en observaciones.' : '';
    });
  }

  function guardarBorrador() { if (st.form) guardar(K.bor + st.form, leerFormulario()); }

  function enviarRegistro(ev) {
    ev.preventDefault();
    const f = st.cat.formularios[st.form];
    const reg = leerFormulario();
    const errores = [];
    document.querySelectorAll('#f-reg .invalido').forEach((e) => e.classList.remove('invalido'));
    if (!reg.fechaRegistro) { errores.push('fecha y hora'); $('r-fecha').classList.add('invalido'); }
    f.campos.forEach((c) => {
      const v = reg[c.k];
      const malo = (c.req && !v) || (c.t === 'num' && v !== '' && !/^-?\d+(\.\d+)?$/.test(v));
      if (malo) { errores.push(c.l); $('c_' + c.k).classList.add('invalido'); }
    });
    const numeros = f.campos.filter((c) => c.t === 'num' && reg[c.k] !== '').length;
    const estados = f.campos.filter((c) => c.t === 'lista' && !c.req && reg[c.k] !== '').length;
    if (!numeros && !estados && !reg.observaciones) errores.push('al menos una medición, estado u observación');
    if (errores.length) { estado('r-estado', 'error', 'Revisa el formulario', 'Falta o es inválido: ' + errores.join(', ') + '.'); return; }

    reg.id = uuid();
    reg.dispositivo = dispositivo();
    const cola = leer(K.cola, []);
    cola.push({ form: st.form, usuario: st.sesion.usuario, nombre: st.sesion.nombre, creado: new Date().toISOString(), registro: reg });
    if (!guardar(K.cola, cola)) { estado('r-estado', 'error', 'No se pudo guardar en el teléfono.', 'Revisa el espacio disponible. El registro NO se envió.'); return; }
    borrar(K.bor + st.form);
    st.ultimoId = reg.id;
    estado('r-estado', 'enviando', 'Enviando…', 'Esperando confirmación del servidor.');
    // Dejar el formulario listo para el siguiente registro (los datos ya están en la cola).
    f.campos.forEach((c) => { if (c.t === 'num' || !c.req) $('c_' + c.k).value = ''; });
    $('r-obs').value = ''; $('r-fecha').value = ahoraLocal(); revisarRangos();
    sincronizar();
  }

  // ── Cola y sincronización ──
  const colaDe = () => st.sesion ? leer(K.cola, []).filter((i) => i.usuario === st.sesion.usuario) : [];

  function marcar(id, cambios) {
    const cola = leer(K.cola, []);
    const i = cola.find((x) => x.registro.id === id);
    if (i) { Object.assign(i, cambios); guardar(K.cola, cola); }
  }
  function quitar(id) { guardar(K.cola, leer(K.cola, []).filter((x) => x.registro.id !== id)); }

  async function sincronizar() {
    if (st.sinc || !st.sesion) { pintarSync(); return; }
    st.sinc = true; pintarSync();
    try {
      for (const item of colaDe().filter((i) => !i.rechazo)) {
        let r;
        try { r = await api('guardar', { form: item.form, registro: item.registro }); }
        catch (e) {
          st.errorSync = e.message;
          if (item.registro.id === st.ultimoId) estado('r-estado', 'pendiente', 'Pendiente de sincronización', e.message + ' El registro está guardado SOLO en este teléfono y se enviará automáticamente.');
          break;
        }
        if (r.ok) {
          quitar(item.registro.id); st.errorSync = '';
          if (item.registro.id === st.ultimoId) {
            estado('r-estado', 'ok', 'Registro guardado correctamente',
              (r.duplicado ? 'Ya estaba en el servidor; no se duplicó. ' : '') + 'Recibido ' + r.recibido + ' · fila ' + r.fila + ' · ID ' + r.id.slice(0, 8));
          }
        } else if (r.codigo === 'SESION') {
          break;
        } else if (r.codigo === 'VALIDACION' || r.codigo === 'PERMISO') {
          marcar(item.registro.id, { rechazo: r.error });
          if (item.registro.id === st.ultimoId) estado('r-estado', 'error', 'El servidor rechazó el registro', r.error + ' Revísalo en «Pendientes».');
        } else {
          st.errorSync = r.error || 'Error del servidor.';
          if (item.registro.id === st.ultimoId) estado('r-estado', 'pendiente', 'Pendiente de sincronización', st.errorSync + ' Se reintentará automáticamente.');
          break;
        }
      }
    } finally {
      st.sinc = false; pintarSync();
      if (!$('v-pendientes').hidden) pintarPendientes();
      if (!$('v-inicio').hidden) pintarInicio();
    }
  }

  function pintarSync() {
    const b = $('sync');
    const todos = leer(K.cola, []);
    const mios = colaDe();
    const rech = mios.filter((i) => i.rechazo).length;
    const otros = todos.length - mios.length;
    if (!st.sesion && !todos.length) { b.classList.add('oculto'); return; }
    b.classList.remove('oculto', 'ok', 'error');
    if (st.sinc) { b.textContent = '⟳ Sincronizando ' + mios.length + ' registro(s)…'; return; }
    if (rech) { b.classList.add('error'); b.textContent = '⚠ ' + rech + ' registro(s) rechazado(s) por el servidor. Toca para revisar.'; return; }
    if (mios.length) { b.textContent = '⏳ ' + mios.length + ' pendiente(s) de sincronización' + (navigator.onLine ? '' : ' · sin conexión') + '. Toca para ver.'; return; }
    if (otros) { b.textContent = '⏳ ' + otros + ' pendiente(s) en este teléfono; se enviarán cuando su usuario ingrese.'; return; }
    b.classList.add('ok'); b.textContent = '✔ Todo sincronizado';
  }

  function pintarPendientes() {
    const cola = leer(K.cola, []);
    $('p-lista').innerHTML = cola.length ? cola.map((i) => {
      const f = st.cat && st.cat.formularios[i.form];
      const equipo = f ? i.registro[f.equipo] : '';
      const propio = st.sesion && i.usuario === st.sesion.usuario;
      return '<div class="item pend' + (i.rechazo ? ' rech' : '') + '"><div class="t">' + esc(f ? f.titulo : i.form) + ' · ' + esc(equipo) + '</div>' +
        '<div class="s">' + esc(i.registro.fechaRegistro.replace('T', ' ')) + ' · ' + esc(i.nombre) + ' · ID ' + esc(i.registro.id.slice(0, 8)) + '</div>' +
        (i.rechazo ? '<div class="s"><b>Rechazado:</b> ' + esc(i.rechazo) + '</div>' : '') +
        (!propio ? '<div class="s">Se enviará cuando ' + esc(i.nombre) + ' ingrese en este teléfono.</div>' : '') +
        (i.rechazo && propio ? '<button class="boton secundario" type="button" data-descartar="' + esc(i.registro.id) + '">Descartar</button>' : '') +
        '</div>';
    }).join('') : '<p class="nota">No hay registros pendientes. Todo está en el servidor.</p>';
  }

  // ── Historial ──
  function prepararHistorial() {
    const f = st.cat.formularios[st.histForm];
    document.querySelectorAll('#f-hist .seg button').forEach((b) => b.classList.toggle('activo', b.dataset.form === st.histForm));
    if (!$('h-desde').value) $('h-desde').value = hoyMenos(7);
    if (!$('h-hasta').value) $('h-hasta').value = hoyMenos(0);
    const sup = st.sesion.rol === 'supervisor';
    $('h-op-c').classList.toggle('oculto', !sup);
    if (sup) $('h-op').innerHTML = '<option value="">Todos</option>' + (st.cat.usuarios || []).map((u) => '<option value="' + esc(u.usuario) + '">' + esc(u.nombre) + '</option>').join('');
    const campoEq = f.campos.find((c) => c.k === f.equipo);
    $('h-eq-l').textContent = campoEq.l;
    $('h-eq').innerHTML = '<option value="">Todos</option>' + (st.cat.listas[campoEq.lista] || []).map((o) => '<option>' + esc(o) + '</option>').join('');
  }

  async function buscar(ev) {
    if (ev) ev.preventDefault();
    estado('h-estado', 'enviando', 'Consultando…');
    $('h-res').innerHTML = ''; $('h-csv').classList.add('oculto');
    try {
      const r = await api('historial', { form: st.histForm, desde: $('h-desde').value, hasta: $('h-hasta').value, operador: $('h-op').value, equipo: $('h-eq').value });
      if (!r.ok) { estado('h-estado', 'error', r.error || 'No se pudo consultar.'); return; }
      st.hist = r.registros;
      estado('h-estado');
      const f = st.cat.formularios[st.histForm];
      const nums = f.campos.filter((c) => c.t === 'num').slice(0, 3);
      $('h-res').innerHTML = '<p class="total">' + r.total + ' registro(s)' + (r.total > r.registros.length ? ' · se muestran los ' + r.registros.length + ' más recientes' : '') + '</p>' +
        r.registros.map((x, i) => '<button class="item" type="button" data-ver="' + i + '"><div class="t">' + esc(x.valores[f.equipo]) + ' · ' + esc(x.fecha.slice(0, 16)) + '</div>' +
          '<div class="s">' + esc(x.operador) + ' · ' + nums.map((c) => esc(c.l.split(' (')[0]) + ': ' + esc(x.valores[c.k] === '' ? '—' : x.valores[c.k])).join(' · ') +
          (x.version > 1 ? ' · <b>corregido</b>' : '') + '</div></button>').join('');
      $('h-csv').classList.toggle('oculto', !r.registros.length);
    } catch (e) { estado('h-estado', 'error', 'No se pudo consultar.', e.message); }
  }

  function exportarCsv() {
    const f = st.cat.formularios[st.histForm];
    const enc = ['ID', 'Fecha del registro', 'Recepción servidor', 'Operador'].concat(f.campos.map((c) => c.l), ['Observaciones', 'Versión', 'Última corrección']);
    const celda = (v) => { let s = typeof v === 'number' ? String(v).replace('.', ',') : String(v == null ? '' : v); if (/^[=+\-@]/.test(s)) s = "'" + s; return '"' + s.replace(/"/g, '""') + '"'; };
    const filas = st.hist.map((x) => [x.id, x.fecha, x.recepcion, x.operador].concat(f.campos.map((c) => x.valores[c.k]), [x.observaciones, x.version, x.ultimaCorreccion]));
    const csv = '﻿' + [enc].concat(filas).map((r) => r.map(celda).join(';')).join('\r\n');
    const a = document.createElement('a');
    a.href = URL.createObjectURL(new Blob([csv], { type: 'text/csv;charset=utf-8' }));
    a.download = 'registro_' + st.histForm + '_' + $('h-desde').value + '_' + $('h-hasta').value + '.csv';
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 5000);
  }

  // ── Detalle y corrección ──
  function verDetalle(x) {
    st.detalle = x;
    const f = st.cat.formularios[st.histForm];
    $('d-titulo').textContent = f.titulo;
    const filas = [['Fecha del registro', x.fecha], ['Recepción servidor', x.recepcion], ['Operador', x.operador]]
      .concat(f.campos.map((c) => [c.l, x.valores[c.k] === '' ? '—' : x.valores[c.k]]),
        [['Observaciones', x.observaciones || '—'], ['Dispositivo', x.dispositivo], ['Versión', x.version], ['Última corrección', x.ultimaCorreccion || '—'], ['ID', x.id]]);
    $('d-datos').innerHTML = filas.map((p) => '<dt>' + esc(p[0]) + '</dt><dd>' + esc(p[1]) + '</dd>').join('');
    const sup = st.sesion.rol === 'supervisor';
    $('f-corr').classList.toggle('oculto', !sup);
    $('d-audit').classList.toggle('oculto', !sup);
    if (sup) {
      $('c-campo').innerHTML = '<option value="">Selecciona…</option><option value="fechaRegistro">Fecha del registro</option>' +
        f.campos.map((c) => '<option value="' + c.k + '">' + esc(c.l) + '</option>').join('') + '<option value="observaciones">Observaciones</option>';
      $('c-valor-c').innerHTML = ''; $('c-motivo').value = ''; estado('c-estado');
      cargarAuditoria(x.id);
    }
    ir('detalle');
  }

  function campoCorreccion() {
    const k = $('c-campo').value;
    const f = st.cat.formularios[st.histForm];
    const c = f.campos.find((x) => x.k === k);
    let html = '<label for="c-valor">Valor nuevo</label>';
    if (!k) html = '';
    else if (k === 'fechaRegistro') html += '<input id="c-valor" type="datetime-local">';
    else if (k === 'observaciones') html += '<textarea id="c-valor" maxlength="1000"></textarea>';
    else if (c.t === 'lista') html += '<select id="c-valor"><option value="">(vacío)</option>' + (st.cat.listas[c.lista] || []).map((o) => '<option>' + esc(o) + '</option>').join('') + '</select>';
    else html += '<input id="c-valor" type="text" inputmode="decimal" maxlength="20">';
    $('c-valor-c').innerHTML = html;
  }

  async function corregir(ev) {
    ev.preventDefault();
    const campo = $('c-campo').value, motivo = $('c-motivo').value.trim();
    if (!campo || !$('c-valor')) { estado('c-estado', 'error', 'Selecciona el campo a corregir.'); return; }
    if (motivo.length < 5) { estado('c-estado', 'error', 'Indica el motivo de la corrección.'); return; }
    const c = st.cat.formularios[st.histForm].campos.find((x) => x.k === campo);
    let valor = $('c-valor').value.trim();
    if (c && c.t === 'num') valor = valor.replace(',', '.');
    estado('c-estado', 'enviando', 'Guardando corrección…');
    try {
      const r = await api('corregir', { form: st.histForm, id: st.detalle.id, campo: campo, valor: valor, motivo: motivo });
      if (!r.ok) { estado('c-estado', 'error', r.error || 'No se pudo corregir.'); return; }
      const i = st.hist.findIndex((x) => x.id === r.registro.id);
      if (i >= 0) st.hist[i] = r.registro;
      verDetalle(r.registro);
      estado('c-estado', 'ok', 'Corrección guardada y auditada.');
    } catch (e) { estado('c-estado', 'error', 'No se pudo corregir.', e.message); }
  }

  async function cargarAuditoria(id) {
    $('d-audit-l').innerHTML = '<p class="nota">Cargando…</p>';
    try {
      const r = await api('auditoria', { id: id });
      $('d-audit-l').innerHTML = r.ok && r.registros.length ? r.registros.map((a) => '<div class="audit"><b>' + esc(a.fecha) + '</b> · ' + esc(a.usuario) + ' · ' + esc(a.accion) +
        '<br>' + esc(a.campo) + ': «' + esc(a.anterior) + '» → «' + esc(a.nuevo) + '»<br><i>' + esc(a.motivo) + '</i></div>').join('') : '<p class="nota">Sin correcciones.</p>';
    } catch (e) { $('d-audit-l').innerHTML = '<p class="nota">No se pudo cargar la auditoría.</p>'; }
  }

  // ── Cuenta ──
  function prepararCuenta() {
    ['k-actual', 'k-nuevo', 'k-rep', 'x-nuevo'].forEach((id) => { $(id).value = ''; });
    estado('k-estado'); estado('x-estado');
    const sup = st.sesion.rol === 'supervisor';
    $('f-reset').classList.toggle('oculto', !sup);
    if (sup) $('x-usuario').innerHTML = (st.cat.usuarios || []).map((u) => '<option value="' + esc(u.usuario) + '">' + esc(u.nombre) + ' (' + esc(u.usuario) + ')</option>').join('');
  }

  async function cambiarCodigo(ev) {
    ev.preventDefault();
    if ($('k-nuevo').value !== $('k-rep').value) { estado('k-estado', 'error', 'Los códigos nuevos no coinciden.'); return; }
    if ($('k-nuevo').value.length < 6) { estado('k-estado', 'error', 'El código nuevo debe tener al menos 6 caracteres.'); return; }
    estado('k-estado', 'enviando', 'Guardando…');
    try {
      const r = await api('cambiarCodigo', { actual: $('k-actual').value, nuevo: $('k-nuevo').value });
      if (r.ok) { prepararCuenta(); estado('k-estado', 'ok', 'Código actualizado.'); } else estado('k-estado', 'error', r.error);
    } catch (e) { estado('k-estado', 'error', 'No se pudo cambiar.', e.message); }
  }

  async function restablecer(ev) {
    ev.preventDefault();
    const u = $('x-usuario').value;
    if ($('x-nuevo').value.length < 6) { estado('x-estado', 'error', 'El código debe tener al menos 6 caracteres.'); return; }
    if (!confirm('¿Restablecer el código de ' + u + '?')) return;
    estado('x-estado', 'enviando', 'Guardando…');
    try {
      const r = await api('resetCodigo', { usuario: u, nuevo: $('x-nuevo').value });
      if (r.ok) { $('x-nuevo').value = ''; estado('x-estado', 'ok', 'Código restablecido. Entrégaselo en persona.'); } else estado('x-estado', 'error', r.error);
    } catch (e) { estado('x-estado', 'error', 'No se pudo restablecer.', e.message); }
  }

  // ── Arranque ──
  function iniciar() {
    $('pie-disp').textContent = 'Equipo ' + dispositivo();
    st.sesion = leer(K.ses, null);
    st.cat = leer(K.cat, null);
    if (!st.cat) st.sesion = null;

    document.addEventListener('click', (ev) => {
      const t = ev.target.closest('[data-ir],[data-ver],[data-descartar],[data-form]');
      if (!t) return;
      if (t.dataset.ir) ir(t.dataset.ir);
      else if (t.dataset.ver) verDetalle(st.hist[Number(t.dataset.ver)]);
      else if (t.dataset.descartar) {
        if (confirm('¿Descartar este registro rechazado? No se enviará al servidor.')) { quitar(t.dataset.descartar); pintarPendientes(); pintarSync(); }
      } else if (t.dataset.form) { st.histForm = t.dataset.form; $('h-eq').value = ''; prepararHistorial(); $('h-res').innerHTML = ''; $('h-csv').classList.add('oculto'); }
    });
    $('f-login').addEventListener('submit', ingresar);
    $('f-reg').addEventListener('submit', enviarRegistro);
    $('f-reg').addEventListener('input', () => { revisarRangos(); guardarBorrador(); });
    $('f-reg').addEventListener('change', () => { revisarRangos(); guardarBorrador(); });
    $('f-hist').addEventListener('submit', buscar);
    $('h-csv').addEventListener('click', exportarCsv);
    $('c-campo').addEventListener('change', campoCorreccion);
    $('f-corr').addEventListener('submit', corregir);
    $('f-codigo').addEventListener('submit', cambiarCodigo);
    $('f-reset').addEventListener('submit', restablecer);
    $('b-salir').addEventListener('click', salir);
    $('sync').addEventListener('click', () => { if (st.sesion) { ir('pendientes'); sincronizar(); } });
    $('p-sync').addEventListener('click', sincronizar);
    window.addEventListener('online', () => { conexion(); sincronizar(); });
    window.addEventListener('offline', conexion);
    document.addEventListener('visibilitychange', () => { if (!document.hidden) sincronizar(); });
    setInterval(() => { if (colaDe().length) sincronizar(); }, 30000);

    conexion();
    cargarUsuarios();
    if (st.sesion) {
      ir('inicio');
      sincronizar();
      api('catalogo').then((r) => { if (r.ok) { st.cat = r.catalogo; guardar(K.cat, st.cat); pintarInicio(); } }).catch(() => { /* usa el catálogo guardado */ });
    } else ir('login');
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', iniciar);
  else iniciar();
})();
