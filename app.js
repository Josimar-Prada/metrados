/* Metrados de obra — app de campo sin conexión para el SISTEMA DE VALORIZACIÓN (Porotobango).
   Datos solo en el teléfono (IndexedDB). Envío: archivo CSV + fotos por el menú Compartir → Google Drive.
   El Excel los lee con Power Query (hoja CAMPO_DATOS) y los valida con las mismas reglas del ingreso manual. */
'use strict';
const APP_VERSION = '1.0.7';
const CAT_SHEET = 'CATÁLOGO CELULAR';
const MAX_FOTOS = 3, FOTO_LADO = 1600, FOTO_CAL = 0.72, FOTOS_POR_ENVIO = 6;
const CSV_COLS = ['ID', 'FECHA', 'CODIGO', 'PPTO', 'FRENTE', 'N_VECES', 'LARGO', 'ANCHO', 'ALTO', 'MET_DIRECTO', 'OBSERVACION',
  'USUARIO', 'LATITUD', 'LONGITUD', 'PRECISION_M', 'FOTOS', 'CREADO'];

const $ = id => document.getElementById(id);
const S = { recs: [], cat: null, byCode: new Map(), ids: new Set(), user: '', lotes: [], sel: null, modo: 'med',
  fotosNuevas: [], fotosQuitar: [], gps: null, prep: null };

/* ---------------- IndexedDB ---------------- */
let DB;
function idb() {
  return new Promise((ok, ko) => {
    const rq = indexedDB.open('metrados', 1);
    rq.onupgradeneeded = () => {
      const d = rq.result;
      d.createObjectStore('recs', { keyPath: 'id' });
      d.createObjectStore('fotos', { keyPath: 'id' });
      d.createObjectStore('meta', { keyPath: 'k' });
    };
    rq.onsuccess = () => ok(rq.result); rq.onerror = () => ko(rq.error);
  });
}
function tx(store, mode, fn) {
  return new Promise((ok, ko) => {
    const t = DB.transaction(store, mode); const st = t.objectStore(store); let out;
    Promise.resolve(fn(st)).then(v => { out = v; });
    t.oncomplete = () => ok(out); t.onerror = () => ko(t.error); t.onabort = () => ko(t.error);
  });
}
const rqp = rq => new Promise((ok, ko) => { rq.onsuccess = () => ok(rq.result); rq.onerror = () => ko(rq.error); });
const getAll = s => tx(s, 'readonly', st => rqp(st.getAll()));
const put = (s, v) => tx(s, 'readwrite', st => { st.put(v); });
const del = (s, k) => tx(s, 'readwrite', st => { st.delete(k); });
const getMeta = async k => { const r = await tx('meta', 'readonly', st => rqp(st.get(k))); return r ? r.v : null; };
const setMeta = (k, v) => put('meta', { k, v });

/* ---------------- utilidades ---------------- */
const pad = n => String(n).padStart(2, '0');
function hoy() { const d = new Date(); return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`; }
function ahoraISO() { const d = new Date(); return `${hoy()}T${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`; }
function uuid() {
  if (crypto.randomUUID) return crypto.randomUUID();
  const b = crypto.getRandomValues(new Uint8Array(16)); b[6] = (b[6] & 15) | 64; b[8] = (b[8] & 63) | 128;
  const h = [...b].map(x => x.toString(16).padStart(2, '0')).join('');
  return `${h.slice(0, 8)}-${h.slice(8, 12)}-${h.slice(12, 16)}-${h.slice(16, 20)}-${h.slice(20)}`;
}
const norm = s => String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase();
const num = v => (v === '' || v === null || v === undefined || isNaN(+v)) ? null : +v;
const fmt = (v, d = 2) => v === null || v === undefined || isNaN(v) ? '—' : (+v).toLocaleString('es-PE', { minimumFractionDigits: d, maximumFractionDigits: 4 });
const fechaTxt = f => { const [y, m, d] = f.split('-'); return `${d}/${m}/${y}`; };
function toast(t, ms = 2600) { const e = $('toast'); e.textContent = t; e.classList.remove('hidden'); clearTimeout(toast.t); toast.t = setTimeout(() => e.classList.add('hidden'), ms); }
function msg(el, t, cls) { const e = $(el); if (!t) { e.classList.add('hidden'); return; } e.className = 'msg ' + cls; e.textContent = t; }
function esc(s) { return String(s ?? '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])); }

function metradoDe(r) {
  if (r.modo === 'dir') return num(r.met_directo);
  const v = [r.n_veces, r.largo, r.ancho, r.alto].map(num).filter(x => x !== null);
  if (!v.length) return null;
  return Math.round(v.reduce((a, b) => a * b, 1) * 1e6) / 1e6;   // igual que PRODUCT() del Excel
}
function estadoDe(r) { return S.ids.has(r.id) ? 'car' : r.estado === 'enviado' ? 'env' : 'pend'; }

/* saldo = saldo del catálogo − registros del teléfono que el Excel todavía no tiene */
function saldo(code, exceptId) {
  const p = S.byCode.get(code); if (!p) return null;
  let s = p.saldo;
  for (const r of S.recs) if (r.codigo === code && r.id !== exceptId && !S.ids.has(r.id)) s -= (metradoDe(r) || 0);
  return Math.round(s * 1e6) / 1e6;
}

/* ---------------- catálogo desde el Excel ---------------- */
async function cargarCatalogo(file) {
  msg('catMsg', 'Leyendo el libro… (puede tardar unos segundos)', 'warn');
  try {
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: 'array', sheets: [CAT_SHEET], cellFormula: false, cellHTML: false, cellText: false, cellStyles: false });
    const ws = wb.Sheets[CAT_SHEET];
    if (!ws) throw new Error(`El archivo no tiene la hoja «${CAT_SHEET}». Elija el libro SISTEMA VALORIZACIÓN POROTOBANGO actualizado.`);
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, range: 3, defval: '' });
    const h = rows[0].map(x => String(x).trim());
    const need = ['CODIGO', 'DESCRIPCION', 'UND', 'PU', 'CONTRACTUAL', 'EJECUTADO', 'SALDO', 'PRESUPUESTOS', 'PPTO_DEF', 'TITULO'];
    need.forEach((n, i) => { if (h[i] !== n) throw new Error('Formato de catálogo no reconocido (columna ' + n + ').'); });
    const items = []; const ids = [];
    let sinValor = 0;
    for (const r of rows.slice(1)) {
      const code = String(r[0] || '').trim();
      if (code) {
        if (r[6] === '' || isNaN(+r[6])) sinValor++;
        items.push({ code, desc: String(r[1]).trim(), und: String(r[2]).trim(), pu: +r[3] || 0, contr: +r[4] || 0, ejec: +r[5] || 0,
          saldo: +r[6] || 0, grupos: String(r[7] || '').split('+').filter(Boolean), def: String(r[8] || ''), tit: String(r[9] || '').trim() });
      }
      if (r[11]) ids.push(String(r[11]).trim());
    }
    if (!items.length) throw new Error(rows.length > 50 ? 'El libro no tiene los valores calculados (nunca se guardó en Excel). En la PC: abra el libro en Excel, espere que calcule, use Datos → Actualizar todo, GUARDE (Ctrl+G) y vuelva a cargarlo aquí.' : 'No se encontraron partidas en el catálogo.');
    if (sinValor > items.length / 2) throw new Error('El libro no tiene los saldos calculados. Ábralo en Excel, use Datos → Actualizar todo, GUARDE y vuelva a intentar.');
    const cat = { items, ids, file: file.name, cargado: ahoraISO() };
    await setMeta('cat', cat); aplicarCatalogo(cat);
    const car = S.recs.filter(r => S.ids.has(r.id)).length;
    msg('catMsg', `✔ Catálogo actualizado: ${items.length} partidas. ${ids.length} registros del celular ya están en el Excel (${car} de este teléfono).`, 'ok');
    renderTodo();
  } catch (e) { console.error(e); msg('catMsg', '✖ ' + (e.message || e), 'err'); }
}
function aplicarCatalogo(cat) {
  S.cat = cat; S.byCode = new Map(cat.items.map(p => [p.code, p])); S.ids = new Set(cat.ids);
  for (const p of cat.items) p._n = norm(p.code + ' ' + p.desc + ' ' + p.tit);
}

/* ---------------- búsqueda de partidas ---------------- */
function buscar(q) {
  const ul = $('resultados'); q = norm(q).trim();
  if (!S.cat || q.length < 2) { ul.classList.add('hidden'); return; }
  const words = q.split(/\s+/);
  // primero: código que empieza igual; luego partidas con saldo; al final deducidas / sin saldo
  const hits = [];
  S.cat.items.forEach((p, i) => {
    const pre = norm(p.code).startsWith(q);
    if (pre || words.every(w => p._n.includes(w))) hits.push([pre ? 0 : 1, p.pu > 0 && saldo(p.code) > 1e-9 ? 0 : 1, i, p]);
  });
  hits.sort((a, b) => a[0] - b[0] || a[1] - b[1] || a[2] - b[2]);
  const res = hits.slice(0, 60).map(h => h[3]);
  ul.innerHTML = res.length ? res.map(p => {
    const s = saldo(p.code); const cero = p.pu <= 0 || s <= 1e-9;
    return `<li data-c="${esc(p.code)}"><span class="s ${cero ? 'cero' : ''}">${p.pu <= 0 ? 'deducida' : 'saldo ' + fmt(s) + ' ' + esc(p.und)}</span>
      <span class="c">${esc(p.code)}</span><br>${esc(p.desc)}${p.tit ? `<br><span class="muted small">${esc(p.tit)}</span>` : ''}</li>`;
  }).join('') : '<li class="muted">Sin resultados</li>';
  ul.classList.remove('hidden');
}
function elegir(code) {
  const p = S.byCode.get(code); if (!p) return;
  S.sel = p; $('resultados').classList.add('hidden'); $('fBuscar').value = ''; $('fBuscar').parentElement.classList.add('hidden'); $('recientes').classList.add('hidden');
  $('partidaCard').classList.remove('hidden');
  $('pCodigo').textContent = p.code; $('pDesc').textContent = p.desc; $('pTitulo').textContent = p.tit;
  $('pUnd').textContent = p.und; $('pContr').textContent = fmt(p.contr); $('outUnd').textContent = p.und;
  const sel = $('fPpto');
  if (p.grupos.length > 1) {
    sel.innerHTML = `<option value="">Automático (recomendado)</option>` + p.grupos.map(g => `<option>${esc(g)}</option>`).join('');
    $('pptoWrap').classList.remove('hidden');
  } else { sel.innerHTML = ''; $('pptoWrap').classList.add('hidden'); }
  actualizarSaldo(); recalcular();
  if (!S.gps || Date.now() - S.gps.t > 5 * 60e3) tomarGps(true);
}
function quitarPartida() {
  S.sel = null; $('partidaCard').classList.add('hidden'); $('fBuscar').parentElement.classList.remove('hidden'); $('recientes').classList.remove('hidden');
  $('outUnd').textContent = ''; recalcular();
}
function actualizarSaldo() {
  if (!S.sel) return;
  const s = saldo(S.sel.code, $('editId').value || null);
  $('pSaldo').textContent = S.sel.pu <= 0 ? 'Deducida' : `${fmt(s)} ${S.sel.und}`;
  $('pSaldo').style.color = s <= 1e-9 || S.sel.pu <= 0 ? 'var(--red)' : 'var(--green)';
}
function renderRecientes() {
  const seen = []; for (const r of [...S.recs].sort((a, b) => b.creado.localeCompare(a.creado))) { if (!seen.includes(r.codigo)) seen.push(r.codigo); if (seen.length >= 8) break; }
  $('recientes').innerHTML = seen.filter(c => S.byCode.has(c)).map(c => `<button type="button" class="chip" data-c="${esc(c)}">${esc(c)}</button>`).join('');
  const fr = [...new Set(S.recs.map(r => r.frente).filter(Boolean))].slice(-30);
  $('dlFrentes').innerHTML = fr.map(f => `<option value="${esc(f)}">`).join('');
}

/* ---------------- formulario ---------------- */
function leerForm() {
  return { modo: S.modo, n_veces: $('fVeces').value, largo: $('fLargo').value, ancho: $('fAncho').value, alto: $('fAlto').value, met_directo: $('fDirecto').value };
}
function recalcular() {
  const m = metradoDe(leerForm());
  $('outMet').textContent = m === null ? '—' : fmt(m);
  validar(false);
}
function validar(final) {
  const f = leerForm(); const m = metradoDe(f); const p = S.sel;
  if (!S.cat) return 'Primero cargue el catálogo (Ajustes).';
  if (!$('fFecha').value) return 'Falta la fecha.';
  if (!p) { if (final) msg('msg', 'Elija la partida.', 'err'); else msg('msg'); return 'Elija la partida.'; }
  if (p.pu <= 0) { msg('msg', 'Esta partida está deducida (sin metrado vigente ni P.U.). El Excel la rechazaría.', 'err'); return 'deducida'; }
  if (m === null) { if (final) msg('msg', 'Escriba las medidas o el metrado.', 'err'); else msg('msg'); return 'Falta metrado'; }
  if (m === 0) { msg('msg', 'El metrado es cero.', 'err'); return 'cero'; }
  const s = saldo(p.code, $('editId').value || null);
  if (m > s + 1e-4) { msg('msg', `✖ Supera el metrado contractual: el saldo es ${fmt(s)} ${p.und} — el Excel lo rechazaría; corrija el metrado.`, 'err'); return 'sobre'; }
  if (m < 0) { msg('msg', 'Metrado negativo: se registrará como corrección (descuenta metrado).', 'warn'); return null; }
  msg('msg', `✔ Dentro del saldo. Quedará: ${fmt(s - m)} ${p.und}`, 'ok'); return null;
}
function setModo(m) {
  S.modo = m; $('modoMed').classList.toggle('on', m === 'med'); $('modoDir').classList.toggle('on', m === 'dir');
  $('boxMed').classList.toggle('hidden', m !== 'med'); $('boxDir').classList.toggle('hidden', m !== 'dir'); recalcular();
}
function limpiar(mantenerFecha = true) {
  const f = $('fFecha').value, fr = $('fFrente').value;
  $('frm').reset(); $('editId').value = '';
  $('fFecha').value = mantenerFecha && f ? f : hoy(); $('fFrente').value = mantenerFecha ? fr : '';
  S.fotosNuevas = []; S.fotosQuitar = []; S.fotosExist = []; renderThumbs(); quitarPartida(); setModo('med'); msg('msg');
  $('btnGuardar').textContent = 'Guardar registro';
}
async function guardar(ev) {
  ev.preventDefault();
  if (!S.user) { toast('Primero escriba su nombre en Ajustes'); cambiarTab('tab-cfg'); return; }
  const err = validar(true); if (err) { if (!$('msg').textContent) msg('msg', err, 'err'); return; }
  const f = leerForm(); const m = metradoDe(f);
  if (m < 0 && !confirm('El metrado es negativo (corrección). ¿Guardar?')) return;
  const editId = $('editId').value;
  const old = editId ? S.recs.find(r => r.id === editId) : null;
  const r = old ? { ...old } : { id: uuid(), creado: ahoraISO(), estado: 'pendiente', usuario: S.user, fotos: [] };
  Object.assign(r, { fecha: $('fFecha').value, codigo: S.sel.code, ppto: $('fPpto').value || '', frente: $('fFrente').value.trim(),
    modo: f.modo, n_veces: f.modo === 'med' ? num(f.n_veces) : null, largo: f.modo === 'med' ? num(f.largo) : null,
    ancho: f.modo === 'med' ? num(f.ancho) : null, alto: f.modo === 'med' ? num(f.alto) : null,
    met_directo: f.modo === 'dir' ? num(f.met_directo) : null, observacion: $('fObs').value.trim(), editado: old ? ahoraISO() : null });
  if (!old && S.gps) Object.assign(r, { lat: S.gps.lat, lon: S.gps.lon, acc: S.gps.acc });
  for (const fid of S.fotosQuitar) { await del('fotos', fid); r.fotos = r.fotos.filter(x => x !== fid); }
  let n = r.fotos.length;
  for (const ft of S.fotosNuevas) {
    n++; const fid = uuid();
    const name = `FOTO_${r.fecha.replace(/-/g, '')}_${r.codigo.replace(/\./g, '-')}_${r.id.slice(0, 8)}_${n}.jpg`;
    await put('fotos', { id: fid, rec: r.id, name, buf: ft.buf, type: 'image/jpeg', enviada: false });
    r.fotos.push(fid);
  }
  await put('recs', r);
  if (old) S.recs[S.recs.indexOf(old)] = r; else S.recs.push(r);
  toast(old ? 'Registro actualizado' : `✔ Guardado: ${r.codigo} · ${fmt(m)} ${S.sel.und}`);
  limpiar(true); renderTodo();
}

/* ---------------- fotos ---------------- */
function comprimir(file) {
  return new Promise((ok, ko) => {
    const url = URL.createObjectURL(file); const img = new Image();
    img.onload = () => {
      const k = Math.min(1, FOTO_LADO / Math.max(img.naturalWidth, img.naturalHeight));
      const c = document.createElement('canvas'); c.width = Math.round(img.naturalWidth * k); c.height = Math.round(img.naturalHeight * k);
      c.getContext('2d').drawImage(img, 0, 0, c.width, c.height); URL.revokeObjectURL(url);
      c.toBlob(b => b ? b.arrayBuffer().then(buf => ok({ buf, url: URL.createObjectURL(b) })) : ko(new Error('No se pudo procesar la foto')), 'image/jpeg', FOTO_CAL);
    };
    img.onerror = () => { URL.revokeObjectURL(url); ko(new Error('Formato de foto no soportado')); };
    img.src = url;
  });
}
async function agregarFoto(e) {
  const file = e.target.files[0]; e.target.value = ''; if (!file) return;
  const total = (S.fotosExist || []).length - S.fotosQuitar.length + S.fotosNuevas.length;
  if (total >= MAX_FOTOS) { toast(`Máximo ${MAX_FOTOS} fotos por registro`); return; }
  try { S.fotosNuevas.push(await comprimir(file)); renderThumbs(); } catch (err) { toast(err.message); }
}
function renderThumbs() {
  const ex = (S.fotosExist || []).filter(f => !S.fotosQuitar.includes(f.id));
  $('thumbs').innerHTML = ex.map(f => `<div><img src="${f.url}"><button type="button" data-q="${f.id}">×</button></div>`).join('') +
    S.fotosNuevas.map((f, i) => `<div><img src="${f.url}"><button type="button" data-n="${i}">×</button></div>`).join('');
  $('lblFoto').classList.toggle('hidden', ex.length + S.fotosNuevas.length >= MAX_FOTOS);
}

/* ---------------- GPS ---------------- */
function tomarGps(silencioso) {
  if (!navigator.geolocation) { $('gpsTxt').textContent = 'Ubicación: no disponible'; return; }
  $('gpsTxt').textContent = 'Ubicación: buscando señal GPS…';
  navigator.geolocation.getCurrentPosition(p => {
    S.gps = { lat: +p.coords.latitude.toFixed(6), lon: +p.coords.longitude.toFixed(6), acc: Math.round(p.coords.accuracy), t: Date.now() };
    $('gpsTxt').textContent = `Ubicación: ${S.gps.lat}, ${S.gps.lon} (±${S.gps.acc} m)`;
  }, err => {
    $('gpsTxt').textContent = 'Ubicación: ' + (err.code === 1 ? 'permiso denegado (actívelo en Ajustes del teléfono)' : 'sin señal GPS; se guardará sin ubicación');
    if (!silencioso) toast('No se pudo obtener la ubicación');
  }, { enableHighAccuracy: true, timeout: 20000, maximumAge: 60000 });
}

/* ---------------- lista de registros ---------------- */
function renderLista() {
  const filtro = $('lstFiltro').value, q = norm($('lstBuscar').value);
  let rs = [...S.recs].sort((a, b) => (b.fecha + b.creado).localeCompare(a.fecha + a.creado));
  if (filtro === 'pend') rs = rs.filter(r => estadoDe(r) === 'pend');
  if (filtro === 'env') rs = rs.filter(r => estadoDe(r) !== 'pend');
  if (q) rs = rs.filter(r => norm(r.codigo + ' ' + (S.byCode.get(r.codigo)?.desc || '') + ' ' + r.frente + ' ' + r.observacion).includes(q));
  const cnt = { pend: 0, env: 0, car: 0 }; S.recs.forEach(r => cnt[estadoDe(r)]++);
  $('lstResumen').textContent = `En el teléfono: ${S.recs.length} registros · ${cnt.pend} pendientes de enviar · ${cnt.env} enviados · ${cnt.car} ya cargados en el Excel`;
  let html = '', dia = '';
  for (const r of rs) {
    if (r.fecha !== dia) { dia = r.fecha; html += `<div class="day">${fechaTxt(dia)}</div>`; }
    const p = S.byCode.get(r.codigo) || { desc: '(partida no está en el catálogo)', und: '' };
    const est = estadoDe(r); const m = metradoDe(r);
    const med = r.modo === 'med' ? [r.n_veces, r.largo, r.ancho, r.alto].filter(v => v !== null).map(v => fmt(v)).join(' × ') : 'directo';
    html += `<div class="rec ${est}"><div class="l1"><span class="code">${esc(r.codigo)}${r.ppto ? ' · ' + esc(r.ppto) : ''}</span><span class="met">${fmt(m)} ${esc(p.und)}</span></div>
      <div class="d">${esc(p.desc)}</div>
      <div class="l3"><span>${esc(r.frente || '')}${r.frente ? ' · ' : ''}${med}${r.observacion ? ' · ' + esc(r.observacion) : ''}</span>
      <span>${r.fotos.length ? '📷' + r.fotos.length + ' ' : ''}${r.lat ? '📍 ' : ''}<span class="tag ${est}">${est === 'pend' ? 'pendiente' : est === 'env' ? 'enviado' : 'en el Excel'}</span></span></div>
      <div class="btns">${est === 'pend' ? `<button data-e="${r.id}">Editar</button><button data-d="${r.id}">Eliminar</button>` : `<button data-k="${r.id}">Registrar corrección</button>`}</div></div>`;
  }
  $('lista').innerHTML = html || '<div class="card muted">No hay registros para mostrar.</div>';
}
async function editar(id) {
  const r = S.recs.find(x => x.id === id); if (!r) return;
  limpiar(false); cambiarTab('tab-reg');
  $('editId').value = r.id; $('fFecha').value = r.fecha; $('fFrente').value = r.frente; $('fObs').value = r.observacion;
  elegir(r.codigo); $('fPpto').value = r.ppto || '';
  setModo(r.modo);
  $('fVeces').value = r.n_veces ?? ''; $('fLargo').value = r.largo ?? ''; $('fAncho').value = r.ancho ?? ''; $('fAlto').value = r.alto ?? ''; $('fDirecto').value = r.met_directo ?? '';
  S.fotosExist = [];
  for (const fid of r.fotos) { const f = await tx('fotos', 'readonly', st => rqp(st.get(fid))); if (f) S.fotosExist.push({ id: fid, url: URL.createObjectURL(new Blob([f.buf], { type: 'image/jpeg' })) }); }
  renderThumbs(); recalcular(); $('btnGuardar').textContent = 'Guardar cambios';
}
async function eliminar(id) {
  const r = S.recs.find(x => x.id === id); if (!r || !confirm(`¿Eliminar el registro ${r.codigo} del ${fechaTxt(r.fecha)}?`)) return;
  for (const fid of r.fotos) await del('fotos', fid);
  await del('recs', id); S.recs = S.recs.filter(x => x.id !== id); toast('Registro eliminado'); renderTodo();
}
function correccion(id) {
  const r = S.recs.find(x => x.id === id); if (!r) return;
  limpiar(false); cambiarTab('tab-reg'); $('fFecha').value = r.fecha; $('fFrente').value = r.frente;   // misma fecha (mismo mes) que el registro que corrige
  elegir(r.codigo); setModo('dir'); $('fDirecto').value = -(metradoDe(r) || 0);
  $('fObs').value = `Corrección del registro del ${fechaTxt(r.fecha)} (${fmt(metradoDe(r))})`; recalcular();
  toast('Misma fecha del registro corregido. Ajuste el valor negativo y guarde');
}

/* ---------------- envío ---------------- */
function csvCampo(v) { if (v === null || v === undefined) return ''; const s = String(v); return /[",\n\r]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s; }
async function nombresFotos(r) {
  const out = []; for (const fid of r.fotos) { const f = await tx('fotos', 'readonly', st => rqp(st.get(fid))); if (f) out.push(f.name); } return out;
}
async function prepararEnvio() {
  const pend = S.recs.filter(r => estadoDe(r) === 'pend').sort((a, b) => a.creado.localeCompare(b.creado));
  const lines = [CSV_COLS.join(',')];
  for (const r of pend) {
    const fotos = (await nombresFotos(r)).join(' | ');
    const row = { ID: r.id, FECHA: r.fecha, CODIGO: r.codigo, PPTO: r.ppto, FRENTE: r.frente, N_VECES: r.n_veces, LARGO: r.largo, ANCHO: r.ancho, ALTO: r.alto,
      MET_DIRECTO: r.met_directo, OBSERVACION: r.observacion, USUARIO: r.usuario, LATITUD: r.lat, LONGITUD: r.lon, PRECISION_M: r.acc, FOTOS: fotos, CREADO: r.creado };
    lines.push(CSV_COLS.map(c => csvCampo(row[c])).join(','));
  }
  const d = new Date(); const stamp = `${hoy().replace(/-/g, '')}_${pad(d.getHours())}${pad(d.getMinutes())}${pad(d.getSeconds())}`;
  const user = norm(S.user).replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '').toUpperCase() || 'USUARIO';
  const name = `campo_${stamp}_${user}.csv`;
  const file = new File([lines.join('\r\n') + '\r\n'], name, { type: 'text/csv' });
  // fotos pendientes (próximo lote)
  const fotos = (await getAll('fotos')).filter(f => !f.enviada).sort((a, b) => a.name.localeCompare(b.name));
  const lote = fotos.slice(0, FOTOS_POR_ENVIO).map(f => ({ id: f.id, file: new File([f.buf], f.name, { type: 'image/jpeg' }) }));
  S.prep = { pend, file, name, fotosPend: fotos.length, lote };
  const kb = Math.max(1, Math.round(file.size / 1024));
  const mb = (fotos.reduce((a, f) => a + f.buf.byteLength, 0) / 1048576).toFixed(1);
  $('sendResumen').innerHTML = `<b>${pend.length}</b> registros pendientes → archivo <i>${esc(name)}</i> (${kb} KB)<br><b>${fotos.length}</b> fotos por enviar (${mb} MB), en envíos de ${FOTOS_POR_ENVIO}.`;
  $('btnEnviarDatos').disabled = !pend.length; $('btnEnviarFotos').disabled = !fotos.length;
  $('btnEnviarFotos').textContent = fotos.length ? `2. Enviar fotos (${Math.min(FOTOS_POR_ENVIO, fotos.length)} de ${fotos.length})` : '2. Enviar fotos';
}
function puedeCompartir(files) { try { return !!(navigator.canShare && navigator.canShare({ files })); } catch (e) { return false; } }
function descargar(file) {
  const a = document.createElement('a'); a.href = URL.createObjectURL(file); a.download = file.name; document.body.appendChild(a); a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 4000);
}
async function compartir(files, titulo) {
  if (puedeCompartir(files)) {
    try { await navigator.share({ files, title: titulo }); return true; }
    catch (e) { if (e.name === 'AbortError') return false; throw e; }
  }
  files.forEach(descargar);   // sin menú compartir: se descarga el archivo
  return confirm('El archivo se descargó. ¿Ya lo subió a Google Drive → CAMPO POROTOBANGO?');
}
async function enviarDatos() {
  const P = S.prep; if (!P || !P.pend.length) return;
  msg('sendMsg');
  try {
    const hecho = await compartir([P.file], P.name);   // llamado directo desde el toque (requisito del iPhone)
    if (!hecho) { msg('sendMsg', 'Envío cancelado. Los registros siguen pendientes.', 'warn'); return; }
    if (!confirm('¿Guardó el archivo en Google Drive → CAMPO POROTOBANGO? (Aceptar = marcar como enviados)')) { msg('sendMsg', 'No se marcó nada: puede volver a enviar.', 'warn'); return; }
    const lote = { id: uuid(), name: P.name, cuando: ahoraISO(), n: P.pend.length, ids: P.pend.map(r => r.id) };
    for (const r of P.pend) { r.estado = 'enviado'; r.lote = lote.id; await put('recs', r); }
    S.lotes.unshift(lote); await setMeta('lotes', S.lotes.slice(0, 200));
    msg('sendMsg', `✔ ${P.pend.length} registros enviados en ${P.name}. En la PC: Datos → Actualizar todo.`, 'ok');
  } catch (e) { msg('sendMsg', '✖ No se pudo compartir: ' + e.message, 'err'); }
  renderTodo();
}
async function enviarFotos() {
  const P = S.prep; if (!P || !P.lote.length) return;
  try {
    const hecho = await compartir(P.lote.map(x => x.file), 'Fotos de metrados');
    if (!hecho) { msg('sendMsg', 'Envío de fotos cancelado.', 'warn'); return; }
    for (const x of P.lote) { const f = await tx('fotos', 'readonly', st => rqp(st.get(x.id))); if (f) { f.enviada = true; await put('fotos', f); } }
    msg('sendMsg', `✔ ${P.lote.length} fotos enviadas.` + (P.fotosPend > P.lote.length ? ' Toque otra vez para el siguiente grupo.' : ''), 'ok');
  } catch (e) { msg('sendMsg', '✖ No se pudo compartir: ' + e.message, 'err'); }
  renderTodo();
}
async function reenviarLote(id) {
  const L = S.lotes.find(l => l.id === id); if (!L) return;
  const rs = S.recs.filter(r => L.ids.includes(r.id));
  const lines = [CSV_COLS.join(',')];
  for (const r of rs) {
    const row = { ID: r.id, FECHA: r.fecha, CODIGO: r.codigo, PPTO: r.ppto, FRENTE: r.frente, N_VECES: r.n_veces, LARGO: r.largo, ANCHO: r.ancho, ALTO: r.alto,
      MET_DIRECTO: r.met_directo, OBSERVACION: r.observacion, USUARIO: r.usuario, LATITUD: r.lat, LONGITUD: r.lon, PRECISION_M: r.acc, FOTOS: (await nombresFotos(r)).join(' | '), CREADO: r.creado };
    lines.push(CSV_COLS.map(c => csvCampo(row[c])).join(','));
  }
  const f = new File([lines.join('\r\n') + '\r\n'], L.name, { type: 'text/csv' });
  try { await compartir([f], L.name); } catch (e) { toast(e.message); }
}
function renderLotes() {
  $('lotes').innerHTML = S.lotes.length ? S.lotes.slice(0, 30).map(l => {
    const car = l.ids.filter(i => S.ids.has(i)).length;
    return `<div class="lote">${l.cuando.replace('T', ' ')} · <b>${l.n}</b> registros · ${car === l.n ? '✔ en el Excel' : car ? car + ' en el Excel' : 'aún no cargado en el Excel'}<br>
      <span class="muted">${esc(l.name)}</span> <button class="link" data-l="${l.id}">volver a compartir</button></div>`;
  }).join('') : '<span class="muted">Aún no hay envíos.</span>';
}

/* ---------------- ajustes ---------------- */
async function renderCfg() {
  $('cfgUsuario').value = S.user;
  $('catInfo').innerHTML = S.cat ? `✔ <b>${S.cat.items.length}</b> partidas · cargado el ${S.cat.cargado.replace('T', ' ')}<br>Archivo: ${esc(S.cat.file)} · ${S.cat.ids.length} registros del celular ya en el Excel`
    : '<b style="color:var(--red)">Sin catálogo</b>';
  let t = `Registros: ${S.recs.length}`;
  if (navigator.storage && navigator.storage.estimate) { const e = await navigator.storage.estimate(); t += ` · espacio usado ${(e.usage / 1048576).toFixed(1)} MB`; }
  $('stoInfo').textContent = t;
}
async function limpiarCargados() {
  const rs = S.recs.filter(r => S.ids.has(r.id));
  const fotos = await getAll('fotos'); const pendF = new Set(fotos.filter(f => !f.enviada).map(f => f.rec));
  const borrar = rs.filter(r => !pendF.has(r.id));
  if (!borrar.length) { toast('No hay registros cargados en el Excel para borrar'); return; }
  if (!confirm(`Se borrarán del teléfono ${borrar.length} registros que YA están en el Excel (y sus fotos ya enviadas). ¿Continuar?`)) return;
  for (const r of borrar) { for (const fid of r.fotos) await del('fotos', fid); await del('recs', r.id); }
  const ids = new Set(borrar.map(r => r.id)); S.recs = S.recs.filter(r => !ids.has(r.id)); toast(`${borrar.length} registros borrados del teléfono`); renderTodo();
}
function respaldo() {
  const f = new File([JSON.stringify({ version: APP_VERSION, usuario: S.user, exportado: ahoraISO(), registros: S.recs }, null, 1)], `respaldo_metrados_${hoy()}.json`, { type: 'application/json' });
  compartir([f], f.name).catch(e => toast(e.message));
}

/* ---------------- navegación y arranque ---------------- */
function cambiarTab(id) {
  document.querySelectorAll('.tab').forEach(t => t.classList.toggle('active', t.id === id));
  document.querySelectorAll('nav.bottom button').forEach(b => b.classList.toggle('on', b.dataset.tab === id));
  if (id === 'tab-list') renderLista(); if (id === 'tab-send') { prepararEnvio(); renderLotes(); } if (id === 'tab-cfg') renderCfg();
  window.scrollTo(0, 0);
}
function renderTodo() {
  const pend = S.recs.filter(r => estadoDe(r) === 'pend').length;
  $('pendBadge').textContent = pend || '';
  $('hdrInfo').textContent = `${S.user || 'sin usuario'} · ${S.cat ? S.cat.items.length + ' partidas' : 'sin catálogo'} · ${pend} por enviar`;
  $('noCat').classList.toggle('hidden', !!S.cat);
  renderRecientes(); actualizarSaldo();
  const act = document.querySelector('.tab.active').id;
  if (act === 'tab-list') renderLista(); if (act === 'tab-send') { prepararEnvio(); renderLotes(); } if (act === 'tab-cfg') renderCfg();
}
function red() { const on = navigator.onLine; const b = $('netBadge'); b.textContent = on ? 'con señal' : 'sin señal'; b.className = 'badge ' + (on ? 'on' : 'off'); }

async function init() {
  $('ver').textContent = APP_VERSION;
  DB = await idb();
  S.recs = await getAll('recs'); S.user = (await getMeta('user')) || ''; S.lotes = (await getMeta('lotes')) || [];
  const cat = await getMeta('cat'); if (cat) aplicarCatalogo(cat);
  $('fFecha').value = hoy();
  // eventos
  document.querySelectorAll('nav.bottom button').forEach(b => b.onclick = () => cambiarTab(b.dataset.tab));
  $('fBuscar').oninput = e => buscar(e.target.value);
  $('resultados').onclick = e => { const li = e.target.closest('li[data-c]'); if (li) elegir(li.dataset.c); };
  $('recientes').onclick = e => { const b = e.target.closest('[data-c]'); if (b) elegir(b.dataset.c); };
  $('btnCambiar').onclick = quitarPartida;
  $('modoMed').onclick = () => setModo('med'); $('modoDir').onclick = () => setModo('dir');
  ['fVeces', 'fLargo', 'fAncho', 'fAlto', 'fDirecto'].forEach(i => $(i).oninput = recalcular);
  $('fFecha').onchange = () => validar(false);
  $('frm').onsubmit = guardar; $('btnLimpiar').onclick = () => limpiar(false);
  $('btnGps').onclick = () => tomarGps(false);
  $('fFoto').onchange = agregarFoto;
  $('thumbs').onclick = e => {
    const b = e.target.closest('button'); if (!b) return;
    if (b.dataset.n !== undefined) S.fotosNuevas.splice(+b.dataset.n, 1); else S.fotosQuitar.push(b.dataset.q);
    renderThumbs();
  };
  $('lstFiltro').onchange = renderLista; $('lstBuscar').oninput = renderLista;
  $('lista').onclick = e => { const b = e.target.closest('button'); if (!b) return; if (b.dataset.e) editar(b.dataset.e); if (b.dataset.d) eliminar(b.dataset.d); if (b.dataset.k) correccion(b.dataset.k); };
  $('btnEnviarDatos').onclick = enviarDatos; $('btnEnviarFotos').onclick = enviarFotos;
  $('lotes').onclick = e => { const b = e.target.closest('[data-l]'); if (b) reenviarLote(b.dataset.l); };
  $('btnUsuario').onclick = async () => { S.user = $('cfgUsuario').value.trim().toUpperCase(); await setMeta('user', S.user); toast('Nombre guardado'); renderTodo(); };
  $('cfgXlsx').onchange = e => { const f = e.target.files[0]; e.target.value = ''; if (f) cargarCatalogo(f); };
  $('btnRespaldo').onclick = respaldo; $('btnLimpiarCargados').onclick = limpiarCargados;
  window.addEventListener('online', red); window.addEventListener('offline', red); red();
  if (navigator.storage && navigator.storage.persist) navigator.storage.persist().catch(() => {});
  if ('serviceWorker' in navigator) navigator.serviceWorker.register('sw.js').catch(() => {});
  S.fotosExist = [];
  renderTodo();
  if (!S.user || !S.cat) cambiarTab('tab-cfg');
}
document.addEventListener('DOMContentLoaded', () => init().catch(e => { console.error(e); alert('Error al iniciar: ' + e.message); }));
