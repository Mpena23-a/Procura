/* =========================================================================
   COLVEN · Dashboard de Seguimiento de Equipos y Mantenimiento
   Lee el Excel (misma estructura) y construye el reporte en el navegador.
   ========================================================================= */
(function () {
  'use strict';

  const CFG = window.DASHBOARD_CONFIG;
  const $ = (s, el = document) => el.querySelector(s);
  const VIEWS = ['resumen', 'mtto', 'oper', 'mat', 'rq'];

  const charts = {};       // id -> instancia Chart.js
  let builders = {};       // vista -> función que crea sus gráficos (se ejecuta al abrir la vista)
  let built = {};          // vista -> ya construida
  let drills = {};         // id de cuadro -> función que devuelve la especificación del panel de detalle
  let openDrill = {};      // slot -> id del cuadro abierto
  let drillSpecs = {};     // slot -> especificación abierta (para su botón de acción)
  const tableState = {};   // filtros y orden por segmento
  let DATA = null;
  let downloadUrl = null;
  let curView = null;

  /* ---------------- Paleta ---------------- */
  const C = { navy: '#1D2041', navy3: '#3C4180', yellow: '#FDD649', gold: '#E3A92B', muted: '#646985', line: '#E2E4EE', ok: '#2E9E6B', risk: '#D9534F' };
  const ESTADOS = [
    { key: 'Completado',                color: '#2E9E6B' },
    { key: 'En ejecución',              color: '#3C4180' },
    { key: 'Entrega pendiente',         color: '#5B8DEF' },
    { key: 'Pago / anticipo pendiente', color: '#E3A92B' },
    { key: 'Por aprobación',            color: '#FDD649' },
    { key: 'En gestión (sin orden)',    color: '#9AA0BF' },
    { key: 'Canalizado / no aprobado',  color: '#C9CCDD' },
    { key: 'En seguimiento',            color: '#7F86B5' },
    { key: 'Sin comentario',            color: '#E6E8F0' }
  ];
  const estadoColor = k => (ESTADOS.find(e => e.key === k) || {}).color || '#ccc';
  const REGION_COLORS = { 'Apure': '#1D2041', 'Barinas': '#FDD649', 'Apure / Barinas': '#7F86B5', 'Sin definir': '#D5D8E5' };
  const PRIO_COLORS = { 'Alta': '#D9534F', 'Media': '#FDD649', 'Baja': '#2E9E6B', 'Sin definir': '#D5D8E5' };
  const PAL = ['#1D2041', '#FDD649', '#3C4180', '#E3A92B', '#7F86B5', '#9AA0BF', '#C9CCDD', '#5B8DEF'];
  const CHEV = '<svg viewBox="0 0 24 24"><path d="m6 9 6 6 6-6"/></svg>';

  /* ---------------- Utilidades ---------------- */
  const norm = s => String(s ?? '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().replace(/\s+/g, ' ').trim();
  const clean = s => String(s ?? '').replace(/\s+/g, ' ').trim();
  const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const fmtInt = n => new Intl.NumberFormat('es-VE').format(n || 0);
  const fmtMoney = n => new Intl.NumberFormat('es-VE', { style: 'currency', currency: CFG.currency, maximumFractionDigits: 0 }).format(n || 0);
  const fmtMoneyShort = n => {
    const a = Math.abs(n || 0);
    if (a >= 1e6) return '$' + (n / 1e6).toFixed(1).replace('.', ',') + ' M';
    if (a >= 1e3) return '$' + (n / 1e3).toFixed(1).replace('.', ',') + ' K';
    return '$' + Math.round(n || 0);
  };
  const pct = (a, b) => b ? Math.round((a / b) * 100) : 0;
  const fmtDate = d => d ? d.toLocaleDateString('es-VE', { day: '2-digit', month: 'short', year: 'numeric', timeZone: 'UTC' }) : '';
  const today = () => { const t = new Date(); return Date.UTC(t.getFullYear(), t.getMonth(), t.getDate()); };
  const short = (s, n = 26) => s && s.length > n ? s.slice(0, n - 1) + '…' : s;
  const plural = (n, a, b) => n === 1 ? a : b;
  const debounce = (fn, ms) => { let t; return (...a) => { clearTimeout(t); t = setTimeout(() => fn(...a), ms); }; };

  function toDate(v) {
    if (v === null || v === undefined || v === '') return null;
    if (v instanceof Date) return new Date(Date.UTC(v.getFullYear(), v.getMonth(), v.getDate()));
    if (typeof v === 'number' && v > 20000 && v < 80000) return new Date(Math.round((v - 25569) * 86400000));
    const m = String(v).trim().match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})$/);
    if (m) { let y = +m[3]; if (y < 100) y += 2000; return new Date(Date.UTC(y, +m[2] - 1, +m[1])); }
    return null;
  }
  const daysSince = d => d ? Math.max(0, Math.floor((today() - d.getTime()) / 86400000)) : null;
  const countBy = (arr, fn) => arr.reduce((m, x) => { const k = fn(x); m[k] = (m[k] || 0) + 1; return m; }, {});
  const sumBy = (arr, fn) => arr.reduce((s, x) => s + (fn(x) || 0), 0);
  const titleCase = s => s ? s.charAt(0).toUpperCase() + s.slice(1).toLowerCase() : '';

  /* ---------------- SHA-256 (para la contraseña de carga) ---------------- */
  const K256 = [0x428a2f98, 0x71374491, 0xb5c0fbcf, 0xe9b5dba5, 0x3956c25b, 0x59f111f1, 0x923f82a4, 0xab1c5ed5, 0xd807aa98, 0x12835b01, 0x243185be, 0x550c7dc3, 0x72be5d74, 0x80deb1fe, 0x9bdc06a7, 0xc19bf174, 0xe49b69c1, 0xefbe4786, 0x0fc19dc6, 0x240ca1cc, 0x2de92c6f, 0x4a7484aa, 0x5cb0a9dc, 0x76f988da, 0x983e5152, 0xa831c66d, 0xb00327c8, 0xbf597fc7, 0xc6e00bf3, 0xd5a79147, 0x06ca6351, 0x14292967, 0x27b70a85, 0x2e1b2138, 0x4d2c6dfc, 0x53380d13, 0x650a7354, 0x766a0abb, 0x81c2c92e, 0x92722c85, 0xa2bfe8a1, 0xa81a664b, 0xc24b8b70, 0xc76c51a3, 0xd192e819, 0xd6990624, 0xf40e3585, 0x106aa070, 0x19a4c116, 0x1e376c08, 0x2748774c, 0x34b0bcb5, 0x391c0cb3, 0x4ed8aa4a, 0x5b9cca4f, 0x682e6ff3, 0x748f82ee, 0x78a5636f, 0x84c87814, 0x8cc70208, 0x90befffa, 0xa4506ceb, 0xbef9a3f7, 0xc67178f2];
  function sha256Hex(str) {
    const bytes = new TextEncoder().encode(str);
    const len = bytes.length, total = ((len + 9 + 63) >> 6) << 6;
    const m = new Uint8Array(total); m.set(bytes); m[len] = 0x80;
    const dv = new DataView(m.buffer);
    dv.setUint32(total - 4, (len * 8) >>> 0); dv.setUint32(total - 8, Math.floor(len / 0x20000000));
    let H = [0x6a09e667, 0xbb67ae85, 0x3c6ef372, 0xa54ff53a, 0x510e527f, 0x9b05688c, 0x1f83d9ab, 0x5be0cd19];
    const w = new Uint32Array(64), rot = (x, n) => (x >>> n) | (x << (32 - n));
    for (let i = 0; i < total; i += 64) {
      for (let t = 0; t < 16; t++) w[t] = dv.getUint32(i + t * 4);
      for (let t = 16; t < 64; t++) {
        const s0 = rot(w[t - 15], 7) ^ rot(w[t - 15], 18) ^ (w[t - 15] >>> 3);
        const s1 = rot(w[t - 2], 17) ^ rot(w[t - 2], 19) ^ (w[t - 2] >>> 10);
        w[t] = (w[t - 16] + s0 + w[t - 7] + s1) >>> 0;
      }
      let [a, b, c, d, e, f, g, h] = H;
      for (let t = 0; t < 64; t++) {
        const t1 = (h + (rot(e, 6) ^ rot(e, 11) ^ rot(e, 25)) + ((e & f) ^ (~e & g)) + K256[t] + w[t]) >>> 0;
        const t2 = ((rot(a, 2) ^ rot(a, 13) ^ rot(a, 22)) + ((a & b) ^ (a & c) ^ (b & c))) >>> 0;
        h = g; g = f; f = e; e = (d + t1) >>> 0; d = c; c = b; b = a; a = (t1 + t2) >>> 0;
      }
      H = [a, b, c, d, e, f, g, h].map((v, k) => (H[k] + v) >>> 0);
    }
    return H.map(x => x.toString(16).padStart(8, '0')).join('');
  }

  /* ---------------- Reglas de negocio ---------------- */
  function region(orden, texto) {
    const m = String(orden || '').toUpperCase().match(/^OD[SC]-([A-Z]{2})-/);
    const code = m ? m[1] : '';
    if (code === 'AP') return 'Apure';
    if (code === 'BA') return 'Barinas';
    if (code === 'AB') return 'Apure / Barinas';
    const t = norm(texto);
    const ap = t.includes('apure'), ba = t.includes('barinas');
    if (ap && ba) return 'Apure / Barinas';
    if (ap) return 'Apure';
    if (ba) return 'Barinas';
    return 'Sin definir';
  }

  // Estado estimado a partir del campo "Comentarios" y de si existe número de orden.
  function inferEstado(coment, hasOrder) {
    const c = norm(coment);
    if (/no fue aprobad|canaliz|anulad|cancelad/.test(c)) return 'Canalizado / no aprobado';
    if (/trabajo 100|completad|culminad|finalizad|cerrad/.test(c)) return 'Completado';
    if (/por aprob/.test(c)) return 'Por aprobación';
    if (/por anticipo|pendiente pago|pendiente de pago|sin informacion de pago|para inicio/.test(c)) return 'Pago / anticipo pendiente';
    if (/entrega pendiente|fabricacion|por entregar|en transito/.test(c)) return 'Entrega pendiente';
    if (!hasOrder) return 'En gestión (sin orden)';
    if (/anticipo|pago|ejecucion|termina|extender|otrosi|valuacion|enviada/.test(c)) return 'En ejecución';
    if (!c) return 'Sin comentario';
    return 'En seguimiento';
  }
  const ABIERTO = e => e !== 'Completado' && e !== 'Canalizado / no aprobado';

  /* ---------------- Lectura del Excel ---------------- */
  const sheetRows = ws => XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '', blankrows: false });
  function findHeaderRow(rows) {
    for (let i = 0; i < Math.min(rows.length, 15); i++) {
      const cells = rows[i].map(norm).filter(Boolean);
      if (cells.length >= 3 && cells.some(c => c.startsWith('no'))) return i;
    }
    return 0;
  }

  function pickSheets(wb) {
    const hidden = new Set(((wb.Workbook || {}).Sheets || []).filter(s => s.Hidden).map(s => s.name));
    const visible = wb.SheetNames.filter(n => !hidden.has(n));
    const out = {}, used = new Set();
    for (const key of ['mtto', 'oper', 'mat', 'rq']) {
      const name = visible.find(n => !used.has(n) && CFG.sheets[key].test(n));
      if (name) { out[key] = name; used.add(name); }
    }
    const rest = visible.filter(n => !used.has(n));
    for (const key of ['mtto', 'oper', 'mat', 'rq']) if (!out[key] && rest.length) out[key] = rest.shift();
    return out;
  }

  function parseOrders(ws, segLabel) {
    const rows = sheetRows(ws);
    const hi = findHeaderRow(rows);
    const rawH = (rows[hi] || []).map(clean);
    const h = rawH.map(norm);
    const f = t => h.findIndex(t);
    const ix = {
      no: f(c => /^no\.?$/.test(c)),
      orden: f(c => c.includes('orden') && !c.includes('fecha')),
      solicitud: f(c => c === 'solicitud' || c === 'descripcion'),
      empresa: f(c => c.includes('empresa') || c.includes('proveedor')),
      tipo: f(c => c.includes('tipo')),
      prioridad: f(c => c.includes('prioridad')),
      fSol: f(c => c.includes('fecha') && c.includes('solicitud')),
      fEm: f(c => c.includes('fecha') && (c.includes('emision') || c.includes('orden'))),
      owner: f(c => c.startsWith('encargado') && !c.includes('procura')),
      procura: f(c => c.includes('procura')),
      coment: f(c => c.includes('comentario') || c.includes('observ'))
    };
    const get = (r, k) => ix[k] >= 0 ? r[ix[k]] : '';
    const items = [];
    for (const r of rows.slice(hi + 1)) {
      const orden = clean(get(r, 'orden')), solicitud = clean(get(r, 'solicitud')), empresa = clean(get(r, 'empresa'));
      if (!orden && !solicitud && !empresa) continue;
      const hasOrder = /^OD[SC]|^\w{2,}-/i.test(orden);
      let tipo = clean(get(r, 'tipo')).toUpperCase();
      if (!tipo && hasOrder) tipo = orden.slice(0, 3).toUpperCase();
      const coment = clean(get(r, 'coment'));
      const fSol = toDate(get(r, 'fSol')), fEm = toDate(get(r, 'fEm'));
      const x = {
        seg: segLabel, no: get(r, 'no'), orden, solicitud, empresa, tipo: tipo || 'Sin definir',
        prioridad: titleCase(clean(get(r, 'prioridad'))) || 'Sin definir', fSol, fEm,
        owner: clean(get(r, 'owner')) || 'Sin asignar', procura: clean(get(r, 'procura')) || 'Sin asignar',
        coment, hasOrder, estado: inferEstado(coment, hasOrder), region: region(orden, solicitud),
        dias: fEm ? daysSince(fEm) : (fSol ? daysSince(fSol) : null)
      };
      x._txt = norm([x.orden, x.solicitud, x.empresa, x.coment, x.owner, x.procura].join(' '));
      items.push(x);
    }
    return { items, ownerLabel: ix.owner >= 0 ? rawH[ix.owner] : 'Encargado' };
  }

  function parsePayments(ws) {
    const rows = sheetRows(ws);
    const hi = findHeaderRow(rows);
    const h = (rows[hi] || []).map(norm);
    const f = t => h.findIndex(t);
    const ix = {
      orden: f(c => c.includes('orden')), solicitud: f(c => c === 'solicitud' || c.includes('descripcion')),
      empresa: f(c => c.includes('empresa') || c.includes('proveedor')), monto: f(c => c.includes('monto')),
      pct: f(c => c.includes('%') || c.includes('anticipo'))
    };
    const get = (r, k) => ix[k] >= 0 ? r[ix[k]] : '';
    const items = [];
    for (const r of rows.slice(hi + 1)) {
      const orden = clean(get(r, 'orden')), solicitud = clean(get(r, 'solicitud')), empresa = clean(get(r, 'empresa'));
      if (!orden && !solicitud && !empresa) continue;
      const mRaw = get(r, 'monto');
      const monto = typeof mRaw === 'number' ? mRaw : (parseFloat(String(mRaw).replace(/[^\d,.-]/g, '').replace(/\.(?=\d{3})/g, '').replace(',', '.')) || 0);
      const pRaw = get(r, 'pct');
      let p = null, incidencia = '';
      if (typeof pRaw === 'number') p = pRaw > 1 ? pRaw / 100 : pRaw;
      else if (/^\s*\d+([.,]\d+)?\s*%?\s*$/.test(String(pRaw))) { p = parseFloat(String(pRaw).replace(',', '.')); if (p > 1) p /= 100; }
      else if (clean(pRaw)) incidencia = clean(pRaw);
      const x = {
        orden, solicitud, empresa: empresa || 'Sin empresa', monto, pct: p, incidencia,
        anticipo: p !== null ? monto * p : 0, saldo: p !== null ? monto * (1 - p) : monto,
        tipo: orden ? orden.slice(0, 3).toUpperCase() : 'Sin definir', region: region(orden, solicitud)
      };
      x._txt = norm([x.orden, x.solicitud, x.empresa, x.incidencia].join(' '));
      items.push(x);
    }
    return { items };
  }

  function parseMaterials(ws) {
    const rows = sheetRows(ws);
    const hi = findHeaderRow(rows);
    const headers = (rows[hi] || []).map(clean).filter(Boolean);
    const h = headers.map(norm);
    const f = t => h.findIndex(t);
    const ix = {
      material: f(c => c.includes('material')), min: f(c => c.includes('minima') || c.includes('minimo')),
      max: f(c => c.includes('maxima') || c.includes('maximo')),
      stock: f(c => c.includes('existencia') || c.includes('stock') || c.includes('actual') || c.includes('inventario'))
    };
    const num = v => typeof v === 'number' ? v : parseFloat(String(v).replace(',', '.'));
    const items = [];
    for (const r of rows.slice(hi + 1)) {
      const vals = headers.map((_, i) => r[i]);
      if (!vals.some((v, i) => clean(v) && !/^no\.?$/.test(h[i]))) continue;
      const it = { cells: vals, name: ix.material >= 0 ? clean(vals[ix.material]) : clean(vals[1]) };
      if (ix.stock >= 0 && ix.min >= 0) { const s = num(vals[ix.stock]), mn = num(vals[ix.min]); it.bajoMinimo = !isNaN(s) && !isNaN(mn) && s < mn; }
      items.push(it);
    }
    return { items, headers, ix };
  }

  function processWorkbook(wb, meta) {
    const names = pickSheets(wb);
    return {
      meta, sheetNames: names,
      mtto: names.mtto ? parseOrders(wb.Sheets[names.mtto], 'Mantenimiento') : { items: [], ownerLabel: 'Encargado' },
      oper: names.oper ? parseOrders(wb.Sheets[names.oper], 'Operaciones') : { items: [], ownerLabel: 'Encargado' },
      mat: names.mat ? parseMaterials(wb.Sheets[names.mat]) : { items: [], headers: [], ix: {} },
      rq: names.rq ? parsePayments(wb.Sheets[names.rq]) : { items: [] },
      modified: (wb.Props && (wb.Props.ModifiedDate || wb.Props.CreatedDate)) || null
    };
  }

  /* ---------------- Carga de datos ---------------- */
  async function loadDefault() {
    const url = encodeURI(CFG.dataFile);
    $('#btnDownload').setAttribute('href', url);
    $('#btnDownload').setAttribute('download', CFG.dataFile.split('/').pop());
    try {
      const res = await fetch(url + '?v=' + Date.now(), { cache: 'no-store' });
      if (!res.ok) throw new Error('HTTP ' + res.status);
      const wb = XLSX.read(await res.arrayBuffer(), { type: 'array', cellDates: false });
      render(processWorkbook(wb, { source: 'repo', fileName: CFG.dataFile.split('/').pop(), lastModified: res.headers.get('last-modified') }));
    } catch (err) {
      console.warn('No se pudo leer el Excel publicado:', err);
      showNoData();
    }
  }

  function loadLocalFile(file) {
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array', cellDates: false });
        if (downloadUrl) URL.revokeObjectURL(downloadUrl);
        downloadUrl = URL.createObjectURL(file);
        $('#btnDownload').setAttribute('href', downloadUrl);
        $('#btnDownload').setAttribute('download', file.name);
        render(processWorkbook(wb, { source: 'local', fileName: file.name, lastModified: new Date(file.lastModified).toUTCString() }));
        const b = $('#localBanner');
        b.hidden = false;
        b.innerHTML = `<span>Estás viendo <b>${esc(file.name)}</b> cargado desde tu equipo. Este cambio es solo en tu navegador: para actualizar la página publicada, sube el archivo a la carpeta <code>data/</code> del repositorio.</span>
          <button class="btn btn--light" id="btnBackRepo">Volver a los datos publicados</button>`;
        $('#btnBackRepo').onclick = () => { b.hidden = true; loadDefault(); };
      } catch (err) {
        alert('No se pudo leer el archivo. Verifica que sea un Excel (.xlsx) con la misma estructura.\n\n' + err.message);
      }
    };
    reader.readAsArrayBuffer(file);
  }

  function showNoData() {
    $('#dataMeta').textContent = 'No se encontró el archivo de datos publicado.';
    const html = `<div class="empty" style="margin-top:24px">
        <h3>No se pudieron cargar los datos</h3>
        <p>Si abriste el archivo <code>index.html</code> directamente desde tu computadora, el navegador bloquea la lectura automática del Excel.<br>
        Usa el botón <b>Cargar Excel</b> para visualizarlo, o publícalo en GitHub Pages (ver README).</p>
        <p style="margin-top:16px"><button type="button" class="btn btn--primary" data-upload>Cargar Excel</button></p>
      </div>`;
    VIEWS.forEach(k => { $('#view-' + k).innerHTML = html; });
  }

  /* ---------------- Contraseña de carga ---------------- */
  let unlocked = false;
  try { unlocked = sessionStorage.getItem('colven-upload-ok') === '1'; } catch (e) { /* almacenamiento no disponible */ }
  function requestUpload() {
    if (unlocked) { $('#fileInput').click(); return; }
    const m = $('#pwModal');
    m.hidden = false;
    $('#pwErr').hidden = true;
    $('#pwInput').value = '';
    setTimeout(() => $('#pwInput').focus(), 30);
  }
  const closeModal = () => { $('#pwModal').hidden = true; };
  $('#pwForm').addEventListener('submit', e => {
    e.preventDefault();
    if (sha256Hex($('#pwInput').value) === CFG.uploadPasswordHash) {
      unlocked = true;
      try { sessionStorage.setItem('colven-upload-ok', '1'); } catch (err) { /* sin almacenamiento */ }
      closeModal();
      $('#fileInput').click();
    } else {
      $('#pwErr').hidden = false;
      const card = $('#pwModal .modal__card');
      card.classList.remove('shake'); void card.offsetWidth; card.classList.add('shake');
      $('#pwInput').select();
    }
  });
  $('#pwCancel').addEventListener('click', closeModal);
  $('#pwModal').addEventListener('click', e => { if (e.target.id === 'pwModal') closeModal(); });
  document.addEventListener('keydown', e => { if (e.key === 'Escape' && !$('#pwModal').hidden) closeModal(); });

  /* ---------------- Gráficos ---------------- */
  Chart.defaults.font.family = "'Inter', system-ui, sans-serif";
  Chart.defaults.font.size = 12;
  Chart.defaults.color = '#646985';
  Chart.defaults.animation.duration = 350;
  Chart.defaults.plugins.legend.labels.usePointStyle = true;
  Chart.defaults.plugins.legend.labels.boxWidth = 8;
  Chart.defaults.plugins.tooltip.backgroundColor = '#1D2041';
  Chart.defaults.plugins.tooltip.padding = 10;
  Chart.defaults.plugins.tooltip.cornerRadius = 8;

  function makeChart(id, config) {
    const el = document.getElementById(id);
    if (!el) return null;
    if (charts[id]) charts[id].destroy();
    charts[id] = new Chart(el, config);
    return charts[id];
  }
  function destroyCharts(prefix) {
    for (const id of Object.keys(charts)) if (id.startsWith(prefix)) { charts[id].destroy(); delete charts[id]; }
  }
  const pointer = (e, els) => { if (e.native && e.native.target) e.native.target.style.cursor = els.length ? 'pointer' : 'default'; };

  // Constructores de configuración reutilizables
  function hbar(labels, data, colors, o = {}) {
    return {
      type: 'bar',
      data: { labels: labels.map(l => short(l, o.cut || 30)), datasets: [{ data, backgroundColor: colors, borderRadius: 6, borderSkipped: false, maxBarThickness: 26 }] },
      options: {
        indexAxis: 'y', maintainAspectRatio: false,
        plugins: { legend: { display: false }, tooltip: { callbacks: { title: i => labels[i[0].dataIndex], label: c => o.money ? fmtMoney(c.raw) : (o.suffix ? c.raw + o.suffix : String(c.raw)) } } },
        scales: { x: { grid: { color: C.line }, ticks: o.money ? { callback: v => fmtMoneyShort(v) } : { precision: 0 }, max: o.max }, y: { grid: { display: false }, ticks: { autoSkip: false } } }
      }
    };
  }
  function donut(labels, data, colors, o = {}) {
    const tot = data.reduce((a, b) => a + b, 0);
    return {
      type: 'doughnut',
      data: { labels, datasets: [{ data, backgroundColor: colors, borderColor: '#fff', borderWidth: 2 }] },
      options: { maintainAspectRatio: false, cutout: '60%', plugins: { legend: { position: 'right', labels: { font: { size: 11 } } }, tooltip: { callbacks: { label: c => `${c.label}: ${o.money ? fmtMoney(c.raw) : c.raw} (${pct(c.raw, tot)}%)` } } } }
    };
  }
  function estadoCfg(items) {
    const est = ESTADOS.filter(e => items.some(x => x.estado === e.key));
    return hbar(est.map(e => e.key), est.map(e => items.filter(x => x.estado === e.key).length), est.map(e => e.color));
  }
  function countCfg(items, fn, color = C.navy, top = 8) {
    const m = countBy(items, fn);
    const ks = Object.keys(m).sort((a, b) => m[b] - m[a]).slice(0, top);
    return hbar(ks, ks.map(k => m[k]), color);
  }
  function regionCfg(items) {
    const m = countBy(items, x => x.region);
    const ks = Object.keys(REGION_COLORS).filter(k => m[k]);
    return donut(ks, ks.map(k => m[k]), ks.map(k => REGION_COLORS[k]));
  }
  function prioCfg(items) {
    const m = countBy(items, x => x.prioridad);
    const ks = Object.keys(PRIO_COLORS).filter(k => m[k]);
    return donut(ks, ks.map(k => m[k]), ks.map(k => PRIO_COLORS[k]));
  }
  function segCfg(items) {
    const m = countBy(items, x => x.seg);
    const ks = Object.keys(m);
    return donut(ks, ks.map(k => m[k]), ks.map((_, i) => [C.navy, C.yellow][i % 2]));
  }
  function payStackCfg(rows) {
    const cfg = {
      type: 'bar',
      data: {
        labels: rows.map(x => x.orden || short(x.solicitud, 20)),
        datasets: [
          { label: 'Anticipo', data: rows.map(x => x.anticipo), backgroundColor: C.yellow, borderRadius: 3, maxBarThickness: 26 },
          { label: 'Saldo', data: rows.map(x => x.saldo), backgroundColor: rows.map(x => x.incidencia ? '#F1B3B1' : C.navy), borderRadius: 3, maxBarThickness: 26 }
        ]
      },
      options: {
        indexAxis: 'y', maintainAspectRatio: false,
        scales: { x: { stacked: true, grid: { color: C.line }, ticks: { callback: v => fmtMoneyShort(v) } }, y: { stacked: true, grid: { display: false }, ticks: { autoSkip: false, font: { family: 'Consolas, monospace', size: 11 } } } },
        plugins: {
          legend: { position: 'bottom' },
          tooltip: { callbacks: { label: c => `${c.dataset.label}: ${fmtMoney(c.raw)}`, afterBody: i => { const x = i[0].chart.$rows[i[0].dataIndex]; return x ? [x.empresa, x.incidencia ? '⚠ ' + x.incidencia : `Anticipo ${Math.round((x.pct || 0) * 100)}%`] : []; } } }
        }
      }
    };
    return cfg;
  }
  function provDonutCfg(rows, val = x => x.monto) {
    const m = {};
    rows.forEach(x => { m[x.empresa] = (m[x.empresa] || 0) + val(x); });
    const ks = Object.keys(m).filter(k => m[k] > 0).sort((a, b) => m[b] - m[a]);
    const cfg = donut(ks.map(k => short(k, 28)), ks.map(k => m[k]), ks.map((_, i) => PAL[i % PAL.length]), { money: true });
    cfg.options.plugins.legend.position = 'bottom';
    return cfg;
  }
  const cfgHeight = cfg => {
    if (cfg.type === 'bar' && cfg.options.indexAxis === 'y') return Math.min(440, Math.max(170, cfg.data.labels.length * 32 + (cfg.options.plugins.legend.display === false ? 40 : 70)));
    return 240;
  };

  /* ---------------- Tablas ---------------- */
  const prioBadge = p => `<span class="badge prio-${{ Alta: 'alta', Media: 'media', Baja: 'baja' }[p] || 'none'}">${esc(p)}</span>`;
  const estadoBadge = e => { const col = estadoColor(e); const dark = ['#FDD649', '#C9CCDD', '#E6E8F0', '#9AA0BF'].includes(col); return `<span class="badge" style="background:${col};color:${dark ? '#1D2041' : '#fff'}">${esc(e)}</span>`; };
  const dash = '<span class="muted">—</span>';

  function orderCols(o = {}) {
    return [
      { t: 'No. de orden', f: x => x.orden ? `<span class="mono">${esc(x.orden)}</span>` : '<span class="muted">Sin orden</span>', cls: 'nowrap' },
      o.seg && { t: 'Segmento', f: x => esc(x.seg), cls: 'nowrap' },
      { t: 'Solicitud', f: x => esc(x.solicitud), cls: 'wide' },
      { t: 'Empresa', f: x => esc(x.empresa) || dash, cls: 'mid' },
      { t: 'Tipo', f: x => esc(x.tipo), cls: 'nowrap' },
      { t: 'Prioridad', f: x => prioBadge(x.prioridad) },
      { t: 'Estado (estimado)', f: x => estadoBadge(x.estado) },
      { t: 'Emisión', f: x => fmtDate(x.fEm) || (x.fSol ? `<span class="muted">Sol. ${fmtDate(x.fSol)}</span>` : dash), cls: 'nowrap' },
      { t: 'Días', f: x => x.dias ?? dash, cls: 'num' },
      { t: 'Enc. Procura', f: x => esc(x.procura), cls: 'nowrap' },
      { t: 'Comentarios', f: x => esc(x.coment) || dash, cls: 'wide' }
    ].filter(Boolean);
  }
  const payCols = () => [
    { t: 'No. de orden', f: x => `<span class="mono">${esc(x.orden)}</span>`, cls: 'nowrap' },
    { t: 'Solicitud', f: x => esc(x.solicitud), cls: 'wide' },
    { t: 'Empresa', f: x => esc(x.empresa), cls: 'mid' },
    { t: 'Monto', f: x => fmtMoney(x.monto), cls: 'num' },
    { t: '% Anticipo', f: x => x.pct !== null ? Math.round(x.pct * 100) + '%' : `<span class="badge prio-alta">${esc(x.incidencia)}</span>`, cls: 'num' },
    { t: 'Anticipo', f: x => x.pct !== null ? fmtMoney(x.anticipo) : dash, cls: 'num' },
    { t: 'Saldo', f: x => fmtMoney(x.saldo), cls: 'num' }
  ];
  const tableHTML = (cols, rows) => `<table><thead><tr>${cols.map(c => `<th style="cursor:default">${esc(c.t)}</th>`).join('')}</tr></thead>
    <tbody>${rows.map(x => `<tr>${cols.map(c => `<td class="${c.cls || ''}">${c.f(x)}</td>`).join('')}</tr>`).join('')}</tbody></table>`;

  /* ---------------- Paneles de detalle (drill-down) ---------------- */
  function ordersDrill(title, sub, items, o = {}) {
    const spec = { title, sub, count: items.length, action: o.action };
    if (!items.length) { spec.empty = o.emptyMsg || 'No hay registros en esta categoría.'; return spec; }
    spec.stats = o.stats || [
      { l: 'Registros', v: fmtInt(items.length) },
      { l: 'Abiertas', v: fmtInt(items.filter(x => ABIERTO(x.estado)).length) },
      { l: 'Con orden', v: fmtInt(items.filter(x => x.hasOrder).length) },
      { l: 'Prioridad alta', v: fmtInt(items.filter(x => x.prioridad === 'Alta').length) }
    ];
    spec.charts = o.charts || [{ t: 'Por estado', cfg: estadoCfg(items) }, { t: 'Por responsable de Procura', cfg: countCfg(items, x => x.procura) }];
    spec.table = { cols: orderCols({ seg: o.showSeg }), rows: o.rows || items };
    return spec;
  }

  function paymentsDrill(title, sub, items, mode, action) {
    const spec = { title, sub, count: items.length, action };
    if (!items.length) { spec.empty = 'No hay registros en esta categoría.'; return spec; }
    const tot = sumBy(items, x => x.monto), ant = sumBy(items, x => x.anticipo);
    spec.stats = [
      { l: 'Monto', v: fmtMoney(tot) }, { l: 'Anticipos', v: `${fmtMoney(ant)} · ${pct(ant, tot)}%` },
      { l: 'Saldo', v: fmtMoney(tot - ant) }, { l: 'Órdenes', v: fmtInt(items.length) }
    ];
    const sorted = items.slice().sort((a, b) => b.monto - a.monto);
    if (mode === 'monto') {
      const st = payStackCfg(sorted);
      spec.charts = [{ t: 'Anticipo vs. saldo por orden', cfg: st, rows: sorted }, { t: 'Monto por proveedor', cfg: provDonutCfg(items), h: 300 }];
    } else if (mode === 'anticipo') {
      const s = items.filter(x => x.pct !== null).sort((a, b) => b.anticipo - a.anticipo);
      spec.charts = [
        { t: 'Anticipo por orden', cfg: hbar(s.map(x => x.orden), s.map(x => x.anticipo), C.gold, { money: true }) },
        { t: '% de anticipo acordado', cfg: hbar(s.map(x => x.orden), s.map(x => Math.round(x.pct * 100)), C.navy3, { suffix: '%', max: 100 }) }
      ];
    } else if (mode === 'saldo') {
      const s = items.slice().sort((a, b) => b.saldo - a.saldo);
      spec.charts = [
        { t: 'Saldo por pagar por orden', cfg: hbar(s.map(x => x.orden), s.map(x => x.saldo), s.map(x => x.incidencia ? '#F1B3B1' : C.navy), { money: true }) },
        { t: 'Saldo por proveedor', cfg: provDonutCfg(items, x => x.saldo), h: 300 }
      ];
    }
    spec.table = { cols: payCols(), rows: sorted };
    return spec;
  }

  function materialsDrill(title, sub, items, action) {
    const seg = DATA.mat;
    const spec = { title, sub, count: items.length, action };
    if (!items.length) {
      spec.empty = seg.items.length ? 'No hay materiales en esta categoría.' : `La hoja <b>${esc(DATA.sheetNames.mat || 'Materiales Criticos')}</b> aún no tiene registros. Columnas disponibles: ${seg.headers.map(esc).join(' · ')}.`;
      return spec;
    }
    spec.stats = [{ l: 'Materiales', v: fmtInt(items.length) }, { l: 'Bajo mínimo', v: seg.ix.stock >= 0 ? fmtInt(items.filter(x => x.bajoMinimo).length) : '—' }];
    spec.table = { cols: seg.headers.map((hd, i) => ({ t: hd, f: x => esc(x.cells[i]), cls: typeof items[0].cells[i] === 'number' ? 'num' : '' })), rows: items };
    return spec;
  }

  function drillHTML(slot, spec) {
    const ch = spec.charts || [];
    return `<div class="drill__panel">
      <div class="drill__head">
        <div><p class="drill__eyebrow">Detalle</p><h3>${esc(spec.title)}<span class="drill__count">${fmtInt(spec.count)}</span></h3>${spec.sub ? `<p class="sub">${esc(spec.sub)}</p>` : ''}</div>
        <div class="drill__actions">${spec.action ? `<button type="button" class="btn btn--light" data-drill-go="${slot}">${esc(spec.action.label)} →</button>` : ''}
          <button type="button" class="drill__close" data-drill-close="${slot}" aria-label="Cerrar detalle">×</button></div>
      </div>
      ${spec.stats ? `<div class="drill__stats">${spec.stats.map(s => `<div><span>${esc(s.l)}</span><b>${s.v}</b></div>`).join('')}</div>` : ''}
      ${spec.empty ? `<div class="empty"><p>${spec.empty}</p></div>` : ''}
      ${ch.length ? `<div class="drill__charts">${ch.map((c, i) => `<div class="drill__chart"><h4>${esc(c.t)}</h4><div class="chart-box" style="height:${c.h || cfgHeight(c.cfg)}px"><canvas id="dc-${slot}-${i}"></canvas></div></div>`).join('')}</div>` : ''}
      ${spec.table && spec.table.rows.length ? `<div class="table-wrap drill__table">${tableHTML(spec.table.cols, spec.table.rows)}</div>` : ''}
    </div>`;
  }

  function closeDrill(slot) {
    const box = document.getElementById('drill-' + slot);
    destroyCharts(`dc-${slot}-`);
    if (box) { box.classList.remove('is-open'); box.innerHTML = ''; }
    const grid = document.querySelector(`[data-slot="${slot}"]`);
    if (grid) grid.querySelectorAll('.is-active').forEach(el => el.classList.remove('is-active'));
    openDrill[slot] = null; drillSpecs[slot] = null;
  }

  function toggleDrill(slot, id, tile) {
    const view = tile.closest('.view');
    // cerrar cualquier otro panel abierto en la misma vista
    for (const s of Object.keys(openDrill)) {
      if (s !== slot && openDrill[s]) { const b = document.getElementById('drill-' + s); if (b && b.closest('.view') === view) closeDrill(s); }
    }
    if (openDrill[slot] === id) { closeDrill(slot); return; }
    const make = drills[id];
    if (!make) return;
    const spec = make();
    destroyCharts(`dc-${slot}-`);
    const box = document.getElementById('drill-' + slot);
    box.innerHTML = drillHTML(slot, spec);
    box.classList.add('is-open');
    view.querySelectorAll('[data-slot] .is-active').forEach(el => el.classList.remove('is-active'));
    tile.classList.add('is-active');
    openDrill[slot] = id; drillSpecs[slot] = spec;
    setTimeout(() => {
      if (openDrill[slot] !== id) return;   // se cerró o cambió antes de dibujar
      (spec.charts || []).forEach((c, i) => { const chart = makeChart(`dc-${slot}-${i}`, c.cfg); if (chart && c.rows) chart.$rows = c.rows; });
      const r = box.getBoundingClientRect();
      if (r.top > window.innerHeight * 0.7) window.scrollBy({ top: r.top - 140, behavior: 'smooth' });
    }, 0);
  }

  function runAction(a) {
    if (!a) return;
    if (a.filter) tableState[a.seg] = Object.assign(defaultState(a.seg), a.filter);
    const go = () => { if (a.filter || a.openDetail) { refreshTable(a.seg); openDetail(a.seg); } };
    if (curView !== a.seg) { location.hash = a.seg; setTimeout(go, 80); } else go();
  }

  /* ---------------- Render principal ---------------- */
  function render(d) {
    DATA = d;
    destroyCharts('');
    builders = {}; built = {}; drills = {}; openDrill = {}; drillSpecs = {};
    for (const k of Object.keys(tableState)) delete tableState[k];

    const mod = d.meta.lastModified ? new Date(d.meta.lastModified) : (d.modified ? new Date(d.modified) : null);
    const modTxt = mod && !isNaN(mod) ? mod.toLocaleString('es-VE', { day: '2-digit', month: 'long', year: 'numeric', hour: '2-digit', minute: '2-digit' }) : '—';
    $('#dataMeta').textContent = `Datos actualizados: ${modTxt} · Corte al ${new Date().toLocaleDateString('es-VE', { day: '2-digit', month: 'long', year: 'numeric' })}`;
    $('#footerSource').textContent = `Fuente: ${d.meta.fileName}`;

    registerOrderDrills('mtto', d.mtto);
    registerOrderDrills('oper', d.oper);
    registerOtherDrills(d);

    renderResumen(d);
    renderOrders('mtto', 'Mantenimiento', d.mtto, 'Órdenes de servicio (ODS) y de compra (ODC) asociadas al mantenimiento de equipos.');
    renderOrders('oper', 'Operaciones', d.oper, 'Órdenes de servicio y compra solicitadas por el área de Operaciones.');
    renderMaterials(d.mat);
    renderPayments(d.rq);
    curView = null;
    route();
  }

  /* ---------------- Componentes ---------------- */
  const kpi = (label, value, hint, cls, drill) =>
    `<button type="button" class="card kpi ${cls || ''}" data-drill="${drill}" aria-label="${esc(label)}: ver detalle">
       <div class="kpi__label">${label}</div><div class="kpi__value">${value}</div><div class="kpi__hint">${hint || '&nbsp;'}</div>
       <span class="kpi__more">Ver detalle ${CHEV}</span></button>`;

  const ICONS = {
    mtto: '<path d="M14.7 6.3a4 4 0 0 0-5.4 5.4L3 18l3 3 6.3-6.3a4 4 0 0 0 5.4-5.4l-2.6 2.6-2.4-.6-.6-2.4z"/>',
    oper: '<path d="M3 21h18M5 21V9l7-5 7 5v12M9 21v-6h6v6"/>',
    mat: '<path d="M21 8 12 3 3 8v8l9 5 9-5zM3 8l9 5 9-5M12 13v8"/>',
    rq: '<path d="M12 2v20M17 6H9.5a3.5 3.5 0 0 0 0 7h5a3.5 3.5 0 0 1 0 7H6"/>'
  };

  function detailBlock(key, count, inner) {
    return `<div class="detail" id="detail-${key}">
      <button type="button" class="detail__toggle" data-detail="${key}">
        <span>Ver detalle completo <span class="muted">(${count} registros)</span></span>
        <svg class="chev" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="m6 9 6 6 6-6"/></svg>
      </button>
      <div class="detail__body">${inner}</div>
    </div>`;
  }
  function openDetail(key) {
    const el = $('#detail-' + key);
    if (el) { el.classList.add('is-open'); setTimeout(() => el.scrollIntoView({ behavior: 'smooth', block: 'start' }), 60); }
  }

  /* ---------------- Registro de detalles por cuadro ---------------- */
  function registerOrderDrills(key, seg) {
    const it = seg.items, L = key === 'mtto' ? 'Mantenimiento' : 'Operaciones';
    drills[`${key}-total`] = () => ordersDrill(`Solicitudes · ${L}`, 'Todas las solicitudes registradas en la hoja', it);
    drills[`${key}-orden`] = () => {
      const r = it.filter(x => x.hasOrder);
      return ordersDrill(`Con orden emitida · ${L}`, 'ODS = orden de servicio · ODC = orden de compra', r, {
        charts: [{ t: 'Por tipo de orden', cfg: countCfg(r, x => x.tipo, [C.navy, C.yellow, C.navy3]) }, { t: 'Por región', cfg: regionCfg(r) }],
        action: { label: 'Filtrar tabla', seg: key, filter: { orden: 'con' } }
      });
    };
    drills[`${key}-sinorden`] = () => {
      const r = it.filter(x => !x.hasOrder);
      return ordersDrill(`En gestión, sin orden · ${L}`, 'Solicitudes que aún no tienen número de orden (cotización, validación o elaboración)', r, {
        action: { label: 'Filtrar tabla', seg: key, filter: { orden: 'sin' } }, emptyMsg: 'Todas las solicitudes tienen orden emitida.'
      });
    };
    drills[`${key}-alta`] = () => {
      const r = it.filter(x => x.prioridad === 'Alta');
      return ordersDrill(`Prioridad alta · ${L}`, 'Solicitudes marcadas con prioridad Alta', r, { action: { label: 'Filtrar tabla', seg: key, filter: { prioridad: 'Alta' } } });
    };
    drills[`${key}-aprob`] = () => {
      const r = it.filter(x => x.estado === 'Por aprobación');
      return ordersDrill(`Por aprobación · ${L}`, 'Órdenes en espera de firma o aprobación', r, {
        charts: [{ t: 'Por responsable de Procura', cfg: countCfg(r, x => x.procura) }, { t: 'Por tipo de orden', cfg: countCfg(r, x => x.tipo, C.navy3) }],
        action: { label: 'Filtrar tabla', seg: key, filter: { estado: 'Por aprobación' } }
      });
    };
    drills[`${key}-entrega`] = () => {
      const r = it.filter(x => x.estado === 'Entrega pendiente');
      return ordersDrill(`Entrega pendiente · ${L}`, 'Órdenes en fabricación o con entrega pendiente del proveedor', r, {
        charts: [{ t: 'Por proveedor', cfg: countCfg(r, x => x.empresa || 'Sin empresa') }, { t: 'Por región', cfg: regionCfg(r) }],
        action: { label: 'Filtrar tabla', seg: key, filter: { estado: 'Entrega pendiente' } }
      });
    };
    drills[`${key}-edad`] = () => {
      const r = it.filter(x => x.hasOrder && x.fEm && ABIERTO(x.estado)).sort((a, b) => b.dias - a.dias);
      const U = CFG.agingAlertDays;
      const buckets = [['0–15 días', 0, 15], ['16–30 días', 16, 30], [`31–${U} días`, 31, U], [`Más de ${U} días`, U + 1, Infinity]];
      const spec = ordersDrill(`Antigüedad de órdenes abiertas · ${L}`, 'Días transcurridos desde la fecha de emisión de la orden', r, {
        stats: r.length ? [
          { l: 'Órdenes', v: fmtInt(r.length) },
          { l: 'Promedio', v: Math.round(sumBy(r, x => x.dias) / r.length) + ' días' },
          { l: 'Máximo', v: r[0].dias + ' días' },
          { l: `> ${U} días`, v: fmtInt(r.filter(x => x.dias > U).length) }
        ] : null,
        charts: [
          { t: 'Días desde la emisión', cfg: hbar(r.map(x => x.orden), r.map(x => x.dias), r.map(x => x.dias > U ? C.risk : C.navy), { suffix: ' días' }) },
          { t: 'Rangos de antigüedad', cfg: hbar(buckets.map(b => b[0]), buckets.map(b => r.filter(x => x.dias >= b[1] && x.dias <= b[2]).length), [C.ok, C.navy3, C.gold, C.risk]) }
        ],
        action: { label: `Ver > ${U} días en tabla`, seg: key, filter: { aging: true } },
        emptyMsg: 'No hay órdenes abiertas con fecha de emisión registrada.'
      });
      return spec;
    };
  }

  function registerOtherDrills(d) {
    const ord = [...d.mtto.items, ...d.oper.items];
    const rq = d.rq.items;
    // Resumen general
    drills['res-total'] = () => ordersDrill('Solicitudes registradas', 'Mantenimiento + Operaciones', ord, {
      showSeg: true,
      charts: [{ t: 'Por estado', cfg: estadoCfg(ord) }, { t: 'Por segmento', cfg: segCfg(ord) }]
    });
    drills['res-emitidas'] = () => {
      const r = ord.filter(x => x.hasOrder);
      return ordersDrill('Órdenes emitidas', `${fmtInt(r.length)} con número de orden · ${fmtInt(ord.length - r.length)} aún sin orden`, r, {
        showSeg: true,
        stats: [
          { l: 'Con orden', v: `${fmtInt(r.length)} · ${pct(r.length, ord.length)}%` },
          { l: 'ODS', v: fmtInt(r.filter(x => x.tipo === 'ODS').length) },
          { l: 'ODC', v: fmtInt(r.filter(x => x.tipo === 'ODC').length) },
          { l: 'Sin orden', v: fmtInt(ord.length - r.length) }
        ],
        charts: [{ t: 'Por segmento', cfg: segCfg(r) }, { t: 'Por región', cfg: regionCfg(r) }]
      });
    };
    drills['res-alta'] = () => {
      const r = ord.filter(x => x.prioridad === 'Alta' && ABIERTO(x.estado));
      return ordersDrill('Prioridad alta abiertas', 'Solicitudes de prioridad Alta que no están completadas ni canalizadas', r, { showSeg: true });
    };
    drills['res-monto'] = () => paymentsDrill('Monto comprometido', 'Órdenes con requerimiento de anticipo', rq, 'monto', { label: 'Abrir sección', seg: 'rq' });

    // Tarjetas de segmento
    drills['seg-mtto'] = () => ordersDrill('Mantenimiento', 'Resumen del segmento', d.mtto.items, { action: { label: 'Abrir sección', seg: 'mtto' } });
    drills['seg-oper'] = () => ordersDrill('Operaciones', 'Resumen del segmento', d.oper.items, { action: { label: 'Abrir sección', seg: 'oper' } });
    drills['seg-mat'] = () => materialsDrill('Materiales críticos', 'Inventario de materiales críticos', d.mat.items, { label: 'Abrir sección', seg: 'mat' });
    drills['seg-rq'] = () => paymentsDrill('Anticipos y montos', 'Resumen del segmento', rq, 'monto', { label: 'Abrir sección', seg: 'rq' });

    // Anticipos
    drills['rq-total'] = () => paymentsDrill('Monto total', 'Monto de cada orden y su distribución por proveedor', rq, 'monto');
    drills['rq-ant'] = () => paymentsDrill('Anticipos', 'Anticipo calculado como Monto × % de anticipo', rq.filter(x => x.pct !== null), 'anticipo');
    drills['rq-saldo'] = () => paymentsDrill('Saldo por pagar', 'Monto pendiente después del anticipo (contra entrega o valuación)', rq, 'saldo');
    drills['rq-incid'] = () => {
      const r = rq.filter(x => x.incidencia);
      const spec = paymentsDrill('Anticipos con incidencia', 'Registros cuyo % de anticipo indica una novedad en lugar de un porcentaje', r, 'none');
      if (r.length) spec.stats = [{ l: 'Registros', v: fmtInt(r.length) }, { l: 'Monto afectado', v: fmtMoney(sumBy(r, x => x.monto)) }, ...r.slice(0, 3).map(x => ({ l: x.orden, v: esc(x.incidencia) }))];
      return spec;
    };

    // Materiales
    drills['mat-total'] = () => materialsDrill('Materiales registrados', 'Listado completo de la hoja de materiales críticos', d.mat.items);
    drills['mat-bajo'] = () => {
      const s = materialsDrill('Materiales bajo nivel mínimo', 'Existencia actual por debajo de la cantidad mínima', d.mat.items.filter(x => x.bajoMinimo));
      if (d.mat.ix.stock < 0) s.empty = 'Para calcular este indicador agrega una columna <b>“Existencia actual”</b> en la hoja de materiales críticos.';
      return s;
    };
    drills['mat-campos'] = () => ({
      title: 'Campos de control', sub: 'Columnas de la hoja de materiales críticos', count: d.mat.headers.length,
      stats: d.mat.headers.map((hd, i) => ({ l: 'Columna ' + (i + 1), v: esc(hd) })), action: { label: 'Abrir sección', seg: 'mat' }
    });
  }

  /* ---------------- Resumen general ---------------- */
  function renderResumen(d) {
    const ord = [...d.mtto.items, ...d.oper.items];
    const m = d.mtto.items, o = d.oper.items;
    const conOrden = ord.filter(x => x.hasOrder).length;
    const alta = ord.filter(x => x.prioridad === 'Alta' && ABIERTO(x.estado)).length;
    const montoTot = sumBy(d.rq.items, x => x.monto), antTot = sumBy(d.rq.items, x => x.anticipo);
    const abiertas = ord.filter(x => ABIERTO(x.estado)).length;

    const stat = (id, v, l) => `<button type="button" class="seg__stat" data-drill="${id}"><b>${v}</b>${l}</button>`;
    const segCard = (key, tag, big, bigLabel, s1, s2, progress) => `
      <div class="card seg" data-drill="seg-${key}" role="button" tabindex="0" aria-label="${tag}: ver detalle">
        <div class="seg__top"><span class="seg__tag">${tag}</span><span class="seg__icon"><svg viewBox="0 0 24 24">${ICONS[key]}</svg></span></div>
        <div class="seg__big">${big}<small>${bigLabel}</small></div>
        ${progress !== undefined ? `<div class="bar" title="${progress}% con orden emitida"><span style="width:${progress}%"></span></div>` : ''}
        <div class="seg__stats">${s1}${s2}</div>
        <span class="seg__link">Ver detalle ${CHEV}</span>
      </div>`;

    const segs = [
      segCard('mtto', 'Mantenimiento', fmtInt(m.length), 'solicitudes',
        stat('mtto-alta', fmtInt(m.filter(x => x.prioridad === 'Alta').length), 'prioridad alta'),
        stat('mtto-sinorden', fmtInt(m.filter(x => !x.hasOrder).length), 'sin orden emitida'), pct(m.filter(x => x.hasOrder).length, m.length)),
      segCard('oper', 'Operaciones', fmtInt(o.length), 'solicitudes',
        stat('oper-aprob', fmtInt(o.filter(x => x.estado === 'Por aprobación').length), 'por aprobación'),
        stat('oper-entrega', fmtInt(o.filter(x => x.estado === 'Entrega pendiente').length), 'entrega pendiente'), pct(o.filter(x => x.hasOrder).length, o.length)),
      segCard('mat', 'Materiales críticos', fmtInt(d.mat.items.length), 'materiales',
        stat('mat-bajo', d.mat.ix && d.mat.ix.stock >= 0 ? fmtInt(d.mat.items.filter(x => x.bajoMinimo).length) : '—', 'bajo mínimo'),
        stat('mat-campos', fmtInt(d.mat.headers.length), 'campos de control')),
      segCard('rq', 'Anticipos y montos', fmtMoneyShort(montoTot), 'comprometido',
        stat('rq-ant', fmtMoneyShort(antTot), 'anticipos'),
        stat('rq-incid', fmtInt(d.rq.items.filter(x => x.incidencia).length), 'con incidencia'))
    ].join('');

    // Puntos de atención
    const alerts = [];
    const altaSin = ord.filter(x => x.prioridad === 'Alta' && !x.hasOrder);
    if (altaSin.length) alerts.push({ c: C.risk, n: altaSin.length, t: plural(altaSin.length, 'solicitud de prioridad alta aún sin orden emitida', 'solicitudes de prioridad alta aún sin orden emitida'), s: 'Mantenimiento / Operaciones', go: { seg: m.some(x => x.prioridad === 'Alta' && !x.hasOrder) ? 'mtto' : 'oper', filter: { orden: 'sin', prioridad: 'Alta' } } });
    const porAprob = ord.filter(x => x.estado === 'Por aprobación');
    if (porAprob.length) alerts.push({ c: C.yellow, n: porAprob.length, t: plural(porAprob.length, 'orden en espera de aprobación', 'órdenes en espera de aprobación'), s: `${m.filter(x => x.estado === 'Por aprobación').length} en Mantenimiento · ${o.filter(x => x.estado === 'Por aprobación').length} en Operaciones`, go: { seg: m.some(x => x.estado === 'Por aprobación') ? 'mtto' : 'oper', filter: { estado: 'Por aprobación' } } });
    const pagoPend = ord.filter(x => x.estado === 'Pago / anticipo pendiente');
    if (pagoPend.length) alerts.push({ c: C.gold, n: pagoPend.length, t: plural(pagoPend.length, 'orden con pago o anticipo pendiente', 'órdenes con pago o anticipo pendiente'), s: 'Requieren gestión de Finanzas', go: { seg: m.some(x => x.estado === 'Pago / anticipo pendiente') ? 'mtto' : 'oper', filter: { estado: 'Pago / anticipo pendiente' } } });
    const viejas = ord.filter(x => x.fEm && x.dias > CFG.agingAlertDays && ABIERTO(x.estado));
    if (viejas.length) alerts.push({ c: C.navy3, n: viejas.length, t: `${plural(viejas.length, 'orden abierta', 'órdenes abiertas')} con más de ${CFG.agingAlertDays} días desde su emisión`, s: 'Revisar avance con proveedores', go: { seg: viejas.filter(x => m.includes(x)).length >= viejas.length / 2 ? 'mtto' : 'oper', filter: { aging: true } } });
    const inc = d.rq.items.filter(x => x.incidencia);
    if (inc.length) alerts.push({ c: C.risk, n: inc.length, t: plural(inc.length, 'anticipo con incidencia reportada', 'anticipos con incidencia reportada'), s: inc.map(x => x.incidencia).slice(0, 2).join(' · '), go: { seg: 'rq', openDetail: true } });
    if (!d.mat.items.length) alerts.push({ c: '#C9CCDD', n: '—', t: 'Materiales críticos sin registros cargados', s: 'La hoja existe pero aún no tiene información', go: { seg: 'mat' } });

    $('#view-resumen').innerHTML = `
      <div class="section-head">
        <div><h2>Resumen general</h2><p>Visión consolidada de los cuatro segmentos. Haz clic en cualquier cuadro para ver su detalle.</p></div>
      </div>
      <div class="grid kpis" data-slot="res-kpi">
        ${kpi('Solicitudes registradas', fmtInt(ord.length), `${fmtInt(abiertas)} abiertas · Mantenimiento + Operaciones`, '', 'res-total')}
        ${kpi('Órdenes emitidas', pct(conOrden, ord.length) + '%', `${fmtInt(conOrden)} de ${fmtInt(ord.length)} con número de orden`, 'accent', 'res-emitidas')}
        ${kpi('Prioridad alta abiertas', fmtInt(alta), 'Requieren seguimiento cercano', 'risk', 'res-alta')}
        ${kpi('Monto comprometido', fmtMoneyShort(montoTot), `${fmtMoneyShort(antTot)} en anticipos (${pct(antTot, montoTot)}%)`, 'ok', 'res-monto')}
      </div>
      <div class="drill" id="drill-res-kpi"></div>
      <div class="grid cols-4" data-slot="res-seg">${segs}</div>
      <div class="drill" id="drill-res-seg"></div>
      <div class="grid cols-3">
        <div class="card">
          <h3>Estado de las órdenes por segmento</h3>
          <p class="sub">Estado estimado a partir de los comentarios · clic en una barra para ver esas órdenes</p>
          <div class="chart-box"><canvas id="ch-res-estado"></canvas></div>
        </div>
        <div class="card">
          <h3>Puntos de atención</h3>
          <p class="sub">Generados automáticamente · clic para ver el detalle</p>
          <ul class="alerts">${alerts.map((a, i) => `<li data-alert="${i}" tabindex="0"><span class="dot" style="background:${a.c}"></span><div><b>${a.n}</b> ${esc(a.t)}<small>${esc(a.s)}</small></div></li>`).join('') || '<li>Sin alertas</li>'}</ul>
        </div>
      </div>
      <div class="grid cols-2">
        <div class="card"><h3>Distribución por región</h3><p class="sub">Según el código de la orden o la descripción</p><div class="chart-box"><canvas id="ch-res-region"></canvas></div></div>
        <div class="card"><h3>Prioridad de las solicitudes</h3><p class="sub">Mantenimiento y Operaciones · clic para filtrar</p><div class="chart-box"><canvas id="ch-res-prio"></canvas></div></div>
      </div>`;

    $('#view-resumen').querySelectorAll('[data-alert]').forEach(li => {
      const go = () => runAction(alerts[+li.dataset.alert].go);
      li.onclick = go; li.onkeydown = e => { if (e.key === 'Enter') go(); };
    });

    builders.resumen = () => {
      const segLabels = ['Mantenimiento', 'Operaciones'], segKeys = ['mtto', 'oper'];
      const present = ESTADOS.map(e => e.key).filter(k => ord.some(x => x.estado === k));
      makeChart('ch-res-estado', {
        type: 'bar',
        data: { labels: segLabels, datasets: present.map(k => ({ label: k, data: [m.filter(x => x.estado === k).length, o.filter(x => x.estado === k).length], backgroundColor: estadoColor(k), borderRadius: 4, borderSkipped: false })) },
        options: {
          indexAxis: 'y', maintainAspectRatio: false, scales: { x: { stacked: true, grid: { color: C.line }, ticks: { precision: 0 } }, y: { stacked: true, grid: { display: false } } }, plugins: { legend: { position: 'bottom' } },
          onHover: pointer,
          onClick: (e, els) => { if (els.length) runAction({ seg: segKeys[els[0].index], filter: { estado: present[els[0].datasetIndex] } }); }
        }
      });
      const regItems = [...ord, ...d.rq.items.filter(x => !ord.some(z => z.orden && z.orden === x.orden))];
      makeChart('ch-res-region', regionCfg(regItems));
      const prioKeys = Object.keys(PRIO_COLORS);
      makeChart('ch-res-prio', {
        type: 'bar',
        data: { labels: segLabels, datasets: prioKeys.map(k => ({ label: k, data: [m.filter(x => x.prioridad === k).length, o.filter(x => x.prioridad === k).length], backgroundColor: PRIO_COLORS[k], borderRadius: 4 })) },
        options: {
          maintainAspectRatio: false, scales: { y: { grid: { color: C.line }, ticks: { precision: 0 } }, x: { grid: { display: false } } }, plugins: { legend: { position: 'bottom' } },
          onHover: pointer,
          onClick: (e, els) => { if (els.length) runAction({ seg: segKeys[els[0].index], filter: { prioridad: prioKeys[els[0].datasetIndex] } }); }
        }
      });
    };
  }

  /* ---------------- Segmentos de órdenes (MTTO / OPER) ---------------- */
  function defaultState(key) {
    if (key === 'rq') return { search: '', sort: 'monto', dir: -1 };
    return { search: '', estado: '', prioridad: '', owner: '', tipo: '', orden: '', sort: null, dir: 1, aging: false };
  }
  const isFiltered = st => !!(st.search || st.estado || st.prioridad || st.owner || st.tipo || st.orden || st.aging);

  function filteredRows(key) {
    const st = tableState[key] || (tableState[key] = defaultState(key));
    const q = norm(st.search);
    return DATA[key].items.filter(x =>
      (!q || x._txt.includes(q)) &&
      (!st.estado || x.estado === st.estado) &&
      (!st.prioridad || x.prioridad === st.prioridad) &&
      (!st.tipo || x.tipo === st.tipo) &&
      (!st.owner || x.owner === st.owner || x.procura === st.owner) &&
      (!st.orden || (st.orden === 'con') === x.hasOrder) &&
      (!st.aging || (x.fEm && x.dias > CFG.agingAlertDays && ABIERTO(x.estado))));
  }

  // Datos de los gráficos que responden a los filtros del detalle
  function tipoRegionData(rows, regs) {
    const tipos = [...new Set(rows.map(x => x.tipo))].sort();
    return { labels: tipos, series: regs.map(r => tipos.map(t => rows.filter(x => x.tipo === t && x.region === r).length)) };
  }
  function provData(rows) {
    const m = countBy(rows.filter(x => x.empresa), x => x.empresa);
    const top = Object.keys(m).sort((a, b) => m[b] - m[a] || a.localeCompare(b)).slice(0, 8);
    return { top, values: top.map(k => m[k]) };
  }

  function updateFilteredCharts(key, rows) {
    const all = DATA[key].items;
    const active = isFiltered(tableState[key]);
    document.querySelectorAll(`[data-flt="${key}"]`).forEach(t => { t.hidden = !active; t.textContent = `Filtrado · ${rows.length} de ${all.length}`; });
    const tc = charts[`ch-${key}-tipo`];
    if (tc) {
      const d = tipoRegionData(rows, tc.$regs);
      tc.data.labels = d.labels;
      tc.data.datasets.forEach((ds, i) => { ds.data = d.series[i]; });
      tc.update();
      $(`#ce-${key}-tipo`).hidden = rows.length > 0;
    }
    const pc = charts[`ch-${key}-prov`];
    if (pc) {
      const d = provData(rows);
      pc.$top = d.top;
      pc.data.labels = d.top.map(t => short(t, 28));
      pc.data.datasets[0].data = d.values;
      pc.update();
      $(`#ce-${key}-prov`).hidden = d.top.length > 0;
    }
  }

  function renderOrders(key, label, seg, desc) {
    const it = seg.items;
    tableState[key] = defaultState(key);
    if (!it.length) {
      $('#view-' + key).innerHTML = `<div class="section-head"><div><h2>${label}</h2><p>${desc}</p></div></div><div class="empty"><h3>Sin registros</h3><p>La hoja no contiene solicitudes.</p></div>`;
      return;
    }
    const conOrden = it.filter(x => x.hasOrder);
    const abiertas = it.filter(x => ABIERTO(x.estado));
    const edades = conOrden.filter(x => x.fEm && ABIERTO(x.estado)).map(x => x.dias);
    const edadProm = edades.length ? Math.round(edades.reduce((a, b) => a + b, 0) / edades.length) : null;
    const alta = it.filter(x => x.prioridad === 'Alta');
    const hasPrio = it.some(x => x.prioridad !== 'Sin definir');
    const hasOwner = it.some(x => x.owner !== 'Sin asignar');
    const fltTag = `<span class="flt-tag" data-flt="${key}" hidden></span>`;

    $('#view-' + key).innerHTML = `
      <div class="section-head"><div><h2>${label}</h2><p>${desc} Haz clic en cualquier cuadro para ver su detalle.</p></div></div>
      <div class="grid kpis" data-slot="${key}-kpi">
        ${kpi('Solicitudes', fmtInt(it.length), `${fmtInt(abiertas.length)} abiertas`, '', `${key}-total`)}
        ${kpi('Con orden emitida', fmtInt(conOrden.length), `${pct(conOrden.length, it.length)}% del total · ${conOrden.filter(x => x.tipo === 'ODS').length} ODS / ${conOrden.filter(x => x.tipo === 'ODC').length} ODC`, 'accent', `${key}-orden`)}
        ${kpi('En gestión (sin orden)', fmtInt(it.length - conOrden.length), 'Cotización, validación o elaboración', '', `${key}-sinorden`)}
        ${hasPrio ? kpi('Prioridad alta', fmtInt(alta.length), `${pct(alta.length, it.length)}% de las solicitudes`, 'risk', `${key}-alta`)
                  : kpi('Por aprobación', fmtInt(it.filter(x => x.estado === 'Por aprobación').length), 'En espera de firma', 'risk', `${key}-aprob`)}
        ${kpi('Antigüedad promedio', edadProm !== null ? edadProm + ' días' : '—', edadProm !== null ? 'Órdenes abiertas, desde su emisión' : 'Sin fechas de emisión registradas', 'ok', `${key}-edad`)}
      </div>
      <div class="drill" id="drill-${key}-kpi"></div>
      <div class="grid ${hasOwner ? 'cols-3eq' : 'cols-2'}">
        <div class="card"><h3>Estado de las solicitudes</h3><p class="sub">Clic en una barra para filtrar el detalle</p><div class="chart-box"><canvas id="ch-${key}-estado"></canvas></div></div>
        <div class="card"><h3>Carga por responsable de Procura</h3><p class="sub">Clic en una columna para filtrar</p><div class="chart-box"><canvas id="ch-${key}-proc"></canvas></div></div>
        ${hasOwner ? `<div class="card"><h3>${esc(seg.ownerLabel)}</h3><p class="sub">Clic en una columna para filtrar</p><div class="chart-box"><canvas id="ch-${key}-owner"></canvas></div></div>` : ''}
      </div>
      <div class="grid cols-2">
        <div class="card"><h3>Tipo de orden y región ${fltTag}</h3><p class="sub">ODS = servicio · ODC = compra · se ajusta a los filtros del detalle</p>
          <div class="chart-box"><canvas id="ch-${key}-tipo"></canvas><div class="chart-empty" id="ce-${key}-tipo" hidden>Sin datos con los filtros aplicados</div></div></div>
        <div class="card"><h3>Principales proveedores ${fltTag}</h3><p class="sub">Número de órdenes por empresa · se ajusta a los filtros del detalle</p>
          <div class="chart-box"><canvas id="ch-${key}-prov"></canvas><div class="chart-empty" id="ce-${key}-prov" hidden>Sin proveedores con los filtros aplicados</div></div></div>
      </div>
      ${detailBlock(key, it.length, `
        <div class="filters">
          <input type="search" placeholder="Buscar por orden, solicitud, empresa o comentario…" data-f="search">
          <select data-f="estado"><option value="">Todos los estados</option>${ESTADOS.filter(e => it.some(x => x.estado === e.key)).map(e => `<option>${e.key}</option>`).join('')}</select>
          ${hasPrio ? `<select data-f="prioridad"><option value="">Todas las prioridades</option>${Object.keys(PRIO_COLORS).filter(p => it.some(x => x.prioridad === p)).map(p => `<option>${p}</option>`).join('')}</select>` : ''}
          <select data-f="tipo"><option value="">Todos los tipos</option>${[...new Set(it.map(x => x.tipo))].sort().map(t => `<option>${esc(t)}</option>`).join('')}</select>
          <select data-f="orden"><option value="">Con y sin orden</option><option value="con">Con orden emitida</option><option value="sin">Sin orden</option></select>
          <select data-f="owner"><option value="">Todos los responsables</option>${[...new Set(it.flatMap(x => [x.owner, x.procura]))].filter(x => x !== 'Sin asignar').sort().map(t => `<option>${esc(t)}</option>`).join('')}</select>
          <button type="button" class="btn btn--light" data-reset>Limpiar</button>
          <span class="count" id="count-${key}"></span>
        </div>
        <div class="table-wrap"><table id="tbl-${key}"></table></div>
        <p class="note">El <b>estado</b> se estima automáticamente según palabras clave del campo “Comentarios” (p. ej. “POR APROB”, “ANTICIPO”, “ENTREGA PENDIENTE”). Para el dato exacto consulta el comentario original o descarga el Excel.</p>
      `)}`;

    bindDetail(key);
    refreshTable(key);

    builders[key] = () => {
      const est = ESTADOS.filter(e => it.some(x => x.estado === e.key));
      const ec = hbar(est.map(e => e.key), est.map(e => it.filter(x => x.estado === e.key).length), est.map(e => e.color));
      ec.options.onHover = pointer;
      ec.options.onClick = (e, els) => { if (els.length) setFilter(key, { estado: est[els[0].index].key }); };
      makeChart(`ch-${key}-estado`, ec);

      const stackedBy = (field, canvas) => {
        const people = [...new Set(it.map(x => x[field]))].sort((a, b) => it.filter(x => x[field] === b).length - it.filter(x => x[field] === a).length);
        makeChart(canvas, {
          type: 'bar',
          data: { labels: people, datasets: est.map(e => ({ label: e.key, data: people.map(p => it.filter(x => x[field] === p && x.estado === e.key).length), backgroundColor: e.color, borderRadius: 3 })) },
          options: {
            maintainAspectRatio: false, scales: { x: { stacked: true, grid: { display: false } }, y: { stacked: true, grid: { color: C.line }, ticks: { precision: 0 } } },
            plugins: { legend: { position: 'bottom', labels: { font: { size: 11 } } } }, onHover: pointer,
            onClick: (e, els) => { if (els.length && people[els[0].index] !== 'Sin asignar') setFilter(key, { owner: people[els[0].index] }); }
          }
        });
      };
      stackedBy('procura', `ch-${key}-proc`);
      if (hasOwner) stackedBy('owner', `ch-${key}-owner`);

      const rows = filteredRows(key);
      const regs = Object.keys(REGION_COLORS).filter(r => it.some(x => x.region === r));
      const td = tipoRegionData(rows, regs);
      const tc = makeChart(`ch-${key}-tipo`, {
        type: 'bar',
        data: { labels: td.labels, datasets: regs.map((r, i) => ({ label: r, data: td.series[i], backgroundColor: REGION_COLORS[r], borderRadius: 3, maxBarThickness: 90 })) },
        options: { maintainAspectRatio: false, scales: { x: { stacked: true, grid: { display: false } }, y: { stacked: true, grid: { color: C.line }, ticks: { precision: 0 } } }, plugins: { legend: { position: 'bottom' } } }
      });
      tc.$regs = regs;

      const pd = provData(rows);
      const pc = makeChart(`ch-${key}-prov`, {
        type: 'bar',
        data: { labels: pd.top.map(t => short(t, 28)), datasets: [{ data: pd.values, backgroundColor: C.navy, borderRadius: 6, maxBarThickness: 24 }] },
        options: {
          indexAxis: 'y', maintainAspectRatio: false,
          plugins: { legend: { display: false }, tooltip: { callbacks: { title: i => i[0].chart.$top[i[0].dataIndex] } } },
          scales: { x: { grid: { color: C.line }, ticks: { precision: 0 } }, y: { grid: { display: false }, ticks: { autoSkip: false } } }
        }
      });
      pc.$top = pd.top;
      updateFilteredCharts(key, rows);
    };
  }

  function setFilter(key, f) {
    tableState[key] = Object.assign(defaultState(key), f);
    refreshTable(key);
    openDetail(key);
  }

  function bindDetail(key) {
    const root = $('#detail-' + key);
    root.querySelector('[data-detail]').onclick = () => root.classList.toggle('is-open');
    const onSearch = debounce(() => refreshTable(key), 150);
    root.querySelectorAll('[data-f]').forEach(el => {
      el.addEventListener(el.tagName === 'INPUT' ? 'input' : 'change', () => {
        tableState[key][el.dataset.f] = el.value;
        tableState[key].aging = false;
        if (el.tagName === 'INPUT') onSearch(); else refreshTable(key);
      });
    });
    const reset = root.querySelector('[data-reset]');
    if (reset) reset.onclick = () => { tableState[key] = defaultState(key); refreshTable(key); };
  }

  function refreshTable(key) {
    if (key === 'rq') return refreshPayTable();
    if (key === 'mat' || !DATA[key] || !DATA[key].items.length) return;
    const seg = DATA[key];
    const st = tableState[key] || (tableState[key] = defaultState(key));
    const root = $('#detail-' + key);
    root.querySelectorAll('[data-f]').forEach(el => { if (el.value !== (st[el.dataset.f] || '')) el.value = st[el.dataset.f] || ''; });

    let rows = filteredRows(key);
    const cols = [
      { k: 'orden', t: 'No. de orden', f: x => x.orden ? `<span class="mono">${esc(x.orden)}</span>` : '<span class="muted">Sin orden</span>', cls: 'nowrap' },
      { k: 'solicitud', t: 'Solicitud', f: x => esc(x.solicitud), cls: 'wide' },
      { k: 'empresa', t: 'Empresa', f: x => esc(x.empresa) || dash, cls: 'mid' },
      { k: 'tipo', t: 'Tipo', f: x => esc(x.tipo), cls: 'nowrap' },
      { k: 'prioridad', t: 'Prioridad', f: x => prioBadge(x.prioridad), sortVal: x => ({ Alta: 0, Media: 1, Baja: 2 }[x.prioridad] ?? 3) },
      { k: 'estado', t: 'Estado (estimado)', f: x => estadoBadge(x.estado) },
      { k: 'fEm', t: 'Emisión', f: x => fmtDate(x.fEm) || (x.fSol ? `<span class="muted">Sol. ${fmtDate(x.fSol)}</span>` : dash), cls: 'nowrap', sortVal: x => (x.fEm || x.fSol) ? (x.fEm || x.fSol).getTime() : 0 },
      { k: 'dias', t: 'Días', f: x => x.dias ?? dash, cls: 'num', sortVal: x => x.dias ?? -1 },
      { k: 'owner', t: seg.ownerLabel.replace(/^Encargado\s*/i, 'Enc. '), f: x => esc(x.owner), cls: 'nowrap' },
      { k: 'procura', t: 'Enc. Procura', f: x => esc(x.procura), cls: 'nowrap' },
      { k: 'coment', t: 'Comentarios', f: x => esc(x.coment) || dash, cls: 'wide' }
    ].filter(c => c.k !== 'prioridad' || seg.items.some(x => x.prioridad !== 'Sin definir'))
     .filter(c => c.k !== 'owner' || seg.items.some(x => x.owner !== 'Sin asignar'));

    if (st.sort) {
      const c = cols.find(z => z.k === st.sort);
      const val = c && c.sortVal ? c.sortVal : (x => norm(x[st.sort]));
      rows = rows.slice().sort((a, b) => (val(a) > val(b) ? 1 : val(a) < val(b) ? -1 : 0) * st.dir);
    }

    const tbl = $('#tbl-' + key);
    tbl.innerHTML = `<thead><tr>${cols.map(c => `<th data-sort="${c.k}">${esc(c.t)}<span class="sort">${st.sort === c.k ? (st.dir > 0 ? '▲' : '▼') : '↕'}</span></th>`).join('')}</tr></thead>
      <tbody>${rows.map(x => `<tr>${cols.map(c => `<td class="${c.cls || ''}">${c.f(x)}</td>`).join('')}</tr>`).join('') || `<tr><td colspan="${cols.length}" class="muted" style="text-align:center;padding:24px">Sin resultados con los filtros aplicados</td></tr>`}</tbody>`;
    $('#count-' + key).textContent = `${rows.length} de ${seg.items.length} registros${st.aging ? ` · solo órdenes abiertas > ${CFG.agingAlertDays} días` : ''}`;
    tbl.querySelectorAll('th').forEach(th => th.onclick = () => { st.dir = st.sort === th.dataset.sort ? -st.dir : 1; st.sort = th.dataset.sort; refreshTable(key); });
    updateFilteredCharts(key, rows);
  }

  /* ---------------- Materiales críticos ---------------- */
  function renderMaterials(seg) {
    const v = $('#view-mat');
    const head = `<div class="section-head"><div><h2>Materiales críticos</h2><p>Inventario de materiales críticos con sus niveles mínimos y máximos de reposición. Haz clic en cualquier cuadro para ver su detalle.</p></div></div>`;
    const hasStock = seg.ix && seg.ix.stock >= 0;
    const bajo = seg.items.filter(x => x.bajoMinimo).length;
    const tiles = `<div class="grid kpis" data-slot="mat-kpi">
        ${kpi('Materiales registrados', fmtInt(seg.items.length), seg.items.length ? 'Materiales bajo control' : 'Aún no hay información cargada', '', 'mat-total')}
        ${kpi('Bajo nivel mínimo', hasStock ? fmtInt(bajo) : '—', hasStock ? `${pct(bajo, seg.items.length)}% del total` : 'Se calcula al registrar existencias', 'risk', 'mat-bajo')}
        ${kpi('Campos de control', fmtInt(seg.headers.length), 'Columnas de la hoja', '', 'mat-campos')}
      </div><div class="drill" id="drill-mat-kpi"></div>`;
    if (!seg.items.length) {
      v.innerHTML = head + tiles + `
        <div class="empty">
          <h3>Esta sección aún no tiene registros</h3>
          <p>Cuando se completen filas en la hoja <b>${esc(DATA.sheetNames.mat || 'Materiales Criticos')}</b> del Excel, aparecerán aquí automáticamente.</p>
          <div class="cols">${seg.headers.map(h => `<span>${esc(h)}</span>`).join('')}</div>
          <p class="note">Sugerencia: agregar una columna <b>“Existencia actual”</b> permitirá alertar los materiales por debajo del mínimo.</p>
        </div>`;
      return;
    }
    v.innerHTML = head + tiles + detailBlock('mat', seg.items.length, `
        <div class="table-wrap"><table>
          <thead><tr>${seg.headers.map(h => `<th style="cursor:default">${esc(h)}</th>`).join('')}</tr></thead>
          <tbody>${seg.items.map(x => `<tr${x.bajoMinimo ? ' style="outline:2px solid #F6C7C5"' : ''}>${x.cells.map(c => `<td class="${typeof c === 'number' ? 'num' : ''}">${esc(c)}</td>`).join('')}</tr>`).join('')}</tbody>
        </table></div>`);
    $('#detail-mat [data-detail]').onclick = () => $('#detail-mat').classList.toggle('is-open');
  }

  /* ---------------- Anticipos y montos (RQ) ---------------- */
  const filteredPay = () => { const q = norm((tableState.rq || {}).search); return DATA.rq.items.filter(x => !q || x._txt.includes(q)); };

  function renderPayments(seg) {
    const v = $('#view-rq');
    const it = seg.items;
    tableState.rq = defaultState('rq');
    const head = `<div class="section-head"><div><h2>Anticipos y montos</h2><p>Montos de las órdenes con requerimiento de anticipo y porcentaje acordado con el proveedor. Haz clic en cualquier cuadro para ver su detalle.</p></div></div>`;
    if (!it.length) { v.innerHTML = head + '<div class="empty"><h3>Sin registros</h3></div>'; return; }
    const tot = sumBy(it, x => x.monto), ant = sumBy(it, x => x.anticipo);
    const inc = it.filter(x => x.incidencia);
    const fltTag = '<span class="flt-tag" data-flt="rq" hidden></span>';

    v.innerHTML = head + `
      <div class="grid kpis" data-slot="rq-kpi">
        ${kpi('Monto total', fmtMoney(tot), `${fmtInt(it.length)} órdenes`, '', 'rq-total')}
        ${kpi('Anticipos', fmtMoney(ant), `${pct(ant, tot)}% del monto total`, 'accent', 'rq-ant')}
        ${kpi('Saldo por pagar', fmtMoney(tot - ant), 'Pendiente contra entrega / valuación', 'ok', 'rq-saldo')}
        ${kpi('Con incidencia', fmtInt(inc.length), inc.length ? esc(inc.map(x => x.incidencia).join(' · ')) : 'Sin incidencias', 'risk', 'rq-incid')}
      </div>
      <div class="drill" id="drill-rq-kpi"></div>
      <div class="grid cols-2">
        <div class="card"><h3>Anticipo vs. saldo por orden ${fltTag}</h3><p class="sub">Montos en ${CFG.currency} · se ajusta a la búsqueda del detalle</p>
          <div class="chart-box tall"><canvas id="ch-rq-orden"></canvas><div class="chart-empty" id="ce-rq-orden" hidden>Sin datos con la búsqueda aplicada</div></div></div>
        <div class="card"><h3>Monto por proveedor ${fltTag}</h3><p class="sub">Participación en el monto comprometido</p>
          <div class="chart-box tall"><canvas id="ch-rq-emp"></canvas><div class="chart-empty" id="ce-rq-emp" hidden>Sin datos con la búsqueda aplicada</div></div></div>
      </div>
      ${detailBlock('rq', it.length, `
        <div class="filters">
          <input type="search" placeholder="Buscar por orden, solicitud o empresa…" data-f="search">
          <button type="button" class="btn btn--light" data-reset>Limpiar</button>
          <span class="count" id="count-rq"></span>
        </div>
        <div class="table-wrap"><table id="tbl-rq"></table></div>`)}`;

    const root = $('#detail-rq');
    root.querySelector('[data-detail]').onclick = () => root.classList.toggle('is-open');
    const onSearch = debounce(refreshPayTable, 150);
    root.querySelector('[data-f]').oninput = e => { tableState.rq.search = e.target.value; onSearch(); };
    root.querySelector('[data-reset]').onclick = () => { tableState.rq = defaultState('rq'); refreshPayTable(); };
    refreshPayTable();

    builders.rq = () => {
      const rows = filteredPay().sort((a, b) => b.monto - a.monto);
      const c1 = makeChart('ch-rq-orden', payStackCfg(rows)); c1.$rows = rows;
      makeChart('ch-rq-emp', provDonutCfg(rows));
      updateRqCharts(filteredPay());
    };
  }

  function updateRqCharts(rows) {
    const active = !!(tableState.rq && tableState.rq.search);
    document.querySelectorAll('[data-flt="rq"]').forEach(t => { t.hidden = !active; t.textContent = `Filtrado · ${rows.length} de ${DATA.rq.items.length}`; });
    const c1 = charts['ch-rq-orden'], c2 = charts['ch-rq-emp'];
    if (c1) {
      const s = rows.slice().sort((a, b) => b.monto - a.monto);
      const cfg = payStackCfg(s);
      c1.$rows = s;
      c1.data.labels = cfg.data.labels;
      c1.data.datasets[0].data = cfg.data.datasets[0].data;
      c1.data.datasets[1].data = cfg.data.datasets[1].data;
      c1.data.datasets[1].backgroundColor = cfg.data.datasets[1].backgroundColor;
      c1.update();
      $('#ce-rq-orden').hidden = rows.length > 0;
    }
    if (c2) {
      const cfg = provDonutCfg(rows);
      c2.data.labels = cfg.data.labels;
      c2.data.datasets[0].data = cfg.data.datasets[0].data;
      c2.data.datasets[0].backgroundColor = cfg.data.datasets[0].backgroundColor;
      c2.update();
      $('#ce-rq-emp').hidden = rows.length > 0;
    }
  }

  function refreshPayTable() {
    if (!DATA || !DATA.rq.items.length) return;
    const it = DATA.rq.items;
    const st = tableState.rq || (tableState.rq = defaultState('rq'));
    const input = $('#detail-rq [data-f]');
    if (input && input.value !== st.search) input.value = st.search;
    let rows = filteredPay();
    const cols = [
      { k: 'orden', t: 'No. de orden', f: x => `<span class="mono">${esc(x.orden)}</span>`, cls: 'nowrap' },
      { k: 'solicitud', t: 'Solicitud', f: x => esc(x.solicitud), cls: 'wide' },
      { k: 'empresa', t: 'Empresa', f: x => esc(x.empresa), cls: 'mid' },
      { k: 'monto', t: 'Monto', f: x => fmtMoney(x.monto), cls: 'num' },
      { k: 'pct', t: '% Anticipo', f: x => x.pct !== null ? Math.round(x.pct * 100) + '%' : `<span class="badge prio-alta">${esc(x.incidencia)}</span>`, cls: 'num', sortVal: x => x.pct ?? -1 },
      { k: 'anticipo', t: 'Anticipo', f: x => x.pct !== null ? fmtMoney(x.anticipo) : dash, cls: 'num' },
      { k: 'saldo', t: 'Saldo', f: x => fmtMoney(x.saldo), cls: 'num' }
    ];
    if (st.sort) {
      const c = cols.find(z => z.k === st.sort);
      const val = c && c.sortVal ? c.sortVal : (x => typeof x[st.sort] === 'number' ? x[st.sort] : norm(x[st.sort]));
      rows = rows.slice().sort((a, b) => (val(a) > val(b) ? 1 : val(a) < val(b) ? -1 : 0) * st.dir);
    }
    const tot = sumBy(rows, x => x.monto), ant = sumBy(rows, x => x.anticipo);
    const tf = 'font-weight:700;background:#F3F4F8';
    const tbl = $('#tbl-rq');
    tbl.innerHTML = `<thead><tr>${cols.map(c => `<th data-sort="${c.k}">${c.t}<span class="sort">${st.sort === c.k ? (st.dir > 0 ? '▲' : '▼') : '↕'}</span></th>`).join('')}</tr></thead>
      <tbody>${rows.map(x => `<tr>${cols.map(c => `<td class="${c.cls}">${c.f(x)}</td>`).join('')}</tr>`).join('') || `<tr><td colspan="7" class="muted" style="text-align:center;padding:24px">Sin resultados</td></tr>`}
      <tr><td colspan="3" style="${tf}">Total</td><td class="num" style="${tf}">${fmtMoney(tot)}</td><td style="${tf}"></td><td class="num" style="${tf}">${fmtMoney(ant)}</td><td class="num" style="${tf}">${fmtMoney(tot - ant)}</td></tr></tbody>`;
    $('#count-rq').textContent = `${rows.length} de ${it.length} registros`;
    tbl.querySelectorAll('th').forEach(th => th.onclick = () => { st.dir = st.sort === th.dataset.sort ? -st.dir : 1; st.sort = th.dataset.sort; refreshPayTable(); });
    updateRqCharts(rows);
  }

  /* ---------------- Navegación ---------------- */
  function route() {
    const tab = (location.hash || '#resumen').slice(1);
    const v = VIEWS.includes(tab) ? tab : 'resumen';
    if (v !== curView) {
      document.querySelectorAll('.tab').forEach(a => a.classList.toggle('is-active', a.dataset.tab === v));
      document.querySelectorAll('.view').forEach(s => s.classList.toggle('is-active', s.id === 'view-' + v));
      if (curView !== null) window.scrollTo(0, 0);
      curView = v;
    }
    if (DATA && !built[v] && builders[v]) {
      built[v] = true;
      setTimeout(() => { try { builders[v](); } catch (err) { console.error(err); } }, 0);
    }
  }

  /* ---------------- Eventos globales ---------------- */
  window.addEventListener('hashchange', route);
  document.addEventListener('click', e => {
    const up = e.target.closest('[data-upload]');
    if (up) { e.preventDefault(); requestUpload(); return; }
    const close = e.target.closest('[data-drill-close]');
    if (close) { closeDrill(close.dataset.drillClose); return; }
    const go = e.target.closest('[data-drill-go]');
    if (go) { runAction((drillSpecs[go.dataset.drillGo] || {}).action); return; }
    const t = e.target.closest('[data-drill]');
    if (t) {
      const grid = t.closest('[data-slot]');
      if (grid) { e.preventDefault(); toggleDrill(grid.dataset.slot, t.dataset.drill, t); }
    }
  });
  document.addEventListener('keydown', e => {
    if ((e.key === 'Enter' || e.key === ' ') && e.target.matches && e.target.matches('.seg[data-drill]')) {
      e.preventDefault();
      const grid = e.target.closest('[data-slot]');
      toggleDrill(grid.dataset.slot, e.target.dataset.drill, e.target);
    }
  });
  $('#fileInput').addEventListener('change', e => { if (e.target.files[0]) loadLocalFile(e.target.files[0]); e.target.value = ''; });
  $('#btnPrint').addEventListener('click', () => {
    const v = document.querySelector('.view.is-active');
    if (v) v.querySelectorAll('.detail').forEach(d => d.classList.add('is-open'));
    window.print();
  });

  // Estado inicial
  document.querySelectorAll('.view').forEach(v => { v.innerHTML = '<div class="loading"><div><div class="spinner"></div>Cargando datos del Excel…</div></div>'; });
  route();
  loadDefault();
})();
