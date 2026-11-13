/* =========================================================================
   COLVEN · Dashboard de Seguimiento de Equipos y Mantenimiento
   Lee el Excel (misma estructura) y construye el reporte en el navegador.
   ========================================================================= */
(function () {
  'use strict';

  const CFG = window.DASHBOARD_CONFIG;
  const $ = (s, el = document) => el.querySelector(s);
  const charts = {};              // registro de gráficos para destruir al recargar
  const tableState = {};          // filtros y orden por segmento
  let DATA = null;                // datos procesados
  let downloadUrl = null;         // object URL cuando se carga un archivo local

  /* ---------------- Paleta ---------------- */
  const C = {
    navy: '#1D2041', navy3: '#3C4180', yellow: '#FDD649', gold: '#E3A92B',
    muted: '#646985', line: '#E2E4EE', ok: '#2E9E6B', risk: '#D9534F'
  };
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
  const REGION_COLORS = { 'Apure': '#1D2041', 'Barinas': '#FDD649', 'Apure / Barinas': '#7F86B5', 'Sin definir': '#E2E4EE' };

  /* ---------------- Utilidades ---------------- */
  const norm = s => String(s ?? '').normalize('NFD').replace(/[̀-ͯ]/g, '')
    .toLowerCase().replace(/\s+/g, ' ').trim();
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
  const titleCase = s => s ? s.charAt(0).toUpperCase() + s.slice(1).toLowerCase() : '';

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
    if (/por anticipo|pendiente pago|pendiente de pago|sin informacion de pago|anticipo para inicio|para inicio/.test(c)) return 'Pago / anticipo pendiente';
    if (/entrega pendiente|fabricacion|por entregar|en transito/.test(c)) return 'Entrega pendiente';
    if (!hasOrder) return 'En gestión (sin orden)';
    if (/anticipo|pago|ejecucion|termina|extender|otrosi|valuacion|enviada/.test(c)) return 'En ejecución';
    if (!c) return 'Sin comentario';
    return 'En seguimiento';
  }
  const ABIERTO = e => e !== 'Completado' && e !== 'Canalizado / no aprobado';

  /* ---------------- Lectura del Excel ---------------- */
  function sheetRows(ws) {
    return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '', blankrows: false });
  }
  function findHeaderRow(rows) {
    for (let i = 0; i < Math.min(rows.length, 15); i++) {
      const cells = rows[i].map(norm).filter(Boolean);
      if (cells.length >= 3 && cells.some(c => c.startsWith('no'))) return i;
    }
    return 0;
  }
  function headerIndex(h, test) { return h.findIndex(test); }

  function pickSheets(wb) {
    const hidden = new Set(((wb.Workbook || {}).Sheets || []).filter(s => s.Hidden).map(s => s.name));
    const visible = wb.SheetNames.filter(n => !hidden.has(n));
    const out = {};
    const used = new Set();
    for (const key of ['mtto', 'oper', 'mat', 'rq']) {
      const name = visible.find(n => !used.has(n) && CFG.sheets[key].test(n));
      if (name) { out[key] = name; used.add(name); }
    }
    // respaldo por posición
    const rest = visible.filter(n => !used.has(n));
    for (const key of ['mtto', 'oper', 'mat', 'rq']) if (!out[key] && rest.length) out[key] = rest.shift();
    return out;
  }

  function parseOrders(ws) {
    const rows = sheetRows(ws);
    const hi = findHeaderRow(rows);
    const rawH = (rows[hi] || []).map(clean);
    const h = rawH.map(norm);
    const ix = {
      no: headerIndex(h, c => /^no\.?$/.test(c)),
      orden: headerIndex(h, c => c.includes('orden') && !c.includes('fecha')),
      solicitud: headerIndex(h, c => c === 'solicitud' || c === 'descripcion'),
      empresa: headerIndex(h, c => c.includes('empresa') || c.includes('proveedor')),
      tipo: headerIndex(h, c => c.includes('tipo')),
      prioridad: headerIndex(h, c => c.includes('prioridad')),
      fSol: headerIndex(h, c => c.includes('fecha') && c.includes('solicitud')),
      fEm: headerIndex(h, c => c.includes('fecha') && (c.includes('emision') || c.includes('orden'))),
      owner: headerIndex(h, c => c.startsWith('encargado') && !c.includes('procura')),
      procura: headerIndex(h, c => c.includes('procura')),
      coment: headerIndex(h, c => c.includes('comentario') || c.includes('observ'))
    };
    const get = (r, k) => ix[k] >= 0 ? r[ix[k]] : '';
    const items = [];
    for (const r of rows.slice(hi + 1)) {
      const orden = clean(get(r, 'orden'));
      const solicitud = clean(get(r, 'solicitud'));
      const empresa = clean(get(r, 'empresa'));
      if (!orden && !solicitud && !empresa) continue;           // fila vacía (solo correlativo)
      const hasOrder = /^OD[SC]|^\w{2,}-/i.test(orden);
      let tipo = clean(get(r, 'tipo')).toUpperCase();
      if (!tipo && hasOrder) tipo = orden.slice(0, 3).toUpperCase();
      const prioridad = titleCase(clean(get(r, 'prioridad'))) || 'Sin definir';
      const coment = clean(get(r, 'coment'));
      const fSol = toDate(get(r, 'fSol'));
      const fEm = toDate(get(r, 'fEm'));
      const estado = inferEstado(coment, hasOrder);
      items.push({
        no: get(r, 'no'), orden, solicitud, empresa: empresa || '', tipo: tipo || 'Sin definir',
        prioridad, fSol, fEm, owner: clean(get(r, 'owner')) || 'Sin asignar',
        procura: clean(get(r, 'procura')) || 'Sin asignar', coment, hasOrder, estado,
        region: region(orden, solicitud),
        dias: fEm ? daysSince(fEm) : (fSol ? daysSince(fSol) : null),
        diasDesde: fEm ? 'emisión' : (fSol ? 'solicitud' : '')
      });
    }
    return { items, ownerLabel: ix.owner >= 0 ? rawH[ix.owner] : 'Encargado', headers: rawH };
  }

  function parsePayments(ws) {
    const rows = sheetRows(ws);
    const hi = findHeaderRow(rows);
    const h = (rows[hi] || []).map(norm);
    const ix = {
      orden: headerIndex(h, c => c.includes('orden')),
      solicitud: headerIndex(h, c => c === 'solicitud' || c.includes('descripcion')),
      empresa: headerIndex(h, c => c.includes('empresa') || c.includes('proveedor')),
      monto: headerIndex(h, c => c.includes('monto')),
      pct: headerIndex(h, c => c.includes('%') || c.includes('anticipo'))
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
      items.push({
        orden, solicitud, empresa, monto, pct: p, incidencia,
        anticipo: p !== null ? monto * p : 0,
        saldo: p !== null ? monto * (1 - p) : monto,
        tipo: orden ? orden.slice(0, 3).toUpperCase() : 'Sin definir',
        region: region(orden, solicitud)
      });
    }
    return { items };
  }

  function parseMaterials(ws) {
    const rows = sheetRows(ws);
    const hi = findHeaderRow(rows);
    const headers = (rows[hi] || []).map(clean).filter(Boolean);
    const h = headers.map(norm);
    const ix = {
      material: headerIndex(h, c => c.includes('material')),
      min: headerIndex(h, c => c.includes('minima') || c.includes('minimo')),
      max: headerIndex(h, c => c.includes('maxima') || c.includes('maximo')),
      stock: headerIndex(h, c => c.includes('existencia') || c.includes('stock') || c.includes('actual') || c.includes('inventario'))
    };
    const items = [];
    for (const r of rows.slice(hi + 1)) {
      const vals = headers.map((_, i) => r[i]);
      const filled = vals.filter((v, i) => clean(v) && !/^no\.?$/.test(h[i]));
      if (!filled.length) continue;
      const num = v => typeof v === 'number' ? v : parseFloat(String(v).replace(',', '.'));
      const it = { cells: vals };
      if (ix.stock >= 0 && ix.min >= 0) {
        const s = num(vals[ix.stock]), mn = num(vals[ix.min]);
        it.bajoMinimo = !isNaN(s) && !isNaN(mn) && s < mn;
      }
      items.push(it);
    }
    return { items, headers, ix };
  }

  function processWorkbook(wb, meta) {
    const names = pickSheets(wb);
    const d = { meta, sheetNames: names };
    d.mtto = names.mtto ? parseOrders(wb.Sheets[names.mtto]) : { items: [], ownerLabel: 'Encargado', headers: [] };
    d.oper = names.oper ? parseOrders(wb.Sheets[names.oper]) : { items: [], ownerLabel: 'Encargado', headers: [] };
    d.mat = names.mat ? parseMaterials(wb.Sheets[names.mat]) : { items: [], headers: [], ix: {} };
    d.rq = names.rq ? parsePayments(wb.Sheets[names.rq]) : { items: [] };
    d.modified = (wb.Props && (wb.Props.ModifiedDate || wb.Props.CreatedDate)) || null;
    return d;
  }

  /* ---------------- Carga ---------------- */
  async function loadDefault() {
    const url = encodeURI(CFG.dataFile);
    $('#btnDownload').setAttribute('href', url);
    $('#btnDownload').setAttribute('download', CFG.dataFile.split('/').pop());
    try {
      const res = await fetch(url + '?v=' + Date.now(), { cache: 'no-store' });
      if (!res.ok) throw new Error('HTTP ' + res.status);
      const buf = await res.arrayBuffer();
      const wb = XLSX.read(buf, { type: 'array', cellDates: false });
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
        <p style="margin-top:16px"><label class="btn btn--primary" style="display:inline-flex">Cargar Excel
          <input type="file" accept=".xlsx,.xlsm,.xls" hidden onchange="window.__loadFile(this.files[0])"></label></p>
      </div>`;
    ['resumen', 'mtto', 'oper', 'mat', 'rq'].forEach(k => { $('#view-' + k).innerHTML = html; });
  }

  /* ---------------- Render principal ---------------- */
  function render(d) {
    DATA = d;
    Object.values(charts).forEach(c => c.destroy());
    for (const k in charts) delete charts[k];

    const mod = d.meta.lastModified ? new Date(d.meta.lastModified) : (d.modified ? new Date(d.modified) : null);
    const modTxt = mod && !isNaN(mod) ? mod.toLocaleString('es-VE', { day: '2-digit', month: 'long', year: 'numeric', hour: '2-digit', minute: '2-digit' }) : '—';
    $('#dataMeta').textContent = `Datos actualizados: ${modTxt} · Corte al ${new Date().toLocaleDateString('es-VE', { day: '2-digit', month: 'long', year: 'numeric' })}`;
    $('#footerSource').textContent = `Fuente: ${d.meta.fileName}`;

    renderResumen(d);
    renderOrders('mtto', 'Mantenimiento', d.mtto, 'Órdenes de servicio (ODS) y de compra (ODC) asociadas al mantenimiento de equipos.');
    renderOrders('oper', 'Operaciones', d.oper, 'Órdenes de servicio y compra solicitadas por el área de Operaciones.');
    renderMaterials(d.mat);
    renderPayments(d.rq);
    route();
  }

  /* ---------------- Componentes ---------------- */
  const kpi = (label, value, hint, cls = '') =>
    `<div class="card kpi ${cls}"><div class="kpi__label">${label}</div><div class="kpi__value">${value}</div><div class="kpi__hint">${hint || '&nbsp;'}</div></div>`;

  const ICONS = {
    mtto: '<path d="M14.7 6.3a4 4 0 0 0-5.4 5.4L3 18l3 3 6.3-6.3a4 4 0 0 0 5.4-5.4l-2.6 2.6-2.4-.6-.6-2.4z"/>',
    oper: '<path d="M3 21h18M5 21V9l7-5 7 5v12M9 21v-6h6v6"/>',
    mat: '<path d="M21 8 12 3 3 8v8l9 5 9-5zM3 8l9 5 9-5M12 13v8"/>',
    rq: '<path d="M12 2v20M17 6H9.5a3.5 3.5 0 0 0 0 7h5a3.5 3.5 0 0 1 0 7H6"/>'
  };

  function detailBlock(key, count, inner) {
    return `<div class="detail" id="detail-${key}">
      <button class="detail__toggle" data-detail="${key}">
        <span>Ver detalle completo <span class="muted">(${count} registros)</span></span>
        <svg class="chev" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="m6 9 6 6 6-6"/></svg>
      </button>
      <div class="detail__body">${inner}</div>
    </div>`;
  }

  function openDetail(key) {
    const el = $('#detail-' + key);
    if (el) { el.classList.add('is-open'); setTimeout(() => el.scrollIntoView({ behavior: 'smooth', block: 'start' }), 50); }
  }

  function makeChart(id, config) {
    const el = document.getElementById(id);
    if (!el) return;
    charts[id] = new Chart(el, config);
  }

  Chart.defaults.font.family = "'Inter', system-ui, sans-serif";
  Chart.defaults.font.size = 12;
  Chart.defaults.color = '#646985';
  Chart.defaults.plugins.legend.labels.usePointStyle = true;
  Chart.defaults.plugins.legend.labels.boxWidth = 8;
  Chart.defaults.plugins.tooltip.backgroundColor = '#1D2041';
  Chart.defaults.plugins.tooltip.padding = 10;
  Chart.defaults.plugins.tooltip.cornerRadius = 8;

  /* ---------------- Resumen general ---------------- */
  function renderResumen(d) {
    const ord = [...d.mtto.items, ...d.oper.items];
    const conOrden = ord.filter(x => x.hasOrder).length;
    const alta = ord.filter(x => x.prioridad === 'Alta' && ABIERTO(x.estado)).length;
    const montoTot = d.rq.items.reduce((s, x) => s + x.monto, 0);
    const antTot = d.rq.items.reduce((s, x) => s + x.anticipo, 0);
    const abiertas = ord.filter(x => ABIERTO(x.estado)).length;

    const segCard = (key, tag, big, bigLabel, s1, s1l, s2, s2l, progress) => `
      <a class="card seg" href="#${key}">
        <div class="seg__top"><span class="seg__tag">${tag}</span><span class="seg__icon"><svg viewBox="0 0 24 24">${ICONS[key]}</svg></span></div>
        <div class="seg__big">${big}<small>${bigLabel}</small></div>
        ${progress !== undefined ? `<div class="bar" title="${progress}% con orden emitida"><span style="width:${progress}%"></span></div>` : ''}
        <div class="seg__stats"><div><b>${s1}</b>${s1l}</div><div><b>${s2}</b>${s2l}</div></div>
        <span class="seg__link">Ver segmento →</span>
      </a>`;

    const m = d.mtto.items, o = d.oper.items;
    const segs = [
      segCard('mtto', 'Mantenimiento', fmtInt(m.length), 'solicitudes',
        fmtInt(m.filter(x => x.prioridad === 'Alta').length), 'prioridad alta',
        fmtInt(m.filter(x => !x.hasOrder).length), 'sin orden emitida', pct(m.filter(x => x.hasOrder).length, m.length)),
      segCard('oper', 'Operaciones', fmtInt(o.length), 'solicitudes',
        fmtInt(o.filter(x => x.estado === 'Por aprobación').length), 'por aprobación',
        fmtInt(o.filter(x => x.estado === 'Entrega pendiente').length), 'entrega pendiente', pct(o.filter(x => x.hasOrder).length, o.length)),
      segCard('mat', 'Materiales críticos', fmtInt(d.mat.items.length), 'materiales',
        d.mat.ix && d.mat.ix.stock >= 0 ? fmtInt(d.mat.items.filter(x => x.bajoMinimo).length) : '—', 'bajo mínimo',
        fmtInt(d.mat.headers.length), 'campos de control'),
      segCard('rq', 'Anticipos y montos', fmtMoneyShort(montoTot), 'comprometido',
        fmtMoneyShort(antTot), 'anticipos', fmtInt(d.rq.items.filter(x => x.incidencia).length), 'con incidencia')
    ].join('');

    // Puntos de atención
    const alerts = [];
    const altaSin = ord.filter(x => x.prioridad === 'Alta' && !x.hasOrder);
    if (altaSin.length) alerts.push({ c: C.risk, n: altaSin.length, t: 'solicitudes de prioridad alta aún sin orden emitida', s: 'Mantenimiento / Operaciones', go: ['mtto', { estado: 'En gestión (sin orden)', prioridad: 'Alta' }] });
    const porAprob = ord.filter(x => x.estado === 'Por aprobación');
    if (porAprob.length) alerts.push({ c: C.yellow, n: porAprob.length, t: 'órdenes en espera de aprobación', s: `${m.filter(x => x.estado === 'Por aprobación').length} en Mantenimiento · ${o.filter(x => x.estado === 'Por aprobación').length} en Operaciones`, go: [m.some(x => x.estado === 'Por aprobación') ? 'mtto' : 'oper', { estado: 'Por aprobación' }] });
    const pagoPend = ord.filter(x => x.estado === 'Pago / anticipo pendiente');
    if (pagoPend.length) alerts.push({ c: C.gold, n: pagoPend.length, t: 'órdenes con pago o anticipo pendiente', s: 'Requieren gestión de Finanzas', go: [m.some(x => x.estado === 'Pago / anticipo pendiente') ? 'mtto' : 'oper', { estado: 'Pago / anticipo pendiente' }] });
    const viejas = ord.filter(x => x.fEm && x.dias > CFG.agingAlertDays && ABIERTO(x.estado));
    if (viejas.length) alerts.push({ c: C.navy3, n: viejas.length, t: `órdenes abiertas con más de ${CFG.agingAlertDays} días desde su emisión`, s: 'Revisar avance con proveedores', go: [viejas.filter(x => m.includes(x)).length >= viejas.length / 2 ? 'mtto' : 'oper', { aging: true }] });
    const inc = d.rq.items.filter(x => x.incidencia);
    if (inc.length) alerts.push({ c: C.risk, n: inc.length, t: inc.length === 1 ? 'anticipo con incidencia reportada' : 'anticipos con incidencia reportada', s: inc.map(x => x.incidencia).slice(0, 2).join(' · '), go: ['rq', {}] });
    if (!d.mat.items.length) alerts.push({ c: '#C9CCDD', n: '—', t: 'Materiales críticos sin registros cargados', s: 'La hoja existe pero aún no tiene información', go: ['mat', {}] });

    const est = ESTADOS.map(e => e.key);
    $('#view-resumen').innerHTML = `
      <div class="section-head">
        <div><h2>Resumen general</h2><p>Visión consolidada de los cuatro segmentos. Haz clic en cualquier tarjeta para ver su detalle.</p></div>
      </div>
      <div class="grid kpis">
        ${kpi('Solicitudes registradas', fmtInt(ord.length), `${fmtInt(abiertas)} abiertas · Mantenimiento + Operaciones`)}
        ${kpi('Órdenes emitidas', pct(conOrden, ord.length) + '%', `${fmtInt(conOrden)} de ${fmtInt(ord.length)} con número de orden`, 'accent')}
        ${kpi('Prioridad alta abiertas', fmtInt(alta), 'Requieren seguimiento cercano', 'risk')}
        ${kpi('Monto comprometido', fmtMoneyShort(montoTot), `${fmtMoneyShort(antTot)} en anticipos (${pct(antTot, montoTot)}%)`, 'ok')}
      </div>
      <div class="grid cols-4">${segs}</div>
      <div class="grid cols-3">
        <div class="card">
          <h3>Estado de las órdenes por segmento</h3>
          <p class="sub">Estado estimado a partir de los comentarios registrados en el Excel</p>
          <div class="chart-box"><canvas id="ch-res-estado"></canvas></div>
        </div>
        <div class="card">
          <h3>Puntos de atención</h3>
          <p class="sub">Generados automáticamente</p>
          <ul class="alerts">${alerts.map((a, i) => `<li data-alert="${i}"><span class="dot" style="background:${a.c}"></span><div><b>${a.n}</b> ${esc(a.t)}<small>${esc(a.s)}</small></div></li>`).join('') || '<li>Sin alertas</li>'}</ul>
        </div>
      </div>
      <div class="grid cols-2">
        <div class="card"><h3>Distribución por región</h3><p class="sub">Según el código de la orden o la descripción</p><div class="chart-box"><canvas id="ch-res-region"></canvas></div></div>
        <div class="card"><h3>Prioridad de las solicitudes</h3><p class="sub">Mantenimiento y Operaciones</p><div class="chart-box"><canvas id="ch-res-prio"></canvas></div></div>
      </div>`;

    $('#view-resumen').querySelectorAll('[data-alert]').forEach(li => li.onclick = () => {
      const a = alerts[+li.dataset.alert]; const [seg, f] = a.go;
      tableState[seg] = Object.assign(defaultState(), f);
      location.hash = seg;
      setTimeout(() => { refreshTable(seg); openDetail(seg); }, 30);
    });

    const segLabels = ['Mantenimiento', 'Operaciones'];
    const present = est.filter(k => ord.some(x => x.estado === k));
    makeChart('ch-res-estado', {
      type: 'bar',
      data: { labels: segLabels, datasets: present.map(k => ({ label: k, data: [m.filter(x => x.estado === k).length, o.filter(x => x.estado === k).length], backgroundColor: estadoColor(k), borderRadius: 4, borderSkipped: false })) },
      options: { indexAxis: 'y', maintainAspectRatio: false, scales: { x: { stacked: true, grid: { color: C.line }, ticks: { precision: 0 } }, y: { stacked: true, grid: { display: false } } }, plugins: { legend: { position: 'bottom' } },
        onClick: (e, els) => { if (!els.length) return; const el = els[0]; const seg = el.index === 0 ? 'mtto' : 'oper'; tableState[seg] = Object.assign(defaultState(), { estado: present[el.datasetIndex] }); location.hash = seg; setTimeout(() => { refreshTable(seg); openDetail(seg); }, 30); } }
    });

    const reg = countBy([...ord, ...d.rq.items.filter(x => !ord.some(o2 => o2.orden && o2.orden === x.orden))].filter(x => x.orden || x.solicitud), x => x.region);
    const regKeys = Object.keys(reg).sort((a, b) => reg[b] - reg[a]);
    makeChart('ch-res-region', {
      type: 'doughnut',
      data: { labels: regKeys, datasets: [{ data: regKeys.map(k => reg[k]), backgroundColor: regKeys.map(k => REGION_COLORS[k] || '#ccc'), borderColor: '#fff', borderWidth: 2 }] },
      options: { maintainAspectRatio: false, cutout: '62%', plugins: { legend: { position: 'right' } } }
    });

    const prioKeys = ['Alta', 'Media', 'Baja', 'Sin definir'];
    makeChart('ch-res-prio', {
      type: 'bar',
      data: { labels: segLabels, datasets: prioKeys.map((k, i) => ({ label: k, data: [m.filter(x => x.prioridad === k).length, o.filter(x => x.prioridad === k).length], backgroundColor: ['#D9534F', '#FDD649', '#2E9E6B', '#E2E4EE'][i], borderRadius: 4 })) },
      options: { maintainAspectRatio: false, scales: { y: { grid: { color: C.line }, ticks: { precision: 0 } }, x: { grid: { display: false } } }, plugins: { legend: { position: 'bottom' } } }
    });
  }

  /* ---------------- Segmentos de órdenes (MTTO / OPER) ---------------- */
  const defaultState = () => ({ search: '', estado: '', prioridad: '', owner: '', tipo: '', sort: null, dir: 1, aging: false });

  function renderOrders(key, label, seg, desc) {
    const it = seg.items;
    tableState[key] = tableState[key] || defaultState();
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

    $('#view-' + key).innerHTML = `
      <div class="section-head"><div><h2>${label}</h2><p>${desc}</p></div></div>
      <div class="grid kpis">
        ${kpi('Solicitudes', fmtInt(it.length), `${fmtInt(abiertas.length)} abiertas`)}
        ${kpi('Con orden emitida', fmtInt(conOrden.length), `${pct(conOrden.length, it.length)}% del total · ${it.filter(x => x.tipo === 'ODS').length} ODS / ${it.filter(x => x.tipo === 'ODC').length} ODC`, 'accent')}
        ${kpi('En gestión (sin orden)', fmtInt(it.length - conOrden.length), 'Cotización, validación o elaboración')}
        ${hasPrio ? kpi('Prioridad alta', fmtInt(alta.length), `${pct(alta.length, it.length)}% de las solicitudes`, 'risk')
                  : kpi('Por aprobación', fmtInt(it.filter(x => x.estado === 'Por aprobación').length), 'En espera de firma', 'risk')}
        ${kpi('Antigüedad promedio', edadProm !== null ? edadProm + ' días' : '—', edadProm !== null ? 'Órdenes abiertas, desde su emisión' : 'Sin fechas de emisión registradas', 'ok')}
      </div>
      <div class="grid cols-2">
        <div class="card"><h3>Estado de las solicitudes</h3><p class="sub">Haz clic en una barra para filtrar el detalle</p><div class="chart-box"><canvas id="ch-${key}-estado"></canvas></div></div>
        <div class="card"><h3>Carga por responsable de Procura</h3><p class="sub">Solicitudes asignadas por estado</p><div class="chart-box"><canvas id="ch-${key}-proc"></canvas></div></div>
      </div>
      <div class="grid cols-2">
        ${hasOwner ? `<div class="card"><h3>${esc(seg.ownerLabel)}</h3><p class="sub">Solicitudes por responsable técnico</p><div class="chart-box"><canvas id="ch-${key}-owner"></canvas></div></div>` : ''}
        <div class="card"><h3>Tipo de orden y región</h3><p class="sub">ODS = servicio · ODC = compra</p><div class="chart-box"><canvas id="ch-${key}-tipo"></canvas></div></div>
        ${!hasOwner ? `<div class="card"><h3>Principales proveedores</h3><p class="sub">Número de órdenes por empresa</p><div class="chart-box"><canvas id="ch-${key}-emp"></canvas></div></div>` : ''}
      </div>
      ${detailBlock(key, it.length, `
        <div class="filters">
          <input type="search" placeholder="Buscar por orden, solicitud, empresa o comentario…" data-f="search">
          <select data-f="estado"><option value="">Todos los estados</option>${ESTADOS.filter(e => it.some(x => x.estado === e.key)).map(e => `<option>${e.key}</option>`).join('')}</select>
          ${hasPrio ? `<select data-f="prioridad"><option value="">Todas las prioridades</option>${['Alta', 'Media', 'Baja', 'Sin definir'].filter(p => it.some(x => x.prioridad === p)).map(p => `<option>${p}</option>`).join('')}</select>` : ''}
          <select data-f="tipo"><option value="">Todos los tipos</option>${[...new Set(it.map(x => x.tipo))].sort().map(t => `<option>${esc(t)}</option>`).join('')}</select>
          <select data-f="owner"><option value="">Todos los responsables</option>${[...new Set(it.flatMap(x => [x.owner, x.procura]))].filter(x => x !== 'Sin asignar').sort().map(t => `<option>${esc(t)}</option>`).join('')}</select>
          <button class="btn btn--light" data-reset>Limpiar</button>
          <span class="count" id="count-${key}"></span>
        </div>
        <div class="table-wrap"><table id="tbl-${key}"></table></div>
        <p class="note">El <b>estado</b> se estima automáticamente según palabras clave del campo “Comentarios” (p. ej. “POR APROB”, “ANTICIPO”, “ENTREGA PENDIENTE”). Para el dato exacto consulta el comentario original o descarga el Excel.</p>
      `)}`;

    bindDetail(key);
    refreshTable(key);

    // Gráficos
    const est = ESTADOS.filter(e => it.some(x => x.estado === e.key));
    makeChart(`ch-${key}-estado`, {
      type: 'bar',
      data: { labels: est.map(e => e.key), datasets: [{ data: est.map(e => it.filter(x => x.estado === e.key).length), backgroundColor: est.map(e => e.color), borderRadius: 6, borderSkipped: false }] },
      options: { indexAxis: 'y', maintainAspectRatio: false, plugins: { legend: { display: false } }, scales: { x: { grid: { color: C.line }, ticks: { precision: 0 } }, y: { grid: { display: false } } },
        onClick: (e, els) => { if (!els.length) return; tableState[key] = Object.assign(defaultState(), { estado: est[els[0].index].key }); refreshTable(key); openDetail(key); },
        onHover: (e, els) => { e.native.target.style.cursor = els.length ? 'pointer' : 'default'; } }
    });

    const stackedBy = (field, canvas, onPick) => {
      const people = [...new Set(it.map(x => x[field]))].sort((a, b) => it.filter(x => x[field] === b).length - it.filter(x => x[field] === a).length);
      makeChart(canvas, {
        type: 'bar',
        data: { labels: people, datasets: est.map(e => ({ label: e.key, data: people.map(p => it.filter(x => x[field] === p && x.estado === e.key).length), backgroundColor: e.color, borderRadius: 3 })) },
        options: { maintainAspectRatio: false, scales: { x: { stacked: true, grid: { display: false } }, y: { stacked: true, grid: { color: C.line }, ticks: { precision: 0 } } }, plugins: { legend: { position: 'bottom', labels: { font: { size: 11 } } } },
          onClick: (e, els) => { if (!els.length) return; tableState[key] = Object.assign(defaultState(), { owner: people[els[0].index] === 'Sin asignar' ? '' : people[els[0].index] }); refreshTable(key); openDetail(key); } }
      });
    };
    stackedBy('procura', `ch-${key}-proc`);
    if (hasOwner) stackedBy('owner', `ch-${key}-owner`);

    const tipos = [...new Set(it.map(x => x.tipo))].sort();
    const regs = Object.keys(REGION_COLORS).filter(r => it.some(x => x.region === r));
    makeChart(`ch-${key}-tipo`, {
      type: 'bar',
      data: { labels: tipos, datasets: regs.map(r => ({ label: r, data: tipos.map(t => it.filter(x => x.tipo === t && x.region === r).length), backgroundColor: REGION_COLORS[r], borderRadius: 3 })) },
      options: { maintainAspectRatio: false, scales: { x: { stacked: true, grid: { display: false } }, y: { stacked: true, grid: { color: C.line }, ticks: { precision: 0 } } }, plugins: { legend: { position: 'bottom' } } }
    });

    if (!hasOwner) {
      const emp = countBy(it.filter(x => x.empresa), x => x.empresa);
      const top = Object.keys(emp).sort((a, b) => emp[b] - emp[a]).slice(0, 7);
      makeChart(`ch-${key}-emp`, {
        type: 'bar',
        data: { labels: top.map(t => t.length > 28 ? t.slice(0, 27) + '…' : t), datasets: [{ data: top.map(t => emp[t]), backgroundColor: C.navy, borderRadius: 6 }] },
        options: { indexAxis: 'y', maintainAspectRatio: false, plugins: { legend: { display: false }, tooltip: { callbacks: { title: i => top[i[0].dataIndex] } } }, scales: { x: { grid: { color: C.line }, ticks: { precision: 0 } }, y: { grid: { display: false } } } }
      });
    }
  }

  function bindDetail(key) {
    const root = $('#detail-' + key);
    root.querySelector('[data-detail]').onclick = () => root.classList.toggle('is-open');
    root.querySelectorAll('[data-f]').forEach(el => {
      el.addEventListener(el.tagName === 'INPUT' ? 'input' : 'change', () => {
        tableState[key][el.dataset.f] = el.value; tableState[key].aging = false; refreshTable(key);
      });
    });
    const reset = root.querySelector('[data-reset]');
    if (reset) reset.onclick = () => { tableState[key] = defaultState(); refreshTable(key); };
  }

  const prioBadge = p => `<span class="badge prio-${{ Alta: 'alta', Media: 'media', Baja: 'baja' }[p] || 'none'}">${esc(p)}</span>`;
  const estadoBadge = e => { const col = estadoColor(e); const dark = ['#FDD649', '#C9CCDD', '#E6E8F0', '#9AA0BF'].includes(col); return `<span class="badge" style="background:${col};color:${dark ? '#1D2041' : '#fff'}">${esc(e)}</span>`; };

  function refreshTable(key) {
    if (key === 'rq') return refreshPayTable();
    if (key === 'mat') return;
    const seg = DATA[key]; if (!seg || !seg.items.length) return;
    const st = tableState[key] || (tableState[key] = defaultState());
    const root = $('#detail-' + key);
    root.querySelectorAll('[data-f]').forEach(el => { el.value = st[el.dataset.f] || ''; });

    const q = norm(st.search);
    let rows = seg.items.filter(x =>
      (!q || norm([x.orden, x.solicitud, x.empresa, x.coment, x.owner, x.procura].join(' ')).includes(q)) &&
      (!st.estado || x.estado === st.estado) &&
      (!st.prioridad || x.prioridad === st.prioridad) &&
      (!st.tipo || x.tipo === st.tipo) &&
      (!st.owner || x.owner === st.owner || x.procura === st.owner) &&
      (!st.aging || (x.fEm && x.dias > CFG.agingAlertDays && ABIERTO(x.estado))));

    const cols = [
      { k: 'orden', t: 'No. de orden', f: x => x.orden ? `<span class="mono">${esc(x.orden)}</span>` : '<span class="muted">Sin orden</span>', cls: 'nowrap' },
      { k: 'solicitud', t: 'Solicitud', f: x => esc(x.solicitud), cls: 'wide' },
      { k: 'empresa', t: 'Empresa', f: x => esc(x.empresa) || '<span class="muted">—</span>', cls: 'mid' },
      { k: 'tipo', t: 'Tipo', f: x => esc(x.tipo), cls: 'nowrap' },
      { k: 'prioridad', t: 'Prioridad', f: x => prioBadge(x.prioridad), sortVal: x => ({ Alta: 0, Media: 1, Baja: 2 }[x.prioridad] ?? 3) },
      { k: 'estado', t: 'Estado (estimado)', f: x => estadoBadge(x.estado) },
      { k: 'fEm', t: 'Emisión', f: x => fmtDate(x.fEm) || (x.fSol ? `<span class="muted">Sol. ${fmtDate(x.fSol)}</span>` : '<span class="muted">—</span>'), cls: 'nowrap', sortVal: x => (x.fEm || x.fSol || { getTime: () => 0 }).getTime() },
      { k: 'dias', t: 'Días', f: x => x.dias ?? '<span class="muted">—</span>', cls: 'num', sortVal: x => x.dias ?? -1 },
      { k: 'owner', t: seg.ownerLabel.replace(/^Encargado\s*/i, 'Enc. '), f: x => esc(x.owner), cls: 'nowrap' },
      { k: 'procura', t: 'Enc. Procura', f: x => esc(x.procura), cls: 'nowrap' },
      { k: 'coment', t: 'Comentarios', f: x => esc(x.coment) || '<span class="muted">—</span>', cls: 'wide' }
    ].filter(c => c.k !== 'prioridad' || seg.items.some(x => x.prioridad !== 'Sin definir'))
     .filter(c => c.k !== 'owner' || seg.items.some(x => x.owner !== 'Sin asignar'));

    if (st.sort) {
      const c = cols.find(c => c.k === st.sort);
      const val = c && c.sortVal ? c.sortVal : (x => norm(x[st.sort]));
      rows = rows.slice().sort((a, b) => (val(a) > val(b) ? 1 : val(a) < val(b) ? -1 : 0) * st.dir);
    }

    $('#tbl-' + key).innerHTML = `<thead><tr>${cols.map(c => `<th data-sort="${c.k}">${esc(c.t)}<span class="sort">${st.sort === c.k ? (st.dir > 0 ? '▲' : '▼') : '↕'}</span></th>`).join('')}</tr></thead>
      <tbody>${rows.map(x => `<tr>${cols.map(c => `<td class="${c.cls || ''}">${c.f(x)}</td>`).join('')}</tr>`).join('') || `<tr><td colspan="${cols.length}" class="muted" style="text-align:center;padding:24px">Sin resultados con los filtros aplicados</td></tr>`}</tbody>`;
    $('#count-' + key).textContent = `${rows.length} de ${seg.items.length} registros${st.aging ? ` · solo órdenes > ${CFG.agingAlertDays} días` : ''}`;
    $('#tbl-' + key).querySelectorAll('th').forEach(th => th.onclick = () => {
      st.dir = st.sort === th.dataset.sort ? -st.dir : 1; st.sort = th.dataset.sort; refreshTable(key);
    });
  }

  /* ---------------- Materiales críticos ---------------- */
  function renderMaterials(seg) {
    const v = $('#view-mat');
    const head = `<div class="section-head"><div><h2>Materiales críticos</h2><p>Inventario de materiales críticos con sus niveles mínimos y máximos de reposición.</p></div></div>`;
    if (!seg.items.length) {
      v.innerHTML = head + `
        <div class="grid kpis">
          ${kpi('Materiales registrados', '0', 'Aún no hay información cargada')}
          ${kpi('Bajo nivel mínimo', '—', 'Se calcula al registrar existencias', 'risk')}
        </div>
        <div class="empty">
          <h3>Esta sección aún no tiene registros</h3>
          <p>Cuando se completen filas en la hoja <b>${esc(DATA.sheetNames.mat || 'Materiales Criticos')}</b> del Excel, aparecerán aquí automáticamente.</p>
          <div class="cols">${seg.headers.map(h => `<span>${esc(h)}</span>`).join('')}</div>
          <p class="note">Sugerencia: agregar una columna <b>“Existencia actual”</b> permitirá alertar los materiales por debajo del mínimo.</p>
        </div>`;
      return;
    }
    const bajo = seg.items.filter(x => x.bajoMinimo).length;
    v.innerHTML = head + `
      <div class="grid kpis">
        ${kpi('Materiales registrados', fmtInt(seg.items.length), 'Materiales bajo control')}
        ${kpi('Bajo nivel mínimo', seg.ix.stock >= 0 ? fmtInt(bajo) : '—', seg.ix.stock >= 0 ? `${pct(bajo, seg.items.length)}% del total` : 'Agrega “Existencia actual” para calcularlo', 'risk')}
      </div>
      ${detailBlock('mat', seg.items.length, `
        <div class="table-wrap"><table>
          <thead><tr>${seg.headers.map(h => `<th>${esc(h)}</th>`).join('')}</tr></thead>
          <tbody>${seg.items.map(x => `<tr${x.bajoMinimo ? ' style="outline:2px solid #F6C7C5"' : ''}>${x.cells.map(c => `<td class="${typeof c === 'number' ? 'num' : ''}">${esc(c)}</td>`).join('')}</tr>`).join('')}</tbody>
        </table></div>`)}`;
    $('#detail-mat [data-detail]').onclick = () => $('#detail-mat').classList.toggle('is-open');
  }

  /* ---------------- Anticipos y montos (RQ) ---------------- */
  function renderPayments(seg) {
    const v = $('#view-rq');
    const it = seg.items;
    const head = `<div class="section-head"><div><h2>Anticipos y montos</h2><p>Montos de las órdenes con requerimiento de anticipo y porcentaje acordado con el proveedor.</p></div></div>`;
    if (!it.length) { v.innerHTML = head + '<div class="empty"><h3>Sin registros</h3></div>'; return; }
    const tot = it.reduce((s, x) => s + x.monto, 0);
    const ant = it.reduce((s, x) => s + x.anticipo, 0);
    const inc = it.filter(x => x.incidencia);
    tableState.rq = tableState.rq || { search: '', sort: 'monto', dir: -1 };

    v.innerHTML = head + `
      <div class="grid kpis">
        ${kpi('Monto total', fmtMoney(tot), `${fmtInt(it.length)} órdenes`)}
        ${kpi('Anticipos', fmtMoney(ant), `${pct(ant, tot)}% del monto total`, 'accent')}
        ${kpi('Saldo por pagar', fmtMoney(tot - ant), 'Pendiente contra entrega / valuación', 'ok')}
        ${kpi('Con incidencia', fmtInt(inc.length), inc.length ? esc(inc.map(x => x.incidencia).join(' · ')) : 'Sin incidencias', 'risk')}
      </div>
      <div class="grid cols-2">
        <div class="card"><h3>Anticipo vs. saldo por orden</h3><p class="sub">Montos en ${CFG.currency}</p><div class="chart-box tall"><canvas id="ch-rq-orden"></canvas></div></div>
        <div class="card"><h3>Monto por proveedor</h3><p class="sub">Participación en el monto comprometido</p><div class="chart-box tall"><canvas id="ch-rq-emp"></canvas></div></div>
      </div>
      ${detailBlock('rq', it.length, `
        <div class="filters">
          <input type="search" placeholder="Buscar por orden, solicitud o empresa…" data-f="search">
          <span class="count" id="count-rq"></span>
        </div>
        <div class="table-wrap"><table id="tbl-rq"></table></div>`)}`;

    const root = $('#detail-rq');
    root.querySelector('[data-detail]').onclick = () => root.classList.toggle('is-open');
    root.querySelector('[data-f]').oninput = e => { tableState.rq.search = e.target.value; refreshPayTable(); };
    refreshPayTable();

    const sorted = it.slice().sort((a, b) => b.monto - a.monto);
    makeChart('ch-rq-orden', {
      type: 'bar',
      data: {
        labels: sorted.map(x => x.orden || x.solicitud.slice(0, 20)),
        datasets: [
          { label: 'Anticipo', data: sorted.map(x => x.anticipo), backgroundColor: C.yellow, borderRadius: 3 },
          { label: 'Saldo', data: sorted.map(x => x.saldo), backgroundColor: sorted.map(x => x.incidencia ? '#F1B3B1' : C.navy), borderRadius: 3 }
        ]
      },
      options: { indexAxis: 'y', maintainAspectRatio: false, scales: { x: { stacked: true, grid: { color: C.line }, ticks: { callback: v => fmtMoneyShort(v) } }, y: { stacked: true, grid: { display: false }, ticks: { font: { family: 'Consolas, monospace', size: 11 } } } },
        plugins: { legend: { position: 'bottom' }, tooltip: { callbacks: { label: c => `${c.dataset.label}: ${fmtMoney(c.raw)}`, afterBody: i => { const x = sorted[i[0].dataIndex]; return [x.empresa, x.incidencia ? '⚠ ' + x.incidencia : `Anticipo ${Math.round((x.pct || 0) * 100)}%`]; } } } } }
    });

    const emp = {};
    it.forEach(x => { emp[x.empresa || 'Sin empresa'] = (emp[x.empresa || 'Sin empresa'] || 0) + x.monto; });
    const ek = Object.keys(emp).sort((a, b) => emp[b] - emp[a]);
    const pal = ['#1D2041', '#FDD649', '#3C4180', '#E3A92B', '#7F86B5', '#9AA0BF', '#C9CCDD'];
    makeChart('ch-rq-emp', {
      type: 'doughnut',
      data: { labels: ek, datasets: [{ data: ek.map(k => emp[k]), backgroundColor: ek.map((_, i) => pal[i % pal.length]), borderColor: '#fff', borderWidth: 2 }] },
      options: { maintainAspectRatio: false, cutout: '58%', plugins: { legend: { position: 'bottom', labels: { font: { size: 11 } } }, tooltip: { callbacks: { label: c => `${c.label}: ${fmtMoney(c.raw)} (${pct(c.raw, tot)}%)` } } } }
    });
  }

  function refreshPayTable() {
    const it = DATA.rq.items; const st = tableState.rq;
    const q = norm(st.search);
    let rows = it.filter(x => !q || norm([x.orden, x.solicitud, x.empresa, x.incidencia].join(' ')).includes(q));
    const cols = [
      { k: 'orden', t: 'No. de orden', f: x => `<span class="mono">${esc(x.orden)}</span>`, cls: 'nowrap' },
      { k: 'solicitud', t: 'Solicitud', f: x => esc(x.solicitud), cls: 'wide' },
      { k: 'empresa', t: 'Empresa', f: x => esc(x.empresa), cls: 'mid' },
      { k: 'monto', t: 'Monto', f: x => fmtMoney(x.monto), cls: 'num' },
      { k: 'pct', t: '% Anticipo', f: x => x.pct !== null ? Math.round(x.pct * 100) + '%' : `<span class="badge prio-alta">${esc(x.incidencia)}</span>`, cls: 'num', sortVal: x => x.pct ?? -1 },
      { k: 'anticipo', t: 'Anticipo', f: x => x.pct !== null ? fmtMoney(x.anticipo) : '<span class="muted">—</span>', cls: 'num' },
      { k: 'saldo', t: 'Saldo', f: x => fmtMoney(x.saldo), cls: 'num' }
    ];
    if (st.sort) {
      const c = cols.find(c => c.k === st.sort);
      const val = c && c.sortVal ? c.sortVal : (x => typeof x[st.sort] === 'number' ? x[st.sort] : norm(x[st.sort]));
      rows = rows.slice().sort((a, b) => (val(a) > val(b) ? 1 : val(a) < val(b) ? -1 : 0) * st.dir);
    }
    const tot = rows.reduce((s, x) => s + x.monto, 0), ant = rows.reduce((s, x) => s + x.anticipo, 0);
    $('#tbl-rq').innerHTML = `<thead><tr>${cols.map(c => `<th data-sort="${c.k}">${c.t}<span class="sort">${st.sort === c.k ? (st.dir > 0 ? '▲' : '▼') : '↕'}</span></th>`).join('')}</tr></thead>
      <tbody>${rows.map(x => `<tr>${cols.map(c => `<td class="${c.cls}">${c.f(x)}</td>`).join('')}</tr>`).join('')}
      <tr><td colspan="3" style="font-weight:700;background:#F3F4F8">Total</td><td class="num" style="font-weight:700;background:#F3F4F8">${fmtMoney(tot)}</td><td style="background:#F3F4F8"></td><td class="num" style="font-weight:700;background:#F3F4F8">${fmtMoney(ant)}</td><td class="num" style="font-weight:700;background:#F3F4F8">${fmtMoney(tot - ant)}</td></tr></tbody>`;
    $('#count-rq').textContent = `${rows.length} de ${it.length} registros`;
    $('#tbl-rq').querySelectorAll('th').forEach(th => th.onclick = () => { st.dir = st.sort === th.dataset.sort ? -st.dir : 1; st.sort = th.dataset.sort; refreshPayTable(); });
  }

  /* ---------------- Navegación ---------------- */
  function route() {
    const tab = (location.hash || '#resumen').slice(1);
    const valid = ['resumen', 'mtto', 'oper', 'mat', 'rq'].includes(tab) ? tab : 'resumen';
    document.querySelectorAll('.tab').forEach(a => a.classList.toggle('is-active', a.dataset.tab === valid));
    document.querySelectorAll('.view').forEach(s => s.classList.toggle('is-active', s.id === 'view-' + valid));
    Object.values(charts).forEach(c => c.resize());
    window.scrollTo({ top: 0, behavior: 'instant' in window ? 'instant' : 'auto' });
  }

  window.addEventListener('hashchange', route);
  $('#fileInput').addEventListener('change', e => { if (e.target.files[0]) loadLocalFile(e.target.files[0]); e.target.value = ''; });
  $('#btnPrint').addEventListener('click', () => {
    const v = document.querySelector('.view.is-active');
    v && v.querySelectorAll('.detail').forEach(d => d.classList.add('is-open'));
    window.print();
  });
  window.__loadFile = f => f && loadLocalFile(f);

  // estado inicial de carga
  document.querySelectorAll('.view').forEach(v => v.innerHTML = '<div class="loading"><div><div class="spinner"></div>Cargando datos del Excel…</div></div>');
  loadDefault();
})();
