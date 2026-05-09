/* Finanzas JM & Pili — dashboard */

const FMT = new Intl.NumberFormat('es-UY', { style: 'currency', currency: 'UYU', maximumFractionDigits: 0 });
const FMT_CMP = new Intl.NumberFormat('es-UY', { maximumFractionDigits: 0 });
const PALETTE = ['#059669','#0ea5e9','#f59e0b','#dc2626','#7c3aed','#0284c7','#ea580c','#be123c','#15803d','#1d4ed8'];

// Íconos por categoría (emojis simples, paleta del theme se aplica al fondo)
const ICONS = {
  // Gastos fijos
  'Supermercado':         '🛒',
  'Servicios':            '⚡',
  'Cuotas':               '💳',
  'Gastos comunes':       '📋',
  'Niñera':               '👶',
  'Jardín':               '🌿',
  'Seguro de vida':       '🛡️',
  'Mapfre':               '🛡️',
  'CJPPU':                '💼',
  // Variables esenciales
  'Salud':                '💊',
  'Estación de servicio': '⛽',
  'Comida trabajo':       '🍴',
  // Variables discrecionales
  'Social / Amigos':      '👥',
  'Regalos':              '🎁',
  'Viajes':               '✈️',
  'Hogar - Mejoras':      '🏠',
  'Ropa':                 '👕',
  'Cosmética':            '✨',
  'Deportes / Gym':       '🏋️',
  'Delivery / Pedidos':   '🍔',
  'Entretenimiento':      '🎬',
  'Varios':               '🔹',
  'Ahorros':              '🏦',
  // Ingresos
  'Sueldo JM':            '💰',
  'Sueldo Pili':          '💵',
  'Otros ingresos':       '💎',
  // Sistema
  'No va':                '⏸️',
  'Traspaso':             '↔️',
  'Sin clasificar':       '❓',
};
const iconFor = cat => ICONS[cat] || '🔹';

let DATA = null;
let mesActualIdx = 0;
let charts = {};

// ── carga ─────────────────────────────────────────────────────
async function load() {
  const res = await fetch('data/finanzas.json?v=' + Date.now());
  // 401 = sesión Cloudflare Access expirada → recargar para re-autenticar
  if (res.status === 401) {
    location.reload();
    throw new Error('Sesión expirada, recargando…');
  }
  if (!res.ok) throw new Error('No se pudo cargar el JSON');
  const json = await res.json();
  if (json && json.error) throw new Error(json.error === 'offline' ? 'Sin conexión' : 'No se pudo cargar');
  DATA = json;
  mesActualIdx = DATA.meses.length - 1;
  init();
  document.getElementById('loader').classList.add('hidden');
}

function init() {
  // header
  const fecha = new Date(DATA.generado);
  document.getElementById('generado').textContent = 'Actualizado ' + fecha.toLocaleDateString('es-UY', { day:'numeric', month:'short', hour:'2-digit', minute:'2-digit' });

  // selector de mes
  const sel = document.getElementById('mesSelect');
  sel.innerHTML = DATA.meses.map((m, i) => `<option value="${i}">${m.label}</option>`).join('');
  sel.value = mesActualIdx;
  sel.onchange = e => { mesActualIdx = +e.target.value; renderResumen(); };
  document.getElementById('prevMes').onclick = () => { if (mesActualIdx > 0) { mesActualIdx--; sel.value = mesActualIdx; renderResumen(); } };
  document.getElementById('nextMes').onclick = () => { if (mesActualIdx < DATA.meses.length - 1) { mesActualIdx++; sel.value = mesActualIdx; renderResumen(); } };

  // tabs
  document.querySelectorAll('.tab').forEach(btn => {
    btn.onclick = () => switchTab(btn.dataset.tab);
  });

  // recargar
  document.getElementById('btnReload').onclick = () => location.reload();

  renderResumen();
  renderHistorico();
  renderCuadro();
}

// Devuelve el índice en cuadro.ca.meses para el mes seleccionado en DATA.meses
function cuadroIdxFor(mesId) {
  return DATA.cuadro.ca.meses.indexOf(mesId);
}

function renderSaldoActual(mesIdx) {
  const ca       = DATA.cuadro.ca;
  const mesData  = DATA.meses[mesIdx];
  const cuadroIdx = cuadroIdxFor(mesData.id);
  const esUltimo  = mesIdx === DATA.meses.length - 1;

  let saldo = null, label = '';
  if (cuadroIdx >= 0 && ca.saldo_final[cuadroIdx] != null) {
    saldo = ca.saldo_final[cuadroIdx];
    label = esUltimo ? 'Saldo Caja de Ahorro' : `Saldo al cierre de ${mesData.label}`;
  } else {
    // fallback al último saldo conocido
    for (let i = ca.saldo_final.length - 1; i >= 0; i--) {
      if (ca.saldo_final[i] != null) { saldo = ca.saldo_final[i]; break; }
    }
    label = 'Saldo Caja de Ahorro';
  }

  document.getElementById('saldoLabel').textContent = label;
  document.getElementById('saldoActual').textContent = saldo != null ? FMT.format(saldo) : '—';

  // Meta: si es el mes actual, mostrar fecha; si es histórico, delta vs mes anterior
  const metaEl = document.getElementById('saldoActualMeta');
  if (esUltimo) {
    metaEl.textContent = 'al día de hoy';
  } else if (cuadroIdx > 0 && ca.saldo_final[cuadroIdx - 1] != null && saldo != null) {
    const delta = saldo - ca.saldo_final[cuadroIdx - 1];
    const sign = delta >= 0 ? '+' : '−';
    const cls  = delta >= 0 ? 'pos' : 'neg';
    metaEl.innerHTML = `<span class="${cls}">${sign}${FMT.format(Math.abs(delta))}</span> vs ${ca.labels[cuadroIdx - 1]}`;
  } else {
    metaEl.textContent = '';
  }
}

function switchTab(name) {
  document.querySelectorAll('.tab').forEach(t => t.classList.toggle('active', t.dataset.tab === name));
  document.querySelectorAll('.view').forEach(v => v.classList.toggle('hidden', v.dataset.view !== name));
}

// ── tab RESUMEN ─────────────────────────────────────────────────
function renderResumen() {
  const mes = DATA.meses[mesActualIdx];

  // Saldo (contextual al mes seleccionado)
  renderSaldoActual(mesActualIdx);

  // Mini chart (últimos 12 meses, mes seleccionado destacado, solo gastos)
  const last12 = DATA.meses.slice(-12);
  drawBar(
    'chartMini',
    last12.map(m => m.label.split(' ')[0]),
    last12.map(m => m.gastos_total),
    mesActualIdx - (DATA.meses.length - last12.length)
  );

  // Categorías + sparklines
  renderCats(mes, mesActualIdx);

  // Top movs
  renderTopMovs(mes);
}

// Helper: serie de últimos N meses para una categoría dada
function serieCategoria(catName, beforeIdx, n = 6) {
  const start = Math.max(0, beforeIdx - n + 1);
  const out = [];
  for (let i = start; i <= beforeIdx; i++) {
    out.push(DATA.meses[i]?.categorias?.[catName]?.total || 0);
  }
  return out;
}

// Mini SVG sparkline para inline (sin Chart.js, super liviano)
function sparklineSVG(values, w = 60, h = 18) {
  if (!values.length) return '';
  const max = Math.max(...values, 1);
  const min = 0;
  const xStep = w / Math.max(values.length - 1, 1);
  const pts = values.map((v, i) => {
    const x = i * xStep;
    const y = h - ((v - min) / (max - min || 1)) * (h - 2) - 1;
    return `${x.toFixed(1)},${y.toFixed(1)}`;
  }).join(' ');
  return `<svg class="spark" width="${w}" height="${h}" viewBox="0 0 ${w} ${h}" preserveAspectRatio="none">
    <polyline points="${pts}" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/>
  </svg>`;
}

function renderCats(mes, mesIdx) {
  const list = document.getElementById('catList');
  const cats = Object.entries(mes.categorias)
    .map(([nombre, d]) => ({ nombre, ...d, prom: DATA.promedios[nombre] || 0 }))
    .sort((a, b) => b.total - a.total);

  list.innerHTML = cats.map(c => {
    const promRatio = c.prom > 0 ? c.total / c.prom : 0;
    let cls = 'ok';
    if (promRatio > 1.2) cls = 'bad';
    else if (promRatio > 1) cls = 'warn';
    const fillW = Math.min(promRatio, 1.5) / 1.5 * 100;
    const markerL = c.prom > 0 ? (1 / 1.5) * 100 : 0;

    const subTxt = Object.keys(c.sub || {}).length
      ? Object.entries(c.sub).sort((a,b)=>b[1]-a[1]).map(([k,v]) => `${escapeHtml(k)}: ${FMT.format(v)}`).join(' · ')
      : '';

    const movsHtml = (c.movs || []).slice(0, 30).map(m => `
      <div class="cat-mov">
        <div class="m-info">${m.f.slice(8,10)}/${m.f.slice(5,7)} — ${escapeHtml(m.d2 || m.c || '')}</div>
        <div class="m-amt ${m.m > 0 ? 'cred' : ''}">${FMT.format(Math.abs(m.m))}</div>
      </div>
    `).join('');

    // Sparkline + tendencia
    const serie = serieCategoria(c.nombre, mesIdx, 6);
    const spark = sparklineSVG(serie, 56, 16);
    const tendencia = (() => {
      if (serie.length < 3) return '';
      const recent = serie.slice(-3).reduce((s,v)=>s+v,0) / 3;
      const older  = serie.slice(0, -3).reduce((s,v)=>s+v,0) / Math.max(serie.length - 3, 1);
      if (older === 0) return '';
      const diff = (recent - older) / older;
      if (diff > 0.15) return '<span class="trend up" title="creciendo">↗</span>';
      if (diff < -0.15) return '<span class="trend down" title="bajando">↘</span>';
      return '<span class="trend flat" title="estable">→</span>';
    })();

    return `
      <div class="cat-row expandable" data-cat="${escapeHtml(c.nombre)}">
        <div class="cat-row-head">
          <span class="cat-icon">${iconFor(c.nombre)}</span>
          <div class="cat-name-wrap">
            <div class="cat-name">${escapeHtml(c.nombre)} ${tendencia}</div>
          </div>
          <div class="cat-spark">${spark}</div>
          <div class="cat-amount">${FMT.format(c.total)}</div>
        </div>
        <div class="cat-bar">
          <div class="cat-bar-fill ${cls}" style="width:${fillW}%"></div>
          ${c.prom > 0 ? `<div class="cat-bar-marker" style="left:${markerL}%"></div>` : ''}
        </div>
        <div class="cat-meta">
          ${c.prom > 0 ? `prom. ${FMT.format(c.prom)}` : 'sin promedio histórico'}
          ${(c.movs || []).length ? ` · ${c.movs.length} mov.` : ''}
        </div>
        ${subTxt ? `<div class="cat-sub">${subTxt}</div>` : ''}
        <div class="cat-detail">${movsHtml}</div>
      </div>
    `;
  }).join('');

  // expand on click
  list.querySelectorAll('.cat-row.expandable').forEach(row => {
    row.onclick = () => row.classList.toggle('open');
  });
}

// Bar chart con línea horizontal de promedio + barra del mes seleccionado destacada
function drawBarConPromedio(canvasId, labels, data, highlightIdx, promedio) {
  if (charts[canvasId]) charts[canvasId].destroy();
  const ctx = document.getElementById(canvasId);
  const colors = data.map((v, i) => {
    if (i === highlightIdx) return '#059669';
    return v > promedio ? 'rgba(220,38,38,.35)' : 'rgba(5,150,105,.30)';
  });
  charts[canvasId] = new Chart(ctx, {
    type: 'bar',
    data: { labels, datasets: [{ data, backgroundColor: colors, borderRadius: 6 }] },
    options: {
      responsive: true, maintainAspectRatio: false,
      animation: { duration: 350 },
      plugins: {
        legend: { display: false },
        tooltip: { callbacks: { label: c => FMT.format(c.raw) } },
        annotation: undefined
      },
      scales: {
        x: { grid: { display: false }, ticks: { font: { size: 11 } } },
        y: { grid: { color: 'rgba(148,163,184,.15)' }, ticks: { font: { size: 11 }, callback: v => '$' + FMT_CMP.format(v) } }
      }
    },
    plugins: [{
      id: 'avg-line',
      afterDatasetsDraw(chart) {
        if (!promedio) return;
        const { ctx: c, scales: { y } } = chart;
        const yPos = y.getPixelForValue(promedio);
        c.save();
        c.beginPath();
        c.setLineDash([4, 4]);
        c.strokeStyle = 'rgba(100,116,139,.55)';
        c.lineWidth = 1.2;
        c.moveTo(chart.chartArea.left, yPos);
        c.lineTo(chart.chartArea.right, yPos);
        c.stroke();
        c.setLineDash([]);
        c.fillStyle = 'rgba(100,116,139,.85)';
        c.font = '600 10px Inter, system-ui';
        c.textAlign = 'right';
        c.fillText('prom.', chart.chartArea.right - 4, yPos - 4);
        c.restore();
      }
    }]
  });
}

function renderTopMovs(mes) {
  const cont = document.getElementById('topMovs');
  if (!mes.top_movs || !mes.top_movs.length) {
    cont.innerHTML = '<div style="color:var(--muted); padding:.5rem 0; font-size:.85rem;">Sin movimientos grandes en el mes.</div>';
    return;
  }
  cont.innerHTML = mes.top_movs.slice(0, 10).map(m => `
    <div class="mov-row">
      <div class="mov-info">
        <div class="mov-desc">${escapeHtml(m.d2 || m.c || m.desc)}</div>
        <div class="mov-meta">${m.f.slice(8,10)}/${m.f.slice(5,7)} · ${escapeHtml(m.desc)}</div>
      </div>
      <div class="mov-amt">${FMT.format(Math.abs(m.m))}</div>
    </div>
  `).join('');
}

// ── tab HISTÓRICO ─────────────────────────────────────────────────
function renderHistorico() {
  const h = DATA.historico;

  drawLine('chartHist', h.labels, [
    { label: 'Gastos',   data: h.gastos,   borderColor: '#dc2626', bg: 'rgba(220,38,38,.10)' },
    { label: 'Ingresos', data: h.ingresos, borderColor: '#059669', bg: 'rgba(5,150,105,.10)' },
  ]);

  const cats = Object.keys(h.por_cat).slice(0, 6);
  drawLine('chartCats', h.labels, cats.map((cat, i) => ({
    label: cat, data: h.por_cat[cat], borderColor: PALETTE[i % PALETTE.length], bg: 'transparent'
  })));

  const legend = document.getElementById('catsLegend');
  legend.innerHTML = cats.map((c, i) =>
    `<div class="legend-item"><span class="legend-dot" style="background:${PALETTE[i % PALETTE.length]}"></span>${escapeHtml(c)}</div>`
  ).join('');

  // stats
  const totalGastos = DATA.meses.reduce((s, m) => s + m.gastos_total, 0);
  const totalIng    = DATA.meses.reduce((s, m) => s + m.ingresos_total, 0);
  const promG       = totalGastos / DATA.meses.length;
  const promI       = totalIng    / DATA.meses.length;
  const maxMes      = DATA.meses.reduce((a, b) => a.gastos_total > b.gastos_total ? a : b);
  const minMes      = DATA.meses.reduce((a, b) => a.gastos_total < b.gastos_total ? a : b);

  document.getElementById('statsGrid').innerHTML = `
    <div class="stat"><div class="stat-label">Promedio gastos</div><div class="stat-value">${FMT.format(promG)}</div></div>
    <div class="stat"><div class="stat-label">Promedio ingresos</div><div class="stat-value">${FMT.format(promI)}</div></div>
    <div class="stat"><div class="stat-label">Mes más caro</div><div class="stat-value">${FMT.format(maxMes.gastos_total)}<br><span style="font-size:.7rem;color:var(--muted)">${maxMes.label}</span></div></div>
    <div class="stat"><div class="stat-label">Mes más barato</div><div class="stat-value">${FMT.format(minMes.gastos_total)}<br><span style="font-size:.7rem;color:var(--muted)">${minMes.label}</span></div></div>
    <div class="stat"><div class="stat-label">Meses registrados</div><div class="stat-value">${DATA.meses.length}</div></div>
  `;
}

// ── tab CUADRO ─────────────────────────────────────────────────
let cuadroViewMode = 'tabla';   // default siempre Tabla; el usuario puede pasar a Lista

function renderCuadro() {
  const selCuenta = document.getElementById('cuadroCuenta');
  const selAno    = document.getElementById('cuadroAno');

  // poblar selector de año
  const anos = new Set();
  ['ca','cc'].forEach(k => DATA.cuadro[k].meses.forEach(m => anos.add(m.split('-')[0])));
  const anosList = [...anos].sort();
  selAno.innerHTML = '<option value="all">Todos</option>' +
    anosList.map(a => `<option value="${a}">${a}</option>`).join('');
  selAno.value = anosList[anosList.length - 1] || 'all';

  // Toggle vista Tabla/Lista
  document.querySelectorAll('.vt-btn').forEach(btn => {
    btn.classList.toggle('active', btn.dataset.viewMode === cuadroViewMode);
    btn.onclick = () => {
      cuadroViewMode = btn.dataset.viewMode;
      document.querySelectorAll('.vt-btn').forEach(b => b.classList.toggle('active', b === btn));
      draw();
    };
  });

  const draw = () => {
    const wrapTabla = document.getElementById('cuadroTablaWrap');
    const wrapLista = document.getElementById('cuadroLista');
    if (cuadroViewMode === 'lista') {
      wrapTabla.hidden = true;
      wrapLista.hidden = false;
      drawCuadroLista(selCuenta.value, selAno.value);
    } else {
      wrapTabla.hidden = false;
      wrapLista.hidden = true;
      drawCuadro(selCuenta.value, selAno.value);
    }
  };

  selCuenta.onchange = draw;
  selAno.onchange    = draw;
  draw();
}

// Vista Lista: una tarjeta por categoría con sparkline + total + tap para detalle
function drawCuadroLista(cuenta, ano) {
  const c = DATA.cuadro[cuenta];
  const idxs = c.meses.map((m, i) => ({m, i})).filter(o => ano === 'all' || o.m.startsWith(ano));
  const colIdx = idxs.map(o => o.i);
  const labels = idxs.map(o => c.labels[o.i]);
  const slice = arr => colIdx.map(i => arr[i]);

  const sumArr = arr => arr.reduce((s, v) => s + (v || 0), 0);
  const fmtCell = v => {
    if (v == null || v === 0) return '<span class="amt-zero">—</span>';
    return FMT_CMP.format(Math.round(Math.abs(v)));
  };

  const renderItem = (entry, kind) => {
    const vals = slice(entry.valores);
    if (!vals.some(v => v && v !== 0)) return '';
    const total = sumArr(vals);
    const ico = iconFor(entry.cat);
    const spark = sparklineSVG(vals, 70, 22);
    const cls = kind === 'ing' ? 'amt-pos' : 'amt-neg';

    // Detalle expandible: subcats + tabla mes a mes
    const subs = Object.entries(entry.subcats || {})
      .filter(([_, v]) => slice(v).some(x => x && x !== 0))
      .map(([sub, v]) => {
        const tot = sumArr(slice(v));
        return `<div class="li-sub"><span>${escapeHtml(sub)}</span><span>${FMT.format(Math.abs(tot))}</span></div>`;
      }).join('');

    const detalleMeses = labels.map((l, j) => {
      const v = vals[j];
      if (!v) return '';
      return `<div class="li-mes"><span>${l}</span><span class="${cls}">${FMT.format(Math.abs(v))}</span></div>`;
    }).join('');

    return `
      <div class="li-row" data-kind="${kind}">
        <div class="li-head">
          <span class="li-icon">${ico}</span>
          <div class="li-title">
            <div class="li-name">${escapeHtml(entry.cat)}</div>
            <div class="li-sub-text">${labels.length} ${labels.length === 1 ? 'mes' : 'meses'}</div>
          </div>
          <div class="li-spark ${cls}">${spark}</div>
          <div class="li-total ${cls}">${FMT.format(Math.abs(total))}</div>
        </div>
        <div class="li-detail">
          ${subs ? `<div class="li-block"><div class="li-block-title">Subcategorías</div>${subs}</div>` : ''}
          <div class="li-block"><div class="li-block-title">Mes a mes</div>${detalleMeses}</div>
        </div>
      </div>
    `;
  };

  // Saldos resumen
  const saldoIni = c.saldo_inicial[colIdx[0]];
  const saldoFin = c.saldo_final[colIdx[colIdx.length - 1]];
  const totalIng = sumArr(slice(c.ingresos_total));
  const totalGas = sumArr(slice(c.gastos_total));

  let html = `
    <div class="li-summary">
      <div class="li-sum-row"><span>Saldo inicial</span><span>${saldoIni != null ? FMT.format(saldoIni) : '—'}</span></div>
      <div class="li-sum-row"><span>Total ingresos</span><span class="amt-pos">+${FMT.format(totalIng)}</span></div>
      <div class="li-sum-row"><span>Total gastos</span><span class="amt-neg">−${FMT.format(totalGas)}</span></div>
      <div class="li-sum-row strong"><span>Saldo final</span><span>${saldoFin != null ? FMT.format(saldoFin) : '—'}</span></div>
    </div>
    <div class="li-section-title">Ingresos</div>
  `;

  c.ingresos.forEach(ing => html += renderItem(ing, 'ing'));
  html += `<div class="li-section-title">Gastos</div>`;
  c.gastos.forEach(g => html += renderItem(g, 'gas'));

  document.getElementById('cuadroLista').innerHTML = html;

  // expand on tap
  document.querySelectorAll('.li-row').forEach(row => {
    row.onclick = () => row.classList.toggle('open');
  });
}

function drawCuadro(cuenta, ano) {
  const c = DATA.cuadro[cuenta];
  // filtrar columnas por año
  const idxs = c.meses.map((m, i) => ({m, i})).filter(o => ano === 'all' || o.m.startsWith(ano));
  const cols = idxs.map(o => c.labels[o.i]);
  const colIdx = idxs.map(o => o.i);

  const fmt = v => {
    if (v == null) return '<span class="amt-zero">—</span>';
    const n = Math.round(v);
    if (n === 0) return '<span class="amt-zero">—</span>';
    const s = FMT_CMP.format(Math.abs(n));
    if (n < 0) return `<span class="amt-neg">(${s})</span>`;
    return `<span>${s}</span>`;
  };

  const slice = arr => colIdx.map(i => arr[i]);
  const hasData = arr => slice(arr).some(v => v != null && v !== 0);

  // Construir tabla
  const rows = [];
  // header
  rows.push(`<thead><tr>
    <th class="col-label">Concepto</th>
    ${cols.map(l => `<th>${l}</th>`).join('')}
    <th>Total</th>
  </tr></thead>`);

  const cellsAndTotal = (arr, cls = '') => {
    const vals = slice(arr);
    const total = vals.reduce((s, v) => s + (v || 0), 0);
    return vals.map(v => `<td class="${cls}">${fmt(v)}</td>`).join('') +
           `<td class="${cls}"><b>${fmt(total)}</b></td>`;
  };

  let body = '<tbody>';

  // Saldo Inicial
  body += `<tr class="row-saldo">
    <td class="col-label">Saldo Inicial</td>
    ${cellsAndTotal(c.saldo_inicial)}
  </tr>`;

  // INGRESOS
  body += `<tr class="row-section">
    <td class="col-label">Ingresos</td>
    ${cellsAndTotal(c.ingresos_total, 'amt-pos')}
  </tr>`;
  c.ingresos.forEach((ing, i) => {
    if (!hasData(ing.valores)) return;     // ocultar si no tiene nada en el período visible
    body += `<tr class="row-cat" data-target="ing-${i}">
      <td class="col-label">${escapeHtml(ing.cat)}</td>
      ${cellsAndTotal(ing.valores, 'amt-pos')}
    </tr>`;
    if (ing.subcats) {
      Object.entries(ing.subcats).forEach(([sub, vals]) => {
        if (!hasData(vals)) return;
        body += `<tr class="row-sub hidden" data-parent="ing-${i}">
          <td class="col-label">${escapeHtml(sub)}</td>
          ${cellsAndTotal(vals, 'amt-pos')}
        </tr>`;
      });
    }
  });

  // GASTOS
  body += `<tr class="row-section">
    <td class="col-label">Gastos</td>
    ${cellsAndTotal(c.gastos_total, 'amt-neg')}
  </tr>`;
  c.gastos.forEach((g, i) => {
    if (!hasData(g.valores)) return;       // ocultar categorías vacías
    body += `<tr class="row-cat" data-target="gas-${i}">
      <td class="col-label">${escapeHtml(g.cat)}</td>
      ${cellsAndTotal(g.valores, 'amt-neg')}
    </tr>`;
    Object.entries(g.subcats || {}).forEach(([sub, vals]) => {
      if (!hasData(vals)) return;
      body += `<tr class="row-sub hidden" data-parent="gas-${i}">
        <td class="col-label">${escapeHtml(sub)}</td>
        ${cellsAndTotal(vals, 'amt-neg')}
      </tr>`;
    });
  });

  // Sin clasificar (si hay)
  if (c.sin_clasificar && c.sin_clasificar.valores.some(v => v !== 0)) {
    body += `<tr class="row-cat" style="background:#fef3c7">
      <td class="col-label">⚠ Sin clasificar</td>
      ${cellsAndTotal(c.sin_clasificar.valores)}
    </tr>`;
  }

  // Saldo Final
  body += `<tr class="row-saldo">
    <td class="col-label">Saldo Final</td>
    ${cellsAndTotal(c.saldo_final)}
  </tr>`;

  body += '</tbody>';

  document.getElementById('cuadroTable').innerHTML = rows.join('') + body;

  // toggle subcategorías
  document.querySelectorAll('#cuadroTable .row-cat').forEach(row => {
    const target = row.dataset.target;
    if (!target) return;
    const subs = document.querySelectorAll(`#cuadroTable .row-sub[data-parent="${target}"]`);
    if (!subs.length) return;
    row.style.cursor = 'pointer';
    row.onclick = () => {
      row.classList.toggle('open');
      subs.forEach(s => s.classList.toggle('hidden'));
    };
  });
}

// (tab CC removido — la cuenta corriente se ve en el cuadro)
function _renderCC_unused() {
  const meses = DATA.cc.meses;
  if (!meses.length) {
    document.getElementById('ccSaldo').textContent = '—';
    return;
  }
  const ult = meses[meses.length - 1];
  document.getElementById('ccSaldo').textContent = ult.saldo != null ? FMT.format(ult.saldo) : '—';
  const totalIngresos = meses.reduce((s, m) => s + (m.ingresos || 0), 0);
  const totalEgresos  = meses.reduce((s, m) => s + (m.egresos  || 0), 0);
  document.getElementById('ccMeta').innerHTML =
    `Total ingresos acumulados: <b>${FMT.format(totalIngresos)}</b> · egresos: <b>${FMT.format(totalEgresos)}</b>`;

  drawLine('chartCC', meses.map(m => m.label.split(' ')[0]), [
    { label: 'Ingresos', data: meses.map(m => m.ingresos), borderColor: '#059669', bg: 'rgba(5,150,105,.10)' },
    { label: 'Egresos',  data: meses.map(m => m.egresos),  borderColor: '#dc2626', bg: 'rgba(220,38,38,.10)' },
  ]);

  document.getElementById('ccTabla').innerHTML = `
    <div class="cc-head"><div>Mes</div><div>Ingresos</div><div>Egresos</div><div>Saldo</div></div>
    ${meses.slice().reverse().map(m => `
      <div class="cc-row">
        <div>${m.label}</div>
        <div class="ing">${FMT.format(m.ingresos)}</div>
        <div class="egr">${FMT.format(m.egresos)}</div>
        <div>${m.saldo != null ? FMT.format(m.saldo) : '—'}</div>
      </div>
    `).join('')}
  `;
}

// ── chart helpers ─────────────────────────────────────────────────
function drawBar(canvasId, labels, data, highlightIdx = -1) {
  if (charts[canvasId]) charts[canvasId].destroy();
  const ctx = document.getElementById(canvasId);
  const colors = data.map((_, i) => i === highlightIdx ? '#059669' : 'rgba(5,150,105,.30)');
  charts[canvasId] = new Chart(ctx, {
    type: 'bar',
    data: { labels, datasets: [{ data, backgroundColor: colors, borderRadius: 6 }] },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { display: false }, tooltip: { callbacks: { label: c => FMT.format(c.raw) } } },
      scales: {
        x: { grid: { display: false }, ticks: { font: { size: 10 } } },
        y: { grid: { color: 'rgba(148,163,184,.15)' }, ticks: { font: { size: 10 }, callback: v => '$' + FMT_CMP.format(v) } }
      }
    }
  });
}

function drawLine(canvasId, labels, datasets) {
  if (charts[canvasId]) charts[canvasId].destroy();
  const ctx = document.getElementById(canvasId);
  charts[canvasId] = new Chart(ctx, {
    type: 'line',
    data: {
      labels,
      datasets: datasets.map(d => ({
        label: d.label, data: d.data,
        borderColor: d.borderColor, backgroundColor: d.bg || 'transparent',
        borderWidth: 2, tension: .3, pointRadius: 2, pointHoverRadius: 4, fill: !!d.bg && d.bg !== 'transparent'
      }))
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { display: datasets.length > 1, labels: { font: { size: 11 } } },
                 tooltip: { callbacks: { label: c => `${c.dataset.label}: ${FMT.format(c.raw)}` } } },
      scales: {
        x: { grid: { display: false }, ticks: { font: { size: 10 }, maxRotation: 0 } },
        y: { grid: { color: 'rgba(148,163,184,.15)' }, ticks: { font: { size: 10 }, callback: v => '$' + FMT_CMP.format(v) } }
      }
    }
  });
}

// ── utils ─────────────────────────────────────────────────
function mediaUlt(arr, n, key, beforeIdx) {
  const start = Math.max(0, beforeIdx - n);
  const slice = arr.slice(start, beforeIdx);
  if (!slice.length) return 0;
  return slice.reduce((s, m) => s + (m[key] || 0), 0) / slice.length;
}

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
}

// ── go ─────────────────────────────────────────────────
load().catch(err => {
  document.getElementById('loader').innerHTML = `
    <div style="padding:2rem; text-align:center; max-width:340px;">
      <div style="font-size:3rem; margin-bottom:1rem;">📡</div>
      <div style="color:var(--text); font-weight:600; font-size:1.05rem; margin-bottom:.5rem;">
        No se pudo cargar
      </div>
      <div style="color:var(--muted); font-size:.85rem; margin-bottom:1.25rem;">
        Probablemente sin conexión o se cerró la sesión.
      </div>
      <button onclick="location.reload()" style="
        background: var(--accent); color: #fff; border: 0;
        padding: .75rem 1.5rem; border-radius: 10px;
        font-size: .95rem; font-weight: 600; cursor: pointer;
        font-family: inherit;
      ">Reintentar</button>
      <div style="margin-top:1rem; font-size:.7rem; color:var(--muted);">
        ${err.message}
      </div>
    </div>`;
});

// service worker
if ('serviceWorker' in navigator) {
  window.addEventListener('load', () => {
    navigator.serviceWorker.register('sw.js').then(reg => {
      // Chequear updates cada vez que la app vuelve al foreground
      document.addEventListener('visibilitychange', () => {
        if (document.visibilityState === 'visible') reg.update().catch(()=>{});
      });

      // Cuando se detecta SW nuevo instalado, mostrar banner para actualizar
      reg.addEventListener('updatefound', () => {
        const nw = reg.installing;
        if (!nw) return;
        nw.addEventListener('statechange', () => {
          if (nw.state === 'installed' && navigator.serviceWorker.controller) {
            mostrarBannerUpdate(nw);
          }
        });
      });
    }).catch(()=>{});

    // Cuando el nuevo SW toma el control, recargar la página
    let recargando = false;
    navigator.serviceWorker.addEventListener('controllerchange', () => {
      if (recargando) return;
      recargando = true;
      location.reload();
    });
  });
}

function mostrarBannerUpdate(nw) {
  if (document.getElementById('updateBanner')) return;
  const b = document.createElement('div');
  b.id = 'updateBanner';
  b.style.cssText = `
    position: fixed; left: 50%; bottom: 20px; transform: translateX(-50%);
    background: var(--accent); color: #fff; padding: .75rem 1.25rem;
    border-radius: 999px; box-shadow: 0 4px 14px rgba(0,0,0,.25);
    font-weight: 600; font-size: .9rem; cursor: pointer; z-index: 200;
    display: flex; align-items: center; gap: .5rem;
  `;
  b.innerHTML = '✨ Hay una versión nueva — tocá para actualizar';
  b.onclick = () => {
    b.textContent = 'Actualizando…';
    nw.postMessage({ type: 'SKIP_WAITING' });
    // Por si el SW no responde al mensaje, recargamos en 2s
    setTimeout(() => location.reload(), 2000);
  };
  document.body.appendChild(b);
}
