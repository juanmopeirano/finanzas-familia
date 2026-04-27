/* Finanzas JM & Pili — dashboard */

const FMT = new Intl.NumberFormat('es-UY', { style: 'currency', currency: 'UYU', maximumFractionDigits: 0 });
const FMT_CMP = new Intl.NumberFormat('es-UY', { maximumFractionDigits: 0 });
const PALETTE = ['#059669','#0ea5e9','#f59e0b','#dc2626','#7c3aed','#0284c7','#ea580c','#be123c','#15803d','#1d4ed8'];

let DATA = null;
let mesActualIdx = 0;
let charts = {};

// ── carga ─────────────────────────────────────────────────────
async function load() {
  const res = await fetch('data/finanzas.json?v=' + Date.now());
  if (!res.ok) throw new Error('No se pudo cargar el JSON');
  DATA = await res.json();
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

  renderSaldoActual();
  renderResumen();
  renderHistorico();
  renderCuadro();
}

function renderSaldoActual() {
  const ca = DATA.cuadro.ca;
  // último saldo final no nulo
  let saldo = null, mesIdx = -1;
  for (let i = ca.saldo_final.length - 1; i >= 0; i--) {
    if (ca.saldo_final[i] != null) { saldo = ca.saldo_final[i]; mesIdx = i; break; }
  }
  document.getElementById('saldoActual').textContent = saldo != null ? FMT.format(saldo) : '—';
  if (mesIdx >= 0) {
    document.getElementById('saldoActualMeta').textContent = `al cierre de ${ca.labels[mesIdx]}`;
  }
}

function switchTab(name) {
  document.querySelectorAll('.tab').forEach(t => t.classList.toggle('active', t.dataset.tab === name));
  document.querySelectorAll('.view').forEach(v => v.classList.toggle('hidden', v.dataset.view !== name));
}

// ── tab RESUMEN ─────────────────────────────────────────────────
function renderResumen() {
  const mes = DATA.meses[mesActualIdx];
  const promTotal = mediaUlt(DATA.meses, 12, 'gastos_total', mesActualIdx);

  // Hero
  document.getElementById('heroAmount').textContent = FMT.format(mes.gastos_total);

  const delta = mes.gastos_total - promTotal;
  const pct = promTotal > 0 ? Math.round((delta / promTotal) * 100) : 0;
  let deltaClass = 'delta-good', deltaTxt = `${pct}% bajo el promedio`;
  if (pct > 0)  { deltaClass = pct > 10 ? 'delta-bad' : 'delta-warn'; deltaTxt = `${pct}% sobre el promedio`; }
  if (pct === 0) deltaTxt = 'igual al promedio';
  document.getElementById('heroMeta').innerHTML =
    `Promedio últ. 12 meses: <b>${FMT.format(promTotal)}</b> · <span class="${deltaClass}">${deltaTxt}</span>`;

  const ratio = promTotal > 0 ? Math.min(mes.gastos_total / promTotal, 1.5) : 0;
  const fill = document.getElementById('heroBarFill');
  fill.style.width = (Math.min(ratio, 1) * 100) + '%';
  fill.classList.remove('warn','bad');
  if (ratio > 1.1) fill.classList.add('bad');
  else if (ratio > 1)  fill.classList.add('warn');

  // Mini chart
  const last12 = DATA.meses.slice(-12);
  drawBar('chartMini', last12.map(m => m.label.split(' ')[0]), last12.map(m => m.gastos_total), mesActualIdx - (DATA.meses.length - last12.length));

  // Categorías
  renderCats(mes);

  // Top movs
  renderTopMovs(mes);
}

function renderCats(mes) {
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
      ? Object.entries(c.sub).sort((a,b)=>b[1]-a[1]).map(([k,v]) => `${k}: ${FMT.format(v)}`).join(' · ')
      : '';

    const movsHtml = (c.movs || []).slice(0, 30).map(m => `
      <div class="cat-mov">
        <div class="m-info">${m.f.slice(8,10)}/${m.f.slice(5,7)} — ${m.d2 || m.c || ''}</div>
        <div class="m-amt ${m.m > 0 ? 'cred' : ''}">${FMT.format(Math.abs(m.m))}</div>
      </div>
    `).join('');

    return `
      <div class="cat-row expandable" data-cat="${escapeHtml(c.nombre)}">
        <div class="cat-row-head">
          <div class="cat-name">${escapeHtml(c.nombre)}</div>
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
        ${subTxt ? `<div class="cat-sub">${escapeHtml(subTxt)}</div>` : ''}
        <div class="cat-detail">${movsHtml}</div>
      </div>
    `;
  }).join('');

  // expand on click
  list.querySelectorAll('.cat-row.expandable').forEach(row => {
    row.onclick = () => row.classList.toggle('open');
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
    <div class="stat"><div class="stat-label">Total acumulado</div><div class="stat-value">${FMT.format(totalGastos)}</div></div>
    <div class="stat"><div class="stat-label">Meses registrados</div><div class="stat-value">${DATA.meses.length}</div></div>
  `;
}

// ── tab CUADRO ─────────────────────────────────────────────────
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

  const draw = () => drawCuadro(selCuenta.value, selAno.value);
  selCuenta.onchange = draw;
  selAno.onchange    = draw;
  draw();
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
  document.getElementById('loader').innerHTML = `<div style="padding:2rem; color:var(--danger); text-align:center;">Error: ${err.message}<br><small>¿Existe data/finanzas.json?</small></div>`;
});

// service worker
if ('serviceWorker' in navigator) {
  window.addEventListener('load', () => navigator.serviceWorker.register('sw.js').catch(()=>{}));
}
