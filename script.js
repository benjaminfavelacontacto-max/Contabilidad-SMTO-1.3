/* ══════════════════════════════════════════════════════
   FinDash — script.js
   Soporta dos modos de lectura:
     A) Multi-hoja: hojas "INGRESOS" y "EGRESOS" separadas
        (formato SMTO: columnas específicas por hoja)
     B) Hoja única: columnas Fecha, Tipo, Categoría, Monto
   El año se detecta automáticamente desde los datos.
══════════════════════════════════════════════════════ */

'use strict';

// ─── ESTADO GLOBAL ────────────────────────────────────
let allRows        = [];  // Todos los registros (todos los años)
let yearRows       = [];  // Registros del año seleccionado
let filteredRows   = [];  // Registros filtrados (año + mes)
let allYears       = [];  // Lista de años disponibles
let barChartInst   = null;
let donutChartInst = null;
let tipoBarInst    = null;
let detectedYear   = new Date().getFullYear();

// ─── ESTADO TABLA DE TRANSACCIONES ────────────────────
let txCurrentRows  = [];    // filas actuales antes de filtros de columna
let txColFilters   = {};    // colKey → Set<string> | vacío = sin filtro
let txOpenCol      = null;  // columna cuyo dropdown está abierto
let txSortCol      = null;  // columna activa de ordenamiento
let txSortDir      = null;  // 'asc' | 'desc' | null

// ─── CONTROLADORES TABLAS DESGLOSE ─────────────────────
// Patrón: cada tabla tiene su propio estado encapsulado.
function makeCatCtrl(tableId, bodyId, badgeId, label) {
  return { allRows:[], colFilters:{}, openCol:null, tableId, bodyId, badgeId, label };
}
// Se instancian globalmente; se recargan en cada renderCategoryTable()
const catIngCtrl = makeCatCtrl('catTableIng', 'catBodyIng', 'catBadgeIng', 'Ingresos');
const catEgrCtrl = makeCatCtrl('catTableEgr', 'catBodyEgr', 'catBadgeEgr', 'Egresos');
let   catActiveCtrl = null;  // controller con dropdown abierto

// ─── PALETA ──────────────────────────────────────────
const PALETTE = [
  '#6366f1','#8b5cf6','#ec4899','#f43f5e',
  '#f97316','#f59e0b','#10b981','#14b8a6',
  '#06b6d4','#3b82f6','#a3e635','#84cc16',
  '#e879f9','#fb7185','#fbbf24','#34d399',
];

const MONTHS_ES   = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'];
const MONTHS_LONG = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];

// ══════════════════════════════════════════════════════
// 1. DRAG & DROP / FILE INPUT
// ══════════════════════════════════════════════════════

function triggerFileInput() { document.getElementById('fileInput').click(); }

function handleDragOver(e) {
  e.preventDefault();
  document.getElementById('dropZone').classList.add('drag-over');
}
function handleDragLeave(e) {
  e.preventDefault();
  document.getElementById('dropZone').classList.remove('drag-over');
}
function handleDrop(e) {
  e.preventDefault();
  document.getElementById('dropZone').classList.remove('drag-over');
  const file = e.dataTransfer.files[0];
  if (file) processFile(file);
}
function handleFileChange(e) {
  const file = e.target.files[0];
  if (file) processFile(file);
}

// ══════════════════════════════════════════════════════
// 2. PROCESAMIENTO DEL ARCHIVO
// ══════════════════════════════════════════════════════

function processFile(file) {
  const ext = file.name.split('.').pop().toLowerCase();
  if (!['xlsx','xls'].includes(ext)) {
    showError('El archivo debe ser .xlsx o .xls. Por favor verifica el formato.');
    return;
  }
  hideError();
  showLoading();

  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const data = new Uint8Array(e.target.result);
      // Sin cellDates:true — trabajamos con strings via raw:false
      const wb = XLSX.read(data, { type: 'array' });

      let rows = [];

      // ── MODO A: hojas INGRESOS + EGRESOS (formato SMTO) ──
      const hasIngresos = wb.SheetNames.some(n => n.toUpperCase().includes('INGRESO'));
      const hasEgresos  = wb.SheetNames.some(n => n.toUpperCase().includes('EGRESO'));

      if (hasIngresos && hasEgresos) {
        rows = parseMultiSheet(wb);
      } else {
        // ── MODO B: hoja única genérica ──
        rows = parseSingleSheet(wb.Sheets[wb.SheetNames[0]]);
      }

      if (rows.length === 0) {
        throw new Error(
          `No se encontraron transacciones válidas en el archivo. ` +
          `Hojas: [${wb.SheetNames.join(', ')}]. ` +
          'Verifica que tenga hojas INGRESOS y EGRESOS con columnas de Fecha, Tipo y Total.'
        );
      }

      buildDashboard(rows);
    } catch (err) {
      hideLoading();
      showError(err.message || 'Error al leer el archivo. Verifica que sea un Excel válido.');
    }
  };
  reader.onerror = () => { hideLoading(); showError('No se pudo leer el archivo. Intenta de nuevo.'); };
  reader.readAsArrayBuffer(file);
}

// ══════════════════════════════════════════════════════
// 3A. PARSER MULTI-HOJA
//     Estrategia: sheet_to_json con raw:false + dateNF.
//     Todo llega como string → sin problemas de tipos.
// ══════════════════════════════════════════════════════

/** Lee un worksheet como array de arrays de strings (sin tipos). */
function wsToStrings(ws) {
  return XLSX.utils.sheet_to_json(ws, {
    header:  1,
    raw:     false,       // TODO como string formateado
    dateNF:  'yyyy-mm-dd', // Fechas → "2017-01-04"
    defval:  '',          // Celdas vacías → ""
  });
}

/** Elige hoja por nombre: exacta primero, luego la más corta con la palabra. */
function findSheetName(wb, keyword) {
  const kw = keyword.toUpperCase();
  const exact = wb.SheetNames.find(n => {
    const u = n.trim().toUpperCase();
    return u === kw || u === kw + 'S';
  });
  if (exact) return exact;
  const partials = wb.SheetNames.filter(n => n.trim().toUpperCase().includes(kw));
  return partials.length ? partials.sort((a,b)=>a.length-b.length)[0] : null;
}

/**
 * Busca la fila de encabezados y mapea nombres → índice de columna.
 * @returns {{ headerIdx, colMap }} o null
 */
function detectHeader(rows, dateHints, maxScan) {
  maxScan = Math.min(maxScan || 25, rows.length);
  for (let i = 0; i < maxScan; i++) {
    const row = rows[i];
    if (!row || !row.some(c => c)) continue;
    // Verificar si esta fila contiene una columna de fecha
    const hasDate = row.some(cell => {
      const cu = String(cell).toUpperCase().trim();
      return dateHints.some(h => cu.includes(h.toUpperCase()));
    });
    if (!hasDate) continue;
    // Construir mapa de nombre→índice
    const colMap = {};
    row.forEach((cell, idx) => {
      const k = String(cell).toUpperCase().trim();
      if (k) colMap[k] = idx;
    });
    return { headerIdx: i, colMap };
  }
  return null;
}

/**
 * Busca el índice de columna usando hints (exacto → parcial).
 */
function findColIdx(colMap, hints) {
  // Exacta
  for (const h of hints) {
    const k = h.toUpperCase().trim();
    if (colMap[k] !== undefined) return colMap[k];
  }
  // Parcial: el nombre de la columna contiene el hint
  for (const h of hints) {
    const k = h.toUpperCase().trim();
    const key = Object.keys(colMap).find(ck => ck.includes(k));
    if (key !== undefined) return colMap[key];
  }
  return -1;
}

/**
 * Parser principal: strings → filas normalizadas.
 * Recibe el worksheet, el tipo (Ingreso/Egreso) y hints de columnas.
 */
function parseSheetStrings(ws, tipoReg, dateHints, totalHints, tipoHints, nameHints,
                           importeHints, ivaHints, retHints) {
  if (!ws) return [];
  const rows = wsToStrings(ws);
  if (!rows.length) return [];

  const hdr = detectHeader(rows, dateHints);
  if (!hdr) return [];

  const { headerIdx, colMap } = hdr;
  const cFecha   = findColIdx(colMap, dateHints);
  const cTotal   = findColIdx(colMap, totalHints);
  const cTipo    = findColIdx(colMap, tipoHints);
  const cNombre  = findColIdx(colMap, nameHints);
  const cImporte = importeHints ? findColIdx(colMap, importeHints) : -1;
  const cIva     = ivaHints     ? findColIdx(colMap, ivaHints)     : -1;
  const cRet     = retHints     ? findColIdx(colMap, retHints)     : -1;

  if (cFecha === -1 || cTotal === -1) return [];

  const result = [];
  for (let i = headerIdx + 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row || !row.some(c => c)) continue;

    const rawFecha = row[cFecha] || '';
    const rawTotal = row[cTotal] || '';
    if (!rawFecha && !rawTotal) continue;

    // Fecha puede ser null — esas filas se incluyen como "Sin fecha"
    const fecha = parseDate(rawFecha);

    const total = parseMonto(rawTotal);
    if (isNaN(total) || total === 0) continue;

    // ── Detectar y OMITIR filas de totales / resumen ──────
    // Una fila de total desnuda: sin fecha, sin proveedor, sin tipo → solo el número
    const rawTipo   = cTipo   >= 0 ? String(row[cTipo]   || '').trim() : '';
    const rawNombre = cNombre >= 0 ? String(row[cNombre] || '').trim() : '';
    const isTotalRow =
      !rawFecha && !rawTipo && !rawNombre;           // sin fecha + sin contexto
    const isLabeledTotal = /^(total|totales|suma|subtotal)/i.test(rawNombre) ||
                           /^(total|totales|suma|subtotal)/i.test(rawTipo);
    if (isTotalRow || isLabeledTotal) continue;

    const tipo    = rawTipo   || 'Sin tipo';
    const nombre  = rawNombre || '—';
    const importe = cImporte >= 0 ? (parseMonto(row[cImporte] || '0') || 0) : Math.abs(total);
    const iva     = cIva     >= 0 ? (parseMonto(row[cIva]     || '0') || 0) : 0;
    const ret     = cRet     >= 0 ? (parseMonto(row[cRet]     || '0') || 0) : 0;

    result.push({
      fecha,                                          // Date | null
      year:          fecha ? fecha.getFullYear() : 'Sin fecha',
      mes:           fecha ? fecha.getMonth()    : null,
      tipo_registro: tipoReg,
      tipo,
      categoria:     tipo,
      subcategoria:  nombre,
      monto:         Math.abs(total),
      importe:       Math.abs(importe),
      iva:           Math.abs(iva),
      ret:           Math.abs(ret),
    });
  }
  return result;
}

function parseMultiSheet(wb) {
  const nameIng = findSheetName(wb, 'INGRESO');
  const nameEgr = findSheetName(wb, 'EGRESO');
  if (!nameIng || !nameEgr) throw new Error('No se encontraron hojas INGRESOS y EGRESOS.');

  const rowsIng = parseSheetStrings(
    wb.Sheets[nameIng], 'Ingreso',
    ['FECHA DE PAGO', 'FECHA PAGO', 'FECHA'],           // date
    ['TOTAL'],                                           // total
    ['TIPO'],                                            // tipo
    ['NOMBRE DEL CLIENTE', 'NOMBRE CLIENTE', 'NOMBRE'], // nombre
    ['IMPORTE'],                                         // importe
    ['IVA'],                                             // iva
    null                                                 // sin ret en INGRESOS
  );

  const rowsEgr = parseSheetStrings(
    wb.Sheets[nameEgr], 'Egreso',
    ['FECHA FAC', 'FECHA FACTURA', 'FECHA'],            // date
    ['TOTAL'],                                           // total
    ['TIPO'],                                            // tipo
    ['PROVEEDOR'],                                       // nombre
    ['IMPORTE'],                                         // importe
    ['IVA'],                                             // iva
    ['RET', 'RET/ ISR', 'RET/ISR', 'RETENCION']        // ret
  );

  const combined = [...rowsIng, ...rowsEgr];
  if (combined.length === 0) return [];

  // Auto-detectar el año más frecuente (solo filas con fecha válida)
  const yearCount = {};
  for (const r of combined) {
    if (r.year !== 'Sin fecha') yearCount[r.year] = (yearCount[r.year] || 0) + 1;
  }
  const topYear = Object.entries(yearCount).sort((a,b) => b[1]-a[1])[0];
  detectedYear = topYear ? parseInt(topYear[0], 10) : new Date().getFullYear();

  // Retornar TODOS los años (el filtrado por año ocurre en buildDashboard)
  return combined;
}

// ══════════════════════════════════════════════════════
// 3B. PARSER HOJA ÚNICA GENÉRICA
// ══════════════════════════════════════════════════════

const COL_MAP = {
  fecha:      ['fecha','date','fec','dia','day','periodo','fecha de pago','fecha fac'],
  tipo:       ['tipo','type','clase','movimiento','naturaleza'],
  categoria:  ['categoria','categoría','category','cat','rubro','concepto','descripcion'],
  subcategoria:['subcategoria','subcategoría','subcategory','subcat','detalle'],
  monto:      ['monto','amount','valor','importe','total','cantidad','sum'],
};

function findCol(headers, aliases) {
  const hLow = headers.map(h => String(h).toLowerCase().trim());
  for (const alias of aliases) {
    const idx = hLow.indexOf(alias);
    if (idx !== -1) return headers[idx];
  }
  return null;
}

function normalizeSingleSheet(raw) {
  const headers    = Object.keys(raw[0]);
  const colFecha   = findCol(headers, COL_MAP.fecha);
  const colTipo    = findCol(headers, COL_MAP.tipo);
  const colCat     = findCol(headers, COL_MAP.categoria);
  const colSubcat  = findCol(headers, COL_MAP.subcategoria);
  const colMonto   = findCol(headers, COL_MAP.monto);

  if (!colFecha) throw new Error('No se encontró columna de Fecha.');
  if (!colMonto) throw new Error('No se encontró columna de Monto.');

  const rows = [];
  for (const row of raw) {
    try {
      const fecha = parseDate(row[colFecha]);   // puede ser null → 'Sin fecha'
      const monto = parseMonto(row[colMonto]);
      if (isNaN(monto) || monto === 0) continue;

      let tipoReg = 'Egreso';
      if (colTipo && row[colTipo]) {
        const rt = String(row[colTipo]).toLowerCase().trim();
        if (rt.includes('ingreso')||rt.includes('income')||rt.includes('entrada')||rt==='+'||rt.includes('crédito')) tipoReg = 'Ingreso';
        else if (rt.includes('egreso')||rt.includes('gasto')||rt.includes('expense')||rt.includes('salida')||rt==='-') tipoReg = 'Egreso';
        else tipoReg = monto >= 0 ? 'Ingreso' : 'Egreso';
      } else {
        tipoReg = monto >= 0 ? 'Ingreso' : 'Egreso';
      }

      const cat    = colCat    && row[colCat]    ? String(row[colCat]).trim()    : 'Sin categoría';
      const subcat = colSubcat && row[colSubcat] ? String(row[colSubcat]).trim() : '—';
      const tipoVal= colTipo   && row[colTipo]   ? String(row[colTipo]).trim()   : tipoReg;

      rows.push({
        fecha,
        year:          fecha ? fecha.getFullYear() : 'Sin fecha',
        mes:           fecha ? fecha.getMonth()    : null,
        tipo_registro: tipoReg,
        tipo:          capitalizar(tipoVal),
        categoria:     capitalizar(cat),
        subcategoria:  capitalizar(subcat),
        monto:         Math.abs(monto),
      });
    } catch (_) { /* fila con error: ignorar */ }
  }

  if (rows.length === 0) return [];

  // Auto-detectar año dominante (solo filas con fecha válida)
  const yearCount = {};
  for (const r of rows) {
    if (r.year !== 'Sin fecha') yearCount[r.year] = (yearCount[r.year] || 0) + 1;
  }
  const topYear = Object.entries(yearCount).sort((a,b) => b[1]-a[1])[0];
  detectedYear = topYear ? parseInt(topYear[0], 10) : new Date().getFullYear();

  // Retornar TODOS los años
  return rows;
}

/** Wrapper: recibe un worksheet, normaliza con la función genérica. */
function parseSingleSheet(ws) {
  const raw = XLSX.utils.sheet_to_json(ws, {
    raw: false, dateNF: 'yyyy-mm-dd', defval: '',
  });
  if (!raw || raw.length === 0) throw new Error('La hoja no contiene datos.');
  return normalizeSingleSheet(raw);
}

// ══════════════════════════════════════════════════════
// 4. CONSTRUCCIÓN DEL DASHBOARD
// ══════════════════════════════════════════════════════

function buildDashboard(rows) {
  allRows  = rows;
  allYears = [...new Set(rows.map(r => r.year))].sort((a,b) => a - b);

  // Por defecto mostrar TODOS los datos (para que el total coincida con el Excel)
  yearRows     = rows;
  filteredRows = rows;

  renderKPIs(filteredRows);
  renderBarChart(filteredRows);
  renderDonutChart(filteredRows);
  renderTipoCharts(filteredRows);
  renderCategoryTable(filteredRows);
  renderTxTable(filteredRows);
  buildYearFilter(rows);
  buildMonthFilter(rows);

  hideLoading();
  document.getElementById('uploadSection').classList.add('hidden');
  document.getElementById('dashboardSection').classList.remove('hidden');
  document.getElementById('dashSubtitle').textContent = 'Datos completos';
  document.getElementById('yearFilterWrapper').classList.remove('hidden');
  document.getElementById('monthFilterWrapper').classList.remove('hidden');
  document.getElementById('exportCsvBtn').classList.remove('hidden');
  window.scrollTo({ top: 0, behavior: 'smooth' });
}

// ── KPIs ──
function renderKPIs(rows) {
  const ing = rows.filter(r => r.tipo_registro === 'Ingreso');
  const egr = rows.filter(r => r.tipo_registro === 'Egreso');
  const totalIng = ing.reduce((s,r) => s + r.monto, 0);
  const totalEgr = egr.reduce((s,r) => s + r.monto, 0);
  const balance  = totalIng - totalEgr;
  const tasa     = totalIng > 0 ? (balance / totalIng * 100) : 0;

  document.getElementById('kpiIncome').textContent        = formatMoney(totalIng);
  document.getElementById('kpiIncomeDetail').textContent  = `${ing.length} transacciones`;
  document.getElementById('kpiExpense').textContent       = formatMoney(totalEgr);
  document.getElementById('kpiExpenseDetail').textContent = `${egr.length} transacciones`;
  document.getElementById('kpiBalance').textContent       = formatMoney(balance);
  document.getElementById('kpiBalanceDetail').textContent = balance >= 0 ? '✓ Balance positivo' : '⚠ Balance negativo';
  document.getElementById('kpiBalance').style.color       = balance >= 0 ? 'var(--income)' : 'var(--expense)';
  document.getElementById('kpiRate').textContent          = `${tasa.toFixed(1)}%`;
}

// ── Barras: Ingresos vs Egresos por mes ──
function renderBarChart(rows) {
  const meses = Array.from({length:12}, ()=>({ing:0, egr:0}));
  const conDatos = new Set();
  for (const r of rows) {
    if (r.mes === null) continue;  // filas sin fecha no se grafican por mes
    meses[r.mes][r.tipo_registro==='Ingreso'?'ing':'egr'] += r.monto;
    conDatos.add(r.mes);
  }
  const sorted   = [...conDatos].sort((a,b)=>a-b);
  const labels   = sorted.map(m => MONTHS_ES[m]);
  const ingData  = sorted.map(m => meses[m].ing);
  const egrData  = sorted.map(m => meses[m].egr);

  const isDark   = document.documentElement.getAttribute('data-theme') !== 'light';
  const grid     = isDark ? 'rgba(255,255,255,0.06)' : 'rgba(0,0,0,0.06)';
  const tick     = isDark ? '#64748b' : '#94a3b8';
  const ctx      = document.getElementById('barChart').getContext('2d');

  if (barChartInst) barChartInst.destroy();
  barChartInst = new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [
        { label:'Ingresos', data:ingData, backgroundColor:'rgba(16,185,129,0.75)', borderRadius:6, borderSkipped:false },
        { label:'Egresos',  data:egrData, backgroundColor:'rgba(244,63,94,0.75)',  borderRadius:6, borderSkipped:false },
      ],
    },
    options: {
      responsive:true, maintainAspectRatio:false,
      interaction:{mode:'index',intersect:false},
      plugins:{
        legend:{display:false},
        tooltip: tooltipDefaults(isDark, v => formatMoney(v)),
      },
      scales:{
        x:{grid:{display:false}, ticks:{color:tick, font:{family:'Inter',size:12}}},
        y:{grid:{color:grid}, border:{display:false},
           ticks:{color:tick, font:{family:'Inter',size:12}, callback:v=>formatMoneyShort(v)}},
      },
    },
  });
}

// ── Dona: Egresos por tipo ──
function renderDonutChart(rows) {
  const egr    = rows.filter(r => r.tipo_registro === 'Egreso');
  const byTipo = agrupar(egr, 'tipo');
  const total  = egr.reduce((s,r) => s + r.monto, 0);

  // Top 8 + "Otros"
  const all       = Object.entries(byTipo).sort((a,b) => b[1] - a[1]);
  const top       = all.slice(0, 8);
  const othersSum = all.slice(8).reduce((s,[,v]) => s + v, 0);
  const slices    = othersSum > 0 ? [...top, ['Otros', othersSum]] : top;

  const isDark     = document.documentElement.getAttribute('data-theme') !== 'light';
  const lblColor   = '#ffffff';                          // blanco para contraste
  const subColor   = isDark ? '#94a3b8' : '#64748b';
  const centerMain = isDark ? '#f1f5f9' : '#0f172a';

  // Plugin: dibuja TOTAL + valor en el centro del anillo (canvas-space)
  const centerPlugin = {
    id: 'donutCenter',
    beforeDraw(chart) {
      const { ctx: c, chartArea } = chart;
      if (!chartArea) return;
      const cx = chartArea.left + chartArea.width  / 2;
      const cy = chartArea.top  + chartArea.height / 2;
      c.save();
      c.textAlign    = 'center';
      c.textBaseline = 'middle';
      // Sub-label
      c.font      = '500 11px Inter, system-ui, sans-serif';
      c.fillStyle = subColor;
      c.fillText('TOTAL', cx, cy - 15);
      // Valor principal
      c.font      = '800 18px Inter, system-ui, sans-serif';
      c.fillStyle = centerMain;
      c.fillText(formatMoneyShort(total), cx, cy + 10);
      c.restore();
    },
  };

  const ctxEl = document.getElementById('donutChart').getContext('2d');
  if (donutChartInst) donutChartInst.destroy();
  donutChartInst = new Chart(ctxEl, {
    type: 'doughnut',
    plugins: [centerPlugin],
    data: {
      labels: slices.map(([t]) => t),
      datasets: [{
        data:            slices.map(([,v]) => v),
        backgroundColor: slices.map((_,i) => PALETTE[i % PALETTE.length]),
        borderColor:     isDark ? '#1e293b' : '#f8fafc',
        borderWidth:     2,
        hoverOffset:     10,
      }],
    },
    options: {
      responsive:        true,
      maintainAspectRatio: false,
      cutout:            '68%',
      layout:            { padding: { top: 4, bottom: 4, left: 4, right: 4 } },
      plugins: {
        legend: {
          position: 'bottom',
          align:    'center',
          labels: {
            color:         lblColor,
            font:          { family: 'Inter', size: 11, weight: '500' },
            padding:       16,
            boxWidth:      10,
            usePointStyle: true,
            pointStyleWidth: 8,
          },
        },
        tooltip: {
          backgroundColor: isDark ? '#1e293b' : '#fff',
          titleColor:      isDark ? '#f1f5f9' : '#0f172a',
          bodyColor:       isDark ? '#94a3b8' : '#475569',
          borderColor:     isDark ? 'rgba(255,255,255,0.12)' : 'rgba(0,0,0,0.08)',
          borderWidth: 1, padding: 12, cornerRadius: 10,
          callbacks: {
            label: ctx => {
              const v   = ctx.dataset.data[ctx.dataIndex];
              const pct = total > 0 ? (v / total * 100).toFixed(1) : '0';
              return ` ${ctx.label}: ${formatMoney(v)} (${pct}%)`;
            },
          },
        },
      },
    },
  });
}

// ── Barras horizontales: Ingresos Y Egresos por tipo ──
function renderTipoCharts(rows) {
  renderTipoBar(
    rows.filter(r => r.tipo_registro === 'Ingreso'),
    'tipoIngChart',
    'rgba(16,185,129,0.8)',
    'tipoIngBadge'
  );
  renderTipoBar(
    rows.filter(r => r.tipo_registro === 'Egreso'),
    'tipoEgrChart',
    'rgba(244,63,94,0.8)',
    'tipoEgrBadge'
  );
}

function renderTipoBar(rows, canvasId, color, badgeId) {
  const byTipo = agrupar(rows, 'tipo');
  const sorted = Object.entries(byTipo).sort((a,b)=>b[1]-a[1]);
  const labels = sorted.map(([t])=>t);
  const data   = sorted.map(([,v])=>v);

  document.getElementById(badgeId).textContent = `${sorted.length} tipos`;

  // Altura dinámica: 36px por barra + 40px padding, mínimo 220px
  const dynamicHeight = Math.max(220, sorted.length * 36 + 40);
  const wrapper = document.getElementById(canvasId).parentElement;
  wrapper.style.height = dynamicHeight + 'px';

  const isDark = document.documentElement.getAttribute('data-theme') !== 'light';
  const grid   = isDark ? 'rgba(255,255,255,0.06)' : 'rgba(0,0,0,0.06)';
  const tick   = isDark ? '#64748b' : '#94a3b8';

  const existingChart = Chart.getChart(canvasId);
  if (existingChart) existingChart.destroy();

  const ctx = document.getElementById(canvasId).getContext('2d');
  new Chart(ctx, {
    type:'bar',
    data:{
      labels,
      datasets:[{
        data,
        backgroundColor: color,
        borderRadius:5,
        borderSkipped:false,
      }],
    },
    options:{
      indexAxis:'y',
      responsive:true, maintainAspectRatio:false,
      plugins:{
        legend:{display:false},
        tooltip: tooltipDefaults(isDark, v => formatMoney(v)),
      },
      scales:{
        x:{grid:{color:grid}, border:{display:false},
           ticks:{color:tick, font:{family:'Inter',size:11}, callback:v=>formatMoneyShort(v)}},
        y:{grid:{display:false}, ticks:{color:isDark?'#94a3b8':'#475569', font:{family:'Inter',size:11}}},
      },
    },
  });
}

// ══════════════════════════════════════════════════════
// TABLAS DESGLOSE POR TIPO — controlador genérico
// ══════════════════════════════════════════════════════

const CAT_COLS = [
  { key:'cat',   label:'Tipo',          align:'left',  filterable:true,  numeric:false },
  { key:'total', label:'Total',         align:'right', filterable:false, numeric:true  },
  { key:'pct',   label:'% del Total',   align:'right', filterable:false, numeric:false },
  { key:'count', label:'Transacciones', align:'right', filterable:false, numeric:false },
];

function catGetVal(row, key) { return String(row[key] ?? ''); }

/** Construye las entries de desglose para un array de rows del mismo tipo */
function buildCatEntries(rows) {
  const total = rows.reduce((s, r) => s + r.monto, 0);
  const byTipo = agrupar(rows, 'tipo');
  const cnt    = contar(rows, 'tipo');
  return Object.keys(byTipo)
    .map(cat => ({
      cat,
      total: byTipo[cat],
      count: cnt[cat] || 0,
      pct:   total > 0 ? byTipo[cat] / total * 100 : 0,
    }))
    .sort((a, b) => b.total - a.total);
}

function renderCategoryTable(rows) {
  const ingRows = rows.filter(r => r.tipo_registro === 'Ingreso');
  const egrRows = rows.filter(r => r.tipo_registro === 'Egreso');

  // Cargar controladores y renderizar cada tabla
  catIngCtrl.allRows    = buildCatEntries(ingRows);
  catIngCtrl.colFilters = {};
  catIngCtrl.openCol    = null;
  buildCatHeader(catIngCtrl);
  refreshCatTable(catIngCtrl);

  catEgrCtrl.allRows    = buildCatEntries(egrRows);
  catEgrCtrl.colFilters = {};
  catEgrCtrl.openCol    = null;
  buildCatHeader(catEgrCtrl);
  refreshCatTable(catEgrCtrl);
}

function buildCatHeader(ctrl) {
  const thead = document.querySelector(`#${ctrl.tableId} thead tr`);
  if (!thead) return;
  thead.innerHTML = '';
  CAT_COLS.forEach(col => {
    const th = document.createElement('th');
    th.className = col.align === 'right' ? 'cat-th-right' : 'cat-th-left';
    if (col.filterable) {
      th.innerHTML = `
        <span class="th-label">${col.label}</span>
        <button class="tx-filter-btn cat-filter-btn" title="Filtrar por ${col.label}">
          <svg width="11" height="11" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24">
            <polygon points="22 3 2 3 10 12.46 10 19 14 21 14 12.46 22 3"/>
          </svg>
        </button>`;
      const btn = th.querySelector('.cat-filter-btn');
      btn._catCtrl = ctrl;
      btn._catKey  = col.key;
      btn.addEventListener('click', e => {
        e.stopPropagation();
        toggleCatDropdown(e.currentTarget._catCtrl, e.currentTarget._catKey, e.currentTarget);
      });
    } else {
      th.innerHTML = `<span class="th-label">${col.label}</span>`;
    }
    thead.appendChild(th);
  });
}

function getCatFiltered(ctrl) {
  return ctrl.allRows.filter(row => {
    for (const [key, allowed] of Object.entries(ctrl.colFilters)) {
      if (!allowed || allowed.size === 0) continue;
      if (!allowed.has(catGetVal(row, key))) return false;
    }
    return true;
  });
}

function refreshCatTable(ctrl) {
  const visible = getCatFiltered(ctrl);
  renderCatBody(ctrl, visible);

  const badge = document.getElementById(ctrl.badgeId);
  if (badge) badge.textContent =
    visible.length === ctrl.allRows.length
      ? `${visible.length} tipos`
      : `${visible.length} de ${ctrl.allRows.length} tipos`;

  // Actualizar estado de botones de filtro
  CAT_COLS.filter(c => c.filterable).forEach(col => {
    // Encuentra el botón de ESTE ctrl buscando por _catCtrl
    const allBtns = document.querySelectorAll('.cat-filter-btn');
    allBtns.forEach(btn => {
      if (btn._catCtrl !== ctrl || btn._catKey !== col.key) return;
      const isActive = !!(ctrl.colFilters[col.key] && ctrl.colFilters[col.key].size > 0);
      btn.classList.toggle('active', isActive);
      let dot = btn.querySelector('.tx-filter-dot');
      if (isActive && !dot) { dot = document.createElement('span'); dot.className='tx-filter-dot'; btn.appendChild(dot); }
      if (!isActive && dot) dot.remove();
    });
  });
}

function renderCatBody(ctrl, entries) {
  const tbody = document.getElementById(ctrl.bodyId);
  if (!tbody) return;
  const frag = document.createDocumentFragment();

  for (const e of entries) {
    const tr = document.createElement('tr');
    const pctBar = e.pct > 0
      ? `<div class="progress-cell">
           <span class="pct-label">${e.pct.toFixed(1)}%</span>
           <div class="progress-bar"><div class="progress-fill" style="width:${Math.min(e.pct,100)}%"></div></div>
         </div>`
      : '<span class="td-nil">—</span>';
    tr.innerHTML = `
      <td class="cat-td-nombre"><strong>${escHtml(e.cat)}</strong></td>
      <td class="text-right cat-td-total">${formatMoney(e.total)}</td>
      <td class="text-right cat-td-pct">${pctBar}</td>
      <td class="text-right cat-td-count">${e.count.toLocaleString('es-MX')}</td>`;
    frag.appendChild(tr);
  }
  tbody.innerHTML = '';
  tbody.appendChild(frag);
}

// ── DROPDOWN PARA TABLAS DESGLOSE (genérico) ──
function toggleCatDropdown(ctrl, colKey, btn) {
  const existing = document.getElementById('txDropdownPanel');
  if (catActiveCtrl === ctrl && ctrl.openCol === colKey && existing) {
    closeCatDropdown(); return;
  }
  closeCatDropdown();
  catActiveCtrl = ctrl;
  ctrl.openCol  = colKey;
  openCatDropdown(ctrl, colKey, btn);
}

function openCatDropdown(ctrl, colKey, anchorBtn) {
  const col       = CAT_COLS.find(c => c.key === colKey);
  const allVals   = [...new Set(ctrl.allRows.map(r => catGetVal(r, colKey)))].sort();
  const activeSet = ctrl.colFilters[colKey] || null;

  const panel = document.createElement('div');
  panel.id        = 'txDropdownPanel';
  panel.className = 'tx-dropdown-panel';
  panel.innerHTML = `
    <div class="tx-dp-title">${escHtml(ctrl.label + ' — ' + (col?.label || colKey))}</div>
    <div class="tx-dp-head">
      <div class="tx-dp-search-wrap">
        <svg class="tx-dp-search-icon" width="13" height="13" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24">
          <circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/>
        </svg>
        <input id="txDpSearch" class="tx-dp-search" placeholder="Buscar…" autocomplete="off"/>
      </div>
    </div>
    <div id="txDpList" class="tx-dp-list"></div>
    <div class="tx-dp-foot">
      <button id="txDpApply" class="tx-dp-btn-apply">Aplicar</button>
      <button id="txDpClear" class="tx-dp-btn-clear">Limpiar</button>
    </div>`;
  document.body.appendChild(panel);

  const list = panel.querySelector('#txDpList');
  function renderOptions(filterText) {
    list.innerHTML = '';
    const filtered = allVals.filter(v => !filterText || v.toLowerCase().includes(filterText));
    const saRow = document.createElement('label');
    saRow.className = 'tx-dp-item tx-dp-select-all-item';
    const allChecked  = filtered.every(v => !activeSet || activeSet.has(v));
    const someChecked = filtered.some(v => !activeSet || activeSet.has(v));
    saRow.innerHTML = `<input type="checkbox" id="txDpCbAll" ${allChecked?'checked':''}>
      <span><strong>Seleccionar todo</strong></span><span class="tx-dp-count">${filtered.length}</span>`;
    const cbAll = saRow.querySelector('#txDpCbAll');
    if (!allChecked && someChecked) cbAll.indeterminate = true;
    cbAll.addEventListener('change', () => list.querySelectorAll('.tx-dp-value-cb').forEach(i=>i.checked=cbAll.checked));
    list.appendChild(saRow);
    const sep = document.createElement('div'); sep.className='tx-dp-sep'; list.appendChild(sep);
    filtered.forEach(val => {
      const checked = !activeSet || activeSet.has(val);
      const item = document.createElement('label');
      item.className = 'tx-dp-item';
      item.innerHTML = `<input type="checkbox" class="tx-dp-value-cb" value="${escHtml(val)}" ${checked?'checked':''}>
        <span title="${escHtml(val)}">${escHtml(val)}</span>`;
      list.appendChild(item);
    });
    list.addEventListener('change', e => {
      if (e.target.classList.contains('tx-dp-value-cb')) {
        const cbs = [...list.querySelectorAll('.tx-dp-value-cb')];
        const sel = cbs.filter(i=>i.checked).length;
        const cbA = list.querySelector('#txDpCbAll');
        if (cbA) { cbA.checked=sel===cbs.length; cbA.indeterminate=sel>0&&sel<cbs.length; }
      }
    });
  }
  renderOptions('');
  panel.querySelector('#txDpSearch').addEventListener('input', e => renderOptions(e.target.value.toLowerCase()));
  panel.querySelector('#txDpApply').addEventListener('click', () => {
    const checked = [...list.querySelectorAll('.tx-dp-value-cb:checked')].map(i=>i.value);
    if (checked.length === allVals.length || checked.length === 0) delete ctrl.colFilters[colKey];
    else ctrl.colFilters[colKey] = new Set(checked);
    closeCatDropdown(); refreshCatTable(ctrl);
  });
  panel.querySelector('#txDpClear').addEventListener('click', () => {
    delete ctrl.colFilters[colKey]; closeCatDropdown(); refreshCatTable(ctrl);
  });

  const rect = anchorBtn.getBoundingClientRect();
  panel.style.top  = `${rect.bottom + 4}px`;
  panel.style.left = `${Math.min(rect.left, window.innerWidth - 288)}px`;
  requestAnimationFrame(() => {
    const ph = panel.offsetHeight;
    if (rect.bottom + ph > window.innerHeight - 8) panel.style.top = `${rect.top - ph - 4}px`;
  });
  setTimeout(() => document.addEventListener('click', outsideCatClick), 0);
}

function outsideCatClick(e) {
  const panel = document.getElementById('txDropdownPanel');
  if (panel && !panel.contains(e.target)) closeCatDropdown();
}
function closeCatDropdown() {
  const panel = document.getElementById('txDropdownPanel');
  if (panel) panel.remove();
  if (catActiveCtrl) { catActiveCtrl.openCol = null; catActiveCtrl = null; }
  document.removeEventListener('click', outsideCatClick);
}

// ══════════════════════════════════════════════════════
// TABLA AVANZADA DE TRANSACCIONES
// Filtros tipo Excel + totales dinámicos + sin límite de filas
// ══════════════════════════════════════════════════════

const TX_COLS = [
  { key:'fecha_str',    label:'Fecha',              align:'left',  filterable:true,  numeric:false },
  { key:'tipo_registro',label:'Movimiento',         align:'left',  filterable:true,  numeric:false },
  { key:'tipo',         label:'Tipo',               align:'left',  filterable:true,  numeric:false },
  { key:'subcategoria', label:'Proveedor / Cliente', align:'left', filterable:true,  numeric:false },
  { key:'importe',      label:'Importe',            align:'right', filterable:false, numeric:true  },
  { key:'iva',          label:'IVA',                align:'right', filterable:false, numeric:true  },
  { key:'ret',          label:'Ret',                align:'right', filterable:false, numeric:true  },
  { key:'monto',        label:'Total',              align:'right', filterable:false, numeric:true  },
];

function txGetVal(row, key) {
  if (key === 'fecha_str') return formatDate(row.fecha);
  if (key === 'year')      return String(row.year ?? 'Sin fecha');
  return String(row[key] ?? '');
}

/** Construye / actualiza el botón de filtro de Año en el header de la tabla */
function buildTxYearControl() {
  const btn = document.getElementById('txYearFilterBtn');
  if (!btn) return;
  btn.onclick = e => { e.stopPropagation(); toggleTxDropdown('year', btn); };
  syncYearBtnState();
}

function syncYearBtnState() {
  const btn = document.getElementById('txYearFilterBtn');
  if (!btn) return;
  const isActive = !!(txColFilters['year'] && txColFilters['year'].size > 0);
  btn.classList.toggle('active', isActive);
  const lbl = btn.querySelector('.tx-yr-label');
  if (lbl) lbl.textContent = isActive ? `Año (${txColFilters['year'].size})` : 'Año';
}

/** Cuenta cuántos filtros de columna están activos */
function txActiveFilterCount() {
  return Object.values(txColFilters).filter(s => s && s.size > 0).length;
}

function renderTxTable(rows) {
  txCurrentRows = [...rows].sort((a, b) => {
    if (!a.fecha && !b.fecha) return 0;
    if (!a.fecha) return 1;   // sin fecha → al final
    if (!b.fecha) return -1;
    return b.fecha - a.fecha;
  });
  txColFilters = {};
  txOpenCol    = null;
  txSortCol    = null;
  txSortDir    = null;
  // Reset botones rápidos
  ['Ingreso','Egreso'].forEach(t => {
    const b = document.getElementById(`qf${t}`);
    if (b) b.classList.remove('active');
  });
  buildTxHeader();
  buildTxYearControl();
  refreshTxTable();
}

function refreshTxTable() {
  const filtered = getTxFiltered();
  const visible  = txSortRows(filtered);   // sort DESPUÉS de filtrar
  renderTxBody(visible);
  renderTxFooter(filtered);  // totales sobre TODOS los filtrados (independiente del sort)

  // Badge de registros
  const badge = document.getElementById('txBadge');
  badge.textContent = visible.length === txCurrentRows.length
    ? `${visible.length.toLocaleString('es-MX')} registros`
    : `${visible.length.toLocaleString('es-MX')} de ${txCurrentRows.length.toLocaleString('es-MX')} registros`;

  // Resaltar botones de filtro activos + contador
  TX_COLS.filter(c => c.filterable).forEach(c => {
    const btn = document.querySelector(`.tx-filter-btn[data-col="${c.key}"]`);
    if (!btn) return;
    const isActive = !!(txColFilters[c.key] && txColFilters[c.key].size > 0);
    btn.classList.toggle('active', isActive);
    // Actualizar el dot indicador dentro del botón
    let dot = btn.querySelector('.tx-filter-dot');
    if (isActive) {
      if (!dot) {
        dot = document.createElement('span');
        dot.className = 'tx-filter-dot';
        btn.appendChild(dot);
      }
    } else {
      if (dot) dot.remove();
    }
  });

  // Botón "Limpiar todos" — aparece solo si hay filtros activos
  const clearAllBtn = document.getElementById('txClearAllBtn');
  if (clearAllBtn) {
    const count = txActiveFilterCount();
    clearAllBtn.style.display = count > 0 ? 'inline-flex' : 'none';
    clearAllBtn.textContent   = count > 1 ? `Limpiar ${count} filtros` : 'Limpiar filtro';
  }
  // Sincronizar botón de año
  syncYearBtnState();
}

function getTxFiltered() {
  return txCurrentRows.filter(row => {
    for (const [key, allowed] of Object.entries(txColFilters)) {
      if (!allowed || allowed.size === 0) continue;
      if (!allowed.has(txGetVal(row, key))) return false;
    }
    return true;
  });
}

function buildTxHeader() {
  const thead = document.querySelector('#txTable thead tr');
  if (!thead) return;
  thead.innerHTML = '';
  TX_COLS.forEach(col => {
    const th = document.createElement('th');
    th.className = (col.align === 'right' ? 'text-right ' : '') + 'tx-th-sortable';
    th.dataset.col = col.key;

    // Icono de ordenamiento
    const sortIcon = `<span class="tx-sort-icon" data-col="${col.key}">
      <svg width="10" height="10" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24" class="sort-neutral">
        <path d="M7 15l5 5 5-5M7 9l5-5 5 5"/>
      </svg></span>`;

    if (col.filterable) {
      th.innerHTML = `
        <span class="th-label">${col.label}</span>
        <button class="tx-filter-btn" data-col="${col.key}" title="Filtrar por ${col.label}">
          <svg width="11" height="11" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24">
            <polygon points="22 3 2 3 10 12.46 10 19 14 21 14 12.46 22 3"/>
          </svg>
        </button>
        ${sortIcon}`;
      th.querySelector('.tx-filter-btn').addEventListener('click', e => {
        e.stopPropagation();
        toggleTxDropdown(col.key, e.currentTarget);
      });
    } else {
      th.innerHTML = `<span class="th-label">${col.label}</span>${sortIcon}`;
    }

    // Click en la columna (no en el botón filtro) → ordenar
    th.addEventListener('click', e => {
      if (e.target.closest('.tx-filter-btn')) return;
      cycleTxSort(col.key);
    });

    thead.appendChild(th);
  });
}

/** Cicla el ordenamiento: null → asc → desc → null */
function cycleTxSort(colKey) {
  if (txSortCol !== colKey) { txSortCol = colKey; txSortDir = 'asc'; }
  else if (txSortDir === 'asc') { txSortDir = 'desc'; }
  else { txSortCol = null; txSortDir = null; }
  updateSortIcons();
  refreshTxTable();
}

/** Actualiza los iconos de sort en el thead */
function updateSortIcons() {
  document.querySelectorAll('.tx-sort-icon').forEach(el => {
    const col = el.dataset.col;
    el.innerHTML = col === txSortCol
      ? (txSortDir === 'asc'
          ? `<svg width="10" height="10" fill="none" stroke="var(--accent)" stroke-width="2.5" viewBox="0 0 24 24"><path d="M12 19V5M5 12l7-7 7 7"/></svg>`
          : `<svg width="10" height="10" fill="none" stroke="var(--accent)" stroke-width="2.5" viewBox="0 0 24 24"><path d="M12 5v14M19 12l-7 7-7-7"/></svg>`)
      : `<svg width="10" height="10" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24" class="sort-neutral"><path d="M7 15l5 5 5-5M7 9l5-5 5 5"/></svg>`;
  });
}

/** Ordena rows según txSortCol / txSortDir */
function txSortRows(rows) {
  if (!txSortCol || !txSortDir) return rows;
  const col   = TX_COLS.find(c => c.key === txSortCol);
  const isNum = col && col.numeric;
  return [...rows].sort((a, b) => {
    let av, bv;
    if (isNum) {
      av = typeof a[txSortCol] === 'number' ? a[txSortCol] : 0;
      bv = typeof b[txSortCol] === 'number' ? b[txSortCol] : 0;
    } else {
      av = txGetVal(a, txSortCol).toLowerCase();
      bv = txGetVal(b, txSortCol).toLowerCase();
    }
    if (av < bv) return txSortDir === 'asc' ? -1 :  1;
    if (av > bv) return txSortDir === 'asc' ?  1 : -1;
    return 0;
  });
}

/** Aplica filtro rápido de tipo (Ingresos / Egresos) */
function applyQuickFilter(tipoReg) {
  const btn = document.getElementById(`qf${tipoReg}`);
  // Toggle: si ya está activo este filtro, quitar
  const current = txColFilters['tipo_registro'];
  const isActive = current && current.size === 1 && current.has(tipoReg);
  if (isActive) {
    delete txColFilters['tipo_registro'];
  } else {
    txColFilters['tipo_registro'] = new Set([tipoReg]);
  }
  // Sincronizar estado visual de los botones rápidos
  ['Ingreso','Egreso'].forEach(t => {
    const b = document.getElementById(`qf${t}`);
    if (b) b.classList.toggle('active', !!(txColFilters['tipo_registro']?.has(t)));
  });
  refreshTxTable();
}

function renderTxBody(rows) {
  const tbody = document.getElementById('txTableBody');
  const frag  = document.createDocumentFragment();

  rows.forEach(r => {
    const isInc = r.tipo_registro === 'Ingreso';
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td class="td-date">${escHtml(formatDate(r.fecha))}</td>
      <td><span class="type-badge ${isInc ? 'type-income' : 'type-expense'}">${r.tipo_registro}</span></td>
      <td><strong class="td-tipo">${escHtml(r.tipo)}</strong></td>
      <td class="td-nombre">${escHtml(r.subcategoria)}</td>
      <td class="text-right td-num">${r.importe > 0 ? formatMoney(r.importe) : '<span class="td-nil">—</span>'}</td>
      <td class="text-right td-num">${r.iva     > 0 ? formatMoney(r.iva)     : '<span class="td-nil">—</span>'}</td>
      <td class="text-right td-num">${r.ret     > 0 ? formatMoney(r.ret)     : '<span class="td-nil">—</span>'}</td>
      <td class="text-right td-total ${isInc ? 'td-total-inc' : 'td-total-egr'}">
        ${isInc ? '+' : '-'}${formatMoney(r.monto)}
      </td>`;
    frag.appendChild(tr);
  });

  tbody.innerHTML = '';
  tbody.appendChild(frag);
}

function renderTxFooter(rows) {
  const sumImporte = rows.reduce((s, r) => s + r.importe, 0);
  const sumIva     = rows.reduce((s, r) => s + r.iva,     0);
  const sumRet     = rows.reduce((s, r) => s + r.ret,     0);
  const sumTotal   = rows.reduce((s, r) => s + r.monto,   0);

  const get = id => document.getElementById(id);
  if (get('footImporte')) get('footImporte').textContent = formatMoney(sumImporte);
  if (get('footIva'))     get('footIva').textContent     = formatMoney(sumIva);
  if (get('footRet'))     get('footRet').textContent     = formatMoney(sumRet);
  if (get('footTotal'))   get('footTotal').textContent   = formatMoney(sumTotal);
  if (get('footCount'))   get('footCount').textContent   = `${rows.length.toLocaleString('es-MX')} registros`;
}

// ── DROPDOWN TIPO EXCEL ──────────────────────────────

function toggleTxDropdown(colKey, btn) {
  const existing = document.getElementById('txDropdownPanel');
  if (txOpenCol === colKey && existing) { closeTxDropdown(); return; }
  closeTxDropdown();
  txOpenCol = colKey;
  openTxDropdown(colKey, btn);
}

function openTxDropdown(colKey, anchorBtn) {
  const colLabel   = TX_COLS.find(c => c.key === colKey)?.label || colKey;
  // Para el key 'year': orden numérico ascendente con 'Sin fecha' al final
  const rawVals  = [...new Set(txCurrentRows.map(r => txGetVal(r, colKey)))];
  const allVals  = colKey === 'year'
    ? rawVals.sort((a, b) => {
        if (a === 'Sin fecha') return 1;
        if (b === 'Sin fecha') return -1;
        return Number(a) - Number(b);
      })
    : rawVals.sort();
  const activeSet  = txColFilters[colKey] || null;   // null → todos activos
  const totalCount = allVals.length;

  // ── Panel ──
  const panel = document.createElement('div');
  panel.id        = 'txDropdownPanel';
  panel.className = 'tx-dropdown-panel';
  panel.innerHTML = `
    <div class="tx-dp-title">${escHtml(colLabel)}</div>
    <div class="tx-dp-head">
      <div class="tx-dp-search-wrap">
        <svg class="tx-dp-search-icon" width="13" height="13" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24">
          <circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/>
        </svg>
        <input id="txDpSearch" class="tx-dp-search" placeholder="Buscar en ${totalCount} valores…" autocomplete="off"/>
      </div>
    </div>
    <div id="txDpList" class="tx-dp-list"></div>
    <div class="tx-dp-foot">
      <button id="txDpApply" class="tx-dp-btn-apply">Aplicar</button>
      <button id="txDpClear" class="tx-dp-btn-clear">Limpiar</button>
    </div>`;
  document.body.appendChild(panel);

  const list = panel.querySelector('#txDpList');

  // ── Renderiza checkboxes con "Seleccionar todo" al inicio ──
  function renderOptions(filterText) {
    list.innerHTML = '';
    const filtered = allVals.filter(v => !filterText || v.toLowerCase().includes(filterText));

    // Fila "Seleccionar todo" (solo sin búsqueda activa o con todos visibles)
    const selectAllRow = document.createElement('label');
    selectAllRow.className = 'tx-dp-item tx-dp-select-all-item';
    const allChecked   = filtered.every(v => !activeSet || activeSet.has(v));
    const someChecked  = filtered.some(v => !activeSet || activeSet.has(v));
    selectAllRow.innerHTML = `
      <input type="checkbox" id="txDpCbAll" ${allChecked ? 'checked' : ''}>
      <span><strong>Seleccionar todo</strong></span>
      <span class="tx-dp-count">${filtered.length}</span>`;
    const cbAll = selectAllRow.querySelector('#txDpCbAll');
    if (!allChecked && someChecked) cbAll.indeterminate = true;
    cbAll.addEventListener('change', () => {
      list.querySelectorAll('.tx-dp-value-cb').forEach(i => i.checked = cbAll.checked);
    });
    list.appendChild(selectAllRow);

    // Separador
    const sep = document.createElement('div');
    sep.className = 'tx-dp-sep';
    list.appendChild(sep);

    // Items individuales
    filtered.forEach(val => {
      const checked = !activeSet || activeSet.has(val);
      const item = document.createElement('label');
      item.className = 'tx-dp-item';
      item.innerHTML = `
        <input type="checkbox" class="tx-dp-value-cb" value="${escHtml(val)}" ${checked ? 'checked' : ''}>
        <span title="${escHtml(val)}">${escHtml(val)}</span>`;
      list.appendChild(item);
    });

    // Actualizar estado del "Seleccionar todo" cuando cambia un item
    list.addEventListener('change', e => {
      if (e.target.classList.contains('tx-dp-value-cb')) {
        const cbs      = [...list.querySelectorAll('.tx-dp-value-cb')];
        const total    = cbs.length;
        const selected = cbs.filter(i => i.checked).length;
        const cbA      = list.querySelector('#txDpCbAll');
        if (cbA) {
          cbA.checked       = selected === total;
          cbA.indeterminate = selected > 0 && selected < total;
        }
      }
    });
  }

  renderOptions('');

  // ── Búsqueda ──
  panel.querySelector('#txDpSearch').addEventListener('input', e => {
    renderOptions(e.target.value.toLowerCase());
  });

  // ── Aplicar ──
  panel.querySelector('#txDpApply').addEventListener('click', () => {
    const checked = [...list.querySelectorAll('.tx-dp-value-cb:checked')].map(i => i.value);
    if (checked.length === allVals.length) {
      delete txColFilters[colKey];
    } else if (checked.length === 0) {
      // Ninguno seleccionado: sin cambio (o limpiar)
      delete txColFilters[colKey];
    } else {
      txColFilters[colKey] = new Set(checked);
    }
    closeTxDropdown();
    refreshTxTable();
  });

  // ── Limpiar este filtro ──
  panel.querySelector('#txDpClear').addEventListener('click', () => {
    delete txColFilters[colKey];
    closeTxDropdown();
    refreshTxTable();
  });

  // ── Posicionamiento viewport-relativo (position:fixed → sin scrollY) ──
  const rect   = anchorBtn.getBoundingClientRect();
  const panelW = 280;
  const left   = Math.min(rect.left, window.innerWidth - panelW - 8);
  panel.style.top  = `${rect.bottom + 4}px`;
  panel.style.left = `${left}px`;

  // Ajustar si se sale por abajo
  requestAnimationFrame(() => {
    const ph = panel.offsetHeight;
    if (rect.bottom + ph > window.innerHeight - 8) {
      panel.style.top = `${rect.top - ph - 4}px`;
    }
  });

  // Cerrar al hacer clic fuera
  setTimeout(() => document.addEventListener('click', outsideTxClick), 0);
}

function outsideTxClick(e) {
  const panel = document.getElementById('txDropdownPanel');
  if (panel && !panel.contains(e.target)) closeTxDropdown();
}

function closeTxDropdown() {
  const panel = document.getElementById('txDropdownPanel');
  if (panel) panel.remove();
  txOpenCol = null;
  document.removeEventListener('click', outsideTxClick);
}

/** Limpia todos los filtros de columna activos (incluye año) y resetea sort */
function clearAllTxFilters() {
  txColFilters = {};
  txSortCol    = null;
  txSortDir    = null;
  // Reset botones rápidos
  ['Ingreso','Egreso'].forEach(t => {
    const b = document.getElementById(`qf${t}`);
    if (b) b.classList.remove('active');
  });
  updateSortIcons();
  refreshTxTable();
  syncYearBtnState();
}

// ══════════════════════════════════════════════════════
// 5A. FILTRO POR AÑO
// ══════════════════════════════════════════════════════

function buildYearFilter(rows) {
  const years = [...new Set(rows.map(r => r.year))].sort((a,b) => a - b);
  const sel   = document.getElementById('yearFilter');
  while (sel.options.length > 1) sel.remove(1);
  for (const y of years) {
    const opt = document.createElement('option');
    opt.value = y; opt.textContent = y;
    sel.appendChild(opt);
  }
  sel.value = 'all';  // mostrar todo por defecto
  sel.onchange = () => applyYearFilter(sel.value);
}

function applyYearFilter(val) {
  yearRows     = val === 'all' ? allRows : allRows.filter(r => r.year === parseInt(val, 10));
  filteredRows = yearRows;
  buildMonthFilter(yearRows);
  document.getElementById('monthFilter').value = 'all';
  renderKPIs(filteredRows);
  renderBarChart(filteredRows);
  renderDonutChart(filteredRows);
  renderTipoCharts(filteredRows);
  renderCategoryTable(filteredRows);
  renderTxTable(filteredRows);
  const label = val === 'all' ? 'Todos los años' : `Año ${val}`;
  document.getElementById('dashSubtitle').textContent = label;
}

// ══════════════════════════════════════════════════════
// 5B. FILTRO POR MES
// ══════════════════════════════════════════════════════

function buildMonthFilter(rows) {
  const meses = [...new Set(rows.map(r=>r.mes))].sort((a,b)=>a-b);
  const sel   = document.getElementById('monthFilter');
  while (sel.options.length > 1) sel.remove(1);
  for (const m of meses) {
    const opt = document.createElement('option');
    opt.value = m; opt.textContent = MONTHS_LONG[m];
    sel.appendChild(opt);
  }
  sel.onchange = () => applyMonthFilter(sel.value);
}

function applyMonthFilter(val) {
  filteredRows = val==='all' ? yearRows : yearRows.filter(r=>r.mes===parseInt(val,10));
  renderKPIs(filteredRows);
  renderBarChart(filteredRows);
  renderDonutChart(filteredRows);
  renderTipoCharts(filteredRows);
  renderCategoryTable(filteredRows);
  renderTxTable(filteredRows);
  const yearSel = document.getElementById('yearFilter').value;
  const yearLabel = yearSel === 'all' ? 'Todos los años' : `Año ${yearSel}`;
  const label = val==='all' ? yearLabel : `${MONTHS_LONG[parseInt(val,10)]} — ${yearLabel}`;
  document.getElementById('dashSubtitle').textContent = label;
}

// ══════════════════════════════════════════════════════
// 6. EXPORTAR CSV
// ══════════════════════════════════════════════════════

function exportCSV() {
  const header = ['Fecha','Tipo','Categoría/Tipo','Proveedor/Cliente','Monto'];
  const lines  = [header.join(',')];
  for (const r of filteredRows) {
    lines.push([
      formatDate(r.fecha), r.tipo_registro,
      `"${r.tipo}"`, `"${r.subcategoria}"`, r.monto.toFixed(2),
    ].join(','));
  }
  const blob = new Blob([lines.join('\n')], {type:'text/csv;charset=utf-8;'});
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a');
  a.href=url; a.download=`findash_${detectedYear}.csv`;
  document.body.appendChild(a); a.click();
  document.body.removeChild(a); URL.revokeObjectURL(url);
}

// ══════════════════════════════════════════════════════
// 7. TEMA DARK / LIGHT
// ══════════════════════════════════════════════════════

function toggleTheme() {
  const html   = document.documentElement;
  const isDark = html.getAttribute('data-theme')==='dark';
  html.setAttribute('data-theme', isDark?'light':'dark');
  document.getElementById('iconSun').classList.toggle('hidden',!isDark);
  document.getElementById('iconMoon').classList.toggle('hidden',isDark);
  if (allRows.length > 0) {
    renderBarChart(filteredRows);
    renderDonutChart(filteredRows);
    renderTipoCharts(filteredRows);
  }
}

// ══════════════════════════════════════════════════════
// 8. RESET
// ══════════════════════════════════════════════════════

function resetApp() {
  allRows=[]; yearRows=[]; filteredRows=[]; allYears=[];
  txCurrentRows=[]; txColFilters={}; txOpenCol=null;
  closeTxDropdown();
  document.getElementById('dashboardSection').classList.add('hidden');
  document.getElementById('yearFilterWrapper').classList.add('hidden');
  document.getElementById('monthFilterWrapper').classList.add('hidden');
  document.getElementById('exportCsvBtn').classList.add('hidden');
  document.getElementById('uploadSection').classList.remove('hidden');
  document.getElementById('yearFilter').value='all';
  document.getElementById('monthFilter').value='all';
  document.getElementById('fileInput').value='';
  [barChartInst,donutChartInst].forEach(c=>{if(c){c.destroy();}});
  barChartInst=donutChartInst=null;
  // Destruir charts de tipo
  ['tipoIngChart','tipoEgrChart'].forEach(id=>{
    const c = Chart.getChart(id);
    if(c) c.destroy();
  });
  hideError();
  window.scrollTo({top:0,behavior:'smooth'});
}

// ══════════════════════════════════════════════════════
// 9. HELPERS UI
// ══════════════════════════════════════════════════════

function showLoading() {
  document.getElementById('uploadSection').classList.add('hidden');
  document.getElementById('loadingSection').classList.remove('hidden');
}
function hideLoading() { document.getElementById('loadingSection').classList.add('hidden'); }
function showError(msg) {
  document.getElementById('errorText').textContent = msg;
  document.getElementById('errorMsg').classList.remove('hidden');
  document.getElementById('uploadSection').classList.remove('hidden');
  hideLoading();
}
function hideError() { document.getElementById('errorMsg').classList.add('hidden'); }

// ══════════════════════════════════════════════════════
// 10. HELPERS DE DATOS
// ══════════════════════════════════════════════════════

function agrupar(rows, key) {
  return rows.reduce((acc,r)=>{ acc[r[key]]=(acc[r[key]]||0)+r.monto; return acc; },{});
}
function contar(rows, key) {
  return rows.reduce((acc,r)=>{ acc[r[key]]=(acc[r[key]]||0)+1; return acc; },{});
}

// ── Parsear fecha ──
function parseDate(val) {
  if (val == null || val === '') return null;

  // Ya es Date (SheetJS con cellDates:true)
  if (val instanceof Date) return isNaN(val.getTime()) ? null : val;

  // Número → serial de Excel (ej: 42739 = 4-Ene-2017)
  if (typeof val === 'number') {
    // Excel serial: días desde 1-Ene-1900, con bug de año bisiesto 1900
    // Equivalente JS: (serial - 25569) días desde 1-Ene-1970 en UTC
    if (val < 1 || val > 2958465) return null; // rango sensato (1900–9999)
    const ms  = Math.round((val - 25569) * 86400 * 1000);
    const utc = new Date(ms);
    // Convertir a fecha local (evitar off-by-one por timezone)
    return new Date(utc.getUTCFullYear(), utc.getUTCMonth(), utc.getUTCDate());
  }

  // String
  const str = String(val).trim();
  if (!str) return null;

  // ISO y otros formatos parseables por el motor JS
  const direct = new Date(str);
  if (!isNaN(direct.getTime())) {
    return new Date(direct.getFullYear(), direct.getMonth(), direct.getDate());
  }

  // DD/MM/YYYY o DD-MM-YYYY
  const ddmm = str.match(/^(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{4})$/);
  if (ddmm) {
    const d = new Date(+ddmm[3], +ddmm[2]-1, +ddmm[1]);
    if (!isNaN(d.getTime())) return d;
  }

  // YYYY/MM/DD
  const yyyymm = str.match(/^(\d{4})[\/\-\.](\d{1,2})[\/\-\.](\d{1,2})$/);
  if (yyyymm) {
    const d = new Date(+yyyymm[1], +yyyymm[2]-1, +yyyymm[3]);
    if (!isNaN(d.getTime())) return d;
  }

  // Serial de Excel guardado como string numérico (ej: "42739" → 4-Ene-2017)
  // Rango: 1 (1-Ene-1900) a 2958465 (31-Dic-9999)
  if (/^\d{4,6}$/.test(str)) {
    const serial = parseInt(str, 10);
    if (serial > 1 && serial < 2958465) {
      const ms  = Math.round((serial - 25569) * 86400 * 1000);
      const utc = new Date(ms);
      return new Date(utc.getUTCFullYear(), utc.getUTCMonth(), utc.getUTCDate());
    }
  }

  return null;
}

function parseMonto(val) {
  if (val===null||val===undefined||val==='') return NaN;
  if (typeof val==='number') return val;
  const c = String(val).replace(/[\$\s,]/g,'').replace(/[()]/g,m=>m==='('?'-':'');
  return parseFloat(c);
}

function capitalizar(s) {
  if (!s||s==='—') return s;
  return s.charAt(0).toUpperCase()+s.slice(1);
}

function formatMoney(n) {
  return new Intl.NumberFormat('es-MX',{style:'currency',currency:'MXN',minimumFractionDigits:0,maximumFractionDigits:0}).format(n);
}
function formatMoneyShort(n) {
  if(Math.abs(n)>=1_000_000)return`$${(n/1_000_000).toFixed(1)}M`;
  if(Math.abs(n)>=1_000)return`$${(n/1_000).toFixed(0)}K`;
  return`$${n}`;
}
function formatDate(d) {
  if (!d) return 'Sin fecha';
  return d.toLocaleDateString('es-MX',{day:'2-digit',month:'short',year:'numeric'});
}
function escHtml(s) {
  if(!s)return'';
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

// Opciones base para tooltips de Chart.js
// horizontal=true para barras horizontales (indexAxis:'y')
function tooltipDefaults(isDark, labelFn, isDonut=false, horizontal=false) {
  return {
    backgroundColor: isDark?'#1e293b':'#fff',
    titleColor:      isDark?'#f1f5f9':'#0f172a',
    bodyColor:       isDark?'#94a3b8':'#475569',
    borderColor:     isDark?'rgba(255,255,255,0.1)':'rgba(0,0,0,0.1)',
    borderWidth:1, padding:12, cornerRadius:10,
    callbacks: {
      label: isDonut
        ? ctx => labelFn(ctx.parsed, ctx)
        // Acceso directo al array de datos — más confiable que ctx.parsed en ambos ejes
        // Para barras horizontales (sin label de dataset) se usa el label de categoría
        : ctx => {
            const val  = ctx.dataset.data[ctx.dataIndex];
            const lbl  = ctx.dataset.label || ctx.label || '';
            return ` ${lbl ? lbl + ': ' : ''}${labelFn(val)}`;
          },
    },
  };
}


// ══════════════════════════════════════════════════════════════════
// MÓDULO 2 — IMPORTACIÓN ARCHIVO 2
// Pipeline: Ingesta → Transformación → Consolidación → Dashboards
// ══════════════════════════════════════════════════════════════════

'use strict';

// ─── ESTADO AISLADO (no se mezcla con allRows del Dashboard) ─────
let rawDataFile2    = {};   // { sheetName: [[rawRow]] }
let processedData2  = {};   // { sheetName: [normalizedRow] }
let finalDataset2   = [];   // todos los rows combinados
let file2FileName   = '';
let sheetChartInsts = {};   // { sheetName: { bar, donut } }
let activeSheetTab2 = null;
let sheetFilters2   = {};   // { sheetName: { year, mes, tipo, search, sortCol, sortDir } }
let sheetIdxToName2 = {};   // { idx: sheetName } para onclick handlers seguros

// ─── CONFIGURACIÓN DE HOJAS (Sistema de Mapeo Dinámico) ──────────
// Para agregar nuevas hojas en el futuro, solo agrega una entrada aquí.
const SHEET_PATTERNS = [
  { match: /BBVA.+MXN.+CHEQ/i,       banco: 'BBVA',        tipo: 'Cheques MXN',     moneda: 'MXN' },
  { match: /BBVA.+MXN.+CONCENT/i,    banco: 'BBVA',        tipo: 'Concentradora',   moneda: 'MXN' },
  { match: /BBVA.+USD/i,             banco: 'BBVA',        tipo: 'Cheques USD',     moneda: 'USD' },
  { match: /BBVA.+MXN.+CR[EÉ]D/i,   banco: 'BBVA',        tipo: 'Crédito MXN',     moneda: 'MXN' },
  { match: /BBVA/i,                  banco: 'BBVA',        tipo: 'Cuenta MXN',      moneda: 'MXN' },
  { match: /MONEX.+FONDO/i,          banco: 'Monex',       tipo: 'Fondo Ahorro',    moneda: 'MXN' },
  { match: /MONEX.+USD/i,            banco: 'Monex',       tipo: 'Cheques USD',     moneda: 'USD' },
  { match: /MONEX.+MXN/i,            banco: 'Monex',       tipo: 'Cheques MXN',     moneda: 'MXN' },
  { match: /MONEX/i,                 banco: 'Monex',       tipo: 'Cuenta',          moneda: 'MXN' },
  { match: /CLARA/i,                 banco: 'Clara',       tipo: 'Tarj. Crédito',   moneda: 'MXN' },
  { match: /KAPITAL.+FLEX/i,         banco: 'Kapital',     tipo: 'Flex',            moneda: 'MXN' },
  { match: /KAPITAL.+FACTOR/i,       banco: 'Kapital',     tipo: 'Factoraje',       moneda: 'MXN' },
  { match: /KAPITAL.+TARJ/i,         banco: 'Kapital',     tipo: 'Tarj. Crédito',   moneda: 'MXN' },
  { match: /KAPITAL/i,               banco: 'Kapital',     tipo: 'Cheques MXN',     moneda: 'MXN' },
  { match: /KONF[IÍ]O?.+TARJ/i,     banco: 'Konfio',      tipo: 'Tarj. Crédito',   moneda: 'MXN' },
  { match: /KONF[IÍ]O?/i,           banco: 'Konfio',      tipo: 'Crédito',         moneda: 'MXN' },
  { match: /XEPELIN/i,               banco: 'Xepelin',     tipo: 'Crédito',         moneda: 'MXN' },
  { match: /INBURSA/i,               banco: 'Inbursa',     tipo: 'Ahorro MXN',      moneda: 'MXN' },
  { match: /TEXAS.+BANK/i,           banco: 'Texas Bank',  tipo: 'Cheques USD',     moneda: 'USD' },
  { match: /INTERCAM.+USD.+INV/i,   banco: 'Intercam',    tipo: 'Inversión USD',   moneda: 'USD' },
  { match: /INTERCAM.+USD/i,         banco: 'Intercam',    tipo: 'Cheques USD',     moneda: 'USD' },
  { match: /INTERCAM.+MXN.+INV/i,   banco: 'Intercam',    tipo: 'Inversión MXN',   moneda: 'MXN' },
  { match: /INTERCAM.+MXN/i,         banco: 'Intercam',    tipo: 'Cheques MXN',     moneda: 'MXN' },
  { match: /INTERCAM/i,              banco: 'Intercam',    tipo: 'Cuenta',          moneda: 'MXN' },
  { match: /BANAMEX.+CONCENT/i,     banco: 'Banamex',     tipo: 'Concentradora',   moneda: 'MXN' },
  { match: /BANAMEX/i,               banco: 'Banamex',     tipo: 'Cheques MXN',     moneda: 'MXN' },
  { match: /USD/i,                   banco: '—',           tipo: 'USD',             moneda: 'USD' },
  { match: /MXN/i,                   banco: '—',           tipo: 'MXN',             moneda: 'MXN' },
];

/** Clasifica una hoja por nombre usando SHEET_PATTERNS (dinámico). */
function classifySheet(sheetName) {
  for (const p of SHEET_PATTERNS) {
    if (p.match.test(sheetName)) {
      return { banco: p.banco, tipo_cuenta: p.tipo, moneda: p.moneda };
    }
  }
  const words = sheetName.trim().split(/\s+/);
  const moneda = sheetName.toUpperCase().includes('USD') ? 'USD' : 'MXN';
  return { banco: words[0] || sheetName, tipo_cuenta: 'Cuenta', moneda };
}

// ─── SWITCHING DE TABS PRINCIPAL ─────────────────────────────────
function switchMainTab(targetId, btn) {
  document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('.main-tab-btn').forEach(b => b.classList.remove('active'));
  document.getElementById(targetId).classList.add('active');
  if (btn) btn.classList.add('active');
  window.scrollTo({ top: 0, behavior: 'smooth' });
}

// ─── DRAG & DROP ARCHIVO 2 ────────────────────────────────────────
function import2DragOver(e) {
  e.preventDefault();
  document.getElementById('import2DropZone').classList.add('drag-over');
}
function import2DragLeave(e) {
  e.preventDefault();
  document.getElementById('import2DropZone').classList.remove('drag-over');
}
function import2Drop(e) {
  e.preventDefault();
  document.getElementById('import2DropZone').classList.remove('drag-over');
  const file = e.dataTransfer.files[0];
  if (file) processFile2(file);
}
function import2Change(e) {
  const file = e.target.files[0];
  if (file) processFile2(file);
}
function triggerImport2FileInput() {
  document.getElementById('import2FileInput').click();
}

// ─── ETAPA 1: INGESTA ────────────────────────────────────────────
function processFile2(file) {
  const ext = file.name.split('.').pop().toLowerCase();
  if (!['xlsx', 'xls'].includes(ext)) {
    showImport2Error('El archivo debe ser .xlsx o .xls. Por favor verifica el formato.');
    return;
  }
  hideImport2Error();
  file2FileName = file.name;

  const reader = new FileReader();
  reader.onload = (ev) => {
    try {
      const data = new Uint8Array(ev.target.result);
      const wb   = XLSX.read(data, { type: 'array' });

      // Leer TODAS las hojas en formato RAW (sin transformar)
      rawDataFile2 = {};
      for (const name of wb.SheetNames) {
        rawDataFile2[name] = wsToStrings(wb.Sheets[name]);
      }

      renderSheetsDetected(wb.SheetNames);
      setPipelineStep(1);
    } catch (err) {
      showImport2Error(err.message || 'Error al leer el archivo. Verifica que sea un Excel válido.');
    }
  };
  reader.onerror = () => showImport2Error('No se pudo leer el archivo. Intenta de nuevo.');
  reader.readAsArrayBuffer(file);
}

/** Muestra las hojas detectadas con su clasificación automática. */
function renderSheetsDetected(sheetNames) {
  const card = document.getElementById('import2SheetsDetected');
  card.classList.remove('hidden');
  document.getElementById('import2FileName').textContent   = file2FileName;
  document.getElementById('import2SheetsCount').textContent = `${sheetNames.length} hojas detectadas · Estado: aislado ✓`;

  const container = document.getElementById('import2SheetsPills');
  container.innerHTML = '';

  for (const name of sheetNames) {
    const meta = classifySheet(name);
    const pill = document.createElement('div');
    pill.className = 'sheet-pill';
    pill.innerHTML = `
      <div class="sheet-pill-name" title="${escHtml(name)}">${escHtml(name)}</div>
      <div class="sheet-pill-meta">
        <span class="sheet-pill-banco">${escHtml(meta.banco)}</span>
        <span class="sheet-pill-tipo">${escHtml(meta.tipo_cuenta)}</span>
        <span class="sheet-pill-moneda sheet-pill-moneda-${meta.moneda}">${meta.moneda}</span>
      </div>`;
    container.appendChild(pill);
  }
}

// ─── ETAPA 2: TRANSFORMACIÓN ─────────────────────────────────────

/** Detecta el layout de columnas usando las primeras filas de datos.
 *  Estrategia verificada con el archivo real SMTO:
 *  - 8-col: Date, Tipo, Folio, Factura, Concepto, Ingreso, Egreso, Saldo
 *  - 7-col: Date, Tipo, Folio, Concepto, Ingreso, Egreso, Saldo
 *  - 6-col: Date, Folio(largeInt), Concepto, Ingreso, Egreso, Saldo  (Monex Fondo)
 *  Devuelve también facturaCol para el campo de referencia/factura.
 */
function detectLayoutPositional(sampleRows) {
  const n = sampleRows.length;
  if (n === 0) return { ingCol: 5, egrCol: 6, descCols: [1, 4], facturaCol: 3 };

  // Caso especial: col 1 = folio de 8+ dígitos (Monex Fondo Ahorro)
  const col1LargeInt = sampleRows.filter(r => {
    if (!r || r[1] == null) return false;
    const v = String(r[1]).trim();
    const num = parseInt(v, 10);
    return !isNaN(num) && v.length >= 8 && Math.abs(num) > 9999999;
  }).length;

  if (col1LargeInt > n * 0.6) {
    // 6-col: col1=folio, col2=concepto, col3=ingreso, col4=egreso
    return { ingCol: 3, egrCol: 4, descCols: [2], facturaCol: 1 };
  }

  // Verificar si col 7 tiene valores numéricos (→ formato 8-col)
  const col7Numeric = sampleRows.filter(r => {
    if (!r || r[7] == null) return false;
    const v = String(r[7]).trim();
    if (!v || /^(x|X|-|NA|na)$/.test(v)) return false;
    return !isNaN(parseMonto(v));
  }).length;

  if (col7Numeric > n * 0.35) {
    // 8-col: col1=Tipo, col2=Folio, col3=Factura, col4=Concepto, col5=Ingreso, col6=Egreso
    return { ingCol: 5, egrCol: 6, descCols: [1, 4], facturaCol: 3 };
  }

  // 7-col: col1=Tipo, col2=Folio/Factura, col3=Concepto, col4=Ingreso, col5=Egreso
  return { ingCol: 4, egrCol: 5, descCols: [1, 3], facturaCol: 2 };
}

/** Extrae filas de datos dados los índices de columnas.
 *  @param {Array}  rawRows       - todas las filas de la hoja
 *  @param {number} startRow      - primera fila de datos (después del header)
 *  @param {object} cols          - { colFecha, colIngreso, colEgreso, colDesc, colTipo, colFactura, colConcepto }
 *  @param {string} sheetName
 *  @param {object} meta          - { banco, tipo_cuenta, moneda }
 *  @param {string} tableName     - nombre de la tabla origen (ej. "Tabla 2023")
 *  @param {number} endRow        - última fila a procesar (exclusiva)
 *  @param {number|null} yearOverride - año de la tabla (detectado del título, o null)
 */
function extractFromCols(rawRows, startRow, cols, sheetName, meta,
                          tableName = '', endRow = null, yearOverride = null) {
  const { colFecha, colIngreso, colEgreso, colDesc = -1, colTipo = -1,
          colFactura = -1, colConcepto = -1 } = cols;
  const limit  = endRow != null ? Math.min(endRow, rawRows.length) : rawRows.length;
  const result = [];

  for (let i = startRow; i < limit; i++) {
    const row = rawRows[i];
    if (!row || !row.some(c => c != null && String(c).trim())) continue;

    // ── Validar fecha ──────────────────────────────────────────────
    const rawFecha = colFecha >= 0 && row[colFecha] != null ? String(row[colFecha]).trim() : '';
    if (!rawFecha) continue;
    const fecha = parseDate(rawFecha);
    if (!fecha || fecha.getFullYear() < 2000 || fecha.getFullYear() > 2035) continue;

    // ── Columnas de texto ──────────────────────────────────────────
    const rawTipo     = colTipo     >= 0 && row[colTipo]     ? String(row[colTipo]).trim()     : '';
    const rawDesc     = colDesc     >= 0 && row[colDesc]     ? String(row[colDesc]).trim()     : '';
    const rawFactura  = colFactura  >= 0 && row[colFactura]  ? String(row[colFactura]).trim()  : '';
    const rawConcepto = colConcepto >= 0 && row[colConcepto] ? String(row[colConcepto]).trim() : '';

    // ── Omitir filas de totales / resúmenes ────────────────────────
    const textCheck = rawDesc || rawTipo || rawConcepto;
    if (/^(total|saldo|resumen|subtotal|suma|balance|promedio|tipo de cambio)/i.test(textCheck)) continue;

    // ── Montos ─────────────────────────────────────────────────────
    const rawIng = colIngreso >= 0 && row[colIngreso] != null ? String(row[colIngreso]) : '';
    const rawEgr = colEgreso  >= 0 && row[colEgreso]  != null ? String(row[colEgreso])  : '';
    const ing = parseMonto(rawIng);
    const egr = parseMonto(rawEgr);

    let monto = 0;
    let tipo_registro = null;

    if (!isNaN(ing) && ing > 0 && (isNaN(egr) || egr === 0)) {
      monto = ing; tipo_registro = 'Ingreso';
    } else if (!isNaN(egr) && egr > 0 && (isNaN(ing) || ing === 0)) {
      monto = egr; tipo_registro = 'Egreso';
    } else if (!isNaN(ing) && ing < 0) {
      monto = Math.abs(ing); tipo_registro = 'Egreso';
    } else if (!isNaN(egr) && egr < 0) {
      monto = Math.abs(egr); tipo_registro = 'Ingreso';
    } else if (!isNaN(ing) && ing !== 0) {
      monto = Math.abs(ing); tipo_registro = ing > 0 ? 'Ingreso' : 'Egreso';
    } else {
      continue;
    }

    if (monto === 0 || !tipo_registro) continue;

    // ── Construir campos de texto limpios ──────────────────────────
    const cleanNA  = v => (v && !/^(NA|na|n\.a\.|-|0)$/i.test(v)) ? v : '';
    const factura  = cleanNA(rawFactura);
    const concepto = rawConcepto || rawDesc || '';

    // descripcion_corta = col Tipo/Desc; descripcion = combinación legible
    const descParts = [rawTipo, concepto].filter(v => cleanNA(v));
    const fullDesc  = descParts.join(' · ') || rawTipo || meta.tipo_cuenta;

    const yearFinal = yearOverride || fecha.getFullYear();

    result.push({
      fecha,
      year:               yearFinal,
      mes:                fecha.getMonth(),
      tipo_registro,
      descripcion:        fullDesc,
      descripcion_corta:  rawTipo || '',
      factura,
      concepto,
      monto,
      ingreso:            tipo_registro === 'Ingreso' ? monto : 0,
      egreso:             tipo_registro === 'Egreso'  ? monto : 0,
      hoja:               sheetName,
      banco:              meta.banco,
      tipo_cuenta:        meta.tipo_cuenta,
      moneda:             meta.moneda,
      tabla_origen:       tableName || '',
      // Compatibilidad con renderers existentes (Módulo 1)
      tipo:               rawTipo || meta.tipo_cuenta,
      categoria:          meta.banco,
      subcategoria:       concepto || fullDesc,
      importe:            monto,
      iva:                0,
      ret:                0,
    });
  }
  return result;
}

// ══════════════════════════════════════════════════════════════════
// DETECCIÓN DE MÚLTIPLES TABLAS POR HOJA
// Cada hoja puede contener N tablas (una por año) con estructura:
//   Título: "Bancomer Cheques SMTO 2023"
//   Header: Fecha | Desc | Factura | Concepto | Ingreso | Egreso
//   Datos:  filas de movimientos
//   Totales: fila de sumas (NO se procesa como movimiento)
// ══════════════════════════════════════════════════════════════════

/** Palabras clave para identificar columnas de encabezado. */
const HDR_FECHA    = ['FECHA','DATE','FECHA OPERACION','FECHA VALOR'];
const HDR_INGRESO  = ['INGRESO','ABONO','CRÉDITO','CREDITO','ABONOS','INGRESOS'];
const HDR_EGRESO   = ['EGRESO','CARGO','DÉBITO','DEBITO','CARGOS','EGRESOS'];
const HDR_CONCEPTO = ['CONCEPTO','DESCRIPCIÓN','DESCRIPCION','DETALLE','CONCEPTO/DESCRIPCIÓN'];
const HDR_FACTURA  = ['FACTURA','FOLIO','REF','REFERENCIA','NO. FACTURA','NÚMERO','NUMERO'];
const HDR_DESC     = ['DESC','TIPO','DESCRIPCION CORTA','TIPO MOVIMIENTO'];

/** Nombres de meses en español para detectar tablas resumen. */
const SUMMARY_MONTHS = [
  'ENERO','FEBRERO','MARZO','ABRIL','MAYO','JUNIO',
  'JULIO','AGOSTO','SEPTIEMBRE','OCTUBRE','NOVIEMBRE','DICIEMBRE',
  'ENE','FEB','MAR','ABR','MAY','JUN','JUL','AGO','SEP','OCT','NOV','DIC',
];

/** ¿Es una fila de resumen mensual (NO procesar)? */
function isSummaryTableHeader(row) {
  const vals = row.map(c => String(c || '').toUpperCase().trim());
  return vals.filter(v => SUMMARY_MONTHS.includes(v)).length >= 3;
}

/** ¿Parece un encabezado de tabla de movimientos? */
function looksLikeDataTableHeader(row) {
  const vals = row.map(c => String(c || '').toUpperCase().trim());
  const hasFecha = vals.some(v => HDR_FECHA.includes(v) || v.includes('FECHA'));
  const hasIng   = vals.some(v => HDR_INGRESO.includes(v));
  const hasEgr   = vals.some(v => HDR_EGRESO.includes(v));
  return hasFecha && (hasIng || hasEgr);
}

/** Busca un año (2020-2035) en las filas anteriores al header de tabla. */
function extractTableYear(rawRows, headerRowIdx) {
  for (let i = Math.max(0, headerRowIdx - 6); i < headerRowIdx; i++) {
    const row = rawRows[i];
    if (!row) continue;
    for (const cell of row) {
      const m = String(cell || '').match(/\b(202[0-9]|203[0-5])\b/);
      if (m) return parseInt(m[1], 10);
    }
  }
  return null;
}

/**
 * Detecta y describe TODAS las tablas de movimientos en una hoja.
 * Ignora tablas de resumen (columnas = meses).
 * @returns {Array<{headerRowIdx, dataStartIdx, dataEndIdx, yearDetected,
 *                  colFecha, colDesc, colFactura, colConcepto, colIngreso, colEgreso, tableName}>}
 */
function detectTablesInSheet(rawRows) {
  const tables = [];
  let i = 0;

  while (i < rawRows.length) {
    const row = rawRows[i];
    if (!row || !row.some(c => c != null && String(c).trim())) { i++; continue; }

    if (looksLikeDataTableHeader(row) && !isSummaryTableHeader(row)) {
      const vals = row.map(c => String(c || '').toUpperCase().trim());

      // ── Mapear columnas desde el header ───────────────────────
      const findCol = (keywords) => vals.findIndex(v => keywords.includes(v));

      const colFecha   = (() => {
        let idx = vals.findIndex(v => HDR_FECHA.includes(v));
        if (idx < 0) idx = vals.findIndex(v => v.includes('FECHA'));
        return idx >= 0 ? idx : 0;
      })();
      const colIngreso = findCol(HDR_INGRESO);
      const colEgreso  = findCol(HDR_EGRESO);

      // DESC y CONCEPTO: cuidado de no asignar el mismo índice
      let colDesc     = findCol(HDR_DESC);
      let colConcepto = findCol(HDR_CONCEPTO);
      if (colConcepto === colDesc && colConcepto >= 0) colConcepto = -1;

      // Si no hay DESC separado, usar el primer campo textual
      if (colDesc < 0 && colConcepto >= 0) {
        // busca otra columna de texto que no sea fecha/ingreso/egreso
        colDesc = vals.findIndex((v, idx) =>
          idx !== colFecha && idx !== colIngreso && idx !== colEgreso &&
          idx !== colConcepto && v && v.length > 0 &&
          !['$','SALDO','BALANCE','TIPO DE CAMBIO'].includes(v)
        );
      }

      let colFactura = findCol(HDR_FACTURA);

      const yearDetected = extractTableYear(rawRows, i);
      const headerRowIdx = i;
      const dataStartIdx = i + 1;

      // ── Buscar fin de tabla ────────────────────────────────────
      let emptyRun   = 0;
      let dataEndIdx = rawRows.length;

      for (let j = dataStartIdx; j < rawRows.length; j++) {
        const jrow = rawRows[j];
        if (!jrow || !jrow.some(c => c != null && String(c).trim())) {
          emptyRun++;
          if (emptyRun >= 3) { dataEndIdx = j - emptyRun + 1; break; }
        } else {
          emptyRun = 0;
          // Nuevo encabezado de tabla → termina la actual
          if (j > headerRowIdx + 3 && looksLikeDataTableHeader(jrow) && !isSummaryTableHeader(jrow)) {
            dataEndIdx = j;
            break;
          }
        }
      }

      tables.push({
        headerRowIdx,
        dataStartIdx,
        dataEndIdx,
        yearDetected,
        colFecha,
        colDesc,
        colFactura,
        colConcepto,
        colIngreso:  colIngreso >= 0 ? colIngreso : 5,
        colEgreso:   colEgreso  >= 0 ? colEgreso  : 6,
        tableName:   yearDetected ? `Tabla ${yearDetected}` : `Tabla ${tables.length + 1}`,
      });

      i = dataEndIdx;
    } else {
      i++;
    }
  }

  return tables;
}

/**
 * Normaliza una hoja completa.
 * Primero intenta detectar MÚLTIPLES TABLAS con encabezados explícitos.
 * Si no encuentra ninguna, cae al modo posicional (hojas sin headers).
 */
function normalizeSheetRows(rawRows, sheetName, meta) {
  if (!rawRows || rawRows.length === 0) return [];

  // ── INTENTO 1: Detección multi-tabla con encabezados explícitos ──
  const tables = detectTablesInSheet(rawRows);

  if (tables.length > 0) {
    const allExtracted = [];
    for (const tbl of tables) {
      const rows = extractFromCols(
        rawRows,
        tbl.dataStartIdx,
        {
          colFecha:    tbl.colFecha,
          colIngreso:  tbl.colIngreso,
          colEgreso:   tbl.colEgreso,
          colDesc:     tbl.colDesc     >= 0 ? tbl.colDesc     : tbl.colConcepto,
          colTipo:     tbl.colDesc     >= 0 ? tbl.colDesc     : -1,
          colFactura:  tbl.colFactura  >= 0 ? tbl.colFactura  : -1,
          colConcepto: tbl.colConcepto >= 0 ? tbl.colConcepto : tbl.colDesc,
        },
        sheetName,
        meta,
        tbl.tableName,
        tbl.dataEndIdx,
        tbl.yearDetected,
      );
      allExtracted.push(...rows);
    }
    return allExtracted;
  }

  // ── INTENTO 2: Detección posicional (hojas sin headers explícitos) ──
  let startRow = -1;
  for (let i = 0; i < Math.min(rawRows.length, 300); i++) {
    const row = rawRows[i];
    if (!row || !row[0]) continue;
    const date = parseDate(row[0]);
    if (date && date.getFullYear() >= 2000 && date.getFullYear() <= 2035) {
      startRow = i;
      break;
    }
  }
  if (startRow < 0) return [];

  const sampleRows = rawRows.slice(startRow, startRow + 10).filter(r => r && r[0]);
  const layout     = detectLayoutPositional(sampleRows);

  return extractFromCols(
    rawRows,
    startRow,
    {
      colFecha:    0,
      colIngreso:  layout.ingCol,
      colEgreso:   layout.egrCol,
      colDesc:     layout.descCols[layout.descCols.length - 1],
      colTipo:     layout.descCols.length > 1 ? layout.descCols[0] : -1,
      colFactura:  layout.facturaCol,
      colConcepto: layout.descCols[layout.descCols.length - 1],
    },
    sheetName,
    meta,
    '',
    rawRows.length,
    null,
  );
}

/** Ejecuta la transformación de todas las hojas. */
function runTransformation() {
  processedData2 = {};
  let totalRows  = 0;

  for (const [sheetName, rawRows] of Object.entries(rawDataFile2)) {
    const meta       = classifySheet(sheetName);
    const normalized = normalizeSheetRows(rawRows, sheetName, meta);
    processedData2[sheetName] = normalized;
    totalRows += normalized.length;
  }

  renderTransformSummary(totalRows);
  renderTransformPreview();

  document.getElementById('stage2').classList.remove('hidden');
  setPipelineStep(2);
  document.getElementById('stage2').scrollIntoView({ behavior: 'smooth', block: 'start' });
}

/** Renderiza las tarjetas de resumen por hoja. */
function renderTransformSummary(totalRows) {
  const grid = document.getElementById('transformSummary');
  grid.innerHTML = '';

  for (const [name, rows] of Object.entries(processedData2)) {
    const meta = classifySheet(name);
    const ing  = rows.filter(r => r.tipo_registro === 'Ingreso').reduce((s, r) => s + r.monto, 0);
    const egr  = rows.filter(r => r.tipo_registro === 'Egreso').reduce((s, r) => s + r.monto, 0);
    const ok   = rows.length > 0 ? '✓' : '—';

    const card = document.createElement('div');
    card.className = 'transform-sheet-card';
    card.innerHTML = `
      <div class="tsc-header">
        <div class="tsc-name" title="${escHtml(name)}">${escHtml(name)}</div>
        <span class="tsc-badge tsc-badge-${meta.moneda}">${meta.moneda}</span>
      </div>
      <div class="tsc-meta">${escHtml(meta.banco)} · ${escHtml(meta.tipo_cuenta)}</div>
      <div class="tsc-stats">
        <div class="tsc-stat">
          <span class="tsc-stat-label">Registros</span>
          <span class="tsc-stat-val">${rows.length > 0 ? rows.length.toLocaleString('es-MX') : '<span style="color:var(--text-3)">0 — verificar</span>'}</span>
        </div>
        <div class="tsc-stat">
          <span class="tsc-stat-label tsc-ing">Ingresos</span>
          <span class="tsc-stat-val tsc-ing">${formatMoneyShort(ing)}</span>
        </div>
        <div class="tsc-stat">
          <span class="tsc-stat-label tsc-egr">Egresos</span>
          <span class="tsc-stat-val tsc-egr">${formatMoneyShort(egr)}</span>
        </div>
      </div>`;
    grid.appendChild(card);
  }

  document.getElementById('transformPreviewBadge').textContent =
    `${totalRows.toLocaleString('es-MX')} registros totales normalizados`;
}

/** Muestra las primeras filas normalizadas como vista previa. */
function renderTransformPreview() {
  const sample = [];
  for (const rows of Object.values(processedData2)) {
    sample.push(...rows.slice(0, 3));
    if (sample.length >= 18) break;
  }

  const tbody = document.getElementById('transformPreviewBody');
  if (!tbody) return;
  tbody.innerHTML = '';

  for (const r of sample.slice(0, 18)) {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td class="td-date" style="max-width:120px;overflow:hidden;text-overflow:ellipsis;" title="${escHtml(r.hoja)}">${escHtml(r.hoja)}</td>
      <td class="td-date">${escHtml(r.banco)}</td>
      <td><span class="sheet-pill-moneda sheet-pill-moneda-${r.moneda}">${r.moneda}</span></td>
      <td class="td-date">${escHtml(formatDate(r.fecha))}</td>
      <td class="td-nombre" style="max-width:200px;overflow:hidden;text-overflow:ellipsis;">${escHtml(r.descripcion)}</td>
      <td class="text-right td-num">${formatMoney(r.monto)}</td>
      <td><span class="type-badge ${r.tipo_registro === 'Ingreso' ? 'type-income' : 'type-expense'}">${r.tipo_registro}</span></td>`;
    tbody.appendChild(tr);
  }
}

// ─── ETAPA 3: CONSOLIDACIÓN ────────────────────────────────────────

/** Combina todos los datos normalizados y muestra la vista consolidada. */
function runConsolidation() {
  finalDataset2 = [];
  for (const rows of Object.values(processedData2)) {
    finalDataset2.push(...rows);
  }
  // Ordenar por fecha ascendente
  finalDataset2.sort((a, b) => {
    if (!a.fecha && !b.fecha) return 0;
    if (!a.fecha) return 1;
    if (!b.fecha) return -1;
    return a.fecha - b.fecha;
  });

  renderConsolidationKPIs();
  renderConsolidationTable();

  document.getElementById('stage3').classList.remove('hidden');
  setPipelineStep(3);
  document.getElementById('stage3').scrollIntoView({ behavior: 'smooth', block: 'start' });
}

/** KPIs del consolidado. */
function renderConsolidationKPIs() {
  const rows   = finalDataset2;
  const ing    = rows.filter(r => r.tipo_registro === 'Ingreso');
  const egr    = rows.filter(r => r.tipo_registro === 'Egreso');
  const totalI = ing.reduce((s, r) => s + r.monto, 0);
  const totalE = egr.reduce((s, r) => s + r.monto, 0);
  const balance  = totalI - totalE;
  const cuentas  = new Set(rows.map(r => r.hoja)).size;

  document.getElementById('consolidationKPIs').innerHTML = `
    <div class="kpi-grid">
      <div class="kpi-card kpi-income">
        <div class="kpi-label">
          <svg width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><line x1="12" y1="19" x2="12" y2="5"/><polyline points="5 12 12 5 19 12"/></svg>
          Total Ingresos
        </div>
        <div class="kpi-value">${formatMoney(totalI)}</div>
        <div class="kpi-detail">${ing.length.toLocaleString('es-MX')} movimientos</div>
      </div>
      <div class="kpi-card kpi-expense">
        <div class="kpi-label">
          <svg width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><line x1="12" y1="5" x2="12" y2="19"/><polyline points="19 12 12 19 5 12"/></svg>
          Total Egresos
        </div>
        <div class="kpi-value">${formatMoney(totalE)}</div>
        <div class="kpi-detail">${egr.length.toLocaleString('es-MX')} movimientos</div>
      </div>
      <div class="kpi-card kpi-balance">
        <div class="kpi-label">
          <svg width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><line x1="12" y1="1" x2="12" y2="23"/><path d="M17 5H9.5a3.5 3.5 0 0 0 0 7h5a3.5 3.5 0 0 1 0 7H6"/></svg>
          Balance Neto
        </div>
        <div class="kpi-value" style="color:${balance >= 0 ? 'var(--income)' : 'var(--expense)'}">
          ${formatMoney(balance)}
        </div>
        <div class="kpi-detail">${balance >= 0 ? '✓ Positivo' : '⚠ Negativo'}</div>
      </div>
      <div class="kpi-card kpi-rate">
        <div class="kpi-label">
          <svg width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><rect x="3" y="3" width="7" height="7"/><rect x="14" y="3" width="7" height="7"/><rect x="14" y="14" width="7" height="7"/><rect x="3" y="14" width="7" height="7"/></svg>
          Total Movimientos
        </div>
        <div class="kpi-value">${rows.length.toLocaleString('es-MX')}</div>
        <div class="kpi-detail">${cuentas} cuentas bancarias</div>
      </div>
    </div>`;
}

/** Tabla de resumen por institución. */
function renderConsolidationTable() {
  const rows   = finalDataset2;
  const bancos = [...new Set(rows.map(r => r.banco))].sort();
  const tbody  = document.getElementById('consolidationTableBody');
  if (!tbody) return;
  tbody.innerHTML = '';

  const badge = document.getElementById('consolidationBadge');
  if (badge) badge.textContent = `${bancos.length} instituciones`;

  let sumI = 0, sumE = 0;

  for (const banco of bancos) {
    const bRows = rows.filter(r => r.banco === banco);
    const i     = bRows.filter(r => r.tipo_registro === 'Ingreso').reduce((s, r) => s + r.monto, 0);
    const e     = bRows.filter(r => r.tipo_registro === 'Egreso').reduce((s, r) => s + r.monto, 0);
    const b     = i - e;
    sumI += i; sumE += e;

    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td><strong>${escHtml(banco)}</strong></td>
      <td class="text-right td-total-inc">+${formatMoney(i)}</td>
      <td class="text-right td-total-egr">-${formatMoney(e)}</td>
      <td class="text-right" style="color:${b >= 0 ? 'var(--income)' : 'var(--expense)'}; font-weight:600">
        ${formatMoney(b)}
      </td>
      <td class="text-right">${bRows.length.toLocaleString('es-MX')}</td>`;
    tbody.appendChild(tr);
  }

  const trTot = document.createElement('tr');
  trTot.className = 'tx-totals-row';
  trTot.innerHTML = `
    <td><strong>TOTAL GENERAL</strong></td>
    <td class="text-right tx-foot-num td-total-inc">+${formatMoney(sumI)}</td>
    <td class="text-right tx-foot-num td-total-egr">-${formatMoney(sumE)}</td>
    <td class="text-right tx-foot-num tx-foot-total">${formatMoney(sumI - sumE)}</td>
    <td class="text-right tx-foot-num">${rows.length.toLocaleString('es-MX')}</td>`;
  tbody.appendChild(trTot);
}

// ─── ETAPA 3: EXPORTAR EXCEL ──────────────────────────────────────

/** Genera y descarga el Excel consolidado con múltiples hojas. */
function exportConsolidatedExcel() {
  if (!finalDataset2.length) {
    showImport2Error('No hay datos consolidados. Completa la etapa de consolidación primero.');
    return;
  }

  const wb = XLSX.utils.book_new();

  // ── Hoja 1: Concentrado (todos los movimientos) ──
  const concentrado = [
    ['Fecha', 'Hoja', 'Banco', 'Tipo de Cuenta', 'Moneda', 'Descripción', 'Monto', 'Tipo'],
    ...finalDataset2.map(r => [
      r.fecha ? r.fecha.toLocaleDateString('es-MX') : 'Sin fecha',
      r.hoja, r.banco, r.tipo_cuenta, r.moneda,
      r.descripcion, r.monto, r.tipo_registro,
    ]),
  ];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(concentrado), 'Concentrado');

  // ── Hojas por banco ──
  const bancos = [...new Set(finalDataset2.map(r => r.banco))].sort();
  for (const banco of bancos) {
    const bRows = finalDataset2.filter(r => r.banco === banco);
    const data  = [
      ['Fecha', 'Hoja', 'Tipo de Cuenta', 'Moneda', 'Descripción', 'Monto', 'Tipo'],
      ...bRows.map(r => [
        r.fecha ? r.fecha.toLocaleDateString('es-MX') : 'Sin fecha',
        r.hoja, r.tipo_cuenta, r.moneda, r.descripcion, r.monto, r.tipo_registro,
      ]),
    ];
    // Nombre de hoja: máximo 31 chars (límite Excel)
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(data), banco.substring(0, 31));
  }

  // ── Hoja Resumen (totales por banco) ──
  const resumen = [
    ['Institución', 'Total Ingresos', 'Total Egresos', 'Balance Neto', 'Movimientos'],
    ...bancos.map(banco => {
      const bRows = finalDataset2.filter(r => r.banco === banco);
      const i     = bRows.filter(r => r.tipo_registro === 'Ingreso').reduce((s, r) => s + r.monto, 0);
      const e     = bRows.filter(r => r.tipo_registro === 'Egreso').reduce((s, r) => s + r.monto, 0);
      return [banco, i, e, i - e, bRows.length];
    }),
  ];
  const totalI = finalDataset2.filter(r => r.tipo_registro === 'Ingreso').reduce((s, r) => s + r.monto, 0);
  const totalE = finalDataset2.filter(r => r.tipo_registro === 'Egreso').reduce((s, r) => s + r.monto, 0);
  resumen.push(['TOTAL GENERAL', totalI, totalE, totalI - totalE, finalDataset2.length]);
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(resumen), 'Resumen');

  const fecha = new Date().toISOString().slice(0, 10);
  XLSX.writeFile(wb, `Concentrado_SMTO_${fecha}.xlsx`);
}

// ─── ETAPA 4: DASHBOARDS POR HOJA ────────────────────────────────

/** Pasa a la etapa 4 y renderiza todos los dashboards. */
function goToDashboards() {
  renderAllSheetDashboards();
  document.getElementById('stage4').classList.remove('hidden');
  setPipelineStep(4);
  document.getElementById('stage4').scrollIntoView({ behavior: 'smooth', block: 'start' });
}

/** Construye los tabs y panels para cada hoja. */
function renderAllSheetDashboards() {
  const sheetNames = Object.keys(processedData2);
  if (!sheetNames.length) return;

  // Tabs de navegación
  const tabsNav = document.getElementById('sheetTabsNav');
  tabsNav.innerHTML = '';
  sheetNames.forEach((name, i) => {
    const meta = classifySheet(name);
    const btn  = document.createElement('button');
    btn.className     = 'sheet-tab-btn' + (i === 0 ? ' active' : '');
    btn.dataset.sheet = name;
    btn.innerHTML     = `
      <span class="sheet-tab-name" title="${escHtml(name)}">${escHtml(name)}</span>
      <span class="sheet-tab-moneda sheet-pill-moneda-${meta.moneda}">${meta.moneda}</span>`;
    btn.addEventListener('click', () => switchSheetTab(name));
    tabsNav.appendChild(btn);
  });

  // Dashboard panels
  const container = document.getElementById('sheetDashboardsContainer');
  container.innerHTML = '';
  sheetChartInsts  = {};
  sheetFilters2    = {};
  sheetIdxToName2  = {};

  sheetNames.forEach((name, i) => {
    sheetIdxToName2[i] = name;
    const panel = document.createElement('div');
    panel.className     = 'sheet-dashboard-panel' + (i === 0 ? ' active' : '');
    panel.dataset.sheet = name;
    panel.id            = `sheet-panel-${i}`;
    panel.innerHTML     = buildSheetDashboardHTML(name, i);
    container.appendChild(panel);
  });

  // Inicializar filtros + tabla de TODAS las hojas (renderiza en background)
  sheetNames.forEach((name, i) => initSheetFilters2(name, i));

  // Renderizar charts solo de la hoja activa
  if (sheetNames.length > 0) {
    activeSheetTab2 = sheetNames[0];
    renderSheetCharts(sheetNames[0], 0);
  }
}

/**
 * Genera el HTML esqueleto del dashboard por hoja.
 * KPIs y tabla se renderizan dinámicamente via initSheetFilters2 → applySheetFilters2.
 */
function buildSheetDashboardHTML(sheetName, idx) {
  const meta  = classifySheet(sheetName);
  const rows  = processedData2[sheetName] || [];
  const years = [...new Set(rows.map(r => r.year).filter(Boolean))].sort();
  const yearStr = years.join(', ');

  // ── Opciones del selector de año ──────────────────────────────
  const yearOpts = years.map(y =>
    `<option value="${y}">${y}</option>`
  ).join('');

  // ── Opciones del selector de mes ──────────────────────────────
  const mesOpts = MONTHS_LONG.map((m, mi) =>
    `<option value="${mi}">${m}</option>`
  ).join('');

  return `
    <div class="sheet-dash-header">
      <div>
        <h3 class="sheet-dash-title">${escHtml(sheetName)}</h3>
        <p class="sheet-dash-subtitle">${escHtml(meta.banco)} · ${escHtml(meta.tipo_cuenta)} · ${meta.moneda}${yearStr ? ' · ' + yearStr : ''}</p>
      </div>
      <span class="table-badge" id="sheet-badge-${idx}">${rows.length.toLocaleString('es-MX')} movimientos</span>
    </div>

    <!-- ── Barra de filtros ──────────────────────────────────── -->
    <div class="sfb-wrap">
      <div class="sfb-left">
        <select class="sfb-select" id="sfb-year-${idx}"
          onchange="onSheetYearChange(${idx}, this.value)">
          <option value="all">Todos los años</option>
          ${yearOpts}
        </select>
        <select class="sfb-select" id="sfb-mes-${idx}"
          onchange="onSheetMesChange(${idx}, this.value)">
          <option value="all">Todos los meses</option>
          ${mesOpts}
        </select>
        <div class="sfb-quick-wrap">
          <button class="sfb-quick sfb-active" id="sfb-todos-${idx}"
            onclick="setSheetTipo(${idx}, 'all', this)">Ver todo</button>
          <button class="sfb-quick sfb-income" id="sfb-ing-${idx}"
            onclick="setSheetTipo(${idx}, 'Ingreso', this)">↑ Ingresos</button>
          <button class="sfb-quick sfb-expense" id="sfb-egr-${idx}"
            onclick="setSheetTipo(${idx}, 'Egreso', this)">↓ Egresos</button>
        </div>
      </div>
      <div class="sfb-right">
        <div class="sfb-search-wrap">
          <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24" class="sfb-search-icon"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
          <input type="text" class="sfb-search" id="sfb-search-${idx}"
            placeholder="Buscar en todos los campos…"
            oninput="onSheetSearch(${idx}, this.value)">
        </div>
      </div>
    </div>

    <!-- ── KPIs dinámicos ────────────────────────────────────── -->
    <div class="kpi-grid sheet-kpi-grid" id="sheet-kpis-${idx}"></div>

    <!-- ── Gráficas ──────────────────────────────────────────── -->
    <div class="charts-grid">
      <div class="chart-card chart-bar-card">
        <div class="chart-header">
          <h3 class="chart-title">Ingresos vs Egresos por Mes</h3>
          <div class="chart-legend">
            <span class="legend-dot legend-income"></span><span>Ingresos</span>
            <span class="legend-dot legend-expense"></span><span>Egresos</span>
          </div>
        </div>
        <div class="chart-body"><canvas id="sheet-bar-${idx}"></canvas></div>
      </div>
      <div class="chart-card chart-donut-card">
        <div class="chart-header">
          <h3 class="chart-title">Distribución por Tipo</h3>
        </div>
        <div class="chart-body donut-body"><canvas id="sheet-donut-${idx}"></canvas></div>
      </div>
    </div>

    <!-- ── Tabla interactiva ──────────────────────────────────── -->
    <div class="table-card">
      <div class="table-header">
        <h3 class="chart-title">Movimientos</h3>
        <span class="table-badge" id="sheet-tbl-badge-${idx}">${rows.length.toLocaleString('es-MX')} registros</span>
      </div>
      <div class="table-wrapper tx-table-wrapper">
        <table class="data-table sheet-detail-table" id="sheet-table-${idx}">
          <thead>
            <tr>
              <th class="sort-th" onclick="sortSheetTable(${idx}, 'fecha')">
                Fecha <span class="sort-icon" id="sh-si-fecha-${idx}">⇅</span>
              </th>
              <th class="sort-th" onclick="sortSheetTable(${idx}, 'descripcion_corta')">
                Desc <span class="sort-icon" id="sh-si-descripcion_corta-${idx}">⇅</span>
              </th>
              <th class="sort-th" onclick="sortSheetTable(${idx}, 'factura')">
                Factura <span class="sort-icon" id="sh-si-factura-${idx}">⇅</span>
              </th>
              <th class="sort-th" onclick="sortSheetTable(${idx}, 'concepto')">
                Concepto <span class="sort-icon" id="sh-si-concepto-${idx}">⇅</span>
              </th>
              <th class="sort-th text-right" onclick="sortSheetTable(${idx}, 'ingreso')">
                Ingreso <span class="sort-icon" id="sh-si-ingreso-${idx}">⇅</span>
              </th>
              <th class="sort-th text-right" onclick="sortSheetTable(${idx}, 'egreso')">
                Egreso <span class="sort-icon" id="sh-si-egreso-${idx}">⇅</span>
              </th>
            </tr>
          </thead>
          <tbody id="sheet-tbody-${idx}">
            <tr><td colspan="6" class="tx-empty">Cargando…</td></tr>
          </tbody>
          <tfoot id="sheet-tfoot-${idx}"></tfoot>
        </table>
      </div>
    </div>`;
}

// ══════════════════════════════════════════════════════════════════
// FILTROS INTERACTIVOS POR HOJA
// Estado: sheetFilters2[sheetName] = { year, mes, tipo, search, sortCol, sortDir }
// Lookup: sheetIdxToName2[idx] = sheetName (evita problemas de HTML escaping)
// ══════════════════════════════════════════════════════════════════

function getSheetFilter2(sheetName) {
  if (!sheetFilters2[sheetName]) {
    sheetFilters2[sheetName] = {
      year: 'all', mes: 'all', tipo: 'all',
      search: '', sortCol: 'fecha', sortDir: 'desc',
    };
  }
  return sheetFilters2[sheetName];
}

/** Inicializa el estado por defecto y renderiza la tabla + KPIs. */
function initSheetFilters2(sheetName, idx) {
  getSheetFilter2(sheetName);
  applySheetFilters2(sheetName, idx);
}

/** Filtra, ordena y re-renderiza tabla + KPIs para una hoja. */
function applySheetFilters2(sheetName, idx) {
  const f    = getSheetFilter2(sheetName);
  let rows   = processedData2[sheetName] || [];

  // ── Filtros globales ───────────────────────────────────────────
  if (f.year !== 'all') rows = rows.filter(r => String(r.year) === String(f.year));
  if (f.mes  !== 'all') rows = rows.filter(r => String(r.mes)  === String(f.mes));
  if (f.tipo !== 'all') rows = rows.filter(r => r.tipo_registro === f.tipo);

  // ── Búsqueda de texto ──────────────────────────────────────────
  if (f.search) {
    const s = f.search.toLowerCase();
    rows = rows.filter(r =>
      (r.concepto           || '').toLowerCase().includes(s) ||
      (r.descripcion_corta  || '').toLowerCase().includes(s) ||
      (r.factura            || '').toLowerCase().includes(s) ||
      (r.descripcion        || '').toLowerCase().includes(s) ||
      (r.fecha ? formatDate(r.fecha).toLowerCase().includes(s) : false)
    );
  }

  // ── Ordenamiento ───────────────────────────────────────────────
  const col = f.sortCol;
  const dir = f.sortDir === 'asc' ? 1 : -1;
  rows = [...rows].sort((a, b) => {
    let va, vb;
    if      (col === 'fecha')   { va = a.fecha ? a.fecha.getTime() : 0; vb = b.fecha ? b.fecha.getTime() : 0; }
    else if (col === 'ingreso' || col === 'egreso') { va = a[col] || 0; vb = b[col] || 0; }
    else { va = String(a[col] || '').toLowerCase(); vb = String(b[col] || '').toLowerCase(); }
    return va < vb ? -1 * dir : va > vb ? 1 * dir : 0;
  });

  renderSheetKPIs2(sheetName, idx, rows);
  renderSheetTableBody2(sheetName, idx, rows);
  updateSortIcons2(idx, f);
}

/** Renderiza KPIs reactivos según las filas filtradas. */
function renderSheetKPIs2(sheetName, idx, filteredRows) {
  const totalRows = processedData2[sheetName] || [];
  const ing  = filteredRows.filter(r => r.tipo_registro === 'Ingreso').reduce((s, r) => s + r.monto, 0);
  const egr  = filteredRows.filter(r => r.tipo_registro === 'Egreso').reduce((s, r) => s + r.monto, 0);
  const bal  = ing - egr;
  const nIng = filteredRows.filter(r => r.tipo_registro === 'Ingreso').length;
  const nEgr = filteredRows.filter(r => r.tipo_registro === 'Egreso').length;
  const nota = filteredRows.length < totalRows.length
    ? `de ${totalRows.length.toLocaleString('es-MX')} total`
    : 'total';

  const kpisEl = document.getElementById(`sheet-kpis-${idx}`);
  if (!kpisEl) return;
  kpisEl.innerHTML = `
    <div class="kpi-card kpi-income">
      <div class="kpi-label">Ingresos</div>
      <div class="kpi-value">${formatMoney(ing)}</div>
      <div class="kpi-detail">${nIng.toLocaleString('es-MX')} movimientos</div>
    </div>
    <div class="kpi-card kpi-expense">
      <div class="kpi-label">Egresos</div>
      <div class="kpi-value">${formatMoney(egr)}</div>
      <div class="kpi-detail">${nEgr.toLocaleString('es-MX')} movimientos</div>
    </div>
    <div class="kpi-card kpi-balance">
      <div class="kpi-label">Balance</div>
      <div class="kpi-value" style="color:${bal >= 0 ? 'var(--income)' : 'var(--expense)'}">${formatMoney(bal)}</div>
      <div class="kpi-detail">${bal >= 0 ? '✓ Positivo' : '⚠ Negativo'}</div>
    </div>
    <div class="kpi-card kpi-rate">
      <div class="kpi-label">Movimientos</div>
      <div class="kpi-value">${filteredRows.length.toLocaleString('es-MX')}</div>
      <div class="kpi-detail">${nota}</div>
    </div>`;

  const badge = document.getElementById(`sheet-badge-${idx}`);
  if (badge) badge.textContent =
    `${filteredRows.length.toLocaleString('es-MX')} de ${totalRows.length.toLocaleString('es-MX')} movimientos`;
}

/** Renderiza tbody + tfoot con totales dinámicos. */
function renderSheetTableBody2(sheetName, idx, rows) {
  const tbody = document.getElementById(`sheet-tbody-${idx}`);
  const tfoot = document.getElementById(`sheet-tfoot-${idx}`);
  const badge = document.getElementById(`sheet-tbl-badge-${idx}`);
  if (!tbody) return;

  if (badge) badge.textContent = `${rows.length.toLocaleString('es-MX')} registros`;

  if (!rows.length) {
    tbody.innerHTML = `<tr><td colspan="6" class="tx-empty">Sin resultados para los filtros seleccionados</td></tr>`;
    if (tfoot) tfoot.innerHTML = '';
    return;
  }

  // Límite de filas en DOM para rendimiento (datasets grandes → usar filtros)
  const MAX_ROWS = 2000;
  const displayRows = rows.slice(0, MAX_ROWS);
  const truncated   = rows.length > MAX_ROWS;

  // Renderizar las filas visibles
  tbody.innerHTML = displayRows.map(r => {
    const isIng = r.tipo_registro === 'Ingreso';
    const fechaStr   = r.fecha ? formatDate(r.fecha) : '—';
    const descCorta  = escHtml(r.descripcion_corta || '');
    const factStr    = escHtml(r.factura || '');
    const concStr    = escHtml(r.concepto || r.descripcion || '');
    const ingStr     = isIng ? formatMoney(r.monto) : '';
    const egrStr     = isIng ? '' : formatMoney(r.monto);
    return `<tr>
      <td class="td-date">${fechaStr}</td>
      <td class="td-nombre">${descCorta}</td>
      <td class="td-ref">${factStr}</td>
      <td class="td-nombre td-concepto" title="${concStr}">${concStr}</td>
      <td class="text-right td-total-inc">${ingStr}</td>
      <td class="text-right td-total-egr">${egrStr}</td>
    </tr>`;
  }).join('');

  // Nota de truncado si hay más filas que el límite
  if (truncated) {
    tbody.innerHTML += `<tr><td colspan="6" class="tx-truncate-note">
      ⚠ Mostrando primeros ${MAX_ROWS.toLocaleString('es-MX')} de ${rows.length.toLocaleString('es-MX')} registros.
      Usa los filtros de año, mes o búsqueda para ver registros específicos.
    </td></tr>`;
  }

  // ── Footer con totales dinámicos (sobre TODOS los rows filtrados, no solo los visibles) ──
  const totalIng = rows.filter(r => r.tipo_registro === 'Ingreso').reduce((s, r) => s + r.monto, 0);
  const totalEgr = rows.filter(r => r.tipo_registro === 'Egreso').reduce((s, r) => s + r.monto, 0);
  const balance  = totalIng - totalEgr;
  const balColor = balance >= 0 ? 'var(--income)' : 'var(--expense)';

  if (tfoot) {
    tfoot.innerHTML = `
      <tr class="tfoot-totals">
        <td colspan="4">
          <strong>TOTALES — ${rows.length.toLocaleString('es-MX')} movimientos</strong>
        </td>
        <td class="text-right td-total-inc"><strong>${formatMoney(totalIng)}</strong></td>
        <td class="text-right td-total-egr"><strong>${formatMoney(totalEgr)}</strong></td>
      </tr>
      <tr class="tfoot-balance">
        <td colspan="4"><strong>BALANCE NETO</strong></td>
        <td colspan="2" class="text-right" style="color:${balColor}">
          <strong>${formatMoney(balance)}</strong>
        </td>
      </tr>`;
  }
}

// ─── Handlers de filtros ──────────────────────────────────────────

function setSheetTipo(idx, tipo, btn) {
  const name = sheetIdxToName2[idx];
  if (!name) return;
  getSheetFilter2(name).tipo = tipo;
  // Actualizar botón activo
  ['todos','ing','egr'].forEach(k => {
    const b = document.getElementById(`sfb-${k}-${idx}`);
    if (b) b.classList.remove('sfb-active');
  });
  if (btn) btn.classList.add('sfb-active');
  applySheetFilters2(name, idx);
}

function onSheetSearch(idx, val) {
  const name = sheetIdxToName2[idx];
  if (!name) return;
  getSheetFilter2(name).search = val.trim();
  applySheetFilters2(name, idx);
}

function onSheetYearChange(idx, val) {
  const name = sheetIdxToName2[idx];
  if (!name) return;
  getSheetFilter2(name).year = val;
  applySheetFilters2(name, idx);
  renderSheetCharts(name, idx);  // refrescar gráficas con el año seleccionado
}

function onSheetMesChange(idx, val) {
  const name = sheetIdxToName2[idx];
  if (!name) return;
  getSheetFilter2(name).mes = val;
  applySheetFilters2(name, idx);
}

function sortSheetTable(idx, col) {
  const name = sheetIdxToName2[idx];
  if (!name) return;
  const f = getSheetFilter2(name);
  if (f.sortCol === col) {
    f.sortDir = f.sortDir === 'asc' ? 'desc' : 'asc';
  } else {
    f.sortCol = col;
    f.sortDir = (col === 'ingreso' || col === 'egreso' || col === 'fecha') ? 'desc' : 'asc';
  }
  applySheetFilters2(name, idx);
}

function updateSortIcons2(idx, f) {
  const cols = ['fecha','descripcion_corta','factura','concepto','ingreso','egreso'];
  for (const col of cols) {
    const el = document.getElementById(`sh-si-${col}-${idx}`);
    if (!el) continue;
    if (f.sortCol === col) {
      el.textContent = f.sortDir === 'asc' ? '↑' : '↓';
      el.className   = 'sort-icon sort-active';
    } else {
      el.textContent = '⇅';
      el.className   = 'sort-icon';
    }
  }
}

/** Cambia la hoja activa en los dashboards. */
function switchSheetTab(sheetName) {
  activeSheetTab2 = sheetName;

  document.querySelectorAll('.sheet-tab-btn').forEach(btn => {
    btn.classList.toggle('active', btn.dataset.sheet === sheetName);
  });

  document.querySelectorAll('.sheet-dashboard-panel').forEach(panel => {
    panel.classList.toggle('active', panel.dataset.sheet === sheetName);
  });

  const sheetNames = Object.keys(processedData2);
  const idx = sheetNames.indexOf(sheetName);
  if (idx >= 0) renderSheetCharts(sheetName, idx);
}

/** Renderiza las gráficas de una hoja respetando el filtro de año activo. */
function renderSheetCharts(sheetName, idx) {
  const f    = getSheetFilter2(sheetName);
  let rows   = processedData2[sheetName] || [];
  if (f.year !== 'all') rows = rows.filter(r => String(r.year) === String(f.year));
  const isDark = document.documentElement.getAttribute('data-theme') !== 'light';
  const grid   = isDark ? 'rgba(255,255,255,0.06)' : 'rgba(0,0,0,0.06)';
  const tick   = isDark ? '#64748b' : '#94a3b8';

  // Destruir instancias previas para evitar memory leaks
  if (sheetChartInsts[sheetName]) {
    if (sheetChartInsts[sheetName].bar)   sheetChartInsts[sheetName].bar.destroy();
    if (sheetChartInsts[sheetName].donut) sheetChartInsts[sheetName].donut.destroy();
  }
  sheetChartInsts[sheetName] = {};

  // ── Bar chart: Ingresos vs Egresos por mes ──
  const barEl = document.getElementById(`sheet-bar-${idx}`);
  if (barEl) {
    const meses     = Array.from({ length: 12 }, () => ({ ing: 0, egr: 0 }));
    const conDatos  = new Set();
    for (const r of rows) {
      if (r.mes == null) continue;
      meses[r.mes][r.tipo_registro === 'Ingreso' ? 'ing' : 'egr'] += r.monto;
      conDatos.add(r.mes);
    }
    const sorted2  = [...conDatos].sort((a, b) => a - b);
    const labels   = sorted2.map(m => MONTHS_ES[m]);
    const ingData  = sorted2.map(m => meses[m].ing);
    const egrData  = sorted2.map(m => meses[m].egr);

    sheetChartInsts[sheetName].bar = new Chart(barEl.getContext('2d'), {
      type: 'bar',
      data: {
        labels,
        datasets: [
          { label: 'Ingresos', data: ingData, backgroundColor: 'rgba(16,185,129,0.75)', borderRadius: 6, borderSkipped: false },
          { label: 'Egresos',  data: egrData, backgroundColor: 'rgba(244,63,94,0.75)',  borderRadius: 6, borderSkipped: false },
        ],
      },
      options: {
        responsive: true, maintainAspectRatio: false,
        interaction: { mode: 'index', intersect: false },
        plugins: {
          legend: { display: false },
          tooltip: tooltipDefaults(isDark, v => formatMoney(v)),
        },
        scales: {
          x: { grid: { display: false }, ticks: { color: tick, font: { family: 'Inter', size: 11 } } },
          y: { grid: { color: grid }, border: { display: false },
               ticks: { color: tick, font: { family: 'Inter', size: 11 }, callback: v => formatMoneyShort(v) } },
        },
      },
    });
  }

  // ── Donut chart: distribución por tipo/descripción ──
  const donutEl = document.getElementById(`sheet-donut-${idx}`);
  if (donutEl) {
    // Agrupar por tipo de transacción (col 'tipo' del row)
    const byTipo = {};
    for (const r of rows) {
      const k = r.tipo || 'Otros';
      byTipo[k] = (byTipo[k] || 0) + r.monto;
    }
    const sorted3   = Object.entries(byTipo).sort((a, b) => b[1] - a[1]);
    const top        = sorted3.slice(0, 8);
    const othersSum  = sorted3.slice(8).reduce((s, [, v]) => s + v, 0);
    const slices     = othersSum > 0 ? [...top, ['Otros', othersSum]] : top;
    const total      = slices.reduce((s, [, v]) => s + v, 0);

    const subC   = isDark ? '#94a3b8' : '#64748b';
    const mainC  = isDark ? '#f1f5f9' : '#0f172a';

    const centerPlugin = {
      id: `sheetCenter_${idx}`,
      beforeDraw(chart) {
        const { ctx: c, chartArea } = chart;
        if (!chartArea) return;
        const cx = chartArea.left + chartArea.width  / 2;
        const cy = chartArea.top  + chartArea.height / 2;
        c.save();
        c.textAlign = 'center'; c.textBaseline = 'middle';
        c.font = '500 11px Inter, system-ui'; c.fillStyle = subC;
        c.fillText('TOTAL', cx, cy - 15);
        c.font = '800 18px Inter, system-ui'; c.fillStyle = mainC;
        c.fillText(formatMoneyShort(total), cx, cy + 10);
        c.restore();
      },
    };

    sheetChartInsts[sheetName].donut = new Chart(donutEl.getContext('2d'), {
      type: 'doughnut',
      plugins: [centerPlugin],
      data: {
        labels: slices.map(([t]) => t),
        datasets: [{
          data: slices.map(([, v]) => v),
          backgroundColor: slices.map((_, i) => PALETTE[i % PALETTE.length]),
          borderColor: isDark ? '#1e293b' : '#f8fafc',
          borderWidth: 2,
          hoverOffset: 10,
        }],
      },
      options: {
        responsive: true, maintainAspectRatio: false, cutout: '68%',
        plugins: {
          legend: {
            position: 'bottom',
            labels: {
              color: isDark ? '#f1f5f9' : '#0f172a',
              font: { family: 'Inter', size: 10, weight: '500' },
              padding: 12, boxWidth: 8, usePointStyle: true,
            },
          },
          tooltip: {
            backgroundColor: isDark ? '#1e293b' : '#fff',
            titleColor: isDark ? '#f1f5f9' : '#0f172a',
            bodyColor:  isDark ? '#94a3b8' : '#475569',
            borderColor: isDark ? 'rgba(255,255,255,0.12)' : 'rgba(0,0,0,0.08)',
            borderWidth: 1, padding: 12, cornerRadius: 10,
            callbacks: {
              label: ctx => {
                const v   = ctx.dataset.data[ctx.dataIndex];
                const pct = total > 0 ? (v / total * 100).toFixed(1) : '0';
                return ` ${ctx.label}: ${formatMoney(v)} (${pct}%)`;
              },
            },
          },
        },
      },
    });
  }
}

// ─── PIPELINE PROGRESS BAR ────────────────────────────────────────
function setPipelineStep(step) {
  for (let i = 1; i <= 4; i++) {
    const el = document.getElementById(`pBarS${i}`);
    if (!el) continue;
    el.classList.remove('active', 'completed');
    if (i < step)        el.classList.add('completed');
    else if (i === step) el.classList.add('active');
  }
  for (let i = 1; i <= 3; i++) {
    const line = document.getElementById(`pBarL${i}`);
    if (line) line.classList.toggle('completed', i < step);
  }
}

// ─── ERROR HELPERS ────────────────────────────────────────────────
function showImport2Error(msg) {
  const el = document.getElementById('import2Error');
  if (!el) return;
  const span = el.querySelector('span');
  if (span) span.textContent = msg;
  el.classList.remove('hidden');
}
function hideImport2Error() {
  const el = document.getElementById('import2Error');
  if (el) el.classList.add('hidden');
}

// ─── RESET MÓDULO 2 ──────────────────────────────────────────────
function resetImport2() {
  rawDataFile2    = {};
  processedData2  = {};
  finalDataset2   = [];
  file2FileName   = '';
  activeSheetTab2 = null;
  sheetFilters2   = {};
  sheetIdxToName2 = {};

  // Destruir chart instances
  for (const insts of Object.values(sheetChartInsts)) {
    if (insts && insts.bar)   insts.bar.destroy();
    if (insts && insts.donut) insts.donut.destroy();
  }
  sheetChartInsts = {};

  // Ocultar etapas 2-4
  ['stage2', 'stage3', 'stage4'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.classList.add('hidden');
  });

  // Limpiar UI de etapa 1
  document.getElementById('import2SheetsDetected').classList.add('hidden');
  document.getElementById('import2FileInput').value = '';
  hideImport2Error();
  setPipelineStep(1);
}
