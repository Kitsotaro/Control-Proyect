// CONFIGURACIÓN CENTRAL
const CLIENT_ID = '1070607567316-mdbd97lbkprgpc4spj71e5f8anovr6it.apps.googleusercontent.com';
const SCOPES = 'https://www.googleapis.com/auth/drive.file https://www.googleapis.com/auth/spreadsheets';
const DB_FILE_NAME = 'StockCentral_DB';
// Nombres antiguos: si el archivo ya existe con alguno de estos nombres, se RENOMBRA
// automáticamente al nuevo (sin perder datos) en vez de crear una base nueva.
const LEGACY_DB_NAMES = ['RutaControl_DB'];

const DEFAULT_STOCK_MIN = 4;
const DEFAULT_STOCK_MAX = 6;

let SPREADSHEET_ID = '';
let tokenClient;
let gapiInited = false;
let gsisInited = false;

// ESTADO DE LA APLICACIÓN
let tipoMovimiento = 'ENTRADA';
let catalogoProductos = [];
let productoVentaSeleccionado = null;
let vistaUnidadesSueltas = false;
let stockChartInstance = null;
let estadoModalSeleccionado = 'ACTIVO';

window.addEventListener('DOMContentLoaded', () => {
  const fechaInput = document.getElementById('fecha-mov');
  if (fechaInput) fechaInput.valueAsDate = new Date();

  iniciarRelojCaptura();
  inicializarAutocompletesEntrada();
  inicializarBuscadorVenta();
});

// RELOJ DE CAPTURA EN VIVO (muestra el timestamp que se grabará al guardar)
function iniciarRelojCaptura() {
  const el = document.getElementById('log-timestamp-live');
  if (!el) return;
  const actualizar = () => {
    const now = new Date();
    el.textContent = now.toLocaleString('es-ES', {
      day: '2-digit', month: '2-digit', year: 'numeric',
      hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: false
    });
  };
  actualizar();
  setInterval(actualizar, 1000);
}

// LIMPIAR FORMULARIO ACTIVO
function limpiarFormularioActivo() {
  if (tipoMovimiento === 'ENTRADA') {
    document.getElementById('form-entrada').reset();
  } else {
    document.getElementById('form-venta').reset();
    const tipoEmpEl = document.getElementById('venta-prod-tipo-empaque');
    if (tipoEmpEl) tipoEmpEl.value = '';
    productoVentaSeleccionado = null;
    document.getElementById('btn-save-out').disabled = true;
  }
  document.getElementById('fecha-mov').valueAsDate = new Date();
  generarFolioCorrelativo();
}

// ===== AUTOCOMPLETADO PROPIO (reemplaza al <datalist> nativo, que no se puede estilizar) =====
function crearAutocomplete(inputId, dropdownId, getOpciones, onSeleccionar, renderItem) {
  const input = document.getElementById(inputId);
  const dropdown = document.getElementById(dropdownId);
  if (!input || !dropdown) return;

  function render(query) {
    const opciones = getOpciones(query.trim());
    dropdown.innerHTML = '';
    if (!query.trim() || opciones.length === 0) {
      dropdown.classList.add('hidden');
      return;
    }
    opciones.slice(0, 8).forEach(opt => {
      const item = document.createElement('div');
      item.className = 'autocomplete-item';
      item.innerHTML = renderItem(opt);
      // mousedown (no click) para que dispare ANTES del blur del input
      item.addEventListener('mousedown', (e) => {
        e.preventDefault();
        onSeleccionar(opt);
        dropdown.classList.add('hidden');
      });
      dropdown.appendChild(item);
    });
    dropdown.classList.remove('hidden');
  }

  input.addEventListener('input', () => render(input.value));
  input.addEventListener('focus', () => { if (input.value.trim()) render(input.value); });
  input.addEventListener('blur', () => setTimeout(() => dropdown.classList.add('hidden'), 150));
}

function sugerirValoresUnicos(campo, query) {
  const q = query.toLowerCase();
  const valores = [...new Set(catalogoProductos.map(p => p[campo]).filter(Boolean))];
  return valores.filter(v => v.toLowerCase().includes(q));
}

function inicializarAutocompletesEntrada() {
  crearAutocomplete('prod-marca', 'ac-marca', q => sugerirValoresUnicos('marca', q),
    v => { document.getElementById('prod-marca').value = v; }, v => v);
  crearAutocomplete('prod-linea', 'ac-linea', q => sugerirValoresUnicos('linea', q),
    v => { document.getElementById('prod-linea').value = v; }, v => v);
  crearAutocomplete('prod-tipo-empaque', 'ac-tipoempaque', q => sugerirValoresUnicos('tipoEmpaque', q),
    v => { document.getElementById('prod-tipo-empaque').value = v; }, v => v);
  crearAutocomplete('prod-presentacion', 'ac-presentacion', q => sugerirValoresUnicos('magnitud', q),
    v => { document.getElementById('prod-presentacion').value = v; }, v => v);
}

// ===== BUSCADOR DE VENTA (reemplaza datalist + parsing frágil de texto) =====
function inicializarBuscadorVenta() {
  const input = document.getElementById('buscar-producto-venta');
  const dropdown = document.getElementById('dropdown-productos-venta');
  const btnSave = document.getElementById('btn-save-out');
  if (!input || !dropdown) return;

  function limpiarSeleccion() {
    productoVentaSeleccionado = null;
    ['venta-prod-marca', 'venta-prod-linea', 'venta-prod-presentacion', 'venta-prod-variante', 'venta-prod-volumen', 'venta-prod-tipo-empaque']
      .forEach(id => { const el = document.getElementById(id); if (el) el.value = ''; });
    if (btnSave) btnSave.disabled = true;
  }

  function buscar(query) {
    const q = query.toLowerCase().trim();
    if (!q) return [];
    return catalogoProductos.filter(p => p.estado === 'ACTIVO' && (
      p.sku.toLowerCase().includes(q) ||
      p.marca.toLowerCase().includes(q) ||
      p.linea.toLowerCase().includes(q) ||
      (p.variante || '').toLowerCase().includes(q)
    )).slice(0, 8);
  }

  function seleccionar(p) {
    productoVentaSeleccionado = p;
    input.value = `[${p.sku}] ${p.marca} ${p.linea}`;
    document.getElementById('venta-prod-marca').value = p.marca;
    document.getElementById('venta-prod-linea').value = p.linea;
    document.getElementById('venta-prod-presentacion').value = p.magnitud;
    document.getElementById('venta-prod-variante').value = p.variante;
    document.getElementById('venta-prod-volumen').value = p.volumen;
    const tipoEmpEl = document.getElementById('venta-prod-tipo-empaque');
    if (tipoEmpEl) tipoEmpEl.value = p.tipoEmpaque || '--';
    if (btnSave) btnSave.disabled = false;
  }

  crearAutocomplete('buscar-producto-venta', 'dropdown-productos-venta', buscar, seleccionar, p => `
    <div class="ac-item-main">${p.marca} ${p.linea} ${p.volumen}${p.variante ? ' · ' + p.variante : ''}</div>
    <div class="ac-item-sub">SKU: ${p.sku} · ${p.magnitud} · Stock: ${p.stock}</div>
  `);

  input.addEventListener('input', () => { if (productoVentaSeleccionado) limpiarSeleccion(); });
}

function gapiLoaded() {
  gapi.load('client', async () => {
    await gapi.client.init({
      discoveryDocs: [
        'https://sheets.googleapis.com/$discovery/rest?version=v4',
        'https://www.googleapis.com/discovery/v1/apis/drive/v3/rest'
      ]
    });
    gapiInited = true;
    checkAuthReady();
  });
}

function gisLoaded() {
  tokenClient = google.accounts.oauth2.initTokenClient({
    client_id: CLIENT_ID,
    scope: SCOPES,
    callback: '', 
  });
  gsisInited = true;
  checkAuthReady();
}

window.onload = () => { gapiLoaded(); gisLoaded(); };

function checkAuthReady() {
  if (gapiInited && gsisInited) {
    document.getElementById('status').innerText = 'Listo para conectar.';
  }
}

function handleAuthClick() {
  tokenClient.callback = async (resp) => {
    if (resp.error) throw (resp);
    document.getElementById('auth-section').classList.add('hidden');
    document.getElementById('main-app').classList.remove('hidden');
    document.getElementById('status').innerText = 'Verificando base de datos en Google Drive...';
    await cargarDatosIniciales();
  };

  if (gapi.client.getToken() === null) {
    tokenClient.requestAccessToken({prompt: 'consent'});
  } else {
    tokenClient.requestAccessToken({prompt: ''});
  }
}

// NAVEGACIÓN ENTRE PESTAÑAS
function switchTab(tabId, event) {
  document.querySelectorAll('.tab-content').forEach(t => t.classList.remove('active'));
  document.querySelectorAll('.nav-btn').forEach(b => b.classList.remove('active'));
  
  document.getElementById(tabId).classList.add('active');
  if (event && event.currentTarget) {
    event.currentTarget.classList.add('active');
  }

  if (tabId === 'tab-stock') {
    renderizarDashboardStock();
  } else if (tabId === 'tab-ventas') {
    renderizarVentasHoy();
  }
}

// TOGGLE TIPO MOVIMIENTO (ENTRADA / VENTA)
function setTipoMovimiento(tipo) {
  tipoMovimiento = tipo;
  document.getElementById('btn-tipo-entrada').classList.toggle('active', tipo === 'ENTRADA');
  document.getElementById('btn-tipo-venta').classList.toggle('active', tipo === 'VENTA');

  document.getElementById('form-entrada').classList.toggle('hidden', tipo !== 'ENTRADA');
  document.getElementById('form-venta').classList.toggle('hidden', tipo !== 'VENTA');

  generarFolioCorrelativo();
}

// VALIDACIONES DE PRECIOS CON DECIMALES
function validarPrecioDistribuidor(input) {
  const val = input.value;
  const warn = document.getElementById('warn-pdist');
  if (val.includes('.')) {
    const decimals = val.split('.')[1];
    if (decimals && decimals.length > 5) {
      if (warn) warn.classList.remove('hidden');
      input.value = parseFloat(val).toFixed(5);
      return;
    }
  }
  if (warn) warn.classList.add('hidden');
}

function validarPrecioConsumidor(input) {
  const val = input.value;
  if (val.includes('.')) {
    const decimals = val.split('.')[1];
    if (decimals && decimals.length > 2) {
      input.value = parseFloat(val).toFixed(2);
    }
  }
}

// LECTURA DE CATALOGO Y BD
async function cargarDatosIniciales() {
  try {
    SPREADSHEET_ID = await buscarOCrearBaseDatos();

    const res = await gapi.client.sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: 'CATALOGO!A2:O',
    });

    const numOrDefault = (val, def) => {
      const n = parseFloat(val);
      return (val !== undefined && val !== '' && !isNaN(n)) ? n : def;
    };

    const rows = res.result ? res.result.values || [] : [];
    catalogoProductos = rows.map(r => ({
      sku: r[0], marca: r[1], linea: r[2], magnitud: r[3],
      cantEmp: parseInt(r[4]) || 1, volumen: r[5], variante: r[6] || '',
      pDist: parseFloat(r[7]) || 0, pCons: parseFloat(r[8]) || 0,
      stock: parseFloat(r[9]) || 0, estado: r[10] || 'ACTIVO',
      tipoEmpaque: r[12] || '',
      stockMin: numOrDefault(r[13], DEFAULT_STOCK_MIN),
      stockMax: numOrDefault(r[14], DEFAULT_STOCK_MAX)
    }));

    generarFolioCorrelativo();
    document.getElementById('status').innerText = 'Conectado a la base de datos.';
  } catch (err) {
    document.getElementById('status').innerText = 'Error al cargar: ' + err.message;
  }
}

async function buscarArchivoPorNombre(nombre) {
  const response = await gapi.client.drive.files.list({
    q: `name = '${nombre}' and mimeType = 'application/vnd.google-apps.spreadsheet' and trashed = false`,
    fields: 'files(id, name)',
    spaces: 'drive'
  });
  const files = response.result ? response.result.files : null;
  return (files && files.length > 0) ? files[0] : null;
}

async function buscarOCrearBaseDatos() {
  // 1. Buscar con el nombre actual
  const actual = await buscarArchivoPorNombre(DB_FILE_NAME);
  if (actual) return actual.id;

  // 2. Buscar con nombres antiguos y migrar (renombrar) si se encuentra, sin perder datos
  for (const nombreLegacy of LEGACY_DB_NAMES) {
    const legacy = await buscarArchivoPorNombre(nombreLegacy);
    if (legacy) {
      document.getElementById('status').innerText = `Migrando base de datos de "${nombreLegacy}" a "${DB_FILE_NAME}"...`;
      await gapi.client.drive.files.update({
        fileId: legacy.id,
        resource: { name: DB_FILE_NAME }
      });
      return legacy.id;
    }
  }

  // 3. No existe con ningún nombre -> crear una base nueva desde cero
  document.getElementById('status').innerText = 'Creando nueva base de datos en tu Drive...';

  const createRes = await gapi.client.sheets.spreadsheets.create({
    resource: {
      properties: { title: DB_FILE_NAME },
      sheets: [
        { properties: { title: 'LOG_TRANS' } },
        { properties: { title: 'CATALOGO' } },
        { properties: { title: 'LOG_HISTORICO_PRECIOS' } }
      ]
    }
  });

  const newSpreadsheetId = createRes.result.spreadsheetId;
  await inicializarEncabezadosBD(newSpreadsheetId);
  return newSpreadsheetId;
}

async function inicializarEncabezadosBD(spreadsheetId) {
  const encabezadosLOG = [
    'TRANS_ID', 'TIMESTAMP_LOG', 'FECHA_MOV', 'TIPO_MOV', 'SKU_ITEM',
    'MARCA', 'LINEA_PROD', 'MAGNITUD', 'VOLUMEN', 'VARIANTE', 'CANT_EMPAQUE',
    'CANTIDAD', 'UNID_BONIF', 'P_DISTRIBUIDOR', 'P_CONSUMIDOR',
    'TOTAL_INVERSION', 'MARGEN_UNIT', 'TOTAL_VENTA', 'UTILIDAD_NETA'
  ];

  const encabezadosCatalogo = [
    'SKU_ITEM', 'MARCA', 'LINEA_PROD', 'MAGNITUD', 'CANT_EMPAQUE',
    'VOLUMEN', 'VARIANTE', 'P_DISTRIBUIDOR', 'P_CONSUMIDOR',
    'STOCK_ACTUAL', 'ESTADO_ITEM', 'ULTIMA_MODIF',
    'TIPO_EMPAQUE', 'STOCK_MIN', 'STOCK_MAX'
  ];

  const encabezadosHistorico = [
    'ID_LOG', 'SKU_ITEM', 'P_DIST_ANT', 'P_DIST_NUEVO',
    'P_CONS_ANT', 'P_CONS_NUEVO', 'FECHA_CAMBIO', 'USUARIO'
  ];

  await gapi.client.sheets.spreadsheets.values.batchUpdate({
    spreadsheetId: spreadsheetId,
    resource: {
      valueInputOption: 'USER_ENTERED',
      data: [
        { range: 'LOG_TRANS!A1:S1', values: [encabezadosLOG] },
        { range: 'CATALOGO!A1:O1', values: [encabezadosCatalogo] },
        { range: 'LOG_HISTORICO_PRECIOS!A1:H1', values: [encabezadosHistorico] }
      ]
    }
  });
}



function generarSKUCompacto(marca, linea, volumen, variante) {
  const clean = (t) => (t || '').toUpperCase().replace(/[^A-Z0-9]/g, '').substring(0, 3);
  const m = clean(marca);
  const l = clean(linea);
  const v = clean(volumen);
  const varStr = variante ? '-' + clean(variante) : '';
  return `${m}-${l}-${v}${varStr}`;
}

async function generarFolioCorrelativo() {
  const prefijo = (tipoMovimiento === 'ENTRADA') ? 'IN-' : 'FAC-';
  try {
    const res = await gapi.client.sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: 'LOG_TRANS!A:A',
    });
    const rows = res.result ? res.result.values || [] : [];
    let maxNum = 0;
    rows.forEach(r => {
      if (r[0] && r[0].startsWith(prefijo)) {
        const num = parseInt(r[0].replace(prefijo, ''));
        if (num > maxNum) maxNum = num;
      }
    });
    const siguiente = String(maxNum + 1).padStart(6, '0');
    document.getElementById('trans-id').value = `${prefijo}${siguiente}`;
  } catch (e) {
    document.getElementById('trans-id').value = `${prefijo}000001`;
  }
}

// GUARDAR MOVIMIENTO
async function guardarMovimiento(event) {
  event.preventDefault();
  const btn = (tipoMovimiento === 'ENTRADA') ? document.getElementById('btn-save-in') : document.getElementById('btn-save-out');
  btn.disabled = true;
  document.getElementById('status').innerText = 'Guardando registro...';

  const transId = document.getElementById('trans-id').value;
  const timestampLog = new Date().toISOString().replace('T', ' ').substring(0, 19);
  const fechaMov = document.getElementById('fecha-mov').value;

  let marca, linea, magnitud, cantEmp, volumen, variante, pDist, pCons, sku, rawCantidad, bonif, tipoEmpaque;

  if (tipoMovimiento === 'ENTRADA') {
    marca = document.getElementById('prod-marca').value.trim();
    linea = document.getElementById('prod-linea').value.trim();
    magnitud = document.getElementById('prod-presentacion').value.trim();
    cantEmp = 1;
    volumen = document.getElementById('prod-volumen').value.trim();
    variante = document.getElementById('prod-variante').value.trim();
    tipoEmpaque = document.getElementById('prod-tipo-empaque').value.trim();
    rawCantidad = parseFloat(document.getElementById('mov-cantidad-in').value);
    bonif = parseFloat(document.getElementById('mov-bonificacion').value) || 0;
    pDist = parseFloat(document.getElementById('precio-distribuidor').value) || 0;
    pCons = parseFloat(document.getElementById('precio-consumidor').value) || 0;
    sku = generarSKUCompacto(marca, linea, volumen, variante);
  } else {
    if (!productoVentaSeleccionado) {
      alert('Debes seleccionar un producto válido antes de guardar la venta.');
      btn.disabled = false;
      return;
    }
    marca = productoVentaSeleccionado.marca;
    linea = productoVentaSeleccionado.linea;
    magnitud = productoVentaSeleccionado.magnitud;
    cantEmp = productoVentaSeleccionado.cantEmp;
    volumen = productoVentaSeleccionado.volumen;
    variante = productoVentaSeleccionado.variante;
    rawCantidad = parseFloat(document.getElementById('mov-cantidad-out').value);
    bonif = 0;
    pDist = productoVentaSeleccionado.pDist;
    pCons = productoVentaSeleccionado.pCons;
    sku = productoVentaSeleccionado.sku;
  }

  const cantidadSigno = (tipoMovimiento === 'VENTA') ? -Math.abs(rawCantidad) : Math.abs(rawCantidad);
  const totalInversion = Math.abs(cantidadSigno) * pDist;
  const margenUnit = pCons - pDist;
  const totalVenta = Math.abs(cantidadSigno) * pCons;
  const utilidadNeta = Math.abs(cantidadSigno) * margenUnit;

  try {
    await gapi.client.sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: 'LOG_TRANS!A:S',
      valueInputOption: 'USER_ENTERED',
      resource: {
        values: [[
          transId, timestampLog, fechaMov, tipoMovimiento, sku,
          marca, linea, magnitud, volumen, variante, cantEmp,
          cantidadSigno, bonif, pDist, pCons,
          totalInversion, margenUnit, totalVenta, utilidadNeta
        ]]
      }
    });

    let prodIndex = catalogoProductos.findIndex(p => p.sku === sku);
    if (prodIndex >= 0) {
      catalogoProductos[prodIndex].stock += cantidadSigno;
      if (tipoMovimiento === 'ENTRADA') {
        catalogoProductos[prodIndex].pDist = pDist || catalogoProductos[prodIndex].pDist;
        catalogoProductos[prodIndex].pCons = pCons || catalogoProductos[prodIndex].pCons;
        catalogoProductos[prodIndex].tipoEmpaque = tipoEmpaque || catalogoProductos[prodIndex].tipoEmpaque;
      }
    } else {
      catalogoProductos.push({
        sku, marca, linea, magnitud, cantEmp, volumen, variante,
        pDist, pCons, stock: cantidadSigno, estado: 'ACTIVO',
        tipoEmpaque: tipoEmpaque || '',
        stockMin: DEFAULT_STOCK_MIN, stockMax: DEFAULT_STOCK_MAX
      });
    }

    await reescribirHojaCatalogo();

    document.getElementById('status').innerText = '¡Guardado correctamente!';
    if (tipoMovimiento === 'ENTRADA') document.getElementById('form-entrada').reset();
    else {
      document.getElementById('form-venta').reset();
      document.getElementById('btn-save-out').disabled = true;
      productoVentaSeleccionado = null;
    }
    document.getElementById('fecha-mov').valueAsDate = new Date();
    await cargarDatosIniciales();
  } catch (err) {
    document.getElementById('status').innerText = 'Error: ' + err.message;
  } finally {
    btn.disabled = false;
  }
}

async function reescribirHojaCatalogo() {
  const rows = catalogoProductos.map(p => [
    p.sku, p.marca, p.linea, p.magnitud, p.cantEmp, p.volumen, p.variante,
    p.pDist, p.pCons, p.stock, p.estado, new Date().toISOString(),
    p.tipoEmpaque || '', p.stockMin ?? DEFAULT_STOCK_MIN, p.stockMax ?? DEFAULT_STOCK_MAX
  ]);

  await gapi.client.sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID,
    range: 'CATALOGO!A2:O',
    valueInputOption: 'USER_ENTERED',
    resource: { values: rows }
  });
}

// RENDERS STOCK & VENTAS
function renderizarDashboardStock() {
  let valorTotalBodega = 0;
  let totalUnidadesSueltas = 0;

  catalogoProductos.forEach(p => {
    if (p.estado === 'ACTIVO') {
      valorTotalBodega += (p.stock * p.pDist);
      totalUnidadesSueltas += (p.stock * p.cantEmp);
    }
  });

  document.getElementById('metric-valor-bodega').innerText = `$${valorTotalBodega.toFixed(2)}`;
  document.getElementById('metric-total-unidades').innerText = totalUnidadesSueltas.toString();

  renderizarGraficoStock();
  renderizarListaStock();
}

function renderizarListaStock() {
  const container = document.getElementById('lista-productos-stock');
  const mostrarObsoletos = document.getElementById('chk-mostrar-obsoletos').checked;
  container.innerHTML = '';

  const listaProcesada = catalogoProductos
    .filter(p => mostrarObsoletos || p.estado === 'ACTIVO')
    .sort((a, b) => {
      const getPriority = (stock) => (stock <= 4) ? 1 : (stock <= 6) ? 2 : 3;
      return getPriority(a.stock) - getPriority(b.stock);
    });

  listaProcesada.forEach(p => {
    let semaforoClase = (p.stock <= 4) ? 'desabastecido' : (p.stock <= 6) ? 'stock-bajo' : 'stock-optimo';
    let semaforoTexto = (p.stock <= 4) ? '🔴 Desabastecido' : (p.stock <= 6) ? '🟡 Stock Bajo' : '🟢 Stock Óptimo';

    let displayStock = vistaUnidadesSueltas ? (p.stock * p.cantEmp) + ' Unid.' : p.stock + ' ' + p.magnitud;

    const div = document.createElement('div');
    div.className = `product-item ${semaforoClase}`;
    div.innerHTML = `
      <div class="prod-info">
        <span class="prod-title">${p.marca} ${p.linea} ${p.volumen} ${p.variante}</span>
        <span class="prod-sub">SKU: ${p.sku} | Stock: <strong>${displayStock}</strong></span>
        <span class="prod-sub">Dist: $${p.pDist.toFixed(5)} | Cons: $${p.pCons.toFixed(2)}</span>
        <span class="prod-badge">${semaforoTexto} (${p.estado})</span>
      </div>
      <div>
        <button class="btn-secondary-sm" onclick="abrirModalEdicion('${p.sku}')">✏️ Editar</button>
      </div>
    `;
    container.appendChild(div);
  });
}

function toggleVistaUnidades() {
  vistaUnidadesSueltas = !vistaUnidadesSueltas;
  document.getElementById('lbl-unidad-vista').innerText = vistaUnidadesSueltas ? 'Unidades Sueltas' : 'Cajas';
  renderizarListaStock();
}

function renderizarGraficoStock() {
  const ctx = document.getElementById('stockPieChart').getContext('2d');
  const groupKey = document.getElementById('filter-chart-group').value;

  const agrupado = {};
  catalogoProductos.filter(p => p.estado === 'ACTIVO').forEach(p => {
    const key = (groupKey === 'MARCA') ? p.marca : p.linea;
    agrupado[key] = (agrupado[key] || 0) + p.stock;
  });

  if (stockChartInstance) stockChartInstance.destroy();

  stockChartInstance = new Chart(ctx, {
    type: 'pie',
    data: {
      labels: Object.keys(agrupado),
      datasets: [{
        data: Object.values(agrupado),
        backgroundColor: ['#22c55e', '#3b82f6', '#eab308', '#ef4444', '#a855f7', '#ec4899']
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { position: 'bottom', labels: { color: '#f8fafc' } } }
    }
  });
}

async function renderizarVentasHoy() {
  const fechaFiltro = document.getElementById('filtro-fecha-ventas').value;
  const targetDate = fechaFiltro || new Date().toISOString().split('T')[0];

  try {
    const res = await gapi.client.sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: 'LOG_TRANS!A2:S',
    });

    const rows = res.result ? res.result.values || [] : [];
    let totalVentas = 0;
    let totalGanancia = 0;
    const listaVentas = document.getElementById('lista-ventas-hoy');
    listaVentas.innerHTML = '';

    rows.forEach(r => {
      const fechaMov = r[2];
      const tipo = r[3];
      if (tipo === 'VENTA' && fechaMov === targetDate) {
        const vta = parseFloat(r[17]) || 0;
        const util = parseFloat(r[18]) || 0;
        totalVentas += vta;
        totalGanancia += util;

        const div = document.createElement('div');
        div.className = 'product-item stock-optimo';
        div.innerHTML = `
          <div class="prod-info">
            <span class="prod-title">${r[5]} ${r[6]} ${r[8]}</span>
            <span class="prod-sub">Folio: ${r[0]} | Cant: ${Math.abs(parseFloat(r[11]))}</span>
          </div>
          <div>
            <strong>$${vta.toFixed(2)}</strong>
          </div>
        `;
        listaVentas.appendChild(div);
      }
    });

    document.getElementById('metric-venta-hoy').innerText = `$${totalVentas.toFixed(2)}`;
    document.getElementById('metric-utilidad-hoy').innerText = `$${totalGanancia.toFixed(2)}`;
  } catch (err) {
    document.getElementById('status').innerText = 'Error al cargar ventas: ' + err.message;
  }
}

function resetearFiltroFechaVentas() {
  document.getElementById('filtro-fecha-ventas').value = '';
  renderizarVentasHoy();
}

// MODAL
function abrirModalEdicion(sku) {
  const prod = catalogoProductos.find(p => p.sku === sku);
  if (!prod) return;

  document.getElementById('modal-sku-item').value = prod.sku;
  document.getElementById('modal-titulo').innerText = `Editar: ${prod.marca} ${prod.linea}`;
  document.getElementById('modal-p-distributor').value = prod.pDist;
  document.getElementById('modal-p-consumer').value = prod.pCons;
  
  setEstadoModal(prod.estado);
  document.getElementById('modal-editar').classList.remove('hidden');
}

function setEstadoModal(estado) {
  estadoModalSeleccionado = estado;
  document.getElementById('btn-estado-activo').classList.toggle('active', estado === 'ACTIVO');
  document.getElementById('btn-estado-obsoleto').classList.toggle('active', estado === 'OBSOLETO');
}

function cerrarModal() {
  document.getElementById('modal-editar').classList.add('hidden');
}

async function guardarCambiosModal() {
  const sku = document.getElementById('modal-sku-item').value;
  const prod = catalogoProductos.find(p => p.sku === sku);
  if (!prod) return;

  const newPDist = parseFloat(document.getElementById('modal-p-distributor').value) || 0;
  const newPCons = parseFloat(document.getElementById('modal-p-consumer').value) || 0;

  try {
    await gapi.client.sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: 'LOG_HISTORICO_PRECIOS!A:H',
      valueInputOption: 'USER_ENTERED',
      resource: {
        values: [[
          `LOG-${Date.now()}`, sku, prod.pDist, newPDist, prod.pCons, newPCons,
          new Date().toISOString(), 'USUARIO_ACTIVO'
        ]]
      }
    });

    prod.pDist = newPDist;
    prod.pCons = newPCons;
    prod.estado = estadoModalSeleccionado;

    await reescribirHojaCatalogo();
    cerrarModal();
    renderizarDashboardStock();
    document.getElementById('status').innerText = 'Cambios guardados exitosamente.';
  } catch (err) {
    alert('Error al guardar en LOG: ' + err.message);
  }
}
