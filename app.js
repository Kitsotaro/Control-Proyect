// CONFIGURACIÓN CENTRAL
const CLIENT_ID = '1070607567316-mdbd97lbkprgpc4spj71e5f8anovr6it.apps.googleusercontent.com';
const SCOPES = 'https://www.googleapis.com/auth/drive.file https://www.googleapis.com/auth/spreadsheets';
const DB_FILE_NAME = 'RutaControl_DB'; // Nombre automático de la base de datos en Drive

let SPREADSHEET_ID = ''; // Se asigna automáticamente tras buscar o crear el archivo
let tokenClient;
let gapiInited = false;
let gsisInited = false;

// ESTADO DE LA APLICACIÓN
let tipoMovimiento = 'ENTRADA'; // 'ENTRADA' o 'VENTA'
let catalogoProductos = []; // Almacena ítems leídos de CATALOGO
let vistaUnidadesSueltas = false; // Alternar vista de Cajas a Unidades
let stockChartInstance = null;
let estadoModalSeleccionado = 'ACTIVO';

window.addEventListener('DOMContentLoaded', () => {
  const fechaInput = document.getElementById('fecha-mov');
  if (fechaInput) fechaInput.valueAsDate = new Date();
});

function gapiLoaded() {
  gapi.load('client', async () => {
    await gapi.client.init({});
    await gapi.client.load('https://sheets.googleapis.com/$discovery/rest?version=v4');
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
function switchTab(tabId) {
  document.querySelectorAll('.tab-content').forEach(t => t.classList.remove('active'));
  document.querySelectorAll('.nav-btn').forEach(b => b.classList.remove('active'));
  
  document.getElementById(tabId).classList.add('active');
  event.currentTarget.classList.add('active');

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

  // Ajustar visibilidad de campos según regla
  document.getElementById('group-bonificacion').style.display = (tipo === 'ENTRADA') ? 'block' : 'none';
  document.getElementById('group-precio-distribuidor').style.display = (tipo === 'ENTRADA') ? 'block' : 'none';

  generarFolioCorrelativo();
}

// LECTURA DE CATALOGO Y GESTIÓN AUTOMÁTICA DE BASE DE DATOS
async function cargarDatosIniciales() {
  try {
    // 1. Detección o Auto-creación en Google Drive
    SPREADSHEET_ID = await buscarOCrearBaseDatos();

    // 2. Carga de datos desde la pestaña CATALOGO
    const res = await gapi.client.sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: 'CATALOGO!A2:L',
    });

    const rows = res.result.values || [];
    catalogoProductos = rows.map(r => ({
      sku: r[0], marca: r[1], linea: r[2], magnitud: r[3],
      cantEmp: parseInt(r[4]) || 1, volumen: r[5], variante: r[6] || '',
      pDist: parseFloat(r[7]) || 0, pCons: parseFloat(r[8]) || 0,
      stock: parseFloat(r[9]) || 0, estado: r[10] || 'ACTIVO'
    }));

    poblarDataLists();
    generarFolioCorrelativo();
    document.getElementById('status').innerText = 'Conectado a la base de datos.';
  } catch (err) {
    document.getElementById('status').innerText = 'Error al cargar: ' + err.message;
  }
}

// BÚSQUEDA O CREACIÓN DE HOJA DE CÁLCULO
async function buscarOCrearBaseDatos() {
  const response = await gapi.client.drive.files.list({
    q: `name = '${DB_FILE_NAME}' and mimeType = 'application/vnd.google-apps.spreadsheet' and trashed = false`,
    fields: 'files(id, name)',
    spaces: 'drive'
  });

  const files = response.result.files;

  if (files && files.length > 0) {
    return files[0].id;
  } else {
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
}

// INICIALIZACIÓN DE ENCABEZADOS DE PESTAÑAS
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
    'STOCK_ACTUAL', 'ESTADO_ITEM', 'ULTIMA_MODIF'
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
        { range: 'CATALOGO!A1:L1', values: [encabezadosCatalogo] },
        { range: 'LOG_HISTORICO_PRECIOS!A1:H1', values: [encabezadosHistorico] }
      ]
    }
  });
}

function poblarDataLists() {
  const dlProds = document.getElementById('dl-productos');
  dlProds.innerHTML = '';
  
  catalogoProductos.filter(p => p.estado === 'ACTIVO').forEach(p => {
    const opt = document.createElement('option');
    opt.value = `${p.marca} ${p.linea} ${p.volumen} ${p.variante}`.trim();
    opt.dataset.sku = p.sku;
    dlProds.appendChild(opt);
  });
}

function onSeleccionarProductoOmnibox(val) {
  const prod = catalogoProductos.find(p => `${p.marca} ${p.linea} ${p.volumen} ${p.variante}`.trim() === val.trim());
  if (prod) {
    document.getElementById('prod-marca').value = prod.marca;
    document.getElementById('prod-linea').value = prod.linea;
    document.getElementById('prod-magnitud').value = prod.magnitud;
    document.getElementById('prod-cant-emp').value = prod.cantEmp;
    document.getElementById('prod-volumen').value = prod.volumen;
    document.getElementById('prod-variante').value = prod.variante;
    document.getElementById('precio-distribuidor').value = prod.pDist;
    document.getElementById('precio-consumidor').value = prod.pCons;
  }
}

// GENERADOR DE SKU COMPACTO INMUTABLE
function generarSKUCompacto(marca, linea, volumen, variante) {
  const clean = (t) => t.toUpperCase().replace(/[^A-Z0-9]/g, '').substring(0, 3);
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
    const rows = res.result.values || [];
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

// GUARDAR MOVIMIENTO Y ACTUALIZAR STOCK ATÓMICO
async function guardarMovimiento(event) {
  event.preventDefault();
  const btn = document.getElementById('btn-save');
  btn.disabled = true;
  document.getElementById('status').innerText = 'Guardando registro...';

  const transId = document.getElementById('trans-id').value;
  const timestampLog = new Date().toISOString().replace('T', ' ').substring(0, 19);
  const fechaMov = document.getElementById('fecha-mov').value;
  
  const marca = document.getElementById('prod-marca').value.trim();
  const linea = document.getElementById('prod-linea').value.trim();
  const magnitud = document.getElementById('prod-magnitud').value.trim();
  const cantEmp = parseInt(document.getElementById('prod-cant-emp').value);
  const volumen = document.getElementById('prod-volumen').value.trim();
  const variante = document.getElementById('prod-variante').value.trim();

  const rawCantidad = parseFloat(document.getElementById('mov-cantidad').value);
  const cantidadSigno = (tipoMovimiento === 'VENTA') ? -Math.abs(rawCantidad) : Math.abs(rawCantidad);
  const bonif = (tipoMovimiento === 'ENTRADA') ? parseFloat(document.getElementById('mov-bonificacion').value) || 0 : 0;
  
  const pDist = parseFloat(document.getElementById('precio-distribuidor').value) || 0;
  const pCons = parseFloat(document.getElementById('precio-consumidor').value) || 0;

  const sku = generarSKUCompacto(marca, linea, volumen, variante);

  // Cálculos matemáticos
  const totalInversion = Math.abs(cantidadSigno) * pDist;
  const margenUnit = pCons - pDist;
  const totalVenta = Math.abs(cantidadSigno) * pCons;
  const utilidadNeta = Math.abs(cantidadSigno) * margenUnit;

  try {
    // 1. Guardar en LOG_TRANS
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

    // 2. Actualizar o Insertar en CATALOGO
    let prodIndex = catalogoProductos.findIndex(p => p.sku === sku);
    if (prodIndex >= 0) {
      catalogoProductos[prodIndex].stock += cantidadSigno;
      catalogoProductos[prodIndex].pDist = pDist || catalogoProductos[prodIndex].pDist;
      catalogoProductos[prodIndex].pCons = pCons || catalogoProductos[prodIndex].pCons;
    } else {
      catalogoProductos.push({
        sku, marca, linea, magnitud, cantEmp, volumen, variante,
        pDist, pCons, stock: cantidadSigno, estado: 'ACTIVO'
      });
    }

    await reescribirHojaCatalogo();

    document.getElementById('status').innerText = '¡Guardado correctamente!';
    document.getElementById('finance-form').reset();
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
    p.pDist, p.pCons, p.stock, p.estado, new Date().toISOString()
  ]);

  await gapi.client.sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID,
    range: 'CATALOGO!A2:L',
    valueInputOption: 'USER_ENTERED',
    resource: { values: rows }
  });
}

// RENDERS DE DASHBOARD & GRÁFICO
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

  // Ordenamiento por Semáforo: Desabastecido (1) -> Stock Bajo (2) -> Stock Óptimo (3)
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
        <span class="prod-sub">Dist: $${p.pDist.toFixed(2)} | Cons: $${p.pCons.toFixed(2)}</span>
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

// MODAL Y HISTORIAL LOG
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

  // Registrar en LOG_HISTORICO_PRECIOS
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
    document.getElementById('status').innerText = 'Cambios de producto guardados en el LOG.';
  } catch (err) {
    alert('Error al guardar en LOG: ' + err.message);
  }
}
