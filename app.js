// CONFIGURACIÓN CENTRAL
const CLIENT_ID = '1070607567316-mdbd97lbkprgpc4spj71e5f8anovr6it.apps.googleusercontent.com';
const SCOPES = 'https://www.googleapis.com/auth/drive.file https://www.googleapis.com/auth/spreadsheets';
const DB_FILE_NAME = 'StockCentral_DB';

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
let stockChartInstance = null;
let estadoModalSeleccionado = 'ACTIVO';
let modoDashboard = 'VENTA';
// Umbrales de semáforo por Tipo de Empaque. Ej: { "Caja": {min:4, max:6}, "Sixpack": {min:10, max:20} }
let umbralesPorEmpaque = {};

window.addEventListener('DOMContentLoaded', () => {
  const fechaInput = document.getElementById('fecha-mov');
  if (fechaInput) fechaInput.valueAsDate = new Date();

  const filtroFechaVentas = document.getElementById('filtro-fecha-ventas');
  if (filtroFechaVentas) filtroFechaVentas.valueAsDate = new Date();

  iniciarRelojCaptura();
  inicializarAutocompletesEntrada();
  inicializarBuscadorVenta();
  inicializarPreviaGanancia();
  inicializarValidacionStock();
});

// GANANCIA POR UNIDAD DE EMPAQUE + MARGEN % (informativo, no se escribe en Excel)
function inicializarPreviaGanancia() {
  ['precio-distribuidor', 'precio-consumidor'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.addEventListener('input', actualizarPreviaGanancia);
  });
}

function actualizarPreviaGanancia() {
  const pDist = parseFloat(document.getElementById('precio-distribuidor').value) || 0;
  const pCons = parseFloat(document.getElementById('precio-consumidor').value) || 0;
  const tipoEmp = document.getElementById('prod-tipo-empaque').value.trim();
  const ganancia = pCons - pDist;
  const margenPct = pDist > 0 ? (ganancia / pDist) * 100 : 0;

  document.getElementById('profit-unit-label').textContent = tipoEmp || 'Unidad';
  document.getElementById('profit-ganancia').textContent = `$${ganancia.toFixed(2)}`;
  document.getElementById('profit-margen').textContent = `${margenPct.toFixed(1)}%`;
}

// VALIDACIÓN DE STOCK DISPONIBLE EN VENTA
function inicializarValidacionStock() {
  const input = document.getElementById('mov-cantidad-out');
  if (!input) return;
  input.addEventListener('input', () => {
    if (!productoVentaSeleccionado) return;
    const val = parseFloat(input.value);
    if (!isNaN(val) && val > productoVentaSeleccionado.stock) {
      mostrarDialogo({
        titulo: 'Stock insuficiente',
        mensaje: `Disponible: ${productoVentaSeleccionado.stock} ${productoVentaSeleccionado.tipoEmpaque || ''}.`.trim()
      });
      input.value = productoVentaSeleccionado.stock;
    }
  });
}

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

// DIÁLOGO PERSONALIZADO (reemplaza alert()/confirm() nativos, sin el encabezado del navegador)
function mostrarDialogo({ titulo = 'Aviso', mensaje = '', textoConfirmar = 'Entendido', textoCancelar = null, onConfirmar = null, onCancelar = null }) {
  document.getElementById('dialogo-titulo').textContent = titulo;
  document.getElementById('dialogo-mensaje').textContent = mensaje;

  const btnConfirmar = document.getElementById('dialogo-btn-confirmar');
  const btnCancelar = document.getElementById('dialogo-btn-cancelar');

  btnConfirmar.textContent = textoConfirmar;
  btnConfirmar.onclick = () => {
    cerrarDialogo();
    if (onConfirmar) onConfirmar();
  };

  if (textoCancelar) {
    btnCancelar.textContent = textoCancelar;
    btnCancelar.classList.remove('hidden');
    btnCancelar.onclick = () => {
      cerrarDialogo();
      if (onCancelar) onCancelar();
    };
  } else {
    btnCancelar.classList.add('hidden');
  }

  document.getElementById('dialogo-personalizado').classList.remove('hidden');
}

function cerrarDialogo() {
  document.getElementById('dialogo-personalizado').classList.add('hidden');
}

// TOOLTIPS INFORMATIVOS (ⓘ) — funcionan con tap (móvil) y hover (desktop)
function toggleInfoTip(el, texto) {
  const existente = el.querySelector('.info-tip');
  cerrarTodosLosInfoTips();
  if (existente) return; // si ya estaba abierto, el cierre de arriba basta (toggle)

  const tip = document.createElement('div');
  tip.className = 'info-tip';
  tip.textContent = texto;
  el.appendChild(tip);

  setTimeout(() => {
    document.addEventListener('click', function cerrarAlTocarFuera(e) {
      if (!el.contains(e.target)) {
        cerrarTodosLosInfoTips();
        document.removeEventListener('click', cerrarAlTocarFuera);
      }
    });
  }, 0);
}

function cerrarTodosLosInfoTips() {
  document.querySelectorAll('.info-tip').forEach(t => t.remove());
}

// MOSTRAR/OCULTAR LAS OPCIONES DE IMPORTACIÓN (Plantilla / Importar Excel)
function toggleOpcionesImportar() {
  document.getElementById('opciones-importar').classList.toggle('hidden');
  document.getElementById('hint-importar').classList.toggle('hidden');
}

// ===== IMPORTACIÓN MASIVA DE ENTRADAS DESDE EXCEL =====
const COLUMNAS_IMPORTACION = [
  'Fecha', 'Marca', 'Producto', 'Tipo de Empaque', 'Cantidad', 'Presentacion',
  'Variante', 'Contenido_Vol', 'Cantidad_Bonificada', 'Precio_Distribuidor', 'Precio_Consumidor'
];

function descargarPlantillaExcel() {
  const hoy = new Date().toISOString().split('T')[0];
  const ws = XLSX.utils.aoa_to_sheet([
    COLUMNAS_IMPORTACION,
    [hoy, 'La Constancia', 'Pepsi', 'Caja', 10, 'Botella', 'Uva', '1.5 Lts', 0, 10, 12]
  ]);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Plantilla');
  XLSX.writeFile(wb, 'Plantilla_Importar_Entradas.xlsx');
}

// Convierte lo que venga en la columna Fecha (fecha nativa de Excel, texto, o número de serie) a 'YYYY-MM-DD'
function normalizarFechaImportada(valor) {
  if (valor instanceof Date && !isNaN(valor)) {
    return valor.toISOString().split('T')[0];
  }
  if (typeof valor === 'number') {
    const f = XLSX.SSF.parse_date_code(valor);
    if (f) return `${f.y}-${String(f.m).padStart(2, '0')}-${String(f.d).padStart(2, '0')}`;
  }
  if (typeof valor === 'string' && valor.trim()) {
    const texto = valor.trim();
    if (/^\d{4}-\d{2}-\d{2}$/.test(texto)) return texto;
    const parseada = new Date(texto);
    if (!isNaN(parseada)) return parseada.toISOString().split('T')[0];
  }
  return null;
}

function onArchivoExcelSeleccionado(event) {
  const file = event.target.files[0];
  event.target.value = ''; // permite volver a seleccionar el mismo archivo después
  if (!file) return;

  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const data = new Uint8Array(e.target.result);
      const wb = XLSX.read(data, { type: 'array', cellDates: true });
      const hoja = wb.Sheets[wb.SheetNames[0]];
      const filas = XLSX.utils.sheet_to_json(hoja, { defval: '' });
      procesarFilasImportadas(filas);
    } catch (err) {
      mostrarDialogo({ titulo: 'Error al leer el archivo', mensaje: err.message });
    }
  };
  reader.readAsArrayBuffer(file);
}

function procesarFilasImportadas(filas) {
  const validas = [];
  const errores = [];
  const hoy = new Date().toISOString().split('T')[0];

  // Arranca con los SKUs que YA existen en el catálogo, y va sumando los que se crean
  // con las Entradas de este mismo archivo, en el mismo orden en que aparecen las filas.
  const skusConocidos = new Set(catalogoProductos.map(p => p.sku));

  filas.forEach((fila, i) => {
    const numFila = i + 2; // fila 1 es el encabezado
    const marca = String(fila['Marca'] || '').trim();
    const linea = String(fila['Producto'] || '').trim();
    const tipoEmpaque = String(fila['Tipo de Empaque'] || '').trim();
    const cantidad = parseFloat(fila['Cantidad']);
    const presentacion = String(fila['Presentacion'] || '').trim();
    const variante = String(fila['Variante'] || '').trim();
    const volumen = String(fila['Contenido_Vol'] || '').trim();
    const bonif = parseFloat(fila['Cantidad_Bonificada']) || 0;
    const pDist = parseFloat(fila['Precio_Distribuidor']);
    const pCons = parseFloat(fila['Precio_Consumidor']);

    const fechaCruda = fila['Fecha'];
    let fecha = hoy;
    if (fechaCruda !== '' && fechaCruda !== undefined && fechaCruda !== null) {
      const fechaNormalizada = normalizarFechaImportada(fechaCruda);
      if (!fechaNormalizada) {
        errores.push(`Fila ${numFila}: la fecha "${fechaCruda}" no se pudo interpretar.`);
        return;
      }
      fecha = fechaNormalizada;
    }

    if (!marca || !linea || !presentacion || !volumen) {
      errores.push(`Fila ${numFila}: faltan Marca, Producto, Presentación o Contenido/Vol.`);
      return;
    }
    if (isNaN(cantidad) || cantidad === 0) {
      errores.push(`Fila ${numFila}: Cantidad inválida o en cero.`);
      return;
    }

    const esVenta = cantidad < 0;
    const skuFila = generarSKUCompacto(marca, linea, volumen, variante);

    if (esVenta) {
      if (!skusConocidos.has(skuFila)) {
        errores.push(`Fila ${numFila}: cantidad negativa (venta), pero "${marca} ${linea}" no existe en el catálogo ni fue creado antes en este mismo archivo.`);
        return;
      }
    } else {
      if (!tipoEmpaque || isNaN(pDist) || isNaN(pCons)) {
        errores.push(`Fila ${numFila}: Entrada requiere Tipo de Empaque y Precio Distribuidor/Consumidor válidos.`);
        return;
      }
      skusConocidos.add(skuFila); // ya queda disponible para filas de venta más abajo en el mismo archivo
    }

    validas.push({
      fecha, marca, linea, tipoEmpaque, cantidad: Math.abs(cantidad),
      tipoMov: esVenta ? 'VENTA' : 'ENTRADA',
      presentacion, variante, volumen, bonif, pDist, pCons
    });
  });

  if (validas.length === 0) {
    mostrarDialogo({
      titulo: 'Nada para importar',
      mensaje: errores.length > 0
        ? `No se encontró ninguna fila válida.\n\n${errores.slice(0, 5).join('\n')}`
        : 'El archivo no tiene filas de datos.'
    });
    return;
  }

  let mensaje = `Se encontraron ${validas.length} fila(s) válidas para importar.`;
  if (errores.length > 0) {
    mensaje += `\n\n${errores.length} fila(s) con problemas se omitirán:\n${errores.slice(0, 5).join('\n')}`;
    if (errores.length > 5) mensaje += `\n...y ${errores.length - 5} más.`;
  }

  mostrarDialogo({
    titulo: 'Confirmar importación',
    mensaje,
    textoConfirmar: `Importar ${validas.length} fila(s)`,
    textoCancelar: 'Cancelar',
    onConfirmar: () => ejecutarImportacionMasiva(validas)
  });
}

async function ejecutarImportacionMasiva(filasValidas) {
  document.getElementById('status').innerText = `Importando ${filasValidas.length} movimiento(s)...`;

  // Continuar la numeración de folio (IN- y FAC- por separado) desde el máximo ya usado
  let maxIN = 0, maxFAC = 0;
  try {
    const res = await gapi.client.sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: 'LOG_TRANS!A:A',
    });
    const rows = res.result ? res.result.values || [] : [];
    rows.forEach(r => {
      if (r[0] && r[0].startsWith('IN-')) {
        const num = parseInt(r[0].replace('IN-', ''));
        if (!isNaN(num) && num > maxIN) maxIN = num;
      } else if (r[0] && r[0].startsWith('FAC-')) {
        const num = parseInt(r[0].replace('FAC-', ''));
        if (!isNaN(num) && num > maxFAC) maxFAC = num;
      }
    });
  } catch (e) { /* si falla, continúa desde 0 */ }

  const filasLog = [];

  filasValidas.forEach(f => {
    const timestampLog = new Date().toISOString().replace('T', ' ').substring(0, 19);
    const sku = generarSKUCompacto(f.marca, f.linea, f.volumen, f.variante);

    if (f.tipoMov === 'ENTRADA') {
      maxIN++;
      const transId = `IN-${String(maxIN).padStart(6, '0')}`;
      const totalInversion = f.cantidad * f.pDist;
      const margenUnit = f.pCons - f.pDist;

      filasLog.push([
        transId, timestampLog, f.fecha, 'ENTRADA', sku,
        f.marca, f.linea, f.presentacion, f.volumen, f.variante, 1,
        f.cantidad, f.bonif, f.pDist, f.pCons,
        totalInversion, margenUnit, '', '', f.tipoEmpaque
      ]);

      const prodIndex = catalogoProductos.findIndex(p => p.sku === sku);
      if (prodIndex >= 0) {
        catalogoProductos[prodIndex].stock += f.cantidad;
        catalogoProductos[prodIndex].pDist = f.pDist;
        catalogoProductos[prodIndex].pCons = f.pCons;
        catalogoProductos[prodIndex].tipoEmpaque = f.tipoEmpaque;
      } else {
        catalogoProductos.push({
          sku, marca: f.marca, linea: f.linea, magnitud: f.presentacion,
          cantEmp: 1, volumen: f.volumen, variante: f.variante,
          pDist: f.pDist, pCons: f.pCons, stock: f.cantidad, estado: 'ACTIVO',
          tipoEmpaque: f.tipoEmpaque, idProducto: generarIdProducto()
        });
      }
    } else {
      // VENTA: usa el precio y tipo de empaque ACTUALES del catálogo, igual que una venta manual
      const prod = catalogoProductos.find(p => p.sku === sku);
      if (!prod) return; // ya validado antes; por seguridad, se omite si no aparece

      maxFAC++;
      const transId = `FAC-${String(maxFAC).padStart(6, '0')}`;
      const margenUnit = prod.pCons - prod.pDist;
      const totalInversion = f.cantidad * prod.pDist;
      const totalVenta = f.cantidad * prod.pCons;
      const utilidadNeta = f.cantidad * margenUnit;

      filasLog.push([
        transId, timestampLog, f.fecha, 'VENTA', sku,
        prod.marca, prod.linea, prod.magnitud, prod.volumen, prod.variante, prod.cantEmp,
        -f.cantidad, 0, prod.pDist, prod.pCons,
        totalInversion, margenUnit, totalVenta, utilidadNeta, prod.tipoEmpaque
      ]);

      prod.stock -= f.cantidad;
    }
  });

  try {
    await gapi.client.sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: 'LOG_TRANS!A:T',
      valueInputOption: 'USER_ENTERED',
      resource: { values: filasLog }
    });

    await reescribirHojaCatalogo();
    await cargarDatosIniciales();

    document.getElementById('status').innerText = `¡${filasValidas.length} movimiento(s) importados correctamente!`;
    mostrarDialogo({ titulo: 'Importación completada', mensaje: `Se importaron ${filasValidas.length} movimiento(s) correctamente.` });
  } catch (err) {
    document.getElementById('status').innerText = 'Error al importar: ' + err.message;
    mostrarDialogo({ titulo: 'Error al importar', mensaje: err.message });
  }
}

function limpiarFormularioActivo() {
  if (tipoMovimiento === 'ENTRADA') {
    document.getElementById('form-entrada').reset();
    actualizarPreviaGanancia();
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
function crearAutocomplete(inputId, dropdownId, getOpciones, onSeleccionar, renderItem, hintId) {
  const input = document.getElementById(inputId);
  const dropdown = document.getElementById(dropdownId);
  const hint = hintId ? document.getElementById(hintId) : null;
  if (!input || !dropdown) return;

  let opcionesActuales = [];
  let indiceResaltado = -1;

  function marcarResaltado() {
    Array.from(dropdown.children).forEach((el, i) => {
      const activo = i === indiceResaltado;
      el.classList.toggle('highlighted', activo);
      if (activo) el.scrollIntoView({ block: 'nearest' });
    });
  }

  function seleccionar(opt) {
    onSeleccionar(opt);
    dropdown.classList.add('hidden');
    if (hint) hint.classList.add('hidden');
  }

  function actualizarHint(query) {
    if (!hint) return;
    const q = query.trim().toLowerCase();
    if (!q) { hint.classList.add('hidden'); return; }
    const todasLasOpciones = getOpciones('');
    const existeExacto = todasLasOpciones.some(v => typeof v === 'string' && v.toLowerCase() === q);
    hint.classList.toggle('hidden', existeExacto);
  }

  function render(query) {
    opcionesActuales = getOpciones(query.trim());
    indiceResaltado = opcionesActuales.length > 0 ? 0 : -1;
    dropdown.innerHTML = '';

    if (!query.trim() || opcionesActuales.length === 0) {
      dropdown.classList.add('hidden');
    } else {
      opcionesActuales.slice(0, 8).forEach((opt, i) => {
        const item = document.createElement('div');
        item.className = 'autocomplete-item' + (i === 0 ? ' highlighted' : '');
        item.innerHTML = renderItem(opt);
        item.addEventListener('mousedown', (e) => {
          e.preventDefault();
          seleccionar(opt);
        });
        dropdown.appendChild(item);
      });
      dropdown.classList.remove('hidden');
    }
    actualizarHint(query);
  }

  input.addEventListener('input', () => render(input.value));
  input.addEventListener('focus', () => { if (input.value.trim()) render(input.value); });
  input.addEventListener('blur', () => setTimeout(() => { dropdown.classList.add('hidden'); }, 150));

  input.addEventListener('keydown', (e) => {
    if (dropdown.classList.contains('hidden') || opcionesActuales.length === 0) return;
    const maxIndex = Math.min(opcionesActuales.length, 8) - 1;

    if (e.key === 'ArrowDown') {
      e.preventDefault();
      indiceResaltado = Math.min(indiceResaltado + 1, maxIndex);
      marcarResaltado();
    } else if (e.key === 'ArrowUp') {
      e.preventDefault();
      indiceResaltado = Math.max(indiceResaltado - 1, 0);
      marcarResaltado();
    } else if (e.key === 'Enter') {
      if (indiceResaltado >= 0) {
        e.preventDefault();
        seleccionar(opcionesActuales[indiceResaltado]);
      }
    } else if (e.key === 'Escape') {
      dropdown.classList.add('hidden');
    }
  });
}

function sugerirValoresUnicos(campo, query) {
  const q = query.toLowerCase();
  const valores = [...new Set(catalogoProductos.map(p => p[campo]).filter(Boolean))];
  return valores.filter(v => v.toLowerCase().includes(q));
}

function inicializarAutocompletesEntrada() {
  crearAutocomplete('prod-marca', 'ac-marca', q => sugerirValoresUnicos('marca', q),
    v => { document.getElementById('prod-marca').value = v; }, v => v, 'hint-marca');
  crearAutocomplete('prod-linea', 'ac-linea', q => sugerirValoresUnicos('linea', q),
    v => { document.getElementById('prod-linea').value = v; }, v => v, 'hint-linea');
  crearAutocomplete('prod-tipo-empaque', 'ac-tipoempaque', q => sugerirValoresUnicos('tipoEmpaque', q),
    v => { document.getElementById('prod-tipo-empaque').value = v; actualizarPreviaGanancia(); }, v => v, 'hint-tipoempaque');
  crearAutocomplete('prod-presentacion', 'ac-presentacion', q => sugerirValoresUnicos('magnitud', q),
    v => { document.getElementById('prod-presentacion').value = v; }, v => v, 'hint-presentacion');
  crearAutocomplete('prod-variante', 'ac-variante', q => sugerirValoresUnicos('variante', q),
    v => { document.getElementById('prod-variante').value = v; }, v => v, 'hint-variante');
  crearAutocomplete('prod-volumen', 'ac-volumen', q => sugerirValoresUnicos('volumen', q),
    v => { document.getElementById('prod-volumen').value = v; }, v => v, 'hint-volumen');
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
      range: 'CATALOGO!A2:N',
    });

    const rows = res.result ? res.result.values || [] : [];
    catalogoProductos = rows.map(r => ({
      sku: r[0], marca: r[1], linea: r[2], magnitud: r[3],
      cantEmp: parseInt(r[4]) || 1, volumen: r[5], variante: r[6] || '',
      pDist: parseFloat(r[7]) || 0, pCons: parseFloat(r[8]) || 0,
      stock: parseFloat(r[9]) || 0, estado: r[10] || 'ACTIVO',
      tipoEmpaque: r[12] || '',
      idProducto: parseInt(r[13]) || null
    }));

    await cargarUmbrales();

    generarFolioCorrelativo();
    document.getElementById('status').innerText = 'Conectado a la base de datos.';
  } catch (err) {
    document.getElementById('status').innerText = 'Error al cargar: ' + err.message;
  }
}

// UMBRALES DE SEMÁFORO POR TIPO DE EMPAQUE
async function cargarUmbrales() {
  const res = await gapi.client.sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: 'UMBRALES!A2:C',
  });
  const rows = res.result ? res.result.values || [] : [];
  umbralesPorEmpaque = {};
  rows.forEach(r => {
    if (!r[0]) return;
    const min = parseFloat(r[1]);
    const max = parseFloat(r[2]);
    umbralesPorEmpaque[r[0]] = {
      min: isNaN(min) ? DEFAULT_STOCK_MIN : min,
      max: isNaN(max) ? DEFAULT_STOCK_MAX : max
    };
  });
}

function obtenerUmbral(tipoEmpaque) {
  return umbralesPorEmpaque[tipoEmpaque] || { min: DEFAULT_STOCK_MIN, max: DEFAULT_STOCK_MAX };
}

async function reescribirHojaUmbrales() {
  const rows = Object.keys(umbralesPorEmpaque).map(tipo => [
    tipo, umbralesPorEmpaque[tipo].min, umbralesPorEmpaque[tipo].max
  ]);
  await gapi.client.sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID,
    range: 'UMBRALES!A2:C',
    valueInputOption: 'USER_ENTERED',
    resource: { values: rows }
  });
}

// GENERA UN ID DE PRODUCTO CORRELATIVO Y ÚNICO (no UUID: volumen bajo, se prioriza legibilidad)
function generarIdProducto() {
  const maxId = catalogoProductos.reduce((max, p) => Math.max(max, p.idProducto || 0), 0);
  return maxId + 1;
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
  const existente = await buscarArchivoPorNombre(DB_FILE_NAME);
  if (existente) return existente.id;

  document.getElementById('status').innerText = 'Creando nueva base de datos en tu Drive...';

  const createRes = await gapi.client.sheets.spreadsheets.create({
    resource: {
      properties: { title: DB_FILE_NAME },
      sheets: [
        { properties: { title: 'LOG_TRANS' } },
        { properties: { title: 'CATALOGO' } },
        { properties: { title: 'LOG_HISTORICO_PRECIOS' } },
        { properties: { title: 'UMBRALES' } }
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
    'MARCA', 'LINEA_PROD', 'PRESENTACION', 'VOLUMEN', 'VARIANTE', 'CANT_EMPAQUE',
    'CANTIDAD', 'UNID_BONIF', 'P_DISTRIBUIDOR', 'P_CONSUMIDOR',
    'TOTAL_INVERSION', 'MARGEN_UNIT', 'TOTAL_VENTA', 'UTILIDAD_NETA', 'TIPO_EMPAQUE'
  ];

  const encabezadosCatalogo = [
    'SKU_ITEM', 'MARCA', 'LINEA_PROD', 'PRESENTACION', 'CANT_EMPAQUE',
    'VOLUMEN', 'VARIANTE', 'P_DISTRIBUIDOR', 'P_CONSUMIDOR',
    'STOCK_ACTUAL', 'ESTADO_ITEM', 'ULTIMA_MODIF',
    'TIPO_EMPAQUE', 'ID_PRODUCTO'
  ];

  const encabezadosHistorico = [
    'ID_LOG', 'SKU_ITEM', 'P_DIST_ANT', 'P_DIST_NUEVO',
    'P_CONS_ANT', 'P_CONS_NUEVO', 'FECHA_CAMBIO', 'USUARIO'
  ];

  const encabezadosUmbrales = ['TIPO_EMPAQUE', 'STOCK_MIN', 'STOCK_MAX'];

  await gapi.client.sheets.spreadsheets.values.batchUpdate({
    spreadsheetId: spreadsheetId,
    resource: {
      valueInputOption: 'USER_ENTERED',
      data: [
        { range: 'LOG_TRANS!A1:T1', values: [encabezadosLOG] },
        { range: 'CATALOGO!A1:N1', values: [encabezadosCatalogo] },
        { range: 'LOG_HISTORICO_PRECIOS!A1:H1', values: [encabezadosHistorico] },
        { range: 'UMBRALES!A1:C1', values: [encabezadosUmbrales] }
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
      mostrarDialogo({ titulo: 'Falta el producto', mensaje: 'Debes seleccionar un producto válido antes de guardar la venta.' });
      btn.disabled = false;
      return;
    }
    rawCantidad = parseFloat(document.getElementById('mov-cantidad-out').value);
    if (rawCantidad > productoVentaSeleccionado.stock) {
      mostrarDialogo({
        titulo: 'Stock insuficiente',
        mensaje: `Disponible: ${productoVentaSeleccionado.stock} ${productoVentaSeleccionado.tipoEmpaque || ''}.`.trim()
      });
      btn.disabled = false;
      return;
    }
    marca = productoVentaSeleccionado.marca;
    linea = productoVentaSeleccionado.linea;
    magnitud = productoVentaSeleccionado.magnitud;
    cantEmp = productoVentaSeleccionado.cantEmp;
    volumen = productoVentaSeleccionado.volumen;
    variante = productoVentaSeleccionado.variante;
    tipoEmpaque = productoVentaSeleccionado.tipoEmpaque || '';
    bonif = 0;
    pDist = productoVentaSeleccionado.pDist;
    pCons = productoVentaSeleccionado.pCons;
    sku = productoVentaSeleccionado.sku;
  }

  const cantidadSigno = (tipoMovimiento === 'VENTA') ? -Math.abs(rawCantidad) : Math.abs(rawCantidad);
  const totalInversion = Math.abs(cantidadSigno) * pDist;
  const margenUnit = pCons - pDist;
  const totalVenta = (tipoMovimiento === 'VENTA') ? Math.abs(cantidadSigno) * pCons : '';
  const utilidadNeta = (tipoMovimiento === 'VENTA') ? Math.abs(cantidadSigno) * margenUnit : '';

  try {
    await gapi.client.sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: 'LOG_TRANS!A:T',
      valueInputOption: 'USER_ENTERED',
      resource: {
        values: [[
          transId, timestampLog, fechaMov, tipoMovimiento, sku,
          marca, linea, magnitud, volumen, variante, cantEmp,
          cantidadSigno, bonif, pDist, pCons,
          totalInversion, margenUnit, totalVenta, utilidadNeta, tipoEmpaque
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
        idProducto: generarIdProducto()
      });
    }

    await reescribirHojaCatalogo();

    document.getElementById('status').innerText = '¡Guardado correctamente!';
    if (tipoMovimiento === 'ENTRADA') {
      document.getElementById('form-entrada').reset();
      actualizarPreviaGanancia();
    } else {
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
    p.tipoEmpaque || '', p.idProducto
  ]);

  await gapi.client.sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID,
    range: 'CATALOGO!A2:N',
    valueInputOption: 'USER_ENTERED',
    resource: { values: rows }
  });
}

// RENDERS STOCK & VENTAS
function renderizarDashboardStock() {
  let valorTotalBodega = 0;
  let totalEnBodega = 0;

  catalogoProductos.forEach(p => {
    if (p.estado === 'ACTIVO') {
      valorTotalBodega += (p.stock * p.pDist);
      totalEnBodega += p.stock;
    }
  });

  document.getElementById('metric-valor-bodega').innerText = `$${valorTotalBodega.toFixed(2)}`;
  document.getElementById('metric-total-unidades').innerText = totalEnBodega.toString();

  renderizarGraficoStock();
  poblarFiltroEmpaque();
}

// FILTRO DE TIPO DE EMPAQUE + UMBRALES DE SEMÁFORO EN VIVO
function poblarFiltroEmpaque() {
  const select = document.getElementById('filter-tipo-empaque');
  if (!select) return;

  const tiposDistintos = [...new Set(catalogoProductos.map(p => p.tipoEmpaque).filter(Boolean))].sort();
  const valorAnterior = select.value;

  select.innerHTML = '<option value="TODOS">Todos los productos</option>' +
    tiposDistintos.map(t => `<option value="${t}">${t}</option>`).join('');

  const opcionesDisponibles = [...select.options].map(o => o.value);
  if (valorAnterior && opcionesDisponibles.includes(valorAnterior)) {
    select.value = valorAnterior;
  } else if (opcionesDisponibles.includes('Caja')) {
    select.value = 'Caja';
  } else if (opcionesDisponibles.includes('Cajas')) {
    select.value = 'Cajas';
  } else {
    select.value = 'TODOS';
  }

  onCambioFiltroEmpaque();
}

function onCambioFiltroEmpaque() {
  const filtro = document.getElementById('filter-tipo-empaque').value;
  const wrapper = document.getElementById('umbral-inputs-wrapper');

  if (filtro === 'TODOS') {
    wrapper.classList.add('hidden');
  } else {
    wrapper.classList.remove('hidden');
    const u = obtenerUmbral(filtro);
    document.getElementById('umbral-stock-bajo').value = u.min;
    document.getElementById('umbral-stock-alto').value = u.max;
  }
  renderizarListaStock();
}

async function onCambioUmbral() {
  const filtro = document.getElementById('filter-tipo-empaque').value;
  if (filtro === 'TODOS') return;

  const min = parseFloat(document.getElementById('umbral-stock-bajo').value);
  const max = parseFloat(document.getElementById('umbral-stock-alto').value);
  umbralesPorEmpaque[filtro] = {
    min: isNaN(min) ? DEFAULT_STOCK_MIN : min,
    max: isNaN(max) ? DEFAULT_STOCK_MAX : max
  };

  renderizarListaStock();
  try {
    await reescribirHojaUmbrales();
  } catch (err) {
    mostrarDialogo({ titulo: 'Error al guardar umbral', mensaje: err.message });
  }
}

function renderizarListaStock() {
  const container = document.getElementById('lista-productos-stock');
  const mostrarObsoletos = document.getElementById('chk-mostrar-obsoletos').checked;
  const filtroEmpaque = document.getElementById('filter-tipo-empaque').value;
  container.innerHTML = '';

  let lista = catalogoProductos.filter(p => mostrarObsoletos || p.estado === 'ACTIVO');
  if (filtroEmpaque !== 'TODOS') {
    lista = lista.filter(p => p.tipoEmpaque === filtroEmpaque);
  }

  lista = lista.sort((a, b) => {
    const prioridad = (p) => {
      const u = obtenerUmbral(p.tipoEmpaque);
      return (p.stock <= u.min) ? 1 : (p.stock <= u.max) ? 2 : 3;
    };
    return prioridad(a) - prioridad(b);
  });

  lista.forEach(p => {
    const u = obtenerUmbral(p.tipoEmpaque);
    const semaforoClase = (p.stock <= u.min) ? 'desabastecido' : (p.stock <= u.max) ? 'stock-bajo' : 'stock-optimo';
    const semaforoTexto = (p.stock <= u.min) ? '🔴 Desabastecido' : (p.stock <= u.max) ? '🟡 Stock Bajo' : '🟢 Stock Óptimo';
    const valorBodega = p.stock * p.pDist;

    const div = document.createElement('div');
    div.className = `product-item ${semaforoClase}`;
    div.innerHTML = `
      <div class="prod-info">
        <span class="prod-title">${p.marca} ${p.linea} ${p.volumen} ${p.variante}</span>
        <span class="prod-sub">SKU: ${p.sku} | Stock: <strong>${p.stock} ${p.magnitud}</strong></span>
        <span class="prod-sub">Dist: $${p.pDist.toFixed(5)} | Cons: $${p.pCons.toFixed(2)} | Bod: $${valorBodega.toFixed(2)}</span>
        <span class="prod-badge">${semaforoTexto} (${p.estado})</span>
      </div>
      <div>
        <button class="btn-secondary-sm" onclick="abrirModalEdicion(${p.idProducto})">✏️ Editar</button>
      </div>
    `;
    container.appendChild(div);
  });
}

// GRÁFICO — VISTA 1: EXPLORAR (pastel jerárquico con drill-down) / VISTA 2: VER TODO (barras)
const NIVELES_JERARQUIA = ['marca', 'linea', 'magnitud', 'volumen', 'variante'];
let chartJerarquiaPath = [];

function renderizarGraficoStock() {
  const vista = document.getElementById('filter-chart-group').value;
  if (vista === 'TODO') {
    renderizarGraficoTodo();
  } else {
    renderizarGraficoJerarquico();
  }
}

function onCambioVistaGrafico() {
  chartJerarquiaPath = [];
  renderizarGraficoStock();
}

function irANivelBreadcrumb(index) {
  chartJerarquiaPath = chartJerarquiaPath.slice(0, index + 1);
  renderizarGraficoJerarquico();
}

function renderizarBreadcrumb() {
  const cont = document.getElementById('chart-breadcrumb');
  if (chartJerarquiaPath.length === 0) {
    cont.innerHTML = '<span class="breadcrumb-item active">Todas las marcas</span>';
    return;
  }
  let html = `<span class="breadcrumb-item" onclick="irANivelBreadcrumb(-1)">Todas</span>`;
  chartJerarquiaPath.forEach((paso, i) => {
    const esUltimo = i === chartJerarquiaPath.length - 1;
    html += ` <span class="breadcrumb-sep">›</span> <span class="breadcrumb-item ${esUltimo ? 'active' : ''}" onclick="irANivelBreadcrumb(${i})">${paso.valor}</span>`;
  });
  cont.innerHTML = html;
}

function renderizarGraficoJerarquico() {
  const nivelActual = NIVELES_JERARQUIA[chartJerarquiaPath.length];
  if (!nivelActual) {
    chartJerarquiaPath.pop(); // ya no hay más niveles para explorar, no hacer nada más
    return;
  }

  const productosFiltrados = catalogoProductos.filter(p =>
    p.estado === 'ACTIVO' && chartJerarquiaPath.every(paso => p[paso.nivel] === paso.valor)
  );

  const agrupado = {};
  productosFiltrados.forEach(p => {
    const key = p[nivelActual] || '(Sin dato)';
    agrupado[key] = (agrupado[key] || 0) + p.stock;
  });

  renderizarBreadcrumb();

  const etiquetas = Object.keys(agrupado);
  const ctx = document.getElementById('stockPieChart').getContext('2d');
  if (stockChartInstance) stockChartInstance.destroy();

  stockChartInstance = new Chart(ctx, {
    type: 'pie',
    data: {
      labels: etiquetas,
      datasets: [{
        data: Object.values(agrupado),
        backgroundColor: ['#22c55e', '#3b82f6', '#eab308', '#ef4444', '#a855f7', '#ec4899', '#14b8a6', '#f97316']
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { position: 'bottom', labels: { color: '#f8fafc' } } },
      onClick: (evt, elements) => {
        if (elements.length > 0) {
          chartJerarquiaPath.push({ nivel: nivelActual, valor: etiquetas[elements[0].index] });
          renderizarGraficoJerarquico();
        }
      }
    }
  });
}

function renderizarGraficoTodo() {
  document.getElementById('chart-breadcrumb').innerHTML = '';
  const productosActivos = catalogoProductos.filter(p => p.estado === 'ACTIVO');
  const marcas = [...new Set(productosActivos.map(p => p.marca))];
  const lineas = [...new Set(productosActivos.map(p => p.linea))].sort();
  const colores = ['#22c55e', '#3b82f6', '#eab308', '#ef4444', '#a855f7', '#ec4899', '#14b8a6', '#f97316'];

  const datasets = marcas.map((marca, i) => ({
    label: marca,
    data: lineas.map(linea => productosActivos
      .filter(p => p.marca === marca && p.linea === linea)
      .reduce((sum, p) => sum + p.stock, 0)),
    backgroundColor: colores[i % colores.length]
  }));

  const ctx = document.getElementById('stockPieChart').getContext('2d');
  if (stockChartInstance) stockChartInstance.destroy();

  stockChartInstance = new Chart(ctx, {
    type: 'bar',
    data: { labels: lineas, datasets: datasets },
    options: {
      indexAxis: 'y',
      responsive: true,
      maintainAspectRatio: false,
      scales: {
        x: { stacked: true, ticks: { color: '#94a3b8' }, grid: { color: '#334155' } },
        y: { stacked: true, ticks: { color: '#f8fafc' }, grid: { display: false } }
      },
      plugins: { legend: { position: 'bottom', labels: { color: '#f8fafc' } } }
    }
  });
}

// DASHBOARD: RANGO DE FECHAS
function onToggleRangoFechas() {
  const activo = document.getElementById('chk-rango-fechas').checked;
  document.getElementById('group-fecha-fin').classList.toggle('hidden', !activo);
  document.getElementById('label-fecha-inicio').textContent = activo ? 'Desde' : 'Filtrar por Fecha';
  if (activo && !document.getElementById('filtro-fecha-fin').value) {
    document.getElementById('filtro-fecha-fin').valueAsDate = new Date();
  }
  renderizarVentasHoy();
}

function resetearFiltroFechaVentas() {
  const hoy = new Date();
  document.getElementById('filtro-fecha-ventas').valueAsDate = hoy;
  if (document.getElementById('chk-rango-fechas').checked) {
    document.getElementById('filtro-fecha-fin').valueAsDate = hoy;
  }
  renderizarVentasHoy();
}

// DASHBOARD: SLIDE ENTRADAS / SALIDAS
function setModoDashboard(modo) {
  modoDashboard = modo;
  document.getElementById('btn-dash-entradas').classList.toggle('active', modo === 'ENTRADA');
  document.getElementById('btn-dash-salidas').classList.toggle('active', modo === 'VENTA');
  document.getElementById('detalle-titulo').textContent = modo === 'ENTRADA' ? 'Detalle de Entradas' : 'Detalle de Salidas';
  actualizarEtiquetasMetricas();
  renderizarVentasHoy();
}

function actualizarEtiquetasMetricas() {
  const titulo1 = document.getElementById('metric-venta-hoy-titulo');
  const titulo2 = document.getElementById('metric-utilidad-hoy-titulo');
  if (modoDashboard === 'ENTRADA') {
    titulo1.innerHTML = `Total Invertido <span class="info-icon" onclick="toggleInfoTip(this, 'Suma de lo invertido (cantidad × precio distribuidor) en las entradas del período seleccionado.')">ⓘ</span>`;
    titulo2.innerHTML = `Unidades Ingresadas <span class="info-icon" onclick="toggleInfoTip(this, 'Suma de las cantidades registradas en las entradas del período seleccionado.')">ⓘ</span>`;
  } else {
    titulo1.innerHTML = `Ventas Totales <span class="info-icon" onclick="toggleInfoTip(this, 'Suma de los ingresos por ventas (cantidad × precio consumidor) en el período seleccionado.')">ⓘ</span>`;
    titulo2.innerHTML = `Margen/Ganancia <span class="info-icon" onclick="toggleInfoTip(this, 'Suma de la utilidad neta (precio de venta menos costo) generada por las ventas del período seleccionado.')">ⓘ</span>`;
  }
}

async function renderizarVentasHoy() {
  const usaRango = document.getElementById('chk-rango-fechas').checked;
  const hoy = new Date().toISOString().split('T')[0];
  const fechaInicio = document.getElementById('filtro-fecha-ventas').value || hoy;
  const fechaFin = usaRango ? (document.getElementById('filtro-fecha-fin').value || fechaInicio) : fechaInicio;

  try {
    const res = await gapi.client.sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: 'LOG_TRANS!A2:T',
    });

    const rows = res.result ? res.result.values || [] : [];
    let total1 = 0;
    let total2 = 0;
    const lista = document.getElementById('lista-ventas-hoy');
    lista.innerHTML = '';

    rows.forEach(r => {
      const fechaMov = r[2];
      const tipo = r[3];
      if (tipo !== modoDashboard || fechaMov < fechaInicio || fechaMov > fechaFin) return;

      const cantidad = Math.abs(parseFloat(r[11]) || 0);
      let valorMostrado;

      if (modoDashboard === 'VENTA') {
        const vta = parseFloat(r[17]) || 0;
        const util = parseFloat(r[18]) || 0;
        total1 += vta;
        total2 += util;
        valorMostrado = `$${vta.toFixed(2)}`;
      } else {
        const inv = parseFloat(r[15]) || 0;
        total1 += inv;
        total2 += cantidad;
        valorMostrado = `$${inv.toFixed(2)}`;
      }

      const div = document.createElement('div');
      div.className = `product-item ${modoDashboard === 'ENTRADA' ? 'mov-entrada' : 'mov-salida'}`;
      div.innerHTML = `
        <div class="prod-info">
          <span class="prod-title">${r[5]} ${r[6]} ${r[8]}</span>
          <span class="prod-sub">Folio: ${r[0]} | Cant: ${cantidad}</span>
        </div>
        <div>
          <strong>${valorMostrado}</strong>
        </div>
      `;
      lista.appendChild(div);
    });

    document.getElementById('metric-venta-hoy').innerText = `$${total1.toFixed(2)}`;
    document.getElementById('metric-utilidad-hoy').innerText = modoDashboard === 'ENTRADA' ? total2.toString() : `$${total2.toFixed(2)}`;
  } catch (err) {
    document.getElementById('status').innerText = 'Error al cargar movimientos: ' + err.message;
  }
}

// MODAL
function abrirModalEdicion(idProducto) {
  const prod = catalogoProductos.find(p => p.idProducto === idProducto);
  if (!prod) return;

  document.getElementById('modal-id-producto').value = prod.idProducto;
  document.getElementById('modal-titulo').innerText = `Editar: ${prod.marca} ${prod.linea}`;
  document.getElementById('modal-marca').value = prod.marca;
  document.getElementById('modal-linea').value = prod.linea;
  document.getElementById('modal-presentacion').value = prod.magnitud;
  document.getElementById('modal-volumen').value = prod.volumen;
  document.getElementById('modal-variante').value = prod.variante || '';
  document.getElementById('modal-tipo-empaque').value = prod.tipoEmpaque || '';
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
  const idProducto = parseInt(document.getElementById('modal-id-producto').value);
  const prod = catalogoProductos.find(p => p.idProducto === idProducto);
  if (!prod) return;

  const nuevaMarca = document.getElementById('modal-marca').value.trim();
  const nuevaLinea = document.getElementById('modal-linea').value.trim();
  const nuevaPresentacion = document.getElementById('modal-presentacion').value.trim();
  const nuevoVolumen = document.getElementById('modal-volumen').value.trim();
  const nuevaVariante = document.getElementById('modal-variante').value.trim();
  const nuevoTipoEmpaque = document.getElementById('modal-tipo-empaque').value.trim();
  const newPDist = parseFloat(document.getElementById('modal-p-distributor').value) || 0;
  const newPCons = parseFloat(document.getElementById('modal-p-consumer').value) || 0;

  const cambioIdentidad = (
    nuevaMarca !== prod.marca ||
    nuevaLinea !== prod.linea ||
    nuevaPresentacion !== prod.magnitud ||
    nuevoVolumen !== prod.volumen ||
    nuevaVariante !== (prod.variante || '') ||
    nuevoTipoEmpaque !== (prod.tipoEmpaque || '')
  );

  const ejecutarGuardado = async () => {
    try {
      if (newPDist !== prod.pDist || newPCons !== prod.pCons) {
        await gapi.client.sheets.spreadsheets.values.append({
          spreadsheetId: SPREADSHEET_ID,
          range: 'LOG_HISTORICO_PRECIOS!A:H',
          valueInputOption: 'USER_ENTERED',
          resource: {
            values: [[
              `LOG-${Date.now()}`, prod.sku, prod.pDist, newPDist, prod.pCons, newPCons,
              new Date().toISOString(), 'USUARIO_ACTIVO'
            ]]
          }
        });
      }

      prod.marca = nuevaMarca;
      prod.linea = nuevaLinea;
      prod.magnitud = nuevaPresentacion;
      prod.volumen = nuevoVolumen;
      prod.variante = nuevaVariante;
      prod.tipoEmpaque = nuevoTipoEmpaque;
      prod.pDist = newPDist;
      prod.pCons = newPCons;
      prod.estado = estadoModalSeleccionado;

      // El ID_PRODUCTO nunca cambia (es la llave real); el SKU sí se regenera si cambió la identidad,
      // ya que es solo una etiqueta legible derivada de estos campos.
      if (cambioIdentidad) {
        prod.sku = generarSKUCompacto(nuevaMarca, nuevaLinea, nuevoVolumen, nuevaVariante);
      }

      await reescribirHojaCatalogo();
      cerrarModal();
      renderizarDashboardStock();
      document.getElementById('status').innerText = 'Cambios guardados exitosamente.';
    } catch (err) {
      mostrarDialogo({ titulo: 'Error al guardar', mensaje: 'No se pudo guardar: ' + err.message });
    }
  };

  if (cambioIdentidad) {
    mostrarDialogo({
      titulo: '⚠️ Dato sensible',
      mensaje: 'Cambiar Marca, Producto, Presentación, Contenido, Variante o Tipo de Empaque actualiza el SKU de este producto. Si el cambio es grande (ej. pasar de Coca-Cola a Pepsi), es mejor crear un ítem nuevo y marcar este como obsoleto. ¿Continuar de todas formas?',
      textoConfirmar: 'Sí, guardar cambios',
      textoCancelar: 'Cancelar',
      onConfirmar: ejecutarGuardado
    });
  } else {
    await ejecutarGuardado();
  }
}
