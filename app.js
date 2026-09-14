// CONFIGURACIÓN PRIVADA
const CLIENT_ID = '416407370193-ugjvoa045re61hu1epo44l0v4q5mec8c.apps.googleusercontent.com'; // <-- REEMPLAZA CON TU CLIENT ID REAL
const SCOPES = 'https://www.googleapis.com/auth/drive.file https://www.googleapis.com/auth/spreadsheets';
const SPREADSHEET_NAME = 'Mi_Control_Financiero_Privado';

let tokenClient;
let gapiInited = false;
let gsisInited = false;
let spreadsheetId = null;

// Fijar fecha de hoy por defecto si existe el campo
window.addEventListener('DOMContentLoaded', () => {
  const fechaInput = document.getElementById('fecha');
  if (fechaInput) fechaInput.valueAsDate = new Date();
});

function gapiLoaded() {
  gapi.load('client', async () => {
    await gapi.client.init({});
    await gapi.client.load('https://sheets.googleapis.com/$discovery/rest?version=v4');
    await gapi.client.load('https://www.googleapis.com/discovery/v1/apis/drive/v3/rest');
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

window.onload = () => { 
  gapiLoaded(); 
  gisLoaded(); 
};

function checkAuthReady() {
  if (gapiInited && gsisInited) {
    const status = document.getElementById('status');
    if (status) status.innerText = 'Listo para conectar.';
  }
}

function handleAuthClick() {
  tokenClient.callback = async (resp) => {
    if (resp.error) throw (resp);
    document.getElementById('auth-section').classList.add('hidden');
    document.getElementById('finance-form').classList.remove('hidden');
    document.getElementById('status').innerText = 'Verificando archivo en tu Drive...';
    await inicializarHojaCalculo();
  };

  if (gapi.client.getToken() === null) {
    tokenClient.requestAccessToken({prompt: 'consent'});
  } else {
    tokenClient.requestAccessToken({prompt: ''});
  }
}

async function inicializarHojaCalculo() {
  try {
    // Busca el archivo si eres dueño O si te lo compartieron
    const response = await gapi.client.drive.files.list({
      q: `name = '${SPREADSHEET_NAME}' and mimeType = 'application/vnd.google-apps.spreadsheet' and trashed = false`,
      fields: 'files(id, name, sharedWithMe)',
    });

    const files = response.result.files;
    if (files && files.length > 0) {
      spreadsheetId = files[0].id;
      document.getElementById('status').innerText = 'Conectado a la hoja de cálculo.';
    } else {
      // Si no existe ni se ha compartido, crea una nueva
      const createResponse = await gapi.client.sheets.spreadsheets.create({
        properties: { title: SPREADSHEET_NAME },
        sheets: [{ properties: { title: 'Registros' } }]
      });
      spreadsheetId = createResponse.result.spreadsheetId;
      
      await gapi.client.sheets.spreadsheets.values.append({
        spreadsheetId: spreadsheetId,
        range: 'Registros!A1',
        valueInputOption: 'USER_ENTERED',
        resource: { values: [['Fecha', 'Tipo', 'Concepto', 'Monto']] }
      });
      document.getElementById('status').innerText = 'Nuevo archivo creado en tu Google Drive.';
    }
  } catch (err) {
    document.getElementById('status').innerText = 'Error: ' + err.message;
  }
}

async function guardarRegistro(event) {
  event.preventDefault();
  const btn = document.getElementById('btn-save');
  btn.disabled = true;
  document.getElementById('status').innerText = 'Guardando en tu Google Drive...';

  const fecha = document.getElementById('fecha').value;
  const tipo = document.getElementById('tipo').value;
  const concepto = document.getElementById('concepto').value;
  const monto = parseFloat(document.getElementById('monto').value);

  try {
    await gapi.client.sheets.spreadsheets.values.append({
      spreadsheetId: spreadsheetId,
      range: 'Registros!A:D',
      valueInputOption: 'USER_ENTERED',
      resource: { values: [[fecha, tipo, concepto, monto]] }
    });

    document.getElementById('status').innerText = '¡Guardado con éxito!';
    document.getElementById('concepto').value = '';
    document.getElementById('monto').value = '';
  } catch (err) {
    document.getElementById('status').innerText = 'Error al guardar: ' + err.message;
  } finally {
    btn.disabled = false;
  }
}
