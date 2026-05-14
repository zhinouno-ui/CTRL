const AGENT_USER_SEARCH_URL = 'https://bo.casinodrex.com/agents/user_search';

async function callDrex(method, ...args) {
  if (window.ctrlElectron && typeof window.ctrlElectron.drexAutomation === 'function') {
    return window.ctrlElectron.drexAutomation(method, ...args);
  }

  if (window.drexAutomation && typeof window.drexAutomation[method] === 'function') {
    return window.drexAutomation[method](...args);
  }

  throw new Error('La API de automatización no está disponible. Verificá que el BrowserWindow use electron/app-preload.js o electron/agent-preload.js.');
}

async function prepararCargasAutomaticas() {
  const estado = await callDrex('estadoPagina');
  if (estado.needsLogin) {
    alert(estado.message);
    return estado;
  }

  if (window.ctrlElectron && typeof window.ctrlElectron.openAgentWindow === 'function') {
    await window.ctrlElectron.openAgentWindow();
  }

  if (!estado.url.startsWith(AGENT_USER_SEARCH_URL)) {
    return callDrex('irABusquedaUsuarios');
  }

  return estado;
}

async function buscarYCargarSaldo() {
  await prepararCargasAutomaticas();

  const usuario = document.querySelector('#autoUsuario').value.trim();
  const monto = document.querySelector('#autoMonto').value.trim();

  const busqueda = await callDrex('buscarUsuario', usuario);
  if (busqueda.needsLogin || !busqueda.exists) {
    alert(busqueda.message || 'Usuario no encontrado.');
    return busqueda;
  }

  const carga = await callDrex('cargarSaldo', monto);
  alert(carga.message);
  return carga;
}

async function buscarYRetirarSaldo() {
  await prepararCargasAutomaticas();

  const usuario = document.querySelector('#autoUsuario').value.trim();
  const monto = document.querySelector('#autoMonto').value.trim();

  const busqueda = await callDrex('buscarUsuario', usuario);
  if (busqueda.needsLogin || !busqueda.exists) {
    alert(busqueda.message || 'Usuario no encontrado.');
    return busqueda;
  }

  const retiro = await callDrex('retirarSaldo', monto);
  alert(retiro.message);
  return retiro;
}

async function buscarYCambiarClave() {
  await prepararCargasAutomaticas();

  const usuario = document.querySelector('#autoUsuario').value.trim();
  const clave = document.querySelector('#autoClave').value.trim();

  const busqueda = await callDrex('buscarUsuario', usuario);
  if (busqueda.needsLogin || !busqueda.exists) {
    alert(busqueda.message || 'Usuario no encontrado.');
    return busqueda;
  }

  const cambio = await callDrex('cambiarClave', clave);
  alert(cambio.message);
  return cambio;
}

function montarPanelCargasAutomaticas() {
  const panel = document.createElement('section');
  panel.id = 'cargasAutomaticasPanel';
  panel.innerHTML = `
    <label>Usuario</label>
    <input id="autoUsuario" type="text" minlength="3" placeholder="Ej: aaal67">

    <label>Monto</label>
    <input id="autoMonto" type="number" min="1" step="0.01" placeholder="Ej: 1000">

    <label>Nueva clave</label>
    <input id="autoClave" type="password" placeholder="Solo para cambio de clave">

    <button id="btnAutoCarga" type="button">Cargar saldo</button>
    <button id="btnAutoRetiro" type="button">Retirar saldo</button>
    <button id="btnAutoClave" type="button">Cambiar clave</button>
  `;
  document.body.appendChild(panel);

  document.querySelector('#btnAutoCarga').addEventListener('click', buscarYCargarSaldo);
  document.querySelector('#btnAutoRetiro').addEventListener('click', buscarYRetirarSaldo);
  document.querySelector('#btnAutoClave').addEventListener('click', buscarYCambiarClave);
}

window.addEventListener('DOMContentLoaded', montarPanelCargasAutomaticas);
