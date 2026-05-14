const path = require('node:path');
const { app, BrowserWindow, ipcMain } = require('electron');

const AGENT_USER_SEARCH_URL = 'https://bo.casinodrex.com/agents/user_search';
let agentWindow;
const pendingAutomation = new Map();

function createAgentWindow() {
  agentWindow = new BrowserWindow({
    width: 1400,
    height: 900,
    title: 'Agentes - Cargas automatizadas',
    webPreferences: {
      preload: path.join(__dirname, 'agent-preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true
    }
  });

  agentWindow.loadURL(AGENT_USER_SEARCH_URL);
  agentWindow.on('closed', () => {
    agentWindow = null;
    pendingAutomation.clear();
  });

  return agentWindow;
}

function getAgentWindow() {
  if (agentWindow && !agentWindow.isDestroyed()) return agentWindow;
  return createAgentWindow();
}

function whenAgentReady(win) {
  if (!win.webContents.isLoading()) return Promise.resolve();
  return new Promise(resolve => win.webContents.once('did-finish-load', resolve));
}

async function sendAutomation(method, ...args) {
  const win = getAgentWindow();
  await whenAgentReady(win);
  const requestId = `${Date.now()}-${Math.random().toString(16).slice(2)}`;

  return new Promise((resolve, reject) => {
    pendingAutomation.set(requestId, { resolve, reject });
    win.webContents.send('drex:automation:run', { requestId, method, args });
  });
}

ipcMain.handle('drex:open-agent-window', () => {
  getAgentWindow().focus();
  return { ok: true };
});

ipcMain.handle('drex:automation', async (_event, { method, args = [] } = {}) => {
  return sendAutomation(method, ...args);
});

ipcMain.on('drex:automation:result', (_event, response = {}) => {
  const pending = pendingAutomation.get(response.requestId);
  if (!pending) return;

  pendingAutomation.delete(response.requestId);
  if (response.ok) pending.resolve(response.result);
  else pending.reject(new Error(response.error || 'Error ejecutando automatización.'));
});

app.whenReady().then(createAgentWindow);
app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});
