// Fichator 3000 — processus principal (fenêtre, fichiers, mises à jour)
const { app, BrowserWindow, ipcMain, dialog, shell } = require('electron');
const path = require('path');
const fs = require('fs');
const { autoUpdater } = require('electron-updater');

const CONFIG = path.join(app.getPath('userData'), 'config.json');
let win = null;

/* ---------- Réglages (dossier des fiches, dernière version vue) ---------- */
function readConfig() {
  try { return JSON.parse(fs.readFileSync(CONFIG, 'utf8')); } catch { return {}; }
}
function writeConfig(patch) {
  const c = Object.assign(readConfig(), patch);
  try { fs.mkdirSync(path.dirname(CONFIG), { recursive: true }); fs.writeFileSync(CONFIG, JSON.stringify(c, null, 2)); } catch {}
  return c;
}

/* ---------- Fenêtre ---------- */
function createWindow() {
  win = new BrowserWindow({
    width: 1400, height: 900, minWidth: 980, minHeight: 640,
    title: 'Fichator 3000',
    backgroundColor: '#f4ede1',
    autoHideMenuBar: true,
    webPreferences: { preload: path.join(__dirname, 'preload.js'), contextIsolation: true, nodeIntegration: false, sandbox: true }
  });
  win.loadFile(path.join(__dirname, 'src', 'index.html'));
  // Les liens externes s'ouvrent dans le navigateur, pas dans l'appli
  win.webContents.setWindowOpenHandler(({ url }) => { shell.openExternal(url); return { action: 'deny' }; });
}

/* ---------- Dossier des fiches ---------- */
async function chooseFolder() {
  const cur = readConfig().folder;
  const r = await dialog.showOpenDialog(win, {
    title: 'Choisis le dossier où ranger les fiches',
    defaultPath: cur || app.getPath('documents'),
    properties: ['openDirectory', 'createDirectory']
  });
  if (r.canceled || !r.filePaths[0]) return null;
  writeConfig({ folder: r.filePaths[0] });
  return r.filePaths[0];
}

function folderOk(p) {
  try { return !!p && fs.statSync(p).isDirectory(); } catch { return false; }
}

async function writeFile(target, bytes) {
  await fs.promises.writeFile(target, Buffer.from(bytes));
  return { ok: true, path: target, name: path.basename(target) };
}

ipcMain.handle('config:get', () => {
  const c = readConfig();
  return { folder: c.folder || null, folderOk: folderOk(c.folder), lastSeenVersion: c.lastSeenVersion || null, version: app.getVersion() };
});
ipcMain.handle('config:seen', (_e, v) => { writeConfig({ lastSeenVersion: v }); return true; });
ipcMain.handle('folder:choose', () => chooseFolder());

// Enregistrement direct dans le dossier des fiches
ipcMain.handle('file:save', async (_e, { name, bytes }) => {
  let folder = readConfig().folder;
  if (!folderOk(folder)) {
    if (folder) {
      await dialog.showMessageBox(win, { type: 'warning', title: 'Dossier introuvable',
        message: "Le dossier des fiches n'est pas accessible.", detail: `${folder}\n\nVérifie que le PC est bien connecté au réseau, ou choisis un autre dossier.` });
    }
    folder = await chooseFolder();
    if (!folder) return { ok: false, canceled: true };
  }
  const target = path.join(folder, name);
  if (fs.existsSync(target)) {
    const r = await dialog.showMessageBox(win, {
      type: 'question', title: 'Fiche déjà existante',
      message: `Une fiche « ${name} » existe déjà dans le dossier.`,
      buttons: ['Remplacer', 'Enregistrer sous…', 'Annuler'], defaultId: 1, cancelId: 2
    });
    if (r.response === 2) return { ok: false, canceled: true };
    if (r.response === 1) return saveAs(name, bytes, folder);
  }
  try { return await writeFile(target, bytes); }
  catch (err) {
    if (err && (err.code === 'EBUSY' || err.code === 'EPERM')) {
      return { ok: false, error: "La fiche est ouverte dans Word. Ferme-la puis réessaie." };
    }
    return { ok: false, error: String(err && err.message || err) };
  }
});

async function saveAs(name, bytes, startFolder) {
  const r = await dialog.showSaveDialog(win, {
    title: 'Enregistrer la fiche',
    defaultPath: path.join(startFolder || readConfig().folder || app.getPath('documents'), name),
    filters: [{ name: 'Document Word', extensions: ['docx'] }]
  });
  if (r.canceled || !r.filePath) return { ok: false, canceled: true };
  try { return await writeFile(r.filePath, bytes); }
  catch (err) { return { ok: false, error: String(err && err.message || err) }; }
}
ipcMain.handle('file:saveAs', (_e, { name, bytes }) => saveAs(name, bytes));

ipcMain.handle('file:open', (_e, p) => shell.openPath(p));
ipcMain.handle('file:reveal', (_e, p) => { shell.showItemInFolder(p); return true; });

/* ---------- Mises à jour automatiques ---------- */
function send(ch, data) { if (win && !win.isDestroyed()) win.webContents.send(ch, data); }
function notesText(info) {
  const n = info && info.releaseNotes;
  if (!n) return '';
  if (Array.isArray(n)) return n.map(x => x.note || '').join('\n');
  return String(n);
}
function setupUpdates() {
  if (!app.isPackaged) return; // pas de mise à jour en mode développement
  autoUpdater.autoDownload = true;
  autoUpdater.autoInstallOnAppQuit = true;
  autoUpdater.on('update-available', info => send('update', { state: 'downloading', version: info.version }));
  autoUpdater.on('download-progress', p => send('update', { state: 'progress', percent: Math.round(p.percent || 0) }));
  autoUpdater.on('update-downloaded', info => send('update', { state: 'ready', version: info.version, notes: notesText(info) }));
  autoUpdater.on('error', err => console.warn('Mise à jour :', err && err.message));
  const check = () => autoUpdater.checkForUpdates().catch(() => {});
  setTimeout(check, 4000);
  setInterval(check, 2 * 60 * 60 * 1000); // toutes les 2 h si l'appli reste ouverte
}
ipcMain.handle('update:install', () => { autoUpdater.quitAndInstall(false, true); });

/* ---------- Démarrage ---------- */
if (!app.requestSingleInstanceLock()) {
  app.quit();
} else {
  app.on('second-instance', () => { if (win) { if (win.isMinimized()) win.restore(); win.focus(); } });
  app.whenReady().then(() => { createWindow(); setupUpdates(); });
  app.on('window-all-closed', () => app.quit());
}
