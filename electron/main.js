const { app, BrowserWindow, shell, ipcMain } = require("electron");
const { autoUpdater } = require("electron-updater");
const path = require("node:path");
const fs = require("node:fs");

let win;
function createWindow() {
  win = new BrowserWindow({
    width: 1280,
    height: 900,
    show: false,
    webPreferences: {
      preload: path.join(__dirname, "preload.js"),
      nodeIntegration: false,
      contextIsolation: true,
      sandbox: true
    }
  });
  win.on("ready-to-show", () => win.show());
  win.loadFile(path.join(__dirname, "../renderer/index.html"));
  win.webContents.setWindowOpenHandler(({ url }) => {
    shell.openExternal(url);
    return { action: "deny" };
  });
}

app.whenReady().then(() => {
  createWindow();
  autoUpdater.autoDownload = true;
  autoUpdater.checkForUpdatesAndNotify();
});

app.on("window-all-closed", () => {
  if (process.platform !== "darwin") app.quit();
});
app.on("activate", () => {
  if (BrowserWindow.getAllWindows().length === 0) createWindow();
});

// Save locations
function quotesDir() {
  const dir = path.join(app.getPath("documents"), "Quote Tool", "Quotes");
  fs.mkdirSync(dir, { recursive: true });
  return dir;
}
function dataDir() {
  const dir = app.getPath("userData");
  fs.mkdirSync(dir, { recursive: true });
  return dir;
}

// IPC bridge
ipcMain.handle("paths:get", () => ({ quotes: quotesDir(), data: dataDir() }));
ipcMain.handle("file:save", async (_e, { rel, bytes, folder }) => {
  const base = folder === "quotes" ? quotesDir() : dataDir();
  const out = path.join(base, rel);
  fs.mkdirSync(path.dirname(out), { recursive: true });
  fs.writeFileSync(out, Buffer.from(bytes));
  return out;
});
ipcMain.handle("catalog:read", async () => {
  try {
    const p = path.join(__dirname, "../renderer/vendor_catalog.json");
    if (fs.existsSync(p)) return JSON.parse(fs.readFileSync(p, "utf-8"));
  } catch {}
  return [];
});
