// Come With — desktop app (Electron).
// Opens the live dashboard in a phone-sized window on your computer, so you can
// build and test exactly what the mobile app feels like. It's a real installable
// desktop app too (package it with electron-builder — see README).
const { app, BrowserWindow, shell, Menu } = require('electron');

const URL = process.env.CW_URL || 'https://comewith.org/dashboard.html';
// iPhone 14/15 logical size — a true phone-sized window.
const PHONE = { width: 402, height: 874 };

function createWindow() {
  const win = new BrowserWindow({
    width: PHONE.width,
    height: PHONE.height,
    minWidth: 360,
    minHeight: 600,
    title: 'Come With',
    backgroundColor: '#140a17',
    autoHideMenuBar: true,
    webPreferences: {
      contextIsolation: true,
      nodeIntegration: false,
    },
  });

  // A phone user-agent so the site uses its mobile layout.
  const mobileUA = 'Mozilla/5.0 (iPhone; CPU iPhone OS 17_0 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.0 Mobile/15E148 Safari/604.1';
  win.webContents.setUserAgent(mobileUA);
  win.loadURL(URL, { userAgent: mobileUA });

  // External links (mailto:, new tabs, other sites) open in the real browser,
  // not inside the app window.
  win.webContents.setWindowOpenHandler(({ url }) => {
    if (!url.startsWith('https://comewith.org')) { shell.openExternal(url); return { action: 'deny' }; }
    return { action: 'allow' };
  });
}

app.whenReady().then(() => {
  Menu.setApplicationMenu(null); // clean, app-like chrome
  createWindow();
  app.on('activate', () => { if (BrowserWindow.getAllWindows().length === 0) createWindow(); });
});
app.on('window-all-closed', () => { if (process.platform !== 'darwin') app.quit(); });
