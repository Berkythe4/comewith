# Come With — desktop app (phone-sized)

Opens the live dashboard (`comewith.org/dashboard.html`) in a real desktop
window sized like an iPhone, using a mobile user-agent so you see the mobile
layout. Good for building/testing the mobile experience on your computer.

## Run it (test on your computer)

You need [Node.js](https://nodejs.org) installed (LTS is fine). Then, in a
terminal:

```bash
cd desktop
npm install      # one-time — downloads Electron (~200 MB)
npm start        # opens the phone-sized Come With window
```

Point it at a different URL if needed:

```bash
CW_URL=http://localhost:8080/dashboard.html npm start     # macOS/Linux
set CW_URL=https://comewith.org/dashboard.html && npm start   # Windows cmd
```

## Make a double-click installer (optional)

```bash
npm run dist
```

`electron-builder` produces an installer in `desktop/dist/`:
- **Windows:** `Come With Setup x.y.z.exe`
- **macOS:** `Come With-x.y.z.dmg`
- **Linux:** `Come With-x.y.z.AppImage`

Double-click that to install the app locally like any other program.

## Putting it on your phone

- **iPhone:** a desktop Electron build does NOT install on iPhone. Sideloading
  an iPhone app needs a **Mac + Xcode + an Apple Developer account** — there's no
  path from Windows. For iPhone, the **PWA** (Safari → Share → Add to Home
  Screen) is the practical route and gives the same full-screen app.
- **Android:** installable over USB, but via **Capacitor + Android Studio**, not
  this Electron project. If you want that, we scaffold a `capacitor/` wrapper and
  build an `.apk` you can install with the phone plugged in (`adb install`).

This project is the desktop test harness; the PWA is the phone app.
