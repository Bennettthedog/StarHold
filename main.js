const { app, BrowserWindow, dialog, ipcMain, Menu, shell } = require("electron")
const path = require("path")
const fs = require("fs/promises")
let autoUpdater = null
try {
  ({ autoUpdater } = require("electron-updater"))
} catch {
  autoUpdater = null
}

let mainWindow = null
let assignDamageWindow = null
let assignDamageWindowState = null
let weaponSelectWindow = null
let weaponSelectWindowState = null
let manualDacRollWindow = null
let manualDacRollWindowState = null
const STARHOLD_SETTINGS_FILE = "starhold-settings.json"
const STARHOLD_RELEASES_URL = "https://github.com/Bennettthedog/StarHold/releases"

function sendToMainWindow(channel, payload = {}) {
  if (!mainWindow || mainWindow.isDestroyed()) return
  mainWindow.webContents.send(channel, payload)
}

function createWindow() {
  const win = new BrowserWindow({
    width: 1480,
    height: 940,
    autoHideMenuBar: true,
    webPreferences: {
      preload: path.join(__dirname, "preload.js"),
      contextIsolation: true,
      nodeIntegration: false
    }
  })

  mainWindow = win
  win.on("closed", () => {
    if (mainWindow === win) mainWindow = null
  })

  win.loadFile(path.join(__dirname, "index.html"))
  win.setMenuBarVisibility(false)
}

function createAssignDamageWindow() {
  if (assignDamageWindow && !assignDamageWindow.isDestroyed()) {
    assignDamageWindow.focus()
    return assignDamageWindow
  }

  const win = new BrowserWindow({
    width: 560,
    height: 720,
    minWidth: 480,
    minHeight: 560,
    autoHideMenuBar: true,
    parent: mainWindow && !mainWindow.isDestroyed() ? mainWindow : null,
    webPreferences: {
      preload: path.join(__dirname, "preload.js"),
      contextIsolation: true,
      nodeIntegration: false
    }
  })

  assignDamageWindow = win
  win.on("closed", () => {
    if (assignDamageWindow === win) assignDamageWindow = null
    sendToMainWindow("starhold:assignDamageWindowClosed", {})
  })

  win.webContents.on("did-finish-load", () => {
    if (assignDamageWindow !== win) return
    if (!assignDamageWindowState) return
    win.webContents.send("starhold:assignDamageWindowState", assignDamageWindowState)
  })

  win.loadFile(path.join(__dirname, "assign-damage.html"))
  win.setMenuBarVisibility(false)
  return win
}

function createWeaponSelectWindow() {
  if (weaponSelectWindow && !weaponSelectWindow.isDestroyed()) {
    weaponSelectWindow.focus()
    return weaponSelectWindow
  }

  const win = new BrowserWindow({
    width: 440,
    height: 360,
    minWidth: 400,
    minHeight: 280,
    autoHideMenuBar: true,
    parent: mainWindow && !mainWindow.isDestroyed() ? mainWindow : null,
    webPreferences: {
      preload: path.join(__dirname, "preload.js"),
      contextIsolation: true,
      nodeIntegration: false
    }
  })

  weaponSelectWindow = win
  win.on("closed", () => {
    if (weaponSelectWindow === win) weaponSelectWindow = null
    sendToMainWindow("starhold:weaponSelectWindowClosed", {})
  })

  win.webContents.on("did-finish-load", () => {
    if (weaponSelectWindow !== win) return
    if (!weaponSelectWindowState) return
    win.webContents.send("starhold:weaponSelectWindowState", weaponSelectWindowState)
  })

  win.loadFile(path.join(__dirname, "weapon-select.html"))
  win.setMenuBarVisibility(false)
  return win
}

function createManualDacRollWindow() {
  if (manualDacRollWindow && !manualDacRollWindow.isDestroyed()) {
    manualDacRollWindow.focus()
    return manualDacRollWindow
  }

  const win = new BrowserWindow({
    width: 440,
    height: 310,
    minWidth: 400,
    minHeight: 260,
    autoHideMenuBar: true,
    parent: mainWindow && !mainWindow.isDestroyed() ? mainWindow : null,
    webPreferences: {
      preload: path.join(__dirname, "preload.js"),
      contextIsolation: true,
      nodeIntegration: false
    }
  })

  manualDacRollWindow = win
  win.on("closed", () => {
    if (manualDacRollWindow === win) manualDacRollWindow = null
    sendToMainWindow("starhold:manualDacRollWindowClosed", {})
  })

  win.webContents.on("did-finish-load", () => {
    if (manualDacRollWindow !== win) return
    if (!manualDacRollWindowState) return
    win.webContents.send("starhold:manualDacRollWindowState", manualDacRollWindowState)
  })

  win.loadFile(path.join(__dirname, "manual-dac-roll.html"))
  win.setMenuBarVisibility(false)
  return win
}

function findFirstByExt(files, exts) {
  const lowered = files.map((name) => ({ name, lower: String(name || "").toLowerCase() }))
  for (const ext of exts) {
    const hit = lowered.find((entry) => entry.lower.endsWith(ext))
    if (hit) return hit.name
  }
  return null
}

function getMimeForExt(filePath) {
  const ext = String(path.extname(filePath) || "").toLowerCase()
  if (ext === ".png") return "image/png"
  if (ext === ".jpg" || ext === ".jpeg") return "image/jpeg"
  if (ext === ".webp") return "image/webp"
  if (ext === ".bmp") return "image/bmp"
  if (ext === ".gif") return "image/gif"
  if (ext === ".tif" || ext === ".tiff") return "image/tiff"
  return "application/octet-stream"
}

async function fileToDataUrl(filePath) {
  const buf = await fs.readFile(filePath)
  const mime = getMimeForExt(filePath)
  return `data:${mime};base64,${buf.toString("base64")}`
}

function getWorkspaceDefaultShipFolder() {
  return path.resolve(app.getAppPath(), "..", "..", "Superluminal Ships")
}

function getSettingsFilePath() {
  return path.join(app.getPath("userData"), STARHOLD_SETTINGS_FILE)
}

function normalizeFolderPath(folderPath) {
  const raw = String(folderPath || "").trim()
  if (!raw) return ""
  try {
    return path.resolve(raw)
  } catch {
    return ""
  }
}

function sanitizeSettings(value) {
  const input = value && typeof value === "object" ? value : {}
  return {
    defaultShipFolder: normalizeFolderPath(input.defaultShipFolder)
  }
}

async function pathIsDirectory(folderPath) {
  const normalized = normalizeFolderPath(folderPath)
  if (!normalized) return false
  try {
    const stat = await fs.stat(normalized)
    return stat.isDirectory()
  } catch {
    return false
  }
}

async function readStarholdSettings() {
  try {
    const text = await fs.readFile(getSettingsFilePath(), "utf-8")
    return sanitizeSettings(JSON.parse(text))
  } catch (err) {
    if (err && (err.code === "ENOENT" || err.name === "SyntaxError")) {
      return sanitizeSettings(null)
    }
    throw err
  }
}

async function writeStarholdSettings(nextValues) {
  const current = await readStarholdSettings()
  const merged = sanitizeSettings({ ...current, ...(nextValues || {}) })
  const settingsPath = getSettingsFilePath()
  await fs.mkdir(path.dirname(settingsPath), { recursive: true })
  await fs.writeFile(settingsPath, JSON.stringify(merged, null, 2), "utf-8")
  return merged
}

async function getShipDialogDefaultPath() {
  const settings = await readStarholdSettings()
  if (await pathIsDirectory(settings.defaultShipFolder)) return settings.defaultShipFolder

  const workspaceDefaultPath = getWorkspaceDefaultShipFolder()
  if (await pathIsDirectory(workspaceDefaultPath)) return workspaceDefaultPath

  return app.getPath("documents")
}

async function pickSuperluminalShipFile() {
  const dialogDefaultPath = await getShipDialogDefaultPath()
  const picked = await dialog.showOpenDialog({
    title: "Select SuperLuminal ship folder",
    properties: ["openDirectory"],
    defaultPath: dialogDefaultPath
  })

  if (picked.canceled || !picked.filePaths || picked.filePaths.length === 0) {
    return { ok: false, error: "No folder selected." }
  }

  const folderPath = picked.filePaths[0]
  const files = await fs.readdir(folderPath)
  const jsonCandidates = files
    .filter((name) => String(path.extname(name) || "").toLowerCase() === ".json")
    .sort((a, b) => a.localeCompare(b))
  if (jsonCandidates.length === 0) {
    return { ok: false, error: "No JSON file found in the selected folder." }
  }

  let jsonFile = null
  let jsonPath = ""
  let jsonText = ""
  let parsed = null
  for (const candidate of jsonCandidates) {
    const candidatePath = path.join(folderPath, candidate)
    let candidateText = ""
    let candidateParsed = null
    try {
      candidateText = await fs.readFile(candidatePath, "utf-8")
      candidateParsed = JSON.parse(candidateText)
    } catch {
      continue
    }
    if (!candidateParsed || typeof candidateParsed !== "object" || !candidateParsed.ssd || typeof candidateParsed.ssd !== "object") {
      continue
    }

    jsonFile = candidate
    jsonPath = candidatePath
    jsonText = candidateText
    parsed = candidateParsed
    break
  }

  if (!jsonFile || !jsonPath || !jsonText || !parsed) {
    return { ok: false, error: "No valid SuperLuminal ship JSON (with ssd data) found in the selected folder." }
  }

  const baseName = path.basename(jsonFile, path.extname(jsonFile))

  const imageExts = [".png", ".jpg", ".jpeg", ".webp", ".bmp", ".tif", ".tiff"]
  const sameStemImage = files.find((name) => {
    const stem = path.basename(name, path.extname(name))
    return stem === baseName && imageExts.includes(path.extname(name).toLowerCase())
  })
  const imageFile = sameStemImage || findFirstByExt(files, imageExts)
  if (!imageFile) {
    return { ok: false, error: "No image file found in the selected folder." }
  }

  const imagePath = path.join(folderPath, imageFile)
  const imageDataUrl = await fileToDataUrl(imagePath)

  return {
    ok: true,
    jsonPath,
    folderPath,
    jsonFile,
    imagePath,
    imageFile,
    baseName,
    jsonText,
    imageDataUrl
  }
}

app.whenReady().then(() => {
  Menu.setApplicationMenu(null)
  createWindow()

  app.on("activate", () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow()
  })
})

app.on("window-all-closed", () => {
  if (process.platform !== "darwin") app.quit()
})

ipcMain.handle("starhold:openSuperluminalShip", async () => {
  try {
    return await pickSuperluminalShipFile()
  } catch (err) {
    return { ok: false, error: err?.message || "Failed to open SuperLuminal ship file." }
  }
})

ipcMain.handle("starhold:getRollsConfig", async () => {
  try {
    const filePath = path.join(__dirname, "Data", "Rolls.json")
    const text = await fs.readFile(filePath, "utf-8")
    const rollsConfig = JSON.parse(text)
    return { ok: true, rollsConfig }
  } catch (err) {
    return { ok: false, error: err?.message || "Failed to load Rolls.json." }
  }
})

ipcMain.handle("starhold:getAppVersion", async () => {
  try {
    return { ok: true, version: app.getVersion() }
  } catch (err) {
    return { ok: false, error: err?.message || "Failed to load app version." }
  }
})

ipcMain.handle("starhold:openUpdates", async () => {
  try {
    await shell.openExternal(STARHOLD_RELEASES_URL)
    return { ok: true }
  } catch (err) {
    return { ok: false, error: err?.message || "Failed to open StarHold releases." }
  }
})

ipcMain.handle("starhold:checkForUpdates", async () => {
  if (!autoUpdater || !app.isPackaged) {
    try {
      await shell.openExternal(STARHOLD_RELEASES_URL)
      return { ok: true, openedExternal: true }
    } catch (err) {
      return { ok: false, error: err?.message || "Failed to open StarHold releases." }
    }
  }

  try {
    autoUpdater.checkForUpdates()
    return { ok: true, openedExternal: false }
  } catch (err) {
    return { ok: false, error: err?.message || "Failed to check for updates." }
  }
})

ipcMain.handle("starhold:getSettings", async () => {
  try {
    const settings = await readStarholdSettings()
    return { ok: true, ...settings }
  } catch (err) {
    return { ok: false, error: err?.message || "Failed to load settings." }
  }
})

ipcMain.handle("starhold:chooseDefaultShipFolder", async () => {
  try {
    const current = await readStarholdSettings()
    const dialogDefaultPath = await getShipDialogDefaultPath()
    const picked = await dialog.showOpenDialog({
      title: "Choose default ship folder",
      properties: ["openDirectory", "createDirectory"],
      defaultPath: dialogDefaultPath
    })

    if (picked.canceled || !picked.filePaths || picked.filePaths.length === 0) {
      return { ok: true, canceled: true, defaultShipFolder: current.defaultShipFolder }
    }

    const saved = await writeStarholdSettings({
      defaultShipFolder: picked.filePaths[0]
    })
    return { ok: true, canceled: false, defaultShipFolder: saved.defaultShipFolder }
  } catch (err) {
    return { ok: false, error: err?.message || "Failed to choose default ship folder." }
  }
})

ipcMain.handle("starhold:openAssignDamageWindow", async () => {
  try {
    createAssignDamageWindow()
    return { ok: true }
  } catch (err) {
    return { ok: false, error: err?.message || "Failed to open Assign Damage window." }
  }
})

ipcMain.handle("starhold:openWeaponSelectWindow", async () => {
  try {
    createWeaponSelectWindow()
    return { ok: true }
  } catch (err) {
    return { ok: false, error: err?.message || "Failed to open Weapon Select window." }
  }
})

ipcMain.handle("starhold:openManualDacRollWindow", async () => {
  try {
    createManualDacRollWindow()
    return { ok: true }
  } catch (err) {
    return { ok: false, error: err?.message || "Failed to open Manual DAC Roll window." }
  }
})

ipcMain.handle("starhold:setAssignDamageWindowState", async (_event, payload) => {
  try {
    assignDamageWindowState = payload && typeof payload === "object" ? payload : null
    if (assignDamageWindow && !assignDamageWindow.isDestroyed() && assignDamageWindowState) {
      assignDamageWindow.webContents.send("starhold:assignDamageWindowState", assignDamageWindowState)
    }
    return { ok: true }
  } catch (err) {
    return { ok: false, error: err?.message || "Failed to update Assign Damage window state." }
  }
})

ipcMain.handle("starhold:setAssignDamageWeaponPrompt", async (_event, payload) => {
  try {
    let shown = false
    if (assignDamageWindow && !assignDamageWindow.isDestroyed()) {
      assignDamageWindow.webContents.send("starhold:assignDamageWeaponPrompt", payload && typeof payload === "object" ? payload : null)
      shown = true
      try {
        assignDamageWindow.focus()
      } catch {
        // no-op
      }
    }
    return { ok: true, shown }
  } catch (err) {
    return { ok: false, error: err?.message || "Failed to update Assign Damage weapon prompt." }
  }
})

ipcMain.handle("starhold:setWeaponSelectWindowState", async (_event, payload) => {
  try {
    weaponSelectWindowState = payload && typeof payload === "object" ? payload : null
    let shown = false
    if (weaponSelectWindow && !weaponSelectWindow.isDestroyed() && weaponSelectWindowState) {
      weaponSelectWindow.webContents.send("starhold:weaponSelectWindowState", weaponSelectWindowState)
      shown = true
      try {
        weaponSelectWindow.focus()
      } catch {
        // no-op
      }
    }
    return { ok: true, shown }
  } catch (err) {
    return { ok: false, error: err?.message || "Failed to update Weapon Select window state." }
  }
})

ipcMain.handle("starhold:setAssignDamageManualDacPrompt", async (_event, payload) => {
  try {
    let shown = false
    if (assignDamageWindow && !assignDamageWindow.isDestroyed()) {
      assignDamageWindow.webContents.send("starhold:assignDamageManualDacPrompt", payload && typeof payload === "object" ? payload : null)
      shown = true
      try {
        assignDamageWindow.focus()
      } catch {
        // no-op
      }
    }
    return { ok: true, shown }
  } catch (err) {
    return { ok: false, error: err?.message || "Failed to update Assign Damage manual DAC prompt." }
  }
})

ipcMain.handle("starhold:closeAssignDamageWindow", async () => {
  try {
    if (assignDamageWindow && !assignDamageWindow.isDestroyed()) {
      const win = assignDamageWindow
      await new Promise((resolve) => {
        if (!win || win.isDestroyed()) {
          resolve()
          return
        }

        const done = () => resolve()
        win.once("closed", done)

        try {
          win.close()
        } catch {
          win.removeListener("closed", done)
          resolve()
        }
      })
    }
    return { ok: true }
  } catch (err) {
    return { ok: false, error: err?.message || "Failed to close Assign Damage window." }
  }
})

ipcMain.handle("starhold:closeWeaponSelectWindow", async () => {
  try {
    if (weaponSelectWindow && !weaponSelectWindow.isDestroyed()) {
      const win = weaponSelectWindow
      await new Promise((resolve) => {
        if (!win || win.isDestroyed()) {
          resolve()
          return
        }

        const done = () => resolve()
        win.once("closed", done)

        try {
          win.close()
        } catch {
          win.removeListener("closed", done)
          resolve()
        }
      })
    }
    return { ok: true }
  } catch (err) {
    return { ok: false, error: err?.message || "Failed to close Weapon Select window." }
  }
})

ipcMain.handle("starhold:setManualDacRollWindowState", async (_event, payload) => {
  try {
    manualDacRollWindowState = payload && typeof payload === "object" ? payload : null
    if (manualDacRollWindow && !manualDacRollWindow.isDestroyed() && manualDacRollWindowState) {
      manualDacRollWindow.webContents.send("starhold:manualDacRollWindowState", manualDacRollWindowState)
    }
    return { ok: true }
  } catch (err) {
    return { ok: false, error: err?.message || "Failed to update Manual DAC Roll window state." }
  }
})

ipcMain.handle("starhold:closeManualDacRollWindow", async () => {
  try {
    if (manualDacRollWindow && !manualDacRollWindow.isDestroyed()) {
      const win = manualDacRollWindow
      await new Promise((resolve) => {
        if (!win || win.isDestroyed()) {
          resolve()
          return
        }

        const done = () => resolve()
        win.once("closed", done)

        try {
          win.close()
        } catch {
          win.removeListener("closed", done)
          resolve()
        }
      })
    }
    return { ok: true }
  } catch (err) {
    return { ok: false, error: err?.message || "Failed to close Manual DAC Roll window." }
  }
})

ipcMain.on("starhold:assignDamageWindowSubmit", (_event, payload) => {
  sendToMainWindow("starhold:assignDamageWindowSubmit", payload || {})
})

ipcMain.on("starhold:assignDamageWindowPreview", (_event, payload) => {
  sendToMainWindow("starhold:assignDamageWindowPreview", payload || {})
})

ipcMain.on("starhold:assignDamageWeaponSelectionSubmit", (_event, payload) => {
  sendToMainWindow("starhold:assignDamageWeaponSelectionSubmit", payload || {})
})

ipcMain.on("starhold:assignDamageWeaponSelectionCancel", (_event, payload) => {
  sendToMainWindow("starhold:assignDamageWeaponSelectionCancel", payload || {})
})

ipcMain.on("starhold:assignDamageManualDacPromptSubmit", (_event, payload) => {
  sendToMainWindow("starhold:assignDamageManualDacPromptSubmit", payload || {})
})

ipcMain.on("starhold:assignDamageManualDacPromptCancel", (_event, payload) => {
  sendToMainWindow("starhold:assignDamageManualDacPromptCancel", payload || {})
})

ipcMain.on("starhold:manualDacRollWindowSubmit", (_event, payload) => {
  sendToMainWindow("starhold:manualDacRollWindowSubmit", payload || {})
})

ipcMain.on("starhold:manualDacRollWindowCancel", (_event, payload) => {
  sendToMainWindow("starhold:manualDacRollWindowCancel", payload || {})
})

function formatUpdateName(info) {
  const releaseName = String(info?.releaseName || "").trim()
  const version = String(info?.version || "").trim()
  if (releaseName && version) {
    const lowerName = releaseName.toLowerCase()
    const lowerVersion = version.toLowerCase()
    if (lowerName.includes(lowerVersion)) return releaseName
    return `${releaseName} (v${version})`
  }
  if (releaseName) return releaseName
  if (version) return `v${version}`
  return "Unknown version"
}

if (autoUpdater) {
  autoUpdater.on("update-available", (info) => {
    const updateName = formatUpdateName(info)
    dialog.showMessageBox({
      type: "info",
      title: "Update Found",
      message: `A new update was found: ${updateName}`
    })
  })

  autoUpdater.on("update-not-available", () => {
    dialog.showMessageBox({
      type: "info",
      title: "No Updates",
      message: "You are already on the latest version."
    })
  })

  autoUpdater.on("error", (err) => {
    dialog.showMessageBox({
      type: "error",
      title: "Update Error",
      message: err?.message || "An error occurred while checking for updates."
    })
  })

  autoUpdater.on("update-downloaded", async (info) => {
    const updateName = formatUpdateName(info)
    const res = await dialog.showMessageBox({
      type: "question",
      title: "Update Ready",
      message: `An update has been downloaded (${updateName}). Restart now to install?`,
      buttons: ["Restart", "Later"],
      defaultId: 0,
      cancelId: 1
    })
    if (res.response === 0) {
      autoUpdater.quitAndInstall()
    }
  })
}
