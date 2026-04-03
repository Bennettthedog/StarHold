const { contextBridge, ipcRenderer } = require("electron")

function exposeEventBridge(channel) {
  return (handler) => {
    if (typeof handler !== "function") return () => {}
    const wrapped = (_event, payload) => {
      try {
        handler(payload)
      } catch {
        // no-op
      }
    }
    ipcRenderer.on(channel, wrapped)
    return () => ipcRenderer.removeListener(channel, wrapped)
  }
}

contextBridge.exposeInMainWorld("api", {
  openSuperluminalShip: () => ipcRenderer.invoke("starhold:openSuperluminalShip"),
  getAppVersion: () => ipcRenderer.invoke("starhold:getAppVersion"),
  openUpdates: () => ipcRenderer.invoke("starhold:openUpdates"),
  checkForUpdates: () => ipcRenderer.invoke("starhold:checkForUpdates"),
  getRollsConfig: () => ipcRenderer.invoke("starhold:getRollsConfig"),
  getSettings: () => ipcRenderer.invoke("starhold:getSettings"),
  chooseDefaultShipFolder: () => ipcRenderer.invoke("starhold:chooseDefaultShipFolder"),
  openAssignDamageWindow: () => ipcRenderer.invoke("starhold:openAssignDamageWindow"),
  openWeaponSelectWindow: () => ipcRenderer.invoke("starhold:openWeaponSelectWindow"),
  setAssignDamageWindowState: (payload) => ipcRenderer.invoke("starhold:setAssignDamageWindowState", payload),
  setAssignDamageWeaponPrompt: (payload) => ipcRenderer.invoke("starhold:setAssignDamageWeaponPrompt", payload),
  setWeaponSelectWindowState: (payload) => ipcRenderer.invoke("starhold:setWeaponSelectWindowState", payload),
  setAssignDamageManualDacPrompt: (payload) => ipcRenderer.invoke("starhold:setAssignDamageManualDacPrompt", payload),
  closeAssignDamageWindow: () => ipcRenderer.invoke("starhold:closeAssignDamageWindow"),
  closeWeaponSelectWindow: () => ipcRenderer.invoke("starhold:closeWeaponSelectWindow"),
  openManualDacRollWindow: () => ipcRenderer.invoke("starhold:openManualDacRollWindow"),
  setManualDacRollWindowState: (payload) => ipcRenderer.invoke("starhold:setManualDacRollWindowState", payload),
  closeManualDacRollWindow: () => ipcRenderer.invoke("starhold:closeManualDacRollWindow"),
  previewAssignDamageWindowSelection: (payload) => ipcRenderer.send("starhold:assignDamageWindowPreview", payload),
  submitAssignDamageWindow: (payload) => ipcRenderer.send("starhold:assignDamageWindowSubmit", payload),
  submitAssignDamageWeaponSelection: (payload) => ipcRenderer.send("starhold:assignDamageWeaponSelectionSubmit", payload),
  cancelAssignDamageWeaponSelection: (payload) => ipcRenderer.send("starhold:assignDamageWeaponSelectionCancel", payload),
  submitAssignDamageManualDacPrompt: (payload) => ipcRenderer.send("starhold:assignDamageManualDacPromptSubmit", payload),
  cancelAssignDamageManualDacPrompt: (payload) => ipcRenderer.send("starhold:assignDamageManualDacPromptCancel", payload),
  submitManualDacRollWindow: (payload) => ipcRenderer.send("starhold:manualDacRollWindowSubmit", payload),
  cancelManualDacRollWindow: (payload) => ipcRenderer.send("starhold:manualDacRollWindowCancel", payload),
  onAssignDamageWindowState: exposeEventBridge("starhold:assignDamageWindowState"),
  onAssignDamageWeaponPrompt: exposeEventBridge("starhold:assignDamageWeaponPrompt"),
  onWeaponSelectWindowState: exposeEventBridge("starhold:weaponSelectWindowState"),
  onAssignDamageManualDacPrompt: exposeEventBridge("starhold:assignDamageManualDacPrompt"),
  onAssignDamageWindowPreview: exposeEventBridge("starhold:assignDamageWindowPreview"),
  onAssignDamageWindowClosed: exposeEventBridge("starhold:assignDamageWindowClosed"),
  onWeaponSelectWindowClosed: exposeEventBridge("starhold:weaponSelectWindowClosed"),
  onAssignDamageWindowSubmit: exposeEventBridge("starhold:assignDamageWindowSubmit"),
  onAssignDamageWeaponSelectionSubmit: exposeEventBridge("starhold:assignDamageWeaponSelectionSubmit"),
  onAssignDamageWeaponSelectionCancel: exposeEventBridge("starhold:assignDamageWeaponSelectionCancel"),
  onAssignDamageManualDacPromptSubmit: exposeEventBridge("starhold:assignDamageManualDacPromptSubmit"),
  onAssignDamageManualDacPromptCancel: exposeEventBridge("starhold:assignDamageManualDacPromptCancel"),
  onManualDacRollWindowState: exposeEventBridge("starhold:manualDacRollWindowState"),
  onManualDacRollWindowSubmit: exposeEventBridge("starhold:manualDacRollWindowSubmit"),
  onManualDacRollWindowCancel: exposeEventBridge("starhold:manualDacRollWindowCancel"),
  onManualDacRollWindowClosed: exposeEventBridge("starhold:manualDacRollWindowClosed")
})
