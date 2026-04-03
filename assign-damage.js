const state = {
  context: null,
  submitInProgress: false,
  manualDacSubmitInProgress: false,
  modePromptPromise: null,
  modePromptResolver: null,
  weaponPromptOpen: false,
  manualDacPromptOpen: false,
  manualDacContext: null
}

function byId(id) {
  return document.getElementById(id)
}

function normalizeInteger(value, fallback = 0) {
  const num = Number(value)
  if (!Number.isFinite(num)) return fallback
  return Math.trunc(num)
}

function normalizeDamageMode(value) {
  const raw = String(value || "").trim().toLowerCase()
  return raw === "automatic" ? "automatic" : "manual"
}

function setHint(text, isError = false) {
  const el = byId("assignDamageHint")
  if (!el) return
  el.textContent = String(text || "")
  el.classList.toggle("error", isError === true)
}

function setControlsDisabled(disabled) {
  const ids = [
    "assignTotalDamage",
    "assignShieldNumber",
    "assignDamageMode",
    "btnAssignDamageSubmit",
    "assignNonBearingPhasersToggle"
  ]
  for (const id of ids) {
    const el = byId(id)
    if (el) el.disabled = !!disabled
  }

  const list = byId("assignNonBearingPhasersList")
  if (list) {
    const inputs = list.querySelectorAll('input[type="checkbox"]')
    for (const input of inputs) {
      input.disabled = !!disabled
    }
  }

  if (disabled) {
    setPhaserDropdownOpen(false)
  } else {
    syncPhaserDropdownToggleState()
  }
}

function isPhaserDropdownOpen() {
  const root = byId("assignNonBearingPhasersDropdown")
  return !!root && root.classList.contains("open")
}

function setPhaserDropdownOpen(open) {
  const root = byId("assignNonBearingPhasersDropdown")
  const toggle = byId("assignNonBearingPhasersToggle")
  const panel = byId("assignNonBearingPhasersPanel")
  if (!root || !toggle || !panel) return

  const next = !!open && !toggle.disabled
  root.classList.toggle("open", next)
  panel.classList.toggle("hidden", !next)
  toggle.setAttribute("aria-expanded", next ? "true" : "false")
}

function getPhaserCheckboxCounts() {
  const container = byId("assignNonBearingPhasersList")
  if (!container) return { total: 0, selected: 0 }
  return {
    total: container.querySelectorAll('input[type="checkbox"]').length,
    selected: container.querySelectorAll('input[type="checkbox"]:checked').length
  }
}

function syncPhaserDropdownToggleState() {
  const toggle = byId("assignNonBearingPhasersToggle")
  if (!toggle) return

  const ctx = state.context || {}
  const { total } = getPhaserCheckboxCounts()
  const disabled = ctx.available !== true || total === 0
  toggle.disabled = disabled
  if (disabled) setPhaserDropdownOpen(false)
}

function updatePhaserDropdownSummary() {
  const summary = byId("assignNonBearingPhasersSummary")
  if (!summary) return

  const { total, selected } = getPhaserCheckboxCounts()
  if (total <= 0) {
    summary.textContent = "No phasers available"
    return
  }
  if (selected <= 0) {
    summary.textContent = `Select phasers (${total} total)`
    return
  }
  if (selected === 1) {
    summary.textContent = `1 phaser selected (${total} total)`
    return
  }
  summary.textContent = `${selected} phasers selected (${total} total)`
}

function isModePromptOpen() {
  const overlay = byId("assignModePromptOverlay")
  return !!overlay && !overlay.classList.contains("hidden")
}

function isWeaponPromptOpen() {
  const overlay = byId("assignWeaponPromptOverlay")
  return !!overlay && !overlay.classList.contains("hidden")
}

function isManualDacPromptOpen() {
  const overlay = byId("assignManualDacOverlay")
  return !!overlay && !overlay.classList.contains("hidden")
}

function setModePromptOpen(open) {
  const overlay = byId("assignModePromptOverlay")
  if (!overlay) return
  overlay.classList.toggle("hidden", !open)
  if (open) {
    const preferred = byId("btnAssignModeManual") || byId("btnAssignModeAutomatic") || byId("btnAssignModePromptCancel")
    preferred?.focus()
  }
}

function resolveModePrompt(choice) {
  const resolver = state.modePromptResolver
  state.modePromptResolver = null
  state.modePromptPromise = null
  setModePromptOpen(false)
  if (typeof resolver === "function") {
    resolver(choice || null)
  }
}

function promptForDamageMode() {
  if (state.modePromptPromise) return state.modePromptPromise

  const overlay = byId("assignModePromptOverlay")
  if (!overlay) {
    const useAutomatic = window.confirm("Apply damage in automatic mode?\n\nOK = Automatic\nCancel = Manual")
    return Promise.resolve(useAutomatic ? "automatic" : "manual")
  }

  state.modePromptPromise = new Promise((resolve) => {
    state.modePromptResolver = resolve
    setModePromptOpen(true)
  })
  return state.modePromptPromise
}

function setWeaponPromptOpen(open) {
  const overlay = byId("assignWeaponPromptOverlay")
  if (!overlay) return

  const next = !!open
  state.weaponPromptOpen = next
  overlay.classList.toggle("hidden", !next)
  overlay.setAttribute("aria-hidden", next ? "false" : "true")

  if (!next) return

  const preferred = byId("assignWeaponPromptList")?.querySelector("button")
    || byId("btnAssignWeaponPromptSkip")
    || byId("btnAssignWeaponPromptCancel")
  preferred?.focus?.()
}

function setManualDacError(text) {
  const el = byId("assignManualDacError")
  if (!el) return
  const msg = String(text || "").trim()
  el.textContent = msg || "Enter a whole number between 2 and 12."
  el.classList.toggle("hidden", msg.length === 0)
}

function setManualDacElementHidden(id, hidden) {
  const el = byId(id)
  if (!el) return
  el.classList.toggle("manualDacHidden", hidden === true)
}

function setManualDacCompleteMode(enabled, completionText = "Damage complete") {
  const complete = enabled === true
  const closeButton = byId("btnAssignManualDacClose")
  if (closeButton) {
    closeButton.setAttribute("aria-label", complete ? "Close assign damage window" : "Cancel manual DAC allocation")
  }

  setManualDacElementHidden("assignManualDacPromptText", complete)
  setManualDacElementHidden("assignManualDacInputField", complete)
  setManualDacElementHidden("assignManualDacError", complete)
  setManualDacElementHidden("assignManualDacActions", complete)
  setManualDacElementHidden("assignManualDacComplete", !complete)

  const completeEl = byId("assignManualDacComplete")
  if (completeEl) {
    const text = String(completionText || "").trim()
    completeEl.textContent = text || "Damage complete"
  }
}

function setManualDacPromptOpen(open) {
  const overlay = byId("assignManualDacOverlay")
  if (!overlay) return

  const next = !!open
  state.manualDacPromptOpen = next
  overlay.classList.toggle("hidden", !next)
  overlay.setAttribute("aria-hidden", next ? "false" : "true")

  if (!next) return

  if (state.manualDacContext?.complete === true) {
    const preferred = byId("btnAssignManualDacClose")
    preferred?.focus?.()
    return
  }

  const input = byId("assignManualDacInput")
  if (input) {
    input.focus()
    input.select?.()
    return
  }

  const preferred = byId("btnAssignManualDacApply") || byId("btnAssignManualDacCancel")
  preferred?.focus?.()
}

function clearWeaponPromptUi() {
  const list = byId("assignWeaponPromptList")
  if (list) {
    list.innerHTML = ""
  }
  const skipBtn = byId("btnAssignWeaponPromptSkip")
  if (skipBtn) {
    skipBtn.classList.add("hidden")
  }
  setWeaponPromptOpen(false)
}

function clearManualDacPromptUi() {
  state.manualDacSubmitInProgress = false
  state.manualDacContext = null
  setManualDacError("")
  setManualDacPromptOpen(false)
}

function submitWeaponPromptSelection(optionId) {
  clearWeaponPromptUi()
  if (!window.api || typeof window.api.submitAssignDamageWeaponSelection !== "function") return
  window.api.submitAssignDamageWeaponSelection({
    optionId: String(optionId || "")
  })
}

function cancelWeaponPrompt() {
  clearWeaponPromptUi()
  if (!window.api || typeof window.api.cancelAssignDamageWeaponSelection !== "function") return
  window.api.cancelAssignDamageWeaponSelection({
    canceled: true
  })
}

function skipWeaponPrompt() {
  clearWeaponPromptUi()
  if (!window.api || typeof window.api.cancelAssignDamageWeaponSelection !== "function") return
  window.api.cancelAssignDamageWeaponSelection({
    skip: true
  })
}

function applyWeaponPrompt(payload) {
  const overlay = byId("assignWeaponPromptOverlay")
  const titleEl = byId("assignWeaponPromptTitle")
  const leadEl = byId("assignWeaponPromptLead")
  const listEl = byId("assignWeaponPromptList")

  if (!overlay || !listEl) return

  if (!payload || typeof payload !== "object") {
    clearWeaponPromptUi()
    return
  }

  const title = String(payload.title || "").trim() || "Choose Weapon Hit"
  const lead = String(payload.lead || "").trim() || "Select which box is hit."
  const options = Array.isArray(payload.options) ? payload.options : []
  const allowSkip = payload.allowSkip === true

  titleEl.textContent = title
  leadEl.textContent = lead
  listEl.innerHTML = ""

  const skipBtn = byId("btnAssignWeaponPromptSkip")
  if (skipBtn) {
    skipBtn.classList.toggle("hidden", !allowSkip)
  }

  for (const option of options) {
    if (!option || typeof option !== "object") continue

    const button = document.createElement("button")
    button.type = "button"
    button.className = "weaponSelectButton"

    const primary = document.createElement("span")
    primary.textContent = String(option.label || "Weapon")
    button.appendChild(primary)

    const detailText = String(option.detail || "").trim()
    if (detailText) {
      const detail = document.createElement("small")
      detail.textContent = detailText
      button.appendChild(detail)
    }

    button.addEventListener("click", () => {
      submitWeaponPromptSelection(option.optionId)
    })
    listEl.appendChild(button)
  }

  setWeaponPromptOpen(true)
}

function applyManualDacPrompt(payload) {
  const overlay = byId("assignManualDacOverlay")
  const lead = byId("assignManualDacLead")
  const input = byId("assignManualDacInput")
  if (!overlay) return

  if (!payload || typeof payload !== "object") {
    clearManualDacPromptUi()
    return
  }

  state.manualDacContext = payload
  state.manualDacSubmitInProgress = false

  const complete = payload.complete === true
  setManualDacCompleteMode(complete, payload.completionText)

  if (complete) {
    if (lead) lead.textContent = "Manual DAC"
    setManualDacError("")
    setManualDacPromptOpen(true)
    return
  }

  const hitIndex = Math.max(1, normalizeInteger(payload.hitIndex, 1))
  const totalHits = Math.max(1, normalizeInteger(payload.totalHits, 1))
  const defaultRoll = Math.min(12, Math.max(2, normalizeInteger(payload.defaultRoll, 7)))

  if (lead) {
    lead.textContent = totalHits > 1
      ? `Internal ${hitIndex} of ${totalHits}`
      : "Internal Hit"
  }

  if (input) {
    input.value = String(defaultRoll)
  }

  setManualDacError(String(payload.errorText || ""))
  setManualDacPromptOpen(true)
}

function closeAssignDamageWindow() {
  if (window.api && typeof window.api.closeAssignDamageWindow === "function") {
    window.api.closeAssignDamageWindow()
  } else {
    window.close()
  }
}

function cancelManualDacPrompt() {
  const complete = state.manualDacContext?.complete === true
  clearManualDacPromptUi()
  if (!complete && window.api && typeof window.api.cancelAssignDamageManualDacPrompt === "function") {
    window.api.cancelAssignDamageManualDacPrompt({
      reason: "user-canceled"
    })
  }
  closeAssignDamageWindow()
}

function submitManualDacRoll() {
  if (state.manualDacSubmitInProgress) return
  if (state.manualDacContext?.complete === true) return

  const input = byId("assignManualDacInput")
  const roll = normalizeInteger(input?.value, NaN)
  if (!Number.isFinite(roll) || roll < 2 || roll > 12) {
    setManualDacError("Enter a whole number between 2 and 12.")
    input?.focus()
    input?.select?.()
    return
  }

  if (!window.api || typeof window.api.submitAssignDamageManualDacPrompt !== "function") {
    setManualDacError("Manual DAC prompt IPC is not available.")
    return
  }

  state.manualDacSubmitInProgress = true
  try {
    setManualDacError("")
    setManualDacPromptOpen(false)
    window.api.submitAssignDamageManualDacPrompt({
      roll,
      hitIndex: normalizeInteger(state.manualDacContext?.hitIndex, 1),
      totalHits: normalizeInteger(state.manualDacContext?.totalHits, 1)
    })
  } finally {
    state.manualDacSubmitInProgress = false
  }
}

function renderShieldOptions(options, selectedShieldNumber) {
  const select = byId("assignShieldNumber")
  if (!select) return
  select.innerHTML = ""

  const list = Array.isArray(options) ? options : []
  for (const item of list) {
    const option = document.createElement("option")
    option.value = String(item?.number ?? "")
    const printed = Number.isFinite(Number(item?.value)) ? Math.trunc(Number(item.value)) : null
    option.textContent = printed !== null && printed > 0
      ? `Shield #${item.number} (${printed})`
      : `Shield #${item.number}`
    select.appendChild(option)
  }

  const wanted = String(selectedShieldNumber ?? "")
  if (wanted && list.some((item) => String(item?.number ?? "") === wanted)) {
    select.value = wanted
  } else if (list.length > 0) {
    select.value = String(list[0].number)
  }
}

function syncDamageModeSelect(selectedMode) {
  const select = byId("assignDamageMode")
  if (!select) return
  const normalized = normalizeDamageMode(selectedMode)
  select.value = normalized
}

function renderPhaserList(phasers, selectedKeys) {
  const container = byId("assignNonBearingPhasersList")
  const meta = byId("assignPhaserListMeta")
  if (!container) return

  const list = Array.isArray(phasers) ? phasers : []
  const selected = new Set(Array.isArray(selectedKeys) ? selectedKeys.map((k) => String(k)) : [])
  container.innerHTML = ""

  if (meta) {
    meta.textContent = list.length > 0
      ? `Select any bearing phasers (${list.length} total).`
      : "No phasers found in the active ship JSON."
  }

  if (list.length === 0) {
    const empty = document.createElement("div")
    empty.className = "phaserChoiceEmpty"
    empty.textContent = "No phaser entries were found in this ship's SSD data."
    container.appendChild(empty)
    updatePhaserDropdownSummary()
    syncPhaserDropdownToggleState()
    return
  }

  for (const phaser of list) {
    const row = document.createElement("label")
    row.className = "phaserChoiceItem"

    const checkbox = document.createElement("input")
    checkbox.type = "checkbox"
    checkbox.value = String(phaser?.key || "")
    checkbox.checked = selected.has(String(phaser?.key || ""))

    const text = document.createElement("div")
    text.className = "phaserChoiceLabel"
    const strong = document.createElement("strong")
    strong.textContent = String(phaser?.shortLabel || phaser?.fullLabel || "Phaser")
    const sub = document.createElement("span")
    sub.textContent = `Arc: ${String(phaser?.arc || "?")}`
    text.appendChild(strong)
    text.appendChild(sub)

    row.appendChild(checkbox)
    row.appendChild(text)
    container.appendChild(row)
  }

  updatePhaserDropdownSummary()
  syncPhaserDropdownToggleState()
}

function collectSelectedPhaserKeys() {
  const container = byId("assignNonBearingPhasersList")
  if (!container) return []
  const checked = container.querySelectorAll('input[type="checkbox"]:checked')
  const out = []
  for (const el of checked) {
    const key = String(el.value || "")
    if (!key) continue
    out.push(key)
  }
  return out
}

function sendPreviewSelection() {
  const ctx = state.context || {}
  if (ctx.available !== true) return
  if (!window.api || typeof window.api.previewAssignDamageWindowSelection !== "function") return

  window.api.previewAssignDamageWindowSelection({
    shipIndex: normalizeInteger(ctx.shipIndex, -1),
    selectedNonBearingPhaserKeys: collectSelectedPhaserKeys()
  })
}

function applyContext(context) {
  state.context = context && typeof context === "object" ? context : null
  state.submitInProgress = false
  const ctx = state.context || {}

  const shipDisplay = byId("assignShipDisplay")
  const totalDamageInput = byId("assignTotalDamage")

  if (shipDisplay) {
    shipDisplay.value = String(ctx.shipLabel || "No active ship")
  }

  renderShieldOptions(ctx.shieldOptions || [], ctx.selectedShieldNumber)
  syncDamageModeSelect(ctx.selectedAssignmentMode)

  if (totalDamageInput) {
    totalDamageInput.value = ctx.totalDamage != null && ctx.totalDamage !== "" ? String(ctx.totalDamage) : ""
  }

  const disabled = ctx.available !== true
  setControlsDisabled(disabled)

  if (ctx.hintText) {
    setHint(ctx.hintText, !!ctx.hintError)
  } else if (disabled) {
    setHint("Load and select a ship in the main StarHold window first.")
  } else {
    setHint("Enter damage and choose the shield hit.")
  }

  if (!disabled && totalDamageInput && !isWeaponPromptOpen() && !isManualDacPromptOpen()) {
    totalDamageInput.focus()
    totalDamageInput.select?.()
  }
}

function closeWindow() {
  if (isManualDacPromptOpen()) {
    cancelManualDacPrompt()
    return
  }
  if (isWeaponPromptOpen()) {
    cancelWeaponPrompt()
    return
  }
  if (isModePromptOpen()) {
    resolveModePrompt(null)
  }
  closeAssignDamageWindow()
}

async function submitAssignDamage() {
  const ctx = state.context || {}
  if (ctx.available !== true) {
    setHint("No active ship is available to assign damage.", true)
    return
  }

  if (state.submitInProgress) return

  const totalDamageInput = byId("assignTotalDamage")
  const shieldSelect = byId("assignShieldNumber")
  const modeSelect = byId("assignDamageMode")

  const totalDamage = normalizeInteger(totalDamageInput?.value, 0)
  const shieldNumber = normalizeInteger(shieldSelect?.value, 0)
  const assignmentMode = normalizeDamageMode(modeSelect?.value)

  if (!Number.isFinite(totalDamage) || totalDamage <= 0) {
    setHint("Enter a total damage value greater than 0.", true)
    totalDamageInput?.focus()
    return
  }
  if (!Number.isFinite(shieldNumber) || shieldNumber <= 0) {
    setHint("Select a valid shield.", true)
    shieldSelect?.focus()
    return
  }

  if (!window.api || typeof window.api.submitAssignDamageWindow !== "function") {
    setHint("Assign Damage IPC is not available.", true)
    return
  }

  state.submitInProgress = true
  try {
    setControlsDisabled(true)
    window.api.submitAssignDamageWindow({
      shipIndex: normalizeInteger(ctx.shipIndex, -1),
      totalDamage,
      shieldNumber,
      assignmentMode
    })
  } catch {
    state.submitInProgress = false
    setControlsDisabled(false)
  }
}

function init() {
  const btnClose = byId("btnAssignDamageClose")
  if (btnClose) btnClose.onclick = () => closeWindow()

  const btnCancel = byId("btnAssignDamageCancel")
  if (btnCancel) btnCancel.onclick = () => closeWindow()

  const btnSubmit = byId("btnAssignDamageSubmit")
  if (btnSubmit) btnSubmit.onclick = () => submitAssignDamage()

  const totalDamageInput = byId("assignTotalDamage")
  if (totalDamageInput) {
    totalDamageInput.addEventListener("keydown", (event) => {
      if (event.key === "Enter") submitAssignDamage()
    })
  }

  const modePromptOverlay = byId("assignModePromptOverlay")
  if (modePromptOverlay) {
    modePromptOverlay.addEventListener("click", (event) => {
      if (event.target === modePromptOverlay) {
        resolveModePrompt(null)
      }
    })
  }

  const btnModePromptClose = byId("btnAssignModePromptClose")
  if (btnModePromptClose) btnModePromptClose.addEventListener("click", () => resolveModePrompt(null))

  const btnModePromptCancel = byId("btnAssignModePromptCancel")
  if (btnModePromptCancel) btnModePromptCancel.addEventListener("click", () => resolveModePrompt(null))

  const btnModeManual = byId("btnAssignModeManual")
  if (btnModeManual) btnModeManual.addEventListener("click", () => resolveModePrompt("manual"))

  const btnModeAutomatic = byId("btnAssignModeAutomatic")
  if (btnModeAutomatic) btnModeAutomatic.addEventListener("click", () => resolveModePrompt("automatic"))

  const weaponPromptOverlay = byId("assignWeaponPromptOverlay")
  if (weaponPromptOverlay) {
    weaponPromptOverlay.addEventListener("click", (event) => {
      if (event.target === weaponPromptOverlay) {
        cancelWeaponPrompt()
      }
    })
  }

  const btnWeaponPromptClose = byId("btnAssignWeaponPromptClose")
  if (btnWeaponPromptClose) btnWeaponPromptClose.addEventListener("click", () => cancelWeaponPrompt())

  const btnWeaponPromptSkip = byId("btnAssignWeaponPromptSkip")
  if (btnWeaponPromptSkip) btnWeaponPromptSkip.addEventListener("click", () => skipWeaponPrompt())

  const btnWeaponPromptCancel = byId("btnAssignWeaponPromptCancel")
  if (btnWeaponPromptCancel) btnWeaponPromptCancel.addEventListener("click", () => cancelWeaponPrompt())

  const manualDacOverlay = byId("assignManualDacOverlay")
  if (manualDacOverlay) {
    manualDacOverlay.addEventListener("click", (event) => {
      if (event.target === manualDacOverlay) {
        cancelManualDacPrompt()
      }
    })
  }

  const btnManualDacClose = byId("btnAssignManualDacClose")
  if (btnManualDacClose) btnManualDacClose.addEventListener("click", () => cancelManualDacPrompt())

  const btnManualDacCancel = byId("btnAssignManualDacCancel")
  if (btnManualDacCancel) btnManualDacCancel.addEventListener("click", () => cancelManualDacPrompt())

  const btnManualDacApply = byId("btnAssignManualDacApply")
  if (btnManualDacApply) btnManualDacApply.addEventListener("click", () => submitManualDacRoll())

  const manualDacInput = byId("assignManualDacInput")
  if (manualDacInput) {
    manualDacInput.addEventListener("input", () => setManualDacError(""))
    manualDacInput.addEventListener("keydown", (event) => {
      if (event.key === "Enter") {
        event.preventDefault()
        submitManualDacRoll()
      } else if (event.key === "Escape") {
        event.preventDefault()
        cancelManualDacPrompt()
      }
    })
  }

  window.addEventListener("keydown", (event) => {
    if (event.key === "Escape") {
      if (isManualDacPromptOpen()) {
        event.preventDefault()
        cancelManualDacPrompt()
        return
      }
      if (isWeaponPromptOpen()) {
        event.preventDefault()
        cancelWeaponPrompt()
        return
      }
      if (isModePromptOpen()) {
        event.preventDefault()
        resolveModePrompt(null)
        return
      }
      closeWindow()
    }
  })

  if (window.api && typeof window.api.onAssignDamageWindowState === "function") {
    window.api.onAssignDamageWindowState((payload) => {
      applyContext(payload)
    })
  }

  if (window.api && typeof window.api.onAssignDamageWeaponPrompt === "function") {
    window.api.onAssignDamageWeaponPrompt((payload) => {
      applyWeaponPrompt(payload)
    })
  }

  if (window.api && typeof window.api.onAssignDamageManualDacPrompt === "function") {
    window.api.onAssignDamageManualDacPrompt((payload) => {
      applyManualDacPrompt(payload)
    })
  }

  applyContext({
    available: false,
    shipLabel: "No active ship",
    shieldOptions: [],
    selectedAssignmentMode: "manual",
    phasers: [],
    hintText: "Waiting for active ship data from the main StarHold window..."
  })
}

init()
