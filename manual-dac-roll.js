const state = {
  context: null,
  submitInProgress: false
}

function byId(id) {
  return document.getElementById(id)
}

function normalizeInteger(value, fallback = 0) {
  const num = Number(value)
  if (!Number.isFinite(num)) return fallback
  return Math.trunc(num)
}

function setError(text) {
  const el = byId("manualDacRollError")
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

function setCompleteMode(enabled, completionText = "Damage complete") {
  const complete = enabled === true
  const closeButton = byId("btnManualDacRollClose")
  if (closeButton) {
    closeButton.setAttribute("aria-label", complete ? "Close manual DAC window" : "Cancel manual DAC allocation")
  }

  setManualDacElementHidden("manualDacRollPromptText", complete)
  setManualDacElementHidden("manualDacRollInputField", complete)
  setManualDacElementHidden("manualDacRollError", complete)
  setManualDacElementHidden("manualDacRollActions", complete)
  setManualDacElementHidden("manualDacRollComplete", !complete)

  const completeEl = byId("manualDacRollComplete")
  if (completeEl) {
    const text = String(completionText || "").trim()
    completeEl.textContent = text || "Damage complete"
  }
}

function applyContext(context) {
  state.context = context && typeof context === "object" ? context : {}
  const complete = state.context.complete === true
  setCompleteMode(complete, state.context.completionText)

  const lead = byId("manualDacRollLead")
  if (complete) {
    if (lead) lead.textContent = "Manual DAC"
    setError("")
    return
  }

  const hitIndex = Math.max(1, normalizeInteger(state.context.hitIndex, 1))
  const totalHits = Math.max(1, normalizeInteger(state.context.totalHits, 1))
  const defaultRoll = normalizeInteger(state.context.defaultRoll, 7)

  if (lead) {
    lead.textContent = totalHits > 1
      ? `Internal ${hitIndex} of ${totalHits}`
      : "Internal Hit"
  }

  const input = byId("manualDacRollInput")
  if (input) {
    input.value = String(Math.min(12, Math.max(2, defaultRoll)))
    setTimeout(() => {
      input.focus()
      input.select?.()
    }, 0)
  }

  setError(String(state.context.errorText || ""))
}

function closeWindow() {
  if (window.api && typeof window.api.closeManualDacRollWindow === "function") {
    window.api.closeManualDacRollWindow()
  } else {
    window.close()
  }
}

function cancelRemaining() {
  if (
    state.context?.complete !== true &&
    window.api &&
    typeof window.api.cancelManualDacRollWindow === "function"
  ) {
    window.api.cancelManualDacRollWindow({
      reason: "user-canceled"
    })
  }
  closeWindow()
}

function submitRoll() {
  if (state.submitInProgress) return
  if (state.context?.complete === true) return
  const input = byId("manualDacRollInput")
  const roll = normalizeInteger(input?.value, NaN)
  if (!Number.isFinite(roll) || roll < 2 || roll > 12) {
    setError("Enter a whole number between 2 and 12.")
    input?.focus()
    input?.select?.()
    return
  }

  if (!window.api || typeof window.api.submitManualDacRollWindow !== "function") {
    setError("Manual DAC Roll IPC is not available.")
    return
  }

  state.submitInProgress = true
  try {
    setError("")
    window.api.submitManualDacRollWindow({
      roll,
      hitIndex: normalizeInteger(state.context?.hitIndex, 1),
      totalHits: normalizeInteger(state.context?.totalHits, 1)
    })
  } finally {
    state.submitInProgress = false
  }
}

function init() {
  const btnClose = byId("btnManualDacRollClose")
  if (btnClose) btnClose.onclick = () => cancelRemaining()

  const btnCancel = byId("btnManualDacRollCancel")
  if (btnCancel) btnCancel.onclick = () => cancelRemaining()

  const btnApply = byId("btnManualDacRollApply")
  if (btnApply) btnApply.onclick = () => submitRoll()

  const input = byId("manualDacRollInput")
  if (input) {
    input.addEventListener("input", () => setError(""))
    input.addEventListener("keydown", (event) => {
      if (event.key === "Enter") {
        event.preventDefault()
        submitRoll()
      } else if (event.key === "Escape") {
        event.preventDefault()
        cancelRemaining()
      }
    })
  }

  if (window.api && typeof window.api.onManualDacRollWindowState === "function") {
    window.api.onManualDacRollWindowState((payload) => {
      applyContext(payload)
    })
  }

  window.addEventListener("keydown", (event) => {
    if (event.key === "Escape") {
      event.preventDefault()
      cancelRemaining()
    }
  })
}

init()
