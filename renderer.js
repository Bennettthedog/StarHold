const DAMAGE_ALLOCATION_CHART = {
  caption: "(D4.21) DAMAGE ALLOCATION CHART",
  columns: ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M"],
  rows: [
    {
      dieRoll: "2",
      cells: [
        "Bridge", "Flag\nBridge", "Sensor", "Damage\nControl", "A Hull", "Left\nW En", "Trans",
        "Tractor", "Shuttle", "Lab", "F Hull", "Right\nW En", "Excess\nDamage"
      ]
    },
    {
      dieRoll: "3",
      cells: [
        "Drone", "Phaser", "Impulse", "Left\nW En", "Right\nW En", "A Hull", "Shuttle",
        "Damage\nControl", "Center\nW En", "Lab", "Battery", "Phaser", "Excess\nDamage"
      ]
    },
    {
      dieRoll: "4",
      cells: [
        "Phaser", "Trans", "Right\nW En", "Impulse", "F Hull", "A Hull", "Left\nW En",
        "APR", "Lab", "Trans", "Probe", "Center\nW En", "Excess\nDamage"
      ]
    },
    {
      dieRoll: "5",
      cells: [
        "Right\nW En", "A Hull", "Cargo", "Battery", "Shuttle", "Torp", "Left\nW En",
        "Impulse", "Right\nW En", "Tractor", "Probe", "Any\nWeapon", "Excess\nDamage"
      ]
    },
    {
      dieRoll: "6",
      cells: [
        "F Hull", "Impulse", "Lab", "Left\nW En", "Sensor", "Tractor", "Shuttle",
        "Right\nW En", "Phaser", "Trans", "Battery", "Any\nWeapon", "Excess\nDamage"
      ]
    },
    {
      dieRoll: "7",
      cells: [
        "Cargo", "F Hull", "Battery", "Center\nW En", "Shuttle", "APR", "Lab",
        "Phaser", "Any\nW En", "Probe", "A Hull", "Any\nWeapon", "Excess\nDamage"
      ]
    },
    {
      dieRoll: "8",
      cells: [
        "A Hull", "APR", "Shuttle", "Right\nW En", "Scanner", "Tractor", "Lab",
        "Left\nW En", "Phaser", "Trans", "Battery", "Any\nWeapon", "Excess\nDamage"
      ]
    },
    {
      dieRoll: "9",
      cells: [
        "Left\nW En", "F Hull", "Cargo", "Battery", "Lab", "Drone", "Right\nW En",
        "Impulse", "Left\nW En", "Tractor", "Probe", "Any\nWeapon", "Excess\nDamage"
      ]
    },
    {
      dieRoll: "10",
      cells: [
        "Phaser", "Tractor", "Left\nW En", "Impulse", "A Hull", "F Hull", "Right\nW En",
        "APR", "Lab", "Trans", "Probe", "Center\nW En", "Excess\nDamage"
      ]
    },
    {
      dieRoll: "11",
      cells: [
        "Torp", "Phaser", "Impulse", "Right\nW En", "Left\nW En", "F Hull", "Tractor",
        "Damage\nControl", "Center\nW En", "Lab", "Battery", "Phaser", "Excess\nDamage"
      ]
    },
    {
      dieRoll: "12",
      cells: [
        "Aux\nCon", "Emer\nBridge", "Scanner", "Probe", "F Hull", "Right\nW En", "Trans",
        "Shuttle", "Tractor", "Lab", "A Hull", "Left\nW En", "Excess\nDamage"
      ]
    }
  ]
}

const UNDERLINED_DAMAGE_ALLOCATION_CELLS = new Set([
  "2:A", "2:B", "2:C", "2:D", "2:E",
  "3:A", "3:B", "3:H",
  "4:A", "4:B",
  "5:A", "5:F",
  "6:E",
  "8:E",
  "9:A", "9:F",
  "10:A", "10:B",
  "11:A", "11:B", "11:H",
  "12:A", "12:B", "12:C", "12:D", "12:E"
])

function isUnderlinedDamageCell(dieRoll, column) {
  return UNDERLINED_DAMAGE_ALLOCATION_CELLS.has(`${String(dieRoll)}:${String(column)}`)
}

const DAC_LABEL_FALLBACK_SSD_KEYS = {
  "a hull": ["Aft Hull"],
  "apr": ["APR"],
  "any w en": ["Left Warp", "Center Warp", "Right Warp"],
  "any weapon": ["phaser", "heavy", "Drone", "Pseudo Plasma Torpedo"],
  "aux con": ["Auxiliary Control", "Aux Con"],
  "battery": ["Battery"],
  "bridge": ["Bridge"],
  "cargo": ["Cargo"],
  "center w en": ["Center Warp"],
  "damage control": ["Damage Control"],
  "drone": ["Drone"],
  "emer bridge": ["Emergency Bridge", "Emer Bridge"],
  "excess damage": [],
  "f hull": ["Forward Hull"],
  "flag bridge": ["Flag Bridge"],
  "impulse": ["Impulse"],
  "lab": ["Lab"],
  "left w en": ["Left Warp"],
  "phaser": ["phaser"],
  "probe": ["Probe"],
  "right w en": ["Right Warp"],
  "scanner": ["Scanner"],
  "sensor": ["Sensor"],
  "shuttle": ["Shuttle", "Admin Shuttles"],
  "torp": ["heavy", "Pseudo Plasma Torpedo"],
  "tractor": ["Tractor"],
  "trans": ["Transporter"]
}

const RANK_ORDERED_SSD_KEYS = new Set(["sensor", "scanner", "damage control"])
const USER_SELECTABLE_WEAPON_GROUP_KEYS = new Set(["phaser", "heavy", "drone"])
const WEAPON_SELECTION_SKIP_TOKEN = "__STARHOLD_WEAPON_SKIP__"
const SHIP_ROTATION_INTERVAL_MS = 7000
const DUPLICATE_SHIP_MARKER_COLORS = ["#d92b2b", "#2f6be6", "#2d9b47", "#8b5a2b", "#7a3bd2", "#e24e9c", "#e2b91f"]
const DUPLICATE_SHIP_MARKER_COLOR_NAMES = Object.freeze({
  "#d92b2b": "red",
  "#2f6be6": "blue",
  "#2d9b47": "green",
  "#8b5a2b": "brown",
  "#7a3bd2": "purple",
  "#e24e9c": "pink",
  "#e2b91f": "yellow"
})
const DUPLICATE_SHIP_MARKER_PAIR_COLORS = []
for (let i = 0; i < DUPLICATE_SHIP_MARKER_COLORS.length; i += 1) {
  for (let j = i + 1; j < DUPLICATE_SHIP_MARKER_COLORS.length; j += 1) {
    DUPLICATE_SHIP_MARKER_PAIR_COLORS.push([
      DUPLICATE_SHIP_MARKER_COLORS[i],
      DUPLICATE_SHIP_MARKER_COLORS[j]
    ])
  }
}

const DAMAGE_OVERLAY_PRESETS = Object.freeze({
  classic: {
    label: "Classic",
    shield: {
      stroke: "rgba(255, 74, 74, 0.98)",
      fill: "rgba(255, 48, 48, 0.34)",
      shadow: "rgba(255, 72, 72, 0.45)"
    },
    armor: {
      stroke: "rgba(255, 168, 60, 0.98)",
      fill: "rgba(255, 150, 48, 0.28)",
      shadow: "rgba(255, 158, 64, 0.42)"
    },
    internal: {
      stroke: "rgba(255, 116, 206, 0.98)",
      fill: "rgba(255, 90, 190, 0.26)",
      shadow: "rgba(255, 108, 196, 0.35)"
    }
  },
  crimson: {
    label: "Crimson",
    shield: {
      stroke: "rgba(255, 74, 74, 0.98)",
      fill: "rgba(255, 56, 56, 0.34)",
      shadow: "rgba(255, 74, 74, 0.44)"
    },
    armor: {
      stroke: "rgba(236, 86, 86, 0.98)",
      fill: "rgba(236, 86, 86, 0.3)",
      shadow: "rgba(236, 86, 86, 0.4)"
    },
    internal: {
      stroke: "rgba(212, 70, 70, 0.98)",
      fill: "rgba(212, 70, 70, 0.28)",
      shadow: "rgba(212, 70, 70, 0.36)"
    }
  },
  orange: {
    label: "Orange",
    shield: {
      stroke: "rgba(255, 164, 74, 0.98)",
      fill: "rgba(255, 152, 66, 0.34)",
      shadow: "rgba(255, 160, 72, 0.44)"
    },
    armor: {
      stroke: "rgba(255, 140, 58, 0.98)",
      fill: "rgba(255, 130, 52, 0.3)",
      shadow: "rgba(255, 140, 58, 0.4)"
    },
    internal: {
      stroke: "rgba(235, 118, 46, 0.98)",
      fill: "rgba(235, 118, 46, 0.28)",
      shadow: "rgba(235, 118, 46, 0.36)"
    }
  },
  gold: {
    label: "Gold",
    shield: {
      stroke: "rgba(255, 210, 76, 0.98)",
      fill: "rgba(255, 198, 62, 0.34)",
      shadow: "rgba(255, 206, 70, 0.44)"
    },
    armor: {
      stroke: "rgba(245, 186, 64, 0.98)",
      fill: "rgba(245, 186, 64, 0.3)",
      shadow: "rgba(245, 186, 64, 0.4)"
    },
    internal: {
      stroke: "rgba(222, 166, 48, 0.98)",
      fill: "rgba(222, 166, 48, 0.28)",
      shadow: "rgba(222, 166, 48, 0.36)"
    }
  },
  emerald: {
    label: "Emerald",
    shield: {
      stroke: "rgba(84, 218, 122, 0.98)",
      fill: "rgba(74, 206, 112, 0.34)",
      shadow: "rgba(84, 214, 122, 0.44)"
    },
    armor: {
      stroke: "rgba(60, 194, 102, 0.98)",
      fill: "rgba(60, 194, 102, 0.3)",
      shadow: "rgba(60, 194, 102, 0.4)"
    },
    internal: {
      stroke: "rgba(48, 168, 88, 0.98)",
      fill: "rgba(48, 168, 88, 0.28)",
      shadow: "rgba(48, 168, 88, 0.36)"
    }
  },
  cyan: {
    label: "Cyan",
    shield: {
      stroke: "rgba(96, 216, 246, 0.98)",
      fill: "rgba(88, 206, 236, 0.34)",
      shadow: "rgba(94, 212, 242, 0.44)"
    },
    armor: {
      stroke: "rgba(72, 194, 224, 0.98)",
      fill: "rgba(72, 194, 224, 0.3)",
      shadow: "rgba(72, 194, 224, 0.4)"
    },
    internal: {
      stroke: "rgba(60, 170, 200, 0.98)",
      fill: "rgba(60, 170, 200, 0.28)",
      shadow: "rgba(60, 170, 200, 0.36)"
    }
  },
  violet: {
    label: "Violet",
    shield: {
      stroke: "rgba(170, 120, 248, 0.98)",
      fill: "rgba(158, 108, 238, 0.34)",
      shadow: "rgba(166, 118, 244, 0.44)"
    },
    armor: {
      stroke: "rgba(146, 98, 224, 0.98)",
      fill: "rgba(146, 98, 224, 0.3)",
      shadow: "rgba(146, 98, 224, 0.4)"
    },
    internal: {
      stroke: "rgba(124, 84, 198, 0.98)",
      fill: "rgba(124, 84, 198, 0.28)",
      shadow: "rgba(124, 84, 198, 0.36)"
    }
  }
})
const DEFAULT_DAMAGE_OVERLAY_PRESET = "classic"
const DAMAGE_OVERLAY_STORAGE_KEY = "starhold.damageOverlayPreset"
const STARHOLD_SAVE_SCHEMA = "starhold.game"
const STARHOLD_SAVE_VERSION = 1

const state = {
  doc: null,
  image: null,
  imageDataUrl: null,
  imagePath: "",
  jsonPath: "",
  shipLabel: "",
  crop: null,
  defaultShipFolder: "",
  ships: [],
  activeShipIndex: -1,
  rollsConfig: null,
  rollsLabelEntriesByChartLabel: new Map(),
  rollsLabelEntriesByLabel: new Map(),
  manualDacPromptPromise: null,
  manualDacPromptResolver: null,
  weaponSelectPromise: null,
  weaponSelectResolver: null,
  assignDamageWindowOpen: false,
  weaponSelectWindowOpen: false,
  assignDamageWeaponPromptSyncPromise: null,
  weaponSelectWindowClosingPromise: null,
  ignoreNextWeaponSelectWindowClosed: false,
  shipRotationTimerId: null,
  pendingManualDacChartClearShip: null,
  manualDamageMode: "",
  zoomSectionMode: false,
  zoomDrag: null,
  damageOverlayPreset: DEFAULT_DAMAGE_OVERLAY_PRESET,
  appVersion: ""
}

function byId(id) {
  return document.getElementById(id)
}

function renderAppVersion() {
  const el = byId("appVersion")
  if (!el) return

  const version = String(state.appVersion || "").trim()
  if (!version) {
    el.textContent = ""
    el.style.display = "none"
    el.onclick = null
    return
  }

  el.textContent = `v${version}`
  el.style.display = "inline-block"
  el.style.cursor = "pointer"
  el.style.userSelect = "none"
  el.title = "Check for updates"
  el.onclick = () => {
    handleAppVersionClick()
  }
}

async function handleAppVersionClick() {
  const ok = typeof window.confirm === "function"
    ? window.confirm("Check for updates now?")
    : true
  if (!ok) return

  if (window.api && typeof window.api.checkForUpdates === "function") {
    try {
      const res = await window.api.checkForUpdates()
      if (!res || !res.ok) {
        if (typeof window.alert === "function") {
          window.alert(res?.error || "Failed to check for updates.")
        }
        return
      }

      if (res.openedExternal === true) {
        setStatus("Opened StarHold releases in your browser.")
        return
      }

      setStatus("Checking for updates...")
      return
    } catch {
      if (typeof window.alert === "function") {
        window.alert("Failed to check for updates.")
      }
      return
    }
  }

  if (window.api && typeof window.api.openUpdates === "function") {
    try {
      const res = await window.api.openUpdates()
      if (!res || !res.ok) {
        if (typeof window.alert === "function") {
          window.alert(res?.error || "Failed to open StarHold releases.")
        }
        return
      }
      setStatus("Opened StarHold releases in your browser.")
    } catch {
      if (typeof window.alert === "function") {
        window.alert("Failed to open StarHold releases.")
      }
    }
  }
}

function normalizeLabelToken(value) {
  return String(value || "")
    .replace(/\r?\n/g, " ")
    .replace(/[_-]+/g, " ")
    .replace(/[^\w# ]+/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase()
}

function setRollsConfig(rollsConfig) {
  state.rollsConfig = rollsConfig && typeof rollsConfig === "object" ? rollsConfig : null
  state.rollsLabelEntriesByChartLabel = new Map()
  state.rollsLabelEntriesByLabel = new Map()

  const labels = Array.isArray(state.rollsConfig?.labels) ? state.rollsConfig.labels : []
  for (const entry of labels) {
    if (!entry || typeof entry !== "object") continue
    const chartLabelKey = normalizeLabelToken(entry.chartLabel)
    const labelKey = normalizeLabelToken(entry.label)
    if (chartLabelKey && !state.rollsLabelEntriesByChartLabel.has(chartLabelKey)) {
      state.rollsLabelEntriesByChartLabel.set(chartLabelKey, entry)
    }
    if (labelKey && !state.rollsLabelEntriesByLabel.has(labelKey)) {
      state.rollsLabelEntriesByLabel.set(labelKey, entry)
    }
  }
}

function findRollsLabelEntry(chartCellLabel) {
  const chartKey = normalizeLabelToken(chartCellLabel)
  if (chartKey && state.rollsLabelEntriesByChartLabel.has(chartKey)) {
    return state.rollsLabelEntriesByChartLabel.get(chartKey) || null
  }
  if (chartKey && state.rollsLabelEntriesByLabel.has(chartKey)) {
    return state.rollsLabelEntriesByLabel.get(chartKey) || null
  }
  return null
}

async function loadRollsConfig() {
  if (!window.api || typeof window.api.getRollsConfig !== "function") {
    setRollsConfig(null)
    return
  }

  try {
    const res = await window.api.getRollsConfig()
    if (!res || !res.ok) {
      setRollsConfig(null)
      return
    }
    setRollsConfig(res.rollsConfig || null)
  } catch {
    setRollsConfig(null)
  }
}

async function loadAppVersion() {
  renderAppVersion()
  if (!window.api || typeof window.api.getAppVersion !== "function") return

  try {
    const res = await window.api.getAppVersion()
    if (!res || !res.ok) return
    state.appVersion = String(res.version || "").trim()
    renderAppVersion()
  } catch {
    // Ignore version load failures.
  }
}

function rollTwoSixSidedDice() {
  const die1 = 1 + Math.floor(Math.random() * 6)
  const die2 = 1 + Math.floor(Math.random() * 6)
  return {
    die1,
    die2,
    total: die1 + die2
  }
}

function setStatus(text) {
  const el = byId("statusMessage")
  if (el) el.textContent = String(text || "")
}

function clearAssignDamageManualDacPrompt() {
  if (!state.assignDamageWindowOpen) return
  if (!window.api || typeof window.api.setAssignDamageManualDacPrompt !== "function") return
  Promise.resolve(window.api.setAssignDamageManualDacPrompt(null)).catch(() => {
    // no-op
  })
}

function resolveManualDacRollPrompt(value) {
  clearAssignDamageManualDacPrompt()

  const resolver = state.manualDacPromptResolver
  state.manualDacPromptResolver = null
  state.manualDacPromptPromise = null
  if (typeof resolver === "function") {
    resolver(value)
  }
}

function clearAssignDamageWeaponPrompt() {
  if (state.assignDamageWeaponPromptSyncPromise) return state.assignDamageWeaponPromptSyncPromise
  if (!state.assignDamageWindowOpen) return Promise.resolve()
  if (!window.api || typeof window.api.setAssignDamageWeaponPrompt !== "function") return Promise.resolve()

  const syncPromise = Promise.resolve(window.api.setAssignDamageWeaponPrompt(null))
    .catch(() => {
      // no-op
    })
    .finally(() => {
      if (state.assignDamageWeaponPromptSyncPromise === syncPromise) {
        state.assignDamageWeaponPromptSyncPromise = null
      }
    })

  state.assignDamageWeaponPromptSyncPromise = syncPromise
  return syncPromise
}

function closeStandaloneWeaponSelectWindow() {
  if (state.weaponSelectWindowClosingPromise) return state.weaponSelectWindowClosingPromise
  if (!state.weaponSelectWindowOpen) return Promise.resolve()
  if (!window.api || typeof window.api.closeWeaponSelectWindow !== "function") return Promise.resolve()

  state.ignoreNextWeaponSelectWindowClosed = true

  const closePromise = Promise.resolve(window.api.closeWeaponSelectWindow())
    .then((res) => {
      if (!res || res.ok !== true) {
        state.ignoreNextWeaponSelectWindowClosed = false
      }
      return res
    })
    .catch(() => {
      state.ignoreNextWeaponSelectWindowClosed = false
      // no-op
    })
    .finally(() => {
      if (state.weaponSelectWindowClosingPromise === closePromise) {
        state.weaponSelectWindowClosingPromise = null
      }
    })

  state.weaponSelectWindowClosingPromise = closePromise
  return closePromise
}

function resolveWeaponSelectionPrompt(value) {
  clearAssignDamageWeaponPrompt()
  closeStandaloneWeaponSelectWindow()

  const overlay = byId("weaponSelectOverlay")
  if (overlay) {
    overlay.classList.add("hidden")
    overlay.setAttribute("aria-hidden", "true")
  }

  const list = byId("weaponSelectList")
  if (list) {
    list.innerHTML = ""
  }

  const skipBtn = byId("btnWeaponSelectSkip")
  if (skipBtn) {
    skipBtn.classList.add("hidden")
  }

  const resolver = state.weaponSelectResolver
  state.weaponSelectResolver = null
  state.weaponSelectPromise = null
  if (typeof resolver === "function") {
    resolver(value)
  }
}

function cancelWeaponSelectionPrompt() {
  resolveWeaponSelectionPrompt(null)
}

function setDefaultShipFolder(folderPath) {
  state.defaultShipFolder = String(folderPath || "").trim()
  updateSettingsButton()
}

function updateSettingsButton() {
  const btn = byId("btnSettings")
  if (!btn) return

  const folder = state.defaultShipFolder
  const hasFolder = folder.length > 0
  const title = hasFolder
    ? `Default ship folder:\n${folder}\nClick to change.`
    : "Choose default ship folder"

  btn.title = title
  btn.setAttribute("aria-label", hasFolder ? "Change default ship folder" : "Choose default ship folder")
}

function getActiveShipRecord() {
  const idx = Number(state.activeShipIndex)
  if (!Number.isInteger(idx) || idx < 0 || idx >= state.ships.length) return null
  return state.ships[idx] || null
}

function normalizeInteger(value, fallback = 0) {
  const num = Number(value)
  if (!Number.isFinite(num)) return fallback
  return Math.trunc(num)
}

function clampInteger(value, min, max) {
  const n = normalizeInteger(value, min)
  return Math.min(max, Math.max(min, n))
}

function normalizeDamageAssignmentMode(value) {
  const raw = String(value || "").trim().toLowerCase()
  return raw === "automatic" ? "automatic" : "manual"
}

function normalizeManualDamageMode(value) {
  const raw = String(value || "").trim().toLowerCase()
  return raw === "mark" || raw === "remove" ? raw : ""
}

function normalizeDamageOverlayPreset(value) {
  const raw = String(value || "").trim().toLowerCase()
  if (Object.prototype.hasOwnProperty.call(DAMAGE_OVERLAY_PRESETS, raw)) return raw
  return DEFAULT_DAMAGE_OVERLAY_PRESET
}

function getDamageOverlayPreset() {
  return normalizeDamageOverlayPreset(state.damageOverlayPreset)
}

function getDamageOverlayTheme(presetKey = state.damageOverlayPreset) {
  const preset = normalizeDamageOverlayPreset(presetKey)
  return DAMAGE_OVERLAY_PRESETS[preset] || DAMAGE_OVERLAY_PRESETS[DEFAULT_DAMAGE_OVERLAY_PRESET]
}

function getDamageOverlayLabel(presetKey = state.damageOverlayPreset) {
  const key = normalizeDamageOverlayPreset(presetKey)
  const preset = DAMAGE_OVERLAY_PRESETS[key]
  return String(preset?.label || key)
}

function getDamageOverlayLabelLower(presetKey = state.damageOverlayPreset) {
  return getDamageOverlayLabel(presetKey).toLowerCase()
}

function loadDamageOverlayPresetFromStorage() {
  try {
    const raw = window.localStorage?.getItem(DAMAGE_OVERLAY_STORAGE_KEY)
    return normalizeDamageOverlayPreset(raw)
  } catch {
    return DEFAULT_DAMAGE_OVERLAY_PRESET
  }
}

function storeDamageOverlayPresetToStorage(presetKey) {
  try {
    const normalized = normalizeDamageOverlayPreset(presetKey)
    window.localStorage?.setItem(DAMAGE_OVERLAY_STORAGE_KEY, normalized)
  } catch {
    // Ignore storage failures (private mode, policy blocks, etc.).
  }
}

function renderDamageColorSelector() {
  const select = byId("damageColorSelect")
  if (!select) return
  const preset = getDamageOverlayPreset()
  if (select.value !== preset) {
    select.value = preset
  }
  const label = getDamageOverlayLabel(preset)
  select.title = `Damage overlay color: ${label}`
}

function applyDamageOverlayPreset(nextPreset, options = {}) {
  const silent = options && options.silent === true
  const normalized = normalizeDamageOverlayPreset(nextPreset)
  const changed = state.damageOverlayPreset !== normalized

  state.damageOverlayPreset = normalized
  renderDamageColorSelector()
  storeDamageOverlayPresetToStorage(normalized)

  if (!changed) return

  renderShipCanvas()
  syncAssignDamageWindowStateForActiveShip()
  if (!silent) {
    setStatus(`Damage overlay color set to ${getDamageOverlayLabel(normalized)}.`)
  }
}

function getDamageAssignmentModeLabel(mode) {
  return normalizeDamageAssignmentMode(mode) === "automatic" ? "Automatic" : "Manual"
}

function shieldKeyForNumber(shieldNumber) {
  return `Shield #${normalizeInteger(shieldNumber, 1)}`
}

function getShipShieldValues(doc) {
  const shields = doc?.ssd?.shields
  if (!shields || typeof shields !== "object") return {}
  const out = {}
  for (const [key, rawVal] of Object.entries(shields)) {
    const m = /^shield\s*#\s*(\d+)$/i.exec(String(key || "").trim())
    if (!m) continue
    const n = Number(rawVal)
    if (!Number.isFinite(n)) continue
    out[String(Number(m[1]))] = Math.max(0, Math.trunc(n))
  }
  return out
}

function getPrintedShieldStrength(doc, shieldNumber) {
  const shields = getShipShieldValues(doc)
  const key = String(normalizeInteger(shieldNumber, 1))
  if (Object.prototype.hasOwnProperty.call(shields, key)) return shields[key]
  return 0
}

function getShipShieldBoxEntriesByNumber(doc) {
  if (!doc || typeof doc !== "object") return {}
  const out = {}

  function collectFromContainer(container) {
    if (!container || typeof container !== "object") return
    for (const [key, entries] of Object.entries(container)) {
      const m = /^shield\s*#\s*(\d+)$/i.exec(String(key || "").trim())
      if (!m || !Array.isArray(entries)) continue
      const shieldNumberKey = String(Number(m[1]))
      if (Array.isArray(out[shieldNumberKey]) && out[shieldNumberKey].length > 0) continue

      const list = []
      for (const entry of entries) {
        const rect = parsePosRect(entry?.pos)
        if (!rect) continue
        list.push({
          index: list.length,
          rect
        })
      }
      out[shieldNumberKey] = list
    }
  }

  // SuperLuminal ships store shield boxes under doc.ssd["Shield #N"].
  collectFromContainer(doc?.ssd)
  // Fallback for alternate layouts.
  collectFromContainer(doc)

  return out
}

function getShipShieldBoxEntries(doc, shieldNumber) {
  const key = String(normalizeInteger(shieldNumber, 1))
  const all = getShipShieldBoxEntriesByNumber(doc)
  return Array.isArray(all[key]) ? all[key] : []
}

function getShipArmorBoxGroups(doc) {
  if (!doc || typeof doc !== "object") return {}
  const out = {}

  function collectFromContainer(container) {
    if (!container || typeof container !== "object") return
    for (const [rawKey, entries] of Object.entries(container)) {
      if (!Array.isArray(entries)) continue
      const key = String(rawKey || "").trim()
      if (!/armou?r/i.test(key)) continue

      const id = key.toLowerCase()
      if (out[id] && Array.isArray(out[id].entries) && out[id].entries.length > 0) continue

      const m = /armou?r\s*#\s*(\d+)/i.exec(key)
      const list = []
      for (const entry of entries) {
        const rect = parsePosRect(entry?.pos)
        if (!rect) continue
        list.push({
          index: list.length,
          rect
        })
      }

      out[id] = {
        id,
        key,
        shieldNumber: m ? normalizeInteger(m[1], 0) : null,
        entries: list
      }
    }
  }

  collectFromContainer(doc?.ssd)
  collectFromContainer(doc)
  return out
}

function getShipArmorBoxGroupForShield(doc, shieldNumber) {
  const groupsById = getShipArmorBoxGroups(doc)
  const groups = Object.values(groupsById)
  if (groups.length === 0) return null

  const desiredShield = normalizeInteger(shieldNumber, 1)
  const exact = groups.find((group) => normalizeInteger(group?.shieldNumber, -1) === desiredShield)
  if (exact) return exact

  const generic = groups.find((group) => /^(armor|armour)$/i.test(String(group?.key || "").trim()))
  if (generic) return generic

  if (groups.length === 1) return groups[0]
  return null
}

function ensureShipRuntimeDamageState(ship) {
  if (!ship || typeof ship !== "object") {
    return {
      damagedShieldBoxIndicesByShield: {},
      damagedArmorBoxIndicesByGroup: {},
      damagedSsdBoxIndicesByKey: {},
      damagedShieldBoxOverlayPresetByShield: {},
      damagedArmorBoxOverlayPresetByGroup: {},
      damagedSsdBoxOverlayPresetByKey: {},
      markedDacCells: [],
      passedOverDacCells: [],
      activeDacCell: "",
      shipDestroyed: false,
      internalsTotal: 0
    }
  }

  if (!ship.runtimeDamage || typeof ship.runtimeDamage !== "object") {
    ship.runtimeDamage = {}
  }
  if (!ship.runtimeDamage.damagedShieldBoxIndicesByShield || typeof ship.runtimeDamage.damagedShieldBoxIndicesByShield !== "object") {
    ship.runtimeDamage.damagedShieldBoxIndicesByShield = {}
  }
  if (!ship.runtimeDamage.damagedArmorBoxIndicesByGroup || typeof ship.runtimeDamage.damagedArmorBoxIndicesByGroup !== "object") {
    ship.runtimeDamage.damagedArmorBoxIndicesByGroup = {}
  }
  if (!ship.runtimeDamage.damagedSsdBoxIndicesByKey || typeof ship.runtimeDamage.damagedSsdBoxIndicesByKey !== "object") {
    ship.runtimeDamage.damagedSsdBoxIndicesByKey = {}
  }
  if (!ship.runtimeDamage.damagedShieldBoxOverlayPresetByShield || typeof ship.runtimeDamage.damagedShieldBoxOverlayPresetByShield !== "object") {
    ship.runtimeDamage.damagedShieldBoxOverlayPresetByShield = {}
  }
  if (!ship.runtimeDamage.damagedArmorBoxOverlayPresetByGroup || typeof ship.runtimeDamage.damagedArmorBoxOverlayPresetByGroup !== "object") {
    ship.runtimeDamage.damagedArmorBoxOverlayPresetByGroup = {}
  }
  if (!ship.runtimeDamage.damagedSsdBoxOverlayPresetByKey || typeof ship.runtimeDamage.damagedSsdBoxOverlayPresetByKey !== "object") {
    ship.runtimeDamage.damagedSsdBoxOverlayPresetByKey = {}
  }
  if (!Array.isArray(ship.runtimeDamage.markedDacCells)) {
    ship.runtimeDamage.markedDacCells = []
  } else {
    ship.runtimeDamage.markedDacCells = Array.from(new Set(
      ship.runtimeDamage.markedDacCells.map((key) => String(key || "").trim()).filter(Boolean)
    ))
  }
  if (!Array.isArray(ship.runtimeDamage.passedOverDacCells)) {
    ship.runtimeDamage.passedOverDacCells = []
  } else {
    ship.runtimeDamage.passedOverDacCells = Array.from(new Set(
      ship.runtimeDamage.passedOverDacCells.map((key) => String(key || "").trim()).filter(Boolean)
    ))
  }
  ship.runtimeDamage.activeDacCell = String(ship.runtimeDamage.activeDacCell || "").trim()
  ship.runtimeDamage.shipDestroyed = ship.runtimeDamage.shipDestroyed === true
  if (!Number.isFinite(Number(ship.runtimeDamage.internalsTotal))) {
    ship.runtimeDamage.internalsTotal = 0
  } else {
    ship.runtimeDamage.internalsTotal = Math.max(0, normalizeInteger(ship.runtimeDamage.internalsTotal, 0))
  }

  return ship.runtimeDamage
}

function normalizeDamageOverlayIndexMap(value) {
  const out = {}
  if (!value || typeof value !== "object") return out
  for (const [rawIndex, rawPreset] of Object.entries(value)) {
    const index = normalizeInteger(rawIndex, -1)
    if (index < 0) continue
    out[String(index)] = normalizeDamageOverlayPreset(rawPreset)
  }
  return out
}

function syncDamageOverlayIndexMapForDamageKey(runtime, containerKey, damageKey, indices, options = {}) {
  if (!runtime || typeof runtime !== "object") return {}
  if (!runtime[containerKey] || typeof runtime[containerKey] !== "object") {
    runtime[containerKey] = {}
  }

  const normalizedDamageKey = String(damageKey || "").trim()
  const existing = normalizeDamageOverlayIndexMap(runtime[containerKey][normalizedDamageKey])
  const next = {}
  const hasPresetForNew = String(options.presetForNew || "").trim().length > 0
  const presetForNew = normalizeDamageOverlayPreset(options.presetForNew)

  for (const rawIndex of Array.isArray(indices) ? indices : []) {
    const index = normalizeInteger(rawIndex, -1)
    if (index < 0) continue
    const indexKey = String(index)
    if (Object.prototype.hasOwnProperty.call(existing, indexKey)) {
      next[indexKey] = existing[indexKey]
      continue
    }
    if (hasPresetForNew) {
      next[indexKey] = presetForNew
    }
  }

  runtime[containerKey][normalizedDamageKey] = next
  return next
}

function getOrSeedDamageOverlayPresetForDamageIndex(runtime, containerKey, damageKey, index, fallbackPreset = "") {
  if (!runtime || typeof runtime !== "object") return DEFAULT_DAMAGE_OVERLAY_PRESET
  if (!runtime[containerKey] || typeof runtime[containerKey] !== "object") {
    runtime[containerKey] = {}
  }

  const normalizedDamageKey = String(damageKey || "").trim()
  const map = normalizeDamageOverlayIndexMap(runtime[containerKey][normalizedDamageKey])
  runtime[containerKey][normalizedDamageKey] = map

  const normalizedIndex = normalizeInteger(index, -1)
  if (normalizedIndex < 0) return DEFAULT_DAMAGE_OVERLAY_PRESET

  const indexKey = String(normalizedIndex)
  if (Object.prototype.hasOwnProperty.call(map, indexKey)) {
    return normalizeDamageOverlayPreset(map[indexKey])
  }

  const seedRaw = String(fallbackPreset || "").trim()
  if (!seedRaw) return DEFAULT_DAMAGE_OVERLAY_PRESET

  const seededPreset = normalizeDamageOverlayPreset(seedRaw)
  map[indexKey] = seededPreset
  return seededPreset
}

function isShipDestroyed(ship) {
  return ensureShipRuntimeDamageState(ship).shipDestroyed === true
}

function setShipDestroyed(ship, destroyed = true) {
  const runtime = ensureShipRuntimeDamageState(ship)
  runtime.shipDestroyed = destroyed === true
  return runtime.shipDestroyed
}

function getShipMarkedDacCellSet(ship) {
  const runtime = ensureShipRuntimeDamageState(ship)
  const set = new Set()
  for (const key of Array.isArray(runtime.markedDacCells) ? runtime.markedDacCells : []) {
    const norm = String(key || "").trim()
    if (!norm) continue
    set.add(norm)
  }
  return set
}

function isDacCellMarkedOnShip(ship, dacCellKey) {
  return getShipMarkedDacCellSet(ship).has(String(dacCellKey || "").trim())
}

function markDacCellOnShip(ship, dacCellKey) {
  const runtime = ensureShipRuntimeDamageState(ship)
  const key = String(dacCellKey || "").trim()
  if (!key) return false
  const next = new Set(Array.isArray(runtime.markedDacCells) ? runtime.markedDacCells : [])
  const before = next.size
  next.add(key)
  runtime.markedDacCells = Array.from(next)
  return next.size !== before
}

function getShipPassedOverDacCellSet(ship) {
  const runtime = ensureShipRuntimeDamageState(ship)
  const set = new Set()
  for (const key of Array.isArray(runtime.passedOverDacCells) ? runtime.passedOverDacCells : []) {
    const norm = String(key || "").trim()
    if (!norm) continue
    set.add(norm)
  }
  return set
}

function markPassedOverDacCellOnShip(ship, dacCellKey) {
  const runtime = ensureShipRuntimeDamageState(ship)
  const key = String(dacCellKey || "").trim()
  if (!key) return false
  const next = new Set(Array.isArray(runtime.passedOverDacCells) ? runtime.passedOverDacCells : [])
  const before = next.size
  next.add(key)
  runtime.passedOverDacCells = Array.from(next)
  return next.size !== before
}

function getShipActiveDacCellKey(ship) {
  const runtime = ensureShipRuntimeDamageState(ship)
  return String(runtime.activeDacCell || "").trim()
}

function setShipActiveDacCell(ship, dacCellKey) {
  const runtime = ensureShipRuntimeDamageState(ship)
  const key = String(dacCellKey || "").trim()
  runtime.activeDacCell = key
  return key
}

function clearShipDacChartState(ship) {
  const runtime = ensureShipRuntimeDamageState(ship)
  runtime.markedDacCells = []
  runtime.passedOverDacCells = []
  runtime.activeDacCell = ""
}

function queueManualDacChartClearOnWindowClose(ship) {
  state.pendingManualDacChartClearShip = ship && typeof ship === "object" ? ship : null
}

function flushQueuedManualDacChartClear() {
  const ship = state.pendingManualDacChartClearShip
  state.pendingManualDacChartClearShip = null
  if (!ship || typeof ship !== "object") return false

  clearShipDacChartState(ship)
  if (ship === getActiveShipRecord()) {
    renderDamageChart()
  }
  return true
}

function getShipDamagedShieldBoxIndexSet(ship, shieldNumber) {
  const runtime = ensureShipRuntimeDamageState(ship)
  const key = String(normalizeInteger(shieldNumber, 1))
  const raw = Array.isArray(runtime.damagedShieldBoxIndicesByShield[key])
    ? runtime.damagedShieldBoxIndicesByShield[key]
    : []
  const set = new Set()
  for (const item of raw) {
    const idx = normalizeInteger(item, -1)
    if (idx < 0) continue
    set.add(idx)
  }
  return set
}

function setShipDamagedShieldBoxIndices(ship, shieldNumber, indices, options = {}) {
  const runtime = ensureShipRuntimeDamageState(ship)
  const key = String(normalizeInteger(shieldNumber, 1))
  const uniqueSorted = Array.from(new Set(
    (Array.isArray(indices) ? indices : [])
      .map((n) => normalizeInteger(n, -1))
      .filter((n) => n >= 0)
  )).sort((a, b) => a - b)
  runtime.damagedShieldBoxIndicesByShield[key] = uniqueSorted
  syncDamageOverlayIndexMapForDamageKey(
    runtime,
    "damagedShieldBoxOverlayPresetByShield",
    key,
    uniqueSorted,
    { presetForNew: options.presetForNew }
  )
  return uniqueSorted
}

function getShipAllDamagedShieldBoxRects(ship) {
  const doc = ship?.doc || null
  if (!doc) return []
  const runtime = ensureShipRuntimeDamageState(ship)
  const damagedByShield = runtime?.damagedShieldBoxIndicesByShield
  if (!damagedByShield || typeof damagedByShield !== "object") return []
  const fallbackPreset = getDamageOverlayPreset()

  const allShieldBoxes = getShipShieldBoxEntriesByNumber(doc)
  const out = []

  for (const [shieldKey, rawIndices] of Object.entries(damagedByShield)) {
    const entries = Array.isArray(allShieldBoxes[String(shieldKey)]) ? allShieldBoxes[String(shieldKey)] : []
    if (entries.length === 0) continue
    const seen = new Set()
    for (const raw of Array.isArray(rawIndices) ? rawIndices : []) {
      const idx = normalizeInteger(raw, -1)
      if (idx < 0 || idx >= entries.length || seen.has(idx)) continue
      seen.add(idx)
      const entry = entries[idx]
      if (!entry?.rect) continue
      out.push({
        shieldNumber: normalizeInteger(shieldKey, 0),
        index: idx,
        rect: entry.rect,
        overlayPreset: getOrSeedDamageOverlayPresetForDamageIndex(
          runtime,
          "damagedShieldBoxOverlayPresetByShield",
          String(shieldKey),
          idx,
          fallbackPreset
        )
      })
    }
  }

  return out
}

function getShipDamagedArmorBoxIndexSet(ship, armorGroupId) {
  const runtime = ensureShipRuntimeDamageState(ship)
  const key = String(armorGroupId || "").trim().toLowerCase()
  const raw = Array.isArray(runtime.damagedArmorBoxIndicesByGroup[key])
    ? runtime.damagedArmorBoxIndicesByGroup[key]
    : []
  const set = new Set()
  for (const item of raw) {
    const idx = normalizeInteger(item, -1)
    if (idx < 0) continue
    set.add(idx)
  }
  return set
}

function setShipDamagedArmorBoxIndices(ship, armorGroupId, indices, options = {}) {
  const runtime = ensureShipRuntimeDamageState(ship)
  const key = String(armorGroupId || "").trim().toLowerCase()
  const uniqueSorted = Array.from(new Set(
    (Array.isArray(indices) ? indices : [])
      .map((n) => normalizeInteger(n, -1))
      .filter((n) => n >= 0)
  )).sort((a, b) => a - b)
  runtime.damagedArmorBoxIndicesByGroup[key] = uniqueSorted
  syncDamageOverlayIndexMapForDamageKey(
    runtime,
    "damagedArmorBoxOverlayPresetByGroup",
    key,
    uniqueSorted,
    { presetForNew: options.presetForNew }
  )
  return uniqueSorted
}

function getShipAllDamagedArmorBoxRects(ship) {
  const doc = ship?.doc || null
  if (!doc) return []

  const runtime = ensureShipRuntimeDamageState(ship)
  const damagedByGroup = runtime?.damagedArmorBoxIndicesByGroup
  if (!damagedByGroup || typeof damagedByGroup !== "object") return []
  const fallbackPreset = getDamageOverlayPreset()

  const armorGroups = getShipArmorBoxGroups(doc)
  const out = []
  for (const [groupId, rawIndices] of Object.entries(damagedByGroup)) {
    const group = armorGroups[String(groupId || "").trim().toLowerCase()]
    const entries = Array.isArray(group?.entries) ? group.entries : []
    if (entries.length === 0) continue

    const seen = new Set()
    for (const raw of Array.isArray(rawIndices) ? rawIndices : []) {
      const idx = normalizeInteger(raw, -1)
      if (idx < 0 || idx >= entries.length || seen.has(idx)) continue
      seen.add(idx)
      const entry = entries[idx]
      if (!entry?.rect) continue
      out.push({
        armorGroupId: String(groupId),
        armorKey: String(group?.key || "Armor"),
        index: idx,
        rect: entry.rect,
        overlayPreset: getOrSeedDamageOverlayPresetForDamageIndex(
          runtime,
          "damagedArmorBoxOverlayPresetByGroup",
          String(groupId || "").trim().toLowerCase(),
          idx,
          fallbackPreset
        )
      })
    }
  }

  return out
}

function getShipSsdBoxEntriesByKey(doc, ssdKey) {
  const key = String(ssdKey || "").trim()
  if (!key) return []
  const entries = Array.isArray(doc?.ssd?.[key]) ? doc.ssd[key] : []
  const normalizedKey = normalizeLabelToken(key)
  const parsed = []
  for (const entry of entries) {
    const rect = parsePosRect(entry?.pos)
    if (!rect) continue
    const rank = Number(entry?.rank)
    parsed.push({
      rect,
      rank: Number.isFinite(rank) ? rank : null,
      sourceOrder: parsed.length,
      type: String(entry?.type || "").trim(),
      designation: String(entry?.designation || "").trim(),
      arc: String(entry?.arc || "").trim()
    })
  }
  const out = RANK_ORDERED_SSD_KEYS.has(normalizedKey)
    ? parsed
        .slice()
        .sort((a, b) => {
          const aHasRank = Number.isFinite(a.rank)
          const bHasRank = Number.isFinite(b.rank)
          if (aHasRank && bHasRank && a.rank !== b.rank) return a.rank - b.rank
          if (aHasRank && !bHasRank) return -1
          if (!aHasRank && bHasRank) return 1
          return a.sourceOrder - b.sourceOrder
        })
        .map((entry, index) => ({
          index,
          rect: entry.rect,
          type: entry.type,
          designation: entry.designation,
          arc: entry.arc
        }))
    : parsed.map((entry, index) => ({
        index,
        rect: entry.rect,
        type: entry.type,
        designation: entry.designation,
        arc: entry.arc
      }))
  return out
}

function getShipSsdKeys(doc) {
  const ssd = doc?.ssd
  if (!ssd || typeof ssd !== "object") return []
  return Object.keys(ssd)
}

function resolveSsdKeysFromSpec(doc, spec) {
  const raw = String(spec || "").trim()
  if (!raw) return []

  const ssdKeys = getShipSsdKeys(doc)
  if (ssdKeys.length === 0) return []

  if (ssdKeys.includes(raw)) return [raw]

  const wanted = normalizeLabelToken(raw)
  if (!wanted) return []

  const exactNormalized = ssdKeys.filter((key) => normalizeLabelToken(key) === wanted)
  if (exactNormalized.length > 0) return exactNormalized

  const containsNormalized = ssdKeys.filter((key) => normalizeLabelToken(key).includes(wanted))
  if (containsNormalized.length > 0) return containsNormalized

  // Fall back to matching by SSD entry type (e.g. "Photon Torpedo", "Plasma Torpedo S")
  // so weapon-type specs are not tied to any bearing/arc text.
  const typeMatchedKeys = []
  const torpLikeSpec = /\btorp/.test(wanted)
  for (const key of ssdKeys) {
    const entries = Array.isArray(doc?.ssd?.[key]) ? doc.ssd[key] : []
    if (entries.length === 0) continue

    let matched = false
    for (const entry of entries) {
      const typeWanted = normalizeLabelToken(entry?.type || "")
      if (!typeWanted) continue
      if (typeWanted === wanted || typeWanted.includes(wanted) || wanted.includes(typeWanted)) {
        matched = true
        break
      }
      if (torpLikeSpec && /\btorp/.test(typeWanted)) {
        matched = true
        break
      }
    }
    if (matched) typeMatchedKeys.push(key)
  }
  if (typeMatchedKeys.length > 0) {
    return Array.from(new Set(typeMatchedKeys))
  }

  return []
}

function getShipDamagedSsdBoxIndexSet(ship, ssdKey) {
  const runtime = ensureShipRuntimeDamageState(ship)
  const key = normalizeLabelToken(ssdKey)
  const raw = Array.isArray(runtime.damagedSsdBoxIndicesByKey[key])
    ? runtime.damagedSsdBoxIndicesByKey[key]
    : []
  const set = new Set()
  for (const item of raw) {
    const idx = normalizeInteger(item, -1)
    if (idx < 0) continue
    set.add(idx)
  }
  return set
}

function setShipDamagedSsdBoxIndices(ship, ssdKey, indices, options = {}) {
  const runtime = ensureShipRuntimeDamageState(ship)
  const key = normalizeLabelToken(ssdKey)
  if (!key) return []
  const uniqueSorted = Array.from(new Set(
    (Array.isArray(indices) ? indices : [])
      .map((n) => normalizeInteger(n, -1))
      .filter((n) => n >= 0)
  )).sort((a, b) => a - b)
  runtime.damagedSsdBoxIndicesByKey[key] = uniqueSorted
  syncDamageOverlayIndexMapForDamageKey(
    runtime,
    "damagedSsdBoxOverlayPresetByKey",
    key,
    uniqueSorted,
    { presetForNew: options.presetForNew }
  )
  return uniqueSorted
}

function getShipAllDamagedInternalBoxRects(ship) {
  const doc = ship?.doc || null
  if (!doc) return []

  const runtime = ensureShipRuntimeDamageState(ship)
  const damagedByKey = runtime?.damagedSsdBoxIndicesByKey
  if (!damagedByKey || typeof damagedByKey !== "object") return []
  const fallbackPreset = getDamageOverlayPreset()

  const out = []
  for (const [normalizedKey, rawIndices] of Object.entries(damagedByKey)) {
    const actualKeys = getShipSsdKeys(doc).filter((key) => normalizeLabelToken(key) === normalizedKey)
    for (const actualKey of actualKeys) {
      const entries = getShipSsdBoxEntriesByKey(doc, actualKey)
      if (entries.length === 0) continue
      const seen = new Set()
      for (const rawIdx of Array.isArray(rawIndices) ? rawIndices : []) {
        const idx = normalizeInteger(rawIdx, -1)
        if (idx < 0 || idx >= entries.length || seen.has(idx)) continue
        seen.add(idx)
        out.push({
          ssdKey: actualKey,
          index: idx,
          rect: entries[idx].rect,
          overlayPreset: getOrSeedDamageOverlayPresetForDamageIndex(
            runtime,
            "damagedSsdBoxOverlayPresetByKey",
            normalizedKey,
            idx,
            fallbackPreset
          )
        })
      }
    }
  }
  return out
}

function getAssociatedSsdSpecsForDacLabel(chartCellLabel) {
  const entry = findRollsLabelEntry(chartCellLabel)
  const fromRolls = Array.isArray(entry?.associatedBoxes)
    ? entry.associatedBoxes.map((item) => String(item || "").trim()).filter(Boolean)
    : []
  if (fromRolls.length > 0) return fromRolls

  const fallback = DAC_LABEL_FALLBACK_SSD_KEYS[normalizeLabelToken(chartCellLabel)]
  return Array.isArray(fallback) ? fallback.slice() : []
}

function formatSelectableWeaponEntryLabel(ssdKey, entry) {
  const groupName = String(ssdKey || "").trim() || "Weapon"
  const type = String(entry?.type || "").trim()
  const designation = String(entry?.designation || "").trim()
  const arc = String(entry?.arc || "").trim()
  const number = normalizeInteger(entry?.index, -1)

  let label = type || groupName
  if (designation) {
    label += ` ${designation}`
  } else if (number >= 0) {
    label += ` #${number + 1}`
  }
  if (arc) {
    label += ` (${arc})`
  }
  return label
}

async function promptForSelectableWeaponDamage(chartCellLabel, ssdKey, availableEntries) {
  const options = Array.isArray(availableEntries) ? availableEntries.filter(Boolean) : []
  if (options.length <= 1) {
    return options[0] || null
  }

  const title = String(chartCellLabel || "").replace(/\r?\n/g, " ").replace(/\s+/g, " ").trim() || "Internal"
  const groupLabel = String(ssdKey || "weapon").trim() || "weapon"
  const allowSkip = normalizeLabelToken(ssdKey) === "phaser"

  if (state.assignDamageWindowOpen && window.api && typeof window.api.setAssignDamageWeaponPrompt === "function") {
    if (state.assignDamageWeaponPromptSyncPromise) {
      try {
        await state.assignDamageWeaponPromptSyncPromise
      } catch {
        // no-op
      }
    }

    if (state.weaponSelectPromise) {
      throw new Error("Weapon selection prompt is already active.")
    }

    const entriesByOptionId = new Map()
    const promptOptions = []
    for (let index = 0; index < options.length; index += 1) {
      const entry = options[index]
      const optionId = String(index)
      entriesByOptionId.set(optionId, entry)
      promptOptions.push({
        optionId,
        label: formatSelectableWeaponEntryLabel(ssdKey, entry),
        detail: `Box ${normalizeInteger(entry?.index, 0) + 1}`
      })
    }

    state.weaponSelectPromise = new Promise((resolve) => {
      state.weaponSelectResolver = resolve
    })

    let shownInAssignDamageWindow = false
    try {
      const res = await window.api.setAssignDamageWeaponPrompt({
        title: `${title} Hit`,
        lead: `Select which ${groupLabel} box is hit.`,
        options: promptOptions,
        allowSkip
      })
      shownInAssignDamageWindow = !!(res && res.ok && res.shown === true)
    } catch {
      shownInAssignDamageWindow = false
    }

    if (shownInAssignDamageWindow) {
      const selectedOptionId = await state.weaponSelectPromise
      if (selectedOptionId === WEAPON_SELECTION_SKIP_TOKEN) {
        return { skipSelection: true }
      }
      return entriesByOptionId.get(String(selectedOptionId || "")) || null
    }

    resolveWeaponSelectionPrompt(null)
  }

  if (
    !state.assignDamageWindowOpen &&
    window.api &&
    typeof window.api.openWeaponSelectWindow === "function" &&
    typeof window.api.setWeaponSelectWindowState === "function"
  ) {
    if (state.weaponSelectWindowClosingPromise) {
      try {
        await state.weaponSelectWindowClosingPromise
      } catch {
        // no-op
      }
    }

    if (state.weaponSelectPromise) {
      throw new Error("Weapon selection prompt is already active.")
    }

    const entriesByOptionId = new Map()
    const promptOptions = []
    for (let index = 0; index < options.length; index += 1) {
      const entry = options[index]
      const optionId = String(index)
      entriesByOptionId.set(optionId, entry)
      promptOptions.push({
        optionId,
        label: formatSelectableWeaponEntryLabel(ssdKey, entry),
        detail: `Box ${normalizeInteger(entry?.index, 0) + 1}`
      })
    }

    state.weaponSelectPromise = new Promise((resolve) => {
      state.weaponSelectResolver = resolve
    })

    let shownInWeaponSelectWindow = false
    try {
      const openRes = await window.api.openWeaponSelectWindow()
      if (openRes && openRes.ok) {
        state.weaponSelectWindowOpen = true
        const stateRes = await window.api.setWeaponSelectWindowState({
          title: `${title} Hit`,
          lead: `Select which ${groupLabel} box is hit.`,
          options: promptOptions,
          allowSkip
        })
        shownInWeaponSelectWindow = !!(stateRes && stateRes.ok && stateRes.shown === true)
      }
    } catch {
      shownInWeaponSelectWindow = false
    }

    if (shownInWeaponSelectWindow) {
      const selectedOptionId = await state.weaponSelectPromise
      if (selectedOptionId === WEAPON_SELECTION_SKIP_TOKEN) {
        return { skipSelection: true }
      }
      return entriesByOptionId.get(String(selectedOptionId || "")) || null
    }

    resolveWeaponSelectionPrompt(null)
  }

  const overlay = byId("weaponSelectOverlay")
  const titleEl = byId("weaponSelectTitle")
  const leadEl = byId("weaponSelectLead")
  const listEl = byId("weaponSelectList")
  if (!overlay || !listEl) {
    return options[0] || null
  }

  if (state.weaponSelectPromise) {
    throw new Error("Weapon selection prompt is already active.")
  }

  if (titleEl) {
    titleEl.textContent = `${title} Hit`
  }
  if (leadEl) {
    leadEl.textContent = `Select which ${groupLabel} box is hit.`
  }

  const skipBtn = byId("btnWeaponSelectSkip")
  if (skipBtn) {
    skipBtn.classList.toggle("hidden", !allowSkip)
  }

  listEl.innerHTML = ""
  for (const entry of options) {
    const button = document.createElement("button")
    button.type = "button"
    button.className = "weaponSelectButton"

    const primary = document.createElement("span")
    primary.textContent = formatSelectableWeaponEntryLabel(ssdKey, entry)
    button.appendChild(primary)

    const detail = document.createElement("small")
    detail.textContent = `Box ${normalizeInteger(entry?.index, 0) + 1}`
    button.appendChild(detail)

    button.onclick = () => {
      resolveWeaponSelectionPrompt(entry)
    }
    listEl.appendChild(button)
  }

  overlay.classList.remove("hidden")
  overlay.setAttribute("aria-hidden", "false")

  try {
    window.focus?.()
  } catch {
    // no-op
  }

  const firstButton = listEl.querySelector("button")
  firstButton?.focus?.()

  state.weaponSelectPromise = new Promise((resolve) => {
    state.weaponSelectResolver = resolve
  })
  return state.weaponSelectPromise.then((selectedValue) => {
    if (selectedValue === WEAPON_SELECTION_SKIP_TOKEN) {
      return { skipSelection: true }
    }
    return selectedValue
  })
}

async function tryMarkFirstAvailableSsdBoxForSpecs(ship, chartCellLabel, specs) {
  const doc = ship?.doc || null
  if (!doc) return null
  const presetForNew = getDamageOverlayPreset()

  const list = Array.isArray(specs) ? specs : []
  for (const spec of list) {
    const actualKeys = resolveSsdKeysFromSpec(doc, spec)
    for (const actualKey of actualKeys) {
      const entries = getShipSsdBoxEntriesByKey(doc, actualKey)
      if (entries.length === 0) continue

      const damagedSet = getShipDamagedSsdBoxIndexSet(ship, actualKey)
      const normalizedActualKey = normalizeLabelToken(actualKey)
      if (USER_SELECTABLE_WEAPON_GROUP_KEYS.has(normalizedActualKey)) {
        const availableEntries = entries.filter((entry) => entry && !damagedSet.has(entry.index))
        if (availableEntries.length === 0) continue

        const chosenEntry = await promptForSelectableWeaponDamage(chartCellLabel, actualKey, availableEntries)
        if (chosenEntry?.skipSelection) {
          // Allow phaser results to be skipped without canceling the rest of the allocation.
          // For Any Weapon, this lets the resolver continue on to heavy/drone groups.
          continue
        }
        if (!chosenEntry) {
          return {
            selectionCanceled: true,
            chartCellLabel: String(chartCellLabel || ""),
            ssdKey: actualKey,
            error: "Weapon selection canceled."
          }
        }

        damagedSet.add(chosenEntry.index)
        setShipDamagedSsdBoxIndices(ship, actualKey, Array.from(damagedSet), { presetForNew })
        return {
          chartCellLabel: String(chartCellLabel || ""),
          ssdKey: actualKey,
          index: chosenEntry.index,
          rect: chosenEntry.rect,
          type: chosenEntry.type,
          designation: chosenEntry.designation,
          arc: chosenEntry.arc
        }
      }

      for (const entry of entries) {
        if (!entry || damagedSet.has(entry.index)) continue
        damagedSet.add(entry.index)
        setShipDamagedSsdBoxIndices(ship, actualKey, Array.from(damagedSet), { presetForNew })
        return {
          chartCellLabel: String(chartCellLabel || ""),
          ssdKey: actualKey,
          index: entry.index,
          rect: entry.rect,
          type: entry.type,
          designation: entry.designation,
          arc: entry.arc
        }
      }
    }
  }

  return null
}

function getDacRowForRoll(rollValue) {
  const wanted = String(clampInteger(rollValue, 2, 12))
  return DAMAGE_ALLOCATION_CHART.rows.find((row) => String(row?.dieRoll || "") === wanted) || null
}

async function promptForManualInternalDacRoll(hitIndex, totalHits) {
  if (
    state.assignDamageWindowOpen &&
    window.api &&
    typeof window.api.setAssignDamageManualDacPrompt === "function"
  ) {
    if (state.manualDacPromptPromise) {
      throw new Error("Manual DAC prompt is already active.")
    }

    state.manualDacPromptPromise = new Promise((resolve) => {
      state.manualDacPromptResolver = resolve
    })

    let shownInAssignDamageWindow = false
    try {
      const stateRes = await window.api.setAssignDamageManualDacPrompt({
        hitIndex: Math.max(1, normalizeInteger(hitIndex, 1)),
        totalHits: Math.max(1, normalizeInteger(totalHits, 1)),
        defaultRoll: 7,
        errorText: ""
      })
      shownInAssignDamageWindow = !!(stateRes && stateRes.ok && stateRes.shown === true)
    } catch {
      shownInAssignDamageWindow = false
    }

    if (shownInAssignDamageWindow) {
      return state.manualDacPromptPromise
    }

    resolveManualDacRollPrompt(null)
    return null
  }

  if (
    window.api &&
    typeof window.api.openManualDacRollWindow === "function" &&
    typeof window.api.setManualDacRollWindowState === "function"
  ) {
    if (state.manualDacPromptPromise) {
      throw new Error("Manual DAC prompt is already active.")
    }

    const openRes = await window.api.openManualDacRollWindow()
    if (!openRes || !openRes.ok) {
      throw new Error(openRes?.error || "Failed to open Manual DAC Roll window.")
    }

    const stateRes = await window.api.setManualDacRollWindowState({
      hitIndex: Math.max(1, normalizeInteger(hitIndex, 1)),
      totalHits: Math.max(1, normalizeInteger(totalHits, 1)),
      defaultRoll: 7,
      errorText: ""
    })
    if (!stateRes || !stateRes.ok) {
      throw new Error(stateRes?.error || "Failed to initialize Manual DAC Roll window.")
    }

    state.manualDacPromptPromise = new Promise((resolve) => {
      state.manualDacPromptResolver = resolve
    })
    return state.manualDacPromptPromise
  }

  if (typeof window.prompt !== "function") {
    throw new Error("Manual DAC prompts are unavailable in this window.")
  }

  while (true) {
    try {
      window.focus?.()
    } catch {
      // no-op
    }

    const promptText = totalHits > 1
      ? `Manual internal damage allocation (${hitIndex}/${totalHits})\nEnter a DAC roll (2-12).`
      : "Manual internal damage allocation\nEnter a DAC roll (2-12)."
    const raw = window.prompt(promptText, "7")
    if (raw == null) return null

    const roll = normalizeInteger(String(raw).trim(), NaN)
    if (Number.isFinite(roll) && roll >= 2 && roll <= 12) return roll
    window.alert("Enter a whole number between 2 and 12.")
  }
}

async function applyManualInternalDacHit(ship, rollValue) {
  if (!ship) return { ok: false, error: "No ship selected." }
  if (isShipDestroyed(ship)) return { ok: false, destroyed: true, error: "Ship already destroyed." }

  const row = getDacRowForRoll(rollValue)
  if (!row) return { ok: false, error: "Invalid DAC roll." }

  for (let colIndex = 0; colIndex < DAMAGE_ALLOCATION_CHART.columns.length; colIndex += 1) {
    const column = DAMAGE_ALLOCATION_CHART.columns[colIndex]
    const chartCellLabel = String(row.cells?.[colIndex] || "")
    const dacCellKey = `${row.dieRoll}:${column}`
    const underlined = isUnderlinedDamageCell(row.dieRoll, column)
    setShipActiveDacCell(ship, dacCellKey)

    if (underlined && isDacCellMarkedOnShip(ship, dacCellKey)) {
      markPassedOverDacCellOnShip(ship, dacCellKey)
      continue
    }

    const specs = getAssociatedSsdSpecsForDacLabel(chartCellLabel)
    const hit = await tryMarkFirstAvailableSsdBoxForSpecs(ship, chartCellLabel, specs)
    if (hit?.selectionCanceled) {
      return {
        ok: false,
        canceled: true,
        selectionCanceled: true,
        roll: normalizeInteger(row.dieRoll, 0),
        column,
        dacCellKey,
        chartCellLabel,
        underlined,
        error: String(hit.error || "Weapon selection canceled.")
      }
    }
    if (!hit) {
      markPassedOverDacCellOnShip(ship, dacCellKey)
      continue
    }

    if (underlined) {
      markDacCellOnShip(ship, dacCellKey)
    }

    return {
      ok: true,
      destroyed: false,
      roll: normalizeInteger(row.dieRoll, 0),
      column,
      dacCellKey,
      chartCellLabel,
      underlined,
      hit
    }
  }

  setShipDestroyed(ship, true)
  return {
    ok: false,
    destroyed: true,
    roll: normalizeInteger(row.dieRoll, 0),
    error: "DAC result moved past column M. Ship destroyed."
  }
}

function formatDacEventLogLine(event, hitIndex = 0) {
  const index = Math.max(0, normalizeInteger(hitIndex, 0)) + 1
  const roll = normalizeInteger(event?.roll, NaN)
  const die1 = normalizeInteger(event?.die1, NaN)
  const die2 = normalizeInteger(event?.die2, NaN)

  let rollText = Number.isFinite(roll) ? `Roll ${roll}` : "Roll ?"
  if (Number.isFinite(die1) && Number.isFinite(die2) && Number.isFinite(roll)) {
    rollText = `Rolled ${die1} + ${die2} = ${roll}`
  }

  if (event?.ok) {
    const type = String(event?.hit?.type || "").trim()
    const designation = String(event?.hit?.designation || "").trim()
    const ssdKey = String(event?.hit?.ssdKey || "").trim()
    const chartCellLabel = String(event?.chartCellLabel || "")
      .replace(/\r?\n/g, " ")
      .replace(/\s+/g, " ")
      .trim()

    let hitLabel = type || ssdKey || chartCellLabel || "Unknown"
    if (designation) {
      hitLabel += ` ${designation}`
    }
    return `${index}. ${rollText} -> ${hitLabel}`
  }

  if (event?.destroyed) {
    const errorText = String(event?.error || "DAC result moved past column M. Ship destroyed.").trim()
    return `${index}. ${rollText} -> ${errorText}`
  }

  const errorText = String(event?.error || "No valid DAC result.").trim()
  return `${index}. ${rollText} -> ${errorText}`
}

function buildDacEventLogLines(events) {
  const list = Array.isArray(events) ? events : []
  return list.map((event, index) => formatDacEventLogLine(event, index)).filter(Boolean)
}

function storeDacResultOnAssignment(assignment, result) {
  if (!assignment || typeof assignment !== "object") return
  assignment.dacEvents = Array.isArray(result?.events) ? result.events.slice() : []
  assignment.dacLogLines = buildDacEventLogLines(assignment.dacEvents)
}

function clearDacResultOnAssignment(assignment) {
  if (!assignment || typeof assignment !== "object") return
  assignment.dacEvents = []
  assignment.dacLogLines = []
}

function showAutomaticDacResultList(assignment, result) {
  void result
  renderAutomaticDacAudit()
}

async function applyManualInternalDacHits(ship, internalHits) {
  const total = Math.max(0, normalizeInteger(internalHits, 0))
  if (!ship || total <= 0) {
    return { attempted: 0, resolved: 0, canceled: false, destroyed: isShipDestroyed(ship) }
  }

  let resolved = 0
  let canceled = false
  let destroyed = false
  const events = []

  for (let i = 1; i <= total; i += 1) {
    if (isShipDestroyed(ship)) {
      destroyed = true
      break
    }

    const roll = await promptForManualInternalDacRoll(i, total)
    if (roll == null) {
      canceled = true
      break
    }

    const result = await applyManualInternalDacHit(ship, roll)
    events.push(result)

    // Keep the chart and SSD overlays in sync as each internal is resolved.
    renderDamageChart()
    renderShipCanvas()
    renderDamageAssignmentSummary()

    if (result?.selectionCanceled) {
      canceled = true
      break
    }
    if (result?.destroyed) {
      destroyed = true
      resolved += 1
      break
    }
    if (result?.ok) {
      resolved += 1
    } else {
      // Treat an unresolved non-destroyed result as consumed input and continue.
      resolved += 1
    }
  }

  const completedAllInternals = !canceled && !destroyed && resolved >= total
  if (completedAllInternals) {
    if (
      state.assignDamageWindowOpen &&
      window.api &&
      typeof window.api.setAssignDamageManualDacPrompt === "function"
    ) {
      try {
        await window.api.setAssignDamageManualDacPrompt({
          complete: true,
          completionText: "Damage complete",
          hitIndex: total,
          totalHits: total,
          defaultRoll: 7,
          errorText: ""
        })
      } catch {
        // no-op
      }
    } else if (window.api && typeof window.api.setManualDacRollWindowState === "function") {
      try {
        await window.api.setManualDacRollWindowState({
          complete: true,
          completionText: "Damage complete",
          hitIndex: total,
          totalHits: total,
          defaultRoll: 7,
          errorText: ""
        })
      } catch {
        // no-op
      }
    }
    queueManualDacChartClearOnWindowClose(ship)
  }

  return {
    attempted: total,
    resolved,
    pending: Math.max(0, total - resolved),
    canceled,
    destroyed: destroyed || isShipDestroyed(ship),
    events
  }
}

async function applyAutomaticInternalDacHits(ship, internalHits) {
  const total = Math.max(0, normalizeInteger(internalHits, 0))
  if (!ship || total <= 0) {
    return { attempted: 0, resolved: 0, pending: 0, canceled: false, destroyed: isShipDestroyed(ship), events: [] }
  }

  clearShipDacChartState(ship)
  renderDamageChart()

  let resolved = 0
  let canceled = false
  let destroyed = false
  const events = []

  for (let i = 1; i <= total; i += 1) {
    if (isShipDestroyed(ship)) {
      destroyed = true
      break
    }

    const rolled = rollTwoSixSidedDice()
    const result = await applyManualInternalDacHit(ship, rolled.total)
    events.push({
      ...result,
      die1: rolled.die1,
      die2: rolled.die2
    })

    if (result?.selectionCanceled) {
      canceled = true
      break
    }

    resolved += 1
    if (result?.destroyed) {
      destroyed = true
      break
    }
  }

  renderDamageChart()
  renderShipCanvas()
  renderDamageAssignmentSummary()

  return {
    attempted: total,
    resolved,
    pending: Math.max(0, total - resolved),
    canceled,
    destroyed: destroyed || isShipDestroyed(ship),
    events
  }
}

function applyShieldDamageToShip(ship, shieldNumber, totalDamage) {
  const doc = ship?.doc || null
  const appliedShieldNumber = clampInteger(shieldNumber, 1, 99)
  const presetForNew = getDamageOverlayPreset()
  const shieldBoxEntries = getShipShieldBoxEntries(doc, appliedShieldNumber)
  const shieldBoxCount = shieldBoxEntries.length
  const shieldPrinted = Math.max(0, getPrintedShieldStrength(doc, appliedShieldNumber))

  const damagedSet = getShipDamagedShieldBoxIndexSet(ship, appliedShieldNumber)
  let shieldBoxesPreviouslyDamaged = 0
  for (const idx of damagedSet) {
    if (idx >= 0 && idx < shieldBoxCount) shieldBoxesPreviouslyDamaged += 1
  }

  let shieldAbsorbed = 0
  const newlyDamagedBoxIndices = []
  for (const entry of shieldBoxEntries) {
    if (shieldAbsorbed >= totalDamage) break
    if (!entry || damagedSet.has(entry.index)) continue
    damagedSet.add(entry.index)
    newlyDamagedBoxIndices.push(entry.index)
    shieldAbsorbed += 1
  }

  const damageAfterShields = Math.max(0, normalizeInteger(totalDamage, 0) - shieldAbsorbed)
  const storedShieldDamage = setShipDamagedShieldBoxIndices(ship, appliedShieldNumber, Array.from(damagedSet), { presetForNew })
  const shieldBoxesDamagedTotal = storedShieldDamage.filter((idx) => idx >= 0 && idx < shieldBoxCount).length
  const shieldRemaining = Math.max(0, shieldBoxCount - shieldBoxesDamagedTotal)

  const armorGroup = getShipArmorBoxGroupForShield(doc, appliedShieldNumber)
  const armorEntries = Array.isArray(armorGroup?.entries) ? armorGroup.entries : []
  const armorBoxCount = armorEntries.length
  const armorGroupId = armorGroup?.id || ""
  const armorGroupLabel = armorGroup?.key || ""
  const armorDamagedSet = armorGroupId ? getShipDamagedArmorBoxIndexSet(ship, armorGroupId) : new Set()

  let armorBoxesPreviouslyDamaged = 0
  for (const idx of armorDamagedSet) {
    if (idx >= 0 && idx < armorBoxCount) armorBoxesPreviouslyDamaged += 1
  }

  let armorAbsorbed = 0
  const newlyDamagedArmorBoxIndices = []
  for (const entry of armorEntries) {
    if (armorAbsorbed >= damageAfterShields) break
    if (!entry || armorDamagedSet.has(entry.index)) continue
    armorDamagedSet.add(entry.index)
    newlyDamagedArmorBoxIndices.push(entry.index)
    armorAbsorbed += 1
  }

  let armorBoxesDamagedTotal = armorBoxesPreviouslyDamaged
  let armorRemaining = Math.max(0, armorBoxCount - armorBoxesDamagedTotal)
  if (armorGroupId) {
    const storedArmorDamage = setShipDamagedArmorBoxIndices(ship, armorGroupId, Array.from(armorDamagedSet), { presetForNew })
    armorBoxesDamagedTotal = storedArmorDamage.filter((idx) => idx >= 0 && idx < armorBoxCount).length
    armorRemaining = Math.max(0, armorBoxCount - armorBoxesDamagedTotal)
  }

  const internals = Math.max(0, damageAfterShields - armorAbsorbed)

  const runtime = ensureShipRuntimeDamageState(ship)
  runtime.internalsTotal = Math.max(0, normalizeInteger(runtime.internalsTotal, 0)) + internals

  return {
    shieldNumber: appliedShieldNumber,
    shieldPrinted,
    shieldBoxCount,
    shieldBoxesPreviouslyDamaged,
    shieldBoxesDamagedThisHit: newlyDamagedBoxIndices.length,
    shieldBoxesDamagedTotal,
    newlyDamagedBoxIndices,
    shieldAbsorbed,
    shieldRemaining,
    armorGroupId,
    armorGroupLabel,
    armorBoxCount,
    armorBoxesPreviouslyDamaged,
    armorBoxesDamagedThisHit: newlyDamagedArmorBoxIndices.length,
    armorBoxesDamagedTotal,
    armorDamagedBoxIndices: newlyDamagedArmorBoxIndices,
    armorAbsorbed,
    armorRemaining,
    internals,
    internalsTotal: runtime.internalsTotal
  }
}

function getShipPhaserEntries(doc) {
  const entries = Array.isArray(doc?.ssd?.phaser) ? doc.ssd.phaser : []
  return entries.map((entry, index) => {
    const type = String(entry?.type || "Phaser").trim() || "Phaser"
    const designation = String(entry?.designation || "").trim()
    const arc = String(entry?.arc || "").trim() || "?"
    const key = `phaser-${index}`
    const shortLabel = designation ? `${type} ${designation}` : `${type} ${index + 1}`
    return {
      key,
      index,
      type,
      designation,
      arc,
      rect: parsePosRect(entry?.pos),
      shortLabel,
      fullLabel: designation ? `${type} ${designation} (${arc})` : `${type} #${index + 1} (${arc})`
    }
  })
}

function getShipPhaserCount(doc) {
  return getShipPhaserEntries(doc).length
}

function getPhasersByKey(doc) {
  const map = new Map()
  for (const phaser of getShipPhaserEntries(doc)) {
    map.set(phaser.key, phaser)
  }
  return map
}

function updateManualDamageControls() {
  const hasShip = !!getActiveShipRecord()
  if (!hasShip) {
    state.manualDamageMode = ""
    state.zoomSectionMode = false
    state.zoomDrag = null
  } else {
    state.zoomSectionMode = state.zoomSectionMode === true
    if (state.zoomSectionMode) {
      state.manualDamageMode = ""
    } else {
      state.manualDamageMode = normalizeManualDamageMode(state.manualDamageMode)
    }
  }
  const mode = state.manualDamageMode
  const zoomMode = hasShip && state.zoomSectionMode === true
  const ship = getActiveShipRecord()
  const hasZoom = hasShip && hasShipZoomCrop(ship)

  const btnMark = byId("btnManualMarkDamage")
  if (btnMark) {
    btnMark.disabled = !hasShip || zoomMode
    btnMark.classList.toggle("toggleButtonActive", mode === "mark")
    btnMark.title = hasShip
      ? (zoomMode
        ? "Turn off zoom mode to manually mark damage."
        : (mode === "mark" ? "Manual mark mode is active. Click again to turn it off." : "Turn on manual mark mode"))
      : "Load a ship first"
  }

  const btnRemove = byId("btnManualRemoveDamage")
  if (btnRemove) {
    btnRemove.disabled = !hasShip || zoomMode
    btnRemove.classList.toggle("toggleButtonActive", mode === "remove")
    btnRemove.title = hasShip
      ? (zoomMode
        ? "Turn off zoom mode to manually remove damage."
        : (mode === "remove" ? "Manual remove mode is active. Click again to turn it off." : "Turn on manual remove mode"))
      : "Load a ship first"
  }

  const btnZoomSection = byId("btnZoomSection")
  if (btnZoomSection) {
    btnZoomSection.disabled = !hasShip
    btnZoomSection.classList.toggle("toggleButtonActive", zoomMode)
    btnZoomSection.textContent = zoomMode ? "Cancel Zoom" : "Zoom Section"
    btnZoomSection.title = hasShip
      ? (zoomMode ? "Drag a rectangle on the SSD to zoom, or click to cancel." : "Drag a rectangle on the SSD to zoom to that section.")
      : "Load a ship first"
  }

  const btnResetZoom = byId("btnResetZoom")
  if (btnResetZoom) {
    btnResetZoom.disabled = !hasShip || !hasZoom
    btnResetZoom.title = hasShip
      ? (hasZoom ? "Reset the SSD view back to full image." : "No active zoom to reset.")
      : "Load a ship first"
  }

  const hint = byId("manualDamageHint")
  if (hint) {
    hint.textContent = !hasShip
      ? "Load a ship to manually edit damage"
      : (zoomMode
        ? "Zoom mode: drag a box to zoom in on that SSD section"
        : (mode === "mark"
        ? "Mark mode: click a box to add damage"
        : (mode === "remove" ? "Remove mode: click a damaged box to clear it" : (hasZoom ? "Manual edit off (zoomed view)" : "Manual edit off"))))
  }

  const canvas = byId("shipCanvas")
  if (canvas) {
    canvas.classList.toggle("manualDamageMarkMode", mode === "mark")
    canvas.classList.toggle("manualDamageRemoveMode", mode === "remove")
    canvas.classList.toggle("zoomSectionMode", zoomMode)
  }
}

function setManualDamageMode(mode) {
  const wanted = normalizeManualDamageMode(mode)
  const hasShip = !!getActiveShipRecord()
  if (!hasShip) {
    state.manualDamageMode = ""
    updateManualDamageControls()
    return
  }

  const hadDrag = !!state.zoomDrag
  if (wanted) {
    state.zoomSectionMode = false
    state.zoomDrag = null
  }
  state.manualDamageMode = state.manualDamageMode === wanted ? "" : wanted
  updateManualDamageControls()
  if (hadDrag) renderShipCanvas()
}

function setZoomSectionMode(enabled, options = {}) {
  const hasShip = !!getActiveShipRecord()
  const quiet = options && options.quiet === true
  const next = hasShip && enabled === true
  const changed = state.zoomSectionMode !== next
  const hadDrag = !!state.zoomDrag

  state.zoomSectionMode = next
  if (next) {
    state.manualDamageMode = ""
  } else {
    state.zoomDrag = null
  }

  updateManualDamageControls()
  if (hadDrag || changed) {
    renderShipCanvas()
  }

  if (!quiet && changed) {
    if (next) {
      setStatus("Zoom mode active. Drag a rectangle on the SSD to zoom.")
    } else {
      setStatus("Zoom mode canceled.")
    }
  }
}

function toggleZoomSectionMode() {
  setZoomSectionMode(!(state.zoomSectionMode === true))
}

function resetActiveShipZoom(options = {}) {
  const quiet = options && options.quiet === true
  const ship = getActiveShipRecord()
  if (!ship) {
    state.zoomDrag = null
    state.zoomSectionMode = false
    updateManualDamageControls()
    renderShipCanvas()
    return false
  }

  const hadZoom = hasShipZoomCrop(ship)
  clearShipZoomCrop(ship)
  state.zoomDrag = null
  state.zoomSectionMode = false
  updateManualDamageControls()
  renderShipCanvas()

  if (!quiet) {
    setStatus(hadZoom ? "Zoom reset to full SSD view." : "SSD view is already at full zoom.")
  }
  return hadZoom
}

function getCanvasPointFromMouseEvent(event, rect) {
  return {
    x: Number(event?.clientX) - rect.left,
    y: Number(event?.clientY) - rect.top
  }
}

function beginZoomSectionDrag(event) {
  if (state.zoomSectionMode !== true) return
  if (event?.button !== 0) return

  const context = getActiveShipCanvasInteractionContext()
  if (!context) return

  const point = getCanvasPointFromMouseEvent(event, context.rect)
  if (!Number.isFinite(point.x) || !Number.isFinite(point.y)) return
  if (!isCanvasPointInsideFit(point.x, point.y, context.fit)) return

  const clamped = clampCanvasPointToFit(point.x, point.y, context.fit)
  state.zoomDrag = {
    startX: clamped.x,
    startY: clamped.y,
    currentX: clamped.x,
    currentY: clamped.y
  }
  renderShipCanvas()
  event.preventDefault?.()
}

function updateZoomSectionDrag(event) {
  if (state.zoomSectionMode !== true) return
  const drag = state.zoomDrag
  if (!drag) return

  const context = getActiveShipCanvasInteractionContext()
  if (!context) return

  const point = getCanvasPointFromMouseEvent(event, context.rect)
  if (!Number.isFinite(point.x) || !Number.isFinite(point.y)) return
  const clamped = clampCanvasPointToFit(point.x, point.y, context.fit)
  drag.currentX = clamped.x
  drag.currentY = clamped.y
  renderShipCanvas()
  event.preventDefault?.()
}

function finalizeZoomSectionDrag(event) {
  if (state.zoomSectionMode !== true) return
  const drag = state.zoomDrag
  if (!drag) return

  const context = getActiveShipCanvasInteractionContext()
  state.zoomDrag = null

  if (!context) {
    renderShipCanvas()
    return
  }

  const point = getCanvasPointFromMouseEvent(event, context.rect)
  if (Number.isFinite(point.x) && Number.isFinite(point.y)) {
    const clamped = clampCanvasPointToFit(point.x, point.y, context.fit)
    drag.currentX = clamped.x
    drag.currentY = clamped.y
  }

  const selectedImageRect = getZoomSelectionImageRectFromDrag(drag, context.crop, context.fit)
  if (!selectedImageRect) {
    renderShipCanvas()
    setStatus("Zoom selection was too small. Drag a larger rectangle.")
    return
  }

  const previousZoom = getShipZoomCrop(context.ship)
  const nextZoom = setShipZoomCrop(context.ship, selectedImageRect)
  state.zoomSectionMode = false
  updateManualDamageControls()
  renderShipCanvas()

  if (!nextZoom) {
    setStatus("Zoom selection matched the full SSD view.")
    return
  }

  if (areCropRectsEqual(previousZoom, nextZoom)) {
    setStatus("Zoom updated.")
    return
  }

  setStatus(`Zoomed to x=${nextZoom.x}, y=${nextZoom.y}, w=${nextZoom.w}, h=${nextZoom.h}.`)
}

function updateGamePersistenceButtons() {
  const hasShips = state.ships.length > 0

  const btnSaveGame = byId("btnSaveGame")
  if (btnSaveGame) {
    btnSaveGame.disabled = !hasShips
    btnSaveGame.title = hasShips
      ? "Save all loaded ships, damage, colors, and current StarHold state"
      : "Load at least one ship before saving a game"
  }

  const btnLoadGame = byId("btnLoadGame")
  if (btnLoadGame) {
    btnLoadGame.disabled = false
    btnLoadGame.title = hasShips
      ? "Load a saved StarHold game and replace the currently loaded ships"
      : "Load a saved StarHold game"
  }
}

function updateAssignDamageButton() {
  const btn = byId("btnAssignDamage")
  const hasShip = !!getActiveShipRecord()
  if (btn) {
    btn.disabled = !hasShip
    btn.title = hasShip ? "Assign damage to the currently selected ship" : "Load a ship first"
  }

  const btnCloseShip = byId("btnCloseShip")
  if (btnCloseShip) {
    btnCloseShip.disabled = !hasShip
    btnCloseShip.title = hasShip ? "Close the currently selected ship" : "No ship loaded"
    btnCloseShip.setAttribute("aria-label", hasShip ? "Close current ship" : "No ship loaded")
  }

  updateGamePersistenceButtons()
  updateManualDamageControls()
  updateRotateShipsButton()
}

function updateRotateShipsButton() {
  const btn = byId("btnRotateShips")
  if (!btn) return

  const hasMultipleShips = state.ships.length > 1
  if (!hasMultipleShips && state.shipRotationTimerId != null) {
    clearInterval(state.shipRotationTimerId)
    state.shipRotationTimerId = null
  }

  const active = state.shipRotationTimerId != null
  btn.disabled = !hasMultipleShips
  btn.classList.toggle("toggleButtonActive", active)
  btn.textContent = active ? "Stop Rotation" : "Rotate SSDs"

  if (!hasMultipleShips) {
    btn.title = "Load at least two ships to rotate through them"
    return
  }

  btn.title = active
    ? "Stop automatically switching between loaded ships"
    : `Cycle through loaded ships every ${Math.round(SHIP_ROTATION_INTERVAL_MS / 1000)} seconds`
}

function stopShipRotation(options = {}) {
  const quiet = options && options.quiet === true
  if (state.shipRotationTimerId != null) {
    clearInterval(state.shipRotationTimerId)
    state.shipRotationTimerId = null
  }
  updateRotateShipsButton()
  if (!quiet) {
    setStatus("Automatic SSD rotation stopped.")
  }
}

function advanceToNextLoadedShip() {
  if (state.ships.length <= 1) return false
  const currentIndex = Number.isInteger(state.activeShipIndex) ? state.activeShipIndex : -1
  const nextIndex = currentIndex >= 0
    ? (currentIndex + 1) % state.ships.length
    : 0
  return activateLoadedShip(nextIndex, { quiet: true })
}

function startShipRotation() {
  if (state.ships.length <= 1) {
    updateRotateShipsButton()
    setStatus("Load at least two ships to rotate through them.")
    return
  }

  if (state.shipRotationTimerId != null) {
    stopShipRotation({ quiet: true })
  }

  state.shipRotationTimerId = window.setInterval(() => {
    if (state.ships.length <= 1) {
      stopShipRotation({ quiet: true })
      return
    }
    advanceToNextLoadedShip()
  }, SHIP_ROTATION_INTERVAL_MS)

  updateRotateShipsButton()
  setStatus(`Automatic SSD rotation started. Switching ships every ${Math.round(SHIP_ROTATION_INTERVAL_MS / 1000)} seconds.`)
}

function toggleShipRotation() {
  if (state.shipRotationTimerId != null) {
    stopShipRotation()
    return
  }
  startShipRotation()
}

function setShipLabel(text) {
  const el = byId("shipLabel")
  if (el) el.textContent = String(text || "")
}

function getShipMarkerTitleText(ship) {
  const colors = getShipDuplicateMarkerColors(ship)
  if (!Array.isArray(colors) || colors.length === 0) return "no marker"

  const names = colors.map((raw) => {
    const key = String(raw || "").trim().toLowerCase()
    if (!key) return ""
    return DUPLICATE_SHIP_MARKER_COLOR_NAMES[key] || key
  }).filter(Boolean)

  if (names.length === 0) return "no marker"
  if (names.length === 1) return `${names[0]} marker`
  return `${names.join(" + ")} markers`
}

function updateWindowTitle(ship = getActiveShipRecord()) {
  const appName = "starhold"
  if (!ship) {
    document.title = appName
    return
  }

  const shipName = String(ship.shipLabel || ship.baseName || ship.jsonFile || "Ship").trim() || "Ship"
  const markerText = getShipMarkerTitleText(ship)
  document.title = `${appName} - ${shipName} (${markerText})`
}

function applyShipRecordToState(ship) {
  if (!ship) {
    state.doc = null
    state.image = null
    state.imageDataUrl = null
    state.imagePath = ""
    state.jsonPath = ""
    state.shipLabel = ""
    state.crop = null
    updateWindowTitle(null)
    return
  }

  state.doc = ship.doc || null
  state.image = ship.image || null
  state.imageDataUrl = ship.imageDataUrl || null
  state.imagePath = ship.imagePath || ""
  state.jsonPath = ship.jsonPath || ""
  state.shipLabel = ship.shipLabel || ""
  state.crop = ship.crop || null
  updateWindowTitle(ship)
}

function areCropRectsEqual(a, b) {
  if (!a || !b) return false
  return (
    normalizeInteger(a.x, -1) === normalizeInteger(b.x, -2) &&
    normalizeInteger(a.y, -1) === normalizeInteger(b.y, -2) &&
    normalizeInteger(a.w, -1) === normalizeInteger(b.w, -2) &&
    normalizeInteger(a.h, -1) === normalizeInteger(b.h, -2)
  )
}

function normalizeShipZoomCrop(ship, rawCrop) {
  const base = ship?.crop
  if (!ship || !base || !rawCrop || typeof rawCrop !== "object") return null

  const baseX1 = Number(base.x)
  const baseY1 = Number(base.y)
  const baseX2 = baseX1 + Number(base.w)
  const baseY2 = baseY1 + Number(base.h)
  if (!Number.isFinite(baseX1) || !Number.isFinite(baseY1) || !Number.isFinite(baseX2) || !Number.isFinite(baseY2)) {
    return null
  }
  if (baseX2 <= baseX1 || baseY2 <= baseY1) return null

  const rawX = Number(rawCrop.x)
  const rawY = Number(rawCrop.y)
  const rawW = Number(rawCrop.w)
  const rawH = Number(rawCrop.h)
  if (!Number.isFinite(rawX) || !Number.isFinite(rawY) || !Number.isFinite(rawW) || !Number.isFinite(rawH)) return null
  if (rawW <= 0 || rawH <= 0) return null

  const x1 = Math.max(baseX1, Math.min(baseX2 - 1, Math.floor(rawX)))
  const y1 = Math.max(baseY1, Math.min(baseY2 - 1, Math.floor(rawY)))
  const x2 = Math.min(baseX2, Math.max(x1 + 1, Math.ceil(rawX + rawW)))
  const y2 = Math.min(baseY2, Math.max(y1 + 1, Math.ceil(rawY + rawH)))

  const normalized = {
    x: x1,
    y: y1,
    w: Math.max(1, x2 - x1),
    h: Math.max(1, y2 - y1),
    source: "zoom"
  }

  if (areCropRectsEqual(normalized, base)) return null
  return normalized
}

function getShipZoomCrop(ship) {
  const normalized = normalizeShipZoomCrop(ship, ship?.zoomCrop)
  if (!ship || typeof ship !== "object") return normalized
  ship.zoomCrop = normalized
    ? { x: normalized.x, y: normalized.y, w: normalized.w, h: normalized.h }
    : null
  return normalized
}

function setShipZoomCrop(ship, rawCrop) {
  if (!ship || typeof ship !== "object") return null
  const normalized = normalizeShipZoomCrop(ship, rawCrop)
  ship.zoomCrop = normalized
    ? { x: normalized.x, y: normalized.y, w: normalized.w, h: normalized.h }
    : null
  return ship.zoomCrop
}

function clearShipZoomCrop(ship) {
  if (!ship || typeof ship !== "object") return false
  const had = !!ship.zoomCrop
  ship.zoomCrop = null
  return had
}

function hasShipZoomCrop(ship) {
  return !!getShipZoomCrop(ship)
}

function getShipCanvasSourceCrop(ship = getActiveShipRecord()) {
  if (!ship) return state.crop || null
  return getShipZoomCrop(ship) || ship.crop || null
}

function getActiveShipCanvasInteractionContext() {
  const ship = getActiveShipRecord()
  const canvas = byId("shipCanvas")
  const crop = getShipCanvasSourceCrop(ship)
  if (!ship || !canvas || !crop) return null

  const rect = canvas.getBoundingClientRect()
  if (rect.width <= 0 || rect.height <= 0) return null

  const fit = fitRectIntoCanvas(crop.w, crop.h, rect.width, rect.height, 4)
  return { ship, canvas, crop, rect, fit }
}

function isCanvasPointInsideFit(canvasX, canvasY, fit) {
  if (!fit) return false
  const x1 = Number(fit.x)
  const y1 = Number(fit.y)
  const x2 = x1 + Number(fit.w)
  const y2 = y1 + Number(fit.h)
  return canvasX >= x1 && canvasY >= y1 && canvasX <= x2 && canvasY <= y2
}

function clampCanvasPointToFit(canvasX, canvasY, fit) {
  if (!fit) return { x: canvasX, y: canvasY }
  const x1 = Number(fit.x)
  const y1 = Number(fit.y)
  const x2 = x1 + Number(fit.w)
  const y2 = y1 + Number(fit.h)
  return {
    x: Math.min(x2, Math.max(x1, Number(canvasX))),
    y: Math.min(y2, Math.max(y1, Number(canvasY)))
  }
}

function canvasPointToImagePoint(canvasX, canvasY, crop, fit) {
  if (!crop || !fit) return null
  const fitW = Math.max(1, Number(fit.w))
  const fitH = Math.max(1, Number(fit.h))
  const u = (Number(canvasX) - Number(fit.x)) / fitW
  const v = (Number(canvasY) - Number(fit.y)) / fitH
  const clampedU = Math.min(1, Math.max(0, u))
  const clampedV = Math.min(1, Math.max(0, v))
  return {
    x: Number(crop.x) + clampedU * Number(crop.w),
    y: Number(crop.y) + clampedV * Number(crop.h)
  }
}

function getZoomSelectionImageRectFromDrag(drag, crop, fit) {
  if (!drag || !crop || !fit) return null

  const rawX1 = Math.min(Number(drag.startX), Number(drag.currentX))
  const rawY1 = Math.min(Number(drag.startY), Number(drag.currentY))
  const rawX2 = Math.max(Number(drag.startX), Number(drag.currentX))
  const rawY2 = Math.max(Number(drag.startY), Number(drag.currentY))
  if (!Number.isFinite(rawX1) || !Number.isFinite(rawY1) || !Number.isFinite(rawX2) || !Number.isFinite(rawY2)) {
    return null
  }

  const p1 = clampCanvasPointToFit(rawX1, rawY1, fit)
  const p2 = clampCanvasPointToFit(rawX2, rawY2, fit)
  const width = Math.max(0, p2.x - p1.x)
  const height = Math.max(0, p2.y - p1.y)
  if (width < 8 || height < 8) return null

  const imageP1 = canvasPointToImagePoint(p1.x, p1.y, crop, fit)
  const imageP2 = canvasPointToImagePoint(p2.x, p2.y, crop, fit)
  if (!imageP1 || !imageP2) return null

  return {
    x: Math.min(imageP1.x, imageP2.x),
    y: Math.min(imageP1.y, imageP2.y),
    w: Math.max(1, Math.abs(imageP2.x - imageP1.x)),
    h: Math.max(1, Math.abs(imageP2.y - imageP1.y))
  }
}

function getLoadedShipOptionLabel(ship, index) {
  const primary = String(ship?.shipLabel || ship?.baseName || ship?.jsonFile || `Ship ${index + 1}`).trim()
  const fileName = String(ship?.jsonFile || "").trim()
  if (fileName && fileName !== primary) return `${index + 1}. ${primary} (${fileName})`
  return `${index + 1}. ${primary}`
}

function getShipDuplicateMarkerMatchKey(ship) {
  if (!ship || typeof ship !== "object") return ""

  const jsonPath = String(ship.jsonPath || "").trim()
  if (jsonPath) return `json:${jsonPath.toLowerCase()}`

  const folderPath = String(ship.folderPath || "").trim()
  const jsonFile = String(ship.jsonFile || "").trim()
  if (folderPath || jsonFile) {
    return `file:${folderPath.toLowerCase()}::${jsonFile.toLowerCase()}`
  }

  const label = String(ship.shipLabel || ship.baseName || "").trim()
  return label ? `label:${label.toLowerCase()}` : ""
}

function getShipDuplicateInstanceOrdinal(ship) {
  const matchKey = getShipDuplicateMarkerMatchKey(ship)
  if (!matchKey) return -1

  let ordinal = 0
  for (const candidate of state.ships) {
    if (!candidate || getShipDuplicateMarkerMatchKey(candidate) !== matchKey) continue
    if (candidate === ship) return ordinal
    ordinal += 1
  }

  return -1
}

function getShipDuplicateMarkerColors(ship) {
  const ordinal = getShipDuplicateInstanceOrdinal(ship)
  if (ordinal <= 0) return []

  if (ordinal <= DUPLICATE_SHIP_MARKER_COLORS.length) {
    return [DUPLICATE_SHIP_MARKER_COLORS[ordinal - 1]]
  }

  if (DUPLICATE_SHIP_MARKER_PAIR_COLORS.length === 0) return []
  const pairIndex = ordinal - DUPLICATE_SHIP_MARKER_COLORS.length - 1
  return DUPLICATE_SHIP_MARKER_PAIR_COLORS[pairIndex % DUPLICATE_SHIP_MARKER_PAIR_COLORS.length]
}

function drawDuplicateShipMarkers(ctx, ship, fit) {
  const colors = getShipDuplicateMarkerColors(ship)
  if (!ship || !fit || colors.length === 0) return

  const radius = Math.max(7, Math.min(12, Math.round(Math.min(fit.w, fit.h) * 0.014)))
  const margin = Math.max(10, Math.round(radius * 0.75))
  const gap = Math.max(4, Math.round(radius * 0.45))
  const centerY = fit.y + margin + radius

  ctx.save()
  ctx.shadowColor = "rgba(0, 0, 0, 0.4)"
  ctx.shadowBlur = 10

  for (let i = 0; i < colors.length; i += 1) {
    const centerX = fit.x + margin + radius + (i * ((radius * 2) + gap))
    ctx.beginPath()
    ctx.arc(centerX, centerY, radius, 0, Math.PI * 2)
    ctx.fillStyle = colors[i]
    ctx.fill()
    ctx.lineWidth = 2
    ctx.strokeStyle = "rgba(255, 255, 255, 0.92)"
    ctx.stroke()
  }

  ctx.restore()
}

function renderShipSwitcher() {
  const select = byId("shipSelect")
  const count = byId("shipCount")

  if (count) {
    const total = state.ships.length
    count.textContent = `${total} loaded`
  }

  if (!select) return

  select.innerHTML = ""
  if (state.ships.length === 0) {
    const option = document.createElement("option")
    option.value = ""
    option.textContent = "No ships loaded"
    select.appendChild(option)
    select.disabled = true
    updateAssignDamageButton()
    return
  }

  for (let i = 0; i < state.ships.length; i += 1) {
    const ship = state.ships[i]
    const option = document.createElement("option")
    option.value = String(i)
    option.textContent = getLoadedShipOptionLabel(ship, i)
    select.appendChild(option)
  }

  select.disabled = false
  const safeIndex = state.activeShipIndex >= 0 && state.activeShipIndex < state.ships.length ? state.activeShipIndex : 0
  select.value = String(safeIndex)
  updateAssignDamageButton()
}

function formatShipStatusLine(ship, index, total, prefix = "Showing") {
  const label = String(ship?.shipLabel || ship?.jsonFile || `Ship ${index + 1}`).trim()
  const jsonFile = String(ship?.jsonFile || "").trim()
  const imageFile = String(ship?.imageFile || "").trim()
  const filePart = [jsonFile, imageFile].filter(Boolean).join(" + ")
  return filePart
    ? `${prefix} ship ${index + 1}/${total}: ${label}\n${filePart}`
    : `${prefix} ship ${index + 1}/${total}: ${label}`
}

function activateLoadedShip(index, options = {}) {
  const nextIndex = Number(index)
  const quiet = options && options.quiet === true
  if (!Number.isInteger(nextIndex) || nextIndex < 0 || nextIndex >= state.ships.length) return false

  state.activeShipIndex = nextIndex
  const ship = state.ships[nextIndex]
  applyShipRecordToState(ship)
  renderShipSwitcher()
  setShipLabel(state.shipLabel)
  renderShipInfo()
  renderDamageChart()
  renderDamageAssignmentSummary()
  renderShipCanvas()
  syncAssignDamageWindowStateForActiveShip()

  if (!quiet) {
    setStatus(formatShipStatusLine(ship, nextIndex, state.ships.length))
  }
  return true
}

function closeActiveShip() {
  const idx = Number(state.activeShipIndex)
  if (!Number.isInteger(idx) || idx < 0 || idx >= state.ships.length) {
    setStatus("No loaded ship is selected.")
    return false
  }

  const closingShip = state.ships[idx]
  const closingLabel = String(closingShip?.shipLabel || closingShip?.jsonFile || `Ship ${idx + 1}`).trim()
  state.ships.splice(idx, 1)

  if (state.ships.length === 0) {
    state.activeShipIndex = -1
    applyShipRecordToState(null)
    renderShipSwitcher()
    setShipLabel("")
    renderShipInfo()
    renderDamageChart()
    renderDamageAssignmentSummary()
    renderShipCanvas()
    syncAssignDamageWindowStateForActiveShip()
    setStatus(`Closed ship: ${closingLabel}. No ships loaded.`)
    return true
  }

  const nextIndex = Math.min(idx, state.ships.length - 1)
  activateLoadedShip(nextIndex, { quiet: true })
  const nextShip = state.ships[nextIndex]
  setStatus(
    `Closed ship: ${closingLabel}.\n` +
    formatShipStatusLine(nextShip, nextIndex, state.ships.length)
  )
  return true
}

function buildAssignDamageShieldOptions(doc) {
  const shieldValues = getShipShieldValues(doc)
  const entries = Object.entries(shieldValues)
    .map(([num, value]) => ({ number: normalizeInteger(num, 0), value }))
    .filter((entry) => entry.number > 0)
    .sort((a, b) => a.number - b.number)

  if (entries.length > 0) return entries

  const fallback = []
  for (let i = 1; i <= 6; i += 1) {
    fallback.push({ number: i, value: 0 })
  }
  return fallback
}

function populateAssignDamageShieldSelect(doc, selectedShieldNumber = 1) {
  const select = byId("assignShieldNumber")
  if (!select) return
  select.innerHTML = ""

  const options = buildAssignDamageShieldOptions(doc)
  for (const entry of options) {
    const option = document.createElement("option")
    option.value = String(entry.number)
    option.textContent = entry.value > 0
      ? `Shield #${entry.number} (${entry.value})`
      : `Shield #${entry.number}`
    select.appendChild(option)
  }

  const desired = String(normalizeInteger(selectedShieldNumber, 1))
  const hasDesired = options.some((entry) => String(entry.number) === desired)
  select.value = hasDesired ? desired : String(options[0]?.number || 1)
}

function renderAssignDamagePhaserList(doc, selectedKeys = []) {
  const container = byId("assignNonBearingPhasersList")
  const meta = byId("assignPhaserListMeta")
  if (!container) return

  const selectedSet = new Set(Array.isArray(selectedKeys) ? selectedKeys.map((k) => String(k)) : [])
  const phasers = getShipPhaserEntries(doc)
  container.innerHTML = ""

  if (meta) {
    meta.textContent = phasers.length > 0
      ? `Select any bearing phasers (${phasers.length} total).`
      : "No phasers found in the active ship JSON."
  }

  if (phasers.length === 0) {
    const empty = document.createElement("div")
    empty.className = "phaserChoiceEmpty"
    empty.textContent = "No phaser entries were found in this ship's SSD data."
    container.appendChild(empty)
    return
  }

  for (const phaser of phasers) {
    const row = document.createElement("label")
    row.className = "phaserChoiceItem"

    const input = document.createElement("input")
    input.type = "checkbox"
    input.value = phaser.key
    input.checked = selectedSet.has(phaser.key)

    const text = document.createElement("div")
    text.className = "phaserChoiceLabel"
    const strong = document.createElement("strong")
    strong.textContent = phaser.designation
      ? `${phaser.type} ${phaser.designation}`
      : `${phaser.type} #${phaser.index + 1}`
    const sub = document.createElement("span")
    sub.textContent = `Arc: ${phaser.arc}`
    text.appendChild(strong)
    text.appendChild(sub)

    row.appendChild(input)
    row.appendChild(text)
    container.appendChild(row)
  }
}

function getSelectedAssignDamagePhasers(doc) {
  const container = byId("assignNonBearingPhasersList")
  const phasers = getShipPhaserEntries(doc)
  const byKey = new Map(phasers.map((phaser) => [phaser.key, phaser]))
  if (!container) return []

  const checked = container.querySelectorAll('input[type="checkbox"]:checked')
  const out = []
  for (const el of checked) {
    const key = String(el.value || "")
    const phaser = byKey.get(key)
    if (!phaser) continue
    out.push({
      key: phaser.key,
      label: phaser.fullLabel
    })
  }
  return out
}

function getActiveShipDamageAssignment() {
  const ship = getActiveShipRecord()
  if (!ship || !ship.damageAssignment || typeof ship.damageAssignment !== "object") return null
  return ship.damageAssignment
}

function getShipDamageAssignmentPreview(ship) {
  if (!ship || !ship.damageAssignmentPreview || typeof ship.damageAssignmentPreview !== "object") return null
  return ship.damageAssignmentPreview
}

function buildAssignDamageWindowStateForActiveShip() {
  const ship = getActiveShipRecord()
  if (!ship) {
    return {
      available: false,
      shipIndex: -1,
      shipLabel: "No active ship",
      shieldOptions: [],
      selectedShieldNumber: 1,
      totalDamage: "",
      hintText: "Load and select a ship in the main StarHold window first."
    }
  }

  const prior = ship.damageAssignment && typeof ship.damageAssignment === "object" ? ship.damageAssignment : null

  return {
    available: true,
    shipIndex: state.activeShipIndex,
    shipLabel: String(ship.shipLabel || ship.jsonFile || "Ship"),
    jsonFile: String(ship.jsonFile || ""),
    shieldOptions: buildAssignDamageShieldOptions(ship.doc),
    selectedShieldNumber: prior?.shieldNumber || 1,
    selectedAssignmentMode: normalizeDamageAssignmentMode(prior?.mode),
    totalDamage: Number.isFinite(Number(prior?.totalDamage)) ? String(prior.totalDamage) : "",
    hintText: `Marks shield boxes on the SSD in-memory (${getDamageOverlayLabelLower()} overlay), then armor boxes if present, and saves remaining overflow as internals.`
  }
}

async function syncAssignDamageWindowStateForActiveShip() {
  if (!window.api || typeof window.api.setAssignDamageWindowState !== "function") return
  try {
    await window.api.setAssignDamageWindowState(buildAssignDamageWindowStateForActiveShip())
  } catch {
    // If the window is closed or IPC is unavailable, ignore.
  }
}

function renderAutomaticDacAudit() {
  const root = byId("automaticDacAudit")
  if (!root) return

  const ship = getActiveShipRecord()
  const assignment = getActiveShipDamageAssignment()
  const mode = normalizeDamageAssignmentMode(assignment?.mode)
  const lines = Array.isArray(assignment?.dacLogLines) ? assignment.dacLogLines : []

  if (!ship || !assignment || mode !== "automatic" || lines.length === 0) {
    root.innerHTML = ""
    root.classList.add("hidden")
    return
  }

  root.innerHTML = ""
  root.classList.remove("hidden")

  const heading = document.createElement("h3")
  heading.textContent = "Automatic DAC Audit"
  root.appendChild(heading)

  const note = document.createElement("div")
  note.className = "dacAuditNote"
  note.textContent = String(ship.shipLabel || ship.jsonFile || "Ship")
  root.appendChild(note)

  const list = document.createElement("div")
  list.className = "dacAuditList"
  for (const line of lines) {
    const item = document.createElement("div")
    item.className = "dacAuditLine"
    item.textContent = line
    list.appendChild(item)
  }
  root.appendChild(list)
}

function renderDamageAssignmentSummary() {
  renderAutomaticDacAudit()

  const root = byId("damageAssignmentSummary")
  if (!root) return

  const ship = getActiveShipRecord()
  const assignment = getActiveShipDamageAssignment()
  if (!ship || !assignment) {
    root.innerHTML = ""
    root.classList.add("hidden")
    return
  }

  root.innerHTML = ""
  root.classList.remove("hidden")

  const heading = document.createElement("h3")
  heading.textContent = "Damage Assignment"
  root.appendChild(heading)

  const grid = document.createElement("div")
  grid.className = "damageAssignGrid"

  const rows = [
    ["Ship", String(ship.shipLabel || ship.jsonFile || "Ship")],
    ["Mode", String(assignment.modeLabel || getDamageAssignmentModeLabel(assignment.mode))],
    ["Total Damage", String(assignment.totalDamage)],
    ["Shield Hit", `Shield #${assignment.shieldNumber} (printed ${assignment.shieldPrinted}, ${Number.isFinite(Number(assignment.shieldBoxCount)) ? assignment.shieldBoxCount : "?"} boxes)`],
    ["Shield Absorbed", `${assignment.shieldAbsorbed} box${Number(assignment.shieldAbsorbed) === 1 ? "" : "es"}`],
    ["Shield Remaining", `${assignment.shieldRemaining} box${Number(assignment.shieldRemaining) === 1 ? "" : "es"}`],
  ]

  if (Number(assignment.armorBoxCount) > 0 || Number(assignment.armorAbsorbed) > 0) {
    rows.push(
      ["Armor", `${String(assignment.armorGroupLabel || "Armor")} (${assignment.armorBoxCount} boxes)`],
      ["Armor Absorbed", `${assignment.armorAbsorbed} box${Number(assignment.armorAbsorbed) === 1 ? "" : "es"}`],
      ["Armor Remaining", `${assignment.armorRemaining} box${Number(assignment.armorRemaining) === 1 ? "" : "es"}`]
    )
  }

  rows.push(
    ["Internals", `${assignment.internals}${Number.isFinite(Number(assignment.internalsTotal)) ? ` (total ${assignment.internalsTotal})` : ""}`]
  )

  for (const [key, value] of rows) {
    const k = document.createElement("div")
    k.className = "damageAssignKey"
    k.textContent = key
    const v = document.createElement("div")
    v.className = "damageAssignVal"
    v.textContent = value
    grid.appendChild(k)
    grid.appendChild(v)
  }

  root.appendChild(grid)

  const note = document.createElement("div")
  note.className = "damageAssignNote"
  note.textContent = assignment.note
  root.appendChild(note)

  const dacLogLines = Array.isArray(assignment.dacLogLines) ? assignment.dacLogLines : []
  if (dacLogLines.length > 0) {
    const logTitle = document.createElement("div")
    logTitle.className = "damageAssignLogTitle"
    logTitle.textContent = normalizeDamageAssignmentMode(assignment.mode) === "automatic"
      ? "Automatic DAC Rolls"
      : "DAC Rolls"
    root.appendChild(logTitle)

    const logList = document.createElement("ol")
    logList.className = "damageAssignLog"
    for (const line of dacLogLines) {
      const item = document.createElement("li")
      item.textContent = line
      logList.appendChild(item)
    }
    root.appendChild(logList)
  }
}

async function closeAssignDamageModal() {
  if (!window.api || typeof window.api.closeAssignDamageWindow !== "function") return
  try {
    await window.api.closeAssignDamageWindow()
    state.assignDamageWindowOpen = false
  } catch {
    // no-op
  }
}

function setAssignDamageHintText(text, isError = false) {
  const hint = byId("assignDamageHint")
  if (!hint) return
  hint.textContent = String(text || "")
  hint.classList.toggle("error", isError === true)
}

function resetActiveShipDacDisplayState() {
  const ship = getActiveShipRecord()
  if (!ship) return

  clearShipDacChartState(ship)
  clearDacResultOnAssignment(ship.damageAssignment)
  renderDamageChart()
  renderDamageAssignmentSummary()
}

function openAssignDamageModal() {
  const ship = getActiveShipRecord()
  if (!ship) {
    setStatus("Load and select a ship before assigning damage.")
    return
  }

  if (!window.api || typeof window.api.openAssignDamageWindow !== "function") {
    setStatus("Assign Damage window API is not available.")
    return
  }

  if (!state.assignDamageWindowOpen) {
    resetActiveShipDacDisplayState()
  }

  Promise.resolve()
    .then(() => window.api.openAssignDamageWindow())
    .then((res) => {
      if (!res || !res.ok) throw new Error(res?.error || "Failed to open Assign Damage window.")
      state.assignDamageWindowOpen = true
      return syncAssignDamageWindowStateForActiveShip()
    })
    .catch((err) => {
      setStatus(err?.message || "Failed to open Assign Damage window.")
    })
}

function buildDamageAssignment(ship, totalDamage, shieldNumber, selectedNonBearingPhasers, assignmentMode = "manual") {
  const doc = ship?.doc || null
  const phaserCount = Math.max(0, getShipPhaserCount(doc))
  const mode = normalizeDamageAssignmentMode(assignmentMode)
  const modeLabel = getDamageAssignmentModeLabel(mode)
  const shieldDamageResult = applyShieldDamageToShip(ship, shieldNumber, totalDamage)
  const selectedList = Array.isArray(selectedNonBearingPhasers)
    ? selectedNonBearingPhasers.filter((item) => item && typeof item === "object")
    : []
  const nonBearing = selectedList.length
  const note = mode === "automatic"
    ? `Automatic mode selected. Shield boxes are marked on the SSD in-memory (${getDamageOverlayLabelLower()} overlay). If present, armor boxes are also marked in-memory after shield overflow, and remaining internals are resolved automatically by rolling 2d6 on the DAC. The original ship JSON is not modified.`
    : `Manual mode selected. Shield boxes are marked on the SSD in-memory (${getDamageOverlayLabelLower()} overlay). If present, armor boxes are also marked in-memory after shield overflow, and only remaining damage is saved as internals. The original ship JSON is not modified.`

  return {
    mode,
    modeLabel,
    totalDamage,
    shieldNumber: shieldDamageResult.shieldNumber,
    shieldPrinted: shieldDamageResult.shieldPrinted,
    shieldBoxCount: shieldDamageResult.shieldBoxCount,
    shieldBoxesPreviouslyDamaged: shieldDamageResult.shieldBoxesPreviouslyDamaged,
    shieldBoxesDamagedThisHit: shieldDamageResult.shieldBoxesDamagedThisHit,
    shieldBoxesDamagedTotal: shieldDamageResult.shieldBoxesDamagedTotal,
    shieldDamagedBoxIndices: shieldDamageResult.newlyDamagedBoxIndices,
    shieldAbsorbed: shieldDamageResult.shieldAbsorbed,
    shieldRemaining: shieldDamageResult.shieldRemaining,
    armorGroupId: shieldDamageResult.armorGroupId,
    armorGroupLabel: shieldDamageResult.armorGroupLabel,
    armorBoxCount: shieldDamageResult.armorBoxCount,
    armorBoxesPreviouslyDamaged: shieldDamageResult.armorBoxesPreviouslyDamaged,
    armorBoxesDamagedThisHit: shieldDamageResult.armorBoxesDamagedThisHit,
    armorBoxesDamagedTotal: shieldDamageResult.armorBoxesDamagedTotal,
    armorDamagedBoxIndices: shieldDamageResult.armorDamagedBoxIndices,
    armorAbsorbed: shieldDamageResult.armorAbsorbed,
    armorRemaining: shieldDamageResult.armorRemaining,
    internals: shieldDamageResult.internals,
    internalsTotal: shieldDamageResult.internalsTotal,
    phaserCount,
    nonBearingPhasers: nonBearing,
    nonBearingPhaserKeys: selectedList.map((item) => String(item.key || "")),
    nonBearingPhaserLabels: selectedList.map((item) => String(item.label || "")).filter(Boolean),
    note,
    dacEvents: [],
    dacLogLines: []
  }
}

async function maybeRunManualInternalDacAllocation(ship, assignment) {
  if (!ship || !assignment || typeof assignment !== "object") {
    return { attempted: 0, resolved: 0, pending: 0, canceled: false, destroyed: false, ran: false }
  }

  const internals = Math.max(0, normalizeInteger(assignment.internals, 0))
  if (internals <= 0) {
    return { attempted: 0, resolved: 0, pending: 0, canceled: false, destroyed: isShipDestroyed(ship), ran: false }
  }

  if (normalizeDamageAssignmentMode(assignment.mode) === "automatic") {
    const result = await applyAutomaticInternalDacHits(ship, internals)
    storeDacResultOnAssignment(assignment, result)
    if (result?.destroyed) {
      assignment.note = `${assignment.note} Ship was destroyed during automatic DAC allocation (ran out of columns in the selected DAC row).`
    } else if (result?.canceled && result?.pending > 0) {
      assignment.note = `${assignment.note} Automatic DAC allocation was canceled with ${result.pending} internal hit${result.pending === 1 ? "" : "s"} not applied to SSD boxes.`
    }
    showAutomaticDacResultList(assignment, result)
    return { ...result, ran: true }
  }

  try {
    window.focus?.()
  } catch {
    // no-op
  }

  const result = await applyManualInternalDacHits(ship, internals)
  storeDacResultOnAssignment(assignment, result)
  renderDamageChart()
  renderShipCanvas()

  const pending = Math.max(0, normalizeInteger(result?.pending, 0))
  if (result?.destroyed) {
    assignment.note = `${assignment.note} Ship was destroyed during manual DAC allocation (ran out of columns in the selected DAC row).`
  } else if (result?.canceled && pending > 0) {
    assignment.note = `${assignment.note} Manual DAC allocation was canceled with ${pending} internal hit${pending === 1 ? "" : "s"} not applied to SSD boxes.`
  }

  return { ...result, ran: true }
}

function shouldKeepAssignDamageWindowOpenAfterAllocation(assignment, internalDacResult) {
  if (!state.assignDamageWindowOpen) return false
  if (!assignment || typeof assignment !== "object") return false
  if (normalizeDamageAssignmentMode(assignment.mode) !== "manual") return false
  if (Math.max(0, normalizeInteger(assignment.internals, 0)) <= 0) return false
  if (!internalDacResult || internalDacResult.ran !== true) return false
  if (internalDacResult.canceled || internalDacResult.destroyed) return false

  const attempted = Math.max(0, normalizeInteger(internalDacResult.attempted, 0))
  const resolved = Math.max(0, normalizeInteger(internalDacResult.resolved, 0))
  return attempted > 0 && resolved >= attempted
}

function shouldCloseAssignDamageWindowBeforeAllocation(assignment) {
  if (!state.assignDamageWindowOpen) return false
  if (!assignment || typeof assignment !== "object") return false
  return normalizeDamageAssignmentMode(assignment.mode) === "automatic"
}

async function submitAssignDamage() {
  const ship = getActiveShipRecord()
  if (!ship) {
    closeAssignDamageModal()
    setStatus("Load and select a ship before assigning damage.")
    return
  }

  const totalDamageInput = byId("assignTotalDamage")
  const shieldSelect = byId("assignShieldNumber")
  const modeSelect = byId("assignDamageMode")

  const totalDamage = normalizeInteger(totalDamageInput?.value, 0)
  const shieldNumber = normalizeInteger(shieldSelect?.value, 1)
  const assignmentMode = normalizeDamageAssignmentMode(modeSelect?.value)
  const selectedNonBearingPhasers = getSelectedAssignDamagePhasers(ship.doc)

  if (!Number.isFinite(totalDamage) || totalDamage <= 0) {
    setAssignDamageHintText("Enter a total damage value greater than 0.", true)
    totalDamageInput?.focus()
    return
  }
  if (!Number.isFinite(shieldNumber) || shieldNumber <= 0) {
    setAssignDamageHintText("Select a valid shield.", true)
    shieldSelect?.focus()
    return
  }

  ship.damageAssignment = buildDamageAssignment(ship, totalDamage, shieldNumber, selectedNonBearingPhasers, assignmentMode)
  if (shouldCloseAssignDamageWindowBeforeAllocation(ship.damageAssignment)) {
    await closeAssignDamageModal()
  }
  let internalDacResult = { attempted: 0, resolved: 0, pending: 0, canceled: false, destroyed: false, ran: false }
  try {
    internalDacResult = await maybeRunManualInternalDacAllocation(ship, ship.damageAssignment)
  } catch (err) {
    ship.damageAssignment.note = `${ship.damageAssignment.note} Manual DAC allocation failed: ${err?.message || "Unknown error"}.`
    setStatus(`Manual DAC allocation failed: ${err?.message || "Unknown error"}`)
  }
  renderDamageAssignmentSummary()
  let statusText =
    `Assigned ${ship.damageAssignment.totalDamage} damage to Shield #${ship.damageAssignment.shieldNumber} (${ship.damageAssignment.modeLabel} mode). ` +
    `Internals: ${ship.damageAssignment.internals}.`
  const assignmentModeLabel = normalizeDamageAssignmentMode(ship.damageAssignment.mode) === "automatic" ? "automatic DAC" : "manual DAC"
  const assignmentModeTitle = normalizeDamageAssignmentMode(ship.damageAssignment.mode) === "automatic" ? "Automatic" : "Manual"
  if (Number(ship.damageAssignment.internals) <= 0) {
    statusText += ` No internals to allocate via ${assignmentModeLabel}.`
  } else if (internalDacResult.ran && internalDacResult.destroyed) {
    statusText += ` Ship destroyed during ${assignmentModeLabel}.`
  } else if (internalDacResult.ran && internalDacResult.canceled && internalDacResult.pending > 0) {
    statusText += ` ${assignmentModeTitle} DAC canceled (${internalDacResult.pending} internal hit${internalDacResult.pending === 1 ? "" : "s"} unresolved).`
  } else if (internalDacResult.ran) {
    statusText += ` ${assignmentModeTitle} DAC resolved ${internalDacResult.resolved}/${internalDacResult.attempted} internal hit${internalDacResult.attempted === 1 ? "" : "s"}.`
  }
  if (!shouldKeepAssignDamageWindowOpenAfterAllocation(ship.damageAssignment, internalDacResult)) {
    await closeAssignDamageModal()
  }
  setStatus(statusText)
}

async function handleAssignDamageWindowSubmit(payload) {
  const requestedIndex = normalizeInteger(payload?.shipIndex, state.activeShipIndex)
  if (!Number.isInteger(requestedIndex) || requestedIndex < 0 || requestedIndex >= state.ships.length) {
    syncAssignDamageWindowStateForActiveShip()
    return
  }

  const ship = state.ships[requestedIndex]
  if (!ship) {
    syncAssignDamageWindowStateForActiveShip()
    return
  }

  const totalDamage = normalizeInteger(payload?.totalDamage, 0)
  const shieldNumber = normalizeInteger(payload?.shieldNumber, 0)
  const assignmentMode = normalizeDamageAssignmentMode(payload?.assignmentMode)
  const selectedKeys = Array.isArray(payload?.selectedNonBearingPhaserKeys)
    ? payload.selectedNonBearingPhaserKeys.map((key) => String(key || "")).filter(Boolean)
    : []

  if (!Number.isFinite(totalDamage) || totalDamage <= 0) {
    syncAssignDamageWindowStateForActiveShip()
    return
  }
  if (!Number.isFinite(shieldNumber) || shieldNumber <= 0) {
    syncAssignDamageWindowStateForActiveShip()
    return
  }

  const phasersByKey = getPhasersByKey(ship.doc)
  const selectedPhasers = []
  for (const key of selectedKeys) {
    const phaser = phasersByKey.get(key)
    if (!phaser) continue
    selectedPhasers.push({ key: phaser.key, label: phaser.fullLabel })
  }

  ship.damageAssignment = buildDamageAssignment(ship, totalDamage, shieldNumber, selectedPhasers, assignmentMode)
  ship.damageAssignmentPreview = null
  state.assignDamageWindowOpen = true
  if (shouldCloseAssignDamageWindowBeforeAllocation(ship.damageAssignment)) {
    await closeAssignDamageModal()
  }
  let internalDacResult = { attempted: 0, resolved: 0, pending: 0, canceled: false, destroyed: false, ran: false }
  try {
    internalDacResult = await maybeRunManualInternalDacAllocation(ship, ship.damageAssignment)
  } catch (err) {
    ship.damageAssignment.note = `${ship.damageAssignment.note} Manual DAC allocation failed: ${err?.message || "Unknown error"}.`
    setStatus(`Manual DAC allocation failed: ${err?.message || "Unknown error"}`)
  }

  if (requestedIndex === state.activeShipIndex) {
    renderDamageAssignmentSummary()
    renderShipCanvas()
  }

  let statusText =
    `Assigned ${ship.damageAssignment.totalDamage} damage to Shield #${ship.damageAssignment.shieldNumber} (${ship.damageAssignment.modeLabel} mode). ` +
    `Internals: ${ship.damageAssignment.internals}.`
  const assignmentModeLabel = normalizeDamageAssignmentMode(ship.damageAssignment.mode) === "automatic" ? "automatic DAC" : "manual DAC"
  const assignmentModeTitle = normalizeDamageAssignmentMode(ship.damageAssignment.mode) === "automatic" ? "Automatic" : "Manual"
  if (Number(ship.damageAssignment.internals) <= 0) {
    statusText += ` No internals to allocate via ${assignmentModeLabel}.`
  } else if (internalDacResult.ran && internalDacResult.destroyed) {
    statusText += ` Ship destroyed during ${assignmentModeLabel}.`
  } else if (internalDacResult.ran && internalDacResult.canceled && internalDacResult.pending > 0) {
    statusText += ` ${assignmentModeTitle} DAC canceled (${internalDacResult.pending} internal hit${internalDacResult.pending === 1 ? "" : "s"} unresolved).`
  } else if (internalDacResult.ran) {
    statusText += ` ${assignmentModeTitle} DAC resolved ${internalDacResult.resolved}/${internalDacResult.attempted} internal hit${internalDacResult.attempted === 1 ? "" : "s"}.`
  }
  if (!shouldKeepAssignDamageWindowOpenAfterAllocation(ship.damageAssignment, internalDacResult)) {
    await closeAssignDamageModal()
  }
  setStatus(statusText)
}

function handleAssignDamageWindowPreview(payload) {
  const requestedIndex = normalizeInteger(payload?.shipIndex, -1)
  if (!Number.isInteger(requestedIndex) || requestedIndex < 0 || requestedIndex >= state.ships.length) return

  const ship = state.ships[requestedIndex]
  if (!ship) return

  const selectedKeys = Array.isArray(payload?.selectedNonBearingPhaserKeys)
    ? Array.from(new Set(payload.selectedNonBearingPhaserKeys.map((key) => String(key || "")).filter(Boolean)))
    : []

  ship.damageAssignmentPreview = {
    nonBearingPhaserKeys: selectedKeys
  }

  if (requestedIndex === state.activeShipIndex) {
    renderShipCanvas()
  }
}

function clearAssignDamagePreviewHighlights() {
  let needsRender = false
  for (let i = 0; i < state.ships.length; i += 1) {
    const ship = state.ships[i]
    if (!ship || !ship.damageAssignmentPreview) continue
    ship.damageAssignmentPreview = null
    if (i === state.activeShipIndex) needsRender = true
  }

  if (needsRender) {
    renderShipCanvas()
  }
}

function setCanvasEmptyVisible(visible) {
  const el = byId("canvasEmpty")
  if (!el) return
  el.classList.toggle("hidden", !visible)
}

function parsePosRect(posText) {
  const raw = String(posText || "").trim()
  if (!raw) return null
  const parts = raw.split(",").map((v) => Number(String(v).trim()))
  if (parts.length < 4 || parts.some((n) => !Number.isFinite(n))) return null
  let [x1, y1, x2, y2] = parts
  if (x2 < x1) [x1, x2] = [x2, x1]
  if (y2 < y1) [y1, y2] = [y2, y1]
  return { x1, y1, x2, y2 }
}

function unionBounds(rects) {
  const list = Array.isArray(rects) ? rects.filter(Boolean) : []
  if (list.length === 0) return null
  let x1 = Infinity
  let y1 = Infinity
  let x2 = -Infinity
  let y2 = -Infinity
  for (const r of list) {
    if (r.x1 < x1) x1 = r.x1
    if (r.y1 < y1) y1 = r.y1
    if (r.x2 > x2) x2 = r.x2
    if (r.y2 > y2) y2 = r.y2
  }
  if (!Number.isFinite(x1) || !Number.isFinite(y1) || !Number.isFinite(x2) || !Number.isFinite(y2)) return null
  return { x1, y1, x2, y2 }
}

function collectSsdRects(ssd, predicate = null) {
  if (!ssd || typeof ssd !== "object") return []
  const out = []
  for (const [key, entries] of Object.entries(ssd)) {
    if (predicate && predicate(key) !== true) continue
    if (!Array.isArray(entries)) continue
    for (const entry of entries) {
      const rect = parsePosRect(entry?.pos)
      if (rect) out.push(rect)
    }
  }
  return out
}

function isShieldKey(key) {
  return /^shield\s*#\d+$/i.test(String(key || "").trim())
}

function computeAverageBoxSize(rects) {
  const list = Array.isArray(rects) ? rects : []
  if (list.length === 0) return 10
  let sum = 0
  let count = 0
  for (const r of list) {
    const w = Math.max(1, Math.abs(Number(r.x2) - Number(r.x1)))
    const h = Math.max(1, Math.abs(Number(r.y2) - Number(r.y1)))
    sum += (w + h) / 2
    count += 1
  }
  return count > 0 ? sum / count : 10
}

function clampCropToImage(crop, imgW, imgH) {
  if (!crop || !Number.isFinite(imgW) || !Number.isFinite(imgH) || imgW <= 0 || imgH <= 0) return null
  let x1 = Math.max(0, Math.floor(crop.x1))
  let y1 = Math.max(0, Math.floor(crop.y1))
  let x2 = Math.min(imgW, Math.ceil(crop.x2))
  let y2 = Math.min(imgH, Math.ceil(crop.y2))
  if (x2 <= x1) x2 = Math.min(imgW, x1 + 1)
  if (y2 <= y1) y2 = Math.min(imgH, y1 + 1)
  return {
    x: x1,
    y: y1,
    w: Math.max(1, x2 - x1),
    h: Math.max(1, y2 - y1)
  }
}

function computeShieldCrop(doc, image) {
  if (!image || !Number.isFinite(image.width) || !Number.isFinite(image.height)) return null
  return {
    x: 0,
    y: 0,
    w: Math.max(1, Math.floor(image.width)),
    h: Math.max(1, Math.floor(image.height)),
    source: "full-image",
    shieldRectCount: 0,
    sourceRectCount: 0,
    pad: 0,
    rawBounds: null
  }
}

function fitRectIntoCanvas(srcW, srcH, dstW, dstH, pad = 12) {
  const usableW = Math.max(1, dstW - pad * 2)
  const usableH = Math.max(1, dstH - pad * 2)
  const scale = Math.min(usableW / srcW, usableH / srcH)
  const drawW = Math.max(1, Math.round(srcW * scale))
  const drawH = Math.max(1, Math.round(srcH * scale))
  const x = Math.round((dstW - drawW) / 2)
  const y = Math.round((dstH - drawH) / 2)
  return { x, y, w: drawW, h: drawH }
}

function renderShipInfo() {
  const grid = byId("shipInfoGrid")
  if (!grid) return
  grid.innerHTML = ""

  const doc = state.doc
  const shipData = doc?.shipData && typeof doc.shipData === "object" ? doc.shipData : {}
  const image = state.image
  const crop = state.crop

  const rows = [
    ["Ship", state.shipLabel || "—"],
    ["Empire", shipData.empire || "—"],
    ["Type", shipData.type || "—"],
    ["Turn Mode", shipData.turnMode || "—"],
    ["Movement Cost", shipData.movementCost || "—"],
    ["JSON File", state.jsonPath || "—"],
    ["Image File", state.imagePath || "—"],
    ["Image Size", image ? `${image.width} × ${image.height}` : "—"],
    [
      "Crop Source",
      crop
        ? (crop.source === "full-image"
          ? "Full image (cropping disabled)"
          : (crop.source === "shields" ? `Shield rectangle (${crop.shieldRectCount} shield boxes)` : `Fallback (${crop.sourceRectCount} SSD boxes)`))
        : "—"
    ],
    ["Crop Bounds", crop ? `x=${crop.x}, y=${crop.y}, w=${crop.w}, h=${crop.h}` : "—"],
    ["Shield Padding", crop ? (crop.source === "full-image" ? "Disabled" : `${crop.pad}px`) : "—"]
  ]

  for (const [key, val] of rows) {
    const row = document.createElement("div")
    row.className = "infoRow"
    const k = document.createElement("div")
    k.className = "infoKey"
    k.textContent = key
    const v = document.createElement("div")
    v.className = "infoVal"
    v.textContent = String(val ?? "")
    row.appendChild(k)
    row.appendChild(v)
    grid.appendChild(row)
  }
}

function drawCanvasMessage(ctx, canvasRect, text) {
  ctx.save()
  ctx.fillStyle = "rgba(255, 255, 255, 0.85)"
  ctx.font = '600 14px "Segoe UI", sans-serif'
  ctx.textAlign = "center"
  ctx.textBaseline = "middle"
  ctx.fillText(String(text || ""), canvasRect.width / 2, canvasRect.height / 2)
  ctx.restore()
}

function projectImageRectIntoCanvas(imageRect, crop, fit) {
  if (!imageRect || !crop || !fit) return null

  const cropX1 = Number(crop.x)
  const cropY1 = Number(crop.y)
  const cropX2 = cropX1 + Number(crop.w)
  const cropY2 = cropY1 + Number(crop.h)

  const x1 = Math.max(cropX1, Number(imageRect.x1))
  const y1 = Math.max(cropY1, Number(imageRect.y1))
  const x2 = Math.min(cropX2, Number(imageRect.x2))
  const y2 = Math.min(cropY2, Number(imageRect.y2))
  if (!Number.isFinite(x1) || !Number.isFinite(y1) || !Number.isFinite(x2) || !Number.isFinite(y2)) return null
  if (x2 <= x1 || y2 <= y1) return null

  const scaleX = Number(fit.w) / Math.max(1, Number(crop.w))
  const scaleY = Number(fit.h) / Math.max(1, Number(crop.h))

  const drawX = Number(fit.x) + (x1 - cropX1) * scaleX
  const drawY = Number(fit.y) + (y1 - cropY1) * scaleY
  const drawW = Math.max(1, (x2 - x1) * scaleX)
  const drawH = Math.max(1, (y2 - y1) * scaleY)

  return { x: drawX, y: drawY, w: drawW, h: drawH }
}

function getManualDamageTargets(ship) {
  const doc = ship?.doc || null
  if (!doc) return []

  const targets = []
  const shieldBoxesByNumber = getShipShieldBoxEntriesByNumber(doc)
  for (const [shieldKey, entries] of Object.entries(shieldBoxesByNumber)) {
    const shieldNumber = normalizeInteger(shieldKey, 0)
    for (const entry of Array.isArray(entries) ? entries : []) {
      if (!entry?.rect) continue
      targets.push({
        type: "shield",
        shieldNumber,
        index: normalizeInteger(entry.index, -1),
        rect: entry.rect
      })
    }
  }

  const armorGroups = Object.values(getShipArmorBoxGroups(doc))
  for (const group of armorGroups) {
    for (const entry of Array.isArray(group?.entries) ? group.entries : []) {
      if (!entry?.rect) continue
      targets.push({
        type: "armor",
        armorGroupId: String(group?.id || "").trim().toLowerCase(),
        armorKey: String(group?.key || "Armor").trim() || "Armor",
        index: normalizeInteger(entry.index, -1),
        rect: entry.rect
      })
    }
  }

  for (const rawKey of getShipSsdKeys(doc)) {
    const ssdKey = String(rawKey || "").trim()
    if (!ssdKey) continue
    if (/^shield\s*#\s*\d+$/i.test(ssdKey)) continue
    if (/armou?r/i.test(ssdKey)) continue

    const entries = getShipSsdBoxEntriesByKey(doc, ssdKey)
    for (const entry of entries) {
      if (!entry?.rect) continue
      targets.push({
        type: "internal",
        ssdKey,
        index: normalizeInteger(entry.index, -1),
        rect: entry.rect
      })
    }
  }

  return targets
}

function findManualDamageTargetAtCanvasPoint(ship, canvasX, canvasY, crop, fit) {
  if (!ship || !crop || !fit) return null

  let best = null
  for (const target of getManualDamageTargets(ship)) {
    const drawRect = projectImageRectIntoCanvas(target.rect, crop, fit)
    if (!drawRect) continue

    const x1 = Number(drawRect.x)
    const y1 = Number(drawRect.y)
    const x2 = x1 + Number(drawRect.w)
    const y2 = y1 + Number(drawRect.h)
    if (canvasX < x1 || canvasY < y1 || canvasX > x2 || canvasY > y2) continue

    const area = Math.max(1, Number(drawRect.w) * Number(drawRect.h))
    const typePriority = target.type === "internal" ? 0 : (target.type === "armor" ? 1 : 2)
    if (
      !best ||
      area < best.area ||
      (area === best.area && typePriority < best.typePriority)
    ) {
      best = { target, area, typePriority }
    }
  }

  return best?.target || null
}

function formatManualDamageTargetLabel(target) {
  if (!target || typeof target !== "object") return "box"
  if (target.type === "shield") {
    return `Shield #${normalizeInteger(target.shieldNumber, 0)} box ${normalizeInteger(target.index, 0) + 1}`
  }
  if (target.type === "armor") {
    return `${String(target.armorKey || "Armor").trim() || "Armor"} box ${normalizeInteger(target.index, 0) + 1}`
  }
  return `${String(target.ssdKey || "Internal").trim() || "Internal"} box ${normalizeInteger(target.index, 0) + 1}`
}

function applyManualDamageEditToTarget(ship, target, mode) {
  const wanted = normalizeManualDamageMode(mode)
  if (!ship || !target || !wanted) return { changed: false }
  const presetForNew = getDamageOverlayPreset()

  if (target.type === "shield") {
    const set = getShipDamagedShieldBoxIndexSet(ship, target.shieldNumber)
    const before = set.has(target.index)
    if (wanted === "mark") {
      if (before) return { changed: false }
      set.add(target.index)
    } else {
      if (!before) return { changed: false }
      set.delete(target.index)
    }
    setShipDamagedShieldBoxIndices(ship, target.shieldNumber, Array.from(set), { presetForNew })
    return { changed: true, label: formatManualDamageTargetLabel(target) }
  }

  if (target.type === "armor") {
    const set = getShipDamagedArmorBoxIndexSet(ship, target.armorGroupId)
    const before = set.has(target.index)
    if (wanted === "mark") {
      if (before) return { changed: false }
      set.add(target.index)
    } else {
      if (!before) return { changed: false }
      set.delete(target.index)
    }
    setShipDamagedArmorBoxIndices(ship, target.armorGroupId, Array.from(set), { presetForNew })
    return { changed: true, label: formatManualDamageTargetLabel(target) }
  }

  if (target.type === "internal") {
    const set = getShipDamagedSsdBoxIndexSet(ship, target.ssdKey)
    const before = set.has(target.index)
    if (wanted === "mark") {
      if (before) return { changed: false }
      set.add(target.index)
    } else {
      if (!before) return { changed: false }
      set.delete(target.index)
    }
    setShipDamagedSsdBoxIndices(ship, target.ssdKey, Array.from(set), { presetForNew })

    const runtime = ensureShipRuntimeDamageState(ship)
    const delta = wanted === "mark" ? 1 : -1
    runtime.internalsTotal = Math.max(0, normalizeInteger(runtime.internalsTotal, 0) + delta)
    return { changed: true, label: formatManualDamageTargetLabel(target) }
  }

  return { changed: false }
}

function handleShipCanvasClick(event) {
  if (state.zoomSectionMode === true) return

  const mode = normalizeManualDamageMode(state.manualDamageMode)
  if (!mode) return

  const context = getActiveShipCanvasInteractionContext()
  if (!context) return

  const canvasX = Number(event?.clientX) - context.rect.left
  const canvasY = Number(event?.clientY) - context.rect.top
  if (!Number.isFinite(canvasX) || !Number.isFinite(canvasY)) return

  const target = findManualDamageTargetAtCanvasPoint(context.ship, canvasX, canvasY, context.crop, context.fit)
  if (!target) return

  const result = applyManualDamageEditToTarget(context.ship, target, mode)
  if (!result.changed) return

  renderShipCanvas()
  renderDamageAssignmentSummary()
  setStatus(`${mode === "mark" ? "Marked" : "Removed"} ${result.label}.`)
}

function drawDamagedShieldBoxOverlays(ctx, ship, crop, fit) {
  if (!ctx || !ship || !crop || !fit) return

  const damagedRects = getShipAllDamagedShieldBoxRects(ship)
  if (!Array.isArray(damagedRects) || damagedRects.length === 0) return

  ctx.save()
  ctx.lineJoin = "round"
  ctx.lineCap = "round"
  ctx.lineWidth = 1.75
  ctx.shadowBlur = 6

  for (const item of damagedRects) {
    const palette = getDamageOverlayTheme(item?.overlayPreset).shield
    ctx.strokeStyle = palette.stroke
    ctx.fillStyle = palette.fill
    ctx.shadowColor = palette.shadow

    const drawRect = projectImageRectIntoCanvas(item?.rect, crop, fit)
    if (!drawRect) continue

    const x = Math.round(drawRect.x) + 0.5
    const y = Math.round(drawRect.y) + 0.5
    const w = Math.max(1, Math.round(drawRect.w) - 1)
    const h = Math.max(1, Math.round(drawRect.h) - 1)

    ctx.fillRect(x, y, w, h)
    ctx.strokeRect(x, y, w, h)
  }

  ctx.restore()
}

function drawDamagedArmorBoxOverlays(ctx, ship, crop, fit) {
  if (!ctx || !ship || !crop || !fit) return

  const damagedRects = getShipAllDamagedArmorBoxRects(ship)
  if (!Array.isArray(damagedRects) || damagedRects.length === 0) return

  ctx.save()
  ctx.lineJoin = "round"
  ctx.lineCap = "round"
  ctx.lineWidth = 1.75
  ctx.shadowBlur = 6

  for (const item of damagedRects) {
    const palette = getDamageOverlayTheme(item?.overlayPreset).armor
    ctx.strokeStyle = palette.stroke
    ctx.fillStyle = palette.fill
    ctx.shadowColor = palette.shadow

    const drawRect = projectImageRectIntoCanvas(item?.rect, crop, fit)
    if (!drawRect) continue

    const x = Math.round(drawRect.x) + 0.5
    const y = Math.round(drawRect.y) + 0.5
    const w = Math.max(1, Math.round(drawRect.w) - 1)
    const h = Math.max(1, Math.round(drawRect.h) - 1)

    ctx.fillRect(x, y, w, h)
    ctx.strokeRect(x, y, w, h)
  }

  ctx.restore()
}

function drawDamagedInternalBoxOverlays(ctx, ship, crop, fit) {
  if (!ctx || !ship || !crop || !fit) return

  const damagedRects = getShipAllDamagedInternalBoxRects(ship)
  if (!Array.isArray(damagedRects) || damagedRects.length === 0) return

  ctx.save()
  ctx.lineJoin = "round"
  ctx.lineCap = "round"
  ctx.lineWidth = 1.6
  ctx.shadowBlur = 5

  for (const item of damagedRects) {
    const palette = getDamageOverlayTheme(item?.overlayPreset).internal
    ctx.strokeStyle = palette.stroke
    ctx.fillStyle = palette.fill
    ctx.shadowColor = palette.shadow

    const drawRect = projectImageRectIntoCanvas(item?.rect, crop, fit)
    if (!drawRect) continue

    const x = Math.round(drawRect.x) + 0.5
    const y = Math.round(drawRect.y) + 0.5
    const w = Math.max(1, Math.round(drawRect.w) - 1)
    const h = Math.max(1, Math.round(drawRect.h) - 1)

    ctx.fillRect(x, y, w, h)
    ctx.strokeRect(x, y, w, h)
  }

  ctx.restore()
}

function drawDestroyedShipOverlay(ctx, ship, fit) {
  if (!ctx || !ship || !fit) return
  if (!isShipDestroyed(ship)) return

  const x1 = Number(fit.x)
  const y1 = Number(fit.y)
  const x2 = x1 + Number(fit.w)
  const y2 = y1 + Number(fit.h)

  ctx.save()
  ctx.fillStyle = "rgba(48, 0, 0, 0.2)"
  ctx.fillRect(x1, y1, Math.max(1, fit.w), Math.max(1, fit.h))

  ctx.strokeStyle = "rgba(255, 32, 32, 0.88)"
  ctx.lineWidth = Math.max(4, Math.round(Math.min(fit.w, fit.h) * 0.012))
  ctx.lineCap = "round"
  ctx.shadowColor = "rgba(255, 48, 48, 0.5)"
  ctx.shadowBlur = 12

  ctx.beginPath()
  ctx.moveTo(x1 + 8, y1 + 8)
  ctx.lineTo(x2 - 8, y2 - 8)
  ctx.moveTo(x2 - 8, y1 + 8)
  ctx.lineTo(x1 + 8, y2 - 8)
  ctx.stroke()

  ctx.fillStyle = "rgba(255, 220, 220, 0.94)"
  ctx.font = '700 18px "Segoe UI", sans-serif'
  ctx.textAlign = "center"
  ctx.textBaseline = "middle"
  ctx.fillText("DESTROYED", x1 + fit.w / 2, y1 + fit.h / 2)
  ctx.restore()
}

function drawZoomSectionSelectionOverlay(ctx) {
  if (!ctx || state.zoomSectionMode !== true) return
  const drag = state.zoomDrag
  if (!drag) return

  const x = Math.min(Number(drag.startX), Number(drag.currentX))
  const y = Math.min(Number(drag.startY), Number(drag.currentY))
  const w = Math.abs(Number(drag.currentX) - Number(drag.startX))
  const h = Math.abs(Number(drag.currentY) - Number(drag.startY))
  if (w <= 0 || h <= 0) return

  ctx.save()
  ctx.fillStyle = "rgba(255, 92, 92, 0.16)"
  ctx.strokeStyle = "rgba(255, 196, 196, 0.95)"
  ctx.lineWidth = 1.5
  ctx.setLineDash([7, 5])
  ctx.fillRect(x, y, w, h)
  ctx.strokeRect(Math.round(x) + 0.5, Math.round(y) + 0.5, Math.max(1, Math.round(w) - 1), Math.max(1, Math.round(h) - 1))
  ctx.restore()
}

function renderShipCanvas() {
  const canvas = byId("shipCanvas")
  if (!canvas) return
  const rect = canvas.getBoundingClientRect()
  if (rect.width <= 0 || rect.height <= 0) return

  const dpr = window.devicePixelRatio || 1
  canvas.width = Math.max(1, Math.floor(rect.width * dpr))
  canvas.height = Math.max(1, Math.floor(rect.height * dpr))
  const ctx = canvas.getContext("2d")
  ctx.setTransform(dpr, 0, 0, dpr, 0, 0)
  ctx.clearRect(0, 0, rect.width, rect.height)

  ctx.fillStyle = "rgba(18, 20, 26, 0.92)"
  ctx.fillRect(0, 0, rect.width, rect.height)

  const activeShip = getActiveShipRecord()
  const image = state.image
  const crop = getShipCanvasSourceCrop(activeShip)
  if (!image || !crop) {
    setCanvasEmptyVisible(true)
    drawCanvasMessage(ctx, rect, "No ship image available.")
    return
  }
  setCanvasEmptyVisible(false)

  const fit = fitRectIntoCanvas(crop.w, crop.h, rect.width, rect.height, 4)
  ctx.save()
  ctx.shadowColor = "rgba(0,0,0,0.45)"
  ctx.shadowBlur = 16
  ctx.drawImage(
    image,
    crop.x, crop.y, crop.w, crop.h,
    fit.x, fit.y, fit.w, fit.h
  )
  ctx.restore()

  drawDamagedArmorBoxOverlays(ctx, activeShip, crop, fit)
  drawDamagedShieldBoxOverlays(ctx, activeShip, crop, fit)
  drawDamagedInternalBoxOverlays(ctx, activeShip, crop, fit)
  drawDestroyedShipOverlay(ctx, activeShip, fit)
  drawDuplicateShipMarkers(ctx, activeShip, fit)
  drawZoomSectionSelectionOverlay(ctx)

  ctx.save()
  ctx.strokeStyle = "rgba(255, 92, 92, 0.9)"
  ctx.lineWidth = 2
  ctx.strokeRect(fit.x + 0.5, fit.y + 0.5, Math.max(1, fit.w - 1), Math.max(1, fit.h - 1))
  ctx.restore()
}

function renderDamageChart() {
  const table = byId("damageChartTable")
  if (!table) return
  table.innerHTML = ""

  const activeShip = getActiveShipRecord()
  const markedDacCells = getShipMarkedDacCellSet(activeShip)
  const passedOverDacCells = getShipPassedOverDacCellSet(activeShip)
  const activeDacCellKey = getShipActiveDacCellKey(activeShip)

  const caption = document.createElement("caption")
  caption.textContent = DAMAGE_ALLOCATION_CHART.caption
  table.appendChild(caption)

  const thead = document.createElement("thead")
  const hr = document.createElement("tr")

  const corner = document.createElement("th")
  corner.className = "corner"
  corner.scope = "col"
  corner.innerHTML = "DIE<br>ROLL"
  hr.appendChild(corner)

  for (const col of DAMAGE_ALLOCATION_CHART.columns) {
    const th = document.createElement("th")
    th.scope = "col"
    th.textContent = col
    hr.appendChild(th)
  }
  thead.appendChild(hr)
  table.appendChild(thead)

  const tbody = document.createElement("tbody")
  for (const rowDef of DAMAGE_ALLOCATION_CHART.rows) {
    const tr = document.createElement("tr")
    const rowHead = document.createElement("th")
    rowHead.scope = "row"
    rowHead.className = "rowHead"
    rowHead.textContent = rowDef.dieRoll
    tr.appendChild(rowHead)

    for (let i = 0; i < DAMAGE_ALLOCATION_CHART.columns.length; i++) {
      const td = document.createElement("td")
      const text = String(rowDef.cells?.[i] || "")
      const columnKey = DAMAGE_ALLOCATION_CHART.columns[i]
      const dacCellKey = `${String(rowDef.dieRoll)}:${String(columnKey)}`
      const span = document.createElement("span")
      span.className = "cellText"
      if (isUnderlinedDamageCell(rowDef.dieRoll, columnKey)) {
        span.classList.add("underlined")
      }
      if (passedOverDacCells.has(dacCellKey)) {
        td.classList.add("dacCellPassed")
        span.classList.add("dacPassed")
      }
      if (markedDacCells.has(dacCellKey)) {
        td.classList.add("dacCellMarked")
        span.classList.add("dacMarked")
      }
      if (activeDacCellKey && dacCellKey === activeDacCellKey) {
        td.classList.add("dacCellActive")
        span.classList.add("dacActive")
      }
      span.textContent = text
      td.appendChild(span)
      tr.appendChild(td)
    }
    tbody.appendChild(tr)
  }

  table.appendChild(tbody)
}

function getShipDisplayLabel(doc, fallback) {
  const shipData = doc?.shipData || {}
  const shipName = String(doc?.shipName || "").trim()
  const empire = String(shipData.empire || "").trim()
  const type = String(shipData.type || "").trim()
  return shipName || [empire, type].filter(Boolean).join(" ") || fallback || "Ship"
}

function loadImage(dataUrl) {
  return new Promise((resolve, reject) => {
    const img = new Image()
    img.onload = () => resolve(img)
    img.onerror = () => reject(new Error("Failed to load image data."))
    img.src = dataUrl
  })
}

function cloneJsonValue(value, fallback = null) {
  if (value === undefined) return fallback
  try {
    return JSON.parse(JSON.stringify(value))
  } catch {
    return fallback
  }
}

function getShipCountLabel(count) {
  const n = Math.max(0, normalizeInteger(count, 0))
  return `${n} ship${n === 1 ? "" : "s"}`
}

function cloneCropForSave(rawCrop) {
  if (!rawCrop || typeof rawCrop !== "object") return null

  const x = Number(rawCrop.x)
  const y = Number(rawCrop.y)
  const w = Number(rawCrop.w)
  const h = Number(rawCrop.h)
  if (!Number.isFinite(x) || !Number.isFinite(y) || !Number.isFinite(w) || !Number.isFinite(h)) return null
  if (w <= 0 || h <= 0) return null

  const out = {
    x: Math.floor(x),
    y: Math.floor(y),
    w: Math.max(1, Math.floor(w)),
    h: Math.max(1, Math.floor(h)),
    source: String(rawCrop.source || "").trim() || "saved"
  }

  for (const key of ["shieldRectCount", "sourceRectCount", "pad"]) {
    const n = Number(rawCrop[key])
    if (Number.isFinite(n)) out[key] = n
  }

  const rawBounds = rawCrop.rawBounds && typeof rawCrop.rawBounds === "object"
    ? cloneJsonValue(rawCrop.rawBounds, null)
    : null
  if (rawBounds) out.rawBounds = rawBounds

  return out
}

function normalizeSavedImageCrop(rawCrop, fallbackCrop, image) {
  const crop = cloneCropForSave(rawCrop)
  if (!crop) return fallbackCrop || null

  const imgW = Number(image?.width)
  const imgH = Number(image?.height)
  if (!Number.isFinite(imgW) || !Number.isFinite(imgH) || imgW <= 0 || imgH <= 0) {
    return crop
  }

  const x = Math.max(0, Math.min(Math.floor(crop.x), Math.max(0, Math.floor(imgW) - 1)))
  const y = Math.max(0, Math.min(Math.floor(crop.y), Math.max(0, Math.floor(imgH) - 1)))
  const maxW = Math.max(1, Math.floor(imgW) - x)
  const maxH = Math.max(1, Math.floor(imgH) - y)

  return {
    ...crop,
    x,
    y,
    w: Math.max(1, Math.min(Math.floor(crop.w), maxW)),
    h: Math.max(1, Math.min(Math.floor(crop.h), maxH)),
    source: crop.source || fallbackCrop?.source || "saved"
  }
}

function buildShipSaveRecord(ship, index) {
  if (!ship || typeof ship !== "object") return null
  const runtimeDamage = ensureShipRuntimeDamageState(ship)

  return {
    saveVersion: STARHOLD_SAVE_VERSION,
    shipIndex: index,
    doc: cloneJsonValue(ship.doc, null),
    imageDataUrl: String(ship.imageDataUrl || ""),
    imagePath: String(ship.imagePath || ""),
    jsonPath: String(ship.jsonPath || ""),
    folderPath: String(ship.folderPath || ""),
    jsonFile: String(ship.jsonFile || ""),
    imageFile: String(ship.imageFile || ""),
    baseName: String(ship.baseName || ""),
    shipLabel: String(ship.shipLabel || ""),
    crop: cloneCropForSave(ship.crop),
    zoomCrop: cloneCropForSave(ship.zoomCrop),
    runtimeDamage: cloneJsonValue(runtimeDamage, {}),
    damageAssignment: cloneJsonValue(ship.damageAssignment || null, null)
  }
}

function buildGameSavePayload() {
  const ships = state.ships
    .map((ship, index) => buildShipSaveRecord(ship, index))
    .filter(Boolean)

  return {
    schema: STARHOLD_SAVE_SCHEMA,
    version: STARHOLD_SAVE_VERSION,
    savedAt: new Date().toISOString(),
    appVersion: String(state.appVersion || ""),
    activeShipIndex: Number.isInteger(state.activeShipIndex) ? state.activeShipIndex : -1,
    damageOverlayPreset: getDamageOverlayPreset(),
    ships
  }
}

async function saveGame() {
  if (!window.api || typeof window.api.saveGame !== "function") {
    setStatus("Save game API is not available.")
    return
  }
  if (state.ships.length === 0) {
    setStatus("Load at least one ship before saving a game.")
    return
  }

  const payload = buildGameSavePayload()
  setStatus(`Saving ${getShipCountLabel(payload.ships.length)}...`)
  try {
    const res = await window.api.saveGame(payload)
    if (!res || !res.ok) {
      setStatus(res?.error || "Failed to save game.")
      return
    }
    if (res.canceled) {
      setStatus("Save canceled.")
      return
    }

    setStatus(`Saved ${getShipCountLabel(payload.ships.length)}.\n${res.filePath || ""}`.trim())
  } catch (err) {
    setStatus(err?.message || "Failed to save game.")
  }
}

function getSavedShipList(saveGame) {
  const input = saveGame && typeof saveGame === "object" ? saveGame : null
  if (!input) return []
  return Array.isArray(input.ships) ? input.ships : []
}

async function restoreShipRecordFromSave(savedShip, index) {
  if (!savedShip || typeof savedShip !== "object") {
    throw new Error(`Ship ${index + 1} is not a valid saved ship record.`)
  }

  const doc = cloneJsonValue(savedShip.doc, null)
  if (!doc || typeof doc !== "object") {
    throw new Error(`Ship ${index + 1} is missing saved ship JSON data.`)
  }

  const imageDataUrl = String(savedShip.imageDataUrl || "")
  if (!imageDataUrl) {
    throw new Error(`Ship ${index + 1} is missing saved ship image data.`)
  }

  const img = await loadImage(imageDataUrl)
  const computedCrop = computeShieldCrop(doc, img)
  const crop = normalizeSavedImageCrop(savedShip.crop, computedCrop, img) || computedCrop
  const shipLabel = String(savedShip.shipLabel || "").trim() ||
    getShipDisplayLabel(doc, savedShip.baseName || savedShip.jsonFile || `Ship ${index + 1}`)

  const shipRecord = {
    doc,
    image: img,
    imageDataUrl,
    imagePath: String(savedShip.imagePath || ""),
    jsonPath: String(savedShip.jsonPath || ""),
    folderPath: String(savedShip.folderPath || ""),
    jsonFile: String(savedShip.jsonFile || ""),
    imageFile: String(savedShip.imageFile || ""),
    baseName: String(savedShip.baseName || ""),
    shipLabel,
    crop,
    zoomCrop: null,
    runtimeDamage: cloneJsonValue(savedShip.runtimeDamage, {}),
    damageAssignment: cloneJsonValue(savedShip.damageAssignment || null, null),
    damageAssignmentPreview: null
  }

  ensureShipRuntimeDamageState(shipRecord)
  const savedZoom = normalizeSavedImageCrop(savedShip.zoomCrop, null, img)
  if (savedZoom) {
    setShipZoomCrop(shipRecord, savedZoom)
  }

  return shipRecord
}

async function restoreGameFromSave(saveGame, filePath = "") {
  const input = saveGame && typeof saveGame === "object" ? saveGame : null
  if (!input) {
    throw new Error("The selected file is not a valid StarHold save.")
  }

  const savedShips = getSavedShipList(input)
  if (savedShips.length === 0) {
    throw new Error("The selected save does not contain any ships.")
  }

  const restoredShips = []
  const failures = []
  for (let i = 0; i < savedShips.length; i += 1) {
    try {
      restoredShips.push(await restoreShipRecordFromSave(savedShips[i], i))
    } catch (err) {
      failures.push(`Ship ${i + 1}: ${err?.message || "failed to restore"}`)
    }
  }

  if (restoredShips.length === 0) {
    throw new Error(`No ships could be restored from the selected save.\n${failures.join("\n")}`)
  }

  stopShipRotation({ quiet: true })
  state.manualDamageMode = ""
  state.zoomSectionMode = false
  state.zoomDrag = null
  state.pendingManualDacChartClearShip = null
  state.ships = restoredShips
  state.activeShipIndex = -1
  state.damageOverlayPreset = normalizeDamageOverlayPreset(input.damageOverlayPreset)
  renderDamageColorSelector()
  storeDamageOverlayPresetToStorage(state.damageOverlayPreset)

  const activeIndex = restoredShips.length > 0
    ? clampInteger(input.activeShipIndex, 0, restoredShips.length - 1)
    : -1
  if (activeIndex >= 0) {
    activateLoadedShip(activeIndex, { quiet: true })
  } else {
    applyShipRecordToState(null)
    renderShipSwitcher()
    renderShipInfo()
    renderDamageChart()
    renderDamageAssignmentSummary()
    renderShipCanvas()
  }

  let status = `Loaded ${getShipCountLabel(restoredShips.length)} from saved game.`
  if (filePath) status += `\n${filePath}`
  if (failures.length > 0) {
    status += `\nSkipped ${failures.length} ship${failures.length === 1 ? "" : "s"} that could not be restored.`
  }
  setStatus(status)
}

async function loadGame() {
  if (!window.api || typeof window.api.loadGame !== "function") {
    setStatus("Load game API is not available.")
    return
  }

  if (state.ships.length > 0 && typeof window.confirm === "function") {
    const replace = window.confirm("Load a saved game? This will replace the currently loaded ships in StarHold.")
    if (!replace) return
  }

  setStatus("Choose a StarHold save file...")
  try {
    const res = await window.api.loadGame()
    if (!res || !res.ok) {
      setStatus(res?.error || "Failed to load game.")
      return
    }
    if (res.canceled) {
      setStatus("Load canceled.")
      return
    }

    await restoreGameFromSave(res.saveGame, res.filePath || "")
  } catch (err) {
    setStatus(err?.message || "Failed to load game.")
  }
}

async function loadStarholdSettings() {
  updateSettingsButton()
  if (!window.api || typeof window.api.getSettings !== "function") return

  try {
    const res = await window.api.getSettings()
    if (!res || !res.ok) return
    setDefaultShipFolder(res.defaultShipFolder || "")
  } catch {
    // Ignore settings load failures; opening ship files still works.
  }
}

async function chooseDefaultShipFolder() {
  if (!window.api || typeof window.api.chooseDefaultShipFolder !== "function") {
    setStatus("Settings API is not available.")
    return
  }

  setStatus("Choose a default starting folder for ship files...")
  try {
    const res = await window.api.chooseDefaultShipFolder()
    if (!res || !res.ok) {
      setStatus(res?.error || "Failed to choose default ship folder.")
      return
    }
    if (res.canceled) {
      setStatus("Default ship folder unchanged.")
      return
    }

    setDefaultShipFolder(res.defaultShipFolder || "")
    if (state.defaultShipFolder) {
      setStatus(`Default ship folder set:\n${state.defaultShipFolder}`)
    } else {
      setStatus("Default ship folder cleared.")
    }
  } catch (err) {
    setStatus(err?.message || "Failed to choose default ship folder.")
  }
}

async function openSuperluminalShip() {
  if (!window.api || typeof window.api.openSuperluminalShip !== "function") {
    setStatus("API is not available.")
    return
  }

  setStatus("Opening SuperLuminal ship folder...")
  try {
    const res = await window.api.openSuperluminalShip()
    if (!res || !res.ok) {
      setStatus(res?.error || "No folder selected.")
      return
    }

    let doc
    try {
      doc = JSON.parse(res.jsonText || "")
    } catch (err) {
      setStatus(`JSON parse failed: ${err.message}`)
      return
    }

    const img = await loadImage(res.imageDataUrl)
    const crop = computeShieldCrop(doc, img)
    const shipRecord = {
      doc,
      image: img,
      imageDataUrl: res.imageDataUrl,
      imagePath: res.imagePath || "",
      jsonPath: res.jsonPath || "",
      folderPath: res.folderPath || "",
      jsonFile: res.jsonFile || "",
      imageFile: res.imageFile || "",
      baseName: res.baseName || "",
      shipLabel: getShipDisplayLabel(doc, res.baseName || res.jsonFile || "Ship"),
      crop,
      zoomCrop: null
    }

    state.ships.push(shipRecord)
    const activeIndex = state.ships.length - 1

    activateLoadedShip(activeIndex, { quiet: true })

    if (!crop) {
      setStatus(`Loaded ship ${activeIndex + 1}/${state.ships.length}, but the image could not be displayed.`)
    } else {
      const sourceText = crop.source === "full-image"
        ? "full image"
        : (crop.source === "shields" ? "shield-bounded" : "fallback SSD-bounded")
      setStatus(
        `Loaded ${res.jsonFile} and ${res.imageFile}.\nShowing ${sourceText} ship area (${crop.w}x${crop.h}). Loaded ships: ${state.ships.length}.`
      )
    }
  } catch (err) {
    setStatus(err?.message || "Failed to open ship.")
  }
}

function init() {
  state.damageOverlayPreset = loadDamageOverlayPresetFromStorage()
  updateWindowTitle()
  renderDamageColorSelector()
  renderAppVersion()
  renderDamageChart()
  renderShipSwitcher()
  renderShipInfo()
  renderDamageAssignmentSummary()
  renderShipCanvas()
  updateSettingsButton()
  updateAssignDamageButton()

  const btn = byId("btnOpenShip")
  if (btn) btn.onclick = () => openSuperluminalShip()

  const btnAssignDamage = byId("btnAssignDamage")
  if (btnAssignDamage) btnAssignDamage.onclick = () => openAssignDamageModal()

  const btnSaveGame = byId("btnSaveGame")
  if (btnSaveGame) btnSaveGame.onclick = () => saveGame()

  const btnLoadGame = byId("btnLoadGame")
  if (btnLoadGame) btnLoadGame.onclick = () => loadGame()

  const btnCloseShip = byId("btnCloseShip")
  if (btnCloseShip) btnCloseShip.onclick = () => closeActiveShip()

  const btnSettings = byId("btnSettings")
  if (btnSettings) btnSettings.onclick = () => chooseDefaultShipFolder()

  const btnRotateShips = byId("btnRotateShips")
  if (btnRotateShips) btnRotateShips.onclick = () => toggleShipRotation()

  const btnManualMarkDamage = byId("btnManualMarkDamage")
  if (btnManualMarkDamage) btnManualMarkDamage.onclick = () => setManualDamageMode("mark")

  const btnManualRemoveDamage = byId("btnManualRemoveDamage")
  if (btnManualRemoveDamage) btnManualRemoveDamage.onclick = () => setManualDamageMode("remove")

  const btnZoomSection = byId("btnZoomSection")
  if (btnZoomSection) btnZoomSection.onclick = () => toggleZoomSectionMode()

  const btnResetZoom = byId("btnResetZoom")
  if (btnResetZoom) btnResetZoom.onclick = () => resetActiveShipZoom()

  const damageColorSelect = byId("damageColorSelect")
  if (damageColorSelect) {
    damageColorSelect.onchange = () => applyDamageOverlayPreset(damageColorSelect.value)
  }

  const btnWeaponSelectClose = byId("btnWeaponSelectClose")
  if (btnWeaponSelectClose) btnWeaponSelectClose.onclick = () => cancelWeaponSelectionPrompt()

  const btnWeaponSelectSkip = byId("btnWeaponSelectSkip")
  if (btnWeaponSelectSkip) btnWeaponSelectSkip.onclick = () => resolveWeaponSelectionPrompt(WEAPON_SELECTION_SKIP_TOKEN)

  const btnWeaponSelectCancel = byId("btnWeaponSelectCancel")
  if (btnWeaponSelectCancel) btnWeaponSelectCancel.onclick = () => cancelWeaponSelectionPrompt()

  const btnAssignDamageClose = byId("btnAssignDamageClose")
  if (btnAssignDamageClose) btnAssignDamageClose.onclick = () => closeAssignDamageModal()

  const btnAssignDamageCancel = byId("btnAssignDamageCancel")
  if (btnAssignDamageCancel) btnAssignDamageCancel.onclick = () => closeAssignDamageModal()

  const btnAssignDamageSubmit = byId("btnAssignDamageSubmit")
  if (btnAssignDamageSubmit) btnAssignDamageSubmit.onclick = () => submitAssignDamage()

  if (window.api && typeof window.api.onAssignDamageWindowSubmit === "function") {
    window.api.onAssignDamageWindowSubmit((payload) => {
      state.assignDamageWindowOpen = true
      handleAssignDamageWindowSubmit(payload)
    })
  }

  if (window.api && typeof window.api.onAssignDamageWindowPreview === "function") {
    window.api.onAssignDamageWindowPreview((payload) => {
      state.assignDamageWindowOpen = true
      handleAssignDamageWindowPreview(payload)
    })
  }

  if (window.api && typeof window.api.onAssignDamageWindowClosed === "function") {
    window.api.onAssignDamageWindowClosed(() => {
      state.assignDamageWindowOpen = false
      state.assignDamageWeaponPromptSyncPromise = null
      if (state.manualDacPromptPromise) {
        resolveManualDacRollPrompt(null)
      }
      if (state.weaponSelectPromise) {
        resolveWeaponSelectionPrompt(null)
      }
      clearAssignDamagePreviewHighlights()
      flushQueuedManualDacChartClear()
    })
  }

  if (window.api && typeof window.api.onWeaponSelectWindowClosed === "function") {
    window.api.onWeaponSelectWindowClosed(() => {
      state.weaponSelectWindowOpen = false
      state.weaponSelectWindowClosingPromise = null
      if (state.ignoreNextWeaponSelectWindowClosed) {
        state.ignoreNextWeaponSelectWindowClosed = false
        return
      }
      if (state.weaponSelectPromise && !state.assignDamageWindowOpen) {
        resolveWeaponSelectionPrompt(null)
      }
    })
  }

  if (window.api && typeof window.api.onAssignDamageWeaponSelectionSubmit === "function") {
    window.api.onAssignDamageWeaponSelectionSubmit((payload) => {
      if (!state.weaponSelectPromise) return
      resolveWeaponSelectionPrompt(String(payload?.optionId || ""))
    })
  }

  if (window.api && typeof window.api.onAssignDamageWeaponSelectionCancel === "function") {
    window.api.onAssignDamageWeaponSelectionCancel((payload) => {
      if (!state.weaponSelectPromise) return
      if (payload?.skip === true) {
        resolveWeaponSelectionPrompt(WEAPON_SELECTION_SKIP_TOKEN)
        return
      }
      resolveWeaponSelectionPrompt(null)
    })
  }

  if (window.api && typeof window.api.onAssignDamageManualDacPromptSubmit === "function") {
    window.api.onAssignDamageManualDacPromptSubmit((payload) => {
      const roll = normalizeInteger(payload?.roll, NaN)
      if (!Number.isFinite(roll) || roll < 2 || roll > 12) return
      resolveManualDacRollPrompt(roll)
    })
  }

  if (window.api && typeof window.api.onAssignDamageManualDacPromptCancel === "function") {
    window.api.onAssignDamageManualDacPromptCancel(() => {
      resolveManualDacRollPrompt(null)
    })
  }

  if (window.api && typeof window.api.onManualDacRollWindowSubmit === "function") {
    window.api.onManualDacRollWindowSubmit((payload) => {
      const roll = normalizeInteger(payload?.roll, NaN)
      if (!Number.isFinite(roll) || roll < 2 || roll > 12) return
      resolveManualDacRollPrompt(roll)
    })
  }

  if (window.api && typeof window.api.onManualDacRollWindowCancel === "function") {
    window.api.onManualDacRollWindowCancel(() => {
      resolveManualDacRollPrompt(null)
    })
  }

  if (window.api && typeof window.api.onManualDacRollWindowClosed === "function") {
    window.api.onManualDacRollWindowClosed(() => {
      resolveManualDacRollPrompt(null)
      flushQueuedManualDacChartClear()
    })
  }

  const shipSelect = byId("shipSelect")
  if (shipSelect) {
    shipSelect.onchange = () => {
      const idx = Number(shipSelect.value)
      if (!Number.isInteger(idx)) return
      activateLoadedShip(idx)
    }
  }

  const shipCanvas = byId("shipCanvas")
  if (shipCanvas) {
    shipCanvas.addEventListener("mousedown", (event) => {
      beginZoomSectionDrag(event)
    })
    shipCanvas.addEventListener("mousemove", (event) => {
      updateZoomSectionDrag(event)
    })
    shipCanvas.addEventListener("click", (event) => {
      handleShipCanvasClick(event)
    })
  }
  window.addEventListener("mouseup", (event) => {
    finalizeZoomSectionDrag(event)
  })

  loadStarholdSettings()
  loadRollsConfig()
  loadAppVersion()
  syncAssignDamageWindowStateForActiveShip()

  window.addEventListener("keydown", (event) => {
    if (event.key !== "Escape") return
    if (state.weaponSelectPromise) {
      event.preventDefault()
      cancelWeaponSelectionPrompt()
      return
    }
    if (state.zoomDrag) {
      event.preventDefault()
      state.zoomDrag = null
      renderShipCanvas()
      return
    }
    if (state.zoomSectionMode) {
      event.preventDefault()
      setZoomSectionMode(false)
      return
    }
    if (state.manualDamageMode) {
      setManualDamageMode("")
      return
    }
    closeAssignDamageModal()
  })

  window.addEventListener("resize", () => {
    renderShipCanvas()
  })
}

init()

