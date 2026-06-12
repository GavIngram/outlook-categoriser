"use strict";

// ── Colour palette — hex values match Outlook 365 / OWA category presets ──
// Source: Slipstick Systems reverse-engineering of OWA category colour swatches.
// Preset names (from Office.MailboxEnums.CategoryColor docs):
//   0=Red 1=Orange 2=Brown 3=Yellow 4=Green 5=Teal 6=Olive 7=Blue
//   8=Purple 9=Cranberry 10=Steel 11=DarkSteel 12=Gray 13=DarkGray
//   14=Black 15=DarkRed 16=DarkOrange 17=DarkBrown 18=DarkYellow
//   19=DarkGreen 20=DarkTeal 21=DarkOlive 22=DarkBlue 23=DarkPurple
//   24=DarkCranberry
const COLOR_CYCLE = [
  { preset: 0,  hex: "#dc626d" },  // Red
  { preset: 1,  hex: "#e8825d" },  // Orange
  { preset: 2,  hex: "#bc8f6f" },  // Brown
  { preset: 3,  hex: "#fdee65" },  // Yellow
  { preset: 4,  hex: "#52ce90" },  // Green
  { preset: 5,  hex: "#57d2da" },  // Teal
  { preset: 6,  hex: "#b6d767" },  // Olive
  { preset: 7,  hex: "#5ca9e5" },  // Blue
  { preset: 8,  hex: "#b1aaeb" },  // Purple
  { preset: 9,  hex: "#ee5fb7" },  // Cranberry
  { preset: 10, hex: "#4497a9" },  // Steel
  { preset: 11, hex: "#4bb4b7" },  // DarkSteel
  { preset: 12, hex: "#9fadb1" },  // Gray
  { preset: 13, hex: "#8f8f8f" },  // DarkGray
  { preset: 14, hex: "#474747" },  // Black
  { preset: 15, hex: "#ac4e5e" },  // DarkRed
  { preset: 16, hex: "#df8e64" },  // DarkOrange
  { preset: 17, hex: "#c8956c" },  // DarkBrown
  { preset: 18, hex: "#dac257" },  // DarkYellow
  { preset: 19, hex: "#4ca64c" },  // DarkGreen
  { preset: 20, hex: "#2d6f8f" },  // DarkTeal
  { preset: 21, hex: "#85b44c" },  // DarkOlive
  { preset: 22, hex: "#4179a3" },  // DarkBlue
  { preset: 23, hex: "#a589cb" },  // DarkPurple
  { preset: 24, hex: "#c34e98" },  // DarkCranberry
];

// ── State ──────────────────────────────────────────────────────────────────
let allCategories = [];
let appliedNames  = new Set();
let searchTerm    = "";
let busy          = false;
let ctxCat        = null;
let renamingCat   = null;

// ── DOM refs ───────────────────────────────────────────────────────────────
const $  = id => document.getElementById(id);
const loading        = $("loading");
const noItem         = $("noItem");
const mainContent    = $("mainContent");
const searchInput    = $("searchInput");
const clearBtn       = $("clearBtn");
const createHint     = $("createHint");
const createHintName = $("createHintName");
const epochNote      = $("epochNote");
const listLabel      = $("listLabel");
const categoryList   = $("categoryList");
const emptyState     = $("emptyState");
const createBtn      = $("createBtn");
const createBtnLabel = $("createBtnLabel");
const refreshBtn     = $("refreshBtn");
const ctxMenu        = $("ctxMenu");
const ctxFind        = $("ctxFind");
const ctxRename      = $("ctxRename");
const ctxInactive    = $("ctxInactive");
const statusEl       = $("status");
const statusIcon     = $("statusIcon");
const statusText     = $("statusText");

// ── Initialise ─────────────────────────────────────────────────────────────
Office.onReady(info => {
  if (info.host !== Office.HostType.Outlook) return;
  init();
});

async function init() {
  // Refresh the pinned pane whenever the user changes their email selection
  Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, onItemChanged, () => {});

  showLoading(true);
  try {
    await loadAll();
    showLoading(false);
  } catch (e) {
    showLoading(false);
    showStatus("error", "Could not load categories: " + e.message);
  }
}

// ── Fires when the selected email changes (keeps pinned pane in sync) ──────
function onItemChanged() {
  appliedNames = new Set();
  busy = false;
  renamingCat = null;
  ctxCat = null;
  ctxMenu.classList.remove("visible");
  searchTerm = "";
  searchInput.value = "";
  clearBtn.classList.remove("visible");
  statusEl.className = "";
  clearTimeout(statusTimer);

  if (!Office.context.mailbox.item) {
    mainContent.classList.remove("visible");
    noItem.classList.add("visible");
    return;
  }

  noItem.classList.remove("visible");

  // Master list doesn't change between emails — skip re-fetching it.
  // Just reload the applied state for the new item and re-render in place.
  if (allCategories.length > 0) {
    Office.context.mailbox.item.categories.getAsync(itemResult => {
      if (itemResult.status === Office.AsyncResultStatus.Succeeded) {
        appliedNames = new Set((itemResult.value || []).map(c => c.displayName));
      }
      render();
    });
    return;
  }

  // First load — master list not yet cached, fetch everything.
  showLoading(true);
  loadAll()
    .then(() => showLoading(false))
    .catch(e => {
      showLoading(false);
      showStatus("error", "Could not load: " + e.message);
    });
}

// ── Load master list + currently applied ───────────────────────────────────
function loadAll() {
  return new Promise((resolve, reject) => {
    if (!Office.context.mailbox.item) {
      noItem.classList.add("visible");
      resolve();
      return;
    }

    Office.context.mailbox.masterCategories.getAsync(masterResult => {
      if (masterResult.status !== Office.AsyncResultStatus.Succeeded) {
        return reject(new Error(masterResult.error.message));
      }

      const raw = masterResult.value || [];
      allCategories = raw.map(c => ({
        displayName: c.displayName,
        color: c.color,
        colorHex: colorPresetToHex(c.color),
      }));
      allCategories.sort((a, b) => a.displayName.localeCompare(b.displayName));

      Office.context.mailbox.item.categories.getAsync(itemResult => {
        if (itemResult.status === Office.AsyncResultStatus.Succeeded) {
          appliedNames = new Set((itemResult.value || []).map(c => c.displayName));
        } else {
          appliedNames = new Set();
        }
        mainContent.classList.add("visible");
        render();
        resolve();
      });
    });
  });
}

// ── Render ────────────────────────────────────────────────────────────────
function render() {
  const term = searchTerm.trim();
  const lower = term.toLowerCase();

  const allActive = allCategories.filter(c => !c.displayName.startsWith("ω."));
  const filtered = term
    ? allActive.filter(c => c.displayName.toLowerCase().includes(lower))
    : allActive;

  // List
  categoryList.innerHTML = "";
  if (filtered.length === 0 && term) {
    emptyState.classList.add("visible");
    listLabel.style.display = "none";
  } else {
    emptyState.classList.remove("visible");
    listLabel.style.display = "block";
    listLabel.textContent = term
      ? `Matching projects (${filtered.length})`
      : `All projects`;

    filtered.forEach(cat => {
      const isApplied = appliedNames.has(cat.displayName);
      const li = document.createElement("li");
      li.className = "cat-item" + (isApplied ? " applied" : "");
      li.setAttribute("role", "option");
      li.setAttribute("aria-selected", isApplied ? "true" : "false");

      if (renamingCat && renamingCat.displayName === cat.displayName) {
        li.innerHTML =
          `<span class="cat-dot" style="background:${cat.colorHex}"></span>
           <input class="rename-input" value="${escHtml(friendlyName(cat.displayName))}" />
           <button class="rename-confirm" title="Confirm">✓</button>
           <button class="rename-cancel" title="Cancel">✗</button>`;
        const inp = li.querySelector(".rename-input");
        const ok  = li.querySelector(".rename-confirm");
        const cnl = li.querySelector(".rename-cancel");
        setTimeout(() => { inp.focus(); inp.select(); }, 0);
        inp.addEventListener("keydown", e => {
          if (e.key === "Enter")  confirmRename(cat, inp.value.trim());
          if (e.key === "Escape") cancelRename();
        });
        ok.addEventListener("click",  e => { e.stopPropagation(); confirmRename(cat, inp.value.trim()); });
        cnl.addEventListener("click", e => { e.stopPropagation(); cancelRename(); });
        li.addEventListener("click", e => e.stopPropagation());
      } else {
        li.innerHTML =
          `<span class="cat-dot" style="background:${cat.colorHex}"></span>
           <span class="cat-name" title="${escHtml(cat.displayName)}">${escHtml(friendlyName(cat.displayName))}</span>
           ${isApplied
             ? `<span class="applied-badge">✓ Applied</span><span class="remove-hint">Remove</span>`
             : ""}`;
        li.addEventListener("click", () => toggleCategory(cat, isApplied));
        li.addEventListener("contextmenu", e => { e.preventDefault(); showCtxMenu(e.clientX, e.clientY, cat); });
      }

      categoryList.appendChild(li);
    });
  }

  // Create new hint & button
  const exactMatch = allActive.some(
    c => c.displayName.toLowerCase() === lower || friendlyName(c.displayName).toLowerCase() === lower
  );
  const showCreate = term && !exactMatch;

  if (showCreate) {
    const epoch = Math.floor(Date.now() / 1000);
    const storedName = term + "_" + epoch;
    createHintName.textContent = term;
    epochNote.textContent = `Stored as: ${storedName}`;
    createHint.classList.add("visible");
    createHint.dataset.name = storedName;
    createBtnLabel.textContent = `Create & apply "${term}"`;
    createBtn.classList.add("visible");
  } else {
    createHint.classList.remove("visible");
    createBtn.classList.remove("visible");
  }
}

// ── Toggle (apply / remove) an existing category ──────────────────────────
async function toggleCategory(cat, isApplied) {
  if (busy) return;
  busy = true;
  showStatus("info",
    isApplied ? "Removing…" : "Applying…",
    `<path d="M8 1a7 7 0 100 14A7 7 0 008 1zm0 1a6 6 0 110 12A6 6 0 018 2z"/>`
  );

  const op = isApplied
    ? () => removeFromItem([cat.displayName])
    : () => addToItem([cat.displayName]);

  try {
    await op();
    if (isApplied) {
      appliedNames.delete(cat.displayName);
      showStatus("success", `Removed from "${friendlyName(cat.displayName)}"`, checkIcon());
    } else {
      appliedNames.add(cat.displayName);
      showStatus("success", `Applied to "${friendlyName(cat.displayName)}"`, checkIcon());
    }
    render();
    clearStatusAfter(3000);
  } catch (e) {
    showStatus("error", e.message, crossIcon());
  } finally {
    busy = false;
  }
}

// ── Create new category & apply ───────────────────────────────────────────
createBtn.addEventListener("click", async () => {
  if (busy) return;
  const storedName = createHint.dataset.name;
  const displayTerm = searchInput.value.trim();
  if (!storedName) return;

  busy = true;
  createBtn.disabled = true;
  showStatus("info", `Creating project "${displayTerm}"…`);

  const colorIndex = allCategories.length % COLOR_CYCLE.length;
  const colorEntry = COLOR_CYCLE[colorIndex];

  try {
    await addToMaster([{ displayName: storedName, color: Office.MailboxEnums.CategoryColor["Preset" + colorEntry.preset] }]);
    await addToItem([storedName]);

    allCategories.push({
      displayName: storedName,
      color: Office.MailboxEnums.CategoryColor["Preset" + colorEntry.preset],
      colorHex: colorEntry.hex,
    });
    allCategories.sort((a, b) => a.displayName.localeCompare(b.displayName));
    appliedNames.add(storedName);

    searchInput.value = "";
    searchTerm = "";
    clearBtn.classList.remove("visible");

    showStatus("success", `Created & applied "${displayTerm}"`, checkIcon());
    render();
    clearStatusAfter(3500);
  } catch (e) {
    showStatus("error", "Failed: " + e.message, crossIcon());
  } finally {
    busy = false;
    createBtn.disabled = false;
  }
});

// ── Search input events ────────────────────────────────────────────────────
searchInput.addEventListener("input", () => {
  searchTerm = searchInput.value;
  clearBtn.classList.toggle("visible", searchTerm.length > 0);
  render();
});

searchInput.addEventListener("keydown", e => {
  if (e.key === "Enter" && createBtn.classList.contains("visible") && !busy) {
    createBtn.click();
  }
  if (e.key === "Escape") {
    searchInput.value = "";
    searchTerm = "";
    clearBtn.classList.remove("visible");
    render();
  }
});

clearBtn.addEventListener("click", () => {
  searchInput.value = "";
  searchTerm = "";
  clearBtn.classList.remove("visible");
  searchInput.focus();
  render();
});

refreshBtn.addEventListener("click", async () => {
  if (busy) return;
  refreshBtn.disabled = true;
  renamingCat = null;
  ctxCat = null;
  ctxMenu.classList.remove("visible");
  mainContent.classList.remove("visible");
  allCategories = [];
  appliedNames = new Set();
  searchTerm = "";
  searchInput.value = "";
  clearBtn.classList.remove("visible");
  statusEl.className = "";
  clearTimeout(statusTimer);
  showLoading(true);
  try {
    await loadAll();
  } catch (e) {
    showStatus("error", "Refresh failed: " + e.message);
  } finally {
    showLoading(false);
    refreshBtn.disabled = false;
  }
});

// ── Office API helpers (promisified) ──────────────────────────────────────
function addToMaster(categories) {
  return new Promise((res, rej) => {
    Office.context.mailbox.masterCategories.addAsync(categories, r => {
      // DuplicateCategory is fine — just means it already exists
      if (r.status === Office.AsyncResultStatus.Succeeded ||
          r.error?.code === 9014 /* DuplicateCategory */) {
        res();
      } else {
        rej(new Error(r.error?.message || "Failed to create category"));
      }
    });
  });
}

function addToItem(names) {
  return new Promise((res, rej) => {
    Office.context.mailbox.item.categories.addAsync(names, r => {
      if (r.status === Office.AsyncResultStatus.Succeeded) res();
      else rej(new Error(r.error?.message || "Failed to apply category"));
    });
  });
}

function removeFromItem(names) {
  return new Promise((res, rej) => {
    Office.context.mailbox.item.categories.removeAsync(names, r => {
      if (r.status === Office.AsyncResultStatus.Succeeded) res();
      else rej(new Error(r.error?.message || "Failed to remove category"));
    });
  });
}

// ── UI helpers ────────────────────────────────────────────────────────────
function showLoading(on) {
  loading.style.display = on ? "block" : "none";
}

let statusTimer = null;
function showStatus(type, msg, iconPath) {
  clearTimeout(statusTimer);
  statusEl.className = "visible " + type;
  statusText.textContent = msg;
  statusIcon.innerHTML = iconPath || "";
}
function clearStatusAfter(ms) {
  statusTimer = setTimeout(() => {
    statusEl.className = "";
  }, ms);
}

function checkIcon() {
  return `<path d="M13.854 3.646a.5.5 0 010 .708l-7 7a.5.5 0 01-.708 0l-3.5-3.5a.5.5 0 11.708-.708L6.5 10.293l6.646-6.647a.5.5 0 01.708 0z"/>`;
}
function crossIcon() {
  return `<path d="M2.146 2.854a.5.5 0 11.708-.708L8 7.293l5.146-5.147a.5.5 0 01.708.708L8.707 8l5.147 5.146a.5.5 0 01-.708.708L8 8.707l-5.146 5.147a.5.5 0 01-.708-.708L7.293 8z"/>`;
}

// Strip the _epoch suffix for display purposes
function friendlyName(name) {
  return name.replace(/_\d{9,11}$/, "");
}

function escHtml(s) {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

// Map Office CategoryColor preset to a display hex.
// The Office API returns string enum values ("Preset0", "Preset1", …);
// normalise to integer before looking up.
function colorPresetToHex(preset) {
  const n = typeof preset === "string"
    ? parseInt(preset.replace(/\D/g, ""), 10)
    : preset;
  const found = COLOR_CYCLE.find(c => c.preset === n);
  return found ? found.hex : "#767676";
}

// ── Context menu ──────────────────────────────────────────────────────────
function showCtxMenu(x, y, cat) {
  ctxCat = cat;
  ctxMenu.style.left = x + "px";
  ctxMenu.style.top  = y + "px";
  ctxMenu.classList.add("visible");
  requestAnimationFrame(() => {
    const r = ctxMenu.getBoundingClientRect();
    if (r.right  > window.innerWidth)  ctxMenu.style.left = (x - r.width)  + "px";
    if (r.bottom > window.innerHeight) ctxMenu.style.top  = (y - r.height) + "px";
  });
}

document.addEventListener("click", e => {
  if (!ctxMenu.contains(e.target)) {
    ctxMenu.classList.remove("visible");
    ctxCat = null;
  }
});

ctxFind.addEventListener("click", () => {
  if (!ctxCat) return;
  const name  = friendlyName(ctxCat.displayName);
  const query = `category:"${name}"`;
  ctxCat = null;
  ctxMenu.classList.remove("visible");

  // Copy the KQL search query to clipboard so the user can paste it into
  // the Outlook search bar (no Office.js API exists to trigger search directly).
  let copied = false;
  try {
    const ta = document.createElement("textarea");
    ta.value = query;
    ta.style.cssText = "position:fixed;top:-9999px;left:-9999px";
    document.body.appendChild(ta);
    ta.select();
    copied = document.execCommand("copy");
    document.body.removeChild(ta);
  } catch (_) {}

  showStatus(
    "info",
    copied ? `Copied — paste into Outlook search bar` : `Search: ${query}`,
    `<path d="M11.742 10.344a6.5 6.5 0 10-1.397 1.398h-.001c.03.04.062.078.098.115l3.85 3.85a1 1 0 001.415-1.414l-3.85-3.85a1.007 1.007 0 00-.115-.099zm-5.242 1.156a5.5 5.5 0 110-11 5.5 5.5 0 010 11z"/>`
  );
  clearStatusAfter(5000);
});

ctxRename.addEventListener("click", () => {
  if (!ctxCat) return;
  renamingCat = ctxCat;
  ctxCat = null;
  ctxMenu.classList.remove("visible");
  render();
});

ctxInactive.addEventListener("click", async () => {
  if (!ctxCat) return;
  const cat = ctxCat;
  ctxCat = null;
  ctxMenu.classList.remove("visible");
  await markInactive(cat);
});

// ── Inline rename ─────────────────────────────────────────────────────────
function cancelRename() {
  renamingCat = null;
  render();
}

async function confirmRename(cat, newFriendlyName) {
  if (!newFriendlyName) { cancelRename(); return; }
  if (newFriendlyName === friendlyName(cat.displayName)) { cancelRename(); return; }

  const allActive = allCategories.filter(c => !c.displayName.startsWith("ω."));
  if (allActive.some(c => c !== cat && friendlyName(c.displayName).toLowerCase() === newFriendlyName.toLowerCase())) {
    showStatus("error", `"${newFriendlyName}" already exists`, crossIcon());
    return;
  }

  renamingCat = null;
  busy = true;
  const epoch = Math.floor(Date.now() / 1000);
  const newStoredName = newFriendlyName + "_" + epoch;

  try {
    await addToMaster([{ displayName: newStoredName, color: cat.color }]);
    await removeMaster([cat.displayName]);

    if (appliedNames.has(cat.displayName)) {
      await removeFromItem([cat.displayName]);
      await addToItem([newStoredName]);
      appliedNames.delete(cat.displayName);
      appliedNames.add(newStoredName);
    }

    cat.displayName = newStoredName;
    allCategories.sort((a, b) => a.displayName.localeCompare(b.displayName));
    showStatus("success", `Renamed to "${newFriendlyName}"`, checkIcon());
    clearStatusAfter(3000);
  } catch (e) {
    showStatus("error", "Rename failed: " + e.message, crossIcon());
  } finally {
    busy = false;
    render();
  }
}

// ── Mark inactive ─────────────────────────────────────────────────────────
async function markInactive(cat) {
  if (busy) return;
  busy = true;
  const oldName      = cat.displayName;
  const newStoredName = "ω." + oldName;

  try {
    await addToMaster([{ displayName: newStoredName, color: cat.color }]);
    await removeMaster([oldName]);

    if (appliedNames.has(oldName)) {
      await removeFromItem([oldName]);
      await addToItem([newStoredName]);
      appliedNames.delete(oldName);
      appliedNames.add(newStoredName);
    }

    cat.displayName = newStoredName;
    showStatus("success", `"${friendlyName(oldName)}" marked inactive`, checkIcon());
    clearStatusAfter(3000);
  } catch (e) {
    showStatus("error", "Failed: " + e.message, crossIcon());
  } finally {
    busy = false;
    render();
  }
}

// ── Remove from master list ───────────────────────────────────────────────
function removeMaster(names) {
  return new Promise((res, rej) => {
    Office.context.mailbox.masterCategories.removeAsync(names, r => {
      if (r.status === Office.AsyncResultStatus.Succeeded) res();
      else rej(new Error(r.error?.message || "Failed to remove category"));
    });
  });
}
