"use strict";

// ── Colour palette cycling for new categories ──────────────────────────────
const COLOR_CYCLE = [
  { preset: 0,  hex: "#e74856" },
  { preset: 1,  hex: "#ff8c00" },
  { preset: 2,  hex: "#f7630c" },
  { preset: 3,  hex: "#ca5010" },
  { preset: 4,  hex: "#da3b01" },
  { preset: 5,  hex: "#ef6950" },
  { preset: 6,  hex: "#d13438" },
  { preset: 7,  hex: "#ff4343" },
  { preset: 8,  hex: "#ffb900" },
  { preset: 9,  hex: "#c0cb00" },
  { preset: 10, hex: "#10893e" },
  { preset: 11, hex: "#00b7c3" },
  { preset: 12, hex: "#0078d4" },
  { preset: 13, hex: "#4a4af4" },
  { preset: 14, hex: "#881798" },
  { preset: 15, hex: "#744da9" },
  { preset: 16, hex: "#8e8cd8" },
  { preset: 17, hex: "#038387" },
  { preset: 18, hex: "#107c10" },
  { preset: 19, hex: "#486860" },
  { preset: 20, hex: "#525e54" },
  { preset: 21, hex: "#7e735f" },
  { preset: 22, hex: "#a0522d" },
  { preset: 23, hex: "#767676" },
  { preset: 24, hex: "#4c4c4c" },
];

// ── State ──────────────────────────────────────────────────────────────────
let allCategories = [];
let appliedNames  = new Set();
let searchTerm    = "";
let busy          = false;

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
const categoryCount  = $("categoryCount");
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

  const filtered = term
    ? allCategories.filter(c => c.displayName.toLowerCase().includes(lower))
    : allCategories;

  categoryCount.textContent = allCategories.length
    ? allCategories.length + " project" + (allCategories.length === 1 ? "" : "s")
    : "";

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
      li.innerHTML =
        `<span class="cat-dot" style="background:${cat.colorHex}"></span>
         <span class="cat-name" title="${escHtml(cat.displayName)}">${escHtml(friendlyName(cat.displayName))}</span>
         ${isApplied
           ? `<span class="applied-badge">✓ Applied</span><span class="remove-hint">Remove</span>`
           : ""}`;
      li.addEventListener("click", () => toggleCategory(cat, isApplied));
      categoryList.appendChild(li);
    });
  }

  // Create new hint & button
  const exactMatch = allCategories.some(
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
      color: colorEntry.preset,
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

// Map Office CategoryColor preset integer to a display hex
function colorPresetToHex(preset) {
  const found = COLOR_CYCLE.find(c => c.preset === preset);
  return found ? found.hex : "#767676";
}
