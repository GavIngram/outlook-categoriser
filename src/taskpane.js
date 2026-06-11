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
let allCategories  = [];
let appliedNames   = new Set();
let searchTerm     = "";
let busy           = false;
let isMultiSelect  = false;
let multiSelectItems = []; // SelectedItemDetails[] from getSelectedItemsAsync

// ── DOM refs ───────────────────────────────────────────────────────────────
const $  = id => document.getElementById(id);
const loading           = $("loading");
const noItem            = $("noItem");
const mainContent       = $("mainContent");
const searchInput       = $("searchInput");
const clearBtn          = $("clearBtn");
const createHint        = $("createHint");
const createHintName    = $("createHintName");
const epochNote         = $("epochNote");
const listLabel         = $("listLabel");
const categoryList      = $("categoryList");
const emptyState        = $("emptyState");
const createBtn         = $("createBtn");
const createBtnLabel    = $("createBtnLabel");
const categoryCount     = $("categoryCount");
const statusEl          = $("status");
const statusIcon        = $("statusIcon");
const statusText        = $("statusText");
const multiBanner       = $("multiBanner");
const multiCountEl      = $("multiCount");

// ── Initialise ─────────────────────────────────────────────────────────────
Office.onReady(info => {
  if (info.host !== Office.HostType.Outlook) return;
  init();
});

async function init() {
  // Register ItemChanged so the pinned pane refreshes on selection change
  Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, onItemChanged, () => {});

  // Also register SelectedItemsChanged if available (Mailbox 1.10+)
  try {
    if (Office.EventType.SelectedItemsChanged !== undefined) {
      Office.context.mailbox.addHandlerAsync(Office.EventType.SelectedItemsChanged, onItemChanged, () => {});
    }
  } catch (e) { /* not available in this client */ }

  showLoading(true);
  try {
    await loadAll();
    showLoading(false);
  } catch (e) {
    showLoading(false);
    showStatus("error", "Could not load categories: " + e.message);
  }
}

// ── Selection change handler (fires when user switches email / multi-selects) ─
function onItemChanged() {
  isMultiSelect = false;
  multiSelectItems = [];
  appliedNames = new Set();
  searchTerm = "";
  searchInput.value = "";
  clearBtn.classList.remove("visible");
  statusEl.className = "";
  clearTimeout(statusTimer);

  noItem.classList.remove("visible");
  mainContent.classList.remove("visible");

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
      // item is null — could be multi-select or nothing selected
      if (typeof Office.context.mailbox.getSelectedItemsAsync === "function") {
        Office.context.mailbox.getSelectedItemsAsync(selResult => {
          const items = (selResult.status === Office.AsyncResultStatus.Succeeded)
            ? (selResult.value || [])
            : [];

          if (items.length > 0) {
            isMultiSelect = true;
            multiSelectItems = items;
            loadMasterCategories(resolve, reject);
          } else {
            noItem.classList.add("visible");
            resolve();
          }
        });
      } else {
        noItem.classList.add("visible");
        resolve();
      }
      return;
    }

    isMultiSelect = false;
    multiSelectItems = [];
    loadMasterCategories(resolve, reject);
  });
}

function loadMasterCategories(resolve, reject) {
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

    if (isMultiSelect) {
      // No per-item applied state in multi-select mode
      appliedNames = new Set();
      mainContent.classList.add("visible");
      render();
      resolve();
    } else {
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
    }
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

  // Multi-select banner
  if (isMultiSelect) {
    multiBanner.classList.add("visible");
    multiCountEl.textContent =
      `${multiSelectItems.length} email${multiSelectItems.length === 1 ? "" : "s"} selected — click to apply to all`;
  } else {
    multiBanner.classList.remove("visible");
  }

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
      const isApplied = !isMultiSelect && appliedNames.has(cat.displayName);
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
    createBtnLabel.textContent = isMultiSelect
      ? `Create & apply "${term}" to all`
      : `Create & apply "${term}"`;
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

  if (isMultiSelect) {
    const n = multiSelectItems.length;
    showStatus("info", `Applying to ${n} email${n === 1 ? "" : "s"}…`,
      `<path d="M8 1a7 7 0 100 14A7 7 0 008 1zm0 1a6 6 0 110 12A6 6 0 018 2z"/>`);
    try {
      await applyCategoriesToItems([cat.displayName]);
      showStatus("success",
        `Applied "${friendlyName(cat.displayName)}" to ${n} email${n === 1 ? "" : "s"}`,
        checkIcon());
      clearStatusAfter(3000);
    } catch (e) {
      showStatus("error", e.message, crossIcon());
    } finally {
      busy = false;
    }
    return;
  }

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

  const n = multiSelectItems.length;
  showStatus("info", isMultiSelect
    ? `Creating project "${displayTerm}" and applying to ${n} email${n === 1 ? "" : "s"}…`
    : `Creating project "${displayTerm}"…`);

  const colorIndex = allCategories.length % COLOR_CYCLE.length;
  const colorEntry = COLOR_CYCLE[colorIndex];

  try {
    await addToMaster([{
      displayName: storedName,
      color: Office.MailboxEnums.CategoryColor["Preset" + colorEntry.preset],
    }]);

    if (isMultiSelect) {
      await applyCategoriesToItems([storedName]);
    } else {
      await addToItem([storedName]);
    }

    allCategories.push({
      displayName: storedName,
      color: colorEntry.preset,
      colorHex: colorEntry.hex,
    });
    allCategories.sort((a, b) => a.displayName.localeCompare(b.displayName));

    if (!isMultiSelect) appliedNames.add(storedName);

    searchInput.value = "";
    searchTerm = "";
    clearBtn.classList.remove("visible");

    showStatus("success",
      isMultiSelect
        ? `Created & applied "${displayTerm}" to ${n} email${n === 1 ? "" : "s"}`
        : `Created & applied "${displayTerm}"`,
      checkIcon());
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

// ── EWS helpers for multi-select ───────────────────────────────────────────
const EWS_T_NS = "http://schemas.microsoft.com/exchange/services/2006/types";

async function applyCategoriesToItems(categoryNames) {
  const ids = multiSelectItems.map(i => i.itemId);

  // GetItem to read existing categories (so we merge rather than replace)
  const getXml = buildGetItemSoap(ids);
  const getResponse = await makeEwsRequest(getXml);
  const existing = parseGetItemResponse(getResponse);

  // Build one UpdateItem with all changes
  const changes = ids.map(id => {
    const found = existing.find(e => e.id === id) || { id, changeKey: "", categories: [] };
    const merged = [...new Set([...found.categories, ...categoryNames])];
    return { id: found.id || id, changeKey: found.changeKey, categories: merged };
  });

  const updateXml = buildUpdateItemSoap(changes);
  await makeEwsRequest(updateXml);
}

function buildGetItemSoap(ids) {
  const itemIds = ids.map(id => `<t:ItemId Id="${escXml(id)}"/>`).join("\n        ");
  return `<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
  <soap:Header>
    <t:RequestServerVersion Version="Exchange2013_SP1"/>
  </soap:Header>
  <soap:Body>
    <m:GetItem>
      <m:ItemShape>
        <t:BaseShape>IdOnly</t:BaseShape>
        <t:AdditionalProperties>
          <t:FieldURI FieldURI="item:Categories"/>
        </t:AdditionalProperties>
      </m:ItemShape>
      <m:ItemIds>
        ${itemIds}
      </m:ItemIds>
    </m:GetItem>
  </soap:Body>
</soap:Envelope>`;
}

function buildUpdateItemSoap(changes) {
  const itemChanges = changes.map(({ id, changeKey, categories }) => {
    const ckAttr = changeKey ? ` ChangeKey="${escXml(changeKey)}"` : "";
    const catStrings = categories
      .map(c => `<t:String>${escXml(c)}</t:String>`)
      .join("\n                  ");
    return `<t:ItemChange>
          <t:ItemId Id="${escXml(id)}"${ckAttr}/>
          <t:Updates>
            <t:SetItemField>
              <t:FieldURI FieldURI="item:Categories"/>
              <t:Message>
                <t:Categories>
                  ${catStrings}
                </t:Categories>
              </t:Message>
            </t:SetItemField>
          </t:Updates>
        </t:ItemChange>`;
  }).join("\n        ");

  return `<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
  <soap:Header>
    <t:RequestServerVersion Version="Exchange2013_SP1"/>
  </soap:Header>
  <soap:Body>
    <m:UpdateItem MessageDisposition="SaveOnly" ConflictResolution="AlwaysOverwrite">
      <m:ItemChanges>
        ${itemChanges}
      </m:ItemChanges>
    </m:UpdateItem>
  </soap:Body>
</soap:Envelope>`;
}

function parseGetItemResponse(xml) {
  const doc = new DOMParser().parseFromString(xml, "text/xml");
  const messages = doc.getElementsByTagNameNS(EWS_T_NS, "Message");
  const items = messages.length ? messages : doc.getElementsByTagNameNS(EWS_T_NS, "Item");
  return Array.from(items).map(item => {
    const idEl = item.getElementsByTagNameNS(EWS_T_NS, "ItemId")[0];
    const id = idEl?.getAttribute("Id") || "";
    const changeKey = idEl?.getAttribute("ChangeKey") || "";
    const catContainer = item.getElementsByTagNameNS(EWS_T_NS, "Categories")[0];
    const categories = catContainer
      ? Array.from(catContainer.getElementsByTagNameNS(EWS_T_NS, "String"))
          .map(el => el.textContent.trim())
          .filter(Boolean)
      : [];
    return { id, changeKey, categories };
  });
}

function makeEwsRequest(soapRequest) {
  return new Promise((res, rej) => {
    Office.context.mailbox.makeEwsRequestAsync(soapRequest, r => {
      if (r.status === Office.AsyncResultStatus.Succeeded) res(r.value);
      else rej(new Error(r.error?.message || "EWS request failed"));
    });
  });
}

// ── Office API helpers (promisified) ──────────────────────────────────────
function addToMaster(categories) {
  return new Promise((res, rej) => {
    Office.context.mailbox.masterCategories.addAsync(categories, r => {
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

function escXml(s) {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function colorPresetToHex(preset) {
  const found = COLOR_CYCLE.find(c => c.preset === preset);
  return found ? found.hex : "#767676";
}
