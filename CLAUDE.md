# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Hosting

The add-in is hosted as static files on **GitHub Pages** at `https://gavingram.github.io/outlook-categoriser/`. There is no build step, no server, and no Node.js dependency. All source files are served directly by GitHub Pages.

To test changes: push to the `main` branch and reload the taskpane in Outlook (or OWA).

## Architecture

Vanilla JS Office.js Outlook add-in. No framework, no bundler.

**Entry points defined in `manifest.xml`:**
- `src/taskpane.html` + `src/taskpane.js` — the main UI loaded when the user clicks "Categorise"
- `src/commands.html` — required empty shim for the `FunctionFile` manifest entry (no logic)

**All add-in logic lives in `src/taskpane.js`.** It calls the Office.js API directly:
- `Office.context.mailbox.masterCategories` — read/write the user's full category list
- `Office.context.mailbox.item.categories` — apply/remove categories on the current email
- All async Office calls are wrapped in Promises (`addToMaster`, `addToItem`, `removeFromItem`)

**Category naming convention:** stored names use `DisplayName_<unix epoch>` (e.g. `Acme Rollout_1718067200`). The epoch makes names unique across time. `friendlyName()` strips the suffix for display. The raw stored name is what's passed to the Office API.

**Color assignment:** new categories cycle through `COLOR_CYCLE` (Preset0–Preset24 mapped to `Office.MailboxEnums.CategoryColor`). Index is `allCategories.length % COLOR_CYCLE.length` at creation time.

**State:** four module-level variables — `allCategories`, `appliedNames` (Set), `searchTerm`, `busy`. `render()` is a full re-render from that state.

## Manifest

`manifest.xml` points to `https://gavingram.github.io/outlook-categoriser/`. If the Pages URL ever changes, do a find-and-replace on that base URL throughout the file. Requires Mailbox API 1.8.

To install: sideload `manifest.xml` via OWA (Apps → Upload a custom app), or deploy org-wide via Microsoft 365 Admin Centre (Settings → Integrated apps).
