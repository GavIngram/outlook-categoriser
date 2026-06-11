# Project Categoriser — Outlook Add-in

Quickly assign emails to project categories from a single **Categorise** button in the Outlook ribbon. Works in **Outlook Classic (Windows)** and **Outlook on the Web (OWA)**.

---

## How it works

| Action | What happens |
|--------|-------------|
| Click **Categorise** in the ribbon | Taskpane opens showing all your project categories |
| Type to search | Live-filters the list instantly |
| Click an existing project | Applies (or removes) it from the email immediately |
| Type a name with no match → press Enter or click Create | Creates the category in your master list and applies it in one step |
| Category stored name | `Project Name_<unix epoch>` e.g. `Acme Rollout_1718067200` |
| Display name shown | `Project Name` (epoch suffix stripped in the UI) |

The epoch suffix makes every project name unique even if you reuse names across years, and lets you sort by creation time if needed.

---

## Installing the add-in

The add-in is hosted at **https://gavingram.github.io/outlook-categoriser/** — no local server required.

### Outlook on the Web (OWA)

1. Go to [outlook.office.com](https://outlook.office.com)
2. Open any email
3. Click the **Apps** icon (grid/puzzle piece) in the toolbar
4. Click **Add apps** → **Upload a custom app** (or navigate to [aka.ms/olksideload](https://aka.ms/olksideload))
5. Upload `manifest.xml` from this repository
6. The **Categorise** button will appear when you open an email

### Outlook Classic (Windows)

1. Open Outlook
2. **File** → **Manage Add-ins** — this opens OWA in your browser
3. Follow the OWA steps above — the add-in installs across both clients automatically

---

## Deploying for your whole organisation

For company-wide rollout so everyone gets the button automatically:

1. Go to [admin.microsoft.com](https://admin.microsoft.com) → **Settings** → **Integrated apps** → **Upload custom apps**
2. Upload `manifest.xml`
3. Assign to specific users, groups, or the whole organisation
4. Users get the button automatically — no action needed on their end

---

## File structure

```
outlook-categoriser/
├── manifest.xml          ← Add-in manifest (upload this to install)
├── assets/               ← Icons (16, 32, 64, 80, 128 px)
└── src/
    ├── taskpane.html     ← Add-in UI
    ├── taskpane.js       ← All logic (search, create, apply, remove)
    └── commands.html     ← Required shim for manifest FunctionFile
```

---

## Browsing your projects after categorising

**Outlook Classic:**
- View tab → **Filter Email** → **By Category** → pick your project
- Or: right-click any category column header → Group by Category

**OWA:**
- Filter icon at top of inbox → **Categories** → select project

**Search:**
- `category:"Acme Rollout"` in the search box (use the stored name with epoch, or just part of the name)

---

## Notes & limitations

- Requires **Mailbox API 1.8** — supported in Outlook on Microsoft 365 (desktop and web). Not supported in Outlook 2016 or older.
- The button appears in **read mode only** (when viewing an email). You cannot apply categories while composing — this is an Office.js platform restriction for OWA.
- In **delegate mailbox** scenarios, the delegate can apply existing categories but cannot create new ones (Microsoft platform restriction).
- The epoch suffix in stored names is permanent. The UI always strips it for display. If you need the raw name (e.g. for search), hover over a category in the list — it shows the full stored name in the title attribute.
