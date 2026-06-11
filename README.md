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

## Validating the manifest

**Always validate before deploying.** The admin center and the deployment cmdlets reject an invalid manifest with opaque errors — *"Wrong Package. Your package does not match submission type"* or *"Unable to extract Add-In's details"* — that don't say what's actually wrong. Microsoft's official validator does:

```powershell
npx --yes office-addin-manifest validate manifest.xml
```

It runs Microsoft's schema + acceptance checks and reports the real problem with a line number. Needs Node.js installed; `npx` fetches the tool on demand (no install).

**Gotchas this caught here** (both previously blocked all deployment):
- The root element **must** be `xsi:type="MailApp"`. The legacy `xsi:type="MessageReadApp"` is no longer recognised by the validation gateway and causes the "Wrong Package" / "Package Type Not Identified" rejection.
- `<SupportsPinning>` is only legal inside the **VersionOverrides v1.1** `<Action>`, not v1.0.

---

## Deploying for your whole organisation

Company-wide rollout uses the **Centralized Deployment** PowerShell module (the proven working method here). You need a **Global Admin** account.

### One-time setup

```powershell
Install-Module -Name O365CentralizedAddInDeployment   # first time only
Import-Module  -Name O365CentralizedAddInDeployment
Connect-OrganizationAddInService                       # sign in as Global Admin
```

### Deploy to one person (do this first, to test)

```powershell
New-OrganizationAddIn `
  -ManifestPath 'https://gavingram.github.io/outlook-categoriser/manifest.xml' `
  -Locale 'en-AU' `
  -Members 'someone@virtualgraffiti.com.au'
```

`-ManifestPath` accepts the hosted URL (single source of truth) or a local file path. On success it returns the add-in's **ProductId**, which is the same as the manifest `<Id>`:

```
2e7b12fe-72e1-4e3d-b7c0-f39f7fb383c6
```

The button can take up to 24 hours to appear (usually minutes); users may need to restart Outlook.

### Deploy to everyone

If it isn't uploaded yet, upload first (omitting `-Members` assigns it to nobody initially):

```powershell
New-OrganizationAddIn -ManifestPath 'https://gavingram.github.io/outlook-categoriser/manifest.xml' -Locale 'en-AU'
```

Then assign it to the whole organisation:

```powershell
Set-OrganizationAddInAssignments -ProductId 2e7b12fe-72e1-4e3d-b7c0-f39f7fb383c6 -AssignToEveryone $true
```

To scope to specific people/groups instead of everyone:

```powershell
Set-OrganizationAddInAssignments -ProductId 2e7b12fe-72e1-4e3d-b7c0-f39f7fb383c6 -Add -Members 'a@virtualgraffiti.com.au','team@virtualgraffiti.com.au'
```

### GUI alternative

Once the manifest is valid, the admin center also works: [admin.microsoft.com](https://admin.microsoft.com) → **Settings** → **Integrated apps** → **Add-ins** → **Deploy Add-in** → provide the manifest URL.

---

## Updating the manifest and redeploying

1. Edit `manifest.xml`.
2. **Bump the version** — increase `<Version>` (e.g. `1.0.0.0` → `1.0.0.1`). Updates will not propagate otherwise.
3. Validate: `npx --yes office-addin-manifest validate manifest.xml`
4. Commit and push so GitHub Pages serves the new file:
   ```powershell
   git add manifest.xml; git commit -m "Update manifest to vX.Y.Z"; git push
   ```
5. Push the update into the tenant (re-uses the existing ProductId):
   ```powershell
   Set-OrganizationAddIn `
     -ProductId 2e7b12fe-72e1-4e3d-b7c0-f39f7fb383c6 `
     -ManifestPath 'https://gavingram.github.io/outlook-categoriser/manifest.xml' `
     -Locale 'en-AU'
   ```

> Changes to requested **permissions** or **events** require an admin to re-consent before users receive the update.

### Managing / removing the deployment

```powershell
Get-OrganizationAddIn                                                      # list all add-ins + ProductIds
Set-OrganizationAddIn -ProductId 2e7b12fe-72e1-4e3d-b7c0-f39f7fb383c6 -Enabled $false   # disable (keep assignment)
Remove-OrganizationAddIn -ProductId 2e7b12fe-72e1-4e3d-b7c0-f39f7fb383c6   # remove entirely
```

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
