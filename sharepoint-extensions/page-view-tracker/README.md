# Page View Tracker — SPFx Application Customizer

Tracks views on SharePoint modern **Site Pages** (including news) and logs them to a SharePoint list with **traffic source attribution** from URL query parameters such as `?source=viva`.

**Hub site:** `https://mdigital.sharepoint.com/sites/ITOTCommunityHub`

---

## What it does

1. Runs silently on every modern site page in the hub (Application Customizer).
2. Reads `?source=` from the URL (e.g. Viva Engage links use `?source=viva`).
3. Records viewer, page, timestamp, referrer, news flag, and hub section to **Page View Tracking**.
4. Deduplicates repeat views in the same browser session (same page + same source).

---

## Prerequisites

| Requirement | Notes |
|-------------|-------|
| Node.js **18 LTS** or **22 LTS** | SPFx 1.21 does not support Node 24 — use [nvm-windows](https://github.com/coreybutler/nvm-windows) if needed |
| SharePoint Online tenant | Modern pages |
| Tenant **App Catalog** | To deploy the `.sppkg` |
| SharePoint admin | To approve API access if prompted |

---

## Quick start (development)

```powershell
cd "sharepoint-extensions/page-view-tracker"
npm install
gulp trust-dev-cert
gulp serve
```

1. Create the list using [LIST_SETUP.md](./LIST_SETUP.md).
2. When the browser opens, accept the dev certificate warning.
3. Load a hub site page — check the **Page View Tracking** list for a new item.

`config/serve.json` is preconfigured for the IT/OT Community Hub home page.

---

## Build and package

```powershell
npm install
gulp bundle --ship
gulp package-solution --ship
```

Output: `sharepoint-extensions/page-view-tracker/sharepoint/solution/page-view-tracker.sppkg`

---

## Deploy to the hub (all pages)

### Step 1 — Upload to App Catalog

1. Open your **Tenant App Catalog** site.
2. Upload `page-view-tracker.sppkg`.
3. Check **Make this solution available to all sites in the organization** when prompted.
4. Click **Deploy**.

### Step 2 — Register on the hub site

**Option A — SharePoint Admin Center (recommended)**

1. SharePoint Admin Center → **Settings** → **Custom apps** (or **Apps** → manage apps).
2. Find **page-view-tracker-client-side-solution**.
3. Add an **Application Customizer** deployment scoped to:

   `https://mdigital.sharepoint.com/sites/ITOTCommunityHub`

**Option B — PnP PowerShell**

```powershell
Connect-PnPOnline -Url "https://mdigital.sharepoint.com/sites/ITOTCommunityHub" -Interactive

Add-PnPCustomAction `
  -Name "PageViewTracker" `
  -Title "Page View Tracker" `
  -Location "ClientSideExtension.ApplicationCustomizer" `
  -ClientSideComponentId "7c4e8f2a-1b3d-4e5f-9a6b-2c8d0e1f3a5b" `
  -ClientSideComponentProperties '{"trackingListTitle":"Page View Tracking","trackingListSiteUrl":"https://mdigital.sharepoint.com/sites/ITOTCommunityHub","sourceQueryParam":"source","trackNewsOnly":false,"defaultSource":"direct","dedupeInSession":true}'
```

**Option C — Manual site custom action (XML)**

See `sharepoint/assets/elements.xml` for the component ID and default properties.

### Step 3 — Verify

1. Open a news page:  
   `.../SitePages/Your-News.aspx?source=viva`
2. Confirm a new row in **Page View Tracking** with `Source = viva`.

---

## Configuration properties

Set via `ClientSideComponentProperties` when registering the extension:

| Property | Default | Description |
|----------|---------|-------------|
| `trackingListTitle` | `Page View Tracking` | Target list display name |
| `trackingListSiteUrl` | Current web URL | Hub site URL where the list lives |
| `sourceQueryParam` | `source` | Query parameter name for attribution |
| `trackNewsOnly` | `false` | If `true`, only track SharePoint news posts |
| `defaultSource` | `direct` | Value when no query parameter is present |
| `dedupeInSession` | `true` | Skip duplicate views per session |

Example — news only, custom default:

```json
{
  "trackingListTitle": "Page View Tracking",
  "trackingListSiteUrl": "https://mdigital.sharepoint.com/sites/ITOTCommunityHub",
  "sourceQueryParam": "source",
  "trackNewsOnly": true,
  "defaultSource": "organic",
  "dedupeInSession": true
}
```

---

## Hub-associated sites

This package is registered on the **hub site** by default. Associated spoke sites do **not** inherit the custom action automatically.

To track spoke sites too, either:

- Register the same component ID on each associated site, or
- Use tenant-wide deployment with site URL scoping in SharePoint Admin Center.

---

## Component ID

```
7c4e8f2a-1b3d-4e5f-9a6b-2c8d0e1f3a5b
```

Keep this ID stable when updating the package so existing registrations continue to work.

---

## Related workspace assets

| Asset | Purpose |
|-------|---------|
| [LIST_SETUP.md](./LIST_SETUP.md) | SharePoint list columns and permissions |
| `sharepoint-widgets/page-kpi-dashboard/` | Existing lifetime views/viewers dashboard (Power Automate export) |

This extension adds **per-view source attribution** (Viva, newsletter, etc.) that native SharePoint analytics does not expose in the same way.

---

## Uninstall

1. Remove the custom action / Application Customizer registration from the hub site.
2. Remove the app from the site (Site contents → Apps for SharePoint).
3. Optionally delete the app from the tenant catalog.
