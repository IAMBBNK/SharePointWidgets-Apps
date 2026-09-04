# Predictive Maintenance Form — setup

## Lists

| List | URL |
|------|-----|
| **Predictive Maintenance** (target) | [All items](https://mdigital.sharepoint.com/sites/ITOTCommunityHub/Lists/Predictive%20Maintenance/AllItems.aspx) |
| **Site Master Data** (site codes) | `Site Master Data` on the same hub site |

## Required columns (Predictive Maintenance)

Display names must match SharePoint exactly:

| Column | Type | Form |
|--------|------|------|
| Title | Single line of text | Auto-set on save (`{Site ID} — {Production}`) — not shown in UI |
| Site ID | Single line of text | Dropdown from Site Master Data |
| Site Name | Single line of text | Read-only, prefilled from master data |
| Sector | Choice | Read-only, prefilled from master data |
| Production | Single line of text | Text input (required) |
| Technical Contact | Person or Group | Single person picker |
| System Size | Single line of text | Text input |
| Start of Operation | Date and Time | `datetime-local` input |
| Technical Specials | Single line of text | Text input |

**Sector choices** on this list should align with **Sector** values in Site Master Data. If they differ, the form shows a warning and save may return HTTP 400.

## Access

- Only users listed in `ADMIN_USERS` in `predictiveMaintenanceForm.js` can open the form.
- Client-side check is UX only; list permissions still apply on save (403 if the user cannot contribute/edit).
- Recommended: grant **Contribute** (or Edit) on Predictive Maintenance only to those admins.

Update before deploy:

```javascript
var ADMIN_USERS = [
  "your.email@merckgroup.com",
  "i:0#.f|membership|another.user@merckgroup.com"
];
```

Match is case-insensitive on email, or exact login name.

## Behaviour

- **Entry selector** (admins only): `Create new entry` or any existing Predictive Maintenance item (`Site ID — Production`, or Title / item id fallback).
- **Create new:** empty manual fields; choose Site ID → Site Name and Sector prefilled read-only from Site Master Data.
- **Edit:** loads the selected item; Site ID can be changed (prefill updates).
- **Save:** POST for new items, MERGE for updates.

## Deploy

1. Upload `predictiveMaintenanceForm.css` to  
   `SiteAssets/styles/predictiveMaintenanceForm.css`
2. Upload `predictiveMaintenanceForm.js` to  
   `SiteAssets/scripts/predictiveMaintenanceForm.js`
3. Create a site page → **Embed** web part → paste contents of `predictiveMaintenanceForm.html`
4. After JS/CSS changes, bump `?v=` on the `<link>` and `<script>` tags in the HTML and re-paste

Dependencies (same as other hub widgets):

- jQuery: `SiteAssets/ExternalLibs/jquery371.min.js`

## Files

| File | Role |
|------|------|
| `predictiveMaintenanceForm.html` | Embed markup + asset links |
| `predictiveMaintenanceForm.css` | All widget styles |
| `predictiveMaintenanceForm.js` | Logic, REST, people picker, admin gate |

## Troubleshooting

- **Unauthorized:** add your email/login to `ADMIN_USERS`.
- **Save 400:** open browser console (F12); verify column display names and Sector choice values.
- **Unresolved columns:** console warns `[PM Form] Unresolved columns: …` — fix list column Titles to match the table above.
- **Stale UI after deploy:** hard refresh (Ctrl+F5) and confirm `?v=` was bumped on CSS/JS URLs.
