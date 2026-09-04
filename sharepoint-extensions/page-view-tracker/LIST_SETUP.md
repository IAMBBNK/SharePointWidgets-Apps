# Page View Tracking — SharePoint list setup

**Site:** `https://mdigital.sharepoint.com/sites/ITOTCommunityHub`  
**Extension:** `sharepoint-extensions/page-view-tracker/`

The Application Customizer writes one list item per tracked page view.

---

## 1. Create the list

1. Open **Site contents** on the hub site.
2. **New** → **List** → **Blank list**.
3. Name: **Page View Tracking** (exact name — matches extension default config).

---

## 2. Add columns

Use these **display names** so internal names match the extension.

| Display name | Type | Required | Notes |
|--------------|------|----------|-------|
| Title | Single line | Yes | Exists by default — page title |
| PageUrl | Single line | Yes | Full page URL including query string |
| PagePath | Single line | Yes | Server-relative path, e.g. `/sites/ITOTCommunityHub/SitePages/News.aspx` |
| Source | Single line | Yes | Traffic source from URL, e.g. `viva`, `direct` |
| UserEmail | Single line | No | Viewer e-mail |
| UserDisplayName | Single line | No | Viewer display name |
| ViewedAt | Date and time | Yes | UTC timestamp (include time) |
| Referrer | Single line | No | Browser referrer URL |
| IsNewsPage | Yes/No | No | `Yes` for SharePoint news posts |
| HubSection | Single line | No | Hub section site column when present |
| UserAgent | Multiple lines | No | Browser user agent |

### Column tips

- Add columns in the order above to keep internal names predictable.
- **ViewedAt** should use **Date and time** with time enabled.
- If **HubSection** already exists as a site column on **Site Pages**, the extension reads it automatically. You do not need a separate HubSection column on this list unless you want to store the copied value (the extension writes to this list column if it exists).

---

## 3. Permissions

| Role | Permission | Why |
|------|------------|-----|
| All hub visitors | **Contribute** on this list | Extension creates items as the signed-in user |
| Dashboard / reporting owners | **Read** or **Contribute** | For Power BI, exports, or manual review |

The extension runs in the context of the current user, so each viewer must be allowed to add items. If you prefer a service account pattern, use a Power Automate flow triggered differently instead of direct list writes.

**Alternative ( tighter security):** Break inheritance on the list, grant **Contribute** to **Everyone except external users** (or a hub visitors group), and keep the list hidden from quick launch.

---

## 4. Viva Engage link format

When sharing news on Viva Engage, append a source query parameter:

```text
https://mdigital.sharepoint.com/sites/ITOTCommunityHub/SitePages/Your-News-Post.aspx?source=viva
```

Other examples:

| Link type | Example parameter | Stored `Source` value |
|-----------|-------------------|------------------------|
| Viva Engage | `?source=viva` | `viva` |
| Newsletter | `?source=newsletter` | `newsletter` |
| Teams | `?source=teams` | `teams` |
| No parameter | (none) | `direct` (default) |

Source values are normalized to lowercase.

---

## 5. What gets tracked

| Scenario | Tracked? |
|----------|----------|
| Modern **Site Pages** on the hub | Yes |
| Page opened in **edit** mode | No |
| Same page + same source in one browser session | No (session dedupe) |
| Page refresh in a new session | Yes |
| News only (`trackNewsOnly: true`) | Only `PromotedState = 2` pages |

---

## 6. Reporting ideas

- **Views by source:** group by `Source`, count rows
- **Unique viewers:** count distinct `UserEmail` per `PagePath`
- **Viva vs direct:** filter `Source eq 'viva'`
- **News performance:** filter `IsNewsPage eq 1`

You can connect Power BI to this list or export to CSV for your existing Page KPI dashboard workflow.

---

## 7. Troubleshooting

| Symptom | Check |
|---------|--------|
| No list items | User has Contribute on **Page View Tracking**; list name matches config |
| 403 on write | List permissions for hub visitors |
| Missing `HubSection` | Site column exists on **Site Pages** library |
| Extension not loading | App deployed from tenant catalog and registered on hub site |
| Duplicate rows every refresh | Session storage blocked — expected in strict privacy browsers |

Use browser dev tools → **Console** and filter for `PageViewTracker` after deploying with `gulp serve` locally.
