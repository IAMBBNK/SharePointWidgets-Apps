# Page & Newsletter KPI Dashboard — Setup

Local widget files: `sharepoint-widgets/page-kpi-dashboard/`

**Site:** `https://mdigital.sharepoint.com/sites/ITOTCommunityHub`

This dashboard is separate from the Site Master Data KPI dashboard.

---

## 1. Page KPIs folder (Power Automate export)

The widget reads the **newest `.txt` file** from a document library folder named **Page KPIs**.

### Expected file content

A JSON array of strings in this shape:

```json
["\".\"HubSection:{\"@odata.type\":\"#Microsoft.Azure.Connectors.SharePoint.SPListExpandedReference\",\"Id\":1,\"Value\":\"General\"},Link:https://mdigital.sharepoint.com/sites/ITOTCommunityHub/SitePages/TopicHome.aspx,Title:Home,Views:15380,Viewers:1289"]
```

The script:

- Parses `HubSection` (uses `Value`), `Link`, `Title`, `Views`, `Viewers`
- Treats empty Views/Viewers as `0`
- Deduplicates by Link (falls back to Title); keeps the higher Views value
- Classifies pages:
  - **Menu:** Program, Portfolio, Architecture, General, Service Management, News & Connect, Upskilling & Resources, Digital Plant
  - **Articles:** Articles
  - **Ignored:** Templates, Archive (and blank/unknown HubSection)
- Newsletter story URLs are matched to page Links to show page views/viewers beside NL clicks

### Config

In `pageKpiDashboard.js`, set the primary path (fallbacks are tried automatically):

```js
var PAGE_KPI_FOLDER_SERVER_RELATIVE_URL =
  "/sites/ITOTCommunityHub/Shared Documents/Page KPIs";
```

The script also tries:

- `/sites/ITOTCommunityHub/Documents/Page KPIs`
- `/sites/ITOTCommunityHub/SiteAssets/Page KPIs`
- `/sites/ITOTCommunityHub/Page KPIs` (library root)
- List/library titled **Page KPIs**

**How to get the exact path:** open a KPI `.txt` file in SharePoint → **Details** → copy the folder path, or from the browser URL when viewing the folder.

Expected file names look like:

`KPIs-2026-07-22T08_30_05.4838198Z.txt`

Everyone viewing the dashboard page needs **Read** on that library/folder.

---

## 2. Newsletter KPIs list (manual monthly entry)

**List name:** `Newsletter KPIs`

1. Site contents → **New** → **List** → Blank list  
2. Name: **Newsletter KPIs**  
3. Add the columns below (Title exists by default)

### Columns

| Display name | Type | Notes |
|--------------|------|--------|
| Title | Single line | Edition name, e.g. `IT/OT Newsletter – June 2026` |
| SentDateTime | Date and time | When the newsletter was sent (used for ordering/labels) |
| DeliveredTo | Number | Number of people the newsletter was delivered to |
| Notes | Multiple lines | Optional |
| Opens24hSeries | Multiple lines | 24 incremental hourly opens (comma-separated) |
| Clicks24hSeries | Multiple lines | 24 incremental hourly clicks (comma-separated) |
| Opens7dSeries | Multiple lines | 7 incremental daily opens (comma-separated) |
| Clicks7dSeries | Multiple lines | 7 incremental daily clicks (comma-separated) |
| Story1Title | Single line | Story headline (leave blank if unused) |
| Story1Url | Single line | Story link URL (optional) |
| Story1Clicks | Number | Clicks on story 1 |

| Story2Title | Single line | |
| Story2Url | Single line | |
| Story2Clicks | Number | |

| Story3Title | Single line | |
| Story3Url | Single line | |
| Story3Clicks | Number | |

| Story4Title | Single line | |
| Story4Url | Single line | |
| Story4Clicks | Number | |

| Story5Title | Single line | |
| Story5Url | Single line | |
| Story5Clicks | Number | |

| Story6Title | Single line | |
| Story6Url | Single line | |
| Story6Clicks | Number | |

Use **exact** display names so internal names match `pageKpiDashboard.js`.

### Series format (incremental)

Paste comma-separated integers. Values are **increments** for that hour or day (not running totals).

- **24h series:** exactly 24 values (hour 1 … hour 24)  
- **7d series:** exactly 7 values (day 1 … day 7)  

Example `Opens24hSeries`:

```text
12,8,5,3,2,1,0,0,1,2,4,6,5,3,2,1,1,0,0,0,1,0,0,0
```

Example `Opens7dSeries`:

```text
40,18,12,8,5,3,2
```

The dashboard:

- **Cards:** sum of each series (24h / 7d totals)
- **Charts:** hourly and daily trends for Opens + Clicks

### How to use

1. After each monthly newsletter send, create a **new list item**
2. Set **SentDateTime** and **DeliveredTo**
3. Paste Opens/Clicks series for 24h (hourly) and 7d (daily)
4. Fill up to 6 story titles, optional URLs, and click counts
5. The dashboard dropdown shows **Overall (all editions)** plus each edition  
   - Overall: Delivered to, Opens/Clicks cards, and charts = averages across editions  
   - Overall charts = pointwise average of series across editions

### Permissions

- Editors who enter KPIs: Contribute  
- Dashboard viewers: Read  

---

## 3. Page View Tracking list (Viva Engage views)

The dashboard also reads **`Page View Tracking`** (written by the page-view-tracker extension) to show **Viva Engage views** on news articles.

| Config | Default |
|--------|---------|
| List name | `Page View Tracking` |
| Source filter | `Source eq 'viva'` (from links with `?source=viva`) |

### Behaviour

- News **articles** table gets a **Viva views** column
- Counts tracking rows matched to the article URL / path
- If an article has **no** matching Viva source rows → **N/A** (not `0`)
- Summary card **Viva Engage views** = sum across articles that have data; **N/A** if none do

Viva share links must use:

```text
...?source=viva
```

See `sharepoint-extensions/page-view-tracker/LIST_SETUP.md` for list columns and permissions.

Dashboard viewers need **Read** on **Page View Tracking**.

---

## 4. Deployment

1. Upload `pageKpiDashboard.js` to `SiteAssets/scripts/pageKpiDashboard.js`
2. Embed `pageKpiDashboard.html` on a SharePoint page (CSS is inline)
3. After JS changes, bump the `?v=` query string on the script tag in the HTML

### Config in `pageKpiDashboard.js`

| Variable | Default |
|----------|---------|
| `SITE_URL` | `https://mdigital.sharepoint.com/sites/ITOTCommunityHub` |
| `PAGE_KPI_FOLDER_SERVER_RELATIVE_URL` | `/sites/ITOTCommunityHub/Shared Documents/Page KPIs` |
| `NEWSLETTER_LIST_NAME` | `Newsletter KPIs` |
| `PAGE_VIEW_TRACKING_LIST_NAME` | `Page View Tracking` |
| `VIVA_SOURCE_VALUE` | `viva` |
| `TOP_PAGES_COUNT` | `10` |
