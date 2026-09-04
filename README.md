# Merck SharePoint Scripts

Local workspace for IT/OT Community Hub SharePoint embeds, templates, and related assets.

**SharePoint site:** `https://mdigital.sharepoint.com/sites/ITOTCommunityHub`  
**Deploy target:** `SiteAssets/scripts/` (and `SiteAssets/ExternalLibs/` for libraries)

HTML widgets use **absolute SiteAssets URLs** in their `<script>` / `<link>` tags. Local folder names do not need to match SharePoint paths — upload each widget’s JS/CSS to SiteAssets when deploying.

## Folder structure

```
sharepoint-widgets/     Production embed widgets (upload to SharePoint)
  kpi-dashboard/        Site Master Data KPI dashboard
  page-kpi-dashboard/   Page views + newsletter engagement KPIs
  trainings/            Training portal accordion
  resources/            Key resources portal (JS: KeyResources.js on SharePoint)
  survey/               OT integration feedback survey
  site-master-data-form/ Site Master Data update form (admins / directors)
  site-completion-dashboard/ Site Master Data field-completion dashboard
  timeline/             Vertical timeline hero
  lib/                  Shared libraries (e.g. Chart.js)

prototypes/             Experiments and web part prototypes
templates/              Document / HTML templates (Word, SoW, etc.)
powerpoint/             PowerPoint VBA animation macros
data/                   Data exports (Duolingo, schedule extracts, …)
tools/                  One-off utility scripts (Python, etc.)
```

## Widget deployment checklist

1. Edit files under `sharepoint-widgets/<widget>/`
2. Upload JS (and CSS if any) to `SiteAssets/scripts/`
3. Paste the widget `.html` into a SharePoint embed web part
4. Bump the `?v=` cache-bust query on script/style tags after changes

## Widget index

| Widget | Local path | SharePoint JS name |
|--------|------------|-------------------|
| KPI Dashboard | `sharepoint-widgets/kpi-dashboard/` | `kpiDashboard.js` |
| Page & Newsletter KPIs | `sharepoint-widgets/page-kpi-dashboard/` | `pageKpiDashboard.js` |
| Trainings | `sharepoint-widgets/trainings/` | `itot_training.js` |
| Resources | `sharepoint-widgets/resources/` | `KeyResources.js` |
| Survey | `sharepoint-widgets/survey/` | `survey_tool.js` |
| Site Master Data Form | `sharepoint-widgets/site-master-data-form/` | `siteMasterDataForm.js` |
| Site Completion Dashboard | `sharepoint-widgets/site-completion-dashboard/` | `siteCompletionDashboard.js` |
| Timeline | `sharepoint-widgets/timeline/` | (inline CSS only) |
| Convergence plan | `prototypes/` | `itot-convergence-plan/projects-webpart.js` |

See `sharepoint-widgets/survey/survey_tool_LIST_SETUP.md` for survey list setup.  
See `sharepoint-widgets/page-kpi-dashboard/pageKpiDashboard_LIST_SETUP.md` for Page KPIs folder + Newsletter KPIs list setup.  
See `sharepoint-widgets/site-master-data-form/LIST_NOTES.md` for Site Master Data form deploy and field mapping notes.
