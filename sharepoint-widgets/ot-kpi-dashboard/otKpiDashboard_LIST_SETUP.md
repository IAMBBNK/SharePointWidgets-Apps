# OT KPI Dashboard — SharePoint list setup

**Site:** `https://mdigital.sharepoint.com/sites/ITOTCommunityHub`  
**Widget folder:** `sharepoint-widgets/ot-kpi-dashboard/`

---

## Recommendation: two lists

| List | Purpose |
|------|---------|
| **OT KPI Catalog** | Stable definitions (name, description, formula, category, target). Rarely changes. |
| **OT KPI Values** | Period snapshots (numerator, denominator, optional notes). One row per KPI per reporting period. |

### Why two lists (not one)

- Definitions and measurements change at different rates.
- You can keep history (Jan, Feb, …) without duplicating long description/formula text.
- The dashboard can show all KPIs from the catalog even before any values exist (`—` / awaiting data).
- Editors enter numbers in **Values**; owners maintain wording in **Catalog**.

A single list is only better if you will ever store the **current** score and overwrite it each month with no history.

---

## 1. List: `OT KPI Catalog`

1. Site contents → **New** → **List** → Blank list  
2. Name: **OT KPI Catalog** (exact)

### Columns

| Display name | Type | Required | Notes |
|--------------|------|----------|-------|
| Title | Single line | Yes | Short KPI name (display title) |
| KPICode | Single line | Yes | Stable id, e.g. `ORG-SITE-COVERAGE` — used to join Values |
| Category | Choice | Yes | See choices below |
| Description | Multiple lines | No | Business meaning of the KPI |
| Formula | Multiple lines | No | Calculation formula (documentation) |
| NumeratorLabel | Single line | No | What the numerator counts |
| DenominatorLabel | Single line | No | What the denominator counts |
| TargetPercent | Number | No | Optional target, e.g. `90` |
| SortOrder | Number | No | Display order within category (lower = first) |
| IsActive | Yes/No | No | Default Yes — hide retired KPIs when No |
| SourceSystem | Single line | No | e.g. CEAD, CMDB, SuccessFactors, ServiceNow |
| Owner | Single line | No | Business owner / team |

### Category choices (exact)

```text
Organization & Governance
Training & Compliance
Service Catalog
Asset & Inventory
DPI & Infrastructure
Disaster Recovery
Lifecycle Management
```

### Seed rows (create these catalog items)

Use **exact** `KPICode` values so the widget and Values list stay aligned.

| KPICode | Title | Category | SortOrder |
|---------|-------|----------|-----------|
| ORG-SITE-COVERAGE | Sites OT Organization Coverage | Organization & Governance | 10 |
| ORG-SITE-STAFFING | Site OT Staffing Adequacy | Organization & Governance | 20 |
| ORG-SECTOR-ROLES | Sector OT Organization Coverage | Organization & Governance | 30 |
| ORG-SECTOR-GOV | Sector OT Governance Coverage | Organization & Governance | 40 |
| TRN-FOUNDATIONAL | OT Foundational Training Compliance | Training & Compliance | 10 |
| TRN-CURRICULUM-ASSIGN | OT Role-Based Curriculum Assignment | Training & Compliance | 20 |
| TRN-ROLE-COMPLIANCE | OT Role-Based Training Compliance | Training & Compliance | 30 |
| SVC-APPROVAL | OT Service Offering Approval Compliance | Service Catalog | 10 |
| SVC-MANDATORY-ATTR | OT Service Offering Mandatory Attributes | Service Catalog | 20 |
| SVC-BOP-INSTANCE | BoP OT Service Offering-to-Instance | Service Catalog | 30 |
| INV-CEAD-COMPLETE | OT Application Inventory Completeness | Asset & Inventory | 10 |
| INV-CMDB-REGISTER | OT Asset CMDB Registration Completeness | Asset & Inventory | 20 |
| INV-SI-RELATIONSHIP | OT Service Instance Relationship Completeness | Asset & Inventory | 30 |
| DPI-SITE-ENABLE | DPI Site Enablement Coverage | DPI & Infrastructure | 10 |
| DPI-APP-INTEGRATED | Known OT Applications Integrated with DPI | DPI & Infrastructure | 20 |
| DR-BSO-PLAN | OT Service DR Plan Compliance | Disaster Recovery | 10 |
| DR-BSO-TEST | OT Service DR Test Evidence Compliance | Disaster Recovery | 20 |
| DR-DPI-SITE-PLAN | DPI Site DR Plan Compliance | Disaster Recovery | 30 |
| LIFECYCLE-SERVER-OS | OT Server Operating System Lifecycle | Lifecycle Management | 10 |
| LIFECYCLE-CLIENT-OS | OT Clients Operating System Lifecycle | Lifecycle Management | 20 |

### Catalog descriptions & formulas (paste into Description / Formula)

| KPICode | Description | Formula (× 100 for %) |
|---------|-------------|------------------------|
| ORG-SITE-COVERAGE | % of OT-relevant sites with a formal Site OT organization created and HR Job Family Operational Technology assigned. | Sites with official Site OT organization and HR Job Family OT applied / Total OT-related sites |
| ORG-SITE-STAFFING | % of OT-relevant sites that meet the minimum staffing threshold for day-to-day OT services. | Sites with official Site OT organization meeting ratio 1 OT FTE : 20 applications / Total OT-related sites |
| ORG-SECTOR-ROLES | % of required Sector OT roles assigned to named individuals with HR Job Family OT assigned. | Head of OT, Sector Architect, Sector Service Manager, Sector Portfolio Manager dedicated and HR Job Family applied / 16 |
| ORG-SECTOR-GOV | % of OT-relevant sites integrated into Sector OT governance forums (Architecture, Portfolio, Service Management CoPs). | Sites integrated in Architecture CoP, Portfolio Management CoP, and Service Management CoPs / Total OT-related sites |
| TRN-FOUNDATIONAL | % of OT Job Family employees who completed mandatory foundational OT training on time. | (OT employees assigned − OT employees overdue) / OT employees assigned |
| TRN-CURRICULUM-ASSIGN | % of OT Job Family employees assigned applicable role-based curricula per the OT Training Matrix. | OT employees assigned one or more OTSKM curricula / OT employees |
| TRN-ROLE-COMPLIANCE | % of employees with assigned OT role-based curricula who completed required training on time. | (OT employees assigned OTSKM curricula − OT employees overdue) / Employees assigned |
| SVC-APPROVAL | % of OT Service Offerings that completed required approval workflow and are in an approved lifecycle state. | Service Offerings with Service Type OT in Status Planned / Rejected Updating / Operational Updating / Total Catalog Service Offerings with Service Type OT |
| SVC-MANDATORY-ATTR | % of OT Service Offerings in catalogue phase where all mandatory attributes are completed. | Service Offerings with Service Type OT in Phase Catalog with all mandatory fields completed / Total Service Offerings with Service Type OT |
| SVC-BOP-INSTANCE | Whether BoP OT Business Service Offerings have at least one related BoP production Service Instance. | Business Service Offerings in Phase Catalog with Flag BoP and Service Type OT that have a related Service Instance (Installed) / Total BoP OT BSOs in scope |
| INV-CEAD-COMPLETE | % of OT-relevant sites with a completed and approved OT application inventory registered in CEAD. | Sites with full OT inventory registered in CEAD / Total OT-related sites |
| INV-CMDB-REGISTER | % of OT assets connected to the DM network registered as CIs in CMDB. | OT-connected assets registered in CMDB (e.g. Client CIs with DNS Domain containing "infra2") / Total endpoints connected to DM network |
| INV-SI-RELATIONSHIP | Completeness of Technical Service Offering assignment on OT Service Instances (invert “missing” count as needed for reporting). | Service Instances (Installed) under OT Business Service Offerings with Technical Service Offering assigned / Total such Service Instances |
| DPI-SITE-ENABLE | % of OT-relevant sites where DPI infrastructure is deployed and available for OT service integration. | Sites with DPI infrastructure / Total OT-related sites |
| DPI-APP-INTEGRATED | % of known OT applications integrated with DPI based on related OT Service Instances. | Service Instances belonging to OT BSOs using DPI-integrated CIs / Total in-scope OT Service Instances / applications |
| DR-BSO-PLAN | % of OT BSOs requiring DR that have an approved and current DR plan. | Business Service Offerings with approved DR Plan / Total OT BSOs in Operations that require DR |
| DR-BSO-TEST | % of OT BSOs requiring DR that have an approved and current DR test report. | Business Service Offerings with approved DR Test Report / Total OT BSOs in Operations that require DR |
| DR-DPI-SITE-PLAN | % of DPI-enabled sites with an approved and current site-level DR plan. | DPI sites with approved Site DR Plan / Number of DPI sites |
| LIFECYCLE-SERVER-OS | % of OT servers within OT BSOs running a supported and approved OS version. | Servers in OT BSOs with current/supported OS / Total servers within OT BSOs |
| LIFECYCLE-CLIENT-OS | % of OT clients within OT BSOs running a supported and approved OS version. | Clients in OT BSOs with current/supported OS / Total clients within OT BSOs |

---

## 2. List: `OT KPI Values`

1. Site contents → **New** → **List** → Blank list  
2. Name: **OT KPI Values** (exact)

### Columns

| Display name | Type | Required | Notes |
|--------------|------|----------|-------|
| Title | Single line | Yes | Convenience label, e.g. `ORG-SITE-COVERAGE — 2026-09` |
| KPICode | Single line | Yes | Must match Catalog `KPICode` |
| Period | Single line | Yes | Reporting period key, e.g. `2026-09` (YYYY-MM) or `2026-Q3` |
| PeriodDate | Date and time | No | First day of period — used for sorting |
| Numerator | Number | Yes | Count for the formula numerator |
| Denominator | Number | Yes | Count for the formula denominator (must be &gt; 0 to show %) |
| PercentOverride | Number | No | Optional manual % if you cannot supply num/den; leave blank to auto-calculate |
| Status | Choice | No | `Draft`, `Published`, `Superseded` — dashboard uses **Published** by default |
| Notes | Multiple lines | No | Data quality / assumptions |
| DataAsOf | Date and time | No | When the source extract was taken |

### Status choices

```text
Draft
Published
Superseded
```

### How to enter a period

For period `2026-09`, create **one Values row per KPICode** (20 rows). Example:

| Title | KPICode | Period | PeriodDate | Numerator | Denominator | Status |
|-------|---------|--------|------------|-----------|-------------|--------|
| ORG-SITE-COVERAGE — 2026-09 | ORG-SITE-COVERAGE | 2026-09 | 2026-09-01 | 42 | 60 | Published |

**Percent** on the dashboard = `100 × Numerator / Denominator` (unless `PercentOverride` is set).

---

## 3. Permissions

| Role | OT KPI Catalog | OT KPI Values |
|------|----------------|---------------|
| Dashboard viewers | Read | Read |
| KPI editors | Read (or Contribute if they maintain definitions) | Contribute |
| Admins | Full Control | Full Control |

---

## 4. Deploy the widget

1. Upload `otKpiDashboard.js` → `SiteAssets/scripts/otKpiDashboard.js`  
2. Create a site page → **Embed** web part → paste `otKpiDashboard.html` (CSS is inline)  
3. After JS changes, bump `?v=` on the script tag in the HTML and re-paste

### Config in `otKpiDashboard.js`

| Variable | Default |
|----------|---------|
| `SITE_URL` | Hub site URL |
| `CATALOG_LIST_NAME` | `OT KPI Catalog` |
| `VALUES_LIST_NAME` | `OT KPI Values` |
| `PUBLISHED_STATUS` | `Published` |

---

## 5. Empty data behaviour

Until **OT KPI Values** has Published rows for a period:

- Catalog KPIs still render (from the list, or built-in seed if the catalog list is empty)
- Score shows **—**
- Subtitle shows **Awaiting data**

When values exist, pick the period in the dashboard dropdown.

---

## 6. Troubleshooting

| Symptom | Check |
|---------|--------|
| No KPIs | Catalog list name; `IsActive` not set to No |
| All scores — | Create Published Values for that Period + matching KPICode |
| Wrong % | Denominator 0, or PercentOverride set |
| Period missing in dropdown | Values rows need `Period` populated |
| 403 | Viewer needs Read on both lists |
