# Site Master Data Update Form — notes

## List

- **Site:** https://mdigital.sharepoint.com/sites/ITOTCommunityHub  
- **List:** `Site Master Data`  
- View (reference): [General Site Info Short](https://mdigital.sharepoint.com/sites/ITOTCommunityHub/Lists/Site%20Master%20Data/General%20Site%20Info%20Short.aspx)

## Deploy

1. Upload `siteMasterDataForm.js` to `SiteAssets/scripts/siteMasterDataForm.js`
2. Create a site page and add an **Embed** web part
3. Paste the contents of `siteMasterDataForm.html` (script tags already point at SiteAssets)
4. After each JS change, bump `?v=` on the script tag in the HTML and re-paste

jQuery must already exist at `SiteAssets/ExternalLibs/jquery371.min.js` (same as survey / KPI widgets).

## Access control (UX)

The form loads all list items the user can read, then keeps only items where the signed-in user appears in:

- **Site Data Admin**, and/or  
- **Site Head**, and/or  
- **Digital Lead / SPOC** (list column title may be misspelled `Digtial Lead / SPOC`)

SharePoint list permissions still apply on save (403 if the user cannot edit).  
`TEST_SHOW_ALL_SITES = true` bypasses this filter for testing.

## Behaviour

- **1 matched site:** form opens prefilled  
- **2+ matched sites:** dropdown → then prefilled form  
- **0 matched sites:** unauthorized message  
- **Save:** MERGE update of all mapped fields + sets **Review Status** to `Completed / Updated`

## Field mapping

On load the script maps **exact column Titles** from the list (`FIELD_DEFS` in `siteMasterDataForm.js`). Notable quirks:

- **Digtial Lead / SPOC** — typo in the list title is matched intentionally  
- **Gear** — used for OT systems inventorized  
- **Current Maturity (BPOG)** / **Target Maturity (BPOG)** — single-line text  
- Calculated / system columns (Migration Completed, Site Image, Created/Modified, …) are not on the form  

Unresolved fields are logged as `[Site Master Data Form] Unresolved fields …`.

## People picker

Uses SharePoint REST:

- `ClientPeoplePickerWebServiceInterface.ClientPeoplePickerSearchUser`
- `/_api/web/ensureuser`

Writes person columns via `FieldNameId` / `{ results: [...] }` for multi-value.
