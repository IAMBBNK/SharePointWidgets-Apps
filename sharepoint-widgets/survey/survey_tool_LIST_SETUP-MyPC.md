# OT Integration Survey — SharePoint List Setup

Local widget files: `sharepoint-widgets/survey/`

Create these lists on **ITOT Community Hub** before using the survey widget.

**Site:** `https://mdigital.sharepoint.com/sites/ITOTCommunityHub`

---

## 1. OT SI Survey Dropdowns (roles)

**List name:** `OT SI - Survey Dropdowns`  
**List URL:** [OT SI - Survey Dropdowns — All Items](https://mdigital.sharepoint.com/sites/ITOTCommunityHub/Lists/OT%20SI%20-%20Survey%20Dropdowns/AllItems.aspx)

This list drives the **role** dropdown. Each row has a **Title** (shown in the dropdown) and **OptionType** set to `Role`.

### Setup

1. Open the list above (or create it with the **exact** name if it does not exist yet)  
2. Ensure a **Choice** column exists:
   - Display name: **OptionType**
   - Choices (exact spelling): `Role`
3. Add items — set **Title** and **OptionType** = `Role` for each row  

### Example items

| Title   | OptionType |
|---------|------------|
| OT Engineer | Role |
| Site Manager | Role |
| IT Project Lead | Role |

Everyone with access to the survey page needs **Read** on this list.

---

## 2. OT SI - Survey Reponses (submissions)

**List name:** `OT SI - Survey Reponses`  
**List URL:** [OT System Integration Survey — All Items](https://mdigital.sharepoint.com/sites/ITOTCommunityHub/Lists/OT%20System%20Integration%20Survey/AllItems.aspx)

1. Open the list above (or create it with the **exact** name if it does not exist yet)  
2. Add the columns below (Title column exists by default)

### Columns

| Display name     | Type            | Internal name (if created as shown) | Notes                          |
|------------------|-----------------|-------------------------------------|--------------------------------|
| Title            | Single line     | Title                               | Auto-filled on submit          |
| ContactEmail     | Single line     | ContactEmail                        | Stored only if follow-up consent is checked |
| ConsentFollowUp  | Yes/No          | ConsentFollowUp                     | Optional consent to store email / follow up |
| ContactRole      | Single line     | ContactRole                         | Selected role                  |
| MilestoneRank    | Multiple lines  | MilestoneRank                       | e.g. `M3: Solution Design;M1: Service Design;…` |
| BestMilestone    | Single line     | BestMilestone                       | Top-ranked milestone           |
| WorstMilestone   | Single line     | WorstMilestone                      | Bottom-ranked milestone        |
| BestPoint1       | Multiple lines  | BestPoint1                          | What went well                 |
| WorstPoint1      | Multiple lines  | WorstPoint1                         | What needs improvement         |
| AdditionalFeedback | Multiple lines | AdditionalFeedback                  | Optional extra milestone feedback |

Use **exact** display names above so internal names match the defaults in `survey_tool.js`.

### Permissions

- **Members:** Contribute (add items)  
- **Owners:** Full control  

---

## Deployment

1. Upload `survey_tool.js` to `SiteAssets/scripts/survey_tool.js`
2. Embed `survey_tool.html` on a SharePoint page (CSS is inline in the HTML)
3. Re-upload JS after logic changes (update the SharePoint script link if needed)

### Config in `survey_tool.js`

| Variable | Default | Notes |
|----------|---------|--------|
| `LOOKUP_LIST_NAME` | `OT SI - Survey Dropdowns` | |
| `LOOKUP_TYPE_FIELD` | `OptionType` | |
| `OPTION_TYPE_ROLE` | `Role` | |
| `LIST_NAME` | `OT SI - Survey Reponses` | |
| `SURVEY_REQUEST_ID` | `LE00XXX` | Paste approved Legal request ID before launch |
| `SURVEY_CONTACT_NAME` | placeholder | Survey contact person |
| `SURVEY_CONTACT_EMAIL` | placeholder | Shown in the intro |
| `SURVEY_CONTROLLER_NAME` | placeholder | Merck subsidiary name + address for privacy notice |
| `SURVEY_OBJECTION_EMAIL` | placeholder | Email for Art. 21 GDPR objections |
| `PRIVACY_STATEMENT_URL` | `https://www.merckgroup.com/en/privacy-statement.html` | Verify link works |

### Privacy / Legal checklist (Survey Approval)

The introduction section at the top of the questionnaire must include:

1. **Purpose of the survey** — `SURVEY_PURPOSE_TEXT`
2. **Contact person** — `SURVEY_CONTACT_NAME` + `SURVEY_CONTACT_EMAIL`
3. **Retention period for personal data** — `SURVEY_RETENTION_TEXT`
4. **Privacy notice** — fixed EU-GDPR text in HTML; fill `SURVEY_CONTROLLER_NAME` and `SURVEY_OBJECTION_EMAIL` in JS

Before launching:

1. Fill in all intro config values (contact, controller subsidiary + address, objection email).
2. After Legal approval, paste the request ID (`LE00XXX`) into `SURVEY_REQUEST_ID`.
3. Verify the **Data Privacy Statement** link opens (`PRIVACY_STATEMENT_URL` — default `https://www.merckgroup.com/en/privacy-statement.html`).
4. For the Legal upload package, print-to-PDF (or screenshots) of the live questionnaire including this introduction section.
