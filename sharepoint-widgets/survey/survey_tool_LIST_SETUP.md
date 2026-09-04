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
| ConsentFollowUp  | Choice          | ConsentFollowUp                     | Choices: `Yes`, `No` — optional consent to store email / follow up |
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

### Privacy / Legal checklist

#### Submitting the survey to Legal

1. Fill in the intro config values above (contact, controller, objection email, retention text if needed).
2. Deploy the survey page so the introduction section is visible.
3. **Upload format (required by Legal):** export the questionnaire as **print-to-PDF**, or — only if PDF is not possible — a text file with **screenshots** of the full survey tool (include the introduction section at the top).
4. Submit the Legal request. The survey already shows a **Request ID** placeholder (`LE00XXX`) in the intro — that is expected for the initial upload.

#### After Legal approval — before launch

1. Copy the approved request ID from Legal (format `LE00XXX`).
2. Paste it into `SURVEY_REQUEST_ID` in `survey_tool.js` (line ~15), then re-upload the JS to SharePoint.
3. Confirm the privacy statement link opens correctly.
4. Launch the survey only after the real request ID is showing on the live page.
