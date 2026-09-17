# Image Carousel — SPFx web part

A page web part for the IT/OT Community Hub: a **customizable button** that opens a **popup image carousel** (1 or 3 cards), similar to a lightbox gallery.

**Hub site:** `https://mdigital.sharepoint.com/sites/ITOTCommunityHub`

---

## What it does

- Renders a button on the page (text only, or text + icon/image).
- Clicking the button opens a full-page overlay with image cards.
- Each card has an image, title, and description.
- Editors choose **1 picture** or **3 pictures** in the web part properties.
- Arrow buttons, dots, Escape, and clicking the gray backdrop close / navigate.

---

## Prerequisites

| Requirement | Notes |
|-------------|-------|
| Node.js **18 LTS** or **22 LTS** | SPFx 1.21 does not support Node 24 |
| SharePoint Online tenant | Modern pages |
| Tenant **App Catalog** | To deploy the `.sppkg` |

---

## Quick start (development)

```powershell
cd "sharepoint-extensions/image-carousel"
npm install
gulp trust-dev-cert
gulp serve
```

The workbench is preconfigured to:

`https://mdigital.sharepoint.com/sites/ITOTCommunityHub/_layouts/workbench.aspx`

Add **Image Carousel**, then edit the web part to change the button and slides.

---

## Build and package

```powershell
npm install
gulp bundle --ship
gulp package-solution --ship
```

Output: `sharepoint-extensions/image-carousel/sharepoint/solution/image-carousel.sppkg`

---

## Deploy

1. Open the **Tenant App Catalog**.
2. Upload `image-carousel.sppkg`.
3. Deploy it and add the app to the IT/OT Community Hub (or make it available tenant-wide).
4. On a modern page: **Edit** → **Add a web part** → **Image Carousel**.

---

## Web part properties

| Property | Default | Description |
|----------|---------|-------------|
| Button text | `View gallery` | Label on the trigger button |
| Icon / image URL | *(empty)* | Optional icon next to the text. Leave empty for text only |
| Button color | `#0f69af` | Merck blue by default |
| Button alignment | Left | Left, center, or right |
| Cards visible | `3 pictures` | `1 picture` or `3 pictures` (always 1 on small screens) |
| Slides JSON | sample images | Array of `{ "image", "title", "description" }` |

Example slides JSON:

```json
[
  {
    "image": "https://mdigital.sharepoint.com/sites/ITOTCommunityHub/SiteAssets/your-image.jpg",
    "title": "Autumn",
    "description": "Autumn, a season full of great things."
  },
  {
    "image": "https://mdigital.sharepoint.com/sites/ITOTCommunityHub/SiteAssets/your-image-2.jpg",
    "title": "Night in The City",
    "description": "Silent and beautiful night in the city"
  }
]
```

Use Site Assets (or another library) URLs for production images.

---

## Component ID

```
3e7b9c14-5a2d-4f81-b6c0-8d4e1a9f2756
```

Keep this ID stable when updating the package so existing pages keep the web part.
