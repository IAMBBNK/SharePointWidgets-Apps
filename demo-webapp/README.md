# Intranet services demo

Standalone Vite + React site: a services offer for SharePoint and Confluence, plus a SharePoint-like demo intranet with mock data.

No live tenant, no `.sppkg`. Widgets are visual remakes of the hub web parts.

## Run

```powershell
cd demo-webapp
npm install
npm run dev
```

Open the URL Vite prints (usually `http://localhost:5173`).

- `/` — services landing
- `/demo` — Contoso Hub
- `/demo/catalog` — every widget

## Build

```powershell
npm run build
```
