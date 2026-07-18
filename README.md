# OXYGEN — Confident Daffodils Owners Association PMS

Property Management System for the Confident Daffodils Owners Association
(136 units): maintenance / waste / LPG fee tracking, bank statement imports,
LPG meter readings & inventory, corpus fund, party hall rental, move-in/out,
reports and more.

## How this project is put together

| Piece | Where it lives | What it does |
|---|---|---|
| **The app itself** | Google Apps Script (script.google.com) | All server code (`.gs`) and pages (`.html`) run on Google's servers; the database is a Google Sheet. |
| **`src/` folder (this repo)** | GitHub | A safety copy of every Apps Script file, with full change history. |
| **`index.html` (this repo)** | GitHub Pages | The public entrance — a full-screen wrapper that embeds the Apps Script web app. |

## The public URL

GitHub Pages serves `index.html` at:

**https://shinenair.github.io/OXYGEN/**

That page simply embeds the Apps Script deployment (`/exec` URL) full-screen.

## Making changes (the update routine)

1. Edit the file(s) in the **Apps Script editor** (script.google.com).
2. **Deploy → Manage deployments → ✏️ Edit → Version: "New version" → Deploy.**
   This updates the app at the SAME URL — nothing else to touch.
3. Upload the same changed file(s) to `src/` in this repo
   (Add file → Upload files) so the backup and history stay current.

⚠️ If you ever use **"New deployment"** instead, the app gets a NEW `/exec`
URL — then update the one line marked **APP URL** in `index.html` here.

## Repo layout

```
index.html   ← GitHub Pages wrapper (the public entrance)
README.md    ← this file
src/         ← all Apps Script project files (.gs, .html, appsscript.json)
```
