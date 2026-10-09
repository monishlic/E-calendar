# E-Calendar

An office **leave, attendance & salary-cycle calendar** — installable as an app on phone and desktop.

It runs as a Progressive Web App (PWA): the interface is a single HTML file hosted free on **GitHub Pages**, and the data lives in your own **Google Sheet** via a **Google Apps Script** backend. No servers, no monthly bills.

> **New here? Open [`SETUP.md`](SETUP.md)** — it walks you through publishing this repo, deploying the backend, and installing the app, step by step (~10 minutes).

---

## What it does

- Visual month calendar of who's on leave, with drag-to-select multi-day entry
- Full-day / half-day leaves; **Sundays and office-closed holidays never count**
- Per-staff opening balances, annual leave quota, and **salary-cycle deduction reports**
- Comp-off / overtime credits (earned and used, excluded from deductions)
- Indian holidays library with one-click import and a per-holiday "office closed" toggle
- **Voice commands** (English + Hindi) for every action — with an optional AI "smart mode"
- PIN protection, freeze-lock for already-paid cycles, and a live connection signal
- Works offline for browsing; syncs the moment you're back online

---

## Project structure

```
index.html                 The whole app (UI + logic in one file) — served by GitHub Pages
manifest.json              PWA manifest (app name, icons, colors)
sw.js                      Service worker (installable + offline shell)
icon-192.png               App icons
icon-512.png
icon-512-maskable.png
apple-touch-icon.png       iOS home-screen icon
favicon.png
backend/
  Code.gs                  Google Apps Script backend (paste into your Apps Script project)
SETUP.md                   Full setup + install + "how to ship updates" guide
README.md                  This file
```

## Tech

- **Frontend:** one self-contained HTML file (vanilla JS, no build step) + a PWA service worker
- **Backend:** Google Apps Script Web App exposing a JSON API
- **Database:** a Google Sheet (tabs: Staff, Leaves, Holidays, Settings, Templates, Compensation)
- **Hosting:** GitHub Pages (HTTPS, which is what makes the app installable)
- **Optional AI:** Google Gemini (free API key) for smart voice understanding

## Cost

₹0 / month. GitHub Pages, Google Sheets, Apps Script, and the Gemini free tier are all free.
