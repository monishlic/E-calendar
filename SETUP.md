# E-Calendar — Setup & Install Guide

You'll do three things, once:

1. **Publish this repo on GitHub** → gives the app a free HTTPS web address.
2. **Deploy the backend** (Google Apps Script) → your data + logic.
3. **Install it as an app** on your phone / desktop.

Total time: **~10 minutes.** You only need a GitHub account (you have one) and a Google account.

> **Why GitHub?** It gives you (a) a free HTTPS link — and HTTPS is exactly what lets the app be *installed* like a native app, and (b) version history, so every future change is saved and can be rolled back.

---

## Part A — Publish on GitHub (4 min)

### A1 · Create the repository
1. Go to **github.com → New repository** (the green **New** button, or `+` top-right → New repository).
2. **Repository name:** `e-calendar` (any name is fine).
3. Keep it **Public** (simplest — free GitHub Pages works on public repos).
   *Safe to be public?* **Yes.** There are no passwords in these files. Your actual data lives in your private Google Sheet and is protected by the PIN you'll set. (Your Gemini key, if you add one, is stored in Apps Script, never in this repo.)
4. **Don't** tick "Add a README" (you already have one). Click **Create repository**.

### A2 · Upload the files
1. On the new empty repo page, click **uploading an existing file** (the link in the middle), or **Add file → Upload files**.
2. Open the **`ecalendar`** folder I gave you. Select **everything inside it** and drag it into the browser:
   - `index.html`, `manifest.json`, `sw.js`
   - `favicon.png`, `icon-192.png`, `icon-512.png`, `icon-512-maskable.png`, `apple-touch-icon.png`
   - the **`backend`** folder (contains `Code.gs`)
   - `README.md`, `SETUP.md`

   ⚠️ **Important:** `index.html` must sit at the **root** of the repo — *not* inside a sub-folder. (If you drag the `ecalendar` folder itself, GitHub will nest everything one level too deep and the site will 404.)
3. Scroll down → **Commit changes**.

### A3 · Turn on GitHub Pages
1. In your repo: **Settings** (top tab) → **Pages** (left sidebar).
2. Under **Build and deployment → Source**, choose **Deploy from a branch**.
3. **Branch:** `main`  ·  Folder: `/ (root)` → **Save**.
4. Wait ~1 minute, then refresh. Pages shows your live link:
   ```
   https://<your-username>.github.io/e-calendar/
   ```
   **This is your app URL.** Open it — you'll see the "Connect" screen. (Keep this link; you'll install from it.)

---

## Part B — Deploy the backend (5 min)

### B1 · Create a Google Sheet
1. Go to **sheets.google.com → Blank**.
2. Rename it **E-Calendar** (top-left title).

### B2 · Open the script editor
1. In the sheet: **Extensions → Apps Script**.
2. Delete whatever is in the default `Code.gs`.

### B3 · Paste the backend
1. Open **`backend/Code.gs`** from this repo, copy **all** of it.
2. Paste into the empty `Code.gs` → press **Ctrl/Cmd + S** to save.

### B4 · (Optional) create the tabs
- In the function dropdown pick **`initializeSheets`** → **▶ Run** → approve the permission prompt (**Advanced → Go to E-Calendar (unsafe) → Allow**).
- If it shows *"An unknown error occurred,"* ignore it — the tabs also create themselves automatically on first use.

### B5 · Deploy as a Web App
1. **Deploy → New deployment** → gear ⚙️ → **Web app**.
2. Set:
   - **Execute as:** **Me**
   - **Who has access:** **Anyone**  ← required so the app can reach it
3. **Deploy** → approve if asked → **copy the Web App URL** (ends in **`/exec`**).

---

## Part C — Connect & secure (2 min)

1. Open your **GitHub Pages URL** (from A3) on your computer.
2. Paste the **`/exec` URL** → **Connect**. The signal turns green and the app loads.
3. **Set a PIN:** Settings → *Access PIN* → set a 4–12 digit PIN (protects your data; asked each time the app opens).
4. **Optional — AI smart voice:** Settings → *AI Smart Voice* → paste a free Gemini key from **aistudio.google.com/apikey** → Save. (Lets voice understand any phrasing. Skip it and voice still works with the built-in parser.)

---

## Part D — Install as an app 📲

Open your GitHub Pages URL, then:

- **Android (Chrome):** you'll get an "Install app" banner — or menu **⋮ → Install app / Add to Home screen**. (Also: Settings → **Install as App** → Install.)
- **iPhone / iPad (Safari):** tap **Share** → **Add to Home Screen**. *(Must be Safari.)*
- **Windows / Mac (Chrome or Edge):** click the **install icon** in the address bar (a little monitor/⊕), or Settings → **Install as App** → Install.

It now opens in its own window with an **E-Calendar** icon — no browser tabs, just like a native app. Everyone in the office can install it from the same link; all share the same live data.

---

## Part E — Shipping future enhancements

This is why we put it on GitHub. When I give you updated files:

**Frontend change (`index.html`, icons, etc.):**
1. In your repo, click the file → the ✏️ pencil → paste the new version → **Commit** (or **Add file → Upload files** to replace).
2. Open **`sw.js`**, bump the version once: change `ecal-v1` → `ecal-v2` → Commit. *(This tells installed apps to pull the update.)*
3. GitHub Pages redeploys in ~1 min; installed apps refresh the next time they're opened.

**Backend change (`Code.gs`):**
1. Paste the new code into Apps Script → **Ctrl/Cmd + S**.
2. **Deploy → Manage deployments → ✏️ Edit → Version: New version → Deploy.** (Keeps the same `/exec` URL — no reconnect needed.)

**Rollback:** GitHub saves every commit. Repo → **Commits** → open an older one → **Revert**, and you're back.

---

## Troubleshooting

**Pages URL shows 404** → `index.html` isn't at the repo root (it's nested in a folder), or Pages isn't ready yet (wait 1–2 min and refresh).

**"Install" option never appears** → You must open the **https GitHub Pages URL**, not a local file. On iPhone, you must use **Safari**. If already installed, that's why it's hidden.

**"Failed to fetch" / "Server returned non-JSON"** → The Apps Script wasn't deployed with **Who has access: Anyone**, or you used the wrong URL (it must end in **`/exec`**). Fix: Apps Script → Deploy → Manage deployments → ✏️ → access **Anyone** → New version → Deploy.

**App didn't update after a change** → Make sure you bumped the `sw.js` version. Then fully close and reopen the installed app (or in a browser tab, reload twice).

**Signal is orange / slow on first open** → Normal. Apps Script "cold-starts" take a few seconds after being idle; it speeds up after the first request.

**Forgot the PIN** → Open your Google Sheet → **Settings** tab → find the `ACCESS_PIN` row → clear that cell. The app will let you set a new PIN next open.

---

## Files in this repo

| File | What it is |
|---|---|
| `index.html` | The whole app — served by GitHub Pages |
| `manifest.json`, `sw.js` | Make it installable + offline |
| `*.png` icons | App icons (home screen / desktop) |
| `backend/Code.gs` | Backend — paste into Google Apps Script |
| `README.md` | Repo overview |
| `SETUP.md` | This guide |

Cost: **₹0 / month.** GitHub Pages + Google Sheets + Apps Script + Gemini free tier.
