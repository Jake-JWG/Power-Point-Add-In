# PowerPoint Slide Status Add-in

This repository contains a beginner-friendly Microsoft PowerPoint Office Add-in that lets you:

- set a status marker on the **current slide**
- set a status marker on **all slides** in one click

Built-in statuses:

- To be updated
- Draft
- For review
- Completed

---

## What it does

The add-in opens a task pane in PowerPoint.

When you choose a status, it adds (or updates) a status text box on each target slide with color-coded text:

- 🔴 To be updated
- 🟠 Draft
- 🔵 For review
- 🟢 Completed

This is designed for month-to-month deck rollover, where you can set all slides back to **To be updated** quickly.

---

## Files

- `manifest.xml` – Office Add-in manifest to sideload in PowerPoint
- `src/taskpane.html` – task pane UI
- `src/taskpane.css` – styling
- `src/taskpane.js` – PowerPoint logic for applying statuses

---

## The most important concept (for first-time users)

PowerPoint does **not** run the add-in code directly from your `.pptx` file.

It loads your add-in from a **web URL** (must be HTTPS), and the `manifest.xml` tells PowerPoint where that URL is.

So the workflow is:

1. Put these files on an HTTPS web host.
2. Update `manifest.xml` to point to your hosted URL.
3. Sideload (install) the manifest in PowerPoint.
4. Open the add-in from the ribbon.

---

## Step-by-step: implement this in PowerPoint

## 1) Host the add-in files on HTTPS

You need to host the repository files (`src/*` + icons) on a URL like:

- `https://your-domain.com/src/taskpane.html`

Good beginner options:

- GitHub Pages (free)
- Netlify / Vercel (free tiers)
- Any company HTTPS static hosting

> Tip: If your host serves files from the repo root, the path `/src/taskpane.html` should work directly.

## 2) Update `manifest.xml` URLs

Open `manifest.xml` and replace every `https://localhost:3000` with your real HTTPS host.

Examples of fields to update:

- `IconUrl`
- `HighResolutionIconUrl`
- `SupportUrl`
- `AppDomain`
- `SourceLocation`
- `bt:Image` URLs
- `Taskpane.Url`

If your site URL is `https://my-status-addin.netlify.app`, then:

- `SourceLocation` should become `https://my-status-addin.netlify.app/src/taskpane.html`

Save the file as `manifest.xml`.

## 3) Sideload manifest in PowerPoint (Windows desktop)

1. Open PowerPoint.
2. Go to **Insert** tab.
3. Click **My Add-ins**.
4. Choose **Manage My Add-ins** (or **Upload My Add-in**, wording can vary).
5. Upload your updated `manifest.xml`.
6. Open the add-in pane when prompted.

If your Office build shows a different path, use:

- **Insert → Add-ins → My Add-ins → Upload My Add-in**

## 4) Sideload manifest in PowerPoint (Mac desktop)

Typical process:

1. Open PowerPoint.
2. **Insert → Add-ins → My Add-ins**.
3. Use the upload option for custom add-ins.
4. Select your `manifest.xml`.

If upload is restricted in your organization, your IT admin may need to deploy it via Microsoft 365 admin center.

## 5) Use the add-in

1. Open any presentation.
2. Open **Slide Status Manager** from the ribbon button.
3. In the task pane:
   - click a status under **Current slide** to mark only the selected slide
   - click a status under **All slides** to set every slide
4. For your monthly rollover workflow, click **To be updated** under **All slides**.

---

## Troubleshooting

- **Add-in does not load**
  - Confirm all manifest URLs are HTTPS and reachable in a browser.
- **Button appears but pane is blank**
  - Check that `/src/taskpane.html`, `/src/taskpane.js`, and `/src/taskpane.css` are all accessible on your host.
- **Company laptop blocks custom add-ins**
  - Ask IT to allow sideloading or deploy through centralized admin deployment.
- **Icons missing**
  - The manifest uses externally hosted icon URLs; verify those URLs are reachable.

- **Codex Create PR says “Binary files are not supported”**
  - This repo is now text-only to avoid that issue.
  - Keep icons hosted by URL (as in `manifest.xml`) instead of committing PNG binaries.

---

## Notes

- This version uses emoji color indicators for status text.
- You can later customize wording, location, size, or styling in `src/taskpane.js` and `src/taskpane.css`.

---

## Windows + GitHub: non-technical click-by-click checklist

Use this exact checklist if you are on **Windows** and want to host using **GitHub**.

### A) Put this project on GitHub

1. Sign in to GitHub in your web browser.
2. Click the **+** (top-right) → **New repository**.
3. Repository name: for example `powerpoint-status-addin`.
4. Choose visibility (Public is easiest for first setup).
5. Click **Create repository**.
6. Upload this project’s files to that repository (or push from Git if you already use it).

### B) Turn on GitHub Pages (this hosts your add-in)

1. Open your repository on GitHub.
2. Click **Settings** tab.
3. In the left menu, click **Pages**.
4. Under **Build and deployment**:
   - **Source**: choose **Deploy from a branch**.
   - **Branch**: choose **main**.
   - **Folder**: choose **/(root)**.
5. Click **Save**.
6. Wait 1–5 minutes.
7. GitHub will show your site URL, for example:
   - `https://YOUR-USERNAME.github.io/powerpoint-status-addin`

### C) Edit `manifest.xml` on Windows (Notepad is fine)

1. On your PC, open the local `manifest.xml` file.
2. Press **Ctrl + H** (Replace).
3. Find what: `https://localhost:3000`
4. Replace with: your GitHub Pages URL, e.g.
   - `https://YOUR-USERNAME.github.io/powerpoint-status-addin`
5. Click **Replace All**.
6. Save the file.

### D) Quick check in your browser

Open these links in your browser (replace with your own URL):

- `https://YOUR-USERNAME.github.io/powerpoint-status-addin/src/taskpane.html`
- `https://YOUR-USERNAME.github.io/powerpoint-status-addin/src/assets/icon-32.png`

If both open, your hosting is ready.

### E) Add it to PowerPoint on Windows

1. Open **PowerPoint**.
2. Open any presentation.
3. Click **Insert** tab.
4. Click **My Add-ins**.
5. Click **Upload My Add-in** (or **Manage My Add-ins** → upload).
6. Select your edited `manifest.xml` file.
7. Click **Open**.
8. The add-in appears as **Slide Status Manager**.

### F) Use it each month

1. Open your monthly copied deck.
2. Open **Slide Status Manager** from the ribbon.
3. Under **All slides**, click **🔴 To be updated**.
4. Then set individual slides to Draft / For review / Completed as work progresses.

### G) If it doesn’t appear

1. Close and reopen PowerPoint.
2. Re-open **My Add-ins**.
3. Re-check that your `manifest.xml` has your GitHub Pages URL (not localhost).
4. Re-check that `taskpane.html` URL opens in browser.
5. If work laptop policies block custom add-ins, ask IT to allow Office add-in sideloading.

---

## Use Codex "Create PR" button to push to GitHub

If you want Codex to push this project via the **Create PR** flow, make sure these are true first:

1. The folder is a Git repository (it is).
2. A GitHub remote is set, for example:
   - `https://github.com/Jake-JWG/Power-Point-Add-In.git`
3. You are authenticated to GitHub in the environment (token or credential helper).

### One-time setup checklist

1. Set the remote (if missing):
   - `git remote add origin https://github.com/Jake-JWG/Power-Point-Add-In.git`
2. Ensure your branch has commits.
3. Ensure there are no uncommitted changes you still want to edit.

### Codex PR flow checklist

1. Ask Codex to make/update files.
2. Ask Codex to **commit** changes.
3. Ask Codex to **create PR**.
4. Click **Create PR** in the Codex UI when prompted.

If the Create PR action fails, the most common causes are:

- missing GitHub authentication
- remote URL not set
- network/proxy blocking github.com
