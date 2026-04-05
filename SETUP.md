# Asset Library — PowerPoint Add-in
## Complete Setup Guide

---

## What's in This Package

```
asset-library/
├── taskpane.html     ← The task pane UI (main app)
├── manifest.xml      ← Add-in installer file (users load this)
├── commands.html     ← Required Office.js stub
└── SETUP.md          ← This guide
```

---

## PART 1 — Host the Files on GitHub Pages (One-Time Setup, You Do This)

GitHub Pages gives you a free permanent HTTPS URL — exactly what Office add-ins require.

### Step 1 — Create a GitHub Account
Go to **github.com** and sign up (free).

### Step 2 — Create a New Repository
1. Click **New repository**
2. Name it exactly: `asset-library`
3. Set it to **Public**
4. Click **Create repository**

### Step 3 — Upload the Files
1. Click **Add file → Upload files**
2. Drag in all 3 files: `taskpane.html`, `commands.html`, `manifest.xml`
3. Click **Commit changes**

### Step 4 — Enable GitHub Pages
1. Go to your repo → **Settings → Pages**
2. Under Source, select **Deploy from a branch**
3. Branch: `main` / Folder: `/ (root)`
4. Click **Save**

Your site will be live at:
```
https://YOUR_GITHUB_USERNAME.github.io/asset-library/
```
(Takes 1–2 minutes to go live)

### Step 5 — Update the Manifest
Open `manifest.xml` and replace ALL instances of:
```
YOUR_GITHUB_USERNAME
```
with your actual GitHub username. Then re-upload the updated `manifest.xml`.

---

## PART 2 — Google Drive Setup

### Step 1 — Create 3 Folders in Shared Drive
In your team's Google Drive, create:
- 📁 `Assets - Images`
- 📁 `Assets - Illustrations`
- 📁 `Assets - Icons`

Upload your images to each folder.

### Step 2 — Make Each Folder Public
For **each** of the 3 folders:
1. Right-click the folder → **Share**
2. Under "General access" → change to **Anyone with the link**
3. Set role to **Viewer**
4. Click **Done**

### Step 3 — Copy Each Folder ID
Open each folder. The URL looks like:
```
https://drive.google.com/drive/folders/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs
```
The **Folder ID** is the long string after `/folders/`:
```
1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs
```
Save the 3 folder IDs — you'll need them in the add-in config.

---

## PART 3 — Google API Key (Free)

The add-in uses the Google Drive API to list files. You need a free API key.

### Step 1 — Go to Google Cloud Console
Visit: **console.cloud.google.com**
Sign in with your Google account.

### Step 2 — Create a Project
1. Click the project dropdown (top left) → **New Project**
2. Name it `Asset Library`
3. Click **Create**

### Step 3 — Enable the Drive API
1. Go to **APIs & Services → Enable APIs and Services**
2. Search for **Google Drive API**
3. Click it → click **Enable**

### Step 4 — Create an API Key
1. Go to **APIs & Services → Credentials**
2. Click **Create Credentials → API Key**
3. Copy the key (starts with `AIza…`)

### Step 5 — Restrict the Key (Recommended)
1. Click on the key you just created
2. Under **Application restrictions** → select **HTTP referrers**
3. Add: `https://YOUR_GITHUB_USERNAME.github.io/*`
4. Under **API restrictions** → select **Restrict key** → choose **Google Drive API**
5. Click **Save**

---

## PART 4 — Install the Add-in in PowerPoint (Each User Does This)

### On Windows:
1. Open PowerPoint
2. Go to **File → Options → Trust Center → Trust Center Settings**
3. Click **Trusted Add-in Catalogs**
4. — OR — easier method:
   - Go to **Insert → Get Add-ins → My Add-ins**
   - Click **Upload My Add-in**
   - Browse to and select `manifest.xml`
   - Click **Upload**

### On Mac:
1. Open PowerPoint
2. Go to **Insert → My Add-ins**
3. Click **Upload My Add-in** (bottom left)
4. Select the `manifest.xml` file
5. Click **Upload**

The **Asset Library** button will appear in the **Home** tab ribbon.

---

## PART 5 — First-Time Configuration (Each User Does This Once)

1. Click **Asset Library** in the Home tab to open the task pane
2. Click **⚙ Configure** (top right of the pane)
3. Paste in your 3 **Folder IDs** (from Part 2, Step 3)
4. Paste in your **Google API Key** (from Part 3, Step 4)
5. Click **Save & Reload Assets**

Your assets will appear. Click any asset to insert it into the current slide.

---

## Updating Assets

To add/remove assets: simply add or delete files from the Google Drive folders.
The add-in always fetches the latest files — no reinstallation needed.

---

## Sharing With New Team Members

Send them:
1. The `manifest.xml` file
2. This guide (PART 4 + PART 5 only)
3. The 3 Folder IDs and API Key

They follow Part 4 to install, then Part 5 to configure.

---

## Troubleshooting

| Problem | Fix |
|---|---|
| No assets showing | Check that Drive folders are set to "Anyone with link → Viewer" |
| API error | Verify the API key is correct and Drive API is enabled |
| Insert fails | The image may be too large; try smaller files first |
| Add-in doesn't appear | Make sure PowerPoint trusts the add-in; re-upload manifest |
| Blank task pane | Check that GitHub Pages is live (visit the URL in a browser) |
