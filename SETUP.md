# SurgaMed Sales Wallboard — Live Public Link Setup

This puts the flashy dashboard on a free public web link that **auto-refreshes every 15 minutes in the cloud** —
no PC required. The link works on the office TV and as your share link.

**Total cost:** $0.  **One-time setup:** ~15–20 minutes.  **Requires a Microsoft 365 admin for Step 2.**

Files in this folder:
- `index.html` – the dashboard (reads `data.json`)
- `data.json` – the live numbers (overwritten hourly by the cloud job)
- `refresh.mjs` – the cloud job that reads your 4 files and rewrites `data.json`
- `package.json` – its one dependency
- `.github/workflows/refresh.yml` – the hourly schedule that runs the job

---

## Step 1 — Put the files on GitHub (5 min)
1. Create a free account at github.com (skip if you have one).
2. Create a **new repository**, e.g. `surgamed-board`. Make it **Public** (required for the free public link).
3. Upload **all** files from this folder, keeping the folder structure — the workflow MUST stay at
   `.github/workflows/refresh.yml`. (Use “Add file → Upload files”, then drag the whole folder.)

## Step 2 — Create the Microsoft “app registration” (admin, ~10 min)
This lets the cloud job read the four files. Do this at **entra.microsoft.com** signed in as a Microsoft 365 admin.
1. **Identity → App registrations → New registration.** Name it `SurgaMed Dashboard`. Register.
2. On the Overview page copy the **Application (client) ID** and the **Directory (tenant) ID**.
3. **Certificates & secrets → New client secret.** Copy the secret **Value** now (it’s shown only once).
4. **API permissions → Add a permission → Microsoft Graph → Application permissions →** search **Files.Read.All** → add.
   Then click **Grant admin consent**. (Files.Read.All lets it read org files; if you prefer tighter scope,
   ask your admin about `Sites.Selected` granted only to the four OneDrive sites.)

## Step 3 — Give GitHub the three secrets (2 min)
In your GitHub repo: **Settings → Secrets and variables → Actions → New repository secret.** Add three:
- `MS_TENANT_ID` = the Directory (tenant) ID
- `MS_CLIENT_ID` = the Application (client) ID
- `MS_CLIENT_SECRET` = the secret Value
(These are stored securely by GitHub and are **not** part of the public page.)

## Step 4 — Turn on the public page (2 min)
**Settings → Pages → Build and deployment → Source: Deploy from a branch → Branch: `main` / `/ (root)` → Save.**
After a minute your link is: **https://YOUR-USERNAME.github.io/surgamed-board/**

## Step 5 — Run it once & confirm
**Actions tab → refresh-dashboard → Run workflow.** It signs in, reads the 4 files, and rewrites `data.json`.
Open your link — you should see live numbers. From now on it refreshes **every 15 minutes** automatically.

## Step 6 — Put it on the office TV
On the screen’s device, open the link in a browser **full-screen / kiosk mode**. The page reloads itself
every 15 minutes, and the data behind it refreshes every 15 minutes. Done.

---

### Adjusting things
- **Schedule:** the `cron` in `refresh.yml` is `*/15 * * * *` = every 15 minutes, all day (UTC, DST-proof). To run only during business hours use e.g. `*/15 13-22 * * *` (≈9am–6pm US Eastern).
- **Team goal:** change `GOAL_30D` near the top of `refresh.mjs`.
- **A rep renames a file or tab:** update that rep’s line in the `REPS` list in `refresh.mjs`.

### Privacy
This is a **public** page — anyone with the link can see it (your choice). It shows customer first names and
revenue. No passwords or emails are exposed, and your Microsoft secrets live only in GitHub Secrets, never on the page.
If you later want it private, GitHub Pages can be restricted, or move to a host with password protection.
