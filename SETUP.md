# AMP Archive — cloud setup (nothing runs locally)

The website rebuilds **in GitHub's cloud** by reading the OneDrive archive through
the Microsoft Graph API. Nobody installs anything; you click a button on GitHub
once or twice a quarter and the site updates itself.

## One-time setup

### 1. Register an app in Microsoft Entra (needs an admin — IT)
1. Go to **entra.microsoft.com** → **App registrations** → **New registration**.
2. Name: `AMP Archive Reader`. Accounts: *this organizational directory only*. Register.
3. Copy the **Application (client) ID** and **Directory (tenant) ID**.
4. **Certificates & secrets → New client secret** → copy the **Value** (shown once).
5. **API permissions → Add → Microsoft Graph → Application permissions →** add
   **`Files.Read.All`** → then **"Grant admin consent"** (this button needs an admin).

> The admin-consent click is the only step that requires IT. Everything else you can do.

### 2. Add the credentials as GitHub Secrets
In this repo: **Settings → Secrets and variables → Actions → New repository secret.**
Add four:

| Secret name | Value |
|---|---|
| `AZURE_CLIENT_ID` | Application (client) ID |
| `AZURE_TENANT_ID` | Directory (tenant) ID |
| `AZURE_CLIENT_SECRET` | the client secret **Value** |
| `SHARE_URL` | the OneDrive share link to the archive folder |

> Secrets are encrypted and never appear in the code or the site. Never paste them into files.

### 3. Turn on GitHub Pages
**Settings → Pages → Source: Deploy from a branch → Branch `main`, Folder `/docs` → Save.**
Your site will be at `https://adamgodina.github.io/amp-archive-test/`.

## Running it (every quarter)
1. Faculty add new files to the OneDrive archive folder (any time).
2. Go to the **Actions** tab → **"Publish AMP Archive"** → **Run workflow**.
3. ~1–2 minutes later the site reflects the current OneDrive contents. Done.

## Notes
- Reads file **names, folders, and links only** — never file contents, nothing downloaded.
- File links open each document directly in OneDrive (viewers still need access to the folder).
- Whoever clicks "Run workflow" must be a collaborator on this repo. Faculty who only
  add files never need GitHub access.
- Local scripts (`_db_build/`, the `.command` launchers) are optional and kept off GitHub;
  the cloud workflow is the source of truth.
