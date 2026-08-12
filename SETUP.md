# AMP Archive — cloud setup (least-privilege, nothing local)

The website rebuilds **in GitHub's cloud** by reading ONE SharePoint site through
the Microsoft Graph API. No files are downloaded, nothing is installed on anyone's
computer, and **no secret is stored** anywhere.

Access is locked down two ways:
- **Scope:** the app can read **exactly one SharePoint site** (`Sites.Selected`) — nothing else in the tenant.
- **Auth:** GitHub proves its identity with a short-lived **OIDC token** — there is **no client secret** to expire or leak.

## One-time setup

### 1. Create the archive's home (you / a faculty owner)
- Create a **SharePoint site** for the archive (e.g. `https://nuwildcat.sharepoint.com/sites/AMPArchive`).
- Copy the pitch files into its document library. Note the site URL and, if the files
  sit in a subfolder, that folder's name.

### 2. Register the app (you) + two admin actions (IT)
1. **entra.microsoft.com → App registrations → New registration.** Name `AMP Archive Reader`,
   single tenant. Copy the **Application (client) ID** and **Directory (tenant) ID**.
2. **API permissions → Microsoft Graph → Application permissions →** add **`Sites.Selected`**
   → **IT grants admin consent.** *(This grants access to no sites yet — it's harmless by itself.)*
3. **Certificates & secrets → Federated credentials → Add credential → "GitHub Actions deploying
   Azure resources":**
   - Organization: `adamgodina`  •  Repository: `amp-archive-test`  •  Entity: **Branch** `main`
   - This lets this repo authenticate with **no secret**.
4. **IT grants the app read access to just the one site** (Graph `POST /sites/{siteId}/permissions`
   with role `read` for this app). This is the step that actually gives access — to one site only.

### 3. Add repo settings (you)
**Settings → Secrets and variables → Actions → New repository secret**, add:

| Secret | Value |
|---|---|
| `AZURE_CLIENT_ID` | Application (client) ID |
| `AZURE_TENANT_ID` | Directory (tenant) ID |
| `SHAREPOINT_SITE_URL` | e.g. `https://nuwildcat.sharepoint.com/sites/AMPArchive` |
| `ARCHIVE_FOLDER` | subfolder in the library, or leave blank for the whole library |

> Note: there is **no `AZURE_CLIENT_SECRET`** — auth is secret-less via OIDC.

### 4. Turn on Pages (you)
**Settings → Pages → Deploy from a branch → `main` / `/docs` → Save.**
Site: `https://adamgodina.github.io/amp-archive-test/`

## Running it (each quarter)
1. Faculty add files to the SharePoint archive.
2. **Actions tab → "Publish AMP Archive" → Run workflow.**
3. ~1–2 min later the site matches the current archive. File links open each document in SharePoint.

## What IT is actually being asked for
- **Read-only** Graph `Sites.Selected` (consent) — no data access on its own.
- **Read access to one named site.**
- A **federated credential** trust for one GitHub repo — **no secret to manage.**

That's it — no org-wide access, no stored credential.
