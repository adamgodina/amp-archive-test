#!/usr/bin/env python3
"""
Cloud builder (least-privilege). Reads ONE SharePoint site's document library
via Microsoft Graph and writes docs/index.html. Runs in GitHub Actions.

Access model:
  * Permission: Graph APPLICATION permission `Sites.Selected` (read) -- grants
    NOTHING until an admin grants this app read access to the one archive site.
    So the app can read exactly that site and nothing else in the tenant.
  * Auth: GitHub OIDC federated credential -- NO client secret is stored.
    (If AZURE_CLIENT_SECRET is set, it's used instead, for local testing.)

Environment (set as GitHub Secrets):
  AZURE_CLIENT_ID        app registration Application (client) ID
  AZURE_TENANT_ID        Directory (tenant) ID
  SHAREPOINT_SITE_URL    e.g. https://nuwildcat.sharepoint.com/sites/AMPArchive
  ARCHIVE_FOLDER         (optional) subfolder within the site library, e.g. "AMP Historical"
  SITE_TITLE             (optional) page title
  AZURE_CLIENT_SECRET    (optional) only for local testing; unused in the cloud
"""
import base64, json, os, re, sys, datetime
from urllib.parse import urlparse
try:
    import msal, requests
except ImportError:
    sys.exit("Missing deps. Run: pip install msal requests")

GRAPH = "https://graph.microsoft.com/v1.0"
PITCH_EXTS = {".pptx", ".ppt", ".pdf", ".xlsx", ".xls", ".xlsm", ".xlsb", ".docx", ".doc"}
COMPANY_TICKER = re.compile(r"^(.*?)\s*\(([A-Za-z.&]{1,6})\)\s*$")
DOC_KEYWORDS = [
    ("Presentation", ["presentation", "pitch", "deck", "slides"]),
    ("Model", ["model", "dcf"]),
    ("Report", ["report", "write-up", "writeup", "memo"]),
    ("Update", ["update"]),
    ("Feedback", ["feedback"]),
]


def doc_type(stem):
    s = stem.lower()
    for label, kws in DOC_KEYWORDS:
        if any(k in s for k in kws):
            return label
    return "Document"


def extract(rel_parts, stem, ext):
    company = ticker = year = period = section = None
    for p in reversed(rel_parts[:-1]):
        m = COMPANY_TICKER.match(p)
        if m:
            company = m.group(1).strip() or None
            ticker = m.group(2).upper().replace(".", "")
            break
    for p in rel_parts:
        m = re.search(r"\b(0[1-9]|1[0-2])[-.](20\d\d)\b", p)
        if m:
            period, year = f"{m.group(2)}-{m.group(1)}", m.group(2)
            break
    m = re.search(r"\b(20\d\d)[-._]?(0[1-9]|1[0-2])[-._]?(0[1-9]|[12]\d|3[01])\b", stem)
    if m:
        period, year = f"{m.group(1)}-{m.group(2)}-{m.group(3)}", m.group(1)
    if not year:
        for p in reversed(rel_parts):
            m = re.search(r"\b(20\d\d)\b", p)
            if m:
                year = m.group(1)
                break
    for p in rel_parts:
        pl = p.lower()
        if "pitch" in pl:
            section = "Pitches"
        elif "update" in pl:
            section = "Updates"
    return {"ticker": ticker, "company": company, "year": year, "period": period,
            "section": section, "doc_type": doc_type(stem), "ext": ext.lower().lstrip(".")}


def github_oidc_assertion():
    """Fetch a GitHub Actions OIDC token to prove identity to Entra (no secret)."""
    url = os.environ["ACTIONS_ID_TOKEN_REQUEST_URL"] + "&audience=api://AzureADTokenExchange"
    tok = os.environ["ACTIONS_ID_TOKEN_REQUEST_TOKEN"]
    r = requests.get(url, headers={"Authorization": "Bearer " + tok})
    r.raise_for_status()
    return r.json()["value"]


def get_token():
    cid, tid = os.environ["AZURE_CLIENT_ID"], os.environ["AZURE_TENANT_ID"]
    authority = f"https://login.microsoftonline.com/{tid}"
    secret = os.environ.get("AZURE_CLIENT_SECRET")
    credential = secret if secret else {"client_assertion": github_oidc_assertion()}
    app = msal.ConfidentialClientApplication(cid, authority=authority, client_credential=credential)
    res = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in res:
        sys.exit("Auth failed: " + json.dumps(res.get("error_description", res)))
    return res["access_token"]


def resolve_root(session):
    site_url = os.environ["SHAREPOINT_SITE_URL"].rstrip("/")
    u = urlparse(site_url)
    site = session.get(f"{GRAPH}/sites/{u.netloc}:{u.path}").json()
    if "id" not in site:
        sys.exit("Could not resolve SharePoint site: " + json.dumps(site))
    drive = session.get(f"{GRAPH}/sites/{site['id']}/drive?$select=id").json()
    drive_id = drive["id"]
    folder = os.environ.get("ARCHIVE_FOLDER", "").strip("/")
    root = session.get(f"{GRAPH}/drives/{drive_id}/root:/{folder}" if folder
                       else f"{GRAPH}/drives/{drive_id}/root").json()
    if "id" not in root:
        sys.exit("Could not resolve archive folder: " + json.dumps(root))
    return drive_id, root["id"]


def walk(session, drive_id, root_id):
    files, stack = [], [(root_id, [])]
    while stack:
        item_id, rel = stack.pop()
        url = (f"{GRAPH}/drives/{drive_id}/items/{item_id}/children"
               "?$top=200&$select=id,name,file,folder,webUrl")
        while url:
            data = session.get(url).json()
            for c in data.get("value", []):
                nm = c["name"]
                if "folder" in c:
                    stack.append((c["id"], rel + [nm]))
                elif "file" in c and not nm.lower().endswith("_error.txt"):
                    if os.path.splitext(nm)[1].lower() in PITCH_EXTS:
                        files.append((rel + [nm], nm, c.get("webUrl", "")))
            url = data.get("@odata.nextLink")
    return files


TEMPLATE = r"""<!doctype html><html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1"><title>__TITLE__</title>
<style>
 :root{--bg:#0f1419;--panel:#171d26;--panel2:#1e2632;--line:#2a3340;--text:#e8edf3;--muted:#8a97a8;--accent:#4da3ff;--pill:#243042}
 *{box-sizing:border-box} body{margin:0;font:14px/1.5 -apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;background:var(--bg);color:var(--text)}
 header{padding:20px 24px;border-bottom:1px solid var(--line);background:var(--panel)} h1{margin:0;font-size:20px}
 .sub{color:var(--muted);font-size:13px;margin-left:8px} .stats{margin-top:8px;color:var(--muted);font-size:13px} .stats b{color:var(--accent)}
 .controls{display:flex;flex-wrap:wrap;gap:10px;padding:16px 24px;background:var(--panel2);border-bottom:1px solid var(--line);position:sticky;top:0;z-index:5}
 input[type=search],select{background:var(--panel);color:var(--text);border:1px solid var(--line);border-radius:8px;padding:8px 11px;font-size:13px;outline:none}
 input[type=search]{flex:1;min-width:220px} input:focus,select:focus{border-color:var(--accent)}
 .wrap{padding:0 24px 60px} .count{padding:12px 0;color:var(--muted);font-size:13px}
 table{width:100%;border-collapse:collapse} th,td{text-align:left;padding:9px 12px;border-bottom:1px solid var(--line);vertical-align:top}
 th{position:sticky;top:64px;background:var(--bg);color:var(--muted);font-size:12px;text-transform:uppercase;letter-spacing:.4px;cursor:pointer;white-space:nowrap}
 th:hover{color:var(--text)} tr:hover td{background:var(--panel)} .tk{font-weight:700;color:var(--accent)} .muted{color:var(--muted)}
 .pill{display:inline-block;padding:2px 8px;border-radius:20px;background:var(--pill);font-size:11px;color:var(--muted)}
 a.file{color:var(--accent);text-decoration:none} a.file:hover{text-decoration:underline} .empty{padding:50px;text-align:center;color:var(--muted)}
 .foot{padding:14px 24px;color:var(--muted);font-size:12px;border-top:1px solid var(--line)}
</style></head><body>
<header><h1>__TITLE__ <span class="sub">read-only archive</span></h1><div class="stats" id="stats"></div></header>
<div class="controls">
 <input type="search" id="q" placeholder="Search ticker, company, filename…" autofocus>
 <select id="fSection"><option value="">All sections</option></select>
 <select id="fYear"><option value="">All years</option></select>
 <select id="fType"><option value="">All types</option></select>
</div>
<div class="wrap"><div class="count" id="count"></div>
<table><thead><tr id="head"></tr></thead><tbody id="rows"></tbody></table>
<div class="empty" id="empty" style="display:none">No matches.</div></div>
<div class="foot" id="foot"></div>
<script id="data">window.DATA=__DATA__;window.BUILT=__BUILT__;</script>
<script>
const D=window.DATA||[];
const COLS=[["ticker","Ticker"],["company","Company"],["period","Period"],["section","Section"],["doc_type","Type"],["file","File",1]];
let sk="period",sd=-1;
const uniq=k=>[...new Set(D.map(r=>r[k]).filter(Boolean))];
function fill(id,arr,s){const e=document.getElementById(id);(s?arr.sort():arr).forEach(v=>{const o=document.createElement("option");o.value=o.textContent=v;e.appendChild(o)})}
fill("fSection",uniq("section"),true);fill("fYear",uniq("year").sort().reverse(),false);fill("fType",uniq("doc_type"),true);
document.getElementById("stats").innerHTML=`<b>${D.length}</b> files · <b>${new Set(D.map(r=>r.ticker).filter(Boolean)).size}</b> companies`;
document.getElementById("foot").textContent="Last updated: "+(window.BUILT||"");
const q_=document.getElementById("q"),fS=document.getElementById("fSection"),fY=document.getElementById("fYear"),fT=document.getElementById("fType");
function esc(s){return s?(""+s).replace(/[&<>"]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c])):s}
function head(){document.getElementById("head").innerHTML=COLS.map(c=>`<th data-k="${c[0]}">${c[1]}${!c[2]&&c[0]===sk?(sd<0?" ▼":" ▲"):""}</th>`).join("");
 document.querySelectorAll("#head th").forEach(t=>{const k=t.dataset.k;if(COLS.find(c=>c[0]===k)[2])return;t.onclick=()=>{sk===k?sd*=-1:(sk=k,sd=1);render()}})}
function filtered(){const q=q_.value.toLowerCase(),fs=fS.value,fy=fY.value,ft=fT.value;
 return D.filter(r=>{if(fs&&r.section!==fs)return 0;if(fy&&r.year!==fy)return 0;if(ft&&r.doc_type!==ft)return 0;
  if(q){if(![r.ticker,r.company,r.filename].join(" ").toLowerCase().includes(q))return 0}return 1})}
function render(){head();const rows=filtered().sort((a,b)=>{let x=(a[sk]||"")+"",y=(b[sk]||"")+"";return x<y?-sd:x>y?sd:0});
 document.getElementById("count").textContent=`${rows.length} result${rows.length===1?"":"s"}`;
 document.getElementById("empty").style.display=rows.length?"none":"block";
 document.getElementById("rows").innerHTML=rows.slice(0,3000).map(r=>`<tr>
  <td class="tk">${r.ticker||'<span class=muted>?</span>'}</td><td>${esc(r.company)||'<span class=muted>—</span>'}</td>
  <td class="muted">${r.period||"—"}</td><td><span class="pill">${r.section||"—"}</span></td><td>${r.doc_type||"—"}</td>
  <td><a class="file" href="${r.url||"#"}" target="_blank" title="${esc(r.filename)}">${esc(r.filename)} ↗</a></td></tr>`).join("")}
[q_,fS,fY,fT].forEach(e=>e.addEventListener("input",render));render();
</script></body></html>"""


def main():
    session = requests.Session()
    session.headers["Authorization"] = "Bearer " + get_token()
    drive_id, root_id = resolve_root(session)
    files = walk(session, drive_id, root_id)
    recs = []
    for rel_parts, name, url in files:
        r = extract(rel_parts, os.path.splitext(name)[0], os.path.splitext(name)[1])
        r.update({"filename": name, "rel_path": "/".join(rel_parts), "url": url})
        recs.append(r)
    os.makedirs("docs", exist_ok=True)
    open("docs/.nojekyll", "w").close()
    page = (TEMPLATE
            .replace("__TITLE__", os.environ.get("SITE_TITLE", "AMP Stock Pitch Archive"))
            .replace("__DATA__", json.dumps(recs, separators=(",", ":")))
            .replace("__BUILT__", json.dumps(datetime.date.today().isoformat())))
    with open("docs/index.html", "w") as f:
        f.write(page)
    print(f"Wrote docs/index.html from {len(recs)} files (one SharePoint site, via Graph).")


if __name__ == "__main__":
    main()
