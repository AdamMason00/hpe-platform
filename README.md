# HPE Service Platform

Unified web app for **Hyde Park Equipment** (Kubota dealership — South & North stores, ~30 staff).
Static frontend on **GitHub Pages**, data + file storage on a **Google Apps Script** backend.

Two modules:

| Module | Path | Purpose |
| --- | --- | --- |
| **KPI Incentive Manager** | [`/kpi`](kpi/index.html) | Quarterly KPI scoring, tech efficiency, payroll & bonus calculations, manager KPI bonus + PDF export |
| **Warranty Management** | [`/warranty`](warranty/index.html) | Kubota warranty claim submission & tracking |

---

## File structure

```
hpe-platform/
├── index.html            Landing page + login (Google email or PIN) → module hub
├── shared/
│   ├── config.js         Staff roster, stores, KPI defaults, auth emails, WO assignments, backend IDs
│   ├── auth.js           Tier-1 Google email (managers) + Tier-2 PIN (techs/support); sessionStorage
│   └── api.js            Apps Script client (JSONP GET, text/plain POST)
├── kpi/
│   ├── index.html        KPI app shell (sidebar + pages)
│   ├── styles.css        KPI styles
│   └── app.js            KPI logic (scoring, bonuses, uploads, dashboards, PDF export)
├── warranty/
│   └── index.html        Warranty app (sourced from AdamMason00/hpe-warranty)
├── backend/
│   └── Code.gs           Apps Script backend (stored here for version control; deployed separately)
└── README.md
```

---

## Authentication

**Tier 1 — Google email** (managers / admin). Resolved by the backend `getUser`
action against the authorised roster in `shared/config.js`:

- **Admin:** `adam@`, `bapfelbeck@`, `johnwilliams@` — full access to all data, config, uploads.
- **South Store Manager:** `steve@` (Steve Hayes) — South store only.
- **North Store Manager:** `bill@` (Bill Denison) — North store only.

A manual email fallback is provided so managers are never locked out if the
`getUser` round-trip can't read their Google identity (e.g. cross-account).

**Tier 2 — PIN** (techs / support). Each employee gets a unique 4-digit PIN set
by an admin on the **Staff Roster** page (stored in the Staff blob). Techs see
their own efficiency dashboard; support staff see their assigned open-WO queue.

Sessions persist in `sessionStorage` (`hpe_session`) and are shared across modules.

---

## Backend (Google Apps Script)

`backend/Code.gs` is the source of truth for the deployed Web App. It is **not**
served by GitHub Pages — it is kept here for version control.

- **Sheet:** `1Ljh-Ycf1ut6TyV2NgXFRrypRUJzLRPpdUjl_yOIHrBw`
- **Drive folder (uploads + backups):** `1m9wv8eaWhAaLe1qZ0T0P35zUmt4NrSPQ`
- **Photos folder:** `1kbsKqfQp-Ms4YqwOtWxD2p3JYTiPEUu5`

**GET** (JSONP): `ping`, `getUser`, `loadKPI`, `loadWarranty`, `loadPOS`,
`loadEfficiency`, `loadExclusions`, `loadStaff`
**POST**: `saveKPI`, `saveWarranty`, `savePOS`, `saveEfficiency`,
`saveExclusions`, `saveStaff`, `savePINs`, `uploadFile`, `backup`, `importCSV`

- **Auto-backup:** every save writes a timestamped JSON to `Drive/Backups/{module}/`, keeping the last 30.
- **File uploads:** saved to `Drive/Data Uploads/{Q1–Q4|FYTD}/{Efficiency|POS|Warranty}/`.
- **Change logging:** all saves logged to the `ChangeLog` sheet.
- **WO exclusion history:** logged to the `WO_Exclusion_History` sheet.
- **`initialSetup()`** creates the Drive folder tree and all sheet tabs — run it once.

### Deploying the backend

1. Open the bound Apps Script project for the sheet above (Extensions → Apps Script),
   or create a standalone project and paste `backend/Code.gs`.
2. Run **`initialSetup()`** once and authorise the requested scopes.
3. **Deploy → New deployment → Web app.** Execute as **Me**, access **Anyone with the link**.
4. Copy the `/exec` URL into `shared/config.js` → `BACKEND.API_URL`
   (and into `warranty/index.html` → `API_URL` if it differs). The current value is
   already wired to the existing deployment.

> When you change `Code.gs`, create a **new version** of the deployment (or use
> "Manage deployments → Edit → New version") so the live `/exec` URL picks it up.

---

## KPI scoring

Four KPIs are scored per quarter per store (`scoreKPIs(data) → {kpi1..4, points}`):

1. **Tech Efficiency ≥ 75% AND Comeback ≤ 2%**
2. **Service GM ≥ 78%**
3. **Parts GM ≥ 32%**
4. **Open WIP ≤ 2% of service revenue**

**Manager KPI bonus** (Steve / Bill): `payout = KPIs hit × (annual cap ÷ 4 quarters ÷ 4 KPIs)`.
At a $6,000 cap → $1,500 max/quarter, $375 per KPI (3/4 = $1,125). Year-end growth
bonus is tracked separately on the **Growth Bonus** page.

**Tech bonus pool:** quarterly cap (default $4,927), split 73% techs / 27% support,
distributed by FTE — **only if the store scores ≥ 2/4 KPIs**.

Thresholds are admin-editable on the **Configuration** page and persisted to the backend.

---

## Consolidated uploads (Configuration page)

Three drag-and-drop `.xlsx` zones, parsed client-side with SheetJS and archived to Drive:

1. **Labour Efficiencies** — filtered to `TECH` job code, excludes `NON-BILLABLE`;
   efficiency = Hours Billed ÷ Hours Reported per tech per month. WOs over 100% or
   under 75% are flagged on the manager/tech dashboards. Division `S` = South, `M` = North.
2. **POS / Open Work Orders** — filtered to status `O`; prefixes `WS/WM` = work orders,
   `IS/IM` = invoices. Open > 30 days = warning, > 60 days = critical. Per-WO exclude
   checkboxes + notes persist via `saveExclusions`.
3. **Warranty Work Orders** — archived to Drive (claim parsing lives in the Warranty module).

---

## Local development

It's a static site — open `index.html` with any static server, e.g.:

```powershell
python -m http.server 8080
# then browse http://localhost:8080
```

The backend calls work from `localhost` and from the deployed Pages URL because
GETs use JSONP and POSTs use a CORS-simple `text/plain` body.

---

## Enabling GitHub Pages

GitHub Pages must be turned on for the live site:

1. Repo → **Settings → Pages**.
2. **Build and deployment → Source: Deploy from a branch**.
3. Branch: **`main`**, folder: **`/ (root)`** → **Save**.
4. The site publishes at `https://adammason00.github.io/hpe-platform/`.

---

## Roadmap / known follow-ups

- Warranty **Admin** page refinements (remove its upload zone now that uploads live on
  KPI → Configuration; enlarge the WO-exceptions area with notes + exclude checkboxes).
  The warranty app is currently vendored as-is from `AdamMason00/hpe-warranty`.
- Replace default employee last names once confirmed (see `shared/config.js → DEFAULT_STAFF`).