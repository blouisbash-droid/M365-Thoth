# M365 Conditional Access Dashboard

An interactive, Vercel-hosted dashboard that pulls live sign-in audit logs from **Microsoft Graph** and shows you exactly which users, applications, and devices would be affected when you enforce a **device-compliance conditional access policy** requiring:

- **Entra Joined** devices (formerly Azure AD joined)
- **Hybrid Entra Joined** devices (on-prem AD + Entra ID)
- **Intune Enrolled** devices (MDM managed)

---

## Features

| Feature | Details |
|---|---|
| **Live MS Graph data** | Reads `/auditLogs/signIns` with your delegated token |
| **Policy impact estimate** | Instantly see what % of sign-ins would be blocked today |
| **Compliance charts** | Pie, bar, and stacked-timeline charts |
| **OS distribution** | Windows / macOS / iOS / Android / Linux breakdown |
| **Top failing users** | Ranked list of users with the most non-compliant sign-ins |
| **Risk level summary** | Entra ID risk aggregation (P2) |
| **Filter & sort** | By user, application, OS, policy status, sign-in result, date range |
| **Export CSV** | Download the filtered table for reporting or ticketing |
| **Token refresh** | Automatically refreshes the MS Graph token before expiry |

---

## Prerequisites

| Requirement | Notes |
|---|---|
| Microsoft 365 tenant | Any SKU; sign-in log retention requires **Azure AD P1/P2** |
| Azure AD App Registration | See setup below |
| Admin role | Global Admin, Security Admin, Security Reader, Global Reader, or Reports Reader |
| Node.js 18+ | For local development |

---

## 1 — Azure AD App Registration

1. Go to **Azure Portal → Entra ID → App registrations → New registration**.
2. Give it a name (e.g. `M365-CA-Dashboard`).
3. **Redirect URI** — choose **Web** and add:
   - `http://localhost:3000/api/auth/callback/azure-ad` (local dev)
   - `https://your-app.vercel.app/api/auth/callback/azure-ad` (production)
4. After creation, copy the **Application (client) ID** and **Directory (tenant) ID**.
5. Go to **Certificates & secrets → New client secret**, copy the value.
6. Go to **API permissions → Add a permission → Microsoft Graph → Delegated**:
   - `AuditLog.Read.All`
   - `Directory.Read.All`
7. Click **Grant admin consent for <tenant>**.

---

## 2 — Local Development

```bash
# 1. Clone and install
git clone <repo-url>
cd M365-Graph-Reports
npm install

# 2. Create environment file
cp .env.example .env

# 3. Fill in the values in .env
AZURE_AD_CLIENT_ID="..."
AZURE_AD_CLIENT_SECRET="..."
AZURE_AD_TENANT_ID="..."
NEXTAUTH_SECRET="$(openssl rand -base64 32)"
NEXTAUTH_URL="http://localhost:3000"

# 4. Run the dev server
npm run dev
```

Open [http://localhost:3000](http://localhost:3000) and sign in with your Microsoft admin account.

---

## 3 — Deploy to Vercel

### Option A — Vercel CLI

```bash
npm i -g vercel
vercel
```

### Option B — Vercel Dashboard

1. Push this repo to GitHub.
2. Import it in [vercel.com/new](https://vercel.com/new).
3. Add these environment variables under **Project Settings → Environment Variables**:

| Variable | Value |
|---|---|
| `AZURE_AD_CLIENT_ID` | From Azure portal |
| `AZURE_AD_CLIENT_SECRET` | From Azure portal |
| `AZURE_AD_TENANT_ID` | From Azure portal |
| `NEXTAUTH_SECRET` | Run: `openssl rand -base64 32` |
| `NEXTAUTH_URL` | `https://your-app.vercel.app` |

4. Add the production redirect URI in your Azure App Registration:
   `https://your-app.vercel.app/api/auth/callback/azure-ad`

---

## Architecture

```
src/
├── app/
│   ├── api/
│   │   ├── auth/[...nextauth]/   # NextAuth.js handler (Azure AD provider)
│   │   └── signins/              # Proxies MS Graph /auditLogs/signIns
│   ├── dashboard/                # Protected dashboard page (server component)
│   └── page.tsx                  # Login page / redirect
├── components/
│   ├── Dashboard.tsx             # Main client component — data fetching, state
│   ├── SummaryCards.tsx          # KPI cards + device category progress bar
│   ├── ComplianceChart.tsx       # Recharts: pie, bar, timeline, OS charts
│   ├── FilterBar.tsx             # Search + dropdown filters
│   ├── SignInTable.tsx           # TanStack Table with sort, paginate, CSV export
│   ├── Navbar.tsx                # Top navigation with user menu
│   └── LoginPage.tsx             # Microsoft sign-in landing page
└── lib/
    ├── auth.ts                   # NextAuth options + token refresh
    ├── types.ts                  # TypeScript types for Graph data
    └── utils.ts                  # Device category logic, helpers, CSV export
```

## Device Policy Logic

| Trust Type | isManaged | Category | Passes Policy? |
|---|---|---|---|
| `AzureAD` | any | Entra Joined | Yes |
| `ServerAD` | any | Hybrid Entra Joined | Yes |
| `Workplace` | `true` | Intune Enrolled | Yes |
| `Workplace` | `false` | Registered Only | No |
| (none) | (none) | No Device Info | Unknown |

---

## Required MS Graph Scopes

```
openid profile email offline_access AuditLog.Read.All Directory.Read.All
```

All scopes are **delegated** — the signed-in user's permissions apply.
Admin consent is required for `AuditLog.Read.All`.

---

## Data Privacy

- Sign-in data is **never stored server-side**. Each page load fetches fresh data from Microsoft Graph using the signed-in user's session token.
- The app runs entirely within your Vercel deployment; no data leaves your own infrastructure.
