# M365 Conditional Access Dashboard

**Live site: [m365-ca-review.vercel.app](https://m365-ca-review.vercel.app)**

An interactive, MSP-ready dashboard that pulls live sign-in audit logs from **Microsoft Graph** and shows exactly which users, applications, and devices would be affected when you enforce a device-compliance conditional access policy requiring:

- **Entra Joined** devices (formerly Azure AD joined)
- **Hybrid Entra Joined** devices (on-prem AD + Entra ID)
- **Intune Enrolled** devices (MDM managed)

MSPs can connect **multiple customer tenants** — one admin consent click per customer, no customer secrets stored.

---

## Features

| Feature | Details |
|---|---|
| **Multi-tenant** | Switch between customer tenants from the dashboard dropdown |
| **Tenant manager** | Add/remove customers at `/tenants` — generates the consent URL automatically |
| **Live MS Graph data** | App-only tokens per customer tenant; auto-refreshed before expiry |
| **Policy impact estimate** | Instantly see what % of sign-ins would be blocked today |
| **Compliance charts** | Pie, bar, stacked-timeline, and OS distribution charts |
| **Top failing users** | Ranked list of users with the most non-compliant sign-ins |
| **Risk level summary** | Entra ID risk aggregation (requires Azure AD P2) |
| **Filter & sort** | By user, app, OS, policy status, sign-in result, and date range |
| **Export CSV** | Download the filtered table for reporting or ticketing |

---

## MSP Architecture

```
MSP Tenant
  └─ One Azure AD App Registration (multi-tenant)
        ├─ Application permissions: AuditLog.Read.All, Directory.Read.All
        └─ Client credentials used to get app-only tokens per customer

Customer Tenant A  ──(admin consent)──► Service principal created
Customer Tenant B  ──(admin consent)──► Service principal created
Customer Tenant C  ──(admin consent)──► Service principal created
```

No customer credentials are ever stored. The MSP's `client_id` + `client_secret` are used to request tokens scoped to each customer's `tenant_id`.

---

## Setup

### 1 — Azure AD App Registration (MSP tenant)

1. Go to **Azure Portal → Entra ID → App registrations → New registration**
2. Name it (e.g. `M365-CA-Dashboard`)
3. **Supported account types** → **Accounts in any organizational directory (Multi-tenant)**
4. **Redirect URIs** (Web) — add all three:
   - `http://localhost:3000/api/auth/callback/azure-ad`
   - `https://m365-ca-review.vercel.app/api/auth/callback/azure-ad`
   - `https://m365-ca-review.vercel.app/tenants`
5. Copy the **Application (client) ID** and **Directory (tenant) ID**
6. **Certificates & secrets → New client secret** — copy the value
7. **API permissions → Add → Microsoft Graph → Application permissions**:
   - `AuditLog.Read.All`
   - `Directory.Read.All`
8. **API permissions → Add → Microsoft Graph → Delegated permissions**:
   - `AuditLog.Read.All`
   - `Directory.Read.All`
9. Click **Grant admin consent for \<MSP tenant\>**

### 2 — Vercel Deployment

1. Push this repo to GitHub and import at [vercel.com/new](https://vercel.com/new)
2. Set these **Environment Variables** in Project Settings:

| Variable | Value |
|---|---|
| `AZURE_AD_CLIENT_ID` | App registration client ID |
| `AZURE_AD_CLIENT_SECRET` | App registration client secret |
| `AZURE_AD_TENANT_ID` | MSP's Azure AD tenant ID |
| `NEXTAUTH_SECRET` | `openssl rand -base64 32` |
| `NEXTAUTH_URL` | `https://m365-ca-review.vercel.app` |

3. **(Optional but recommended)** Add **Upstash Redis** for dynamic tenant management without redeployment:
   - Vercel dashboard → Project → **Storage** → **Create** → **Upstash Redis**
   - The `UPSTASH_REDIS_REST_URL` and `UPSTASH_REDIS_REST_TOKEN` env vars are injected automatically
   - Without Redis, tenants are managed via the `TENANTS_CONFIG` env var (see below)

### 3 — Local Development

```bash
cp .env.example .env   # fill in your values
npm install
npm run dev            # http://localhost:3000
```

---

## Adding a Customer Tenant

1. Navigate to **Tenants** in the top nav
2. Click **Add Tenant** — enter the customer's display name and Azure AD Tenant ID
   *(found in their Azure Portal → Entra ID → Overview)*
3. Copy the generated **Admin Consent URL** and send it to the customer's **Global Administrator**
4. The customer admin visits the URL, signs in, and clicks **Accept**
5. Click **Test** on the tenant card to confirm the connection is working
6. The customer's tenant now appears in the dashboard tenant selector

---

## Tenant Storage Options

| Option | Setup | Add/remove tenants |
|---|---|---|
| **Upstash Redis** (recommended) | Add via Vercel Storage tab | Instant, no redeploy needed |
| **`TENANTS_CONFIG` env var** | Paste JSON into Vercel env vars | Requires redeploy |

The `/tenants` page generates the correct JSON for `TENANTS_CONFIG` automatically — just copy and paste.

---

## Architecture

```
src/
├── app/
│   ├── api/
│   │   ├── auth/[...nextauth]/   # NextAuth Azure AD provider
│   │   ├── signins/              # MS Graph proxy — delegated OR app-only token
│   │   └── tenants/              # Tenant CRUD + /test connection endpoint
│   ├── dashboard/                # Protected dashboard (server component)
│   ├── tenants/                  # Tenant management page
│   └── page.tsx                  # Login / redirect
├── components/
│   ├── Dashboard.tsx             # Data fetching, tenant selector, all charts
│   ├── TenantManager.tsx         # Add/test/remove tenants UI
│   ├── SummaryCards.tsx          # KPI cards + device breakdown bar
│   ├── ComplianceChart.tsx       # Recharts: pie, bar, timeline, OS
│   ├── FilterBar.tsx             # Search + filter dropdowns
│   ├── SignInTable.tsx           # TanStack Table — sort, paginate, CSV export
│   ├── Navbar.tsx                # Navigation with Dashboard / Tenants links
│   └── LoginPage.tsx             # Microsoft sign-in landing
└── lib/
    ├── auth.ts                   # NextAuth options + token refresh
    ├── tenants.ts                # Tenant CRUD — Upstash Redis + env var fallback
    ├── tokenCache.ts             # App-only token cache (client_credentials)
    ├── types.ts                  # TypeScript types
    └── utils.ts                  # Device logic, CSV export, helpers
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

## Data Privacy

- Sign-in data is **never stored server-side** — fetched fresh from Microsoft Graph on each load
- Customer credentials are **never stored** — only the Azure tenant ID is saved; the MSP's app credentials handle all token acquisition
- The app runs entirely within your Vercel deployment
