/**
 * Azure AD app registration settings.
 *
 * Register a Single-page application (SPA) in Azure Portal:
 *   https://portal.azure.com → Entra ID → App registrations → New registration
 *
 * Required settings:
 *   - Supported account types: "Accounts in any organizational directory" (multi-tenant)
 *     OR "Accounts in this organizational directory only" (single tenant)
 *   - Platform: Single-page application (SPA)
 *   - Redirect URI: the URL you deploy this app to (e.g. https://yourapp.vercel.app)
 *   - Delegated permission: AuditLog.Read.All  (grant admin consent)
 *
 * The Client ID is NOT a secret — it is safe to commit to source control.
 * No client secret is used (PKCE flow).
 */

/** Your app registration's Application (client) ID */
export const AZURE_CLIENT_ID = 'YOUR_CLIENT_ID_HERE';

/**
 * Authority controls which accounts can sign in:
 *   'common'        → any work/school Microsoft account (multi-tenant)
 *   'organizations' → any Azure AD work/school account
 *   '<tenant-id>'   → only users from your specific tenant
 */
export const AZURE_AUTHORITY = 'https://login.microsoftonline.com/common';
