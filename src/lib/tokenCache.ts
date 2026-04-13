/**
 * App-only (client_credentials) token cache for customer tenants.
 *
 * The MSP's own Azure app credentials (AZURE_AD_CLIENT_ID / AZURE_AD_CLIENT_SECRET)
 * are used to acquire tokens for each customer tenant that has granted admin consent.
 *
 * Tokens are cached in module-level memory (survives across warm Lambda invocations;
 * automatically refreshed on cold starts or 60 seconds before expiry).
 */

interface CachedToken {
  accessToken: string;
  expiresAt: number; // Unix ms timestamp
}

const cache = new Map<string, CachedToken>();
const BUFFER_MS = 60_000; // refresh 1 min before actual expiry

export async function getAppOnlyToken(customerTenantId: string): Promise<string> {
  const cached = cache.get(customerTenantId);
  if (cached && Date.now() < cached.expiresAt - BUFFER_MS) {
    return cached.accessToken;
  }

  const clientId = process.env.AZURE_AD_CLIENT_ID;
  const clientSecret = process.env.AZURE_AD_CLIENT_SECRET;

  if (!clientId || !clientSecret) {
    throw new Error(
      'MSP app credentials (AZURE_AD_CLIENT_ID / AZURE_AD_CLIENT_SECRET) are not configured.',
    );
  }

  const resp = await fetch(
    `https://login.microsoftonline.com/${customerTenantId}/oauth2/v2.0/token`,
    {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        grant_type: 'client_credentials',
        client_id: clientId,
        client_secret: clientSecret,
        scope: 'https://graph.microsoft.com/.default',
      }),
      cache: 'no-store',
    },
  );

  const data = await resp.json();

  if (!resp.ok) {
    // Provide an actionable error when consent hasn't been granted yet
    const errCode: string = data?.error ?? '';
    const errDesc: string = data?.error_description ?? '';

    if (errCode === 'invalid_client' || errCode === 'unauthorized_client') {
      throw new Error(
        `Admin consent has not been granted for this tenant. ` +
        `Share the consent URL with the customer's Global Administrator.`,
      );
    }

    throw new Error(errDesc || `Failed to acquire token for tenant ${customerTenantId}`);
  }

  const entry: CachedToken = {
    accessToken: data.access_token,
    expiresAt: Date.now() + (data.expires_in as number) * 1000,
  };

  cache.set(customerTenantId, entry);
  return entry.accessToken;
}

/** Evict a cached token (e.g. after a permission error). */
export function evictToken(customerTenantId: string): void {
  cache.delete(customerTenantId);
}
