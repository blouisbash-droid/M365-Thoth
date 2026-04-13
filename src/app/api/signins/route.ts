import { getServerSession } from 'next-auth';
import { authOptions } from '@/lib/auth';
import { NextRequest, NextResponse } from 'next/server';
import { getTenantById } from '@/lib/tenants';
import { getAppOnlyToken, evictToken } from '@/lib/tokenCache';

// Fields we request from MS Graph — keeps payload lean
const SELECT_FIELDS = [
  'id',
  'createdDateTime',
  'userDisplayName',
  'userPrincipalName',
  'appDisplayName',
  'clientAppUsed',
  'ipAddress',
  'location',
  'deviceDetail',
  'status',
  'conditionalAccessStatus',
  'riskLevelAggregated',
  'isInteractive',
].join(',');

export async function GET(request: NextRequest) {
  const session = await getServerSession(authOptions);
  if (!session) {
    return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });
  }

  const { searchParams } = new URL(request.url);
  const days          = Math.min(parseInt(searchParams.get('days') ?? '7', 10), 30);
  const nextLink      = searchParams.get('nextLink');
  const tenantRecordId = searchParams.get('tenantId'); // internal UUID from our tenant list

  // ── Resolve access token ────────────────────────────────────────────────────
  let accessToken: string;
  let azureTenantId: string | undefined;

  if (tenantRecordId) {
    // Customer tenant — use app-only token
    const tenantConfig = await getTenantById(tenantRecordId);
    if (!tenantConfig) {
      return NextResponse.json({ error: 'Tenant not found' }, { status: 404 });
    }
    azureTenantId = tenantConfig.tenantId;
    try {
      accessToken = await getAppOnlyToken(tenantConfig.tenantId);
    } catch (err: unknown) {
      return NextResponse.json(
        { error: (err as Error).message },
        { status: 403 },
      );
    }
  } else {
    // MSP's own tenant — use the signed-in user's delegated token
    if (!session.accessToken) {
      return NextResponse.json({ error: 'No access token in session' }, { status: 401 });
    }
    accessToken = session.accessToken;
  }

  // ── Build Graph URL ─────────────────────────────────────────────────────────
  let graphUrl: string;

  if (nextLink) {
    graphUrl = nextLink;
  } else {
    const since = new Date();
    since.setDate(since.getDate() - days);
    since.setHours(0, 0, 0, 0);
    const filter = encodeURIComponent(`createdDateTime ge ${since.toISOString()}`);
    graphUrl = `https://graph.microsoft.com/v1.0/auditLogs/signIns?$top=999&$select=${SELECT_FIELDS}&$filter=${filter}`;
  }

  // ── Call Graph ──────────────────────────────────────────────────────────────
  try {
    const resp = await fetch(graphUrl, {
      headers: {
        Authorization: `Bearer ${accessToken}`,
        'Content-Type': 'application/json',
      },
      cache: 'no-store',
    });

    if (!resp.ok) {
      const body = await resp.json().catch(() => ({}));
      const code: string = body?.error?.code ?? '';

      // Evict cached token on auth errors so the next request fetches a fresh one
      if (azureTenantId && (resp.status === 401 || resp.status === 403)) {
        evictToken(azureTenantId);
      }

      if (resp.status === 403 || code === 'Authorization_RequestDenied') {
        return NextResponse.json(
          {
            error:
              tenantRecordId
                ? 'Access denied. Ensure admin consent has been granted and the app has AuditLog.Read.All application permission.'
                : 'Access denied. Your account needs one of: Global Administrator, Security Administrator, Security Reader, Global Reader, or Reports Reader.',
          },
          { status: 403 },
        );
      }

      return NextResponse.json(
        { error: body?.error?.message ?? 'Failed to fetch sign-in logs from MS Graph' },
        { status: resp.status },
      );
    }

    const data = await resp.json();
    return NextResponse.json({
      value: data.value ?? [],
      nextLink: data['@odata.nextLink'] ?? null,
    });
  } catch (err) {
    console.error('Unexpected error fetching sign-in logs', err);
    return NextResponse.json({ error: 'Internal server error' }, { status: 500 });
  }
}
