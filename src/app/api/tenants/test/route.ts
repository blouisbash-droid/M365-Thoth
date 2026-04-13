import { getServerSession } from 'next-auth';
import { authOptions } from '@/lib/auth';
import { NextRequest, NextResponse } from 'next/server';
import { getAppOnlyToken, evictToken } from '@/lib/tokenCache';

/**
 * POST /api/tenants/test
 * Body: { tenantId: string }  ← the Azure tenant ID (GUID), not our internal UUID
 *
 * Tries to acquire an app-only token for the given tenant and calls
 * GET /organization to verify the connection is working.
 * Returns { ok: true, displayName, tenantId } or { ok: false, error }.
 */
export async function POST(request: NextRequest) {
  const session = await getServerSession(authOptions);
  if (!session) {
    return NextResponse.json({ ok: false, error: 'Unauthorized' }, { status: 401 });
  }

  const body = await request.json().catch(() => null);
  const azureTenantId: string = body?.tenantId ?? '';

  if (!azureTenantId) {
    return NextResponse.json({ ok: false, error: 'tenantId is required' }, { status: 400 });
  }

  // 1. Acquire app-only token
  let token: string;
  try {
    token = await getAppOnlyToken(azureTenantId);
  } catch (err: unknown) {
    return NextResponse.json({ ok: false, error: (err as Error).message });
  }

  // 2. Call /organization — lightweight call that confirms consent + permissions
  try {
    const resp = await fetch('https://graph.microsoft.com/v1.0/organization', {
      headers: { Authorization: `Bearer ${token}` },
      cache: 'no-store',
    });

    if (!resp.ok) {
      evictToken(azureTenantId);
      const body2 = await resp.json().catch(() => ({}));
      return NextResponse.json({
        ok: false,
        error: body2?.error?.message ?? `Graph returned HTTP ${resp.status}`,
      });
    }

    const data = await resp.json();
    const org = data.value?.[0];

    return NextResponse.json({
      ok: true,
      displayName: org?.displayName ?? azureTenantId,
      verifiedDomain: org?.verifiedDomains?.find((d: { isDefault: boolean }) => d.isDefault)?.name,
    });
  } catch (err: unknown) {
    evictToken(azureTenantId);
    return NextResponse.json({ ok: false, error: (err as Error).message });
  }
}
