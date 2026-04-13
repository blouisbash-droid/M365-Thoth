import { getServerSession } from 'next-auth';
import { authOptions } from '@/lib/auth';
import { NextRequest, NextResponse } from 'next/server';

// Fields we request from MS Graph – keeps payload lean
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

  if (!session?.accessToken) {
    return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });
  }

  const { searchParams } = new URL(request.url);
  const days = Math.min(parseInt(searchParams.get('days') ?? '7', 10), 30);
  const nextLink = searchParams.get('nextLink');

  let graphUrl: string;

  if (nextLink) {
    // MS Graph already gives us a fully-formed continuation URL
    graphUrl = nextLink;
  } else {
    const since = new Date();
    since.setDate(since.getDate() - days);
    since.setHours(0, 0, 0, 0);
    const sinceISO = since.toISOString();

    const filter = encodeURIComponent(`createdDateTime ge ${sinceISO}`);
    graphUrl = `https://graph.microsoft.com/v1.0/auditLogs/signIns?$top=999&$select=${SELECT_FIELDS}&$filter=${filter}`;
  }

  try {
    const resp = await fetch(graphUrl, {
      headers: {
        Authorization: `Bearer ${session.accessToken}`,
        'Content-Type': 'application/json',
      },
      // Disable Next.js data-cache for this server-only fetch
      cache: 'no-store',
    });

    if (!resp.ok) {
      const body = await resp.json().catch(() => ({}));
      console.error('MS Graph error', resp.status, body);

      // Surface a helpful message when the user lacks the required role
      const code: string = body?.error?.code ?? '';
      if (resp.status === 403 || code === 'Authorization_RequestDenied') {
        return NextResponse.json(
          {
            error:
              'Access denied. Your account needs one of: Global Administrator, ' +
              'Security Administrator, Security Reader, Global Reader, or Reports Reader.',
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
