import { getServerSession } from 'next-auth';
import { authOptions } from '@/lib/auth';
import { NextRequest, NextResponse } from 'next/server';
import {
  listTenants,
  addTenant,
  deleteTenant,
  updateTenant,
  isRedisConfigured,
} from '@/lib/tenants';

function unauthorized() {
  return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });
}

// GET /api/tenants — list all configured tenants
export async function GET() {
  const session = await getServerSession(authOptions);
  if (!session) return unauthorized();

  const tenants = await listTenants();
  return NextResponse.json({
    tenants,
    storageMode: isRedisConfigured() ? 'redis' : 'env',
  });
}

// POST /api/tenants — add a new tenant
export async function POST(request: NextRequest) {
  const session = await getServerSession(authOptions);
  if (!session) return unauthorized();

  const body = await request.json().catch(() => null);
  if (!body?.name || !body?.tenantId) {
    return NextResponse.json(
      { error: 'name and tenantId are required' },
      { status: 400 },
    );
  }

  // Basic GUID validation for tenantId
  const guidRe = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
  if (!guidRe.test(body.tenantId)) {
    return NextResponse.json(
      { error: 'tenantId must be a valid GUID (xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx)' },
      { status: 400 },
    );
  }

  const { tenant, needsRedeploy } = await addTenant({
    name: String(body.name).trim(),
    tenantId: String(body.tenantId).trim().toLowerCase(),
    notes: body.notes ? String(body.notes).trim() : undefined,
  });

  return NextResponse.json({ tenant, needsRedeploy }, { status: 201 });
}

// PATCH /api/tenants — update name/notes for a tenant
export async function PATCH(request: NextRequest) {
  const session = await getServerSession(authOptions);
  if (!session) return unauthorized();

  const body = await request.json().catch(() => null);
  if (!body?.id) {
    return NextResponse.json({ error: 'id is required' }, { status: 400 });
  }

  try {
    const updated = await updateTenant(body.id, {
      name: body.name,
      notes: body.notes,
    });
    return NextResponse.json({ tenant: updated });
  } catch (err: unknown) {
    return NextResponse.json({ error: (err as Error).message }, { status: 400 });
  }
}

// DELETE /api/tenants — remove a tenant
export async function DELETE(request: NextRequest) {
  const session = await getServerSession(authOptions);
  if (!session) return unauthorized();

  const body = await request.json().catch(() => null);
  if (!body?.id) {
    return NextResponse.json({ error: 'id is required' }, { status: 400 });
  }

  try {
    await deleteTenant(body.id);
    return NextResponse.json({ success: true });
  } catch (err: unknown) {
    return NextResponse.json({ error: (err as Error).message }, { status: 400 });
  }
}
