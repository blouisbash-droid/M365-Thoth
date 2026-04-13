/**
 * Tenant storage — supports two backends, in priority order:
 *
 * 1. Upstash Redis  (UPSTASH_REDIS_REST_URL + UPSTASH_REDIS_REST_TOKEN)
 *    Set up via Vercel Marketplace → Storage → Upstash Redis.
 *    Enables full dynamic add / edit / delete without redeployment.
 *
 * 2. TENANTS_CONFIG env var  (JSON array of TenantConfig objects, without the
 *    `id` / `createdAt` fields that the UI manages — or with them if you copy
 *    the full JSON from the dashboard).
 *    Read-only: write operations return the updated JSON for you to paste into
 *    your Vercel env vars and redeploy.
 */
import type { TenantConfig } from './types';

const REDIS_KEY = 'ca-dashboard:tenants';

// ─── Redis helpers ────────────────────────────────────────────────────────────

export function isRedisConfigured(): boolean {
  return !!(
    (process.env.UPSTASH_REDIS_REST_URL && process.env.UPSTASH_REDIS_REST_TOKEN)
  );
}

async function redisRead(): Promise<TenantConfig[] | null> {
  if (!isRedisConfigured()) return null;
  try {
    const { Redis } = await import('@upstash/redis');
    const redis = new Redis({
      url: process.env.UPSTASH_REDIS_REST_URL!,
      token: process.env.UPSTASH_REDIS_REST_TOKEN!,
    });
    const result = await redis.get<TenantConfig[]>(REDIS_KEY);
    return result ?? [];
  } catch {
    return null;
  }
}

async function redisWrite(tenants: TenantConfig[]): Promise<void> {
  const { Redis } = await import('@upstash/redis');
  const redis = new Redis({
    url: process.env.UPSTASH_REDIS_REST_URL!,
    token: process.env.UPSTASH_REDIS_REST_TOKEN!,
  });
  await redis.set(REDIS_KEY, tenants);
}

// ─── Env-var fallback ─────────────────────────────────────────────────────────

function envTenants(): TenantConfig[] {
  try {
    const raw = process.env.TENANTS_CONFIG;
    if (!raw) return [];
    const parsed = JSON.parse(raw) as Partial<TenantConfig>[];
    // Tolerate entries without id/createdAt (hand-crafted env var entries)
    return parsed.map((t, i) => ({
      id: t.id ?? `env-${i}`,
      name: t.name ?? 'Unknown',
      tenantId: t.tenantId ?? '',
      notes: t.notes,
      createdAt: t.createdAt ?? new Date().toISOString(),
    }));
  } catch {
    return [];
  }
}

// ─── Public API ───────────────────────────────────────────────────────────────

export async function listTenants(): Promise<TenantConfig[]> {
  const fromRedis = await redisRead();
  if (fromRedis !== null) {
    // Merge: Redis is authoritative, env-var entries fill in anything not in Redis
    const redisIds = new Set(fromRedis.map(t => t.tenantId));
    const envOnly = envTenants().filter(t => !redisIds.has(t.tenantId));
    return [...fromRedis, ...envOnly];
  }
  return envTenants();
}

export async function getTenantById(id: string): Promise<TenantConfig | null> {
  const all = await listTenants();
  return all.find(t => t.id === id) ?? null;
}

export async function getTenantByAzureId(azureTenantId: string): Promise<TenantConfig | null> {
  const all = await listTenants();
  return all.find(t => t.tenantId === azureTenantId) ?? null;
}

/**
 * Add a new tenant.
 * Returns the saved tenant plus a flag indicating whether a redeployment is
 * needed (i.e. Redis is not configured and the caller should update TENANTS_CONFIG).
 */
export async function addTenant(
  data: Pick<TenantConfig, 'name' | 'tenantId' | 'notes'>,
): Promise<{ tenant: TenantConfig; needsRedeploy: boolean }> {
  const tenant: TenantConfig = {
    ...data,
    id: crypto.randomUUID(),
    createdAt: new Date().toISOString(),
  };

  if (!isRedisConfigured()) {
    return { tenant, needsRedeploy: true };
  }

  const existing = (await redisRead()) ?? [];
  await redisWrite([...existing, tenant]);
  return { tenant, needsRedeploy: false };
}

export async function deleteTenant(id: string): Promise<void> {
  if (!isRedisConfigured()) {
    throw new Error(
      'Upstash Redis is not configured. Remove the tenant manually from your TENANTS_CONFIG environment variable.',
    );
  }
  const all = (await redisRead()) ?? [];
  await redisWrite(all.filter(t => t.id !== id));
}

export async function updateTenant(
  id: string,
  patch: Partial<Pick<TenantConfig, 'name' | 'notes'>>,
): Promise<TenantConfig> {
  if (!isRedisConfigured()) {
    throw new Error('Upstash Redis is not configured. Edit the tenant in your TENANTS_CONFIG environment variable.');
  }
  const all = (await redisRead()) ?? [];
  const idx = all.findIndex(t => t.id === id);
  if (idx === -1) throw new Error('Tenant not found');
  all[idx] = { ...all[idx], ...patch };
  await redisWrite(all);
  return all[idx];
}
