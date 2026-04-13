'use client';

import { useState, useEffect, useMemo, useRef, useCallback } from 'react';
import { SignInLog, FilterState, DateRange, TenantConfig } from '@/lib/types';
import { processSignIn, computeStats, getOSLabel } from '@/lib/utils';
import Link from 'next/link';
import Navbar from './Navbar';
import SummaryCards from './SummaryCards';
import {
  CompliancePieChart,
  DeviceCategoryChart,
  OSDistributionChart,
  TimelineChart,
} from './ComplianceChart';
import FilterBar from './FilterBar';
import SignInTable from './SignInTable';
import { AlertCircle, Loader2, RefreshCw, Building2, ChevronDown } from 'lucide-react';
import { cn } from '@/lib/utils';

// ─── Defaults ─────────────────────────────────────────────────────────────────

const DEFAULT_FILTERS: FilterState = {
  search: '',
  userFilter: '',
  appFilter: '',
  osFilter: '',
  policyStatusFilter: '',
  signInStatusFilter: '',
  dateRange: '7d',
};

const DAY_MAP: Record<DateRange, number> = {
  '1d': 1, '7d': 7, '14d': 14, '30d': 30,
};

// ─── Tenant selector ──────────────────────────────────────────────────────────

interface TenantSelectorProps {
  tenants: TenantConfig[];
  selected: string | null;   // null = MSP own tenant (delegated)
  onChange: (id: string | null) => void;
}

function TenantSelector({ tenants, selected, onChange }: TenantSelectorProps) {
  const [open, setOpen] = useState(false);
  const ref = useRef<HTMLDivElement>(null);

  useEffect(() => {
    function handler(e: MouseEvent) {
      if (ref.current && !ref.current.contains(e.target as Node)) setOpen(false);
    }
    document.addEventListener('mousedown', handler);
    return () => document.removeEventListener('mousedown', handler);
  }, []);

  const current =
    selected == null
      ? { label: 'MSP Tenant (delegated)', sub: 'Your own tenant via signed-in user' }
      : tenants.find(t => t.id === selected)
        ? { label: tenants.find(t => t.id === selected)!.name, sub: tenants.find(t => t.id === selected)!.tenantId }
        : { label: 'Unknown tenant', sub: selected };

  return (
    <div className="relative" ref={ref}>
      <button
        onClick={() => setOpen(o => !o)}
        className="flex items-center gap-2 px-3 py-2 rounded-lg border border-slate-300
                   bg-white text-sm text-slate-700 hover:bg-slate-50 transition-colors
                   min-w-[220px] max-w-xs"
      >
        <Building2 className="w-4 h-4 text-slate-400 shrink-0" />
        <span className="flex-1 text-left truncate">{current.label}</span>
        <ChevronDown className="w-3.5 h-3.5 text-slate-400 shrink-0" />
      </button>

      {open && (
        <div className="absolute left-0 mt-1 w-72 bg-white border border-slate-200 rounded-xl
                        shadow-lg py-1 z-40 max-h-80 overflow-y-auto">
          {/* Own tenant option */}
          <button
            onClick={() => { onChange(null); setOpen(false); }}
            className={cn(
              'w-full text-left px-3 py-2.5 hover:bg-slate-50 transition-colors flex flex-col gap-0.5',
              selected === null && 'bg-blue-50',
            )}
          >
            <span className="text-sm font-medium text-slate-800">MSP Tenant (delegated)</span>
            <span className="text-xs text-slate-400">Your own tenant via signed-in user</span>
          </button>

          {tenants.length > 0 && (
            <div className="border-t border-slate-100 mt-1 pt-1">
              <p className="px-3 py-1 text-xs font-semibold text-slate-400 uppercase tracking-wide">
                Customer Tenants
              </p>
              {tenants.map(t => (
                <button
                  key={t.id}
                  onClick={() => { onChange(t.id); setOpen(false); }}
                  className={cn(
                    'w-full text-left px-3 py-2.5 hover:bg-slate-50 transition-colors flex flex-col gap-0.5',
                    selected === t.id && 'bg-blue-50',
                  )}
                >
                  <span className="text-sm font-medium text-slate-800 truncate">{t.name}</span>
                  <span className="text-xs text-slate-400 font-mono truncate">{t.tenantId}</span>
                </button>
              ))}
            </div>
          )}

          <div className="border-t border-slate-100 mt-1 pt-1 px-3 py-2">
            <Link href="/tenants" className="text-xs text-blue-600 hover:underline">
              + Manage tenants
            </Link>
          </div>
        </div>
      )}
    </div>
  );
}

// ─── Top non-compliant users mini-table ───────────────────────────────────────

function TopFailingUsers({ signIns }: { signIns: SignInLog[] }) {
  const users = useMemo(() => {
    const counts: Record<string, { name: string; upn: string; count: number }> = {};
    signIns
      .filter(s => s.policyStatus === 'fails')
      .forEach(s => {
        const key = s.userPrincipalName || s.userDisplayName;
        if (!counts[key]) counts[key] = { name: s.userDisplayName, upn: s.userPrincipalName, count: 0 };
        counts[key].count++;
      });
    return Object.values(counts).sort((a, b) => b.count - a.count).slice(0, 8);
  }, [signIns]);

  if (users.length === 0) {
    return (
      <div className="card p-5 h-72 flex items-center justify-center">
        <p className="text-sm text-slate-400">No policy failures — all sign-ins pass!</p>
      </div>
    );
  }

  const max = users[0]?.count ?? 1;

  return (
    <div className="card p-5 h-72 overflow-y-auto">
      <h3 className="text-sm font-semibold text-slate-700 mb-1">Top Users — Fails Policy</h3>
      <p className="text-xs text-slate-400 mb-4">Sign-ins from non-compliant devices</p>
      <ul className="space-y-2.5">
        {users.map(u => (
          <li key={u.upn} className="flex items-center gap-3">
            <div className="min-w-0 flex-1">
              <p className="text-xs font-medium text-slate-800 truncate">{u.name || u.upn}</p>
              <p className="text-xs text-slate-400 truncate">{u.upn}</p>
            </div>
            <div className="flex items-center gap-2 shrink-0">
              <div className="w-24 h-1.5 rounded-full bg-slate-100 overflow-hidden">
                <div className="h-full bg-red-400 rounded-full" style={{ width: `${(u.count / max) * 100}%` }} />
              </div>
              <span className="text-xs font-semibold text-slate-600 w-8 text-right">{u.count}</span>
            </div>
          </li>
        ))}
      </ul>
    </div>
  );
}

// ─── Policy impact callout ────────────────────────────────────────────────────

function PolicyImpactCallout({ signIns }: { signIns: SignInLog[] }) {
  const total   = signIns.length;
  const blocked = signIns.filter(s => s.policyStatus !== 'passes').length;
  const pct     = total > 0 ? Math.round((blocked / total) * 100) : 0;

  if (total === 0) return null;

  const severity =
    pct > 30 ? { cls: 'bg-red-50 border-red-200 text-red-800', icon: '🔴' } :
    pct > 10 ? { cls: 'bg-amber-50 border-amber-200 text-amber-800', icon: '🟡' } :
               { cls: 'bg-green-50 border-green-200 text-green-700', icon: '🟢' };

  return (
    <div className={`rounded-xl border p-4 text-sm ${severity.cls}`}>
      <strong>{severity.icon} Policy Impact Estimate:</strong>{' '}
      If you enforce this policy today,{' '}
      <strong>{blocked.toLocaleString()} sign-ins ({pct}%)</strong> in this period would be blocked.{' '}
      {pct > 10
        ? 'Consider enabling report-only mode first to notify affected users before enforcing.'
        : 'Impact looks manageable — consider piloting with a test group before full rollout.'}
    </div>
  );
}

// ─── Risk level mini-table ────────────────────────────────────────────────────

function RiskLevelTable({ signIns }: { signIns: SignInLog[] }) {
  const rows = useMemo(() => {
    const counts: Record<string, number> = {};
    signIns.forEach(s => {
      const level = s.riskLevelAggregated || 'none';
      counts[level] = (counts[level] ?? 0) + 1;
    });
    const order = ['high', 'medium', 'low', 'none', 'hidden', 'unknownFutureValue'];
    return order.filter(k => counts[k] !== undefined).map(level => ({ level, count: counts[level] }));
  }, [signIns]);

  const total = signIns.length;
  const badgeClass: Record<string, string> = {
    high: 'badge-red', medium: 'badge-yellow', low: 'badge-blue', none: 'badge-gray',
  };

  return (
    <ul className="space-y-2">
      {rows.map(({ level, count }) => (
        <li key={level} className="flex items-center gap-3">
          <span className={`badge capitalize ${badgeClass[level] ?? 'badge-gray'}`}>{level}</span>
          <div className="flex-1 h-1.5 rounded-full bg-slate-100 overflow-hidden">
            <div
              className="h-full bg-blue-400 rounded-full"
              style={{ width: total > 0 ? `${(count / total) * 100}%` : '0%' }}
            />
          </div>
          <span className="text-xs font-medium text-slate-600 w-10 text-right">{count.toLocaleString()}</span>
        </li>
      ))}
      {rows.length === 0 && (
        <p className="text-xs text-slate-400">No risk data available (requires Azure AD P2)</p>
      )}
    </ul>
  );
}

// ─── Main Dashboard ───────────────────────────────────────────────────────────

export default function Dashboard() {
  const [tenants, setTenants]           = useState<TenantConfig[]>([]);
  const [selectedTenantId, setSelected] = useState<string | null>(null);
  const [signIns, setSignIns]           = useState<SignInLog[]>([]);
  const [isInitialLoad, setIsInitial]   = useState(true);
  const [isLoadingMore, setLoadMore]    = useState(false);
  const [hasMore, setHasMore]           = useState(false);
  const [nextLink, setNextLink]         = useState<string | null>(null);
  const [error, setError]               = useState<string | null>(null);
  const [filters, setFilters]           = useState<FilterState>(DEFAULT_FILTERS);

  const abortRef = useRef<AbortController | null>(null);

  // ── Load tenant list once on mount ─────────────────────────────────────────

  useEffect(() => {
    fetch('/api/tenants')
      .then(r => r.json())
      .then(d => setTenants(d.tenants ?? []))
      .catch(() => {/* non-fatal */});
  }, []);

  // ── Data fetching ───────────────────────────────────────────────────────────

  const loadData = useCallback(async (days: number, tenantId: string | null) => {
    if (abortRef.current) abortRef.current.abort();
    const ctrl = new AbortController();
    abortRef.current = ctrl;

    setIsInitial(true);
    setError(null);
    setSignIns([]);
    setNextLink(null);
    setHasMore(false);

    let accumulated: SignInLog[] = [];
    let link: string | null = null;
    let firstPage = true;

    // Build base URL — include tenantId when a customer tenant is selected
    const baseUrl = tenantId
      ? `/api/signins?days=${days}&tenantId=${encodeURIComponent(tenantId)}`
      : `/api/signins?days=${days}`;

    try {
      do {
        const url: string = link
          ? `/api/signins?nextLink=${encodeURIComponent(link)}${tenantId ? `&tenantId=${encodeURIComponent(tenantId)}` : ''}`
          : baseUrl;

        const resp = await fetch(url, { signal: ctrl.signal });

        if (!resp.ok) {
          const body = await resp.json().catch(() => ({}));
          throw new Error(body.error ?? `HTTP ${resp.status}`);
        }

        const data = await resp.json();
        const page: SignInLog[] = (data.value ?? []).map(processSignIn);

        accumulated = [...accumulated, ...page];
        link = data.nextLink ?? null;

        setSignIns(accumulated);
        setNextLink(link);
        setHasMore(!!link && accumulated.length < 5000);

        if (firstPage) {
          setIsInitial(false);
          if (link) setLoadMore(true);
          firstPage = false;
        }
      } while (link && accumulated.length < 5000 && !ctrl.signal.aborted);
    } catch (err: unknown) {
      if ((err as Error).name === 'AbortError') return;
      setError((err as Error).message ?? 'An unexpected error occurred');
      setIsInitial(false);
    } finally {
      if (!ctrl.signal.aborted) setLoadMore(false);
    }
  }, []);

  // Reload when date range OR selected tenant changes
  useEffect(() => {
    loadData(DAY_MAP[filters.dateRange], selectedTenantId);
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [filters.dateRange, selectedTenantId]);

  // Manual "load more"
  const handleLoadMore = useCallback(async () => {
    if (!nextLink || isLoadingMore) return;
    setLoadMore(true);
    try {
      const url = selectedTenantId
        ? `/api/signins?nextLink=${encodeURIComponent(nextLink)}&tenantId=${encodeURIComponent(selectedTenantId)}`
        : `/api/signins?nextLink=${encodeURIComponent(nextLink)}`;
      const resp = await fetch(url);
      const data = await resp.json();
      const page: SignInLog[] = (data.value ?? []).map(processSignIn);
      setSignIns(prev => [...prev, ...page]);
      setNextLink(data.nextLink ?? null);
      setHasMore(!!data.nextLink);
    } catch { /* swallow */ }
    finally { setLoadMore(false); }
  }, [nextLink, isLoadingMore, selectedTenantId]);

  // ── Client-side filtering ───────────────────────────────────────────────────

  const filteredSignIns = useMemo(() => signIns.filter(s => {
    if (filters.search) {
      const q = filters.search.toLowerCase();
      const hit =
        (s.userDisplayName ?? '').toLowerCase().includes(q) ||
        (s.userPrincipalName ?? '').toLowerCase().includes(q) ||
        (s.appDisplayName ?? '').toLowerCase().includes(q) ||
        (s.deviceDetail.displayName ?? '').toLowerCase().includes(q);
      if (!hit) return false;
    }
    if (filters.userFilter && s.userPrincipalName !== filters.userFilter) return false;
    if (filters.appFilter  && s.appDisplayName    !== filters.appFilter)  return false;
    if (filters.osFilter) {
      if (getOSLabel(s.deviceDetail.operatingSystem).toLowerCase() !== filters.osFilter.toLowerCase()) return false;
    }
    if (filters.policyStatusFilter && s.policyStatus !== filters.policyStatusFilter) return false;
    if (filters.signInStatusFilter) {
      const success = s.status.errorCode === 0;
      if (filters.signInStatusFilter === 'success' && !success) return false;
      if (filters.signInStatusFilter === 'failure' && success) return false;
    }
    return true;
  }), [signIns, filters]);

  const stats = useMemo(() => computeStats(signIns), [signIns]);

  const filterOptions = useMemo(() => ({
    users:     [...new Set(signIns.map(s => s.userPrincipalName).filter(Boolean))].sort() as string[],
    apps:      [...new Set(signIns.map(s => s.appDisplayName).filter(Boolean))].sort() as string[],
    osSystems: [...new Set(signIns.map(s => getOSLabel(s.deviceDetail.operatingSystem)))].sort(),
  }), [signIns]);

  // Current tenant label for the header
  const tenantLabel = selectedTenantId
    ? tenants.find(t => t.id === selectedTenantId)?.name ?? 'Customer Tenant'
    : 'MSP Tenant';

  // ── Render ───────────────────────────────────────────────────────────────────

  return (
    <div className="min-h-screen bg-slate-50">
      <Navbar />

      <main className="max-w-screen-2xl mx-auto px-4 py-6 space-y-5">

        {/* Page header */}
        <div className="flex items-start justify-between gap-4 flex-wrap">
          <div>
            <h1 className="text-xl font-bold text-slate-900">
              Conditional Access Readiness
              <span className="ml-2 text-base font-normal text-slate-400">· {tenantLabel}</span>
            </h1>
            <p className="text-sm text-slate-500 mt-0.5">
              Sign-in analysis for: <em>Require Entra joined, Hybrid Entra joined, or Intune enrolled device</em>
            </p>
          </div>

          <div className="flex items-center gap-2 flex-wrap">
            {/* Tenant selector */}
            <TenantSelector
              tenants={tenants}
              selected={selectedTenantId}
              onChange={id => {
                setSelected(id);
                // Reset per-tenant filter options when switching tenants
                setFilters(f => ({ ...f, userFilter: '', appFilter: '', osFilter: '' }));
              }}
            />

            <button
              onClick={() => loadData(DAY_MAP[filters.dateRange], selectedTenantId)}
              disabled={isInitialLoad || isLoadingMore}
              className="btn btn-secondary shrink-0"
            >
              <RefreshCw className={cn('w-4 h-4', (isInitialLoad || isLoadingMore) && 'animate-spin')} />
              Refresh
            </button>
          </div>
        </div>

        {/* ── Loading ──────────────────────────────────────────────────────── */}
        {isInitialLoad && (
          <div className="flex flex-col items-center justify-center py-24 gap-3">
            <Loader2 className="w-8 h-8 animate-spin text-blue-500" />
            <p className="text-slate-500 text-sm">
              Fetching sign-in logs from Microsoft Graph
              {selectedTenantId ? ` · ${tenantLabel}` : ''}…
            </p>
          </div>
        )}

        {/* ── Error ───────────────────────────────────────────────────────── */}
        {error && !isInitialLoad && (
          <div className="card p-5 border-red-200 bg-red-50 flex items-start gap-3">
            <AlertCircle className="w-5 h-5 text-red-500 mt-0.5 shrink-0" />
            <div>
              <p className="text-sm font-semibold text-red-700">Failed to load sign-in logs</p>
              <p className="text-sm text-red-600 mt-0.5">{error}</p>
              <div className="flex gap-2 mt-2">
                <button onClick={() => loadData(DAY_MAP[filters.dateRange], selectedTenantId)} className="btn btn-secondary text-xs">
                  Retry
                </button>
                {selectedTenantId && (
                  <Link href="/tenants" className="btn btn-secondary text-xs">
                    Check Tenant Setup
                  </Link>
                )}
              </div>
            </div>
          </div>
        )}

        {/* ── Dashboard content ────────────────────────────────────────────── */}
        {!isInitialLoad && !error && (
          <>
            <SummaryCards stats={stats} isLoadingMore={isLoadingMore} />
            <PolicyImpactCallout signIns={signIns} />

            <div className="grid grid-cols-1 md:grid-cols-2 gap-5">
              <CompliancePieChart stats={stats} />
              <DeviceCategoryChart stats={stats} />
            </div>

            <div className="grid grid-cols-1 md:grid-cols-2 gap-5">
              <TimelineChart signIns={signIns} />
              <OSDistributionChart signIns={signIns} />
            </div>

            <div className="grid grid-cols-1 md:grid-cols-2 gap-5">
              <TopFailingUsers signIns={signIns} />
              <div className="card p-5 h-72 overflow-y-auto">
                <h3 className="text-sm font-semibold text-slate-700 mb-1">Risk Level Summary</h3>
                <p className="text-xs text-slate-400 mb-4">Sign-ins by Entra ID risk level</p>
                <RiskLevelTable signIns={signIns} />
              </div>
            </div>

            <FilterBar
              filters={filters}
              onChange={setFilters}
              options={filterOptions}
              totalShown={filteredSignIns.length}
              totalLoaded={signIns.length}
            />

            <SignInTable
              signIns={filteredSignIns}
              isLoadingMore={isLoadingMore}
              hasMore={hasMore}
              onLoadMore={handleLoadMore}
            />
          </>
        )}
      </main>
    </div>
  );
}
