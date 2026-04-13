'use client';

import { useState, useEffect, useMemo, useRef, useCallback } from 'react';
import { SignInLog, FilterState, DateRange } from '@/lib/types';
import {
  processSignIn, computeStats, getOSLabel,
} from '@/lib/utils';
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
import { AlertCircle, Loader2, RefreshCw } from 'lucide-react';

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
                <div
                  className="h-full bg-red-400 rounded-full"
                  style={{ width: `${(u.count / max) * 100}%` }}
                />
              </div>
              <span className="text-xs font-semibold text-slate-600 w-8 text-right">
                {u.count}
              </span>
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

// ─── Main Dashboard ───────────────────────────────────────────────────────────

export default function Dashboard() {
  const [signIns, setSignIns]         = useState<SignInLog[]>([]);
  const [isInitialLoad, setIsInitial] = useState(true);
  const [isLoadingMore, setLoadMore]  = useState(false);
  const [hasMore, setHasMore]         = useState(false);
  const [nextLink, setNextLink]       = useState<string | null>(null);
  const [error, setError]             = useState<string | null>(null);
  const [filters, setFilters]         = useState<FilterState>(DEFAULT_FILTERS);

  // Abort controller ref so we can cancel in-flight fetches on date-range change
  const abortRef = useRef<AbortController | null>(null);

  // ── Data fetching ───────────────────────────────────────────────────────────

  const loadData = useCallback(async (days: number) => {
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

    try {
      // We auto-page through results until we hit 5 000 records or run out
      do {
        const url: string = link
          ? `/api/signins?nextLink=${encodeURIComponent(link)}`
          : `/api/signins?days=${days}`;

        const resp = await fetch(url, { signal: ctrl.signal });

        if (!resp.ok) {
          const body = await resp.json().catch(() => ({}));
          throw new Error(body.error ?? `HTTP ${resp.status}`);
        }

        const data = await resp.json();
        const page: SignInLog[] = (data.value ?? []).map(processSignIn);

        accumulated = [...accumulated, ...page];
        link = data.nextLink ?? null;

        // Show first page immediately so the UI is interactive
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

  // Initial load + reload on date-range change
  useEffect(() => {
    loadData(DAY_MAP[filters.dateRange]);
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [filters.dateRange]);

  // Manual "load more" (when user clicks the button after auto-load hits the 5k cap)
  const handleLoadMore = useCallback(async () => {
    if (!nextLink || isLoadingMore) return;
    setLoadMore(true);

    try {
      const resp = await fetch(`/api/signins?nextLink=${encodeURIComponent(nextLink)}`);
      const data = await resp.json();
      const page: SignInLog[] = (data.value ?? []).map(processSignIn);
      const next = data.nextLink ?? null;

      setSignIns(prev => [...prev, ...page]);
      setNextLink(next);
      setHasMore(!!next);
    } catch {
      /* swallow */
    } finally {
      setLoadMore(false);
    }
  }, [nextLink, isLoadingMore]);

  // ── Client-side filtering ───────────────────────────────────────────────────

  const filteredSignIns = useMemo(() => {
    return signIns.filter(s => {
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
        const osLabel = getOSLabel(s.deviceDetail.operatingSystem).toLowerCase();
        if (osLabel !== filters.osFilter.toLowerCase()) return false;
      }

      if (filters.policyStatusFilter && s.policyStatus !== filters.policyStatusFilter) return false;

      if (filters.signInStatusFilter) {
        const success = s.status.errorCode === 0;
        if (filters.signInStatusFilter === 'success' && !success) return false;
        if (filters.signInStatusFilter === 'failure' && success) return false;
      }

      return true;
    });
  }, [signIns, filters]);

  // ── Aggregate stats (always over all loaded data, not just filtered) ─────────

  const stats = useMemo(() => computeStats(signIns), [signIns]);

  // ── Filter-bar options (derived from loaded data) ────────────────────────────

  const filterOptions = useMemo(() => {
    const users     = [...new Set(signIns.map(s => s.userPrincipalName).filter(Boolean))].sort() as string[];
    const apps      = [...new Set(signIns.map(s => s.appDisplayName).filter(Boolean))].sort() as string[];
    const osSystems = [...new Set(signIns.map(s => getOSLabel(s.deviceDetail.operatingSystem)))].sort();
    return { users, apps, osSystems };
  }, [signIns]);

  // ── Render ───────────────────────────────────────────────────────────────────

  return (
    <div className="min-h-screen bg-slate-50">
      <Navbar />

      <main className="max-w-screen-2xl mx-auto px-4 py-6 space-y-5">

        {/* Page header */}
        <div className="flex items-start justify-between gap-4 flex-wrap">
          <div>
            <h1 className="text-xl font-bold text-slate-900">Conditional Access Readiness</h1>
            <p className="text-sm text-slate-500 mt-0.5">
              Sign-in analysis for: <em>Require Entra joined, Hybrid Entra joined, or Intune enrolled device</em>
            </p>
          </div>
          <button
            onClick={() => loadData(DAY_MAP[filters.dateRange])}
            disabled={isInitialLoad || isLoadingMore}
            className="btn btn-secondary shrink-0"
          >
            <RefreshCw className={`w-4 h-4 ${(isInitialLoad || isLoadingMore) ? 'animate-spin' : ''}`} />
            Refresh
          </button>
        </div>

        {/* ── Loading skeleton ─────────────────────────────────────────────── */}
        {isInitialLoad && (
          <div className="flex flex-col items-center justify-center py-24 gap-3">
            <Loader2 className="w-8 h-8 animate-spin text-blue-500" />
            <p className="text-slate-500 text-sm">Fetching sign-in logs from Microsoft Graph…</p>
          </div>
        )}

        {/* ── Error ───────────────────────────────────────────────────────── */}
        {error && !isInitialLoad && (
          <div className="card p-5 border-red-200 bg-red-50 flex items-start gap-3">
            <AlertCircle className="w-5 h-5 text-red-500 mt-0.5 shrink-0" />
            <div>
              <p className="text-sm font-semibold text-red-700">Failed to load sign-in logs</p>
              <p className="text-sm text-red-600 mt-0.5">{error}</p>
              <button
                onClick={() => loadData(DAY_MAP[filters.dateRange])}
                className="mt-2 btn btn-secondary text-xs"
              >
                Retry
              </button>
            </div>
          </div>
        )}

        {/* ── Dashboard content ────────────────────────────────────────────── */}
        {!isInitialLoad && !error && (
          <>
            {/* Summary cards */}
            <SummaryCards stats={stats} isLoadingMore={isLoadingMore} />

            {/* Policy impact callout */}
            <PolicyImpactCallout signIns={signIns} />

            {/* Charts row 1 */}
            <div className="grid grid-cols-1 md:grid-cols-2 gap-5">
              <CompliancePieChart stats={stats} />
              <DeviceCategoryChart stats={stats} />
            </div>

            {/* Charts row 2 */}
            <div className="grid grid-cols-1 md:grid-cols-2 gap-5">
              <TimelineChart signIns={signIns} />
              <div className="grid grid-cols-1 gap-5">
                <OSDistributionChart signIns={signIns} />
              </div>
            </div>

            {/* Top failing users */}
            <div className="grid grid-cols-1 md:grid-cols-2 gap-5">
              <TopFailingUsers signIns={signIns} />
              {/* Spacer column intentionally empty — could add more widgets here */}
              <div className="card p-5 h-72 overflow-y-auto">
                <h3 className="text-sm font-semibold text-slate-700 mb-1">Risk Level Summary</h3>
                <p className="text-xs text-slate-400 mb-4">Sign-ins by Entra ID risk level</p>
                <RiskLevelTable signIns={signIns} />
              </div>
            </div>

            {/* Filter bar */}
            <FilterBar
              filters={filters}
              onChange={setFilters}
              options={filterOptions}
              totalShown={filteredSignIns.length}
              totalLoaded={signIns.length}
            />

            {/* Data table */}
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

// ─── Risk level mini-table ────────────────────────────────────────────────────

function RiskLevelTable({ signIns }: { signIns: SignInLog[] }) {
  const rows = useMemo(() => {
    const counts: Record<string, number> = {};
    signIns.forEach(s => {
      const level = s.riskLevelAggregated || 'none';
      counts[level] = (counts[level] ?? 0) + 1;
    });
    const order = ['high', 'medium', 'low', 'none', 'hidden', 'unknownFutureValue'];
    return order
      .filter(k => counts[k] !== undefined)
      .map(level => ({ level, count: counts[level] }));
  }, [signIns]);

  const total = signIns.length;

  const badgeClass: Record<string, string> = {
    high:   'badge-red',
    medium: 'badge-yellow',
    low:    'badge-blue',
    none:   'badge-gray',
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
          <span className="text-xs font-medium text-slate-600 w-10 text-right">
            {count.toLocaleString()}
          </span>
        </li>
      ))}
      {rows.length === 0 && (
        <p className="text-xs text-slate-400">No risk data available (requires Azure AD P2)</p>
      )}
    </ul>
  );
}
