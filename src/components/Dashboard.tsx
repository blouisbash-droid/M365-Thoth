'use client';

import { useState, useEffect, useMemo, useRef, useCallback } from 'react';
import {
  PublicClientApplication,
  InteractionRequiredAuthError,
} from '@azure/msal-browser';
import type { AccountInfo } from '@azure/msal-browser';
import { SignInLog, FilterState, DateRange } from '@/lib/types';
import { processSignIn, computeStats, getOSLabel } from '@/lib/utils';
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
import { AlertCircle, Loader2, RefreshCw, Shield } from 'lucide-react';
import { cn } from '@/lib/utils';
import { MSAL_CONFIG, GRAPH_SCOPES } from '@/lib/msalConfig';

// ─── Constants ────────────────────────────────────────────────────────────────

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

const SELECT_FIELDS = [
  'id', 'createdDateTime', 'userDisplayName', 'userPrincipalName',
  'appDisplayName', 'clientAppUsed', 'ipAddress', 'location',
  'deviceDetail', 'status', 'conditionalAccessStatus',
  'riskLevelAggregated', 'isInteractive',
].join(',');

// ─── Top non-compliant users ──────────────────────────────────────────────────

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
    pct > 30 ? { cls: 'bg-red-50 border-red-200 text-red-800',      icon: '🔴' } :
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

// ─── Risk level summary ───────────────────────────────────────────────────────

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

// ─── MSAL instance (module-level singleton) ───────────────────────────────────

let msalInstance: PublicClientApplication | null = null;

async function getMsal(): Promise<PublicClientApplication> {
  if (!msalInstance) {
    msalInstance = new PublicClientApplication(MSAL_CONFIG);
    await msalInstance.initialize();
  }
  return msalInstance;
}

// ─── Main Dashboard ───────────────────────────────────────────────────────────

export default function Dashboard() {
  // ── Auth state ─────────────────────────────────────────────────────────────
  const [account, setAccount]     = useState<AccountInfo | null>(null);
  const [msalReady, setMsalReady] = useState(false);
  const [authError, setAuthError] = useState<string | null>(null);

  // ── Data state ─────────────────────────────────────────────────────────────
  const [signIns, setSignIns]         = useState<SignInLog[]>([]);
  const [isInitialLoad, setIsInitial] = useState(true);
  const [isLoadingMore, setLoadMore]  = useState(false);
  const [hasMore, setHasMore]         = useState(false);
  const [nextLink, setNextLink]       = useState<string | null>(null);
  const [error, setError]             = useState<string | null>(null);
  const [filters, setFilters]         = useState<FilterState>(DEFAULT_FILTERS);
  const abortRef                      = useRef<AbortController | null>(null);

  // ── Init MSAL on mount ─────────────────────────────────────────────────────
  useEffect(() => {
    getMsal()
      .then(instance => {
        const accounts = instance.getAllAccounts();
        if (accounts.length > 0) setAccount(accounts[0]);
        setMsalReady(true);
      })
      .catch((err: Error) => {
        setAuthError(`Authentication initialisation failed: ${err.message}`);
        setMsalReady(true);
      });
  }, []);

  // ── Token acquisition ──────────────────────────────────────────────────────
  const getToken = useCallback(async (): Promise<string> => {
    const instance = await getMsal();
    const acc = account;
    if (!acc) throw new Error('Not authenticated');
    try {
      const result = await instance.acquireTokenSilent({ scopes: GRAPH_SCOPES, account: acc });
      return result.accessToken;
    } catch (e) {
      if (e instanceof InteractionRequiredAuthError) {
        // Use redirect instead of popup to avoid nested popup errors
        await instance.acquireTokenRedirect({ scopes: GRAPH_SCOPES });
        // Note: acquireTokenRedirect will redirect and return, code below won't execute
        throw new Error('Token acquisition requires redirect');
      }
      throw e;
    }
  }, [account]);

  // ── Sign in ────────────────────────────────────────────────────────────────
  const handleSignIn = useCallback(async () => {
    setAuthError(null);
    try {
      const instance = await getMsal();
      // Use loginRedirect instead of loginPopup to avoid nested popup issues
      await instance.loginRedirect({ scopes: GRAPH_SCOPES });
    } catch (e) {
      const err = e as Error;
      if (err.name !== 'BrowserAuthError' || !err.message.includes('user_cancelled')) {
        setAuthError(`Sign-in failed: ${err.message}`);
      }
    }
  }, []);

  // ── Sign out ───────────────────────────────────────────────────────────────
  const handleSignOut = useCallback(async () => {
    if (!account) return;
    try {
      const instance = await getMsal();
      // Use logoutRedirect for consistency and to avoid popup issues
      await instance.logoutRedirect({ account });
    } catch { /* logout failed */ }
    setAccount(null);
    setSignIns([]);
  }, [account]);

  // ── Data fetching (direct Graph API calls) ────────────────────────────────
  const loadData = useCallback(async (days: number) => {
    if (abortRef.current) abortRef.current.abort();
    const ctrl = new AbortController();
    abortRef.current = ctrl;

    setIsInitial(true);
    setError(null);
    setSignIns([]);
    setNextLink(null);
    setHasMore(false);

    const since = new Date();
    since.setDate(since.getDate() - days);
    since.setHours(0, 0, 0, 0);
    const filterParam = encodeURIComponent(`createdDateTime ge ${since.toISOString()}`);
    const baseUrl = `https://graph.microsoft.com/v1.0/auditLogs/signIns?$top=999&$select=${SELECT_FIELDS}&$filter=${filterParam}`;

    let accumulated: SignInLog[] = [];
    let link: string | null = null;
    let firstPage = true;

    try {
      do {
        const token = await getToken();
        const url: string = link ?? baseUrl;

        const resp = await fetch(url, {
          headers: { Authorization: `Bearer ${token}` },
          signal: ctrl.signal,
        });

        if (!resp.ok) {
          const body = await resp.json().catch(() => ({}));
          if (resp.status === 403) {
            throw new Error(
              'Access denied. Your account needs one of: Global Administrator, Security Administrator, ' +
              'Security Reader, Global Reader, or Reports Reader. Also ensure AuditLog.Read.All ' +
              'admin consent has been granted for this app registration.',
            );
          }
          throw new Error(body?.error?.message ?? `HTTP ${resp.status}`);
        }

        const data = await resp.json();
        const page: SignInLog[] = (data.value ?? []).map(processSignIn);

        accumulated = [...accumulated, ...page];
        link = data['@odata.nextLink'] ?? null;

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
  }, [getToken]);

  // Reload when account signs in or date range changes
  useEffect(() => {
    if (!account || !msalReady) return;
    loadData(DAY_MAP[filters.dateRange]);
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [account, msalReady, filters.dateRange]);

  // Manual "load more"
  const handleLoadMore = useCallback(async () => {
    if (!nextLink || isLoadingMore) return;
    setLoadMore(true);
    try {
      const token = await getToken();
      const resp = await fetch(nextLink, {
        headers: { Authorization: `Bearer ${token}` },
      });
      const data = await resp.json();
      const page: SignInLog[] = (data.value ?? []).map(processSignIn);
      setSignIns(prev => [...prev, ...page]);
      setNextLink(data['@odata.nextLink'] ?? null);
      setHasMore(!!data['@odata.nextLink']);
    } catch { /* swallow */ }
    finally { setLoadMore(false); }
  }, [nextLink, isLoadingMore, getToken]);

  // ── Client-side filtering ──────────────────────────────────────────────────
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

  // ── MSAL initialising ──────────────────────────────────────────────────────
  if (!msalReady) return (
    <div className="min-h-screen flex items-center justify-center bg-slate-50">
      <Loader2 className="w-7 h-7 animate-spin text-blue-500" />
    </div>
  );

  // ── Not signed in → sign-in screen ────────────────────────────────────────
  if (!account) return (
    <div className="min-h-screen flex">
      {/* Left panel */}
      <div className="hidden lg:flex flex-col justify-between w-1/2 bg-gradient-to-br
                      from-blue-700 via-blue-600 to-indigo-700 p-12 text-white">
        <div className="flex items-center gap-3">
          <div className="w-9 h-9 bg-white/20 rounded-xl flex items-center justify-center">
            <Shield className="w-5 h-5" />
          </div>
          <span className="font-semibold text-lg">M365 Conditional Access</span>
        </div>

        <div>
          <h1 className="text-3xl font-bold leading-snug mb-4">
            Understand your sign-in landscape before enforcing device compliance.
          </h1>
          <p className="text-blue-100">
            Pull live sign-in logs from Microsoft Graph and see exactly which users,
            apps, and devices would be affected when you require Entra joined,
            Hybrid joined, or Intune enrolled devices.
          </p>
        </div>

        <p className="text-blue-200 text-xs">
          Requires <strong className="text-white">AuditLog.Read.All</strong> delegated
          permission and one of: Global Admin, Security Admin, Security Reader, Global
          Reader, or Reports Reader.
        </p>
      </div>

      {/* Right panel */}
      <div className="flex-1 flex items-center justify-center p-8 bg-slate-50">
        <div className="w-full max-w-sm">
          {/* Mobile logo */}
          <div className="flex lg:hidden items-center gap-2 mb-8">
            <div className="w-8 h-8 bg-blue-600 rounded-lg flex items-center justify-center">
              <Shield className="w-4 h-4 text-white" />
            </div>
            <span className="font-semibold">M365 Conditional Access</span>
          </div>

          <h2 className="text-2xl font-bold text-slate-900 mb-2">Sign in</h2>
          <p className="text-slate-500 text-sm mb-8">
            Use your Microsoft 365 account to access the dashboard.
          </p>

          {authError && (
            <div className="mb-4 flex items-start gap-2 bg-red-50 border border-red-200
                            rounded-lg p-3 text-red-800 text-sm">
              <AlertCircle className="w-4 h-4 mt-0.5 shrink-0" />
              {authError}
            </div>
          )}

          <button
            onClick={handleSignIn}
            className="w-full flex items-center justify-center gap-3 bg-white border
                       border-slate-300 rounded-lg px-4 py-3 text-sm font-medium
                       text-slate-700 hover:bg-slate-50 active:bg-slate-100
                       transition-colors shadow-sm"
          >
            {/* Microsoft logo */}
            <svg viewBox="0 0 21 21" className="w-5 h-5" aria-hidden>
              <rect x="1"  y="1"  width="9" height="9" fill="#f25022" />
              <rect x="11" y="1"  width="9" height="9" fill="#7fba00" />
              <rect x="1"  y="11" width="9" height="9" fill="#00a4ef" />
              <rect x="11" y="11" width="9" height="9" fill="#ffb900" />
            </svg>
            Sign in with Microsoft
          </button>
        </div>
      </div>
    </div>
  );

  // ── Authenticated: full dashboard ─────────────────────────────────────────
  return (
    <div className="min-h-screen bg-slate-50">
      <Navbar
        userName={account.name}
        userEmail={account.username}
        onSignOut={handleSignOut}
      />

      <main className="max-w-screen-2xl mx-auto px-4 py-6 space-y-5">

        {/* Page header */}
        <div className="flex items-start justify-between gap-4 flex-wrap">
          <div>
            <h1 className="text-xl font-bold text-slate-900">Conditional Access Readiness</h1>
            <p className="text-sm text-slate-500 mt-0.5">
              Sign-in analysis for:{' '}
              <em>Require Entra joined, Hybrid Entra joined, or Intune enrolled device</em>
            </p>
          </div>

          <button
            onClick={() => loadData(DAY_MAP[filters.dateRange])}
            disabled={isInitialLoad || isLoadingMore}
            className="btn btn-secondary shrink-0"
          >
            <RefreshCw className={cn('w-4 h-4', (isInitialLoad || isLoadingMore) && 'animate-spin')} />
            Refresh
          </button>
        </div>

        {/* Loading */}
        {isInitialLoad && (
          <div className="flex flex-col items-center justify-center py-24 gap-3">
            <Loader2 className="w-8 h-8 animate-spin text-blue-500" />
            <p className="text-slate-500 text-sm">Fetching sign-in logs from Microsoft Graph…</p>
          </div>
        )}

        {/* Error */}
        {error && !isInitialLoad && (
          <div className="card p-5 border-red-200 bg-red-50 flex items-start gap-3">
            <AlertCircle className="w-5 h-5 text-red-500 mt-0.5 shrink-0" />
            <div>
              <p className="text-sm font-semibold text-red-700">Failed to load sign-in logs</p>
              <p className="text-sm text-red-600 mt-0.5">{error}</p>
              <button
                onClick={() => loadData(DAY_MAP[filters.dateRange])}
                className="btn btn-secondary text-xs mt-2"
              >
                Retry
              </button>
            </div>
          </div>
        )}

        {/* Dashboard content */}
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
