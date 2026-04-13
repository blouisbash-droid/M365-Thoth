'use client';

import {
  PieChart, Pie, Cell, Tooltip, Legend, ResponsiveContainer,
  BarChart, Bar, XAxis, YAxis, CartesianGrid,
} from 'recharts';
import { DashboardStats, SignInLog } from '@/lib/types';
import { formatDateShort, getOSLabel } from '@/lib/utils';
import { useMemo } from 'react';

// ─── Shared colours ───────────────────────────────────────────────────────────

const POLICY_COLORS = {
  passes:  '#22c55e',  // green-500
  fails:   '#ef4444',  // red-500
  unknown: '#f59e0b',  // amber-500
};

const DEVICE_COLORS = {
  'entra-joined':        '#3b82f6',  // blue-500
  'hybrid-entra-joined': '#8b5cf6',  // violet-500
  'intune-enrolled':     '#06b6d4',  // cyan-500
  'registered-only':     '#f97316',  // orange-500
  'no-device':           '#94a3b8',  // slate-400
};

// ─── Compliance donut ─────────────────────────────────────────────────────────

export function CompliancePieChart({ stats }: { stats: DashboardStats }) {
  const data = [
    { name: 'Passes Policy',  value: stats.passes,  color: POLICY_COLORS.passes },
    { name: 'Fails Policy',   value: stats.fails,   color: POLICY_COLORS.fails },
    { name: 'No Device Info', value: stats.unknown, color: POLICY_COLORS.unknown },
  ].filter(d => d.value > 0);

  if (data.length === 0) return <EmptyState />;

  return (
    <div className="card p-5 h-72">
      <h3 className="text-sm font-semibold text-slate-700 mb-1">Policy Compliance Distribution</h3>
      <p className="text-xs text-slate-400 mb-3">
        Would this sign-in pass the planned device CA policy?
      </p>
      <ResponsiveContainer width="100%" height="80%">
        <PieChart>
          <Pie
            data={data}
            cx="50%"
            cy="50%"
            innerRadius="50%"
            outerRadius="75%"
            dataKey="value"
            paddingAngle={2}
          >
            {data.map((entry, i) => (
              <Cell key={i} fill={entry.color} stroke="none" />
            ))}
          </Pie>
          <Tooltip
            formatter={(value: number) => [value.toLocaleString(), 'Sign-ins']}
            contentStyle={{ fontSize: 12, borderRadius: 8, border: '1px solid #e2e8f0' }}
          />
          <Legend iconType="circle" iconSize={8} wrapperStyle={{ fontSize: 12 }} />
        </PieChart>
      </ResponsiveContainer>
    </div>
  );
}

// ─── Device category bar chart ────────────────────────────────────────────────

export function DeviceCategoryChart({ stats }: { stats: DashboardStats }) {
  const data = [
    { name: 'Entra Joined',        value: stats.entraJoined,        color: DEVICE_COLORS['entra-joined'] },
    { name: 'Hybrid Entra',        value: stats.hybridEntraJoined,  color: DEVICE_COLORS['hybrid-entra-joined'] },
    { name: 'Intune Enrolled',     value: stats.intuneEnrolled,     color: DEVICE_COLORS['intune-enrolled'] },
    { name: 'Registered Only',     value: stats.registeredOnly,     color: DEVICE_COLORS['registered-only'] },
    { name: 'No Device',           value: stats.noDevice,           color: DEVICE_COLORS['no-device'] },
  ];

  return (
    <div className="card p-5 h-72">
      <h3 className="text-sm font-semibold text-slate-700 mb-1">Device Category Breakdown</h3>
      <p className="text-xs text-slate-400 mb-3">
        Trust type reported by Microsoft Entra ID
      </p>
      <ResponsiveContainer width="100%" height="80%">
        <BarChart data={data} layout="vertical" margin={{ left: 12, right: 16 }}>
          <CartesianGrid strokeDasharray="3 3" stroke="#f1f5f9" horizontal={false} />
          <XAxis
            type="number"
            tick={{ fontSize: 11, fill: '#94a3b8' }}
            axisLine={false}
            tickLine={false}
          />
          <YAxis
            type="category"
            dataKey="name"
            width={110}
            tick={{ fontSize: 11, fill: '#64748b' }}
            axisLine={false}
            tickLine={false}
          />
          <Tooltip
            formatter={(value: number) => [value.toLocaleString(), 'Sign-ins']}
            contentStyle={{ fontSize: 12, borderRadius: 8, border: '1px solid #e2e8f0' }}
          />
          <Bar dataKey="value" radius={[0, 4, 4, 0]} maxBarSize={20}>
            {data.map((entry, i) => (
              <Cell key={i} fill={entry.color} />
            ))}
          </Bar>
        </BarChart>
      </ResponsiveContainer>
    </div>
  );
}

// ─── OS distribution ─────────────────────────────────────────────────────────

const OS_COLORS: Record<string, string> = {
  Windows: '#3b82f6',
  macOS:   '#8b5cf6',
  iOS:     '#0ea5e9',
  Android: '#22c55e',
  Linux:   '#f97316',
  Other:   '#94a3b8',
  Unknown: '#cbd5e1',
};

export function OSDistributionChart({ signIns }: { signIns: SignInLog[] }) {
  const data = useMemo(() => {
    const counts: Record<string, number> = {};
    signIns.forEach(s => {
      const label = getOSLabel(s.deviceDetail.operatingSystem);
      counts[label] = (counts[label] ?? 0) + 1;
    });
    return Object.entries(counts)
      .map(([name, value]) => ({ name, value, color: OS_COLORS[name] ?? '#94a3b8' }))
      .sort((a, b) => b.value - a.value);
  }, [signIns]);

  if (data.length === 0) return <EmptyState />;

  return (
    <div className="card p-5 h-72">
      <h3 className="text-sm font-semibold text-slate-700 mb-1">OS Distribution</h3>
      <p className="text-xs text-slate-400 mb-3">All sign-ins by operating system</p>
      <ResponsiveContainer width="100%" height="80%">
        <BarChart data={data} margin={{ left: 4, right: 16 }}>
          <CartesianGrid strokeDasharray="3 3" stroke="#f1f5f9" vertical={false} />
          <XAxis
            dataKey="name"
            tick={{ fontSize: 11, fill: '#64748b' }}
            axisLine={false}
            tickLine={false}
          />
          <YAxis
            tick={{ fontSize: 11, fill: '#94a3b8' }}
            axisLine={false}
            tickLine={false}
            width={36}
          />
          <Tooltip
            formatter={(value: number) => [value.toLocaleString(), 'Sign-ins']}
            contentStyle={{ fontSize: 12, borderRadius: 8, border: '1px solid #e2e8f0' }}
          />
          <Bar dataKey="value" radius={[4, 4, 0, 0]} maxBarSize={40}>
            {data.map((entry, i) => (
              <Cell key={i} fill={entry.color} />
            ))}
          </Bar>
        </BarChart>
      </ResponsiveContainer>
    </div>
  );
}

// ─── Timeline (stacked bars by day) ──────────────────────────────────────────

export function TimelineChart({ signIns }: { signIns: SignInLog[] }) {
  const data = useMemo(() => {
    const byDate: Record<string, { date: string; passes: number; fails: number; unknown: number }> = {};

    signIns.forEach(s => {
      const key = formatDateShort(s.createdDateTime);
      if (!byDate[key]) byDate[key] = { date: key, passes: 0, fails: 0, unknown: 0 };
      byDate[key][s.policyStatus]++;
    });

    return Object.values(byDate).sort(
      (a, b) => new Date(a.date).getTime() - new Date(b.date).getTime(),
    );
  }, [signIns]);

  if (data.length === 0) return <EmptyState />;

  return (
    <div className="card p-5 h-72">
      <h3 className="text-sm font-semibold text-slate-700 mb-1">Sign-ins Over Time</h3>
      <p className="text-xs text-slate-400 mb-3">Daily sign-ins by policy outcome</p>
      <ResponsiveContainer width="100%" height="80%">
        <BarChart data={data} margin={{ left: 4, right: 16 }}>
          <CartesianGrid strokeDasharray="3 3" stroke="#f1f5f9" vertical={false} />
          <XAxis
            dataKey="date"
            tick={{ fontSize: 10, fill: '#64748b' }}
            axisLine={false}
            tickLine={false}
          />
          <YAxis
            tick={{ fontSize: 11, fill: '#94a3b8' }}
            axisLine={false}
            tickLine={false}
            width={36}
          />
          <Tooltip
            contentStyle={{ fontSize: 12, borderRadius: 8, border: '1px solid #e2e8f0' }}
          />
          <Legend iconType="circle" iconSize={8} wrapperStyle={{ fontSize: 11 }} />
          <Bar dataKey="passes"  name="Passes Policy"  stackId="a" fill={POLICY_COLORS.passes}  maxBarSize={40} />
          <Bar dataKey="fails"   name="Fails Policy"   stackId="a" fill={POLICY_COLORS.fails}   maxBarSize={40} />
          <Bar dataKey="unknown" name="No Device Info" stackId="a" fill={POLICY_COLORS.unknown} maxBarSize={40} radius={[3, 3, 0, 0]} />
        </BarChart>
      </ResponsiveContainer>
    </div>
  );
}

// ─── Shared empty state ───────────────────────────────────────────────────────

function EmptyState() {
  return (
    <div className="card p-5 h-72 flex items-center justify-center">
      <p className="text-slate-400 text-sm">No data to display</p>
    </div>
  );
}
