'use client';

import { FilterState, DateRange } from '@/lib/types';
import { Search, X, SlidersHorizontal } from 'lucide-react';
import { cn } from '@/lib/utils';

interface FilterOptions {
  users: string[];
  apps: string[];
  osSystems: string[];
}

interface Props {
  filters: FilterState;
  onChange: (f: FilterState) => void;
  options: FilterOptions;
  totalShown: number;
  totalLoaded: number;
}

function Select({
  label,
  value,
  options,
  onChange,
}: {
  label: string;
  value: string;
  options: { value: string; label: string }[];
  onChange: (v: string) => void;
}) {
  return (
    <select
      value={value}
      onChange={e => onChange(e.target.value)}
      className={cn(
        'text-sm rounded-lg border border-slate-300 px-2.5 py-1.5',
        'bg-white text-slate-700 focus:outline-none focus:ring-2 focus:ring-blue-500',
        'max-w-[180px] truncate',
        value !== '' && 'border-blue-400 bg-blue-50 text-blue-700',
      )}
    >
      <option value="">{label}</option>
      {options.map(o => (
        <option key={o.value} value={o.value}>
          {o.label}
        </option>
      ))}
    </select>
  );
}

const DATE_OPTIONS: { value: DateRange; label: string }[] = [
  { value: '1d',  label: 'Last 24 h' },
  { value: '7d',  label: 'Last 7 days' },
  { value: '14d', label: 'Last 14 days' },
  { value: '30d', label: 'Last 30 days' },
];

const POLICY_OPTIONS = [
  { value: 'passes',  label: 'Passes Policy' },
  { value: 'fails',   label: 'Fails Policy' },
  { value: 'unknown', label: 'No Device Info' },
];

const SIGNIN_OPTIONS = [
  { value: 'success', label: 'Success' },
  { value: 'failure', label: 'Failure' },
];

export default function FilterBar({ filters, onChange, options, totalShown, totalLoaded }: Props) {
  const set = (partial: Partial<FilterState>) => onChange({ ...filters, ...partial });

  const hasActiveFilters =
    filters.search ||
    filters.userFilter ||
    filters.appFilter ||
    filters.osFilter ||
    filters.policyStatusFilter ||
    filters.signInStatusFilter;

  // Normalised OS options from the loaded data
  const osOptions = options.osSystems.map(os => ({ value: os, label: os }));

  // User options — truncate display to email
  const userOptions = options.users.map(u => ({ value: u, label: u }));

  // App options
  const appOptions = options.apps.map(a => ({ value: a, label: a }));

  return (
    <div className="card p-4 space-y-3">
      <div className="flex flex-wrap items-center gap-3">
        <SlidersHorizontal className="w-4 h-4 text-slate-400 shrink-0" />

        {/* Search */}
        <div className="relative">
          <Search className="absolute left-2.5 top-1/2 -translate-y-1/2 w-3.5 h-3.5 text-slate-400 pointer-events-none" />
          <input
            type="text"
            placeholder="Search user, app, device…"
            value={filters.search}
            onChange={e => set({ search: e.target.value })}
            className={cn(
              'pl-8 pr-3 py-1.5 text-sm rounded-lg border border-slate-300',
              'bg-white text-slate-700 placeholder:text-slate-400',
              'focus:outline-none focus:ring-2 focus:ring-blue-500 w-56',
              filters.search && 'border-blue-400',
            )}
          />
        </div>

        {/* Date range */}
        <select
          value={filters.dateRange}
          onChange={e => set({ dateRange: e.target.value as DateRange })}
          className="text-sm rounded-lg border border-slate-300 px-2.5 py-1.5
                     bg-white text-slate-700 focus:outline-none focus:ring-2 focus:ring-blue-500"
        >
          {DATE_OPTIONS.map(o => (
            <option key={o.value} value={o.value}>{o.label}</option>
          ))}
        </select>

        <div className="h-4 border-l border-slate-200" />

        {/* User filter */}
        <Select
          label="All Users"
          value={filters.userFilter}
          options={userOptions}
          onChange={v => set({ userFilter: v })}
        />

        {/* App filter */}
        <Select
          label="All Apps"
          value={filters.appFilter}
          options={appOptions}
          onChange={v => set({ appFilter: v })}
        />

        {/* OS filter */}
        <Select
          label="All OS"
          value={filters.osFilter}
          options={osOptions}
          onChange={v => set({ osFilter: v })}
        />

        {/* Policy status */}
        <Select
          label="All Statuses"
          value={filters.policyStatusFilter}
          options={POLICY_OPTIONS}
          onChange={v => set({ policyStatusFilter: v })}
        />

        {/* Sign-in result */}
        <Select
          label="All Results"
          value={filters.signInStatusFilter}
          options={SIGNIN_OPTIONS}
          onChange={v => set({ signInStatusFilter: v })}
        />

        {/* Clear filters */}
        {hasActiveFilters && (
          <button
            onClick={() =>
              set({
                search: '',
                userFilter: '',
                appFilter: '',
                osFilter: '',
                policyStatusFilter: '',
                signInStatusFilter: '',
              })
            }
            className="flex items-center gap-1 text-xs text-slate-500 hover:text-red-600
                       transition-colors px-2 py-1.5 rounded-lg hover:bg-red-50"
          >
            <X className="w-3.5 h-3.5" />
            Clear filters
          </button>
        )}
      </div>

      {/* Result count */}
      <p className="text-xs text-slate-400">
        Showing <strong className="text-slate-600">{totalShown.toLocaleString()}</strong> of{' '}
        <strong className="text-slate-600">{totalLoaded.toLocaleString()}</strong> loaded sign-ins
      </p>
    </div>
  );
}
