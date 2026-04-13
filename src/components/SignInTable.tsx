'use client';

import {
  useReactTable,
  getCoreRowModel,
  getSortedRowModel,
  getPaginationRowModel,
  flexRender,
  createColumnHelper,
  SortingState,
} from '@tanstack/react-table';
import { useState, useCallback } from 'react';
import { SignInLog } from '@/lib/types';
import {
  cn, formatDate, formatTrustType, policyStatusLabel,
  deviceCategoryLabel, getOSLabel, exportToCSV,
} from '@/lib/utils';
import {
  ChevronUp, ChevronDown, ChevronsUpDown,
  ChevronLeft, ChevronRight, Download, Loader2,
} from 'lucide-react';

const col = createColumnHelper<SignInLog>();

// ─── Badge helpers ────────────────────────────────────────────────────────────

function PolicyBadge({ status }: { status: SignInLog['policyStatus'] }) {
  return (
    <span className={cn(
      'badge text-xs',
      status === 'passes'  && 'badge-green',
      status === 'fails'   && 'badge-red',
      status === 'unknown' && 'badge-yellow',
    )}>
      {policyStatusLabel(status)}
    </span>
  );
}

function SignInStatusBadge({ errorCode }: { errorCode: number }) {
  return errorCode === 0
    ? <span className="badge badge-green">Success</span>
    : <span className="badge badge-red">Failure</span>;
}

function TrustTypeBadge({ trustType }: { trustType: string | null }) {
  if (!trustType) return <span className="badge badge-gray">None</span>;
  const map: Record<string, string> = {
    AzureAD:   'badge-blue',
    ServerAD:  'badge-violet',
    Workplace: 'badge-cyan',
  };
  return <span className={cn('badge', map[trustType] ?? 'badge-gray')}>{formatTrustType(trustType)}</span>;
}

// ─── Column definitions ───────────────────────────────────────────────────────

const columns = [
  col.accessor('createdDateTime', {
    header: 'Date / Time',
    cell: info => (
      <span className="text-xs text-slate-500">{formatDate(info.getValue())}</span>
    ),
    sortingFn: 'datetime',
  }),
  col.accessor('userDisplayName', {
    header: 'User',
    cell: info => {
      const row = info.row.original;
      return (
        <div className="min-w-0">
          <p className="font-medium text-slate-900 truncate max-w-[160px]" title={row.userDisplayName}>
            {row.userDisplayName || '—'}
          </p>
          <p className="text-xs text-slate-400 truncate max-w-[160px]" title={row.userPrincipalName}>
            {row.userPrincipalName}
          </p>
        </div>
      );
    },
  }),
  col.accessor('appDisplayName', {
    header: 'Application',
    cell: info => (
      <span className="truncate max-w-[140px] block" title={info.getValue()}>
        {info.getValue() || '—'}
      </span>
    ),
  }),
  col.accessor(row => row.deviceDetail.displayName, {
    id: 'deviceName',
    header: 'Device',
    cell: info => (
      <span className="truncate max-w-[130px] block text-slate-600" title={info.getValue() ?? ''}>
        {info.getValue() || <span className="text-slate-300 italic">Unknown</span>}
      </span>
    ),
  }),
  col.accessor(row => getOSLabel(row.deviceDetail.operatingSystem), {
    id: 'os',
    header: 'OS',
    cell: info => <span className="text-slate-600">{info.getValue()}</span>,
  }),
  col.accessor(row => row.deviceDetail.trustType, {
    id: 'trustType',
    header: 'Trust Type',
    cell: info => <TrustTypeBadge trustType={info.getValue()} />,
  }),
  col.accessor(row => row.deviceDetail.isManaged, {
    id: 'isManaged',
    header: 'Intune',
    cell: info => {
      const v = info.getValue();
      if (v === null) return <span className="text-slate-300 text-xs">—</span>;
      return v
        ? <span className="badge badge-green">Yes</span>
        : <span className="badge badge-gray">No</span>;
    },
  }),
  col.accessor('policyStatus', {
    header: 'Policy',
    cell: info => <PolicyBadge status={info.getValue()} />,
  }),
  col.accessor(row => row.status.errorCode, {
    id: 'signInStatus',
    header: 'Result',
    cell: info => <SignInStatusBadge errorCode={info.getValue()} />,
    sortingFn: 'basic',
  }),
  col.accessor('conditionalAccessStatus', {
    header: 'CA Status',
    cell: info => {
      const v = info.getValue();
      if (!v || v === 'notApplied') return <span className="text-slate-300 text-xs">—</span>;
      return (
        <span className={cn('badge',
          v === 'success' ? 'badge-green' :
          v === 'failure' ? 'badge-red' : 'badge-gray',
        )}>
          {v}
        </span>
      );
    },
  }),
  col.accessor('ipAddress', {
    header: 'IP',
    cell: info => <span className="text-xs text-slate-400 font-mono">{info.getValue() || '—'}</span>,
  }),
  col.accessor(row => [row.location.city, row.location.countryOrRegion].filter(Boolean).join(', '), {
    id: 'location',
    header: 'Location',
    cell: info => <span className="text-xs text-slate-500">{info.getValue() || '—'}</span>,
  }),
];

// ─── Sort icon ────────────────────────────────────────────────────────────────

function SortIcon({ sorted }: { sorted: false | 'asc' | 'desc' }) {
  if (!sorted) return <ChevronsUpDown className="w-3 h-3 text-slate-300" />;
  return sorted === 'asc'
    ? <ChevronUp className="w-3 h-3 text-blue-500" />
    : <ChevronDown className="w-3 h-3 text-blue-500" />;
}

// ─── Main table component ─────────────────────────────────────────────────────

interface Props {
  signIns: SignInLog[];
  isLoadingMore: boolean;
  hasMore: boolean;
  onLoadMore: () => void;
}

export default function SignInTable({ signIns, isLoadingMore, hasMore, onLoadMore }: Props) {
  const [sorting, setSorting] = useState<SortingState>([
    { id: 'createdDateTime', desc: true },
  ]);
  const [pagination, setPagination] = useState({ pageIndex: 0, pageSize: 50 });

  const table = useReactTable({
    data: signIns,
    columns,
    state: { sorting, pagination },
    onSortingChange: setSorting,
    onPaginationChange: setPagination,
    getCoreRowModel: getCoreRowModel(),
    getSortedRowModel: getSortedRowModel(),
    getPaginationRowModel: getPaginationRowModel(),
  });

  const handleExport = useCallback(() => {
    exportToCSV(signIns, `signin-report-${new Date().toISOString().slice(0, 10)}.csv`);
  }, [signIns]);

  const { pageIndex, pageSize } = table.getState().pagination;
  const pageCount = table.getPageCount();
  const from = pageIndex * pageSize + 1;
  const to = Math.min(from + pageSize - 1, signIns.length);

  return (
    <div className="card">
      {/* Toolbar */}
      <div className="flex items-center justify-between px-4 py-3 border-b border-slate-100">
        <div>
          <h3 className="text-sm font-semibold text-slate-700">Sign-in Logs</h3>
          <p className="text-xs text-slate-400">
            {signIns.length.toLocaleString()} records
            {isLoadingMore && (
              <span className="ml-2 inline-flex items-center gap-1 text-blue-500">
                <Loader2 className="w-3 h-3 animate-spin" /> loading more…
              </span>
            )}
          </p>
        </div>
        <div className="flex items-center gap-2">
          {hasMore && !isLoadingMore && (
            <button onClick={onLoadMore} className="btn btn-secondary text-xs">
              Load more
            </button>
          )}
          <button onClick={handleExport} className="btn btn-secondary text-xs">
            <Download className="w-3.5 h-3.5" />
            Export CSV
          </button>
        </div>
      </div>

      {/* Table */}
      <div className="overflow-x-auto">
        <table className="w-full data-table">
          <thead className="bg-slate-50 border-b border-slate-100">
            {table.getHeaderGroups().map(hg => (
              <tr key={hg.id}>
                {hg.headers.map(header => (
                  <th
                    key={header.id}
                    className={cn(
                      header.column.getCanSort() && 'cursor-pointer hover:bg-slate-100',
                    )}
                    onClick={header.column.getToggleSortingHandler()}
                  >
                    <div className="flex items-center gap-1">
                      {flexRender(header.column.columnDef.header, header.getContext())}
                      {header.column.getCanSort() && (
                        <SortIcon sorted={header.column.getIsSorted()} />
                      )}
                    </div>
                  </th>
                ))}
              </tr>
            ))}
          </thead>
          <tbody>
            {table.getRowModel().rows.length === 0 ? (
              <tr>
                <td colSpan={columns.length} className="py-12 text-center text-slate-400 text-sm">
                  No sign-ins match the current filters.
                </td>
              </tr>
            ) : (
              table.getRowModel().rows.map(row => (
                <tr key={row.id}>
                  {row.getVisibleCells().map(cell => (
                    <td key={cell.id}>
                      {flexRender(cell.column.columnDef.cell, cell.getContext())}
                    </td>
                  ))}
                </tr>
              ))
            )}
          </tbody>
        </table>
      </div>

      {/* Pagination */}
      <div className="flex items-center justify-between px-4 py-3 border-t border-slate-100 bg-slate-50">
        <p className="text-xs text-slate-500">
          {signIns.length > 0
            ? `Showing ${from.toLocaleString()}–${to.toLocaleString()} of ${signIns.length.toLocaleString()}`
            : 'No results'}
        </p>
        <div className="flex items-center gap-2">
          {/* Page size */}
          <select
            value={pageSize}
            onChange={e => {
              table.setPageSize(Number(e.target.value));
              table.setPageIndex(0);
            }}
            className="text-xs border border-slate-300 rounded-md px-2 py-1
                       bg-white text-slate-700 focus:outline-none focus:ring-1 focus:ring-blue-500"
          >
            {[25, 50, 100, 250].map(s => (
              <option key={s} value={s}>{s} / page</option>
            ))}
          </select>

          <button
            onClick={() => table.previousPage()}
            disabled={!table.getCanPreviousPage()}
            className="p-1.5 rounded-md border border-slate-300 disabled:opacity-30
                       hover:bg-slate-100 transition-colors"
            aria-label="Previous page"
          >
            <ChevronLeft className="w-3.5 h-3.5" />
          </button>

          <span className="text-xs text-slate-500 min-w-[70px] text-center">
            {pageIndex + 1} / {pageCount || 1}
          </span>

          <button
            onClick={() => table.nextPage()}
            disabled={!table.getCanNextPage()}
            className="p-1.5 rounded-md border border-slate-300 disabled:opacity-30
                       hover:bg-slate-100 transition-colors"
            aria-label="Next page"
          >
            <ChevronRight className="w-3.5 h-3.5" />
          </button>
        </div>
      </div>
    </div>
  );
}
