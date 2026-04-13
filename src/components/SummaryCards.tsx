import { DashboardStats } from '@/lib/types';
import { CheckCircle2, XCircle, HelpCircle, Users2, TrendingUp } from 'lucide-react';
import { cn } from '@/lib/utils';

interface Props {
  stats: DashboardStats;
  isLoadingMore: boolean;
}

function pct(part: number, total: number) {
  if (total === 0) return '—';
  return `${Math.round((part / total) * 100)}%`;
}

function Card({
  label,
  value,
  sub,
  icon: Icon,
  color,
}: {
  label: string;
  value: number | string;
  sub: string;
  icon: React.ElementType;
  color: 'green' | 'red' | 'yellow' | 'blue' | 'slate';
}) {
  const colourMap = {
    green:  { wrap: 'bg-green-50 border-green-100',  icon: 'bg-green-100 text-green-600', text: 'text-green-700' },
    red:    { wrap: 'bg-red-50 border-red-100',      icon: 'bg-red-100 text-red-600',     text: 'text-red-700' },
    yellow: { wrap: 'bg-yellow-50 border-yellow-100',icon: 'bg-yellow-100 text-yellow-600',text: 'text-yellow-700' },
    blue:   { wrap: 'bg-blue-50 border-blue-100',    icon: 'bg-blue-100 text-blue-600',   text: 'text-blue-700' },
    slate:  { wrap: 'bg-white border-slate-200',     icon: 'bg-slate-100 text-slate-600', text: 'text-slate-600' },
  };
  const c = colourMap[color];

  return (
    <div className={cn('card p-5 flex items-center gap-4', c.wrap)}>
      <div className={cn('w-10 h-10 rounded-xl flex items-center justify-center shrink-0', c.icon)}>
        <Icon className="w-5 h-5" />
      </div>
      <div className="min-w-0">
        <p className="text-xs font-medium text-slate-500 uppercase tracking-wide truncate">{label}</p>
        <p className="text-2xl font-bold text-slate-900 leading-tight">
          {typeof value === 'number' ? value.toLocaleString() : value}
        </p>
        <p className={cn('text-xs font-medium', c.text)}>{sub}</p>
      </div>
    </div>
  );
}

export default function SummaryCards({ stats, isLoadingMore }: Props) {
  const { total, passes, fails, unknown } = stats;

  return (
    <div className="space-y-3">
      <div className="grid grid-cols-2 lg:grid-cols-4 gap-3">
        <Card
          label="Total Sign-ins"
          value={total}
          sub={isLoadingMore ? 'Loading more…' : 'in date range'}
          icon={Users2}
          color="slate"
        />
        <Card
          label="Passes Policy"
          value={passes}
          sub={pct(passes, total)}
          icon={CheckCircle2}
          color="green"
        />
        <Card
          label="Fails Policy"
          value={fails}
          sub={pct(fails, total)}
          icon={XCircle}
          color="red"
        />
        <Card
          label="No Device Info"
          value={unknown}
          sub={pct(unknown, total)}
          icon={HelpCircle}
          color="yellow"
        />
      </div>

      {/* Device breakdown sub-row */}
      <div className="card p-4">
        <p className="text-xs font-semibold text-slate-500 uppercase tracking-wide mb-3">
          Device category breakdown
        </p>
        <div className="grid grid-cols-2 sm:grid-cols-4 gap-4">
          {[
            { label: 'Entra Joined',        value: stats.entraJoined,        color: 'bg-blue-500' },
            { label: 'Hybrid Entra Joined', value: stats.hybridEntraJoined,  color: 'bg-violet-500' },
            { label: 'Intune Enrolled',     value: stats.intuneEnrolled,     color: 'bg-cyan-500' },
            { label: 'Registered Only',     value: stats.registeredOnly,     color: 'bg-orange-400' },
          ].map(({ label, value, color }) => (
            <div key={label} className="flex items-center gap-2.5">
              <span className={cn('w-2.5 h-2.5 rounded-full shrink-0', color)} />
              <div>
                <p className="text-sm font-semibold text-slate-900">{value.toLocaleString()}</p>
                <p className="text-xs text-slate-500">{label}</p>
              </div>
            </div>
          ))}
        </div>

        {/* Progress bar */}
        {total > 0 && (
          <div className="mt-4 h-2 rounded-full overflow-hidden bg-slate-100 flex">
            {[
              { value: stats.entraJoined,        color: 'bg-blue-500' },
              { value: stats.hybridEntraJoined,  color: 'bg-violet-500' },
              { value: stats.intuneEnrolled,     color: 'bg-cyan-500' },
              { value: stats.registeredOnly,     color: 'bg-orange-400' },
              { value: stats.noDevice,           color: 'bg-slate-300' },
            ].map(({ value, color }, i) => (
              <div
                key={i}
                className={cn('h-full transition-all', color)}
                style={{ width: `${(value / total) * 100}%` }}
              />
            ))}
          </div>
        )}
      </div>
    </div>
  );
}
