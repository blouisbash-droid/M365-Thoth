'use client';

import { signIn } from 'next-auth/react';
import { Shield, Lock, BarChart2, Filter, Download, AlertCircle } from 'lucide-react';

interface Props {
  tokenError?: string;
}

const features = [
  { icon: BarChart2, label: 'Compliance overview charts' },
  { icon: Filter,    label: 'Filter by user, app, OS, and policy status' },
  { icon: Lock,      label: 'Device trust type breakdown' },
  { icon: Download,  label: 'Export filtered data to CSV' },
];

export default function LoginPage({ tokenError }: Props) {
  return (
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
          <p className="text-blue-100 mb-8">
            Pull live sign-in logs from Microsoft Graph and see exactly which users,
            apps, and devices would be affected when you require Entra joined,
            Hybrid joined, or Intune enrolled devices.
          </p>
          <ul className="space-y-3">
            {features.map(({ icon: Icon, label }) => (
              <li key={label} className="flex items-center gap-3 text-blue-100">
                <span className="w-7 h-7 rounded-lg bg-white/15 flex items-center justify-center shrink-0">
                  <Icon className="w-3.5 h-3.5" />
                </span>
                {label}
              </li>
            ))}
          </ul>
        </div>

        <p className="text-blue-200 text-xs">
          Requires <strong className="text-white">AuditLog.Read.All</strong> delegated permission
          and one of: Global Admin, Security Admin, Security Reader, Global Reader, or Reports Reader.
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
            Use your Microsoft 365 admin account to access the dashboard.
          </p>

          {/* Token expiry warning */}
          {tokenError === 'RefreshAccessTokenError' && (
            <div className="mb-4 flex items-start gap-2 bg-amber-50 border border-amber-200
                            rounded-lg p-3 text-amber-800 text-sm">
              <AlertCircle className="w-4 h-4 mt-0.5 shrink-0" />
              Session expired. Please sign in again.
            </div>
          )}

          <button
            onClick={() => signIn('azure-ad', { callbackUrl: '/dashboard' })}
            className="w-full flex items-center justify-center gap-3 bg-white border
                       border-slate-300 rounded-lg px-4 py-3 text-sm font-medium
                       text-slate-700 hover:bg-slate-50 active:bg-slate-100
                       transition-colors shadow-sm"
          >
            {/* Microsoft logo SVG */}
            <svg viewBox="0 0 21 21" className="w-5 h-5" aria-hidden>
              <rect x="1"  y="1"  width="9" height="9" fill="#f25022" />
              <rect x="11" y="1"  width="9" height="9" fill="#7fba00" />
              <rect x="1"  y="11" width="9" height="9" fill="#00a4ef" />
              <rect x="11" y="11" width="9" height="9" fill="#ffb900" />
            </svg>
            Sign in with Microsoft
          </button>

          <p className="mt-6 text-center text-xs text-slate-400">
            Sign-in data is read-only and never stored server-side.
          </p>
        </div>
      </div>
    </div>
  );
}
