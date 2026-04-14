'use client';

import { useState } from 'react';
import { Shield, Info } from 'lucide-react';
import { AppConfig, saveConfig } from '@/lib/msalConfig';

interface Props {
  onSaved: (config: AppConfig) => void;
  existing?: AppConfig | null;
}

const GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

export default function ConfigForm({ onSaved, existing }: Props) {
  const [tenantId, setTenantId] = useState(existing?.tenantId ?? '');
  const [clientId, setClientId] = useState(existing?.clientId ?? '');
  const [error, setError]       = useState('');

  function handleSubmit(e: React.FormEvent) {
    e.preventDefault();
    setError('');
    const t = tenantId.trim();
    const c = clientId.trim();

    if (!GUID_RE.test(t)) {
      setError('Tenant ID must be a valid GUID (xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx)');
      return;
    }
    if (!GUID_RE.test(c)) {
      setError('Client ID must be a valid GUID');
      return;
    }

    const config: AppConfig = { tenantId: t, clientId: c };
    saveConfig(config);
    onSaved(config);
  }

  const origin = typeof window !== 'undefined' ? window.location.origin : 'https://your-app-url';

  return (
    <div className="min-h-screen flex items-center justify-center bg-slate-50 p-4">
      <div className="w-full max-w-md">
        {/* Logo */}
        <div className="flex items-center gap-3 mb-8">
          <div className="w-10 h-10 bg-blue-600 rounded-xl flex items-center justify-center">
            <Shield className="w-5 h-5 text-white" />
          </div>
          <div>
            <h1 className="text-lg font-semibold text-slate-900">M365 CA Dashboard</h1>
            <p className="text-xs text-slate-400">Conditional Access Readiness</p>
          </div>
        </div>

        <div className="card p-6">
          <h2 className="text-base font-semibold text-slate-800 mb-1">Connect your tenant</h2>
          <p className="text-sm text-slate-500 mb-6">
            Enter your Azure AD app registration details. These are saved locally in your
            browser — nothing is sent to any server.
          </p>

          <form onSubmit={handleSubmit} className="space-y-4">
            <div>
              <label className="block text-sm font-medium text-slate-700 mb-1.5">
                Tenant ID
              </label>
              <input
                type="text"
                value={tenantId}
                onChange={e => { setTenantId(e.target.value); setError(''); }}
                placeholder="xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
                className="w-full px-3 py-2 border border-slate-300 rounded-lg text-sm
                           focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent
                           font-mono placeholder:font-sans"
                required
              />
            </div>

            <div>
              <label className="block text-sm font-medium text-slate-700 mb-1.5">
                Client ID (App Registration)
              </label>
              <input
                type="text"
                value={clientId}
                onChange={e => { setClientId(e.target.value); setError(''); }}
                placeholder="xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
                className="w-full px-3 py-2 border border-slate-300 rounded-lg text-sm
                           focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent
                           font-mono placeholder:font-sans"
                required
              />
            </div>

            {error && (
              <p className="text-sm text-red-600 bg-red-50 border border-red-200 rounded-lg p-2.5">
                {error}
              </p>
            )}

            <button type="submit" className="btn btn-primary w-full">
              Save &amp; Continue
            </button>
          </form>

          {/* Setup instructions */}
          <div className="mt-5 p-3.5 rounded-lg bg-blue-50 border border-blue-100">
            <div className="flex items-start gap-2">
              <Info className="w-4 h-4 text-blue-500 mt-0.5 shrink-0" />
              <div className="text-xs text-blue-700 space-y-1.5">
                <p className="font-semibold">App Registration requirements:</p>
                <ul className="list-disc ml-3 space-y-1">
                  <li>Platform type: <strong>Single-page application (SPA)</strong></li>
                  <li>
                    Redirect URI: <strong className="font-mono break-all">{origin}</strong>
                  </li>
                  <li>Delegated permission: <strong>AuditLog.Read.All</strong></li>
                  <li>Grant admin consent for the permission</li>
                </ul>
                <p className="mt-1 text-blue-600">
                  The signed-in user also needs one of: Global Admin, Security Admin,
                  Security Reader, Global Reader, or Reports Reader.
                </p>
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}
