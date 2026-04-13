'use client';

import { useState, useEffect, useCallback } from 'react';
import { TenantConfig } from '@/lib/types';
import {
  Building2, Plus, Trash2, CheckCircle2, XCircle,
  Copy, Check, ExternalLink, AlertTriangle, Loader2,
  ChevronDown, ChevronUp, Info,
} from 'lucide-react';
import { cn } from '@/lib/utils';

// ─── Types ────────────────────────────────────────────────────────────────────

interface TenantStatus {
  state: 'untested' | 'testing' | 'ok' | 'error';
  message?: string;
  displayName?: string;
}

// ─── Small helpers ────────────────────────────────────────────────────────────

function CopyButton({ text, label = 'Copy' }: { text: string; label?: string }) {
  const [copied, setCopied] = useState(false);

  const copy = () => {
    navigator.clipboard.writeText(text);
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };

  return (
    <button
      onClick={copy}
      className="btn btn-secondary text-xs"
      title={copied ? 'Copied!' : label}
    >
      {copied ? <Check className="w-3.5 h-3.5 text-green-500" /> : <Copy className="w-3.5 h-3.5" />}
      {copied ? 'Copied' : label}
    </button>
  );
}

// ─── Add Tenant modal ─────────────────────────────────────────────────────────

interface AddModalProps {
  appUrl: string;
  clientId: string;
  onAdded: (tenant: TenantConfig, needsRedeploy: boolean) => void;
  onClose: () => void;
}

function AddTenantModal({ appUrl, clientId, onAdded, onClose }: AddModalProps) {
  const [name, setName]           = useState('');
  const [tenantId, setTenantId]   = useState('');
  const [notes, setNotes]         = useState('');
  const [saving, setSaving]       = useState(false);
  const [error, setError]         = useState('');
  const [step, setStep]           = useState<'form' | 'consent'>(tenantId ? 'consent' : 'form');
  const [addedTenant, setAdded]   = useState<TenantConfig | null>(null);

  const consentUrl = tenantId
    ? `https://login.microsoftonline.com/${tenantId}/adminconsent?client_id=${clientId}&redirect_uri=${appUrl}/tenants`
    : '';

  const save = async () => {
    if (!name.trim() || !tenantId.trim()) return;
    setSaving(true);
    setError('');
    try {
      const resp = await fetch('/api/tenants', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ name, tenantId, notes }),
      });
      const data = await resp.json();
      if (!resp.ok) throw new Error(data.error ?? 'Failed to add tenant');
      setAdded(data.tenant);
      setStep('consent');
      onAdded(data.tenant, data.needsRedeploy);
    } catch (e: unknown) {
      setError((e as Error).message);
    } finally {
      setSaving(false);
    }
  };

  return (
    <div className="fixed inset-0 bg-black/40 backdrop-blur-sm z-50 flex items-center justify-center p-4">
      <div className="bg-white rounded-2xl shadow-xl w-full max-w-lg">
        {/* Header */}
        <div className="flex items-center justify-between p-6 border-b border-slate-100">
          <h2 className="font-semibold text-slate-900">
            {step === 'form' ? 'Add Customer Tenant' : 'Grant Admin Consent'}
          </h2>
          <button onClick={onClose} className="text-slate-400 hover:text-slate-600 text-xl leading-none">&times;</button>
        </div>

        <div className="p-6 space-y-5">
          {step === 'form' ? (
            <>
              {/* Info banner */}
              <div className="bg-blue-50 border border-blue-100 rounded-lg p-4 text-sm text-blue-700 flex gap-3">
                <Info className="w-4 h-4 mt-0.5 shrink-0 text-blue-400" />
                <div>
                  Enter the customer&apos;s Azure AD Tenant ID — you&apos;ll get a consent URL to
                  share with their Global Admin. No customer secrets are stored.
                </div>
              </div>

              {/* Form */}
              <div className="space-y-4">
                <div>
                  <label className="block text-sm font-medium text-slate-700 mb-1">
                    Display Name <span className="text-red-500">*</span>
                  </label>
                  <input
                    type="text"
                    value={name}
                    onChange={e => setName(e.target.value)}
                    placeholder="e.g. Contoso Corp"
                    className="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm
                               focus:outline-none focus:ring-2 focus:ring-blue-500"
                  />
                </div>

                <div>
                  <label className="block text-sm font-medium text-slate-700 mb-1">
                    Azure AD Tenant ID <span className="text-red-500">*</span>
                  </label>
                  <input
                    type="text"
                    value={tenantId}
                    onChange={e => setTenantId(e.target.value.trim())}
                    placeholder="xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
                    className="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm font-mono
                               focus:outline-none focus:ring-2 focus:ring-blue-500"
                  />
                  <p className="text-xs text-slate-400 mt-1">
                    Found in Azure Portal → Entra ID → Overview → Tenant ID
                  </p>
                </div>

                <div>
                  <label className="block text-sm font-medium text-slate-700 mb-1">Notes</label>
                  <textarea
                    value={notes}
                    onChange={e => setNotes(e.target.value)}
                    rows={2}
                    placeholder="Contact name, contract reference, etc. (optional)"
                    className="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm
                               focus:outline-none focus:ring-2 focus:ring-blue-500 resize-none"
                  />
                </div>
              </div>

              {error && <p className="text-sm text-red-600">{error}</p>}

              <div className="flex justify-end gap-2 pt-1">
                <button onClick={onClose} className="btn btn-secondary">Cancel</button>
                <button
                  onClick={save}
                  disabled={!name.trim() || !tenantId.trim() || saving}
                  className="btn btn-primary disabled:opacity-50"
                >
                  {saving && <Loader2 className="w-4 h-4 animate-spin" />}
                  Add Tenant & Get Consent URL
                </button>
              </div>
            </>
          ) : (
            <>
              {/* Step 2 — consent URL */}
              <div className="bg-green-50 border border-green-200 rounded-lg p-4 text-sm text-green-800 flex gap-3">
                <CheckCircle2 className="w-4 h-4 mt-0.5 shrink-0 text-green-500" />
                <div>
                  <strong>{addedTenant?.name ?? name}</strong> has been added.
                  Now share the consent URL below with the customer&apos;s <strong>Global Administrator</strong>.
                </div>
              </div>

              <div>
                <p className="text-sm font-medium text-slate-700 mb-2">Admin Consent URL</p>
                <div className="bg-slate-50 border border-slate-200 rounded-lg p-3 font-mono text-xs
                                text-slate-600 break-all leading-relaxed">
                  {consentUrl}
                </div>
                <div className="flex gap-2 mt-2">
                  <CopyButton text={consentUrl} label="Copy URL" />
                  <a
                    href={consentUrl}
                    target="_blank"
                    rel="noopener noreferrer"
                    className="btn btn-secondary text-xs"
                  >
                    <ExternalLink className="w-3.5 h-3.5" />
                    Open (if you&apos;re the admin)
                  </a>
                </div>
              </div>

              <div className="bg-slate-50 border border-slate-100 rounded-lg p-4 text-sm space-y-2">
                <p className="font-medium text-slate-700">What the customer admin needs to do:</p>
                <ol className="list-decimal list-inside text-slate-600 space-y-1.5 text-xs">
                  <li>Click the consent URL (or paste it in their browser)</li>
                  <li>Sign in with their <strong>Global Administrator</strong> account</li>
                  <li>Review the requested permissions and click <strong>Accept</strong></li>
                  <li>They&apos;ll be redirected back to this page — done!</li>
                </ol>
                <p className="text-xs text-slate-400 pt-1">
                  Permissions requested: <code>AuditLog.Read.All</code>, <code>Directory.Read.All</code>
                </p>
              </div>

              <div className="flex justify-end gap-2 pt-1">
                <button onClick={onClose} className="btn btn-primary">Done</button>
              </div>
            </>
          )}
        </div>
      </div>
    </div>
  );
}

// ─── Tenant card ──────────────────────────────────────────────────────────────

interface TenantCardProps {
  tenant: TenantConfig;
  appUrl: string;
  clientId: string;
  storageMode: 'redis' | 'env';
  onDelete: (id: string) => void;
}

function TenantCard({ tenant, appUrl, clientId, storageMode, onDelete }: TenantCardProps) {
  const [status, setStatus]       = useState<TenantStatus>({ state: 'untested' });
  const [expanded, setExpanded]   = useState(false);
  const [deleting, setDeleting]   = useState(false);

  const consentUrl = `https://login.microsoftonline.com/${tenant.tenantId}/adminconsent?client_id=${clientId}&redirect_uri=${appUrl}/tenants`;

  const test = useCallback(async () => {
    setStatus({ state: 'testing' });
    try {
      const resp = await fetch('/api/tenants/test', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ tenantId: tenant.tenantId }),
      });
      const data = await resp.json();
      if (data.ok) {
        setStatus({ state: 'ok', displayName: data.displayName, message: data.verifiedDomain });
      } else {
        setStatus({ state: 'error', message: data.error });
      }
    } catch {
      setStatus({ state: 'error', message: 'Network error' });
    }
  }, [tenant.tenantId]);

  const remove = async () => {
    if (!confirm(`Remove "${tenant.name}" from the dashboard? This won't affect the customer's tenant.`)) return;
    setDeleting(true);
    try {
      const resp = await fetch('/api/tenants', {
        method: 'DELETE',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ id: tenant.id }),
      });
      if (resp.ok) {
        onDelete(tenant.id);
      } else {
        const data = await resp.json();
        alert(data.error ?? 'Failed to remove tenant');
      }
    } finally {
      setDeleting(false);
    }
  };

  const statusBadge = () => {
    switch (status.state) {
      case 'testing':  return <span className="badge badge-blue"><Loader2 className="w-3 h-3 animate-spin" /> Testing…</span>;
      case 'ok':       return <span className="badge badge-green"><CheckCircle2 className="w-3 h-3" /> Connected</span>;
      case 'error':    return <span className="badge badge-red"><XCircle className="w-3 h-3" /> Error</span>;
      default:         return <span className="badge badge-gray">Not tested</span>;
    }
  };

  return (
    <div className="card overflow-hidden">
      <div className="p-4 flex items-center gap-4">
        {/* Icon */}
        <div className="w-9 h-9 rounded-lg bg-slate-100 flex items-center justify-center shrink-0">
          <Building2 className="w-4 h-4 text-slate-500" />
        </div>

        {/* Name + tenant ID */}
        <div className="min-w-0 flex-1">
          <p className="font-medium text-slate-900 text-sm truncate">{tenant.name}</p>
          <p className="text-xs text-slate-400 font-mono truncate">{tenant.tenantId}</p>
        </div>

        {/* Status */}
        <div className="shrink-0">{statusBadge()}</div>

        {/* Actions */}
        <div className="flex items-center gap-1.5 shrink-0">
          <button
            onClick={test}
            disabled={status.state === 'testing'}
            className="btn btn-secondary text-xs disabled:opacity-50"
          >
            Test
          </button>
          {storageMode === 'redis' && (
            <button
              onClick={remove}
              disabled={deleting}
              className="p-1.5 rounded-lg text-slate-400 hover:text-red-500 hover:bg-red-50 transition-colors disabled:opacity-50"
              title="Remove tenant"
            >
              {deleting ? <Loader2 className="w-4 h-4 animate-spin" /> : <Trash2 className="w-4 h-4" />}
            </button>
          )}
          <button
            onClick={() => setExpanded(e => !e)}
            className="p-1.5 rounded-lg text-slate-400 hover:bg-slate-100 transition-colors"
          >
            {expanded ? <ChevronUp className="w-4 h-4" /> : <ChevronDown className="w-4 h-4" />}
          </button>
        </div>
      </div>

      {/* Expanded details */}
      {expanded && (
        <div className="border-t border-slate-100 p-4 bg-slate-50 space-y-3 text-sm">
          {status.state === 'ok' && (
            <p className="text-green-700">
              <strong>Verified:</strong> {status.displayName}
              {status.message ? ` · ${status.message}` : ''}
            </p>
          )}
          {status.state === 'error' && (
            <p className="text-red-600">
              <strong>Error:</strong> {status.message}
            </p>
          )}
          {tenant.notes && (
            <p className="text-slate-600"><strong>Notes:</strong> {tenant.notes}</p>
          )}
          <p className="text-slate-400 text-xs">
            Added {new Date(tenant.createdAt).toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' })}
          </p>

          {/* Consent URL */}
          <div>
            <p className="text-xs font-medium text-slate-500 mb-1.5">Admin Consent URL</p>
            <div className="bg-white border border-slate-200 rounded-lg p-2.5 font-mono text-xs text-slate-500 break-all">
              {consentUrl}
            </div>
            <div className="flex gap-2 mt-1.5">
              <CopyButton text={consentUrl} label="Copy Consent URL" />
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

// ─── Main TenantManager ───────────────────────────────────────────────────────

interface Props {
  appUrl: string;
  clientId: string;
}

export default function TenantManager({ appUrl, clientId }: Props) {
  const [tenants, setTenants]       = useState<TenantConfig[]>([]);
  const [storageMode, setMode]      = useState<'redis' | 'env'>('env');
  const [loading, setLoading]       = useState(true);
  const [showAdd, setShowAdd]       = useState(false);
  const [envJson, setEnvJson]       = useState<string | null>(null);

  // Fetch tenants on mount
  useEffect(() => {
    fetch('/api/tenants')
      .then(r => r.json())
      .then(d => {
        setTenants(d.tenants ?? []);
        setMode(d.storageMode ?? 'env');
      })
      .finally(() => setLoading(false));
  }, []);

  const handleAdded = (tenant: TenantConfig, needsRedeploy: boolean) => {
    setTenants(prev => {
      // avoid duplicates
      if (prev.find(t => t.id === tenant.id)) return prev;
      const next = [tenant, ...prev];
      if (needsRedeploy) {
        setEnvJson(JSON.stringify(next, null, 2));
      }
      return next;
    });
    setShowAdd(false);
  };

  const handleDelete = (id: string) => {
    setTenants(prev => prev.filter(t => t.id !== id));
  };

  return (
    <div className="max-w-screen-2xl mx-auto px-4 py-6 space-y-6">

      {/* Header */}
      <div className="flex items-center justify-between flex-wrap gap-4">
        <div>
          <h1 className="text-xl font-bold text-slate-900">Customer Tenants</h1>
          <p className="text-sm text-slate-500 mt-0.5">
            Connect customer Entra tenants — one consent click per customer, no secrets stored.
          </p>
        </div>
        <button onClick={() => setShowAdd(true)} className="btn btn-primary">
          <Plus className="w-4 h-4" />
          Add Tenant
        </button>
      </div>

      {/* Storage mode banner */}
      {storageMode === 'env' && (
        <div className="card p-4 border-amber-200 bg-amber-50 flex items-start gap-3">
          <AlertTriangle className="w-5 h-5 text-amber-500 mt-0.5 shrink-0" />
          <div className="text-sm text-amber-800">
            <strong>Dynamic storage not configured.</strong>{' '}
            Tenants added here persist only until the next deployment.
            For permanent dynamic management, add{' '}
            <strong>Upstash Redis</strong> to your Vercel project
            (Project → Storage → Upstash Redis) — the env vars are added automatically.
            <br />
            <span className="text-amber-600 text-xs mt-1 block">
              Alternatively, manage tenants via the <code>TENANTS_CONFIG</code> environment variable (JSON array).
            </span>
          </div>
        </div>
      )}

      {/* TENANTS_CONFIG JSON (shown after adding when Redis not available) */}
      {envJson && (
        <div className="card p-5 border-blue-200 bg-blue-50">
          <div className="flex items-center justify-between mb-3">
            <div>
              <p className="font-semibold text-blue-900 text-sm">Update TENANTS_CONFIG</p>
              <p className="text-xs text-blue-600 mt-0.5">
                Paste this into your <strong>TENANTS_CONFIG</strong> environment variable in the Vercel dashboard, then redeploy.
              </p>
            </div>
            <div className="flex gap-2">
              <CopyButton text={envJson} label="Copy JSON" />
              <button onClick={() => setEnvJson(null)} className="btn btn-secondary text-xs">Dismiss</button>
            </div>
          </div>
          <pre className="text-xs bg-white border border-blue-200 rounded-lg p-3 overflow-x-auto text-slate-700 max-h-56">
            {envJson}
          </pre>
        </div>
      )}

      {/* Tenant list */}
      {loading ? (
        <div className="flex justify-center py-16">
          <Loader2 className="w-6 h-6 animate-spin text-slate-400" />
        </div>
      ) : tenants.length === 0 ? (
        <div className="card p-12 flex flex-col items-center gap-4 text-center">
          <Building2 className="w-10 h-10 text-slate-300" />
          <div>
            <p className="font-semibold text-slate-700">No customer tenants connected yet</p>
            <p className="text-sm text-slate-400 mt-1">
              Click <strong>Add Tenant</strong> to connect your first customer.
            </p>
          </div>
          <button onClick={() => setShowAdd(true)} className="btn btn-primary mt-2">
            <Plus className="w-4 h-4" />
            Add First Tenant
          </button>
        </div>
      ) : (
        <div className="space-y-3">
          {tenants.map(t => (
            <TenantCard
              key={t.id}
              tenant={t}
              appUrl={appUrl}
              clientId={clientId}
              storageMode={storageMode}
              onDelete={handleDelete}
            />
          ))}
        </div>
      )}

      {/* How it works */}
      <HowItWorks />

      {/* Add modal */}
      {showAdd && (
        <AddTenantModal
          appUrl={appUrl}
          clientId={clientId}
          onAdded={handleAdded}
          onClose={() => setShowAdd(false)}
        />
      )}
    </div>
  );
}

// ─── How it works accordion ───────────────────────────────────────────────────

function HowItWorks() {
  const [open, setOpen] = useState(false);

  return (
    <div className="card overflow-hidden">
      <button
        onClick={() => setOpen(o => !o)}
        className="w-full flex items-center justify-between p-4 text-left hover:bg-slate-50 transition-colors"
      >
        <span className="text-sm font-semibold text-slate-700">How MSP multi-tenant access works</span>
        {open ? <ChevronUp className="w-4 h-4 text-slate-400" /> : <ChevronDown className="w-4 h-4 text-slate-400" />}
      </button>

      {open && (
        <div className="border-t border-slate-100 p-5 text-sm text-slate-600 space-y-4">
          <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
            {[
              {
                step: '1',
                title: 'One app registration',
                body: 'Your Azure AD app (registered in the MSP tenant) is configured as multi-tenant with AuditLog.Read.All and Directory.Read.All application permissions.',
              },
              {
                step: '2',
                title: 'Customer admin grants consent',
                body: 'The customer\'s Global Administrator visits the consent URL and approves the requested permissions. This registers your app\'s service principal in their tenant.',
              },
              {
                step: '3',
                title: 'App-only token per tenant',
                body: 'The dashboard uses your client ID and secret to request an app-only token scoped to the customer\'s tenant ID. No customer credentials are ever stored.',
              },
            ].map(({ step, title, body }) => (
              <div key={step} className="flex gap-3">
                <span className="w-6 h-6 rounded-full bg-blue-100 text-blue-700 text-xs font-bold flex items-center justify-center shrink-0 mt-0.5">
                  {step}
                </span>
                <div>
                  <p className="font-semibold text-slate-800">{title}</p>
                  <p className="text-slate-500 mt-0.5 text-xs leading-relaxed">{body}</p>
                </div>
              </div>
            ))}
          </div>

          <div className="bg-slate-50 rounded-lg p-4 text-xs space-y-1.5 text-slate-500">
            <p><strong className="text-slate-700">App registration requirements:</strong></p>
            <ul className="list-disc list-inside space-y-1">
              <li>Supported account types: <strong>Accounts in any organizational directory (Multi-tenant)</strong></li>
              <li>Application permissions granted: <code>AuditLog.Read.All</code>, <code>Directory.Read.All</code></li>
              <li>Redirect URI added: <code>https://m365-ca-review.vercel.app/tenants</code></li>
            </ul>
          </div>
        </div>
      )}
    </div>
  );
}
