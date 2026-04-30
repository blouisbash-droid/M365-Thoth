'use client';

import { AZURE_CLIENT_ID, AZURE_AUTHORITY } from '@/config';
import { useState, useEffect, useCallback, useMemo } from 'react';
import {
  Shield, Users, UserCheck, UserX, Key, Lock, Globe, Monitor,
  Smartphone, AlertTriangle, CheckCircle2, XCircle, Mail, RefreshCw,
  LogOut, ChevronDown, Building2, Eye, Wifi, WifiOff, Crown,
  ShieldCheck, ShieldAlert, ShieldOff, Fingerprint, Server, Inbox,
  ClipboardList, Settings2, BarChart3, Activity, Info, Search,
  Calendar, MapPin, AlertCircle, Laptop, TabletSmartphone, HardDrive
} from 'lucide-react';
import {
  PieChart, Pie, Cell, Tooltip, Legend, ResponsiveContainer,
  BarChart, Bar, XAxis, YAxis, CartesianGrid, LineChart, Line, Area, AreaChart
} from 'recharts';

// ─── MSAL / Graph Auth ───────────────────────────────────────────────────────
import { PublicClientApplication, InteractionRequiredAuthError } from '@azure/msal-browser';

const MSAL_CONFIG = {
  auth: {
    clientId: AZURE_CLIENT_ID,
    authority: AZURE_AUTHORITY,
    redirectUri: typeof window !== 'undefined' ? window.location.origin + '/dashboard': '/dashboard',
  },
  cache: { cacheLocation: 'localStorage' },
};

const SCOPES = [
  'AuditLog.Read.All',
  'Directory.Read.All',
  'Policy.Read.All',
  'DeviceManagementConfiguration.Read.All',
  'DeviceManagementManagedDevices.Read.All',
  'Reports.Read.All',
  'MailboxSettings.Read',
  'SecurityEvents.Read.All',
  'RoleManagement.Read.All',
];

// ─── Graph API helpers ────────────────────────────────────────────────────────
async function graphGet(token, path) {
  const base = 'https://graph.microsoft.com/v1.0';
  const betaBase = 'https://graph.microsoft.com/beta';
  const url = path.startsWith('http') ? path
    : `${path.startsWith('/beta') ? betaBase : base}${path}`;
  const res = await fetch(url, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!res.ok) {
    const body = await res.json().catch(() => ({}));
    throw new Error(`Graph ${path} → ${res.status}: ${body?.error?.message || JSON.stringify(body)}`);
  }
  return res.json();
}

async function graphGetBeta(token, path) {
  const url = `https://graph.microsoft.com/beta${path}`;
  const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  if (!res.ok) throw new Error(`Graph beta ${path} → ${res.status}`);
  return res.json();
}

async function graphGetAll(token, path) {
  let items = [];
  let url = `https://graph.microsoft.com/v1.0${path}`;
  while (url) {
    const res = await fetch(url, {
      headers: { Authorization: `Bearer ${token}` },
    });
    if (!res.ok) {
      const body = await res.json().catch(() => ({}));
      throw new Error(`Graph ${path} → ${res.status}: ${body?.error?.message || JSON.stringify(body)}`);
    }
    const data = await res.json();
    items = items.concat(data.value || []);
    url = data['@odata.nextLink'] || null;
  }
  return items;
}

// ─── Utility helpers ──────────────────────────────────────────────────────────
function fmtDate(str) {
  if (!str) return 'Never';
  return new Date(str).toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
}
function fmtDateTime(str) {
  if (!str) return 'Never';
  return new Date(str).toLocaleString('en-US', { month: 'short', day: 'numeric', hour: '2-digit', minute: '2-digit' });
}
function daysSince(str) {
  if (!str) return null;
  return Math.floor((Date.now() - new Date(str).getTime()) / 86400000);
}

// ─── Color palette ────────────────────────────────────────────────────────────
const COLORS = {
  blue: '#378ADD', green: '#639922', amber: '#BA7517', red: '#E24B4A',
  purple: '#7F77DD', teal: '#1D9E75', coral: '#D85A30', gray: '#888780',
  pink: '#D4537E', cyan: '#0891b2',
};

// ─── StatusPill ───────────────────────────────────────────────────────────────
function StatusPill({ on, labelOn = 'Enabled', labelOff = 'Disabled', unknown = false }) {
  if (unknown) return (
    <span style={{ display: 'inline-flex', alignItems: 'center', gap: 4, fontSize: 12, fontWeight: 500,
      background: '#F1EFE8', color: '#5F5E5A', padding: '2px 8px', borderRadius: 99, border: '0.5px solid #B4B2A9' }}>
      Unknown
    </span>
  );
  return (
    <span style={{ display: 'inline-flex', alignItems: 'center', gap: 4, fontSize: 12, fontWeight: 500,
      background: on ? '#EAF3DE' : '#FCEBEB',
      color: on ? '#3B6D11' : '#A32D2D',
      padding: '2px 8px', borderRadius: 99,
      border: `0.5px solid ${on ? '#C0DD97' : '#F7C1C1'}` }}>
      <span style={{ width: 6, height: 6, borderRadius: '50%', background: on ? '#639922' : '#E24B4A' }} />
      {on ? labelOn : labelOff}
    </span>
  );
}

// ─── Card ─────────────────────────────────────────────────────────────────────
function Card({ children, style = {} }) {
  return (
    <div style={{ background: 'var(--color-background-primary)', borderRadius: 'var(--border-radius-lg)',
      border: '0.5px solid var(--color-border-tertiary)', padding: '1rem 1.25rem', ...style }}>
      {children}
    </div>
  );
}

// ─── SectionTitle ─────────────────────────────────────────────────────────────
function SectionTitle({ icon: Icon, title, sub }) {
  return (
    <div style={{ display: 'flex', alignItems: 'center', gap: 10, marginBottom: 14 }}>
      <div style={{ width: 32, height: 32, borderRadius: 8, background: '#E6F1FB',
        display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 }}>
        <Icon size={16} color="#185FA5" />
      </div>
      <div>
        <div style={{ fontSize: 14, fontWeight: 500, color: 'var(--color-text-primary)' }}>{title}</div>
        {sub && <div style={{ fontSize: 12, color: 'var(--color-text-secondary)' }}>{sub}</div>}
      </div>
    </div>
  );
}

// ─── MetricCard ───────────────────────────────────────────────────────────────
function MetricCard({ label, value, sub, icon: Icon, color = '#378ADD', bgColor = '#E6F1FB' }) {
  return (
    <div style={{ background: 'var(--color-background-secondary)', borderRadius: 'var(--border-radius-md)',
      padding: '1rem', display: 'flex', alignItems: 'center', gap: 12 }}>
      {Icon && (
        <div style={{ width: 36, height: 36, borderRadius: 8, background: bgColor,
          display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 }}>
          <Icon size={18} color={color} />
        </div>
      )}
      <div style={{ minWidth: 0 }}>
        <div style={{ fontSize: 12, color: 'var(--color-text-secondary)', marginBottom: 2 }}>{label}</div>
        <div style={{ fontSize: 22, fontWeight: 500, color: 'var(--color-text-primary)', lineHeight: 1.1 }}>
          {value ?? '—'}
        </div>
        {sub && <div style={{ fontSize: 11, color: 'var(--color-text-secondary)', marginTop: 2 }}>{sub}</div>}
      </div>
    </div>
  );
}

// ─── SecurityIndicator ────────────────────────────────────────────────────────
function SecurityIndicator({ label, description, status, icon: Icon }) {
  const statusConfig = {
    good: { bg: '#EAF3DE', border: '#C0DD97', iconBg: '#C0DD97', iconColor: '#27500A', dot: '#639922' },
    warning: { bg: '#FAEEDA', border: '#FAC775', iconBg: '#FAC775', iconColor: '#633806', dot: '#BA7517' },
    bad: { bg: '#FCEBEB', border: '#F7C1C1', iconBg: '#F7C1C1', iconColor: '#791F1F', dot: '#E24B4A' },
    unknown: { bg: 'var(--color-background-secondary)', border: 'var(--color-border-tertiary)', iconBg: '#D3D1C7', iconColor: '#5F5E5A', dot: '#888780' },
  };
  const s = statusConfig[status] || statusConfig.unknown;
  return (
    <div style={{ background: s.bg, border: `0.5px solid ${s.border}`, borderRadius: 'var(--border-radius-md)',
      padding: '12px 14px', display: 'flex', alignItems: 'flex-start', gap: 10 }}>
      <div style={{ width: 30, height: 30, borderRadius: 6, background: s.iconBg,
        display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0, marginTop: 1 }}>
        <Icon size={15} color={s.iconColor} />
      </div>
      <div style={{ minWidth: 0, flex: 1 }}>
        <div style={{ fontSize: 13, fontWeight: 500, color: 'var(--color-text-primary)', display: 'flex', alignItems: 'center', gap: 6 }}>
          {label}
          <span style={{ width: 7, height: 7, borderRadius: '50%', background: s.dot, flexShrink: 0 }} />
        </div>
        <div style={{ fontSize: 11, color: 'var(--color-text-secondary)', marginTop: 2, lineHeight: 1.5 }}>{description}</div>
      </div>
    </div>
  );
}

// ─── Loading skeleton ─────────────────────────────────────────────────────────
function Skeleton({ width = '100%', height = 20, style = {} }) {
  return (
    <div style={{ width, height, background: 'var(--color-background-secondary)',
      borderRadius: 6, animation: 'pulse 1.5s ease-in-out infinite', ...style }} />
  );
}

// ─── Login page ───────────────────────────────────────────────────────────────
function LoginPage({ onLogin }) {
  return (
    <div style={{ minHeight: '100vh', display: 'flex', alignItems: 'center', justifyContent: 'center',
      background: 'var(--color-background-tertiary)' }}>
      <Card style={{ maxWidth: 420, width: '100%', textAlign: 'center', padding: '2.5rem' }}>
        <div style={{ width: 56, height: 56, borderRadius: 14, background: '#E6F1FB',
          display: 'flex', alignItems: 'center', justifyContent: 'center', margin: '0 auto 1.5rem' }}>
          <Shield size={28} color="#185FA5" />
        </div>
        <h1 style={{ fontSize: 22, fontWeight: 500, margin: '0 0 8px', color: 'var(--color-text-primary)' }}>
          M365 Tenant Snapshot
        </h1>
        <p style={{ fontSize: 14, color: 'var(--color-text-secondary)', margin: '0 0 2rem', lineHeight: 1.6 }}>
          Sign in with your Microsoft 365 admin account to view a comprehensive security and configuration snapshot of your tenant.
        </p>
        <button onClick={onLogin} style={{ width: '100%', padding: '10px 20px', borderRadius: 8, border: '0.5px solid #B5D4F4',
          background: '#185FA5', color: '#fff', fontSize: 14, fontWeight: 500, cursor: 'pointer', display: 'flex',
          alignItems: 'center', justifyContent: 'center', gap: 8 }}>
          <Shield size={16} />
          Sign in with Microsoft
        </button>
        <p style={{ fontSize: 11, color: 'var(--color-text-secondary)', marginTop: '1.5rem', lineHeight: 1.5 }}>
          Requires Global Reader or higher. Data is fetched live from Microsoft Graph and never stored.
        </p>
      </Card>
    </div>
  );
}

// ─── Navbar ───────────────────────────────────────────────────────────────────
function Navbar({ userName, userEmail, tenantName, onSignOut, activeTab, setActiveTab }) {
  const [menuOpen, setMenuOpen] = useState(false);
  const tabs = [
    { id: 'overview', label: 'Overview' },
    { id: 'identity', label: 'Identity & Access' },
    { id: 'devices', label: 'Devices & Intune' },
    { id: 'signins', label: 'Sign-in Activity' },
    { id: 'security', label: 'Security' },
    { id: 'email', label: 'Email & Compliance' },
  ];
  const initials = (userName || userEmail || '?').split(' ').map(p => p[0]).slice(0, 2).join('').toUpperCase();

  return (
    <header style={{ position: 'sticky', top: 0, zIndex: 50, background: 'var(--color-background-primary)',
      borderBottom: '0.5px solid var(--color-border-tertiary)', boxShadow: '0 2px 8px rgba(0,0,0,0.08)', backdropFilter: 'blur(12px)', WebkitBackdropFilter: 'blur(12px)' }}>
      <div style={{ maxWidth: 1400, margin: '0 auto', padding: '0 20px' }}>
        <div style={{ height: 52, display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
            <div style={{ width: 32, height: 32, borderRadius: 8, background: '#185FA5',
              display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 }}>
              <Shield size={16} color="#fff" />
            </div>
            <div>
              <div style={{ fontSize: 13, fontWeight: 500, color: 'var(--color-text-primary)', lineHeight: 1 }}>
                M365 Tenant Snapshot
              </div>
              {tenantName && (
                <div style={{ fontSize: 11, color: 'var(--color-text-secondary)', lineHeight: 1, marginTop: 2 }}>
                  {tenantName}
                </div>
              )}
            </div>
          </div>
          <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
            <div style={{ position: 'relative' }}>
              <button onClick={() => setMenuOpen(o => !o)}
                style={{ display: 'flex', alignItems: 'center', gap: 6, padding: '5px 10px', borderRadius: 7,
                  border: '0.5px solid var(--color-border-secondary)', background: 'transparent', cursor: 'pointer',
                  color: 'var(--color-text-primary)', fontSize: 13 }}>
                <div style={{ width: 24, height: 24, borderRadius: '50%', background: '#E6F1FB',
                  display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 11, fontWeight: 500, color: '#185FA5' }}>
                  {initials}
                </div>
                <span style={{ maxWidth: 160, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                  {userName || userEmail}
                </span>
                <ChevronDown size={13} color="var(--color-text-secondary)" />
              </button>
              {menuOpen && (
                <div style={{ position: 'absolute', right: 0, top: '100%', marginTop: 4, width: 200,
                  background: 'var(--color-background-primary)', borderRadius: 9, border: '0.5px solid var(--color-border-tertiary)',
                  padding: '4px 0', zIndex: 100 }}>
                  <div style={{ padding: '8px 12px', borderBottom: '0.5px solid var(--color-border-tertiary)' }}>
                    <div style={{ fontSize: 11, color: 'var(--color-text-secondary)', wordBreak: 'break-all' }}>{userEmail}</div>
                  </div>
                  <button onClick={() => { onSignOut(); setMenuOpen(false); }}
                    style={{ width: '100%', padding: '8px 12px', display: 'flex', alignItems: 'center', gap: 8,
                      fontSize: 13, color: 'var(--color-text-primary)', background: 'transparent', border: 'none', cursor: 'pointer', textAlign: 'left' }}>
                    <LogOut size={14} color="var(--color-text-secondary)" />
                    Sign out
                  </button>
                </div>
              )}
            </div>
          </div>
        </div>
        <div style={{ display: 'flex', gap: 2, paddingBottom: 0, overflowX: 'auto' }}>
          {tabs.map(tab => (
            <button key={tab.id} onClick={() => setActiveTab(tab.id)}
              style={{ padding: '6px 14px', fontSize: 13, fontWeight: activeTab === tab.id ? 500 : 400,
                color: activeTab === tab.id ? '#185FA5' : 'var(--color-text-secondary)',
                background: 'transparent', border: 'none', borderBottom: activeTab === tab.id ? '2px solid #185FA5' : '2px solid transparent',
                cursor: 'pointer', whiteSpace: 'nowrap', transition: 'all 0.15s' }}>
              {tab.label}
            </button>
          ))}
        </div>
      </div>
    </header>
  );
}

// ─── License friendly name map ────────────────────────────────────────────────
const LICENSE_NAMES = {
  'O365_BUSINESS_ESSENTIALS':'Microsoft 365 Business Basic','O365_BUSINESS_PREMIUM':'Microsoft 365 Business Standard',
  'SMB_BUSINESS':'Microsoft 365 Apps for Business','SMB_BUSINESS_ESSENTIALS':'Microsoft 365 Business Basic',
  'SMB_BUSINESS_PREMIUM':'Microsoft 365 Business Standard','SPB':'Microsoft 365 Business Premium',
  'MICROSOFT_365_BUSINESS':'Microsoft 365 Business Premium','SPE_E3':'Microsoft 365 E3','SPE_E5':'Microsoft 365 E5',
  'ENTERPRISEPACK':'Microsoft 365 E3','ENTERPRISEPREMIUM':'Microsoft 365 E5',
  'ENTERPRISEPREMIUM_NOPSTNCONF':'Microsoft 365 E5 (No Audio Conf)','SPE_F1':'Microsoft 365 F3',
  'M365_F1':'Microsoft 365 F1','DESKLESSPACK':'Microsoft 365 F1','DESKLESSWOFFPACK':'Microsoft 365 F3',
  'STANDARDPACK':'Office 365 E1','STANDARDWOFFPACK':'Office 365 F3','DEVELOPERPACK':'Office 365 E3 Developer',
  'AAD_PREMIUM':'Microsoft Entra ID P1','AAD_PREMIUM_P2':'Microsoft Entra ID P2','AAD_BASIC':'Microsoft Entra ID Basic',
  'EMS':'Enterprise Mobility + Security E3','EMSPREMIUM':'Enterprise Mobility + Security E5',
  'INTUNE_A':'Microsoft Intune Plan 1','INTUNE_A_D':'Microsoft Intune Plan 1 for Education',
  'EXCHANGESTANDARD':'Exchange Online Plan 1','EXCHANGEENTERPRISE':'Exchange Online Plan 2',
  'EXCHANGE_S_DESKLESS':'Exchange Online Kiosk','EXCHANGEARCHIVE_ADDON':'Exchange Online Archiving',
  'SHAREPOINTSTANDARD':'SharePoint Online Plan 1','SHAREPOINTENTERPRISE':'SharePoint Online Plan 2',
  'MCOMEETADV':'Microsoft Teams Audio Conferencing','MCOEV':'Microsoft Teams Phone Standard',
  'MCOPSTN1':'Teams Domestic Calling Plan','MCOPSTN2':'Teams Domestic & International Calling Plan',
  'Teams_Ess':'Microsoft Teams Essentials','TEAMS_EXPLORATORY':'Microsoft Teams Exploratory',
  'ATP_ENTERPRISE':'Microsoft Defender for Office 365 P1','THREAT_INTELLIGENCE':'Microsoft Defender for Office 365 P2',
  'WIN_DEF_ATP':'Microsoft Defender for Endpoint P2','DEFENDER_ENDPOINT_P1':'Microsoft Defender for Endpoint P1',
  'FLOW_FREE':'Power Automate Free','POWERAPPS_VIRAL':'Power Apps Trial',
  'POWER_BI_STANDARD':'Power BI (Free)','POWER_BI_PRO':'Power BI Pro','POWER_BI_PREMIUM_USER':'Power BI Premium Per User',
  'PROJECTPREMIUM':'Project Plan 5','PROJECTPROFESSIONAL':'Project Plan 3','PROJECT_PLAN1_DEPT':'Project Plan 1',
  'VISIOCLIENT':'Visio Plan 2','VISIOONLINE_PLAN1':'Visio Plan 1',
  'WIN10_PRO_ENT_SUB':'Windows 10/11 Enterprise E3','WIN_ENT_E5':'Windows 10/11 Enterprise E5',
  'DYN365_ENTERPRISE_PLAN1':'Dynamics 365 Customer Engagement Plan','DYN365_ENTERPRISE_SALES':'Dynamics 365 Sales Enterprise',
  'DYN365_FINANCIALS_BUSINESS_SKU':'Dynamics 365 Business Central',
  'RIGHTSMANAGEMENT':'Azure Information Protection P1','MIDSIZEPACK':'Office 365 Midsize Business',
  'BUSINESS_VOICE_DIRECTROUTING':'Microsoft 365 Business Voice','POWERAPPS_DEV':'Power Apps Developer Plan', 
  'CPC_E_2C_8GB_128GB':'Windows 365 Enterprise 2vCPU 8GB','PHONESYSTEM_VIRTUALUSER': 'Teams Phone Resource Account',
  'CCIBOTS_PRIVPREV_VIRAL':'Copilot Studio Viral Trial','Microsoft_365_Copilot':'Microsoft 365 Copilot',
  'COPILOT_STUDIO_VIRAL':'Copilot Studio Trial','MCOEV_VIRTUALUSER':'Teams Phone Resource Account',
};
function getFriendlyLicenseName(sku) { return LICENSE_NAMES[sku] || sku; }

// ─── Overview Tab ─────────────────────────────────────────────────────────────
function OverviewTab({ data, dateDays, setDateDays, onRefresh, setActiveTab }) {
  const { org, users, licenses, caStats, signInStats, intuneStats, securityIndicators, entraTier } = data;

  const licenseRows = useMemo(() => {
    if (!licenses) return [];
    return licenses.map(l => ({
      sku: l.skuPartNumber,
      name: getFriendlyLicenseName(l.skuPartNumber),
      total: l.prepaidUnits?.enabled || 0,
      used: l.consumedUnits || 0,
    })).filter(l => l.total > 0).sort((a, b) => b.used - a.used).slice(0, 12);
  }, [licenses]);

  return (
    <div style={{ display: 'grid', gap: 16 }}>
      {/* Date period selector */}
      <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
        <span style={{ fontSize: 13, color: 'var(--color-text-secondary)' }}>Sign-in data period:</span>
        {[7, 14, 30].map(d => (
          <button key={d} onClick={() => { setDateDays(d); onRefresh(d); }}
            style={{ padding: '4px 14px', borderRadius: 99, fontSize: 13, fontWeight: dateDays === d ? 500 : 400,
              border: `0.5px solid ${dateDays === d ? '#185FA5' : 'var(--color-border-secondary)'}`,
              background: dateDays === d ? '#E6F1FB' : 'var(--color-background-primary)',
              color: dateDays === d ? '#185FA5' : 'var(--color-text-secondary)', cursor: 'pointer' }}>
            {d} days
          </button>
        ))}
      </div>

      {/* Org identity banner */}
      {org && (
        <Card style={{ background: 'linear-gradient(135deg, #0C447C 0%, #185FA5 100%)', border: 'none' }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 16 }}>
            <div style={{ width: 52, height: 52, borderRadius: 12, background: 'rgba(255,255,255,0.15)',
              display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 }}>
              <Building2 size={26} color="#fff" />
            </div>
            <div>
              <div style={{ fontSize: 20, fontWeight: 500, color: '#fff' }}>{org.displayName}</div>
              <div style={{ fontSize: 13, color: 'rgba(255,255,255,0.7)', marginTop: 3 }}>
                {org.verifiedDomains?.find(d => d.isDefault)?.name || org.id} · Tenant ID: {org.id}
              </div>
              <div style={{ display: 'flex', gap: 8, marginTop: 8, flexWrap: 'wrap' }}>
                {org.assignedPlans?.some(p => p.servicePlanId && p.capabilityStatus === 'Enabled') && (
                  <span style={{ fontSize: 11, background: 'rgba(255,255,255,0.2)', color: '#fff',
                    padding: '2px 8px', borderRadius: 99 }}>Active</span>
                )}
                <span style={{ fontSize: 11, background: 'rgba(255,255,255,0.2)', color: '#fff',
                  padding: '2px 8px', borderRadius: 99 }}>
                  {org.countryLetterCode || 'Unknown country'}
                </span>
                {entraTier && (
                  <span style={{ fontSize: 11, borderRadius: 99, padding: '2px 10px', fontWeight: 500,
                    background: entraTier.includes('P2') ? '#FAC775' : 'rgba(255,255,255,0.22)',
                    color: entraTier.includes('P2') ? '#412402' : '#fff',
                    border: entraTier.includes('P2') ? 'none' : '0.5px solid rgba(255,255,255,0.4)' }}>
                    {entraTier}
                  </span>
                )}
              </div>
            </div>
          </div>
        </Card>
      )}

      {/* Key metrics grid — clickable */}
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(160px, 1fr))', gap: 12 }}>
        {[
          { label: 'Total Users', value: users?.total?.toLocaleString(), icon: Users, color: '#185FA5', bgColor: '#E6F1FB', sub: `${users?.enabled?.toLocaleString() || 0} enabled`, tab: 'identity' },
          { label: 'Guest Users', value: users?.guests?.toLocaleString(), icon: Globe, color: '#0F6E56', bgColor: '#E1F5EE', sub: 'External identities', tab: 'identity' },
          { label: 'Global Admins', value: users?.globalAdmins?.toLocaleString(), icon: Crown, color: users?.globalAdmins > 3 ? '#A32D2D' : '#185FA5', bgColor: users?.globalAdmins > 3 ? '#FCEBEB' : '#E6F1FB', sub: users?.globalAdmins > 3 ? 'Consider reducing' : 'Recommended ≤ 3', tab: 'security' },
          { label: 'CA Policies', value: caStats?.total, icon: ShieldCheck, color: '#534AB7', bgColor: '#EEEDFE', sub: `${caStats?.enabled || 0} enabled`, tab: 'identity' },
          { label: 'Intune Devices', value: intuneStats?.total?.toLocaleString(), icon: Monitor, color: '#0F6E56', bgColor: '#E1F5EE', sub: `${intuneStats?.compliant || 0} compliant`, tab: 'devices' },
          { label: `Sign-ins (${dateDays}d)`, value: signInStats?.total?.toLocaleString(), icon: Activity, color: '#185FA5', bgColor: '#E6F1FB', sub: `${signInStats?.failPct || 0}% failures`, tab: 'signins' },
        ].map(card => (
          <div key={card.label} onClick={() => setActiveTab(card.tab)}
            style={{ cursor: 'pointer', transition: 'transform 0.1s, box-shadow 0.1s', borderRadius: 'var(--border-radius-md)' }}
            onMouseEnter={e => { e.currentTarget.style.transform = 'translateY(-2px)'; e.currentTarget.style.boxShadow = '0 4px 12px rgba(0,0,0,0.10)'; }}
            onMouseLeave={e => { e.currentTarget.style.transform = ''; e.currentTarget.style.boxShadow = ''; }}>
            <MetricCard label={card.label} value={card.value} icon={card.icon} color={card.color} bgColor={card.bgColor} sub={card.sub} />
          </div>
        ))}
      </div>

      {/* Security posture + license usage */}
      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 16 }}>
        <Card>
          <SectionTitle icon={ShieldAlert} title="Security posture" sub="Key configuration indicators" />
          <div style={{ display: 'grid', gap: 8 }}>
            {securityIndicators?.map((ind, i) => (
              <SecurityIndicator key={i} {...ind} />
            ))}
          </div>
        </Card>
        <Card>
          <SectionTitle icon={Key} title="License allocation" sub="Purchased vs consumed" />
          <div style={{ display: 'grid', gap: 8 }}>
            {licenseRows.map(l => (
              <div key={l.sku}>
                <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: 12, marginBottom: 3 }}>
                  <span style={{ color: 'var(--color-text-primary)', fontWeight: 500, maxWidth: '70%',
                    overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{l.name}</span>
                  <span style={{ color: 'var(--color-text-secondary)' }}>{l.used} / {l.total}</span>
                </div>
                <div style={{ height: 5, background: 'var(--color-background-secondary)', borderRadius: 99, overflow: 'hidden' }}>
                  <div style={{ height: '100%', width: `${Math.min(100, (l.used / l.total) * 100)}%`,
                    background: l.used / l.total > 0.9 ? '#E24B4A' : l.used / l.total > 0.75 ? '#BA7517' : '#185FA5',
                    borderRadius: 99, transition: 'width 0.4s ease' }} />
                </div>
              </div>
            ))}
            {licenseRows.length === 0 && (
              <div style={{ color: 'var(--color-text-secondary)', fontSize: 13 }}>No license data</div>
            )}
          </div>
        </Card>
      </div>
    </div>
  );
}

// ─── Identity Tab ─────────────────────────────────────────────────────────────
function IdentityTab({ data }) {
  const { users, caDetails, namedLocations, authMethods } = data;
  const [search, setSearch] = useState('');
  const [typeFilter, setTypeFilter] = useState('all'); // all | member | guest
  const [statusFilter, setStatusFilter] = useState('all'); // all | active | disabled
  const [caFilter, setCaFilter] = useState('all'); // all | enabled | disabled | reportOnly

  const filteredUsers = useMemo(() => {
    if (!users?.list) return [];
    const q = search.toLowerCase();
    return users.list.filter(u => {
      if (q && !u.displayName?.toLowerCase().includes(q) && !u.userPrincipalName?.toLowerCase().includes(q)) return false;
      if (typeFilter === 'member' && u.userType === 'Guest') return false;
      if (typeFilter === 'guest' && u.userType !== 'Guest') return false;
      if (statusFilter === 'active' && !u.accountEnabled) return false;
      if (statusFilter === 'disabled' && u.accountEnabled) return false;
      return true;
    }).slice(0, 100);
  }, [users, search, typeFilter, statusFilter]);

  const filteredCA = useMemo(() => {
    if (!caDetails) return [];
    if (caFilter === 'all') return caDetails;
    if (caFilter === 'enabled') return caDetails.filter(p => p.state === 'enabled');
    if (caFilter === 'disabled') return caDetails.filter(p => p.state === 'disabled');
    if (caFilter === 'reportOnly') return caDetails.filter(p => p.state === 'enabledForReportingButNotEnforced');
    return caDetails;
  }, [caDetails, caFilter]);

  const policyStateColor = (state) => ({
    enabled: { bg: '#EAF3DE', color: '#27500A', border: '#C0DD97' },
    disabled: { bg: 'var(--color-background-secondary)', color: 'var(--color-text-secondary)', border: 'var(--color-border-tertiary)' },
    enabledForReportingButNotEnforced: { bg: '#FAEEDA', color: '#633806', border: '#FAC775' },
  }[state] || { bg: 'var(--color-background-secondary)', color: 'var(--color-text-secondary)', border: 'var(--color-border-tertiary)' });

  return (
    <div style={{ display: 'grid', gap: 16 }}>
      {/* User overview */}
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(150px, 1fr))', gap: 12 }}>
        <MetricCard label="Total Users" value={users?.total} icon={Users} color="#185FA5" bgColor="#E6F1FB" />
        <MetricCard label="Enabled" value={users?.enabled} icon={UserCheck} color="#3B6D11" bgColor="#EAF3DE" />
        <MetricCard label="Disabled" value={users?.disabled} icon={UserX} color="#A32D2D" bgColor="#FCEBEB" />
        <MetricCard label="Guests" value={users?.guests} icon={Globe} color="#0F6E56" bgColor="#E1F5EE" />
        <MetricCard label="Global Admins" value={users?.globalAdmins} icon={Crown}
          color={users?.globalAdmins > 3 ? '#A32D2D' : '#185FA5'}
          bgColor={users?.globalAdmins > 3 ? '#FCEBEB' : '#E6F1FB'}
          sub="Role members" />
        <MetricCard label="MFA Registered" value={authMethods?.mfaPct != null ? `${authMethods.mfaPct}%` : '—'}
          icon={Fingerprint} color="#534AB7" bgColor="#EEEDFE"
          sub={authMethods?.mfaRegistered != null ? `${authMethods.mfaRegistered} users` : ''} />
      </div>

      {/* User table */}
      <Card>
        <div style={{ marginBottom: 14 }}>
          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 10 }}>
            <SectionTitle icon={Users} title="User directory" sub={`${filteredUsers.length} shown of ${users?.total || 0}`} />
          </div>
          <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap', alignItems: 'center' }}>
            <div style={{ position: 'relative', flex: '1 1 180px', minWidth: 140 }}>
              <Search size={13} color="var(--color-text-secondary)"
                style={{ position: 'absolute', left: 9, top: '50%', transform: 'translateY(-50%)' }} />
              <input value={search} onChange={e => setSearch(e.target.value)}
                placeholder="Search by name or UPN…"
                style={{ paddingLeft: 28, paddingRight: 10, height: 30, fontSize: 12, borderRadius: 7, width: '100%',
                  border: '0.5px solid var(--color-border-secondary)', background: 'var(--color-background-secondary)',
                  color: 'var(--color-text-primary)', outline: 'none' }} />
            </div>
            {[
              { key: 'typeFilter', val: typeFilter, set: setTypeFilter, opts: [['all','All types'],['member','Members only'],['guest','Guests only']] },
              { key: 'statusFilter', val: statusFilter, set: setStatusFilter, opts: [['all','All status'],['active','Active only'],['disabled','Disabled only']] },
            ].map(f => (
              <select key={f.key} value={f.val} onChange={e => f.set(e.target.value)}
                style={{ height: 30, padding: '0 8px', fontSize: 12, borderRadius: 7, cursor: 'pointer',
                  border: '0.5px solid var(--color-border-secondary)', background: 'var(--color-background-secondary)',
                  color: 'var(--color-text-primary)', outline: 'none' }}>
                {f.opts.map(([v, l]) => <option key={v} value={v}>{l}</option>)}
              </select>
            ))}
          </div>
        </div>
        <div style={{ overflowX: 'auto' }}>
          <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
            <thead>
              <tr style={{ borderBottom: '0.5px solid var(--color-border-tertiary)' }}>
                {['Display name', 'UPN', 'Type', 'Last sign-in', 'Last pwd change', 'Status'].map(h => (
                  <th key={h} style={{ padding: '6px 10px', textAlign: 'left', fontSize: 11, fontWeight: 500,
                    color: 'var(--color-text-secondary)', whiteSpace: 'nowrap' }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {filteredUsers.map(u => {
                const days = daysSince(u.signInActivity?.lastSignInDateTime);
                return (
                  <tr key={u.id} style={{ borderBottom: '0.5px solid var(--color-border-tertiary)' }}>
                    <td style={{ padding: '7px 10px', fontWeight: 500, color: 'var(--color-text-primary)',
                      whiteSpace: 'nowrap', maxWidth: 180, overflow: 'hidden', textOverflow: 'ellipsis' }}>
                      {u.displayName}
                    </td>
                    <td style={{ padding: '7px 10px', color: 'var(--color-text-secondary)',
                      maxWidth: 220, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                      {u.userPrincipalName}
                    </td>
                    <td style={{ padding: '7px 10px' }}>
                      <span style={{ fontSize: 11, padding: '2px 7px', borderRadius: 99,
                        background: u.userType === 'Guest' ? '#FAEEDA' : '#E6F1FB',
                        color: u.userType === 'Guest' ? '#633806' : '#0C447C',
                        border: `0.5px solid ${u.userType === 'Guest' ? '#FAC775' : '#B5D4F4'}` }}>
                        {u.userType || 'Member'}
                      </span>
                    </td>
                    <td style={{ padding: '7px 10px', whiteSpace: 'nowrap' }}>
                      {u.signInActivity?.lastSignInDateTime ? (
                        <span style={{ color: days > 90 ? '#A32D2D' : days > 30 ? '#854F0B' : 'var(--color-text-secondary)', fontSize: 12 }}>
                          {fmtDate(u.signInActivity.lastSignInDateTime)}
                          {days != null && <span style={{ marginLeft: 4, fontSize: 11, opacity: 0.7 }}>({days}d)</span>}
                        </span>
                      ) : <span style={{ color: 'var(--color-text-secondary)', fontSize: 11 }}>Never</span>}
                    </td>
                    <td style={{ padding: '7px 10px', color: 'var(--color-text-secondary)', whiteSpace: 'nowrap', fontSize: 12 }}>
                      {fmtDate(u.lastPasswordChangeDateTime)}
                    </td>
                    <td style={{ padding: '7px 10px' }}>
                      <StatusPill on={u.accountEnabled} labelOn="Active" labelOff="Disabled" />
                    </td>
                  </tr>
                );
              })}
              {filteredUsers.length === 0 && (
                <tr><td colSpan={6} style={{ padding: '20px 10px', textAlign: 'center', color: 'var(--color-text-secondary)', fontSize: 13 }}>
                  No users found
                </td></tr>
              )}
            </tbody>
          </table>
        </div>
      </Card>

      {/* CA Policies */}
      <Card>
        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 14 }}>
          <SectionTitle icon={ShieldCheck} title="Conditional Access policies"
            sub={`${filteredCA.length} of ${caDetails?.length || 0} policies`} />
          <select value={caFilter} onChange={e => setCaFilter(e.target.value)}
            style={{ height: 30, padding: '0 8px', fontSize: 12, borderRadius: 7, cursor: 'pointer',
              border: '0.5px solid var(--color-border-secondary)', background: 'var(--color-background-secondary)',
              color: 'var(--color-text-primary)', outline: 'none' }}>
            <option value="all">All states</option>
            <option value="enabled">Enabled only</option>
            <option value="reportOnly">Report only</option>
            <option value="disabled">Disabled only</option>
          </select>
        </div>
        <div style={{ display: 'grid', gap: 8 }}>
          {(filteredCA || []).map(p => {
            const st = policyStateColor(p.state);
            return (
              <div key={p.id} style={{ display: 'flex', alignItems: 'flex-start', gap: 10,
                padding: '10px 12px', borderRadius: 8, border: '0.5px solid var(--color-border-tertiary)',
                background: 'var(--color-background-secondary)' }}>
                <div style={{ marginTop: 2, flexShrink: 0 }}>
                  {p.state === 'enabled' ? <ShieldCheck size={15} color="#3B6D11" /> :
                   p.state === 'enabledForReportingButNotEnforced' ? <Eye size={15} color="#854F0B" /> :
                   <ShieldOff size={15} color="#888780" />}
                </div>
                <div style={{ flex: 1, minWidth: 0 }}>
                  <div style={{ fontSize: 13, fontWeight: 500, color: 'var(--color-text-primary)', marginBottom: 3 }}>
                    {p.displayName}
                  </div>
                  <div style={{ display: 'flex', gap: 6, flexWrap: 'wrap', marginBottom: 4 }}>
                    <span style={{ fontSize: 11, padding: '1px 7px', borderRadius: 99,
                      background: st.bg, color: st.color, border: `0.5px solid ${st.border}` }}>
                      {p.state === 'enabledForReportingButNotEnforced' ? 'Report only' : p.state}
                    </span>
                    {p.conditions?.users?.includeUsers?.includes('All') && (
                      <span style={{ fontSize: 11, padding: '1px 7px', borderRadius: 99,
                        background: '#EEEDFE', color: '#3C3489', border: '0.5px solid #AFA9EC' }}>All users</span>
                    )}
                    {p.grantControls?.builtInControls?.includes('mfa') && (
                      <span style={{ fontSize: 11, padding: '1px 7px', borderRadius: 99,
                        background: '#E1F5EE', color: '#085041', border: '0.5px solid #5DCAA5' }}>Requires MFA</span>
                    )}
                    {p.grantControls?.builtInControls?.includes('compliantDevice') && (
                      <span style={{ fontSize: 11, padding: '1px 7px', borderRadius: 99,
                        background: '#E6F1FB', color: '#042C53', border: '0.5px solid #85B7EB' }}>Compliant device</span>
                    )}
                  </div>
                  <div style={{ fontSize: 11, color: 'var(--color-text-secondary)', display: 'flex', gap: 12 }}>
                    {p.createdDateTime && <span>Created: {fmtDate(p.createdDateTime)}</span>}
                    {p.modifiedDateTime && <span>Modified: {fmtDate(p.modifiedDateTime)}</span>}
                  </div>
                </div>
              </div>
            );
          })}
          {(!caDetails || caDetails.length === 0) && (
            <div style={{ color: 'var(--color-text-secondary)', fontSize: 13, padding: '12px 0' }}>
              No Conditional Access policies found
            </div>
          )}
        </div>
      </Card>

      {/* Named Locations */}
      <Card>
        <SectionTitle icon={MapPin} title="Named locations"
          sub={`${namedLocations?.length || 0} locations configured`} />
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(280px, 1fr))', gap: 8 }}>
          {(namedLocations || []).map(loc => (
            <div key={loc.id} style={{ padding: '10px 12px', borderRadius: 8,
              border: '0.5px solid var(--color-border-tertiary)', background: 'var(--color-background-secondary)' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: 7, marginBottom: 4 }}>
                {loc.isTrusted ? <ShieldCheck size={13} color="#3B6D11" /> : <MapPin size={13} color="#888780" />}
                <span style={{ fontSize: 13, fontWeight: 500, color: 'var(--color-text-primary)' }}>{loc.displayName}</span>
              </div>
              <div style={{ fontSize: 11, color: 'var(--color-text-secondary)' }}>
                {loc['@odata.type']?.includes('countryNamed') ? 'Country-based' : 'IP range-based'}
                {loc.isTrusted && ' · Trusted'}
                {loc.ipRanges && ` · ${loc.ipRanges.length} range(s)`}
                {loc.countriesAndRegions && ` · ${loc.countriesAndRegions.join(', ')}`}
              </div>
            </div>
          ))}
          {(!namedLocations || namedLocations.length === 0) && (
            <div style={{ color: 'var(--color-text-secondary)', fontSize: 13 }}>No named locations configured</div>
          )}
        </div>
      </Card>
    </div>
  );
}

// ─── Devices & Intune Tab ─────────────────────────────────────────────────────
function DevicesTab({ data }) {
  const { intune } = data;

  const complianceData = useMemo(() => {
    if (!intune?.complianceByState) return [];
    return Object.entries(intune.complianceByState).map(([name, value]) => ({ name, value }));
  }, [intune]);

  const osPie = useMemo(() => {
    if (!intune?.byOS) return [];
    return Object.entries(intune.byOS).map(([name, value]) => ({ name, value }));
  }, [intune]);

  const pieColors = ['#378ADD', '#639922', '#BA7517', '#E24B4A', '#7F77DD', '#1D9E75'];

  return (
    <div style={{ display: 'grid', gap: 16 }}>
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(150px, 1fr))', gap: 12 }}>
        <MetricCard label="Total Devices" value={intune?.total} icon={Monitor} color="#185FA5" bgColor="#E6F1FB" />
        <MetricCard label="Compliant" value={intune?.compliant} icon={CheckCircle2} color="#3B6D11" bgColor="#EAF3DE" />
        <MetricCard label="Non-compliant" value={intune?.nonCompliant} icon={XCircle} color="#A32D2D" bgColor="#FCEBEB" />
        <MetricCard label="Enrolled (30d)" value={intune?.recentEnrollments} icon={Smartphone} color="#534AB7" bgColor="#EEEDFE" />
        <MetricCard label="Compliance Policies" value={intune?.compliancePolicies} icon={ClipboardList} color="#0F6E56" bgColor="#E1F5EE" />
        <MetricCard label="Config Profiles" value={intune?.configProfiles} icon={Settings2} color="#854F0B" bgColor="#FAEEDA" />
      </div>

      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 16 }}>
        <Card>
          <SectionTitle icon={BarChart3} title="Compliance status" sub="Device compliance breakdown" />
          {complianceData.length > 0 ? (
            <ResponsiveContainer width="100%" height={200}>
              <PieChart>
                <Pie data={complianceData} cx="50%" cy="50%" innerRadius="45%" outerRadius="70%"
                  dataKey="value" paddingAngle={2}>
                  {complianceData.map((_, i) => (
                    <Cell key={i} fill={pieColors[i % pieColors.length]} stroke="none" />
                  ))}
                </Pie>
                <Tooltip formatter={(v) => [v.toLocaleString(), 'Devices']}
                  contentStyle={{ fontSize: 12, borderRadius: 8, border: '0.5px solid var(--color-border-tertiary)' }} />
                <Legend iconType="circle" iconSize={8} wrapperStyle={{ fontSize: 12 }} />
              </PieChart>
            </ResponsiveContainer>
          ) : <div style={{ color: 'var(--color-text-secondary)', fontSize: 13, height: 200, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>No data</div>}
        </Card>
        <Card>
          <SectionTitle icon={Laptop} title="OS distribution" sub="By operating system" />
          {osPie.length > 0 ? (
            <ResponsiveContainer width="100%" height={200}>
              <BarChart data={osPie} margin={{ left: 0, right: 10 }}>
                <CartesianGrid strokeDasharray="3 3" stroke="var(--color-border-tertiary)" vertical={false} />
                <XAxis dataKey="name" tick={{ fontSize: 11, fill: 'var(--color-text-secondary)' }} axisLine={false} tickLine={false} />
                <YAxis tick={{ fontSize: 11, fill: 'var(--color-text-secondary)' }} axisLine={false} tickLine={false} width={30} />
                <Tooltip contentStyle={{ fontSize: 12, borderRadius: 8, border: '0.5px solid var(--color-border-tertiary)' }} />
                <Bar dataKey="value" name="Devices" radius={[4, 4, 0, 0]} maxBarSize={40}>
                  {osPie.map((_, i) => <Cell key={i} fill={pieColors[i % pieColors.length]} />)}
                </Bar>
              </BarChart>
            </ResponsiveContainer>
          ) : <div style={{ color: 'var(--color-text-secondary)', fontSize: 13, height: 200, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>No data</div>}
        </Card>
      </div>

      {/* Compliance policies */}
      <Card>
        <SectionTitle icon={ClipboardList} title="Compliance policies" sub="Intune device compliance rules" />
        <div style={{ display: 'grid', gap: 8 }}>
          {(intune?.compliancePolicyList || []).map(p => (
            <div key={p.id} style={{ display: 'flex', alignItems: 'center', gap: 10, padding: '10px 12px',
              borderRadius: 8, border: '0.5px solid var(--color-border-tertiary)', background: 'var(--color-background-secondary)' }}>
              <ClipboardList size={14} color="#185FA5" style={{ flexShrink: 0 }} />
              <div style={{ flex: 1 }}>
                <div style={{ fontSize: 13, fontWeight: 500, color: 'var(--color-text-primary)' }}>{p.displayName}</div>
                <div style={{ fontSize: 11, color: 'var(--color-text-secondary)' }}>
                  {p['@odata.type']?.replace('#microsoft.graph.', '').replace('CompliancePolicy', '')} ·{' '}
                  Created {fmtDate(p.createdDateTime)}
                </div>
              </div>
            </div>
          ))}
          {(!intune?.compliancePolicyList || intune.compliancePolicyList.length === 0) && (
            <div style={{ color: 'var(--color-text-secondary)', fontSize: 13 }}>No compliance policies found</div>
          )}
        </div>
      </Card>

      {/* Config profiles */}
      <Card>
        <SectionTitle icon={Settings2} title="Configuration profiles" sub="Intune device configuration" />
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(280px, 1fr))', gap: 8 }}>
          {(intune?.configProfileList || []).map(p => (
            <div key={p.id} style={{ padding: '10px 12px', borderRadius: 8,
              border: '0.5px solid var(--color-border-tertiary)', background: 'var(--color-background-secondary)' }}>
              <div style={{ fontSize: 13, fontWeight: 500, color: 'var(--color-text-primary)', marginBottom: 3 }}>{p.displayName}</div>
              <div style={{ fontSize: 11, color: 'var(--color-text-secondary)' }}>
                {p.platformType || p['@odata.type']?.replace('#microsoft.graph.', '') || 'Unknown platform'}
              </div>
            </div>
          ))}
          {(!intune?.configProfileList || intune.configProfileList.length === 0) && (
            <div style={{ color: 'var(--color-text-secondary)', fontSize: 13 }}>No configuration profiles found</div>
          )}
        </div>
      </Card>
    </div>
  );
}

// ─── Sign-in Activity Tab ─────────────────────────────────────────────────────
const PIE_COLORS = ['#378ADD','#639922','#BA7517','#E24B4A','#7F77DD','#1D9E75','#D85A30','#0891b2','#D4537E','#888780'];

function SignInsTab({ data }) {
  const { signIns } = data;
  const [osFilter, setOsFilter] = useState('');
  const [appFilter, setAppFilter] = useState('');
  const [userSearch, setUserSearch] = useState('');

  const filtered = useMemo(() => {
    if (!signIns?.list) return [];
    const uq = userSearch.toLowerCase();
    return signIns.list.filter(s =>
      (!osFilter || s.deviceDetail?.operatingSystem?.toLowerCase().includes(osFilter.toLowerCase())) &&
      (!appFilter || s.appDisplayName?.toLowerCase().includes(appFilter.toLowerCase())) &&
      (!uq || s.userDisplayName?.toLowerCase().includes(uq) || s.userPrincipalName?.toLowerCase().includes(uq))
    );
  }, [signIns, osFilter, appFilter, userSearch]);

  const locationData = useMemo(() => {
    const counts = {};
    (signIns?.list || []).forEach(s => {
      const key = s.location?.countryOrRegion || 'Unknown';
      counts[key] = (counts[key] || 0) + 1;
    });
    return Object.entries(counts).sort((a, b) => b[1] - a[1]).slice(0, 10)
      .map(([name, value]) => ({ name, value }));
  }, [signIns]);

  const appData = useMemo(() => {
    const counts = {};
    (signIns?.list || []).forEach(s => {
      const key = s.appDisplayName || 'Unknown';
      counts[key] = (counts[key] || 0) + 1;
    });
    return Object.entries(counts).sort((a, b) => b[1] - a[1]).slice(0, 8)
      .map(([name, value]) => ({ name, value }));
  }, [signIns]);

  return (
    <div style={{ display: 'grid', gap: 16 }}>
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(150px, 1fr))', gap: 12 }}>
        <MetricCard label="Total (7d)" value={signIns?.total?.toLocaleString()} icon={Activity} color="#185FA5" bgColor="#E6F1FB" />
        <MetricCard label="Successful" value={signIns?.successful?.toLocaleString()} icon={CheckCircle2} color="#3B6D11" bgColor="#EAF3DE" />
        <MetricCard label="Failed" value={signIns?.failed?.toLocaleString()} icon={XCircle} color="#A32D2D" bgColor="#FCEBEB" />
        <MetricCard label="Countries" value={signIns?.uniqueCountries} icon={Globe} color="#534AB7" bgColor="#EEEDFE" />
        <MetricCard label="Unique Apps" value={signIns?.uniqueApps} icon={Server} color="#0F6E56" bgColor="#E1F5EE" />
        <MetricCard label="Risky sign-ins" value={signIns?.risky?.toLocaleString()} icon={AlertTriangle}
          color={signIns?.risky > 0 ? '#A32D2D' : '#3B6D11'}
          bgColor={signIns?.risky > 0 ? '#FCEBEB' : '#EAF3DE'} />
      </div>

      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 16 }}>
        <Card>
          <SectionTitle icon={Globe} title="Top countries" sub="Sign-in origin by country" />
          <ResponsiveContainer width="100%" height={240}>
            <PieChart>
              <Pie data={locationData} cx="50%" cy="50%" innerRadius="35%" outerRadius="65%"
                dataKey="value" paddingAngle={2} label={({ name, percent }) => `${name} ${(percent*100).toFixed(0)}%`}
                labelLine={false} fontSize={10}>
                {locationData.map((_, i) => <Cell key={i} fill={PIE_COLORS[i % PIE_COLORS.length]} stroke="none" />)}
              </Pie>
              <Tooltip formatter={(v, n) => [v.toLocaleString(), 'Sign-ins']}
                contentStyle={{ fontSize: 12, borderRadius: 8, border: '0.5px solid var(--color-border-tertiary)' }} />
              <Legend iconType="circle" iconSize={8} wrapperStyle={{ fontSize: 11 }} />
            </PieChart>
          </ResponsiveContainer>
        </Card>
        <Card>
          <SectionTitle icon={Server} title="Top applications" sub="Most-used apps by sign-in volume" />
          <ResponsiveContainer width="100%" height={240}>
            <PieChart>
              <Pie data={appData} cx="50%" cy="50%" innerRadius="35%" outerRadius="65%"
                dataKey="value" paddingAngle={2} label={({ name, percent }) => percent > 0.05 ? `${(percent*100).toFixed(0)}%` : ''}
                labelLine={false} fontSize={10}>
                {appData.map((_, i) => <Cell key={i} fill={PIE_COLORS[i % PIE_COLORS.length]} stroke="none" />)}
              </Pie>
              <Tooltip formatter={(v, n, p) => [v.toLocaleString(), p?.payload?.name || 'Sign-ins']}
                contentStyle={{ fontSize: 12, borderRadius: 8, border: '0.5px solid var(--color-border-tertiary)' }} />
              <Legend iconType="circle" iconSize={8} wrapperStyle={{ fontSize: 11 }} />
            </PieChart>
          </ResponsiveContainer>
        </Card>
      </div>

      <Card>
        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 14 }}>
          <SectionTitle icon={Activity} title="Sign-in log" sub={`${filtered.length} entries`} />
          <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
            <div style={{ position: 'relative' }}>
              <Search size={12} color="var(--color-text-secondary)"
                style={{ position: 'absolute', left: 8, top: '50%', transform: 'translateY(-50%)' }} />
              <input value={userSearch} onChange={e => setUserSearch(e.target.value)} placeholder="Search user…"
                style={{ paddingLeft: 26, paddingRight: 8, height: 28, fontSize: 12, borderRadius: 6, border: '0.5px solid var(--color-border-secondary)',
                  background: 'var(--color-background-secondary)', color: 'var(--color-text-primary)', outline: 'none', width: 140 }} />
            </div>
            <input value={osFilter} onChange={e => setOsFilter(e.target.value)} placeholder="Filter OS…"
              style={{ height: 28, padding: '0 10px', fontSize: 12, borderRadius: 6, border: '0.5px solid var(--color-border-secondary)',
                background: 'var(--color-background-secondary)', color: 'var(--color-text-primary)', outline: 'none', width: 100 }} />
            <input value={appFilter} onChange={e => setAppFilter(e.target.value)} placeholder="Filter app…"
              style={{ height: 28, padding: '0 10px', fontSize: 12, borderRadius: 6, border: '0.5px solid var(--color-border-secondary)',
                background: 'var(--color-background-secondary)', color: 'var(--color-text-primary)', outline: 'none', width: 110 }} />
          </div>
        </div>
        <div style={{ overflowX: 'auto' }}>
          <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
            <thead>
              <tr style={{ borderBottom: '0.5px solid var(--color-border-tertiary)' }}>
                {['User', 'App', 'Location', 'Device / OS', 'Time', 'Status'].map(h => (
                  <th key={h} style={{ padding: '6px 10px', textAlign: 'left', fontSize: 11, fontWeight: 500,
                    color: 'var(--color-text-secondary)', whiteSpace: 'nowrap' }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {filtered.slice(0, 50).map(s => (
                <tr key={s.id} style={{ borderBottom: '0.5px solid var(--color-border-tertiary)' }}>
                  <td style={{ padding: '6px 10px', maxWidth: 160, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                    <div style={{ fontWeight: 500, color: 'var(--color-text-primary)' }}>{s.userDisplayName}</div>
                    <div style={{ color: 'var(--color-text-secondary)', fontSize: 11 }}>{s.userPrincipalName}</div>
                  </td>
                  <td style={{ padding: '6px 10px', color: 'var(--color-text-secondary)', maxWidth: 140, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                    {s.appDisplayName}
                  </td>
                  <td style={{ padding: '6px 10px', color: 'var(--color-text-secondary)', whiteSpace: 'nowrap', fontSize: 11 }}>
                    {[s.location?.city, s.location?.countryOrRegion].filter(Boolean).join(', ') || '—'}
                  </td>
                  <td style={{ padding: '6px 10px', color: 'var(--color-text-secondary)', whiteSpace: 'nowrap', fontSize: 11 }}>
                    {s.deviceDetail?.operatingSystem || s.clientAppUsed || '—'}
                  </td>
                  <td style={{ padding: '6px 10px', color: 'var(--color-text-secondary)', whiteSpace: 'nowrap', fontSize: 11 }}>
                    {fmtDateTime(s.createdDateTime)}
                  </td>
                  <td style={{ padding: '6px 10px' }}>
                    <StatusPill on={s.status?.errorCode === 0} labelOn="Success" labelOff="Failed" />
                  </td>
                </tr>
              ))}
              {filtered.length === 0 && (
                <tr><td colSpan={6} style={{ padding: '20px 10px', textAlign: 'center', color: 'var(--color-text-secondary)' }}>
                  No sign-in data
                </td></tr>
              )}
            </tbody>
          </table>
        </div>
      </Card>
    </div>
  );
}

// ─── Security Tab ─────────────────────────────────────────────────────────────
function SecurityTab({ data }) {
  const { securityIndicators, riskSummary, globalAdminsList } = data;

  return (
    <div style={{ display: 'grid', gap: 16 }}>
      <Card>
        <SectionTitle icon={Crown} title="Global administrators"
          sub={`${globalAdminsList?.length || 0} accounts with Global Admin role`} />
        {globalAdminsList?.length > 3 && (
          <div style={{ marginBottom: 12, padding: '10px 14px', borderRadius: 8,
            background: globalAdminsList.length > 5 ? '#FCEBEB' : '#FAEEDA',
            border: `0.5px solid ${globalAdminsList.length > 5 ? '#F7C1C1' : '#FAC775'}`,
            fontSize: 13, color: globalAdminsList.length > 5 ? '#791F1F' : '#633806' }}>
            <AlertTriangle size={14} style={{ display: 'inline', marginRight: 6, verticalAlign: -2 }} />
            {globalAdminsList.length} Global Admins detected —{' '}
            {globalAdminsList.length > 5 ? 'Critical: significantly exceeds recommended maximum of 3.' : 'Elevated: above recommended maximum of 3.'}
            {' '}Consider converting excess admins to lower-privilege roles.
          </div>
        )}
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(260px, 1fr))', gap: 8 }}>
          {(globalAdminsList || []).map(u => (
            <div key={u.id} style={{ display: 'flex', alignItems: 'center', gap: 10, padding: '10px 12px',
              borderRadius: 8, border: '0.5px solid var(--color-border-tertiary)', background: 'var(--color-background-secondary)' }}>
              <div style={{ width: 32, height: 32, borderRadius: '50%',
                background: globalAdminsList.length <= 3 ? '#EAF3DE' : globalAdminsList.length <= 5 ? '#FAEEDA' : '#FCEBEB',
                display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 12, fontWeight: 500,
                color: globalAdminsList.length <= 3 ? '#27500A' : globalAdminsList.length <= 5 ? '#633806' : '#791F1F', flexShrink: 0 }}>
                {u.displayName?.split(' ').map(p => p[0]).slice(0, 2).join('').toUpperCase()}
              </div>
              <div style={{ minWidth: 0 }}>
                <div style={{ fontSize: 13, fontWeight: 500, color: 'var(--color-text-primary)', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                  {u.displayName}
                </div>
                <div style={{ fontSize: 11, color: 'var(--color-text-secondary)', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                  {u.userPrincipalName}
                </div>
              </div>
            </div>
          ))}
          {(!globalAdminsList || globalAdminsList.length === 0) && (
            <div style={{ color: 'var(--color-text-secondary)', fontSize: 13 }}>No Global Admin data available</div>
          )}
        </div>
      </Card>

      <Card>
        <SectionTitle icon={ShieldAlert} title="Security posture indicators" sub="Configuration health check" />
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(300px, 1fr))', gap: 10 }}>
          {(securityIndicators || []).map((ind, i) => (
            <SecurityIndicator key={i} {...ind} />
          ))}
        </div>
      </Card>

      {riskSummary && (
        <Card>
          <SectionTitle icon={AlertCircle} title="Risk summary" sub="Entra ID risk aggregation (requires P2)" />
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(140px, 1fr))', gap: 12 }}>
            {Object.entries(riskSummary).map(([level, count]) => (
              <MetricCard key={level} label={`Risk: ${level}`} value={count}
                icon={level === 'high' ? AlertTriangle : level === 'medium' ? AlertCircle : Info}
                color={level === 'high' ? '#A32D2D' : level === 'medium' ? '#854F0B' : '#185FA5'}
                bgColor={level === 'high' ? '#FCEBEB' : level === 'medium' ? '#FAEEDA' : '#E6F1FB'} />
            ))}
          </div>
        </Card>
      )}
    </div>
  );
}

// ─── Email & Compliance Tab ───────────────────────────────────────────────────
function EmailTab({ data }) {
  const { emailConfig } = data;
  return (
    <div style={{ display: 'grid', gap: 16 }}>
      <Card>
        <SectionTitle icon={Mail} title="Exchange Online security" sub="Email security configuration status" />
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(300px, 1fr))', gap: 10 }}>
          <SecurityIndicator label="External sender email tags"
            description={emailConfig?.externalSenderTags ? 'External sender banners are enabled on incoming mail' : 'External sender warnings are not enabled — users cannot identify external emails'}
            status={emailConfig?.externalSenderTags ? 'good' : 'bad'}
            icon={Mail} />
          <SecurityIndicator label="Microsoft Purview Message Encryption"
            description={emailConfig?.messageEncryption ? 'OME / Purview encryption is configured' : 'Message encryption not detected — sensitive emails may be unencrypted'}
            status={emailConfig?.messageEncryption ? 'good' : emailConfig?.messageEncryption === null ? 'unknown' : 'warning'}
            icon={Lock} />
          <SecurityIndicator label="Anti-phishing policy"
            description={emailConfig?.antiPhishing ? 'Custom anti-phishing policies configured' : 'Only default anti-phishing policy found — consider custom policies for key users'}
            status={emailConfig?.antiPhishing ? 'good' : 'warning'}
            icon={ShieldAlert} />
          <SecurityIndicator label="Safe links"
            description={emailConfig?.safeLinks ? 'Defender for Office 365 Safe Links is active' : 'Safe Links not detected — requires Defender for Office 365 Plan 1 or higher'}
            status={emailConfig?.safeLinks ? 'good' : emailConfig?.safeLinks === null ? 'unknown' : 'warning'}
            icon={Wifi} />
          <SecurityIndicator label="Safe attachments"
            description={emailConfig?.safeAttachments ? 'Defender for Office 365 Safe Attachments is active' : 'Safe Attachments not detected — malicious files may reach end users'}
            status={emailConfig?.safeAttachments ? 'good' : emailConfig?.safeAttachments === null ? 'unknown' : 'warning'}
            icon={HardDrive} />
          <SecurityIndicator label="DMARC / mail authentication"
            description={emailConfig?.dmarc ? 'DMARC policy detected for primary domain' : 'No DMARC policy detected — domain spoofing risk'}
            status={emailConfig?.dmarc ? 'good' : emailConfig?.dmarc === null ? 'unknown' : 'bad'}
            icon={ShieldCheck} />
        </div>
      </Card>

      <Card>
        <SectionTitle icon={Info} title="Exchange Online configuration note" sub="" />
        <div style={{ fontSize: 13, color: 'var(--color-text-secondary)', lineHeight: 1.7 }}>
          Full Exchange Online configuration (transport rules, DMARC, connector settings, mail flow) requires additional permissions
          via the Exchange Online PowerShell module or Exchange admin API, which is separate from Microsoft Graph. The indicators
          above reflect what can be determined from Graph API data and Defender for Office 365 policies.
          For a complete picture, review the{' '}
          <a href="https://admin.exchange.microsoft.com" target="_blank" rel="noreferrer"
            style={{ color: '#185FA5', textDecoration: 'none' }}>Exchange admin center</a>.
        </div>
      </Card>
    </div>
  );
}

// ─── Main Dashboard ───────────────────────────────────────────────────────────
export default function M365TenantDashboard() {
  const [msalInstance, setMsalInstance] = useState(null);
  const [account, setAccount] = useState(null);
  const [token, setToken] = useState(null);
  const [loading, setLoading] = useState(false);
  const [loadingStatus, setLoadingStatus] = useState('');
  const [error, setError] = useState(null);
  const [activeTab, setActiveTab] = useState('overview');
  const [dashData, setDashData] = useState(null);
  const [dateDays, setDateDays] = useState(7);

  // Init MSAL
 useEffect(() => {
  const instance = new PublicClientApplication(MSAL_CONFIG);
  instance.initialize().then(() => {
    instance.handleRedirectPromise().then((result) => {
      if (result?.account) setAccount(result.account);
      const accounts = instance.getAllAccounts();
      if (accounts.length > 0) setAccount(accounts[0]);
    }).catch(() => {});
    setMsalInstance(instance);
  });
}, []);

  // Acquire token silently or interactively
  const acquireToken = useCallback(async () => {
  if (!msalInstance || !account) return null;
  try {
    const result = await msalInstance.acquireTokenSilent({ scopes: SCOPES, account });
    return result.accessToken;
  } catch (e) {
    if (e instanceof InteractionRequiredAuthError) {
      await msalInstance.acquireTokenRedirect({ scopes: SCOPES, account });
    }
    throw e;
  }
}, [msalInstance, account]);

  // Login
  const handleLogin = useCallback(async () => {
  if (!msalInstance) return;
  try {
    await msalInstance.loginRedirect({ scopes: SCOPES });
  } catch (e) {
    setError('Login failed: ' + e.message);
  }
}, [msalInstance]);

  // Sign out
  const handleSignOut = useCallback(async () => {
    if (!msalInstance || !account) return;
    await msalInstance.logoutRedirect({ account });
    setAccount(null);
    setToken(null);
    setDashData(null);
  }, [msalInstance, account]);

  // Fetch all data
  const fetchAll = useCallback(async (tk, days = 7) => {
    const account = msalInstance?.getAllAccounts()?.[0];
    setLoading(true);
    setError(null);
    const data = {};

    try {
      setLoadingStatus('Loading organization info…');
      const orgRes = await graphGet(tk, '/organization');
      data.org = orgRes.value?.[0] || null;

      setLoadingStatus('Loading users…');
      const userList = await graphGetAll(tk, '/users?$select=id,displayName,userPrincipalName,userType,accountEnabled,lastPasswordChangeDateTime,signInActivity&$top=999');
      const guestCount = userList.filter(u => u.userType === 'Guest').length;
      const enabledCount = userList.filter(u => u.accountEnabled && u.userType !== 'Guest').length;
      data.users = {
        list: userList,
        total: userList.length,
        guests: guestCount,
        enabled: enabledCount,
        disabled: userList.filter(u => !u.accountEnabled).length,
        globalAdmins: 0,
      };

      setLoadingStatus('Loading admin roles…');
      try {
        const rolesRes = await graphGet(tk, '/directoryRoles?$filter=roleTemplateId eq \'62e90394-69f5-4237-9190-012177145e10\'');
        const gaRole = rolesRes.value?.[0];
        if (gaRole) {
          const membersRes = await graphGetAll(tk, `/directoryRoles/${gaRole.id}/members?$select=id,displayName,userPrincipalName`);
          data.globalAdminsList = membersRes;
          data.users.globalAdmins = membersRes.length;
        }
      } catch { data.globalAdminsList = []; }

      setLoadingStatus('Loading licenses…');
      try {
        const licRes = await graphGet(tk, '/subscribedSkus');
        data.licenses = licRes.value || [];
      } catch { data.licenses = []; }

      // Detect Entra ID tier
      const hasP2 = (data.licenses || []).some(l =>
        ['AAD_PREMIUM_P2','EMSPREMIUM','SPE_E5','ENTERPRISEPREMIUM'].includes(l.skuPartNumber) && l.consumedUnits > 0);
      const hasP1 = (data.licenses || []).some(l =>
        ['AAD_PREMIUM','EMS','SPE_E3','ENTERPRISEPACK'].includes(l.skuPartNumber) && l.consumedUnits > 0);
      data.entraTier = hasP2 ? 'Microsoft Entra ID P2' : hasP1 ? 'Microsoft Entra ID P1' : 'Microsoft Entra ID Free';

      setLoadingStatus('Loading Conditional Access policies…');
      try {
        // Refresh token before this call — it's deep in the fetch sequence
        let caTk = tk;
        try {
          const fresh = await msalInstance.acquireTokenSilent({ scopes: SCOPES, account });
          caTk = fresh.accessToken;
        } catch {}
        const caPolicies = await graphGetAll(caTk, '/identity/conditionalAccess/policies');
        data.caDetails = caPolicies;
        data.caStats = {
          total: caPolicies.length,
          enabled: caPolicies.filter(p => p.state === 'enabled').length,
          reportOnly: caPolicies.filter(p => p.state === 'enabledForReportingButNotEnforced').length,
          disabled: caPolicies.filter(p => p.state === 'disabled').length,
        };
      } catch { data.caDetails = []; data.caStats = { total: 0, enabled: 0 }; }

      setLoadingStatus('Loading named locations…');
      try {
        const nlRes = await graphGet(tk, '/identity/conditionalAccess/namedLocations');
        data.namedLocations = nlRes.value || [];
      } catch { data.namedLocations = []; }

      setLoadingStatus('Loading sign-in logs…');
      try {
        const since = new Date(Date.now() - days * 86400000).toISOString();
        const siRes = await graphGet(tk, `/auditLogs/signIns?$filter=createdDateTime ge ${since}&$top=500&$select=id,createdDateTime,userDisplayName,userPrincipalName,appDisplayName,clientAppUsed,ipAddress,location,deviceDetail,status,riskLevelAggregated,isInteractive`);
        const signInList = siRes.value || [];
        const countries = new Set(signInList.map(s => s.location?.countryOrRegion).filter(Boolean));
        const apps = new Set(signInList.map(s => s.appDisplayName).filter(Boolean));
        const failed = signInList.filter(s => s.status?.errorCode !== 0);
        const risky = signInList.filter(s => s.riskLevelAggregated && !['none', 'hidden', 'unknownFutureValue'].includes(s.riskLevelAggregated));
        data.signIns = {
          list: signInList,
          total: signInList.length,
          successful: signInList.length - failed.length,
          failed: failed.length,
          uniqueCountries: countries.size,
          uniqueApps: apps.size,
          risky: risky.length,
          failPct: signInList.length ? Math.round((failed.length / signInList.length) * 100) : 0,
        };
        data.signInStats = data.signIns;
        const riskMap = {};
        risky.forEach(s => {
          riskMap[s.riskLevelAggregated] = (riskMap[s.riskLevelAggregated] || 0) + 1;
        });
        data.riskSummary = Object.keys(riskMap).length > 0 ? riskMap : null;
      } catch { data.signIns = { list: [], total: 0, successful: 0, failed: 0, uniqueCountries: 0, uniqueApps: 0, risky: 0, failPct: 0 }; data.signInStats = data.signIns; }

      setLoadingStatus('Loading Intune data…');
      try {
        const [devicesRes, compPoliciesRes, configProfilesRes] = await Promise.all([
          graphGetAll(tk, '/deviceManagement/managedDevices?$select=id,deviceName,operatingSystem,complianceState,enrolledDateTime&$top=999'),
          graphGet(tk, '/deviceManagement/deviceCompliancePolicies'),
          graphGet(tk, '/deviceManagement/deviceConfigurations'),
        ]);
        const devices = devicesRes || [];
        const complianceByState = {};
        const byOS = {};
        devices.forEach(d => {
          complianceByState[d.complianceState || 'unknown'] = (complianceByState[d.complianceState || 'unknown'] || 0) + 1;
          const os = d.operatingSystem || 'Unknown';
          byOS[os] = (byOS[os] || 0) + 1;
        });
        const thirtyDaysAgo = new Date(Date.now() - 30 * 86400000);
        data.intune = {
          total: devices.length,
          compliant: complianceByState.compliant || 0,
          nonCompliant: complianceByState.noncompliant || 0,
          complianceByState,
          byOS,
          recentEnrollments: devices.filter(d => d.enrolledDateTime && new Date(d.enrolledDateTime) > thirtyDaysAgo).length,
          compliancePolicies: compPoliciesRes.value?.length || 0,
          compliancePolicyList: compPoliciesRes.value || [],
          configProfiles: configProfilesRes.value?.length || 0,
          configProfileList: configProfilesRes.value || [],
        };
        data.intuneStats = { total: devices.length, compliant: complianceByState.compliant || 0 };
      } catch { data.intune = { total: 0, compliant: 0, nonCompliant: 0, compliancePolicies: 0, configProfiles: 0, compliancePolicyList: [], configProfileList: [] }; data.intuneStats = { total: 0, compliant: 0 }; }

      setLoadingStatus('Loading auth methods…');
      try {
        const authRes = await graphGetBeta(tk, '/reports/authenticationMethods/usersRegisteredByMethod');
        const total = authRes.value?.length || 0;
        const mfaReg = authRes.value?.filter(u => u.isMfaRegistered)?.length || 0;
        data.authMethods = { mfaRegistered: mfaReg, total, mfaPct: total ? Math.round((mfaReg / total) * 100) : 0 };
      } catch { data.authMethods = null; }

      setLoadingStatus('Evaluating security posture…');
      const hasMFAPolicy = (data.caDetails || []).some(p =>
        p.state !== 'disabled' && (
          p.grantControls?.builtInControls?.includes('mfa') ||
          p.grantControls?.authenticationStrength != null ||
          p.grantControls?.builtInControls?.includes('compliantDevice') ||
          p.grantControls?.builtInControls?.includes('domainJoinedDevice')
        ));
      const hasLegacyBlock = (data.caDetails || []).some(p =>
        p.state !== 'disabled' && (
          p.conditions?.clientAppTypes?.includes('exchangeActiveSync') ||
          p.conditions?.clientAppTypes?.includes('other')
        ) && p.grantControls?.builtInControls?.includes('block'));
      const hasSecurityDefaults = data.org?.isSecurityDefaultsEnabled;
      data.securityIndicators = [
        {
          label: 'MFA policy active',
          description: hasMFAPolicy ? 'At least one enabled CA policy requires MFA' : 'No enforced CA policy requiring MFA detected',
          status: hasMFAPolicy ? 'good' : hasSecurityDefaults ? 'warning' : 'bad',
          icon: Fingerprint,
        },
        {
          label: 'Security defaults',
          description: hasSecurityDefaults ? 'Security defaults are enabled — basic MFA and legacy auth blocking in place' : 'Security defaults disabled — ensure CA policies cover equivalent protections',
          status: hasSecurityDefaults ? 'good' : 'warning',
          icon: Shield,
        },
        {
          label: 'Legacy auth blocked',
          description: hasLegacyBlock ? 'CA policy blocking legacy authentication clients detected' : 'No CA policy found blocking legacy auth — brute force risk may be elevated',
          status: hasLegacyBlock ? 'good' : 'bad',
          icon: WifiOff,
        },
        {
          label: 'Global Admin count',
          description: data.users.globalAdmins <= 3 ? `${data.users.globalAdmins} Global Admins — within recommended range` : `${data.users.globalAdmins} Global Admins — Microsoft recommends ≤ 3`,
          status: data.users.globalAdmins === 0 ? 'bad' : data.users.globalAdmins <= 3 ? 'good' : 'warning',
          icon: Crown,
        },
        {
          label: 'Guest user presence',
          description: guestCount === 0 ? 'No guest users in tenant' : `${guestCount} guest users — review access with Entra ID Access Reviews`,
          status: guestCount === 0 ? 'good' : guestCount > 50 ? 'warning' : 'good',
          icon: Globe,
        },
        {
          label: 'Intune device compliance',
          description: data.intune.total === 0 ? 'No Intune-enrolled devices found' : `${data.intune.compliant} of ${data.intune.total} devices compliant`,
          status: data.intune.total === 0 ? 'unknown' : data.intune.compliant / data.intune.total >= 0.9 ? 'good' : data.intune.compliant / data.intune.total >= 0.7 ? 'warning' : 'bad',
          icon: Monitor,
        },
      ];

      data.emailConfig = {
        externalSenderTags: null,
        messageEncryption: null,
        antiPhishing: null,
        safeLinks: null,
        safeAttachments: null,
        dmarc: null,
      };

      setLoadingStatus('Done!');
      setDashData(data);
    } catch (e) {
      setError('Error loading data: ' + e.message);
    } finally {
      setLoading(false);
    }
  }, []);

  // Fetch on login
  useEffect(() => {
    if (!account || !msalInstance) return;
    acquireToken().then(tk => {
      if (tk) { setToken(tk); fetchAll(tk, dateDays); }
    }).catch(e => setError(e.message));
  }, [account, msalInstance]);

  const tenantName = dashData?.org?.displayName || account?.tenantId;

  if (!msalInstance) return <div style={{ padding: 40, textAlign: 'center', color: 'var(--color-text-secondary)' }}>Initializing…</div>;
  if (!account) return <LoginPage onLogin={handleLogin} />;

  if (loading) return (
    <div style={{ minHeight: '100vh', display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', gap: 16 }}>
      <div style={{ width: 44, height: 44, borderRadius: 12, background: '#E6F1FB', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
        <RefreshCw size={22} color="#185FA5" style={{ animation: 'spin 1s linear infinite' }} />
      </div>
      <div style={{ fontSize: 15, fontWeight: 500, color: 'var(--color-text-primary)' }}>Loading tenant data…</div>
      <div style={{ fontSize: 13, color: 'var(--color-text-secondary)' }}>{loadingStatus}</div>
      <style>{`@keyframes spin { from { transform: rotate(0deg); } to { transform: rotate(360deg); } }`}</style>
    </div>
  );

  if (error) return (
    <div style={{ minHeight: '100vh', display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 20 }}>
      <Card style={{ maxWidth: 480, textAlign: 'center' }}>
        <AlertTriangle size={32} color="#E24B4A" style={{ margin: '0 auto 12px' }} />
        <div style={{ fontSize: 16, fontWeight: 500, marginBottom: 8, color: 'var(--color-text-primary)' }}>Error loading data</div>
        <div style={{ fontSize: 13, color: 'var(--color-text-secondary)', marginBottom: 16 }}>{error}</div>
        <div style={{ display: 'flex', gap: 8, justifyContent: 'center' }}>
          <button onClick={() => acquireToken().then(tk => tk && fetchAll(tk, dateDays))}
            style={{ padding: '8px 16px', borderRadius: 8, border: '0.5px solid #B5D4F4', background: '#185FA5',
              color: '#fff', fontSize: 13, cursor: 'pointer' }}>
            Retry
          </button>
          <button onClick={handleSignOut}
            style={{ padding: '8px 16px', borderRadius: 8, border: '0.5px solid var(--color-border-secondary)',
              background: 'transparent', color: 'var(--color-text-primary)', fontSize: 13, cursor: 'pointer' }}>
            Sign out
          </button>
        </div>
      </Card>
    </div>
  );

  if (!dashData) return null;

  const tabContent = {
    overview: <OverviewTab data={dashData} dateDays={dateDays} setDateDays={setDateDays}
                 setActiveTab={setActiveTab}
                 onRefresh={(d) => acquireToken().then(tk => tk && fetchAll(tk, d))} />,
    identity: <IdentityTab data={dashData} />,
    devices: <DevicesTab data={dashData} />,
    signins: <SignInsTab data={dashData} />,
    security: <SecurityTab data={dashData} />,
    email: <EmailTab data={dashData} />,
  };

  return (
    <div style={{ minHeight: '100vh', background: 'var(--color-background-tertiary)' }}>
      <style>{`
        @keyframes pulse { 0%, 100% { opacity: 1; } 50% { opacity: 0.5; } }
        @keyframes spin { from { transform: rotate(0deg); } to { transform: rotate(360deg); } }
        * { box-sizing: border-box; }
        button:focus { outline: 2px solid #378ADD; outline-offset: 2px; }
        input:focus { outline: 2px solid #378ADD; box-shadow: none; }
      `}</style>
      <Navbar
        userName={account?.name}
        userEmail={account?.username}
        tenantName={tenantName}
        onSignOut={handleSignOut}
        activeTab={activeTab}
        setActiveTab={setActiveTab}
      />
      <main style={{ maxWidth: 1400, margin: '0 auto', padding: '20px 20px 40px' }}>
        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 16 }}>
          <div>
            <h2 style={{ fontSize: 18, fontWeight: 500, margin: 0, color: 'var(--color-text-primary)' }}>
              {activeTab === 'overview' ? 'Tenant overview' :
               activeTab === 'identity' ? 'Identity & access' :
               activeTab === 'devices' ? 'Devices & Intune' :
               activeTab === 'signins' ? 'Sign-in activity' :
               activeTab === 'security' ? 'Security posture' : 'Email & compliance'}
            </h2>
            <p style={{ fontSize: 12, color: 'var(--color-text-secondary)', margin: '4px 0 0' }}>
              Live data from Microsoft Graph · {new Date().toLocaleString('en-US', { dateStyle: 'medium', timeStyle: 'short' })}
            </p>
          </div>
          <button onClick={() => acquireToken().then(tk => tk && fetchAll(tk, dateDays))}
            style={{ display: 'flex', alignItems: 'center', gap: 6, padding: '6px 12px', borderRadius: 7,
              border: '0.5px solid var(--color-border-secondary)', background: 'var(--color-background-primary)',
              color: 'var(--color-text-primary)', fontSize: 13, cursor: 'pointer' }}>
            <RefreshCw size={13} />
            Refresh
          </button>
        </div>
        {tabContent[activeTab]}
      </main>
    </div>
  );
}
