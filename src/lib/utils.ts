import { type ClassValue, clsx } from 'clsx';
import { twMerge } from 'tailwind-merge';
import { DeviceDetail, DeviceCategory, PolicyStatus, SignInLog, DashboardStats } from './types';

// ─── Tailwind class helper ────────────────────────────────────────────────────

export function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}

// ─── Device categorisation ────────────────────────────────────────────────────

/**
 * Maps raw MS Graph deviceDetail to a human-readable category.
 *
 * Priority order:
 *  1. AzureAD  trust type → entra-joined
 *  2. ServerAD trust type → hybrid-entra-joined
 *  3. isManaged === true  → intune-enrolled (covers Workplace + joined devices managed by Intune)
 *  4. Workplace + not managed → registered-only
 *  5. No useful detail → no-device
 */
export function getDeviceCategory(deviceDetail: DeviceDetail | null): DeviceCategory {
  if (!deviceDetail || !deviceDetail.deviceId) return 'no-device';

  if (deviceDetail.trustType === 'AzureAD') return 'entra-joined';
  if (deviceDetail.trustType === 'ServerAD') return 'hybrid-entra-joined';
  if (deviceDetail.isManaged === true) return 'intune-enrolled';
  if (deviceDetail.trustType === 'Workplace') return 'registered-only';

  return 'no-device';
}

/**
 * Returns whether the sign-in would pass the target CA policy:
 *  "Device must be Entra joined, Hybrid Entra joined, OR enrolled in Intune."
 */
export function getPolicyStatus(deviceDetail: DeviceDetail | null): PolicyStatus {
  if (!deviceDetail || !deviceDetail.deviceId) return 'fails';

  const isEntraJoined = deviceDetail.trustType === 'AzureAD';
  const isHybridJoined = deviceDetail.trustType === 'ServerAD';
  const isIntuneEnrolled = deviceDetail.isManaged === true;

  if (isEntraJoined || isHybridJoined || isIntuneEnrolled) return 'passes';

  // Has a device ID but meets none of the criteria (e.g. Workplace, isManaged=false)
  return 'fails';
}

// ─── Display helpers ──────────────────────────────────────────────────────────

export function formatTrustType(trustType: string | null): string {
  switch (trustType) {
    case 'AzureAD':   return 'Entra Joined';
    case 'ServerAD':  return 'Hybrid Entra Joined';
    case 'Workplace': return 'Registered (Workplace)';
    default:          return 'None';
  }
}

/**
 * Normalise a raw OS string (e.g. "Windows 10", "iOS 17.4") to a short label.
 */
export function getOSLabel(os: string | null): string {
  if (!os) return 'Unknown';
  const lower = os.toLowerCase();
  if (lower.includes('windows')) return 'Windows';
  if (lower.includes('macos') || lower.includes('mac os')) return 'macOS';
  if (lower.includes('ios')) return 'iOS';
  if (lower.includes('android')) return 'Android';
  if (lower.includes('linux')) return 'Linux';
  return 'Other';
}

export function formatDate(dateStr: string): string {
  return new Date(dateStr).toLocaleString('en-US', {
    month: 'short',
    day: 'numeric',
    year: 'numeric',
    hour: '2-digit',
    minute: '2-digit',
  });
}

export function formatDateShort(dateStr: string): string {
  return new Date(dateStr).toLocaleDateString('en-US', {
    month: 'short',
    day: 'numeric',
  });
}

export function policyStatusLabel(status: PolicyStatus): string {
  switch (status) {
    case 'passes':  return 'Passes Policy';
    case 'fails':   return 'Fails Policy';
    case 'unknown': return 'No Device Info';
  }
}

export function deviceCategoryLabel(cat: DeviceCategory): string {
  switch (cat) {
    case 'entra-joined':        return 'Entra Joined';
    case 'hybrid-entra-joined': return 'Hybrid Entra Joined';
    case 'intune-enrolled':     return 'Intune Enrolled';
    case 'registered-only':     return 'Registered Only';
    case 'no-device':           return 'No Device';
  }
}

// ─── Sign-in processing ───────────────────────────────────────────────────────

// eslint-disable-next-line @typescript-eslint/no-explicit-any
export function processSignIn(raw: any): SignInLog {
  const deviceDetail: DeviceDetail = raw.deviceDetail ?? {
    deviceId: null,
    displayName: null,
    operatingSystem: null,
    browser: null,
    isCompliant: null,
    isManaged: null,
    trustType: null,
  };

  return {
    id: raw.id,
    createdDateTime: raw.createdDateTime,
    userDisplayName: raw.userDisplayName ?? '',
    userPrincipalName: raw.userPrincipalName ?? '',
    appDisplayName: raw.appDisplayName ?? '',
    clientAppUsed: raw.clientAppUsed ?? '',
    ipAddress: raw.ipAddress ?? '',
    location: raw.location ?? { city: null, state: null, countryOrRegion: null },
    deviceDetail,
    status: raw.status ?? { errorCode: 0, failureReason: null, additionalDetails: null },
    conditionalAccessStatus: raw.conditionalAccessStatus ?? '',
    riskLevelAggregated: raw.riskLevelAggregated ?? 'none',
    isInteractive: raw.isInteractive ?? true,
    policyStatus: getPolicyStatus(deviceDetail),
    deviceCategory: getDeviceCategory(deviceDetail),
  };
}

// ─── Aggregation ──────────────────────────────────────────────────────────────

export function computeStats(signIns: SignInLog[]): DashboardStats {
  const total = signIns.length;
  const passes = signIns.filter(s => s.policyStatus === 'passes').length;
  const fails = signIns.filter(s => s.policyStatus === 'fails').length;
  const unknown = signIns.filter(s => s.policyStatus === 'unknown').length;

  return {
    total,
    passes,
    fails,
    unknown,
    entraJoined: signIns.filter(s => s.deviceCategory === 'entra-joined').length,
    hybridEntraJoined: signIns.filter(s => s.deviceCategory === 'hybrid-entra-joined').length,
    intuneEnrolled: signIns.filter(s => s.deviceCategory === 'intune-enrolled').length,
    registeredOnly: signIns.filter(s => s.deviceCategory === 'registered-only').length,
    noDevice: signIns.filter(s => s.deviceCategory === 'no-device').length,
  };
}

// ─── CSV import ───────────────────────────────────────────────────────────────

function parseCSVLine(line: string): string[] {
  const result: string[] = [];
  let current = '';
  let insideQuotes = false;

  for (let i = 0; i < line.length; i++) {
    const char = line[i];
    const nextChar = line[i + 1];

    if (char === '"') {
      if (insideQuotes && nextChar === '"') {
        current += '"';
        i++;
      } else {
        insideQuotes = !insideQuotes;
      }
    } else if (char === ',' && !insideQuotes) {
      result.push(current.trim());
      current = '';
    } else {
      current += char;
    }
  }

  result.push(current.trim());
  return result;
}

export function parseCSV(csvContent: string): SignInLog[] {
  const lines = csvContent.split('\n').filter(l => l.trim());
  if (lines.length < 2) throw new Error('CSV file is empty or invalid');

  const headers = parseCSVLine(lines[0]);
  const expectedHeaders = [
    'DateTime', 'User', 'UPN', 'Application', 'Client App',
    'Device Name', 'OS', 'Trust Type', 'Intune Managed', 'Compliant',
    'Policy Status', 'Sign-in Status', 'Failure Reason',
    'CA Status', 'Risk Level', 'IP Address', 'City', 'Country',
  ];

  // Validate headers
  if (headers.length < expectedHeaders.length) {
    throw new Error('CSV format is invalid. Missing required columns.');
  }

  const headerMap: Record<string, number> = {};
  expectedHeaders.forEach((header, idx) => {
    headerMap[header] = idx;
  });

  const signIns: SignInLog[] = [];

  for (let i = 1; i < lines.length; i++) {
    const values = parseCSVLine(lines[i]);
    if (values.every(v => !v)) continue; // Skip empty lines

    try {
      const getString = (header: string) => values[headerMap[header]] || '';
      const getBoolean = (header: string) => {
        const val = values[headerMap[header]]?.toLowerCase();
        if (val === 'true') return true;
        if (val === 'false') return false;
        if (val === 'n/a' || val === '') return null;
        return null;
      };

      // Parse Trust Type back to original values
      const trustTypeLabel = getString('Trust Type');
      let trustType: string | null = null;
      if (trustTypeLabel === 'Entra Joined') trustType = 'AzureAD';
      else if (trustTypeLabel === 'Hybrid Entra Joined') trustType = 'ServerAD';
      else if (trustTypeLabel === 'Registered (Workplace)') trustType = 'Workplace';

      // Parse sign-in status
      const signInStatus = getString('Sign-in Status');
      const errorCode = signInStatus === 'Failure' ? 1 : 0;

      const deviceDetail: DeviceDetail = {
        deviceId: 'imported-' + Math.random().toString(36).substr(2, 9),
        displayName: getString('Device Name') || null,
        operatingSystem: getString('OS') || null,
        browser: null,
        isCompliant: getBoolean('Compliant'),
        isManaged: getBoolean('Intune Managed'),
        trustType,
      };

      const signInLog: SignInLog = {
        id: 'imported-' + Math.random().toString(36).substr(2, 9),
        createdDateTime: getString('DateTime'),
        userDisplayName: getString('User'),
        userPrincipalName: getString('UPN'),
        appDisplayName: getString('Application'),
        clientAppUsed: getString('Client App'),
        ipAddress: getString('IP Address'),
        location: {
          city: getString('City') ? getString('City') : null,
          state: null,
          countryOrRegion: getString('Country') ? getString('Country') : null,
        },
        deviceDetail,
        status: {
          errorCode,
          failureReason: getString('Failure Reason') ? getString('Failure Reason') : null,
          additionalDetails: null,
        },
        conditionalAccessStatus: getString('CA Status'),
        riskLevelAggregated: getString('Risk Level') || 'none',
        isInteractive: true,
        policyStatus: 'unknown',
        deviceCategory: 'no-device',
      };

      // Recompute policy status and device category
      signInLog.policyStatus = getPolicyStatus(signInLog.deviceDetail);
      signInLog.deviceCategory = getDeviceCategory(signInLog.deviceDetail);

      signIns.push(signInLog);
    } catch (err) {
      console.warn(`Failed to parse row ${i + 1}:`, err);
    }
  }

  if (signIns.length === 0) throw new Error('No valid sign-in records found in CSV');
  return signIns;
}

// ─── CSV export ───────────────────────────────────────────────────────────────

export function exportToCSV(signIns: SignInLog[], filename = 'signin-report.csv') {
  const headers = [
    'DateTime', 'User', 'UPN', 'Application', 'Client App',
    'Device Name', 'OS', 'Trust Type', 'Intune Managed', 'Compliant',
    'Policy Status', 'Sign-in Status', 'Failure Reason',
    'CA Status', 'Risk Level', 'IP Address', 'City', 'Country',
  ];

  const escape = (v: string | null | undefined | boolean | number) => {
    const s = String(v ?? '');
    return s.includes(',') || s.includes('"') || s.includes('\n')
      ? `"${s.replace(/"/g, '""')}"`
      : s;
  };

  const rows = signIns.map(s => [
    escape(formatDate(s.createdDateTime)),
    escape(s.userDisplayName),
    escape(s.userPrincipalName),
    escape(s.appDisplayName),
    escape(s.clientAppUsed),
    escape(s.deviceDetail.displayName),
    escape(s.deviceDetail.operatingSystem),
    escape(formatTrustType(s.deviceDetail.trustType)),
    escape(s.deviceDetail.isManaged ?? 'N/A'),
    escape(s.deviceDetail.isCompliant ?? 'N/A'),
    escape(policyStatusLabel(s.policyStatus)),
    escape(s.status.errorCode === 0 ? 'Success' : 'Failure'),
    escape(s.status.failureReason),
    escape(s.conditionalAccessStatus),
    escape(s.riskLevelAggregated),
    escape(s.ipAddress),
    escape(s.location.city),
    escape(s.location.countryOrRegion),
  ].join(','));

  const csv = [headers.join(','), ...rows].join('\n');
  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = filename;
  link.click();
  URL.revokeObjectURL(url);
}
