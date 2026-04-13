// ─── MS Graph Sign-In Log Types ──────────────────────────────────────────────

export interface DeviceDetail {
  deviceId: string | null;
  displayName: string | null;
  operatingSystem: string | null;
  browser: string | null;
  isCompliant: boolean | null;
  isManaged: boolean | null;
  /**
   * AzureAD     = Entra joined (formerly Azure AD joined)
   * ServerAD    = Hybrid Entra joined (on-prem AD + Entra)
   * Workplace   = Azure AD registered (includes Intune-enrolled personal devices)
   * ""  / null  = unregistered / browser sign-in
   */
  trustType: string | null;
}

export interface SignInLocation {
  city: string | null;
  state: string | null;
  countryOrRegion: string | null;
}

export interface SignInStatus {
  errorCode: number;
  failureReason: string | null;
  additionalDetails: string | null;
}

/**
 * passes  – device is Entra joined, Hybrid Entra joined, or Intune enrolled
 * fails   – device has partial info but doesn't meet any policy criterion
 * unknown – no device detail available (browser / unknown client)
 */
export type PolicyStatus = 'passes' | 'fails' | 'unknown';

export type DeviceCategory =
  | 'entra-joined'        // trustType === 'AzureAD'
  | 'hybrid-entra-joined' // trustType === 'ServerAD'
  | 'intune-enrolled'     // isManaged === true (Workplace or joined)
  | 'registered-only'     // Workplace, isManaged === false
  | 'no-device';          // no meaningful device detail

export interface SignInLog {
  id: string;
  createdDateTime: string;
  userDisplayName: string;
  userPrincipalName: string;
  appDisplayName: string;
  clientAppUsed: string;
  ipAddress: string;
  location: SignInLocation;
  deviceDetail: DeviceDetail;
  status: SignInStatus;
  conditionalAccessStatus: string;
  riskLevelAggregated: string;
  isInteractive: boolean;
  // Computed fields added on the client
  policyStatus: PolicyStatus;
  deviceCategory: DeviceCategory;
}

// ─── Dashboard Filter State ───────────────────────────────────────────────────

export type DateRange = '1d' | '7d' | '14d' | '30d';

export interface FilterState {
  search: string;
  userFilter: string;
  appFilter: string;
  osFilter: string;
  policyStatusFilter: string;
  signInStatusFilter: string;
  dateRange: DateRange;
}

// ─── Aggregated Stats ─────────────────────────────────────────────────────────

export interface DashboardStats {
  total: number;
  passes: number;
  fails: number;
  unknown: number;
  entraJoined: number;
  hybridEntraJoined: number;
  intuneEnrolled: number;
  registeredOnly: number;
  noDevice: number;
}

// ─── API Response ─────────────────────────────────────────────────────────────

export interface SignInsApiResponse {
  value: SignInLog[];
  nextLink: string | null;
}

// ─── MSP Multi-Tenant Config ──────────────────────────────────────────────────

/**
 * A customer tenant registered with the MSP dashboard.
 * No client secret is stored here — the MSP's own app credentials are used
 * to acquire app-only tokens for every connected tenant (client_credentials flow).
 */
export interface TenantConfig {
  /** Internal UUID, generated on creation */
  id: string;
  /** Human-readable name shown in the dashboard (e.g. "Contoso Corp") */
  name: string;
  /** The customer's Azure AD / Entra tenant ID (GUID) */
  tenantId: string;
  /** Optional free-text notes (contact, contract, etc.) */
  notes?: string;
  /** ISO date this tenant was added */
  createdAt: string;
}
