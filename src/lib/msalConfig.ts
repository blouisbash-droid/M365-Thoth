import type { Configuration } from '@azure/msal-browser';
import { LogLevel } from '@azure/msal-browser';

export const GRAPH_SCOPES = ['AuditLog.Read.All'];
export const CONFIG_STORAGE_KEY = 'm365_dashboard_config';

export interface AppConfig {
  tenantId: string;
  clientId: string;
}

export function getStoredConfig(): AppConfig | null {
  if (typeof window === 'undefined') return null;
  const raw = localStorage.getItem(CONFIG_STORAGE_KEY);
  if (!raw) return null;
  try {
    const parsed = JSON.parse(raw) as Partial<AppConfig>;
    if (!parsed.tenantId || !parsed.clientId) return null;
    return { tenantId: parsed.tenantId, clientId: parsed.clientId };
  } catch {
    return null;
  }
}

export function saveConfig(config: AppConfig): void {
  localStorage.setItem(CONFIG_STORAGE_KEY, JSON.stringify(config));
}

export function clearConfig(): void {
  localStorage.removeItem(CONFIG_STORAGE_KEY);
}

export function buildMsalConfig(config: AppConfig): Configuration {
  return {
    auth: {
      clientId: config.clientId,
      authority: `https://login.microsoftonline.com/${config.tenantId}`,
      redirectUri: typeof window !== 'undefined' ? window.location.origin : '/',
    },
    cache: {
      cacheLocation: 'localStorage',
    },
    system: {
      loggerOptions: {
        loggerCallback: (level: LogLevel, message: string, containsPii: boolean) => {
          if (containsPii) return;
          if (level === LogLevel.Error) console.error('[MSAL]', message);
        },
      },
    },
  };
}
