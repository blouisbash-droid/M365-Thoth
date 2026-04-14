import type { Configuration } from '@azure/msal-browser';
import { LogLevel } from '@azure/msal-browser';
import { AZURE_CLIENT_ID, AZURE_AUTHORITY } from '@/config';

export const GRAPH_SCOPES = ['AuditLog.Read.All'];

export const MSAL_CONFIG: Configuration = {
  auth: {
    clientId: AZURE_CLIENT_ID,
    authority: AZURE_AUTHORITY,
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
