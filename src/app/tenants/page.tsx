import { getServerSession } from 'next-auth';
import { authOptions } from '@/lib/auth';
import { redirect } from 'next/navigation';
import Navbar from '@/components/Navbar';
import TenantManager from '@/components/TenantManager';

export const dynamic = 'force-dynamic';

export default async function TenantsPage() {
  const session = await getServerSession(authOptions);

  if (!session) redirect('/');
  if (session.error === 'RefreshAccessTokenError') redirect('/');

  // Pass server-side env vars as props so the client component can build consent URLs
  const appUrl    = process.env.NEXTAUTH_URL ?? 'https://m365-ca-review.vercel.app';
  const clientId  = process.env.AZURE_AD_CLIENT_ID ?? '';

  return (
    <div className="min-h-screen bg-slate-50">
      <Navbar />
      <TenantManager appUrl={appUrl} clientId={clientId} />
    </div>
  );
}
