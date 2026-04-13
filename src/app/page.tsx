import { getServerSession } from 'next-auth';
import { authOptions } from '@/lib/auth';
import { redirect } from 'next/navigation';
import LoginPage from '@/components/LoginPage';
import LandingPage from '@/components/LandingPage';

export const dynamic = 'force-dynamic';

export default async function Home() {
  const isConfigured = !!(
    process.env.AZURE_AD_CLIENT_ID &&
    process.env.AZURE_AD_CLIENT_SECRET &&
    process.env.AZURE_AD_TENANT_ID &&
    process.env.NEXTAUTH_SECRET
  );

  if (!isConfigured) {
    return <LandingPage />;
  }

  const session = await getServerSession(authOptions);

  if (session && !session.error) {
    redirect('/dashboard');
  }

  return <LoginPage tokenError={session?.error} />;
}
