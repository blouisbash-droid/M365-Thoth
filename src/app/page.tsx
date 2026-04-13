import { getServerSession } from 'next-auth';
import { authOptions } from '@/lib/auth';
import { redirect } from 'next/navigation';
import LoginPage from '@/components/LoginPage';

export default async function Home() {
  const session = await getServerSession(authOptions);

  if (session && !session.error) {
    redirect('/dashboard');
  }

  return <LoginPage tokenError={session?.error} />;
}
