'use client';
import dynamic from 'next/dynamic';

const M365TenantDashboard = dynamic(
  () => import('@/components/M365TenantDashboard'),
  { ssr: false }
);

export default function DashboardPage() {
  return <M365TenantDashboard />;
}