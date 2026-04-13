'use client';

import { useSession, signOut } from 'next-auth/react';
import { Shield, LogOut, ChevronDown } from 'lucide-react';
import { useState, useRef, useEffect } from 'react';

export default function Navbar() {
  const { data: session } = useSession();
  const [open, setOpen] = useState(false);
  const ref = useRef<HTMLDivElement>(null);

  // Close dropdown on outside click
  useEffect(() => {
    function handleClick(e: MouseEvent) {
      if (ref.current && !ref.current.contains(e.target as Node)) setOpen(false);
    }
    document.addEventListener('mousedown', handleClick);
    return () => document.removeEventListener('mousedown', handleClick);
  }, []);

  return (
    <header className="sticky top-0 z-30 bg-white border-b border-slate-200 shadow-sm">
      <div className="max-w-screen-2xl mx-auto px-4 h-14 flex items-center justify-between">
        {/* Logo */}
        <div className="flex items-center gap-2.5">
          <div className="flex items-center justify-center w-8 h-8 rounded-lg bg-blue-600">
            <Shield className="w-4 h-4 text-white" />
          </div>
          <div className="leading-tight">
            <span className="font-semibold text-slate-900 text-sm">M365</span>
            <span className="text-slate-400 text-sm"> / </span>
            <span className="font-semibold text-slate-900 text-sm">Conditional Access</span>
          </div>
        </div>

        {/* User menu */}
        {session?.user && (
          <div className="relative" ref={ref}>
            <button
              onClick={() => setOpen(o => !o)}
              className="flex items-center gap-2 px-3 py-1.5 rounded-lg text-sm
                         text-slate-700 hover:bg-slate-100 transition-colors"
            >
              {/* Avatar initials */}
              <span className="w-6 h-6 rounded-full bg-blue-100 text-blue-700 text-xs
                               font-semibold flex items-center justify-center">
                {(session.user.name ?? session.user.email ?? '?')
                  .split(' ')
                  .map(p => p[0])
                  .slice(0, 2)
                  .join('')
                  .toUpperCase()}
              </span>
              <span className="hidden sm:block max-w-[180px] truncate">
                {session.user.name ?? session.user.email}
              </span>
              <ChevronDown className="w-3.5 h-3.5 text-slate-400" />
            </button>

            {open && (
              <div className="absolute right-0 mt-1 w-48 bg-white rounded-lg shadow-lg border
                              border-slate-200 py-1 z-50">
                <div className="px-3 py-2 border-b border-slate-100">
                  <p className="text-xs text-slate-500 truncate">{session.user.email}</p>
                </div>
                <button
                  onClick={() => signOut({ callbackUrl: '/' })}
                  className="w-full flex items-center gap-2 px-3 py-2 text-sm text-slate-700
                             hover:bg-slate-50 transition-colors"
                >
                  <LogOut className="w-4 h-4 text-slate-400" />
                  Sign out
                </button>
              </div>
            )}
          </div>
        )}
      </div>
    </header>
  );
}
