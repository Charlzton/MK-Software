'use client';

import React from 'react';
import { Eye, LogOut, Settings as SettingsIcon, ShieldCheck } from 'lucide-react';
import { User } from '@/lib/types';

interface NavbarProps {
  user: User;
  activeTab: 'reception' | 'doctor';
  onSelectTab: (tab: 'reception' | 'doctor') => void;
  onOpenSettings: () => void;
  onLogout: () => void;
}

export function Navbar({
  user,
  activeTab,
  onSelectTab,
  onOpenSettings,
  onLogout
}: NavbarProps) {
  // Check role authorization
  const canAccessReception = user.role === 'reception' || user.role === 'admin';
  const canAccessDoctor = user.role === 'doctor' || user.role === 'admin';

  return (
    <header className="sticky top-0 z-40 bg-white border-b border-slate-200">
      <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
        <div className="flex items-center justify-between gap-8 h-16">
          {/* Zone 1: Brand Wordmark */}
          <div className="flex items-center gap-2.5 shrink-0">
            <div className="w-8 h-8 rounded-md bg-slate-900 flex items-center justify-center text-teal-400 shrink-0">
              <Eye className="w-4 h-4 stroke-[2.2]" />
            </div>
            <div className="flex flex-col">
              <span className="text-base font-bold tracking-tight text-slate-900 whitespace-nowrap">
                MK Eye Center
              </span>
              <span className="text-[10px] uppercase tracking-wider text-slate-600 font-medium whitespace-nowrap">
                Clinical Patient Portal
              </span>
            </div>
          </div>

          {/* Zone 2: Navigation Links (single-line, clean unboxed with subtle underline/highlight) */}
          <nav className="flex items-center gap-1 sm:gap-2">
            {canAccessReception && (
              <button
                type="button"
                onClick={() => onSelectTab('reception')}
                className={`px-3.5 py-1.5 text-xs font-medium rounded-md transition-colors whitespace-nowrap shrink-0 ${
                  activeTab === 'reception'
                    ? 'bg-slate-900 text-white'
                    : 'text-slate-600 hover:text-slate-900 hover:bg-slate-100'
                }`}
              >
                Reception Workspace
              </button>
            )}

            {canAccessDoctor && (
              <button
                type="button"
                onClick={() => onSelectTab('doctor')}
                className={`px-3.5 py-1.5 text-xs font-medium rounded-md transition-colors whitespace-nowrap shrink-0 ${
                  activeTab === 'doctor'
                    ? 'bg-slate-900 text-white'
                    : 'text-slate-600 hover:text-slate-900 hover:bg-slate-100'
                }`}
              >
                Doctor Workspace
              </button>
            )}

            <button
              type="button"
              onClick={onOpenSettings}
              className="p-1.5 text-slate-500 hover:text-slate-900 hover:bg-slate-100 rounded-md transition-colors whitespace-nowrap shrink-0"
              title="Integration Settings & Logs"
            >
              <SettingsIcon className="w-4 h-4" />
            </button>
          </nav>

          {/* Zone 3: Account Profile & Sign Out */}
          <div className="flex items-center gap-4 shrink-0">
            <div className="hidden sm:flex items-center gap-2 text-xs text-slate-600">
              <span className="font-semibold text-slate-800">{user.name}</span>
              <span className="text-slate-400" aria-hidden="true">·</span>
              <span className="capitalize text-slate-500 flex items-center gap-1">
                <ShieldCheck className="w-3.5 h-3.5 text-teal-600" />
                {user.role}
              </span>
            </div>

            <button
              type="button"
              onClick={onLogout}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium text-slate-700 bg-slate-100 hover:bg-slate-200 rounded-md transition-colors whitespace-nowrap shrink-0"
              title="End session"
            >
              <LogOut className="w-3.5 h-3.5 text-slate-500" />
              <span>Log Out</span>
            </button>
          </div>
        </div>
      </div>
    </header>
  );
}
