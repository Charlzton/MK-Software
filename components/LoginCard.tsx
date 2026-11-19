'use client';

import React, { useState } from 'react';
import { User } from '@/lib/types';
import type { SheetState } from '@/components/ExcelSpreadsheet';
import { Stethoscope, Users, ShieldAlert, Eye, EyeOff } from 'lucide-react';

export interface LoginPreloadPayload {
  workspace?: 'reception' | 'doctor';
  sheets?: Record<string, SheetState>;
}

interface LoginCardProps {
  onSuccess: (user: User, preload?: LoginPreloadPayload) => void;
  externalError?: string | null;
}

export function LoginCard({ onSuccess, externalError }: LoginCardProps) {
  const [loginMode, setLoginMode] = useState<'staff' | 'admin'>('staff');
  const [staffRole, setStaffRole] = useState<'doctor' | 'reception'>('reception');
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [showPassword, setShowPassword] = useState(false);
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    const trimmedEmail = email.trim();
    if (!trimmedEmail || !password) {
      setError('Please enter both email and password.');
      return;
    }

    setIsLoading(true);
    setError(null);

    const targetRole = loginMode === 'admin' ? 'admin' : staffRole;

    try {
      const res = await fetch('/api/auth/login', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ email: trimmedEmail, password, role: targetRole })
      });

      const data = await res.json();

      if (!res.ok) {
        setError(data.error || 'Authentication failed');
        return;
      }

      if (data.user) {
        const preload: LoginPreloadPayload | undefined =
          data.initialWorkspace && data.initialWorkbook?.sheets
            ? {
                workspace: data.initialWorkspace,
                sheets: data.initialWorkbook.sheets
              }
            : undefined;
        onSuccess(data.user, preload);
      }
    } catch {
      setError('Unable to reach server. Please check network.');
    } finally {
      setIsLoading(false);
    }
  };

  const displayedError = error || externalError;

  return (
    <div className="w-full max-w-[min(100%,24rem)] sm:max-w-[27rem] lg:max-w-[28rem] mx-auto p-5 sm:p-7 md:p-8 bg-white/20 dark:bg-slate-900/40 backdrop-blur-2xl border border-white/40 dark:border-white/10 shadow-[0_16px_48px_rgba(0,0,0,0.35)] rounded-2xl text-slate-900 dark:text-slate-100 relative overflow-hidden">
      {/* Glossy top sheen highlight */}
      <div className="pointer-events-none absolute inset-x-0 top-0 h-[1px] bg-gradient-to-r from-transparent via-white/80 to-transparent" />

      <h1 className="text-xl sm:text-2xl font-bold text-slate-900 dark:text-white text-center tracking-tight mb-4 sm:mb-5">
        {loginMode === 'admin' ? 'Admin Login' : 'Login'}
      </h1>

      <form onSubmit={handleSubmit} className="space-y-3.5 sm:space-y-4">
        {displayedError && (
          <div className="p-2.5 sm:p-3 text-xs text-red-700 bg-red-50/90 border border-red-200 rounded-xl flex items-start gap-2 shadow-xs">
            <ShieldAlert className="w-4 h-4 text-red-600 shrink-0 mt-0.5" />
            <span className="leading-relaxed font-medium break-words">{displayedError}</span>
          </div>
        )}

        {/* Staff Role Selector: Compact horizontal pill-shaped segmented control with sliding active indicator */}
        {loginMode === 'staff' && (
          <div className="relative flex items-center bg-slate-900/10 p-1 rounded-full border border-white/70 shadow-inner w-full max-w-[228px] mx-auto">
            {/* Sliding active pill indicator */}
            <div
              className={`absolute top-1 bottom-1 w-[calc(50%-4px)] rounded-full bg-blue-600 shadow-sm transition-transform duration-200 ease-out pointer-events-none ${
                staffRole === 'doctor' ? 'translate-x-0' : 'translate-x-[calc(100%+0px)]'
              }`}
            />

            <button
              type="button"
              onClick={() => setStaffRole('doctor')}
              className={`relative z-10 flex-1 min-h-[34px] sm:min-h-[30px] py-1 px-2 text-xs font-semibold rounded-full flex items-center justify-center gap-1 transition-colors duration-150 cursor-pointer touch-manipulation ${
                staffRole === 'doctor'
                  ? 'text-white'
                  : 'text-slate-800 hover:text-slate-950'
              }`}
            >
              <Stethoscope className="w-3.5 h-3.5 shrink-0" />
              <span>Doctor</span>
            </button>

            <button
              type="button"
              onClick={() => setStaffRole('reception')}
              className={`relative z-10 flex-1 min-h-[34px] sm:min-h-[30px] py-1 px-2 text-xs font-semibold rounded-full flex items-center justify-center gap-1 transition-colors duration-150 cursor-pointer touch-manipulation ${
                staffRole === 'reception'
                  ? 'text-white'
                  : 'text-slate-800 hover:text-slate-950'
              }`}
            >
              <Users className="w-3.5 h-3.5 shrink-0" />
              <span>Reception</span>
            </button>
          </div>
        )}

        {/* Email Address Field */}
        <div>
          <label htmlFor="email" className="block text-xs font-semibold text-slate-800 mb-1.5">
            Email address
          </label>
          <input
            id="email"
            type="email"
            value={email}
            onChange={(e) => setEmail(e.target.value)}
            disabled={isLoading}
            autoComplete="email"
            required
            className="w-full min-h-[44px] sm:min-h-[38px] px-3.5 py-2 text-base sm:text-xs bg-white/75 border border-white/80 rounded-xl text-slate-900 placeholder:text-slate-400 focus:bg-white focus:outline-hidden focus:border-blue-600 focus:ring-1 focus:ring-blue-600 shadow-xs transition"
          />
        </div>

        {/* Password Field */}
        <div>
          <label htmlFor="password" className="block text-xs font-semibold text-slate-800 mb-1.5">
            Password
          </label>
          <div className="relative">
            <input
              id="password"
              type={showPassword ? 'text' : 'password'}
              value={password}
              onChange={(e) => setPassword(e.target.value)}
              disabled={isLoading}
              autoComplete="current-password"
              required
              className="w-full min-h-[44px] sm:min-h-[38px] pl-3.5 pr-11 py-2 text-base sm:text-xs bg-white/75 border border-white/80 rounded-xl text-slate-900 placeholder:text-slate-400 focus:bg-white focus:outline-hidden focus:border-blue-600 focus:ring-1 focus:ring-blue-600 shadow-xs transition"
            />
            <button
              type="button"
              onClick={() => setShowPassword((prev) => !prev)}
              className="absolute right-1.5 top-1/2 -translate-y-1/2 text-slate-600 hover:text-slate-900 cursor-pointer min-w-[36px] min-h-[36px] flex items-center justify-center rounded-lg transition touch-manipulation"
              title={showPassword ? 'Hide password' : 'Show password'}
              aria-label={showPassword ? 'Hide password' : 'Show password'}
            >
              {showPassword ? <EyeOff className="w-4 h-4" /> : <Eye className="w-4 h-4" />}
            </button>
          </div>
        </div>

        {/* Main Login Button */}
        <button
          type="submit"
          disabled={isLoading}
          className="w-full min-h-[44px] sm:min-h-[40px] py-2.5 px-4 text-xs font-bold uppercase tracking-wider text-white bg-blue-600 hover:bg-blue-700 active:bg-blue-800 rounded-xl transition disabled:opacity-50 cursor-pointer shadow-md mt-2 flex items-center justify-center touch-manipulation"
        >
          {isLoading ? 'Logging in...' : 'Login'}
        </button>

        {/* Understated Text Link with High Readability Contrast */}
        <div className="text-center pt-1 sm:pt-2">
          {loginMode === 'staff' ? (
            <button
              type="button"
              onClick={() => {
                setLoginMode('admin');
                setError(null);
              }}
              className="text-xs text-slate-900 font-semibold hover:text-blue-700 hover:underline transition cursor-pointer tracking-wide py-1 px-2 touch-manipulation"
            >
              Login as Admin
            </button>
          ) : (
            <button
              type="button"
              onClick={() => {
                setLoginMode('staff');
                setError(null);
              }}
              className="text-xs text-slate-900 font-semibold hover:text-blue-700 hover:underline transition cursor-pointer tracking-wide py-1 px-2 touch-manipulation"
            >
              Back to Staff Login
            </button>
          )}
        </div>
      </form>
    </div>
  );
}
