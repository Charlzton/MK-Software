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
    <div
      className="w-full max-w-[min(100%,24.5rem)] sm:max-w-[26.5rem] mx-auto px-6 py-7 sm:px-9 sm:py-9 rounded-[28px] sm:rounded-[30px] relative overflow-hidden
        bg-gradient-to-b from-white/38 via-white/24 to-white/16
        dark:from-slate-900/60 dark:via-slate-900/46 dark:to-slate-950/55
        backdrop-blur-xl backdrop-saturate-[1.35]
        border border-white/65 dark:border-white/18
        shadow-[0_24px_60px_-14px_rgba(15,23,42,0.22),0_4px_16px_-2px_rgba(15,23,42,0.08),inset_0_1px_1px_rgba(255,255,255,0.88),inset_0_-1px_1px_rgba(255,255,255,0.22)]
        dark:shadow-[0_24px_60px_-14px_rgba(0,0,0,0.55),inset_0_1px_1px_rgba(255,255,255,0.22),inset_0_-1px_1px_rgba(255,255,255,0.06)]
        text-slate-900 dark:text-slate-100
        transition-[box-shadow,border-color] duration-300 motion-reduce:transition-none"
    >
      {/* Delicate top specular edge highlight & subtle ambient glass reflections */}
      <div
        aria-hidden="true"
        className="pointer-events-none absolute inset-x-7 top-0 h-[1px] bg-gradient-to-r from-transparent via-white/95 dark:via-white/45 to-transparent"
      />
      <div
        aria-hidden="true"
        className="pointer-events-none absolute -top-24 -left-24 w-56 h-56 rounded-full bg-white/30 dark:bg-white/[0.06] blur-2xl"
      />
      <div
        aria-hidden="true"
        className="pointer-events-none absolute -bottom-24 -right-20 w-52 h-52 rounded-full bg-white/15 dark:bg-sky-400/[0.05] blur-2xl"
      />

      <h1 className="relative z-10 text-[22px] sm:text-[24px] font-semibold sm:font-bold text-slate-900 dark:text-white text-center tracking-tight mb-5 sm:mb-6">
        {loginMode === 'admin' ? 'Admin Login' : 'Login'}
      </h1>

      <form onSubmit={handleSubmit} className="relative z-10 space-y-4">
        {displayedError && (
          <div className="p-3 text-xs text-red-700 dark:text-red-200 bg-red-50/85 dark:bg-red-950/70 backdrop-blur-md border border-red-200/85 dark:border-red-500/35 rounded-2xl flex items-start gap-2 shadow-xs">
            <ShieldAlert className="w-4 h-4 text-red-600 dark:text-red-400 shrink-0 mt-0.5" />
            <span className="leading-relaxed font-medium break-words">{displayedError}</span>
          </div>
        )}

        {/* Staff Role Selector: Compact frosted pill-shaped segmented control with sliding active indicator */}
        {loginMode === 'staff' && (
          <div className="relative flex items-center bg-slate-900/[0.07] dark:bg-white/10 backdrop-blur-md p-1 rounded-full border border-white/75 dark:border-white/15 shadow-[inset_0_1px_3px_rgba(15,23,42,0.08)] w-full max-w-[236px] mx-auto">
            {/* Sliding active pill indicator */}
            <div
              className={`absolute top-1 bottom-1 w-[calc(50%-4px)] rounded-full bg-gradient-to-b from-blue-500 to-blue-600 shadow-[0_2px_8px_rgba(37,99,235,0.35),inset_0_1px_0_rgba(255,255,255,0.35)] transition-transform duration-200 ease-out motion-reduce:transition-none pointer-events-none ${
                staffRole === 'doctor' ? 'translate-x-0' : 'translate-x-[calc(100%+0px)]'
              }`}
            />

            <button
              type="button"
              onClick={() => setStaffRole('doctor')}
              className={`relative z-10 flex-1 min-h-[34px] sm:min-h-[30px] py-1 px-2.5 text-xs font-semibold rounded-full flex items-center justify-center gap-1.5 transition-colors duration-150 motion-reduce:transition-none cursor-pointer touch-manipulation ${
                staffRole === 'doctor'
                  ? 'text-white'
                  : 'text-slate-700 hover:text-slate-950 dark:text-slate-200 dark:hover:text-white'
              }`}
            >
              <Stethoscope className="w-3.5 h-3.5 shrink-0" />
              <span>Doctor</span>
            </button>

            <button
              type="button"
              onClick={() => setStaffRole('reception')}
              className={`relative z-10 flex-1 min-h-[34px] sm:min-h-[30px] py-1 px-2.5 text-xs font-semibold rounded-full flex items-center justify-center gap-1.5 transition-colors duration-150 motion-reduce:transition-none cursor-pointer touch-manipulation ${
                staffRole === 'reception'
                  ? 'text-white'
                  : 'text-slate-700 hover:text-slate-950 dark:text-slate-200 dark:hover:text-white'
              }`}
            >
              <Users className="w-3.5 h-3.5 shrink-0" />
              <span>Reception</span>
            </button>
          </div>
        )}

        {/* Email Address Field */}
        <div>
          <label
            htmlFor="email"
            className="block pl-3.5 text-[11.5px] font-semibold tracking-[0.01em] text-slate-700 dark:text-slate-200 mb-1.5"
          >
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
            className="w-full min-h-[44px] sm:min-h-[42px] px-4 py-2 text-base sm:text-xs font-medium rounded-full
              bg-white/55 hover:bg-white/68 focus:bg-white/85
              dark:bg-slate-900/50 dark:hover:bg-slate-900/65 dark:focus:bg-slate-900/80
              border border-white/80 dark:border-white/18
              shadow-[inset_0_1px_2px_rgba(15,23,42,0.05),0_1px_0_rgba(255,255,255,0.45)]
              dark:shadow-[inset_0_1px_2px_rgba(0,0,0,0.25)]
              text-slate-900 dark:text-white placeholder:text-slate-400 dark:placeholder:text-slate-400
              focus:outline-hidden focus:border-blue-500/80 focus:ring-2 focus:ring-blue-500/20
              dark:focus:border-blue-400/80 dark:focus:ring-blue-400/25
              disabled:opacity-60
              transition-[background-color,border-color,box-shadow] duration-200 motion-reduce:transition-none"
          />
        </div>

        {/* Password Field */}
        <div>
          <label
            htmlFor="password"
            className="block pl-3.5 text-[11.5px] font-semibold tracking-[0.01em] text-slate-700 dark:text-slate-200 mb-1.5"
          >
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
              className="w-full min-h-[44px] sm:min-h-[42px] pl-4 pr-11 py-2 text-base sm:text-xs font-medium rounded-full
                bg-white/55 hover:bg-white/68 focus:bg-white/85
                dark:bg-slate-900/50 dark:hover:bg-slate-900/65 dark:focus:bg-slate-900/80
                border border-white/80 dark:border-white/18
                shadow-[inset_0_1px_2px_rgba(15,23,42,0.05),0_1px_0_rgba(255,255,255,0.45)]
                dark:shadow-[inset_0_1px_2px_rgba(0,0,0,0.25)]
                text-slate-900 dark:text-white placeholder:text-slate-400 dark:placeholder:text-slate-400
                focus:outline-hidden focus:border-blue-500/80 focus:ring-2 focus:ring-blue-500/20
                dark:focus:border-blue-400/80 dark:focus:ring-blue-400/25
                disabled:opacity-60
                transition-[background-color,border-color,box-shadow] duration-200 motion-reduce:transition-none"
            />
            <button
              type="button"
              onClick={() => setShowPassword((prev) => !prev)}
              className="absolute right-2 top-1/2 -translate-y-1/2 text-slate-500 hover:text-slate-800 hover:bg-slate-900/5 dark:text-slate-300 dark:hover:text-white dark:hover:bg-white/10 cursor-pointer min-w-[34px] min-h-[34px] flex items-center justify-center rounded-full transition-colors duration-150 motion-reduce:transition-none touch-manipulation"
              title={showPassword ? 'Hide password' : 'Show password'}
              aria-label={showPassword ? 'Hide password' : 'Show password'}
            >
              {showPassword ? <EyeOff className="w-4 h-4" /> : <Eye className="w-4 h-4" />}
            </button>
          </div>
        </div>

        {/* Refined Pill Login Button */}
        <div className="pt-1.5">
          <button
            type="submit"
            disabled={isLoading}
            className="w-full min-h-[44px] sm:min-h-[42px] py-2.5 px-6 text-xs font-semibold tracking-wide text-white rounded-full
              bg-gradient-to-b from-blue-500 via-blue-600 to-blue-700
              hover:from-blue-500 hover:via-blue-500 hover:to-blue-600
              active:from-blue-700 active:to-blue-800
              border border-blue-400/40
              shadow-[0_10px_24px_-6px_rgba(37,99,235,0.48),0_2px_6px_rgba(15,23,42,0.12),inset_0_1px_0.5px_rgba(255,255,255,0.4)]
              hover:shadow-[0_12px_28px_-6px_rgba(37,99,235,0.58),0_3px_8px_rgba(15,23,42,0.14),inset_0_1px_0.5px_rgba(255,255,255,0.48)]
              active:scale-[0.988] motion-reduce:active:scale-100
              transition-all duration-200 motion-reduce:transition-none
              disabled:opacity-60 disabled:cursor-not-allowed cursor-pointer flex items-center justify-center touch-manipulation"
          >
            {isLoading ? 'Logging in...' : 'Login'}
          </button>
        </div>

        {/* Understated Text Link with High Readability Contrast */}
        <div className="text-center pt-1">
          {loginMode === 'staff' ? (
            <button
              type="button"
              onClick={() => {
                setLoginMode('admin');
                setError(null);
              }}
              className="text-xs text-slate-700 dark:text-slate-300 font-medium hover:text-blue-700 dark:hover:text-blue-400 hover:bg-white/35 dark:hover:bg-white/10 transition-colors duration-150 motion-reduce:transition-none cursor-pointer tracking-wide py-1.5 px-3 rounded-full touch-manipulation"
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
              className="text-xs text-slate-700 dark:text-slate-300 font-medium hover:text-blue-700 dark:hover:text-blue-400 hover:bg-white/35 dark:hover:bg-white/10 transition-colors duration-150 motion-reduce:transition-none cursor-pointer tracking-wide py-1.5 px-3 rounded-full touch-manipulation"
            >
              Back to Staff Login
            </button>
          )}
        </div>
      </form>
    </div>
  );
}
