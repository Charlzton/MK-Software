'use client';

import React, { useState } from 'react';
import { User } from '@/lib/types';
import { Stethoscope, Users, ShieldAlert, Eye, EyeOff, Lock, Mail } from 'lucide-react';

interface LoginCardProps {
  onSuccess: (user: User) => void;
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
        onSuccess(data.user);
      }
    } catch {
      setError('Unable to reach server. Please check network.');
    } finally {
      setIsLoading(false);
    }
  };

  const displayedError = error || externalError;

  return (
    <div className="w-full max-w-md mx-auto p-8 bg-slate-950/65 backdrop-blur-2xl border border-white/20 shadow-2xl rounded-2xl text-slate-100">
      <div className="text-center mb-6">
        <h2 className="text-lg font-bold tracking-tight text-white uppercase">
          {loginMode === 'admin' ? 'Administrator Portal' : 'Clinical Workspace Login'}
        </h2>
        <p className="text-xs text-slate-300 mt-1">
          {loginMode === 'admin'
            ? 'Sign in with administrator credentials'
            : 'Select your role and enter credentials to continue'}
        </p>
      </div>

      {/* Mode Switcher Tabs */}
      <div className="flex bg-slate-900/80 p-1 rounded-xl mb-6 border border-white/10">
        <button
          type="button"
          onClick={() => {
            setLoginMode('staff');
            setError(null);
          }}
          className={`flex-1 py-2 text-xs font-semibold rounded-lg transition cursor-pointer ${
            loginMode === 'staff'
              ? 'bg-blue-600 text-white shadow-md'
              : 'text-slate-300 hover:text-white hover:bg-white/5'
          }`}
        >
          Staff Login
        </button>
        <button
          type="button"
          onClick={() => {
            setLoginMode('admin');
            setError(null);
          }}
          className={`flex-1 py-2 text-xs font-semibold rounded-lg transition cursor-pointer ${
            loginMode === 'admin'
              ? 'bg-blue-600 text-white shadow-md'
              : 'text-slate-300 hover:text-white hover:bg-white/5'
          }`}
        >
          Admin Login
        </button>
      </div>

      <form onSubmit={handleSubmit} className="space-y-4">
        {displayedError && (
          <div className="p-3 text-xs text-red-200 bg-red-950/80 border border-red-500/40 rounded-xl flex items-start gap-2.5 shadow-inner">
            <ShieldAlert className="w-4 h-4 text-red-400 shrink-0 mt-0.5" />
            <span className="leading-relaxed">{displayedError}</span>
          </div>
        )}

        {/* Staff Role Selector */}
        {loginMode === 'staff' && (
          <div>
            <label className="block text-xs font-semibold text-slate-200 mb-2">Select Staff Role</label>
            <div className="grid grid-cols-2 gap-2.5">
              <button
                type="button"
                onClick={() => setStaffRole('reception')}
                className={`p-3 rounded-xl border text-xs font-semibold flex items-center justify-center gap-2 transition cursor-pointer ${
                  staffRole === 'reception'
                    ? 'border-blue-500 bg-blue-600/30 text-white shadow-md ring-1 ring-blue-500'
                    : 'border-white/10 bg-slate-900/50 text-slate-300 hover:bg-white/10 hover:border-white/20'
                }`}
              >
                <Users className="w-4 h-4 text-blue-400" />
                <span>Reception</span>
              </button>

              <button
                type="button"
                onClick={() => setStaffRole('doctor')}
                className={`p-3 rounded-xl border text-xs font-semibold flex items-center justify-center gap-2 transition cursor-pointer ${
                  staffRole === 'doctor'
                    ? 'border-blue-500 bg-blue-600/30 text-white shadow-md ring-1 ring-blue-500'
                    : 'border-white/10 bg-slate-900/50 text-slate-300 hover:bg-white/10 hover:border-white/20'
                }`}
              >
                <Stethoscope className="w-4 h-4 text-blue-400" />
                <span>Doctor</span>
              </button>
            </div>
          </div>
        )}

        <div>
          <label className="block text-xs font-semibold text-slate-200 mb-1.5 flex items-center gap-1.5">
            <Mail className="w-3.5 h-3.5 text-slate-400" />
            <span>Email Address</span>
          </label>
          <input
            type="email"
            placeholder={loginMode === 'admin' ? 'admin@mkeyecenter.com' : `${staffRole}@mkeyecenter.com`}
            value={email}
            onChange={(e) => setEmail(e.target.value)}
            disabled={isLoading}
            autoComplete="email"
            required
            className="w-full px-3.5 py-2.5 text-xs bg-slate-900/80 border border-white/15 rounded-xl text-white placeholder:text-slate-500 focus:outline-hidden focus:border-blue-500 focus:ring-1 focus:ring-blue-500 shadow-inner"
          />
        </div>

        <div>
          <label className="block text-xs font-semibold text-slate-200 mb-1.5 flex items-center gap-1.5">
            <Lock className="w-3.5 h-3.5 text-slate-400" />
            <span>Password</span>
          </label>
          <div className="relative">
            <input
              type={showPassword ? 'text' : 'password'}
              placeholder="••••••••"
              value={password}
              onChange={(e) => setPassword(e.target.value)}
              disabled={isLoading}
              autoComplete="current-password"
              required
              className="w-full pl-3.5 pr-10 py-2.5 text-xs bg-slate-900/80 border border-white/15 rounded-xl text-white placeholder:text-slate-500 focus:outline-hidden focus:border-blue-500 focus:ring-1 focus:ring-blue-500 shadow-inner"
            />
            <button
              type="button"
              onClick={() => setShowPassword((prev) => !prev)}
              className="absolute right-3 top-1/2 -translate-y-1/2 text-slate-400 hover:text-white cursor-pointer p-1"
              title={showPassword ? 'Hide password' : 'Show password'}
            >
              {showPassword ? <EyeOff className="w-4 h-4" /> : <Eye className="w-4 h-4" />}
            </button>
          </div>
        </div>

        <button
          type="submit"
          disabled={isLoading}
          className="w-full py-3 px-4 text-xs font-bold uppercase tracking-wider text-white bg-blue-600 hover:bg-blue-500 active:bg-blue-700 rounded-xl transition disabled:opacity-50 cursor-pointer shadow-lg shadow-blue-600/30 mt-3"
        >
          {isLoading
            ? 'Authenticating...'
            : loginMode === 'admin'
              ? 'Sign In as Administrator'
              : `Sign In as ${staffRole === 'doctor' ? 'Doctor' : 'Reception'}`}
        </button>
      </form>
    </div>
  );
}
