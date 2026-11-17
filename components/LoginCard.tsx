'use client';

import React, { useState } from 'react';
import { User } from '@/lib/types';

interface LoginCardProps {
  onSuccess: (user: User) => void;
  externalError?: string | null;
}

export function LoginCard({ onSuccess, externalError }: LoginCardProps) {
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
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

    try {
      const res = await fetch('/api/auth/login', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ email: trimmedEmail, password })
      });

      const data = await res.json();

      if (!res.ok) {
        setError(data.error || 'Invalid credentials');
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
    <div className="w-full max-w-sm mx-auto p-8 bg-white border border-slate-300 rounded shadow-md">
      <div className="text-center mb-6">
        <h1 className="text-base font-bold text-slate-900 tracking-wide uppercase">MK Eye Center</h1>
        <p className="text-xs text-slate-500 mt-1">Staff Authentication Portal</p>
      </div>

      <form onSubmit={handleSubmit} className="space-y-4">
        {displayedError && (
          <div className="p-2.5 text-xs text-red-700 bg-red-50 border border-red-200 rounded">
            {displayedError}
          </div>
        )}

        <div>
          <label className="block text-xs font-semibold text-slate-700 mb-1">Email</label>
          <input
            type="email"
            placeholder="staff@mkeyecenter.com"
            value={email}
            onChange={(e) => setEmail(e.target.value)}
            disabled={isLoading}
            autoComplete="email"
            required
            className="w-full px-3 py-2 text-xs bg-white border border-slate-300 rounded text-slate-900 placeholder:text-slate-400 focus:outline-hidden focus:border-slate-800 focus:ring-1 focus:ring-slate-800"
          />
        </div>

        <div>
          <label className="block text-xs font-semibold text-slate-700 mb-1">Password</label>
          <input
            type="password"
            placeholder="••••••••"
            value={password}
            onChange={(e) => setPassword(e.target.value)}
            disabled={isLoading}
            autoComplete="current-password"
            required
            className="w-full px-3 py-2 text-xs bg-white border border-slate-300 rounded text-slate-900 placeholder:text-slate-400 focus:outline-hidden focus:border-slate-800 focus:ring-1 focus:ring-slate-800"
          />
        </div>

        <button
          type="submit"
          disabled={isLoading}
          className="w-full py-2 px-4 text-xs font-semibold text-white bg-slate-900 hover:bg-slate-800 active:bg-black rounded transition disabled:opacity-50 cursor-pointer shadow-xs"
        >
          {isLoading ? 'Signing In...' : 'Sign In to Workspaces'}
        </button>
      </form>
    </div>
  );
}
