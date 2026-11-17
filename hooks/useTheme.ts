'use client';

import { useEffect, useCallback, useSyncExternalStore } from 'react';

export type ThemeMode = 'light' | 'dark';

export const THEME_STORAGE_KEY = 'mk_theme';
const THEME_CHANGE_EVENT = 'mk-theme-change';

export function readStoredTheme(): ThemeMode {
  if (typeof window === 'undefined') return 'light';
  try {
    const stored = window.localStorage.getItem(THEME_STORAGE_KEY);
    return stored === 'dark' ? 'dark' : 'light';
  } catch {
    return 'light';
  }
}

export function syncThemeClassToDocument(): ThemeMode {
  const mode = readStoredTheme();
  if (typeof document !== 'undefined') {
    document.documentElement.classList.toggle('dark', mode === 'dark');
  }
  return mode;
}

function subscribeTheme(callback: () => void): () => void {
  if (typeof window === 'undefined') return () => {};
  const handleStorage = (e: StorageEvent) => {
    if (e.key === null || e.key === THEME_STORAGE_KEY) {
      syncThemeClassToDocument();
      callback();
    }
  };
  const handleCustom = () => callback();
  window.addEventListener('storage', handleStorage);
  window.addEventListener(THEME_CHANGE_EVENT, handleCustom);
  return () => {
    window.removeEventListener('storage', handleStorage);
    window.removeEventListener(THEME_CHANGE_EVENT, handleCustom);
  };
}

function getServerThemeSnapshot(): ThemeMode {
  return 'light';
}

export function useTheme() {
  const theme = useSyncExternalStore(
    subscribeTheme,
    readStoredTheme,
    getServerThemeSnapshot
  );

  useEffect(() => {
    document.documentElement.classList.toggle('dark', theme === 'dark');
  }, [theme]);

  const applyTheme = useCallback((next: ThemeMode) => {
    if (typeof document !== 'undefined') {
      const root = document.documentElement;
      root.classList.add('mk-theme-switching');
      root.classList.toggle('dark', next === 'dark');
      void root.getBoundingClientRect();
      window.requestAnimationFrame(() => {
        window.requestAnimationFrame(() => {
          root.classList.remove('mk-theme-switching');
        });
      });
    }
    if (typeof window !== 'undefined') {
      try {
        window.localStorage.setItem(THEME_STORAGE_KEY, next);
      } catch {
        // Ignore storage write restrictions
      }
      window.dispatchEvent(new Event(THEME_CHANGE_EVENT));
    }
  }, []);

  const toggleTheme = useCallback(() => {
    const current = readStoredTheme();
    applyTheme(current === 'dark' ? 'light' : 'dark');
  }, [applyTheme]);

  return { theme, setTheme: applyTheme, toggleTheme };
}
