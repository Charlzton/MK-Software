'use client';

import React from 'react';
import { AlertTriangle, X } from 'lucide-react';

interface ConfirmDialogProps {
  isOpen: boolean;
  title: string;
  message: string;
  confirmLabel?: string;
  confirmVariant?: 'danger' | 'primary';
  isLoading?: boolean;
  onConfirm: () => void;
  onCancel: () => void;
}

export function ConfirmDialog({
  isOpen,
  title,
  message,
  confirmLabel = 'Confirm Delete',
  confirmVariant = 'danger',
  isLoading = false,
  onConfirm,
  onCancel
}: ConfirmDialogProps) {
  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-3 sm:p-4 bg-slate-900/35 dark:bg-black/65 backdrop-blur-xs">
      <div className="relative w-full max-w-[calc(100vw-1.5rem)] sm:max-w-md max-h-[calc(100dvh-2rem)] flex flex-col mk-glass-panel text-[var(--control-text)] rounded-[24px] overflow-hidden">
        <div
          aria-hidden="true"
          className="pointer-events-none absolute inset-x-6 top-0 h-[1px] bg-gradient-to-r from-transparent via-white/95 dark:via-white/35 to-transparent"
        />
        <div className="p-5 sm:p-6 overflow-y-auto min-h-0">
          <div className="flex items-start gap-3 sm:gap-4">
            <div
              className={`p-2.5 rounded-2xl shrink-0 border ${
                confirmVariant === 'danger'
                  ? 'bg-red-50/85 border-red-200/80 text-red-600 dark:bg-red-950/60 dark:border-red-500/30 dark:text-red-400'
                  : 'bg-blue-50/85 border-blue-200/80 text-blue-600 dark:bg-sky-950/60 dark:border-sky-500/30 dark:text-sky-400'
              }`}
            >
              <AlertTriangle className="w-5 h-5" />
            </div>
            <div className="flex-1 min-w-0">
              <h3 className="text-sm sm:text-base font-semibold text-[var(--control-text)]">{title}</h3>
              <p className="mt-1.5 text-xs sm:text-sm text-[var(--control-muted)] leading-relaxed break-words">
                {message}
              </p>
            </div>
            <button
              type="button"
              onClick={onCancel}
              disabled={isLoading}
              aria-label="Close dialog"
              className="text-[var(--control-muted)] hover:text-[var(--control-text)] hover:bg-slate-900/5 dark:hover:bg-white/10 p-1.5 -mr-1 -mt-1 min-w-[36px] min-h-[36px] sm:min-w-0 sm:min-h-0 flex items-center justify-center rounded-full cursor-pointer shrink-0 transition-colors"
            >
              <X className="w-4 h-4" />
            </button>
          </div>
        </div>

        <div className="flex flex-wrap sm:flex-nowrap items-center justify-end gap-2.5 sm:gap-3 px-5 sm:px-6 py-3.5 bg-white/35 dark:bg-slate-900/40 border-t border-[var(--chrome-divider)] shrink-0">
          <button
            type="button"
            onClick={onCancel}
            disabled={isLoading}
            className="mk-glass-btn-secondary min-h-[40px] sm:min-h-[36px] px-4 py-1.5 text-xs font-semibold cursor-pointer touch-manipulation"
          >
            Cancel
          </button>
          <button
            type="button"
            onClick={onConfirm}
            disabled={isLoading}
            className={`min-h-[40px] sm:min-h-[36px] px-4.5 py-1.5 text-xs font-semibold flex items-center justify-center gap-2 cursor-pointer touch-manipulation disabled:opacity-60 ${
              confirmVariant === 'danger' ? 'mk-glass-btn-danger' : 'mk-glass-btn-primary'
            }`}
          >
            {isLoading && (
              <span className="w-3.5 h-3.5 border-2 border-white/30 border-t-white rounded-full animate-spin" />
            )}
            {confirmLabel}
          </button>
        </div>
      </div>
    </div>
  );
}
