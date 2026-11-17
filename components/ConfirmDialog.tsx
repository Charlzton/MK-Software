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
    <div className="fixed inset-0 z-50 flex items-center justify-center p-3 sm:p-4 bg-slate-900/45 dark:bg-black/70 backdrop-blur-xs">
      <div className="relative w-full max-w-[calc(100vw-1.5rem)] sm:max-w-md max-h-[calc(100dvh-2rem)] flex flex-col bg-[var(--chrome-bg)] text-[var(--control-text)] rounded-xl shadow-xl border border-[var(--control-border)] overflow-hidden">
        <div className="p-4 sm:p-5 overflow-y-auto min-h-0">
          <div className="flex items-start gap-3 sm:gap-4">
            <div
              className={`p-2 sm:p-2.5 rounded-lg shrink-0 ${
                confirmVariant === 'danger'
                  ? 'bg-red-50 text-red-600 dark:bg-red-950/60 dark:text-red-400'
                  : 'bg-[var(--control-bg)] text-[var(--control-text)]'
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
              className="text-[var(--control-muted)] hover:text-[var(--control-text)] p-1.5 -mr-1 -mt-1 min-w-[36px] min-h-[36px] sm:min-w-0 sm:min-h-0 flex items-center justify-center rounded-md cursor-pointer shrink-0"
            >
              <X className="w-4 h-4" />
            </button>
          </div>
        </div>

        <div className="flex flex-wrap sm:flex-nowrap items-center justify-end gap-2.5 sm:gap-3 px-4 sm:px-5 py-3 sm:py-3.5 bg-[var(--chrome-elevated)] border-t border-[var(--chrome-border)] shrink-0">
          <button
            type="button"
            onClick={onCancel}
            disabled={isLoading}
            className="min-h-[40px] sm:min-h-0 px-4 py-2 text-xs font-medium text-[var(--control-text)] bg-[var(--control-bg)] border border-[var(--control-border)] rounded-lg hover:bg-[var(--control-bg-hover)] cursor-pointer touch-manipulation"
          >
            Cancel
          </button>
          <button
            type="button"
            onClick={onConfirm}
            disabled={isLoading}
            className={`min-h-[40px] sm:min-h-0 px-4 py-2 text-xs font-medium text-white rounded-lg flex items-center justify-center gap-2 cursor-pointer touch-manipulation ${
              confirmVariant === 'danger'
                ? 'bg-red-600 hover:bg-red-700'
                : 'bg-blue-600 hover:bg-blue-500'
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
