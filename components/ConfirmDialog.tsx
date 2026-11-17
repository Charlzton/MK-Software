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
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/45 dark:bg-black/70 backdrop-blur-xs">
      <div className="relative w-full max-w-md bg-[var(--chrome-bg)] text-[var(--control-text)] rounded-xl shadow-xl border border-[var(--control-border)] overflow-hidden">
        <div className="p-5">
          <div className="flex items-start gap-4">
            <div
              className={`p-2.5 rounded-lg shrink-0 ${
                confirmVariant === 'danger'
                  ? 'bg-red-50 text-red-600 dark:bg-red-950/60 dark:text-red-400'
                  : 'bg-[var(--control-bg)] text-[var(--control-text)]'
              }`}
            >
              <AlertTriangle className="w-5 h-5" />
            </div>
            <div className="flex-1 min-w-0">
              <h3 className="text-base font-semibold text-slate-900 dark:text-[#ececec]">{title}</h3>
              <p className="mt-1.5 text-sm text-slate-600 dark:text-[#b8b8b8] leading-relaxed">{message}</p>
            </div>
            <button
              type="button"
              onClick={onCancel}
              disabled={isLoading}
              className="text-slate-400 hover:text-slate-600 dark:text-[#909090] dark:hover:text-[#ececec] p-1 rounded-md transition-colors cursor-pointer"
            >
              <X className="w-4 h-4" />
            </button>
          </div>
        </div>

        <div className="flex items-center justify-end gap-3 px-5 py-3.5 bg-[var(--chrome-elevated)] border-t border-[var(--chrome-border)]">
          <button
            type="button"
            onClick={onCancel}
            disabled={isLoading}
            className="px-4 py-2 text-xs font-medium text-[var(--control-text)] bg-[var(--control-bg)] border border-[var(--control-border)] rounded-lg hover:bg-[var(--control-bg-hover)] transition-colors cursor-pointer"
          >
            Cancel
          </button>
          <button
            type="button"
            onClick={onConfirm}
            disabled={isLoading}
            className={`px-4 py-2 text-xs font-medium text-white rounded-lg transition-colors flex items-center gap-2 cursor-pointer ${
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
