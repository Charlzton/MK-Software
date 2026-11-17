'use client';

import React, { useState, useEffect, useRef } from 'react';
import { SlidersHorizontal, X } from 'lucide-react';

interface ColumnWidthModalProps {
  isOpen: boolean;
  columnLabel: string;
  initialWidth: number;
  defaultWidth: number;
  minWidth?: number;
  maxWidth?: number;
  onApply: (width: number) => void;
  onResetDefault: () => void;
  onClose: () => void;
}

export function ColumnWidthModal({
  isOpen,
  columnLabel,
  initialWidth,
  defaultWidth,
  minWidth = 56,
  maxWidth = 600,
  onApply,
  onResetDefault,
  onClose
}: ColumnWidthModalProps) {
  if (!isOpen) return null;

  return (
    <ColumnWidthModalContent
      key={`${columnLabel}-${initialWidth}`}
      columnLabel={columnLabel}
      initialWidth={initialWidth}
      defaultWidth={defaultWidth}
      minWidth={minWidth}
      maxWidth={maxWidth}
      onApply={onApply}
      onResetDefault={onResetDefault}
      onClose={onClose}
    />
  );
}

function ColumnWidthModalContent({
  columnLabel,
  initialWidth,
  defaultWidth,
  minWidth,
  maxWidth,
  onApply,
  onResetDefault,
  onClose
}: Omit<ColumnWidthModalProps, 'isOpen'> & { minWidth: number; maxWidth: number }) {
  const [widthInput, setWidthInput] = useState<string>(String(initialWidth));
  const [error, setError] = useState<string | null>(null);
  const inputRef = useRef<HTMLInputElement | null>(null);

  useEffect(() => {
    const raf = window.requestAnimationFrame(() => {
      inputRef.current?.focus();
      inputRef.current?.select();
    });
    return () => window.cancelAnimationFrame(raf);
  }, []);

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    const parsed = Math.round(Number(widthInput));
    if (!Number.isFinite(parsed) || parsed < minWidth || parsed > maxWidth) {
      setError(`Enter a width between ${minWidth}px and ${maxWidth}px.`);
      return;
    }
    onApply(parsed);
  };

  return (
    <div
      className="fixed inset-0 z-[95] flex items-center justify-center p-3 sm:p-4 bg-slate-900/35 dark:bg-black/65 backdrop-blur-xs"
      role="dialog"
      aria-modal="true"
      aria-label={`Resize column ${columnLabel}`}
      onClick={onClose}
      onKeyDown={(e) => {
        e.stopPropagation();
        if (e.key === 'Escape') onClose();
      }}
    >
      <div
        className="relative w-full max-w-[calc(100vw-1.5rem)] sm:max-w-sm mk-glass-panel text-[var(--control-text)] rounded-[24px] overflow-hidden font-sans"
        onClick={(e) => e.stopPropagation()}
      >
        <div
          aria-hidden="true"
          className="pointer-events-none absolute inset-x-6 top-0 h-[1px] bg-gradient-to-r from-transparent via-white/95 dark:via-white/35 to-transparent"
        />
        <div className="px-5 py-3.5 border-b border-[var(--chrome-divider)] flex items-center justify-between gap-2">
          <div className="flex items-center gap-2 min-w-0">
            <SlidersHorizontal className="w-4 h-4 text-[var(--selection-ring)] shrink-0" />
            <h3 className="text-sm font-semibold truncate">Resize Column {columnLabel}</h3>
          </div>
          <button
            type="button"
            onClick={onClose}
            aria-label="Close resize modal"
            className="p-1.5 text-[var(--control-muted)] hover:text-[var(--control-text)] hover:bg-slate-900/5 dark:hover:bg-white/10 rounded-full cursor-pointer transition-colors"
          >
            <X className="w-4 h-4" />
          </button>
        </div>

        <form onSubmit={handleSubmit} className="p-5 space-y-4">
          <div>
            <label htmlFor="mk-col-width-input" className="block pl-3 text-[11.5px] font-semibold text-[var(--control-muted)] mb-1.5">
              Column width in pixels ({minWidth}–{maxWidth} px)
            </label>
            <input
              ref={inputRef}
              id="mk-col-width-input"
              type="number"
              min={minWidth}
              max={maxWidth}
              value={widthInput}
              onChange={(e) => {
                setWidthInput(e.target.value);
                setError(null);
              }}
              className="mk-glass-pill-input w-full min-h-[40px] sm:min-h-[38px] px-4 py-1.5 text-sm sm:text-xs font-mono tabular-nums"
            />
            {error && <p className="mt-1.5 pl-3 text-xs text-red-600 dark:text-red-400">{error}</p>}
          </div>

          <div className="flex flex-wrap items-center justify-between gap-2 pt-3 border-t border-[var(--chrome-divider)]">
            <button
              type="button"
              onClick={onResetDefault}
              className="min-h-[38px] sm:min-h-[34px] px-3 py-1.5 text-xs font-medium text-[var(--control-muted)] hover:text-[var(--control-text)] hover:bg-white/45 dark:hover:bg-white/10 rounded-full cursor-pointer touch-manipulation transition-colors"
            >
              Default ({defaultWidth}px)
            </button>
            <div className="flex items-center gap-2">
              <button
                type="button"
                onClick={onClose}
                className="mk-glass-btn-secondary min-h-[38px] sm:min-h-[34px] px-3.5 py-1.5 text-xs font-semibold cursor-pointer touch-manipulation"
              >
                Cancel
              </button>
              <button
                type="submit"
                className="mk-glass-btn-primary min-h-[38px] sm:min-h-[34px] px-4 py-1.5 text-xs font-semibold cursor-pointer touch-manipulation"
              >
                Apply
              </button>
            </div>
          </div>
        </form>
      </div>
    </div>
  );
}
