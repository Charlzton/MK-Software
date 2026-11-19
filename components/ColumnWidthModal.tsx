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
      className="fixed inset-0 z-[95] flex items-center justify-center p-3 sm:p-4 bg-slate-900/45 dark:bg-black/70 backdrop-blur-xs"
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
        className="w-full max-w-[calc(100vw-1.5rem)] sm:max-w-sm bg-[var(--chrome-bg)] text-[var(--control-text)] border border-[var(--control-border)] rounded-xl shadow-2xl overflow-hidden font-sans"
        onClick={(e) => e.stopPropagation()}
      >
        <div className="px-4 py-3 border-b border-[var(--chrome-border)] flex items-center justify-between gap-2">
          <div className="flex items-center gap-2 min-w-0">
            <SlidersHorizontal className="w-4 h-4 text-[var(--selection-ring)] shrink-0" />
            <h3 className="text-sm font-semibold truncate">Resize Column {columnLabel}</h3>
          </div>
          <button
            type="button"
            onClick={onClose}
            aria-label="Close resize modal"
            className="p-1 text-[var(--control-muted)] hover:text-[var(--control-text)] rounded cursor-pointer"
          >
            <X className="w-4 h-4" />
          </button>
        </div>

        <form onSubmit={handleSubmit} className="p-4 space-y-3">
          <div>
            <label htmlFor="mk-col-width-input" className="block text-xs font-medium text-[var(--control-muted)] mb-1.5">
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
              className="w-full min-h-[40px] sm:min-h-[36px] px-3 py-1.5 text-sm sm:text-xs font-mono tabular-nums bg-[var(--formula-input-bg)] border border-[var(--control-border)] rounded-lg text-[var(--control-text)] outline-none focus:border-[var(--selection-ring)] focus:ring-1 focus:ring-[var(--selection-ring)]"
            />
            {error && <p className="mt-1.5 text-xs text-red-600 dark:text-red-400">{error}</p>}
          </div>

          <div className="flex flex-wrap items-center justify-between gap-2 pt-2 border-t border-[var(--chrome-border)]">
            <button
              type="button"
              onClick={onResetDefault}
              className="min-h-[38px] sm:min-h-0 px-3 py-1.5 text-xs font-medium text-[var(--control-muted)] hover:text-[var(--control-text)] hover:bg-[var(--control-bg-hover)] rounded-lg cursor-pointer touch-manipulation"
            >
              Default ({defaultWidth}px)
            </button>
            <div className="flex items-center gap-2">
              <button
                type="button"
                onClick={onClose}
                className="min-h-[38px] sm:min-h-0 px-3 py-1.5 text-xs font-medium bg-[var(--control-bg)] hover:bg-[var(--control-bg-hover)] border border-[var(--control-border)] rounded-lg cursor-pointer touch-manipulation"
              >
                Cancel
              </button>
              <button
                type="submit"
                className="min-h-[38px] sm:min-h-0 px-3.5 py-1.5 text-xs font-semibold bg-blue-600 hover:bg-blue-500 text-white rounded-lg cursor-pointer touch-manipulation"
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
