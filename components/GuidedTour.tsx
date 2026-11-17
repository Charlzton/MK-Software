'use client';

import React, { useState, useEffect, useCallback, useLayoutEffect } from 'react';
import { X, ChevronLeft, ChevronRight, Sparkles, Check } from 'lucide-react';
import type { UserRole } from '@/lib/types';

export type TourTargetKey =
  | 'brand-header'
  | 'workspace-switcher'
  | 'grid-viewport'
  | 'formula-bar'
  | 'worksheet-tabs'
  | 'save-export'
  | 'utility-controls';

interface TourStep {
  target: TourTargetKey;
  title: string;
  description: string;
  tips?: string[];
}

interface SpotlightRect {
  top: number;
  left: number;
  width: number;
  height: number;
}

interface GuidedTourProps {
  isOpen: boolean;
  role: UserRole;
  workspace: 'reception' | 'doctor';
  hasWorkspaceSwitcher: boolean;
  onClose: () => void;
  onComplete: () => void;
}

function buildSteps(
  role: UserRole,
  workspace: 'reception' | 'doctor',
  hasWorkspaceSwitcher: boolean
): TourStep[] {
  const firstStep: TourStep =
    role === 'admin' && hasWorkspaceSwitcher
      ? {
          target: 'workspace-switcher',
          title: 'Workspace Switcher',
          description: 'Switch seamlessly between Reception and Doctor workbooks.'
        }
      : role === 'doctor'
        ? {
            target: 'brand-header',
            title: 'Doctor Clinical Workspace',
            description: 'Review and update clinical prescriptions and patient records.'
          }
        : {
            target: 'brand-header',
            title: 'Reception Desk Workspace',
            description: 'Manage patient registers and front-desk clinic records.'
          };

  return [
    firstStep,
    {
      target: 'grid-viewport',
      title: 'Spreadsheet Grid',
      description: 'Click any cell to select or navigate quickly using keyboard shortcuts.'
    },
    {
      target: 'formula-bar',
      title: 'Formula Bar',
      description: 'View active cell coordinates and edit entries directly.'
    },
    {
      target: 'worksheet-tabs',
      title: 'Worksheet Tabs',
      description: 'Switch between sheets, add new tabs, or right-click to manage sheets.'
    },
    {
      target: 'save-export',
      title: 'Save & Export',
      description: 'Synchronize changes to the cloud or export your data to Excel.'
    },
    {
      target: 'utility-controls',
      title: 'Theme & Help',
      description: 'Toggle light/dark modes or restart this guide anytime.'
    }
  ];
}


export function GuidedTour(props: GuidedTourProps) {
  if (!props.isOpen) return null;
  return <GuidedTourContent {...props} />;
}

function GuidedTourContent({
  role,
  workspace,
  hasWorkspaceSwitcher,
  onClose,
  onComplete
}: GuidedTourProps) {
  const tourCardRef = React.useRef<HTMLDivElement | null>(null);
  const steps = buildSteps(role, workspace, hasWorkspaceSwitcher);
  const [stepIndex, setStepIndex] = useState(0);
  const [rect, setRect] = useState<SpotlightRect | null>(null);
  const [viewport, setViewport] = useState<{ width: number; height: number }>({
    width: 1280,
    height: 720
  });
  const [measuredCardHeight, setMeasuredCardHeight] = useState<number>(250);

  const currentStep = steps[Math.min(stepIndex, steps.length - 1)];

  const measureTarget = useCallback(() => {
    if (!currentStep) return;
    const vv = window.visualViewport;
    const vw = Math.round(vv?.width ?? window.innerWidth);
    const vh = Math.round(vv?.height ?? window.innerHeight);
    setViewport({ width: vw, height: vh });

    if (tourCardRef.current) {
      const cardRect = tourCardRef.current.getBoundingClientRect();
      if (cardRect.height > 0) {
        setMeasuredCardHeight(Math.ceil(cardRect.height));
      }
    }

    const el = document.querySelector(`[data-tour="${currentStep.target}"]`) as HTMLElement | null;
    if (!el) {
      setRect(null);
      return;
    }

    const r = el.getBoundingClientRect();
    if (r.width === 0 || r.height === 0) {
      setRect(null);
      return;
    }

    const pad = vw < 480 ? 4 : 6;
    const rawTop = Math.max(4, r.top - pad);
    const rawLeft = Math.max(4, r.left - pad);
    const clampedWidth = Math.max(24, Math.min(r.width + pad * 2, vw - rawLeft - 4));
    const clampedHeight = Math.max(24, Math.min(r.height + pad * 2, vh - rawTop - 4));
    setRect({
      top: rawTop,
      left: rawLeft,
      width: clampedWidth,
      height: clampedHeight
    });
  }, [currentStep]);

  useLayoutEffect(() => {
    const raf = window.requestAnimationFrame(measureTarget);
    const raf2 = window.requestAnimationFrame(() => {
      window.requestAnimationFrame(measureTarget);
    });
    const handleResize = () => measureTarget();
    window.addEventListener('resize', handleResize);
    window.addEventListener('scroll', handleResize, true);
    window.visualViewport?.addEventListener('resize', handleResize);
    window.visualViewport?.addEventListener('scroll', handleResize);
    return () => {
      window.cancelAnimationFrame(raf);
      window.cancelAnimationFrame(raf2);
      window.removeEventListener('resize', handleResize);
      window.removeEventListener('scroll', handleResize, true);
      window.visualViewport?.removeEventListener('resize', handleResize);
      window.visualViewport?.removeEventListener('scroll', handleResize);
    };
  }, [stepIndex, measureTarget]);

  useEffect(() => {
    const handleKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        e.preventDefault();
        e.stopPropagation();
        onClose();
      } else if (e.key === 'ArrowRight') {
        e.preventDefault();
        e.stopPropagation();
        if (stepIndex < steps.length - 1) {
          setStepIndex((i) => i + 1);
        } else {
          onComplete();
        }
      } else if (e.key === 'ArrowLeft') {
        e.preventDefault();
        e.stopPropagation();
        setStepIndex((i) => Math.max(0, i - 1));
      }
    };
    window.addEventListener('keydown', handleKey, true);
    return () => window.removeEventListener('keydown', handleKey, true);
  }, [stepIndex, steps.length, onClose, onComplete]);

  if (!currentStep) return null;

  const isFirst = stepIndex === 0;
  const isLast = stepIndex === steps.length - 1;
  const roleLabel =
    role === 'admin' ? 'Admin Guide' : role === 'doctor' ? 'Doctor Guide' : 'Reception Guide';

  // Compute safe, viewport-clamped popover coordinates across 320px phones to ultra-wide displays
  const margin = viewport.width < 420 ? 12 : 16;
  const maxCardHeight = Math.max(180, viewport.height - margin * 2);
  const cardWidth = Math.min(360, Math.max(240, viewport.width - margin * 2));
  const effectiveCardHeight = Math.min(measuredCardHeight, maxCardHeight);
  let cardTop = Math.max(margin, Math.round((viewport.height - effectiveCardHeight) / 2));
  let cardLeft = Math.round((viewport.width - cardWidth) / 2);

  if (rect) {
    const spaceBelow = viewport.height - (rect.top + rect.height) - margin;
    const spaceAbove = rect.top - margin;

    if (currentStep.target === 'grid-viewport') {
      if (viewport.height < 520) {
        cardTop = Math.round((viewport.height - effectiveCardHeight) / 2);
      } else {
        cardTop = Math.max(
          margin,
          Math.min(viewport.height - effectiveCardHeight - margin, rect.top + 28)
        );
      }
      cardLeft = Math.round((viewport.width - cardWidth) / 2);
    } else if (spaceBelow >= effectiveCardHeight + 12) {
      cardTop = rect.top + rect.height + 12;
      cardLeft = rect.left + rect.width / 2 - cardWidth / 2;
    } else if (spaceAbove >= effectiveCardHeight + 12) {
      cardTop = rect.top - effectiveCardHeight - 12;
      cardLeft = rect.left + rect.width / 2 - cardWidth / 2;
    } else {
      cardTop = Math.round((viewport.height - effectiveCardHeight) / 2);
      cardLeft = Math.round((viewport.width - cardWidth) / 2);
    }

    cardLeft = Math.max(margin, Math.min(viewport.width - cardWidth - margin, Math.round(cardLeft)));
    cardTop = Math.max(
      margin,
      Math.min(viewport.height - effectiveCardHeight - margin, Math.round(cardTop))
    );
  }

  return (
    <div
      className="fixed inset-0 z-[80]"
      role="dialog"
      aria-modal="true"
      aria-label={`Guided tour: ${currentStep.title}`}
      onKeyDown={(e) => e.stopPropagation()}
    >
      {/* Dimmed spotlight cutout or full backdrop */}
      {rect ? (
        <div
          className="fixed pointer-events-none rounded-lg border-2 border-[var(--selection-ring)] transition-all duration-150 ease-out"
          style={{
            top: rect.top,
            left: rect.left,
            width: rect.width,
            height: rect.height,
            boxShadow: '0 0 0 9999px rgba(15, 23, 42, 0.58)'
          }}
        />
      ) : (
        <div className="fixed inset-0 bg-slate-900/60" />
      )}

      {/* Click-blocker backdrop outside card so grid clicks don't accidentally edit cells while tour is open */}
      <div className="fixed inset-0" onClick={onClose} aria-hidden="true" />

      {/* Tour Card */}
      <div
        ref={tourCardRef}
        onClick={(e) => e.stopPropagation()}
        style={{
          top: cardTop,
          left: cardLeft,
          width: cardWidth,
          maxHeight: maxCardHeight
        }}
        className="fixed z-[85] rounded-[22px] mk-glass-panel p-4 sm:p-5 text-[var(--control-text)] flex flex-col overflow-hidden"
      >
        <div
          aria-hidden="true"
          className="pointer-events-none absolute inset-x-6 top-0 h-[1px] bg-gradient-to-r from-transparent via-white/95 dark:via-white/35 to-transparent"
        />
        <div className="flex items-start justify-between gap-2 mb-2 shrink-0">
          <div className="flex flex-wrap items-center gap-x-1.5 gap-y-0.5 min-w-0">
            <Sparkles className="w-3.5 h-3.5 text-[var(--selection-ring)] shrink-0" />
            <span className="text-[11px] font-semibold uppercase tracking-wider text-[var(--selection-ring)]">
              {roleLabel}
            </span>
            <span className="text-[11px] text-[var(--control-muted)] font-mono tabular-nums">
              • Step {stepIndex + 1} of {steps.length}
            </span>
          </div>
          <button
            type="button"
            onClick={onClose}
            aria-label="Close guided tour"
            className="p-1.5 -m-1.5 min-w-[32px] min-h-[32px] sm:min-w-0 sm:min-h-0 flex items-center justify-center text-[var(--control-muted)] hover:text-[var(--control-text)] hover:bg-slate-900/5 dark:hover:bg-white/10 rounded-full cursor-pointer shrink-0 touch-manipulation transition-colors"
          >
            <X className="w-4 h-4" />
          </button>
        </div>

        <div className="overflow-y-auto min-h-0 flex-1 pr-0.5">
          <h3 className="text-sm font-bold text-[var(--control-text)] mb-1.5">
            {currentStep.title}
          </h3>
          <p className="text-xs text-[var(--control-muted)] leading-relaxed mb-3">
            {currentStep.description}
          </p>

          {currentStep.tips && currentStep.tips.length > 0 && (
            <ul className="space-y-1 mb-3 bg-white/45 dark:bg-slate-900/45 border border-white/75 dark:border-white/10 rounded-2xl px-3 py-2.5">
              {currentStep.tips.map((tip, idx) => (
                <li
                  key={idx}
                  className="text-[11px] text-[var(--control-text)] flex items-start gap-1.5 leading-snug"
                >
                  <span className="text-[var(--selection-ring)] font-bold shrink-0">•</span>
                  <span>{tip}</span>
                </li>
              ))}
            </ul>
          )}
        </div>

        <div className="flex items-center justify-between gap-2 pt-3 mt-0.5 border-t border-[var(--chrome-divider)] shrink-0">
          <button
            type="button"
            onClick={onClose}
            className="min-h-[34px] sm:min-h-0 text-[11px] font-medium text-[var(--control-muted)] hover:text-[var(--control-text)] px-2.5 py-1 rounded-full hover:bg-white/45 dark:hover:bg-white/10 cursor-pointer touch-manipulation transition-colors"
          >
            Skip Tour
          </button>

          <div className="flex items-center gap-1.5">
            {!isFirst && (
              <button
                type="button"
                onClick={() => setStepIndex((i) => Math.max(0, i - 1))}
                className="mk-glass-btn-secondary min-h-[34px] sm:min-h-[30px] px-3 py-1 text-xs font-semibold flex items-center gap-1 cursor-pointer touch-manipulation"
              >
                <ChevronLeft className="w-3.5 h-3.5" />
                <span>Back</span>
              </button>
            )}
            <button
              type="button"
              onClick={() => {
                if (isLast) {
                  onComplete();
                } else {
                  setStepIndex((i) => i + 1);
                }
              }}
              className="mk-glass-btn-primary min-h-[34px] sm:min-h-[30px] px-3.5 py-1 text-xs font-semibold flex items-center gap-1 cursor-pointer touch-manipulation"
            >
              {isLast ? (
                <>
                  <Check className="w-3.5 h-3.5" />
                  <span>Finish Tour</span>
                </>
              ) : (
                <>
                  <span>Next</span>
                  <ChevronRight className="w-3.5 h-3.5" />
                </>
              )}
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}
