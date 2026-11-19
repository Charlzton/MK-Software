'use client';

import React, { useEffect, useLayoutEffect, useRef, useState } from 'react';
import {
  Scissors,
  Copy,
  ClipboardPaste,
  Trash2,
  ArrowUp,
  ArrowDown,
  ArrowLeft,
  ArrowRight,
  SlidersHorizontal,
  MousePointerClick,
  Pencil,
  ChevronLeft,
  ChevronRight,
  Loader2
} from 'lucide-react';

export type ContextMenuTarget =
  | {
      kind: 'cell';
      x: number;
      y: number;
      row: number;
      col: number;
      label: string;
      hasContent: boolean;
      canPaste: boolean;
    }
  | {
      kind: 'row';
      x: number;
      y: number;
      startRow: number;
      endRow: number;
      label: string;
      isHeaderRow: boolean;
      hasContent: boolean;
      canPaste: boolean;
    }
  | {
      kind: 'col';
      x: number;
      y: number;
      col: number;
      endCol: number;
      label: string;
      hasContent: boolean;
      canPaste: boolean;
      canInsert: boolean;
      canDelete: boolean;
    }
  | {
      kind: 'tab';
      x: number;
      y: number;
      sheetName: string;
      canMoveLeft: boolean;
      canMoveRight: boolean;
      canDelete: boolean;
      isDeletingThis: boolean;
    };

export type ContextMenuActionId =
  | 'cell_select'
  | 'cell_cut'
  | 'cell_copy'
  | 'cell_paste'
  | 'cell_delete'
  | 'row_select'
  | 'row_cut'
  | 'row_copy'
  | 'row_paste'
  | 'row_insert_above'
  | 'row_insert_below'
  | 'row_delete'
  | 'col_select'
  | 'col_cut'
  | 'col_copy'
  | 'col_paste'
  | 'col_insert_left'
  | 'col_insert_right'
  | 'col_delete'
  | 'col_resize'
  | 'tab_rename'
  | 'tab_move_left'
  | 'tab_move_right'
  | 'tab_delete';

interface SpreadsheetContextMenuProps {
  menu: ContextMenuTarget | null;
  isBusy?: boolean;
  onAction: (action: ContextMenuActionId) => void;
  onClose: () => void;
}

export function SpreadsheetContextMenu({
  menu,
  isBusy = false,
  onAction,
  onClose
}: SpreadsheetContextMenuProps) {
  const menuRef = useRef<HTMLDivElement | null>(null);
  const [coords, setCoords] = useState<{ left: number; top: number } | null>(null);

  useLayoutEffect(() => {
    if (!menu) return;
    const updatePosition = () => {
      const vw = window.innerWidth;
      const vh = window.innerHeight;
      const rect = menuRef.current?.getBoundingClientRect();
      const width = rect?.width || 208;
      const height = rect?.height || (menu.kind === 'tab' ? 140 : 280);

      let left = menu.x + 2;
      let top = menu.kind === 'tab' ? menu.y - height - 6 : menu.y + 2;

      if (left + width > vw - 8) {
        left = Math.max(8, menu.x - width - 2);
      }
      if (left < 8) left = 8;

      if (top + height > vh - 8) {
        top = Math.max(8, menu.y - height - 2);
      }
      if (top < 8) {
        top = Math.max(8, Math.min(vh - height - 8, menu.y + 4));
      }

      setCoords({ left: Math.round(left), top: Math.round(top) });
    };

    updatePosition();
    const raf = window.requestAnimationFrame(updatePosition);
    return () => window.cancelAnimationFrame(raf);
  }, [menu]);

  useEffect(() => {
    if (!menu) return;

    const handlePointerDown = (e: PointerEvent) => {
      if (menuRef.current && menuRef.current.contains(e.target as Node)) return;
      onClose();
    };
    const handleKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        e.preventDefault();
        e.stopPropagation();
        onClose();
      }
    };
    const handleScroll = () => {
      onClose();
    };

    window.addEventListener('pointerdown', handlePointerDown, true);
    window.addEventListener('keydown', handleKeyDown, true);
    window.addEventListener('resize', handleScroll);
    return () => {
      window.removeEventListener('pointerdown', handlePointerDown, true);
      window.removeEventListener('keydown', handleKeyDown, true);
      window.removeEventListener('resize', handleScroll);
    };
  }, [menu, onClose]);

  if (!menu) return null;

  const itemClass =
    'w-full text-left px-3 py-2 sm:py-1.5 hover:bg-[var(--control-bg-hover)] disabled:opacity-40 disabled:hover:bg-transparent disabled:cursor-not-allowed flex items-center gap-2.5 cursor-pointer touch-manipulation text-xs font-medium text-[var(--control-text)] transition-colors duration-75';

  const dangerItemClass =
    'w-full text-left px-3 py-2 sm:py-1.5 hover:bg-red-50 dark:hover:bg-red-950/60 text-red-600 dark:text-red-400 disabled:opacity-40 disabled:hover:bg-transparent disabled:cursor-not-allowed flex items-center gap-2.5 cursor-pointer touch-manipulation text-xs font-medium transition-colors duration-75';

  const headerTitle =
    menu.kind === 'tab'
      ? menu.sheetName
      : menu.label;

  return (
    <div
      ref={menuRef}
      role="menu"
      aria-label={`Context menu for ${headerTitle}`}
      onContextMenu={(e) => {
        e.preventDefault();
        e.stopPropagation();
      }}
      onClick={(e) => e.stopPropagation()}
      onPointerDown={(e) => e.stopPropagation()}
      style={{
        left: coords?.left ?? menu.x,
        top: coords?.top ?? menu.y,
        visibility: coords ? 'visible' : 'hidden'
      }}
      className="fixed z-[90] w-52 max-w-[calc(100vw-16px)] max-h-[calc(100dvh-16px)] overflow-y-auto bg-[var(--chrome-bg)] border border-[var(--control-border)] rounded-xl shadow-2xl py-1 text-xs text-[var(--control-text)] font-sans select-none"
    >
      <div className="px-3 py-1.5 font-semibold text-[10px] uppercase tracking-wider text-[var(--control-muted)] border-b border-[var(--chrome-border)] truncate">
        {headerTitle}
      </div>

      {menu.kind === 'cell' && (
        <>
          <button
            type="button"
            role="menuitem"
            disabled={isBusy}
            onClick={() => onAction('cell_cut')}
            className={itemClass}
          >
            <Scissors className="w-3.5 h-3.5 text-[var(--control-muted)] shrink-0" />
            <span>Cut</span>
          </button>
          <button
            type="button"
            role="menuitem"
            disabled={isBusy}
            onClick={() => onAction('cell_copy')}
            className={itemClass}
          >
            <Copy className="w-3.5 h-3.5 text-[var(--control-muted)] shrink-0" />
            <span>Copy</span>
          </button>
          <button
            type="button"
            role="menuitem"
            disabled={isBusy || !menu.canPaste}
            onClick={() => onAction('cell_paste')}
            className={itemClass}
          >
            <ClipboardPaste className="w-3.5 h-3.5 text-[var(--control-muted)] shrink-0" />
            <span>Paste</span>
          </button>
          <div className="my-1 border-t border-[var(--chrome-border)]" />
          <button
            type="button"
            role="menuitem"
            disabled={isBusy || !menu.hasContent}
            onClick={() => onAction('cell_delete')}
            className={dangerItemClass}
          >
            <Trash2 className="w-3.5 h-3.5 shrink-0" />
            <span>Delete Contents</span>
          </button>
          <button
            type="button"
            role="menuitem"
            disabled={isBusy}
            onClick={() => onAction('cell_select')}
            className={`${itemClass} border-t border-[var(--chrome-border)] mt-1 pt-2 sm:pt-1.5`}
          >
            <MousePointerClick className="w-3.5 h-3.5 text-[var(--control-muted)] shrink-0" />
            <span>Select Cell</span>
          </button>
        </>
      )}

      {menu.kind === 'row' && (
        <>
          <button
            type="button"
            role="menuitem"
            disabled={isBusy}
            onClick={() => onAction('row_select')}
            className={itemClass}
          >
            <MousePointerClick className="w-3.5 h-3.5 text-[var(--control-muted)] shrink-0" />
            <span>{menu.endRow > menu.startRow ? 'Select Rows' : 'Select Row'}</span>
          </button>
          <div className="my-1 border-t border-[var(--chrome-border)]" />
          <button
            type="button"
            role="menuitem"
            disabled={isBusy || menu.isHeaderRow}
            title={menu.isHeaderRow ? 'Column header row cannot be cut' : undefined}
            onClick={() => onAction('row_cut')}
            className={itemClass}
          >
            <Scissors className="w-3.5 h-3.5 text-[var(--control-muted)] shrink-0" />
            <span>{menu.endRow > menu.startRow ? 'Cut Rows' : 'Cut Row'}</span>
          </button>
          <button
            type="button"
            role="menuitem"
            disabled={isBusy}
            onClick={() => onAction('row_copy')}
            className={itemClass}
          >
            <Copy className="w-3.5 h-3.5 text-[var(--control-muted)] shrink-0" />
            <span>{menu.endRow > menu.startRow ? 'Copy Rows' : 'Copy Row'}</span>
          </button>
          <button
            type="button"
            role="menuitem"
            disabled={isBusy || !menu.canPaste}
            onClick={() => onAction('row_paste')}
            className={itemClass}
          >
            <ClipboardPaste className="w-3.5 h-3.5 text-[var(--control-muted)] shrink-0" />
            <span>{menu.endRow > menu.startRow ? 'Paste Rows' : 'Paste Row'}</span>
          </button>
          <div className="my-1 border-t border-[var(--chrome-border)]" />
          <button
            type="button"
            role="menuitem"
            disabled={isBusy || menu.isHeaderRow}
            title={menu.isHeaderRow ? 'Cannot insert a data row above column headers' : undefined}
            onClick={() => onAction('row_insert_above')}
            className={itemClass}
          >
            <ArrowUp className="w-3.5 h-3.5 text-[var(--control-muted)] shrink-0" />
            <span>Insert Row Above</span>
          </button>
          <button
            type="button"
            role="menuitem"
            disabled={isBusy}
            onClick={() => onAction('row_insert_below')}
            className={itemClass}
          >
            <ArrowDown className="w-3.5 h-3.5 text-[var(--control-muted)] shrink-0" />
            <span>Insert Row Below</span>
          </button>
          <div className="my-1 border-t border-[var(--chrome-border)]" />
          <button
            type="button"
            role="menuitem"
            disabled={isBusy || menu.isHeaderRow}
            title={menu.isHeaderRow ? 'Column header row cannot be deleted' : undefined}
            onClick={() => onAction('row_delete')}
            className={dangerItemClass}
          >
            <Trash2 className="w-3.5 h-3.5 shrink-0" />
            <span>{menu.endRow > menu.startRow ? 'Delete Rows' : 'Delete Row'}</span>
          </button>
        </>
      )}

      {menu.kind === 'col' && (
        <>
          <button
            type="button"
            role="menuitem"
            disabled={isBusy}
            onClick={() => onAction('col_select')}
            className={itemClass}
          >
            <MousePointerClick className="w-3.5 h-3.5 text-[var(--control-muted)] shrink-0" />
            <span>Select Column</span>
          </button>
          <div className="my-1 border-t border-[var(--chrome-border)]" />
          <button
            type="button"
            role="menuitem"
            disabled={isBusy}
            onClick={() => onAction('col_cut')}
            className={itemClass}
          >
            <Scissors className="w-3.5 h-3.5 text-[var(--control-muted)] shrink-0" />
            <span>Cut Column</span>
          </button>
          <button
            type="button"
            role="menuitem"
            disabled={isBusy}
            onClick={() => onAction('col_copy')}
            className={itemClass}
          >
            <Copy className="w-3.5 h-3.5 text-[var(--control-muted)] shrink-0" />
            <span>Copy Column</span>
          </button>
          <button
            type="button"
            role="menuitem"
            disabled={isBusy || !menu.canPaste}
            onClick={() => onAction('col_paste')}
            className={itemClass}
          >
            <ClipboardPaste className="w-3.5 h-3.5 text-[var(--control-muted)] shrink-0" />
            <span>Paste Column</span>
          </button>
          <div className="my-1 border-t border-[var(--chrome-border)]" />
          <button
            type="button"
            role="menuitem"
            disabled={isBusy || !menu.canInsert}
            onClick={() => onAction('col_insert_left')}
            className={itemClass}
          >
            <ArrowLeft className="w-3.5 h-3.5 text-[var(--control-muted)] shrink-0" />
            <span>Insert Column Left</span>
          </button>
          <button
            type="button"
            role="menuitem"
            disabled={isBusy || !menu.canInsert}
            onClick={() => onAction('col_insert_right')}
            className={itemClass}
          >
            <ArrowRight className="w-3.5 h-3.5 text-[var(--control-muted)] shrink-0" />
            <span>Insert Column Right</span>
          </button>
          <button
            type="button"
            role="menuitem"
            disabled={isBusy}
            onClick={() => onAction('col_resize')}
            className={itemClass}
          >
            <SlidersHorizontal className="w-3.5 h-3.5 text-[var(--control-muted)] shrink-0" />
            <span>Resize Column...</span>
          </button>
          <div className="my-1 border-t border-[var(--chrome-border)]" />
          <button
            type="button"
            role="menuitem"
            disabled={isBusy || !menu.canDelete}
            onClick={() => onAction('col_delete')}
            className={dangerItemClass}
          >
            <Trash2 className="w-3.5 h-3.5 shrink-0" />
            <span>Delete Column</span>
          </button>
        </>
      )}

      {menu.kind === 'tab' && (
        <>
          <button
            type="button"
            role="menuitem"
            disabled={isBusy || menu.isDeletingThis}
            onClick={() => onAction('tab_rename')}
            className={itemClass}
          >
            <Pencil className="w-3.5 h-3.5 text-[var(--control-muted)] shrink-0" />
            <span>Rename</span>
          </button>
          {(menu.canMoveLeft || menu.canMoveRight) && (
            <div className="flex items-center border-t border-[var(--chrome-border)]">
              <button
                type="button"
                role="menuitem"
                disabled={isBusy || !menu.canMoveLeft}
                onClick={() => onAction('tab_move_left')}
                className="flex-1 px-2.5 py-2 sm:py-1.5 hover:bg-[var(--control-bg-hover)] disabled:opacity-35 flex items-center justify-center gap-1 text-[11px] font-medium cursor-pointer touch-manipulation"
              >
                <ChevronLeft className="w-3.5 h-3.5" />
                <span>Move Left</span>
              </button>
              <div className="w-px h-5 bg-[var(--chrome-border)]" />
              <button
                type="button"
                role="menuitem"
                disabled={isBusy || !menu.canMoveRight}
                onClick={() => onAction('tab_move_right')}
                className="flex-1 px-2.5 py-2 sm:py-1.5 hover:bg-[var(--control-bg-hover)] disabled:opacity-35 flex items-center justify-center gap-1 text-[11px] font-medium cursor-pointer touch-manipulation"
              >
                <span>Move Right</span>
                <ChevronRight className="w-3.5 h-3.5" />
              </button>
            </div>
          )}
          <button
            type="button"
            role="menuitem"
            disabled={isBusy || !menu.canDelete}
            onClick={() => onAction('tab_delete')}
            className={`${dangerItemClass} border-t border-[var(--chrome-border)]`}
          >
            {menu.isDeletingThis ? (
              <>
                <Loader2 className="w-3.5 h-3.5 animate-spin shrink-0" />
                <span>Deleting...</span>
              </>
            ) : (
              <>
                <Trash2 className="w-3.5 h-3.5 shrink-0" />
                <span>Delete</span>
              </>
            )}
          </button>
        </>
      )}
    </div>
  );
}
