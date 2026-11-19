'use client';

import React, { useState, useEffect, useRef, useCallback, useMemo } from 'react';
import {
  LogOut,
  Sun,
  Moon,
  Loader2,
  HelpCircle,
  MoreVertical,
  Check,
  X,
  ChevronLeft,
  ChevronRight,
  RefreshCw,
  WifiOff
} from 'lucide-react';
import { User } from '@/lib/types';
import { ConfirmDialog } from '@/components/ConfirmDialog';
import { GuidedTour } from '@/components/GuidedTour';
import {
  SpreadsheetContextMenu,
  type ContextMenuTarget,
  type ContextMenuActionId
} from '@/components/SpreadsheetContextMenu';
import { ColumnWidthModal } from '@/components/ColumnWidthModal';
import { validateWorksheetName } from '@/lib/store';
import { exportSheetsToExcel } from '@/lib/excelUtils';
import { useTheme, type ThemeMode } from '@/hooks/useTheme';
import {
  getColumnVisualMeta,
  ColumnVisualMeta
} from '@/lib/clinicalColumns';
import {
  KNOWN_DOCTORS,
  KNOWN_TEST_TYPES,
  KNOWN_STAFF
} from '@/lib/data/initialData';

export interface SheetState {
  id?: string;
  name?: string;
  tabOrder?: number;
  headers: string[];
  headerVersion?: number;
  rows: (string | number)[][];
  rowVersions?: Record<number, number>;
  cellVersions?: Record<number, Record<number, number>>;
  rowUpdatedAt?: Record<number, string>;
  columnWidths?: Record<number, number>;
  rowHeights?: Record<number, number>;
}

interface DirtyCellEdit {
  sheetId: string;
  sheetName: string;
  rowIndex: number;
  colIndex: number;
  value: string | number;
  originalValue: string | number;
  expectedCellVersion: number;
}

interface DirtyHeaderEdit {
  sheetId: string;
  sheetName: string;
  colIndex: number;
  newHeader: string;
  originalHeader: string;
  expectedHeaderVersion: number;
}

interface CellDiffEntry {
  row: number;
  col: number;
  before: string | number;
  after: string | number;
}

interface HistoryStep {
  diffs: CellDiffEntry[];
  selectionBefore: SelectionState;
  selectionAfter: SelectionState;
}

interface SheetLayoutSnapshot {
  cols: Record<number, number>;
  rows: Record<number, number>;
}

const MAX_UNDO_HISTORY = 100;

type SelectionMode = 'cell' | 'range' | 'row' | 'col' | 'all';

interface SelectionState {
  mode: SelectionMode;
  anchorRow: number;
  anchorCol: number;
  focusRow: number;
  focusCol: number;
}

export interface ExcelSpreadsheetProps {
  workspace: 'reception' | 'doctor';
  user?: User;
  allowedWorkspaces?: ('reception' | 'doctor')[];
  initialSheets: Record<string, SheetState>;
  activeWorkspace: 'reception' | 'doctor';
  onWorkspaceSwitchRequest: (targetWorkspace: 'reception' | 'doctor') => void;
  onUnsavedStatusChange: (hasUnsaved: boolean) => void;
  onReloadWorkspace?: () => Promise<void>;
  onLogout: () => Promise<void> | void;
  theme?: ThemeMode;
  onToggleTheme?: () => void;
  isLoading?: boolean;
  loadError?: string | null;
  defaultSheetName?: string;
  canAdminImport?: boolean;
  onRequestInitialImport?: () => void;
  importNotice?: { type: 'success' | 'info' | 'error'; message: string } | null;
  onDismissImportNotice?: () => void;
}

const ROW_HEADER_COL_WIDTH = 50;
const CELL_GUTTER = 2;
const DEFAULT_COL_WIDTH = 112;
const MIN_COL_WIDTH = 56;
const MAX_COL_WIDTH = 600;
const DEFAULT_DATA_ROW_HEIGHT = 42;
const DEFAULT_HEADER_ROW_HEIGHT = 44;
const DEFAULT_COL_LETTER_ROW_HEIGHT = 28;
const MIN_ROW_HEIGHT = 28;
const MAX_ROW_HEIGHT = 320;
const MAX_COLS_LIMIT = 256;

function getColumnLetter(colIndex: number): string {
  let temp = colIndex;
  let letter = '';
  while (temp >= 0) {
    letter = String.fromCharCode((temp % 26) + 65) + letter;
    temp = Math.floor(temp / 26) - 1;
  }
  return letter;
}

function isNonEmptyCell(val: unknown): boolean {
  return val !== '' && val !== null && val !== undefined;
}

function getAlignmentClass(align: 'left' | 'center' | 'right'): string {
  if (align === 'center') return 'text-center';
  if (align === 'right') return 'text-right';
  return 'text-left';
}

function getUniformHeaderClasses(
  isSelected: boolean,
  isDirty: boolean
): string {
  if (isSelected) {
    return 'bg-[var(--col-title-selected)] text-[var(--col-title-text)]';
  }
  if (isDirty) {
    return 'bg-[var(--cell-dirty-bg)] text-[var(--cell-text)]';
  }
  return 'bg-[var(--col-title-bg)] hover:bg-[var(--col-title-hover)] text-[var(--col-title-text)]';
}

function getUniformCellBgClass(
  isSelected: boolean,
  isDirty: boolean
): string {
  if (isSelected) {
    return 'bg-[var(--cell-selected-bg)] text-[var(--cell-text)]';
  }
  if (isDirty) {
    return 'bg-[var(--cell-dirty-bg)] text-[var(--cell-text)]';
  }
  return 'bg-[var(--cell-bg)] hover:bg-[var(--cell-hover)] text-[var(--cell-text)]';
}

const SpreadsheetThemeToggle = React.memo(function SpreadsheetThemeToggle({
  overrideTheme,
  onOverrideToggle
}: {
  overrideTheme?: ThemeMode;
  onOverrideToggle?: () => void;
}) {
  const { theme: storedTheme, toggleTheme } = useTheme();
  const activeTheme = overrideTheme ?? storedTheme;
  const handleToggle = onOverrideToggle ?? toggleTheme;
  const isDark = activeTheme === 'dark';

  return (
    <button
      type="button"
      onClick={handleToggle}
      aria-label={isDark ? 'Switch to Light Mode' : 'Switch to Dark Mode'}
      title={isDark ? 'Switch to Light Mode' : 'Switch to Dark Mode'}
      className="p-1.5 bg-transparent border-0 shadow-none text-slate-700 hover:text-slate-950 dark:text-amber-300 dark:hover:text-amber-200 hover:scale-110 transition-transform duration-75 focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-[var(--selection-ring)] rounded-sm cursor-pointer shrink-0 flex items-center justify-center"
    >
      {isDark ? <Moon className="w-[18px] h-[18px]" /> : <Sun className="w-[18px] h-[18px]" />}
    </button>
  );
});

function normalizeHeaderKey(header: unknown): string {
  return String(header ?? '')
    .trim()
    .toLowerCase()
    .replace(/[\s_]+/g, ' ');
}

function parseOptionalFeeValue(raw: unknown): {
  empty: boolean;
  invalid: boolean;
  num: number;
} {
  if (raw === '' || raw === null || raw === undefined) {
    return { empty: true, invalid: false, num: 0 };
  }
  const str = String(raw).trim().replace(/,/g, '');
  if (str === '') {
    return { empty: true, invalid: false, num: 0 };
  }
  if (!/^-?\d+(\.\d+)?$/.test(str)) {
    return { empty: false, invalid: true, num: 0 };
  }
  const num = Number(str);
  if (!Number.isFinite(num)) {
    return { empty: false, invalid: true, num: 0 };
  }
  return { empty: false, invalid: false, num };
}

function remapMovedIndex(
  oldIdx: number,
  startIdx: number,
  endIdx: number,
  targetIdx: number
): number {
  const count = endIdx - startIdx + 1;
  if (targetIdx < startIdx) {
    if (oldIdx >= startIdx && oldIdx <= endIdx) return oldIdx - (startIdx - targetIdx);
    if (oldIdx >= targetIdx && oldIdx < startIdx) return oldIdx + count;
    return oldIdx;
  }
  if (oldIdx >= startIdx && oldIdx <= endIdx) return oldIdx + (targetIdx - endIdx - 1);
  if (oldIdx > endIdx && oldIdx < targetIdx) return oldIdx - count;
  return oldIdx;
}

/**
 * Determine which grid row (1 = Excel Row 2, i.e. dataRowIdx = 0) holds the latest
 * chronological record in row order (ignoring trailing empty rows and historical edit timestamps):
 * 1. Highest populated dataRowIdx scanning backwards from the bottom of the sheet.
 * 2. Fallback to 1 (first data row) when the worksheet has no populated rows.
 */
function findLatestRecordGridRow(sheet: SheetState | undefined): number {
  if (!sheet || !Array.isArray(sheet.rows) || sheet.rows.length === 0) {
    return 1;
  }

  for (let dataRowIdx = sheet.rows.length - 1; dataRowIdx >= 0; dataRowIdx--) {
    const row = sheet.rows[dataRowIdx];
    if (Array.isArray(row) && row.some(isNonEmptyCell)) {
      return dataRowIdx + 1;
    }
  }

  return 1;
}

function parseClipboardTsv(text: string): string[][] {
  const normalized = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
  const lines = normalized.split('\n');
  // Trim single trailing newline typically appended by spreadsheets/browsers
  if (lines.length > 1 && lines[lines.length - 1] === '') {
    lines.pop();
  }
  return lines.map((line) => line.split('\t'));
}

function isTextInputElement(target: EventTarget | null): boolean {
  if (!(target instanceof HTMLElement)) return false;
  const tag = target.tagName;
  if (tag === 'INPUT' || tag === 'TEXTAREA' || tag === 'SELECT') return true;
  if (target.isContentEditable) return true;
  if (target.closest('[role="dialog"]')) return true;
  return false;
}

function areNumberMapsEqual(
  a: Record<number, number> | undefined,
  b: Record<number, number> | undefined
): boolean {
  const keysA = a ? Object.keys(a) : [];
  const keysB = b ? Object.keys(b) : [];
  if (keysA.length !== keysB.length) return false;
  for (const k of keysA) {
    const idx = Number(k);
    if (a?.[idx] !== b?.[idx]) return false;
  }
  return true;
}

function extractLayoutsFromSheets(sheetMap: Record<string, SheetState>): {
  colWidths: Record<string, Record<number, number>>;
  rowHeights: Record<string, Record<number, number>>;
  snapshots: Record<string, SheetLayoutSnapshot>;
} {
  const colWidths: Record<string, Record<number, number>> = {};
  const rowHeights: Record<string, Record<number, number>> = {};
  const snapshots: Record<string, SheetLayoutSnapshot> = {};

  for (const [name, sheet] of Object.entries(sheetMap)) {
    const dimKey = sheet.id || name;
    const cols = { ...(sheet.columnWidths || {}) };
    const rows = { ...(sheet.rowHeights || {}) };
    if (Object.keys(cols).length > 0) colWidths[dimKey] = cols;
    if (Object.keys(rows).length > 0) rowHeights[dimKey] = rows;
    snapshots[dimKey] = { cols, rows };
  }

  return { colWidths, rowHeights, snapshots };
}

interface SpreadsheetDataRowProps {
  gridRowIdx: number;
  rowCells: (string | number)[] | undefined;
  totalCols: number;
  rowHeaderWidth: number;
  colMetaList: ColumnVisualMeta[];
  manualHeight: number | undefined;
  isWrapEnabled: boolean;
  isRowHighlighted: boolean;
  isFullRowMode: boolean;
  anchorCol: number;
  isEditingRow: boolean;
  editValueIfEditing: string;
  editInputListId?: string;
  selectedColStart: number;
  selectedColEnd: number;
  dirtyColsKey?: string;
  workspace: 'reception' | 'doctor';
  paidColIdx: number;
  onRowHeaderClick: (gridRowIdx: number, shiftKey: boolean) => void;
  onRowHeaderContextMenu: (gridRowIdx: number, x: number, y: number) => void;
  onRowHeaderTouchStart: (gridRowIdx: number, e: React.PointerEvent<HTMLTableCellElement>) => void;
  onRowDragStart: (gridRowIdx: number, e: React.PointerEvent<HTMLTableCellElement>) => void;
  onRowResizeStart: (gridRowIdx: number, e: React.PointerEvent<HTMLDivElement>) => void;
  onCellClick: (gridRowIdx: number, colIdx: number, shiftKey: boolean, detail?: number) => void;
  onCellDoubleClick: (gridRowIdx: number, colIdx: number, detail: number) => void;
  onCellContextMenu: (gridRowIdx: number, colIdx: number, x: number, y: number) => void;
  onCellTouchStart: (gridRowIdx: number, colIdx: number, e: React.PointerEvent<HTMLTableCellElement>) => void;
  onPaidValueChange: (gridRowIdx: number, colIdx: number, val: string) => void;
  onEditValueChange: (val: string) => void;
  onCommitEdit: () => void;
  setEditInputRef: (el: HTMLInputElement | null) => void;
  setRowElementRef: (gridRowIdx: number, el: HTMLTableRowElement | null) => void;
}

const SpreadsheetDataRow = React.memo(function SpreadsheetDataRow({
  gridRowIdx,
  rowCells,
  totalCols,
  rowHeaderWidth,
  colMetaList,
  manualHeight,
  isWrapEnabled,
  isRowHighlighted,
  isFullRowMode,
  anchorCol,
  isEditingRow,
  editValueIfEditing,
  editInputListId,
  selectedColStart,
  selectedColEnd,
  dirtyColsKey,
  workspace,
  paidColIdx,
  onRowHeaderClick,
  onRowHeaderContextMenu,
  onRowHeaderTouchStart,
  onRowDragStart,
  onRowResizeStart,
  onCellClick,
  onCellDoubleClick,
  onCellContextMenu,
  onCellTouchStart,
  onPaidValueChange,
  onEditValueChange,
  onCommitEdit,
  setEditInputRef,
  setRowElementRef
}: SpreadsheetDataRowProps) {
  const excelRowNum = gridRowIdx + 1;
  const dirtyColSet = useMemo(() => {
    if (!dirtyColsKey) return null;
    return new Set(dirtyColsKey.split(',').map((n) => Number(n)));
  }, [dirtyColsKey]);

  const effectiveRowHeight = manualHeight ?? DEFAULT_DATA_ROW_HEIGHT;

  return (
    <tr
      ref={(el) => setRowElementRef(gridRowIdx, el)}
      style={{ height: `${effectiveRowHeight}px` }}
    >
      <th
        onClick={(e) => onRowHeaderClick(gridRowIdx, e.shiftKey)}
        onPointerDown={(e) => {
          onRowHeaderTouchStart(gridRowIdx, e);
          onRowDragStart(gridRowIdx, e);
        }}
        onContextMenu={(e) => {
          e.preventDefault();
          e.stopPropagation();
          onRowHeaderContextMenu(gridRowIdx, e.clientX, e.clientY);
        }}
        style={{ width: `${rowHeaderWidth}px`, minWidth: `${rowHeaderWidth}px` }}
        title="Click to select row, drag to reorder row, right-click for options"
        className={`px-1 text-center text-[11px] tabular-nums align-middle sticky left-0 z-[15] select-none touch-manipulation rounded-l-[8px] rounded-r-none shadow-[-4px_0_0_0_var(--grid-canvas)] ${
          isRowHighlighted && isFullRowMode
            ? 'bg-[var(--header-tile-active)] text-[var(--header-text-active)] font-semibold cursor-grab active:cursor-grabbing'
            : isRowHighlighted
              ? 'bg-[var(--header-tile-range)] text-[var(--col-title-text)] font-semibold cursor-pointer'
              : 'bg-[var(--header-tile-bg)] hover:bg-[var(--header-tile-hover)] text-[var(--header-text)] font-medium cursor-pointer'
        }`}
      >
        <span>{excelRowNum}</span>
        <div
          onPointerDown={(e) => onRowResizeStart(gridRowIdx, e)}
          onClick={(e) => e.stopPropagation()}
          onDoubleClick={(e) => e.stopPropagation()}
          title="Drag to resize row height"
          className="absolute left-1 right-0 bottom-0 h-3 md:h-1.5 cursor-row-resize touch-none hover:bg-[var(--resize-accent)]/70 rounded-full z-20"
        />
      </th>

      {Array.from({ length: totalCols }).map((_, cIdx) => {
        const isAnchor = anchorCol === cIdx;
        const isInSel =
          selectedColStart !== -1 && cIdx >= selectedColStart && cIdx <= selectedColEnd;
        const isPaidDropdownCell = workspace === 'reception' && paidColIdx === cIdx;
        const isCellEditing = isAnchor && isEditingRow && !isPaidDropdownCell;
        const isDirty = dirtyColSet ? dirtyColSet.has(cIdx) : false;
        const cellVal = rowCells ? (rowCells[cIdx] ?? '') : '';
        const strVal = cellVal !== undefined && cellVal !== null ? String(cellVal) : '';
        const colMeta = colMetaList[cIdx];
        const alignClass = getAlignmentClass(colMeta?.align ?? 'left');
        const showRangeTint =
          isInSel && !(isAnchor && !isFullRowMode && selectedColStart === selectedColEnd);
        const bgClass = getUniformCellBgClass(showRangeTint, isDirty);
        const hasCustomHistoricalPaidVal =
          isPaidDropdownCell && strVal !== '' && strVal !== 'Paid' && strVal !== 'Unpaid';
        const paidToneClass =
          strVal.trim().toLowerCase() === 'paid'
            ? 'text-emerald-700 dark:text-emerald-300 font-semibold'
            : strVal.trim().toLowerCase() === 'unpaid'
              ? 'text-amber-700 dark:text-amber-300 font-semibold'
              : 'text-[var(--cell-text)]';

        return (
          <td
            key={cIdx}
            onClick={(e) => onCellClick(gridRowIdx, cIdx, e.shiftKey, e.detail)}
            onDoubleClick={(e) => {
              if (!isPaidDropdownCell) {
                onCellDoubleClick(gridRowIdx, cIdx, e.detail);
              }
            }}
            onPointerDown={(e) => onCellTouchStart(gridRowIdx, cIdx, e)}
            onContextMenu={(e) => {
              e.preventDefault();
              e.stopPropagation();
              onCellContextMenu(gridRowIdx, cIdx, e.clientX, e.clientY);
            }}
            className={`${isPaidDropdownCell ? 'px-1 py-0.5' : 'px-2 sm:px-3 py-2'} relative text-xs leading-[20px] font-normal tabular-nums align-middle rounded-[5px] touch-manipulation ${alignClass} ${bgClass} ${
              isAnchor && !isCellEditing
                ? 'ring-2 ring-inset ring-[var(--selection-ring)] z-10'
                : ''
            }`}
          >
            {isDirty && !isCellEditing && (
              <span
                className="pointer-events-none absolute top-1 right-1 w-1.5 h-1.5 rounded-full bg-amber-500 dark:bg-amber-300 z-20"
                title="Unsaved edit"
              />
            )}
            {isPaidDropdownCell ? (
              <select
                tabIndex={-1}
                aria-label={`Payment status row ${excelRowNum}`}
                value={strVal}
                onChange={(e) => {
                  onPaidValueChange(gridRowIdx, cIdx, e.target.value);
                  e.currentTarget.blur();
                }}
                onDoubleClick={(e) => e.stopPropagation()}
                className={`w-full h-[30px] px-1.5 bg-transparent text-xs rounded-[4px] outline-none cursor-pointer ${paidToneClass}`}
              >
                <option value="" className="bg-[var(--chrome-bg)] text-[var(--control-text)]">
                  {' '}
                </option>
                <option value="Paid" className="bg-[var(--chrome-bg)] text-[var(--control-text)]">
                  Paid
                </option>
                <option value="Unpaid" className="bg-[var(--chrome-bg)] text-[var(--control-text)]">
                  Unpaid
                </option>
                {hasCustomHistoricalPaidVal && (
                  <option
                    value={strVal}
                    className="bg-[var(--chrome-bg)] text-[var(--control-text)]"
                  >
                    {strVal}
                  </option>
                )}
              </select>
            ) : isCellEditing ? (
              <input
                ref={setEditInputRef}
                type="text"
                list={editInputListId}
                value={editValueIfEditing}
                onChange={(e) => onEditValueChange(e.target.value)}
                onBlur={onCommitEdit}
                className={`absolute inset-0 w-full h-full px-2 sm:px-3 text-base sm:text-xs rounded-[5px] bg-[var(--formula-input-bg)] text-[var(--control-text)] border-2 border-[var(--selection-ring)] outline-none z-10 ${alignClass}`}
              />
            ) : (
              <div
                className={
                  isWrapEnabled
                    ? 'whitespace-normal break-words leading-[20px] min-h-[20px]'
                    : 'truncate whitespace-nowrap leading-[20px] h-[20px]'
                }
                style={
                  manualHeight
                    ? { maxHeight: `${Math.max(20, manualHeight - 16)}px`, overflowY: 'auto' }
                    : undefined
                }
              >
                {strVal}
              </div>
            )}
          </td>
        );
      })}
    </tr>
  );
});

export function ExcelSpreadsheet({
  workspace,
  user,
  allowedWorkspaces,
  initialSheets,
  activeWorkspace,
  onWorkspaceSwitchRequest,
  onUnsavedStatusChange,
  onReloadWorkspace,
  onLogout,
  theme,
  onToggleTheme,
  isLoading = false,
  loadError = null,
  defaultSheetName,
  canAdminImport = false,
  onRequestInitialImport,
  importNotice = null,
  onDismissImportNotice
}: ExcelSpreadsheetProps) {
  const initialTabName = useMemo(() => {
    if (defaultSheetName && initialSheets[defaultSheetName]) return defaultSheetName;
    const keys = Object.keys(initialSheets).sort((a, b) => {
      const orderA = initialSheets[a]?.tabOrder ?? Number.MAX_SAFE_INTEGER;
      const orderB = initialSheets[b]?.tabOrder ?? Number.MAX_SAFE_INTEGER;
      if (orderA !== orderB) return orderA - orderB;
      return a.localeCompare(b);
    });
    if (keys.length > 0) return keys[0];
    return '';
  }, [defaultSheetName, initialSheets]);

  const initialLatestRow = useMemo(
    () => findLatestRecordGridRow(initialTabName ? initialSheets[initialTabName] : undefined),
    [initialSheets, initialTabName]
  );

  const [activeSheetName, setActiveSheetName] = useState<string>(initialTabName);
  const [sheets, setSheets] = useState<Record<string, SheetState>>(initialSheets);

  // Dirty edit tracking keyed by coordinate; refs mirror state for immediate Save / Ctrl+S access
  const [dirtyCells, setDirtyCells] = useState<Record<string, DirtyCellEdit>>({});
  const [dirtyHeaders, setDirtyHeaders] = useState<Record<string, DirtyHeaderEdit>>({});
  const dirtyCellsRef = useRef<Record<string, DirtyCellEdit>>({});
  const dirtyHeadersRef = useRef<Record<string, DirtyHeaderEdit>>({});

  const [saveStatus, setSaveStatus] = useState<'idle' | 'pending' | 'saving' | 'saved' | 'error'>('idle');
  const [saveMessage, setSaveMessage] = useState<string>('');
  const [saveError, setSaveError] = useState<string | null>(null);
  const [hasConflict, setHasConflict] = useState<boolean>(false);
  const [isOnline, setIsOnline] = useState<boolean>(true);
  const isSavingRef = useRef<boolean>(false);
  const saveStatusRef = useRef<'idle' | 'pending' | 'saving' | 'saved' | 'error'>('idle');
  const hasConflictRef = useRef<boolean>(false);
  const autosaveTimerRef = useRef<number | null>(null);
  const retryTimerRef = useRef<number | null>(null);
  const retryAttemptRef = useRef<number>(0);
  const flushAgainAfterFlightRef = useRef<boolean>(false);
  const inFlightSavePromiseRef = useRef<Promise<boolean> | null>(null);
  const flushSaveQueueRef = useRef<(opts?: { source?: 'autosave' | 'manual' | 'retry' | 'export' | 'structural' }) => Promise<boolean>>(() => Promise.resolve(true));

  // Selection state (supports cell, range, row(s), column(s), all)
  const [selection, setSelection] = useState<SelectionState>({
    mode: 'cell',
    anchorRow: initialLatestRow,
    anchorCol: 0,
    focusRow: initialLatestRow,
    focusCol: 0
  });
  const selectedCell = useMemo(
    () => ({ row: selection.anchorRow, col: selection.anchorCol }),
    [selection.anchorRow, selection.anchorCol]
  );

  // Latest-record scroll positioning generation (only increments on initial data load or sheet open/switch, never on save)
  const [scrollPositionRequest, setScrollPositionRequest] = useState<{
    sheetName: string;
    targetRow: number;
    nonce: number;
  }>({
    sheetName: initialTabName,
    targetRow: initialLatestRow,
    nonce: initialTabName ? 1 : 0
  });
  const lastExecutedScrollNonceRef = useRef<number>(0);



  const [isEditing, setIsEditing] = useState<boolean>(false);
  const [editValue, setEditValue] = useState<string>('');
  const [isWrapEnabled, setIsWrapEnabled] = useState<boolean>(true);

  // Per-sheet persisted column widths & manual row heights
  const initialLayouts = useMemo(() => extractLayoutsFromSheets(initialSheets), [initialSheets]);
  const [colWidthsBySheet, setColWidthsBySheet] = useState<Record<string, Record<number, number>>>(
    () => initialLayouts.colWidths
  );
  const [rowHeightsBySheet, setRowHeightsBySheet] = useState<Record<string, Record<number, number>>>(
    () => initialLayouts.rowHeights
  );
  const colWidthsBySheetRef = useRef<Record<string, Record<number, number>>>(initialLayouts.colWidths);
  const rowHeightsBySheetRef = useRef<Record<string, Record<number, number>>>(initialLayouts.rowHeights);
  const [savedLayoutBySheet, setSavedLayoutBySheet] = useState<Record<string, SheetLayoutSnapshot>>(
    () => initialLayouts.snapshots
  );
  const savedLayoutBySheetRef = useRef<Record<string, SheetLayoutSnapshot>>(initialLayouts.snapshots);

  // Per-sheet undo/redo stacks (up to 100 steps per worksheet, keyed by sheetId || sheetName)
  const undoStackBySheetRef = useRef<Record<string, HistoryStep[]>>({});
  const redoStackBySheetRef = useRef<Record<string, HistoryStep[]>>({});

  // Synchronize when initialSheets prop updates from Supabase fetch / reload
  const [prevInitialSheets, setPrevInitialSheets] = useState(initialSheets);
  if (initialSheets !== prevInitialSheets) {
    setPrevInitialSheets(initialSheets);
    setSheets(initialSheets);
    setDirtyCells({});
    setDirtyHeaders({});
    setSaveStatus('idle');
    setSaveMessage('');
    setSaveError(null);
    setHasConflict(false);
    setIsEditing(false);

    const freshLayouts = extractLayoutsFromSheets(initialSheets);
    setColWidthsBySheet(freshLayouts.colWidths);
    setRowHeightsBySheet(freshLayouts.rowHeights);
    setSavedLayoutBySheet(freshLayouts.snapshots);

    const keys = Object.keys(initialSheets).sort((a, b) => {
      const orderA = initialSheets[a]?.tabOrder ?? Number.MAX_SAFE_INTEGER;
      const orderB = initialSheets[b]?.tabOrder ?? Number.MAX_SAFE_INTEGER;
      if (orderA !== orderB) return orderA - orderB;
      return a.localeCompare(b);
    });
    let nextActive = activeSheetName;
    if (!nextActive || !initialSheets[nextActive]) {
      nextActive =
        defaultSheetName && initialSheets[defaultSheetName]
          ? defaultSheetName
          : keys[0] || '';
      setActiveSheetName(nextActive);
    }

    const latestRow = findLatestRecordGridRow(nextActive ? initialSheets[nextActive] : undefined);
    setSelection({
      mode: 'cell',
      anchorRow: latestRow,
      anchorCol: 0,
      focusRow: latestRow,
      focusCol: 0
    });
    if (nextActive) {
      setScrollPositionRequest((prev) => ({
        sheetName: nextActive,
        targetRow: latestRow,
        nonce: prev.nonce + 1
      }));
    }
  }

  useEffect(() => {
    dirtyCellsRef.current = {};
    dirtyHeadersRef.current = {};
    undoStackBySheetRef.current = {};
    redoStackBySheetRef.current = {};
  }, [initialSheets]);

  useEffect(() => {
    dirtyCellsRef.current = dirtyCells;
  }, [dirtyCells]);

  useEffect(() => {
    dirtyHeadersRef.current = dirtyHeaders;
  }, [dirtyHeaders]);

  useEffect(() => {
    colWidthsBySheetRef.current = colWidthsBySheet;
  }, [colWidthsBySheet]);

  useEffect(() => {
    rowHeightsBySheetRef.current = rowHeightsBySheet;
  }, [rowHeightsBySheet]);

  useEffect(() => {
    savedLayoutBySheetRef.current = savedLayoutBySheet;
  }, [savedLayoutBySheet]);

  const dirtyLayoutEntries = useMemo(() => {
    const entries: Array<{
      dimKey: string;
      sheetId: string;
      sheetName: string;
      columnWidths: Record<number, number>;
      rowHeights: Record<number, number>;
    }> = [];

    for (const [sheetName, sheet] of Object.entries(sheets)) {
      if (!sheet.id) continue;
      const dimKey = sheet.id;
      const curCols = colWidthsBySheet[dimKey] || {};
      const curRows = rowHeightsBySheet[dimKey] || {};
      const savedSnap = savedLayoutBySheet[dimKey] || { cols: {}, rows: {} };
      if (
        !areNumberMapsEqual(curCols, savedSnap.cols) ||
        !areNumberMapsEqual(curRows, savedSnap.rows)
      ) {
        entries.push({
          dimKey,
          sheetId: sheet.id,
          sheetName,
          columnWidths: curCols,
          rowHeights: curRows
        });
      }
    }
    return entries;
  }, [sheets, colWidthsBySheet, rowHeightsBySheet, savedLayoutBySheet]);

  const hasUnsavedChanges =
    Object.keys(dirtyCells).length > 0 ||
    Object.keys(dirtyHeaders).length > 0 ||
    dirtyLayoutEntries.length > 0;

  useEffect(() => {
    onUnsavedStatusChange(hasUnsavedChanges || saveStatus === 'saving' || saveStatus === 'error');
  }, [hasUnsavedChanges, saveStatus, onUnsavedStatusChange]);

  useEffect(() => {
    saveStatusRef.current = saveStatus;
    hasConflictRef.current = hasConflict;
  }, [saveStatus, hasConflict]);

  const justResizedRef = useRef<boolean>(false);
  const resizeCleanupRef = useRef<(() => void) | null>(null);

  useEffect(() => {
    return () => {
      if (resizeCleanupRef.current) {
        resizeCleanupRef.current();
      }
    };
  }, []);

  const rootRef = useRef<HTMLDivElement>(null);
  const gridContainerRef = useRef<HTMLDivElement>(null);
  const tableWrapperRef = useRef<HTMLDivElement>(null);
  const tableRef = useRef<HTMLTableElement>(null);
  const theadRef = useRef<HTMLTableSectionElement>(null);
  const tbodyRef = useRef<HTMLTableSectionElement>(null);
  const editInputRef = useRef<HTMLInputElement | null>(null);
  const rowDropLineRef = useRef<HTMLDivElement | null>(null);
  const colDropLineRef = useRef<HTMLDivElement | null>(null);
  const reorderDragCleanupRef = useRef<(() => void) | null>(null);

  // Direct DOM refs for fast rAF drag resizing without per-frame React re-renders
  const colElementsRef = useRef<Record<number, HTMLTableColElement | null>>({});
  const headerLetterThRefs = useRef<Record<number, HTMLTableCellElement | null>>({});
  const headerTitleTdRefs = useRef<Record<number, HTMLTableCellElement | null>>({});
  const rowElementsRef = useRef<Record<number, HTMLTableRowElement | null>>({});
  const headerRowThRef = useRef<HTMLTableCellElement | null>(null);

  // Worksheet management state: immediate add, inline rename, right-click context menu, pointer drag & drop reorder
  const [isCreatingSheet, setIsCreatingSheet] = useState(false);
  const [isReloading, setIsReloading] = useState(false);
  const [isLoggingOut, setIsLoggingOut] = useState(false);
  const [operationError, setOperationError] = useState<string | null>(null);
  const [renamingSheetName, setRenamingSheetName] = useState<string | null>(null);
  const [renameInputValue, setRenameInputValue] = useState<string>('');
  const [, setIsRenamingSheet] = useState(false);
  const [contextMenu, setContextMenu] = useState<ContextMenuTarget | null>(null);
  const [resizeColDialog, setResizeColDialog] = useState<{
    colIndex: number;
    currentWidth: number;
    defaultWidth: number;
  } | null>(null);
  const [isStructuralBusy, setIsStructuralBusy] = useState(false);
  const [pendingConfirmAction, setPendingConfirmAction] = useState<{
    title: string;
    message: string;
    confirmLabel: string;
    confirmVariant: 'danger' | 'primary';
    run: () => void;
  } | null>(null);
  const internalClipboardRef = useRef<{
    text: string;
    kind: 'cells' | 'row' | 'col';
  } | null>(null);
  const [isDeletingSheet, setIsDeletingSheet] = useState(false);
  const [deletingSheetName, setDeletingSheetName] = useState<string | null>(null);
  const [dragTabState, setDragTabState] = useState<{
    sourceName: string;
    targetIndex: number;
  } | null>(null);
  const tabElementRefs = useRef<Record<string, HTMLDivElement | null>>({});
  const tabDragCleanupRef = useRef<(() => void) | null>(null);
  const gridLongPressCleanupRef = useRef<(() => void) | null>(null);
  const suppressNextTabClickRef = useRef<boolean>(false);
  const renameInputRef = useRef<HTMLInputElement | null>(null);
  const renameCommitGuardRef = useRef<boolean>(false);
  const [pendingBulkClear, setPendingBulkClear] = useState<{
    updates: Array<{ row: number; col: number; value: string | number }>;
    populatedCount: number;
    includesHeader: boolean;
  } | null>(null);

  // Guided Walkthrough Tour state
  const [isTourOpen, setIsTourOpen] = useState(false);
  const hasAutoCheckedTourRef = useRef(false);

  // Responsive viewport compactness (narrows sticky row-number column slightly below 640px)
  const [isCompactViewport, setIsCompactViewport] = useState(false);
  const lastTouchTapRef = useRef<{ row: number; col: number; time: number } | null>(null);

  useEffect(() => {
    const mq = window.matchMedia('(max-width: 639px)');
    const syncCompact = () => setIsCompactViewport(mq.matches);
    const rafId = window.requestAnimationFrame(syncCompact);
    mq.addEventListener('change', syncCompact);
    return () => {
      window.cancelAnimationFrame(rafId);
      mq.removeEventListener('change', syncCompact);
    };
  }, []);

  const rowHeaderWidth = isCompactViewport ? 40 : ROW_HEADER_COL_WIDTH;

  // Focus & select existing tab name immediately whenever rename mode opens
  useEffect(() => {
    if (renamingSheetName !== null && renameInputRef.current) {
      renameInputRef.current.focus();
      renameInputRef.current.select();
    }
  }, [renamingSheetName]);

  useEffect(() => {
    return () => {
      if (tabDragCleanupRef.current) {
        tabDragCleanupRef.current();
      }
      if (gridLongPressCleanupRef.current) {
        gridLongPressCleanupRef.current();
      }
      if (reorderDragCleanupRef.current) {
        reorderDragCleanupRef.current();
      }
      if (autosaveTimerRef.current !== null) {
        window.clearTimeout(autosaveTimerRef.current);
      }
      if (retryTimerRef.current !== null) {
        window.clearTimeout(retryTimerRef.current);
      }
    };
  }, []);

  const sheetKeys = useMemo(() => {
    return Object.keys(sheets).sort((a, b) => {
      const orderA = sheets[a]?.tabOrder ?? Number.MAX_SAFE_INTEGER;
      const orderB = sheets[b]?.tabOrder ?? Number.MAX_SAFE_INTEGER;
      if (orderA !== orderB) return orderA - orderB;
      return a.localeCompare(b);
    });
  }, [sheets]);

  const sheetKeysRef = useRef(sheetKeys);
  const sheetsRef = useRef(sheets);
  useEffect(() => {
    sheetKeysRef.current = sheetKeys;
    sheetsRef.current = sheets;
  });

  const currentSheet = useMemo(() => {
    return (
      sheets[activeSheetName] || {
        headers: [],
        rows: []
      }
    );
  }, [sheets, activeSheetName]);

  const maxRowCols = useMemo(() => {
    return currentSheet.rows.reduce((max, r) => Math.max(max, r ? r.length : 0), 0);
  }, [currentSheet.rows]);

  // Support wide worksheets (up to 52+ columns as in Lens Rx or Medications)
  const totalCols = Math.max(currentSheet.headers.length, maxRowCols, 26);
  const displayRowsCount = Math.max(currentSheet.rows.length + 30, 60);

  // Header-driven clinical classification, width, and alignment per column
  const colMetaList: ColumnVisualMeta[] = useMemo(() => {
    const list: ColumnVisualMeta[] = [];
    for (let c = 0; c < totalCols; c++) {
      list.push(
        getColumnVisualMeta(workspace, activeSheetName, currentSheet.headers[c] ?? '', c)
      );
    }
    return list;
  }, [workspace, activeSheetName, currentSheet.headers, totalCols]);

  // Dynamic header resolution so Reception Total calculation and Paid dropdown follow rearranged columns
  const receptionSpecialCols = useMemo(() => {
    let testFeeColIdx = -1;
    let medColIdx = -1;
    let totalColIdx = -1;
    let paidColIdx = -1;

    if (workspace === 'reception') {
      currentSheet.headers.forEach((h, idx) => {
        const norm = normalizeHeaderKey(h);
        if (norm === 'test fee' && testFeeColIdx === -1) testFeeColIdx = idx;
        else if (norm === 'medication' && medColIdx === -1) medColIdx = idx;
        else if (norm === 'total' && totalColIdx === -1) totalColIdx = idx;
        else if (norm === 'paid' && paidColIdx === -1) paidColIdx = idx;
      });
    }

    return { testFeeColIdx, medColIdx, totalColIdx, paidColIdx };
  }, [workspace, currentSheet.headers]);

  // Lightweight clinical autocomplete suggestions strictly when editing a data cell
  const autocompleteOptions = useMemo(() => {
    if (!isEditing || selectedCell.row === 0) return [];
    const colHeader = normalizeHeaderKey(currentSheet.headers[selectedCell.col] ?? '');
    const seen = new Set<string>();
    const pushUnique = (val: unknown) => {
      if (val === null || val === undefined) return;
      const s = String(val).trim();
      if (!s || seen.has(s)) return;
      seen.add(s);
    };

    if (colHeader === 'examined by' || colHeader === 'dic' || colHeader.includes('doctor')) {
      KNOWN_DOCTORS.forEach(pushUnique);
    } else if (colHeader === 'test type' || colHeader === 'review test') {
      KNOWN_TEST_TYPES.forEach(pushUnique);
    } else if (colHeader === 'recorded by' || colHeader.startsWith('recorded by (')) {
      KNOWN_STAFF.forEach(pushUnique);
    } else if (colHeader === 'paid') {
      pushUnique('Paid');
      pushUnique('Unpaid');
    }

    // Avoid cluttering phone or fee numeric inputs with historical column autocomplete
    const isNumericOrPhoneCol =
      colHeader.includes('phone') ||
      colHeader.includes('contact') ||
      colHeader.includes('fee') ||
      colHeader.includes('total') ||
      colHeader.includes('bal') ||
      colHeader.includes('medication') ||
      colHeader.includes('drops') ||
      colHeader === 'new' ||
      colHeader === 'renewal';

    if (!isNumericOrPhoneCol) {
      for (let r = currentSheet.rows.length - 1; r >= 0 && seen.size < 35; r--) {
        pushUnique(currentSheet.rows[r]?.[selectedCell.col]);
      }
    }

    return Array.from(seen).slice(0, 35);
  }, [isEditing, selectedCell.row, selectedCell.col, currentSheet.headers, currentSheet.rows]);

  const activeDimKey = currentSheet.id || activeSheetName || '__default__';
  const currentColWidths = colWidthsBySheet[activeDimKey];
  const currentRowHeights = rowHeightsBySheet[activeDimKey];

  const getColWidth = useCallback(
    (colIndex: number): number => {
      return (
        currentColWidths?.[colIndex] ??
        colMetaList[colIndex]?.defaultWidth ??
        DEFAULT_COL_WIDTH
      );
    },
    [currentColWidths, colMetaList]
  );

  const getManualRowHeight = useCallback(
    (gridRowIdx: number): number | undefined => {
      return currentRowHeights?.[gridRowIdx];
    },
    [currentRowHeights]
  );

  const totalTableWidth = useMemo(() => {
    let sum = rowHeaderWidth + (totalCols + 2) * CELL_GUTTER;
    for (let c = 0; c < totalCols; c++) {
      sum += getColWidth(c);
    }
    return sum;
  }, [totalCols, getColWidth, rowHeaderWidth]);

  // Fast per-row & header dirty edit lookups so React.memo row optimization stays intact
  const dirtyColsByRow = useMemo(() => {
    const rowColsMap = new Map<number, number[]>();
    const activeId = currentSheet.id;
    for (const edit of Object.values(dirtyCells)) {
      if (
        (activeId && edit.sheetId === activeId) ||
        (!activeId && edit.sheetName === activeSheetName)
      ) {
        const cols = rowColsMap.get(edit.rowIndex) || [];
        cols.push(edit.colIndex);
        rowColsMap.set(edit.rowIndex, cols);
      }
    }
    const serialized = new Map<number, string>();
    for (const [rIdx, cols] of rowColsMap.entries()) {
      cols.sort((a, b) => a - b);
      serialized.set(rIdx, cols.join(','));
    }
    return serialized;
  }, [dirtyCells, currentSheet.id, activeSheetName]);

  const dirtyHeaderCols = useMemo(() => {
    const cols = new Set<number>();
    const activeId = currentSheet.id;
    for (const edit of Object.values(dirtyHeaders)) {
      if (
        (activeId && edit.sheetId === activeId) ||
        (!activeId && edit.sheetName === activeSheetName)
      ) {
        cols.add(edit.colIndex);
      }
    }
    return cols;
  }, [dirtyHeaders, currentSheet.id, activeSheetName]);

  // Normalized selection bounds for highlighting, copy, and clear
  const selectionBounds = useMemo(() => {
    if (selection.mode === 'all') {
      return {
        rMin: 0,
        rMax: displayRowsCount - 1,
        cMin: 0,
        cMax: totalCols - 1
      };
    }
    if (selection.mode === 'row') {
      return {
        rMin: Math.min(selection.anchorRow, selection.focusRow),
        rMax: Math.max(selection.anchorRow, selection.focusRow),
        cMin: 0,
        cMax: totalCols - 1
      };
    }
    if (selection.mode === 'col') {
      return {
        rMin: 0,
        rMax: displayRowsCount - 1,
        cMin: Math.min(selection.anchorCol, selection.focusCol),
        cMax: Math.max(selection.anchorCol, selection.focusCol)
      };
    }
    return {
      rMin: Math.min(selection.anchorRow, selection.focusRow),
      rMax: Math.max(selection.anchorRow, selection.focusRow),
      cMin: Math.min(selection.anchorCol, selection.focusCol),
      cMax: Math.max(selection.anchorCol, selection.focusCol)
    };
  }, [selection, displayRowsCount, totalCols]);

  const isCellInSelection = useCallback(
    (r: number, c: number): boolean => {
      return (
        r >= selectionBounds.rMin &&
        r <= selectionBounds.rMax &&
        c >= selectionBounds.cMin &&
        c <= selectionBounds.cMax
      );
    },
    [selectionBounds]
  );

  const isRowHeaderHighlighted = useCallback(
    (r: number): boolean => {
      return r >= selectionBounds.rMin && r <= selectionBounds.rMax;
    },
    [selectionBounds]
  );

  const isColHeaderHighlighted = useCallback(
    (c: number): boolean => {
      return c >= selectionBounds.cMin && c <= selectionBounds.cMax;
    },
    [selectionBounds]
  );

  const selectionAddressLabel = useMemo(() => {
    if (selection.mode === 'all') return 'ALL';
    if (selection.mode === 'row') {
      return selectionBounds.rMin === selectionBounds.rMax
        ? `R${selectionBounds.rMin + 1}`
        : `R${selectionBounds.rMin + 1}:${selectionBounds.rMax + 1}`;
    }
    if (selection.mode === 'col') {
      return selectionBounds.cMin === selectionBounds.cMax
        ? `C:${getColumnLetter(selectionBounds.cMin)}`
        : `${getColumnLetter(selectionBounds.cMin)}:${getColumnLetter(selectionBounds.cMax)}`;
    }
    if (selection.mode === 'range') {
      return `${getColumnLetter(selectionBounds.cMin)}${selectionBounds.rMin + 1}:${getColumnLetter(
        selectionBounds.cMax
      )}${selectionBounds.rMax + 1}`;
    }
    return `${getColumnLetter(selection.anchorCol)}${selection.anchorRow + 1}`;
  }, [selection, selectionBounds]);

  // Auto-position scroll viewport to the latest record when a worksheet opens, reloads, or workspace switches
  const isActiveWorkspace = workspace === activeWorkspace;
  const wasActiveWorkspaceRef = useRef<boolean>(isActiveWorkspace);

  useEffect(() => {
    if (!isActiveWorkspace || isLoading) return;
    if (scrollPositionRequest.nonce === 0) return;
    if (scrollPositionRequest.sheetName !== activeSheetName) return;
    if (lastExecutedScrollNonceRef.current === scrollPositionRequest.nonce) return;

    const reqNonce = scrollPositionRequest.nonce;
    const targetRow = scrollPositionRequest.targetRow;
    let rafId1 = 0;
    let rafId2 = 0;

    rafId1 = window.requestAnimationFrame(() => {
      rafId2 = window.requestAnimationFrame(() => {
        const el = gridContainerRef.current;
        if (!el || el.clientHeight === 0) return;

        const tbodyEl = tbodyRef.current;
        const targetTr = tbodyEl?.children[targetRow - 1] as HTMLElement | undefined;
        if (targetRow > 1 && !targetTr) return;

        // Commit nonce strictly inside the executed animation frame after DOM rows are verified
        lastExecutedScrollNonceRef.current = reqNonce;

        const theadHeight = theadRef.current?.offsetHeight ?? 58;
        if (targetRow <= 1) {
          el.scrollTop = 0;
          el.scrollLeft = 0;
          return;
        }

        const viewportDataHeight = Math.max(120, el.clientHeight - theadHeight);
        if (targetTr) {
          const rowTopInTbody = Math.max(0, targetTr.offsetTop - theadHeight);
          el.scrollTop = Math.max(0, rowTopInTbody - Math.floor(viewportDataHeight * 0.55));
        }
        el.scrollLeft = 0;
      });
    });

    return () => {
      window.cancelAnimationFrame(rafId1);
      window.cancelAnimationFrame(rafId2);
    };
  }, [isActiveWorkspace, isLoading, activeSheetName, scrollPositionRequest]);

  // Keep keyboard-selected row visible below sticky headers when navigating with Arrow/Tab/Enter
  const ensureRowVisible = useCallback((targetGridRow: number) => {
    const container = gridContainerRef.current;
    if (!container) return;
    const theadHeight = theadRef.current?.offsetHeight ?? 58;

    if (targetGridRow <= 0) {
      container.scrollTop = 0;
      return;
    }

    const targetTr = tbodyRef.current?.children[targetGridRow - 1] as HTMLElement | undefined;
    if (!targetTr) return;

    const rowTop = targetTr.offsetTop;
    const rowBottom = rowTop + targetTr.offsetHeight;
    const visibleTop = container.scrollTop + theadHeight;
    const visibleBottom = container.scrollTop + container.clientHeight;

    if (rowTop < visibleTop) {
      container.scrollTop = Math.max(0, rowTop - theadHeight);
    } else if (rowBottom > visibleBottom) {
      container.scrollTop = rowBottom - container.clientHeight;
    }
  }, []);

  // Retrieve cell value by grid coordinates: row 0 is Row 1 (headers), row 1+ is Row 2..N
  const getCellValue = useCallback(
    (r: number, c: number): string | number => {
      if (r === 0) {
        return currentSheet.headers[c] ?? '';
      }
      const dataRowIdx = r - 1;
      if (dataRowIdx < currentSheet.rows.length) {
        const rowArr = currentSheet.rows[dataRowIdx];
        return rowArr ? (rowArr[c] ?? '') : '';
      }
      return '';
    },
    [currentSheet]
  );

  // Targeted RPC Save Queue (Google Sheets-style ~1000ms debounce, single-flight serialization, offline & lifecycle safety)
  const computePendingLayouts = useCallback(() => {
    const pending: Array<{
      dimKey: string;
      sheetId: string;
      sheetName: string;
      columnWidths: Record<number, number>;
      rowHeights: Record<number, number>;
    }> = [];

    for (const [sheetName, sheet] of Object.entries(sheetsRef.current)) {
      if (!sheet.id) continue;
      const dimKey = sheet.id;
      const curCols = colWidthsBySheetRef.current[dimKey] || {};
      const curRows = rowHeightsBySheetRef.current[dimKey] || {};
      const savedSnap = savedLayoutBySheetRef.current[dimKey] || { cols: {}, rows: {} };
      if (
        !areNumberMapsEqual(curCols, savedSnap.cols) ||
        !areNumberMapsEqual(curRows, savedSnap.rows)
      ) {
        pending.push({
          dimKey,
          sheetId: sheet.id,
          sheetName,
          columnWidths: curCols,
          rowHeights: curRows
        });
      }
    }
    return pending;
  }, []);

  const clearSaveTimers = useCallback(() => {
    if (autosaveTimerRef.current !== null) {
      window.clearTimeout(autosaveTimerRef.current);
      autosaveTimerRef.current = null;
    }
    if (retryTimerRef.current !== null) {
      window.clearTimeout(retryTimerRef.current);
      retryTimerRef.current = null;
    }
  }, []);

  const hasAnyDirtyWork = useCallback((): boolean => {
    return (
      Object.keys(dirtyCellsRef.current).length > 0 ||
      Object.keys(dirtyHeadersRef.current).length > 0 ||
      computePendingLayouts().length > 0
    );
  }, [computePendingLayouts]);

  const scheduleAutosave = useCallback(
    (delayMs: number = 1000) => {
      if (retryTimerRef.current !== null) {
        window.clearTimeout(retryTimerRef.current);
        retryTimerRef.current = null;
      }
      if (autosaveTimerRef.current !== null) {
        window.clearTimeout(autosaveTimerRef.current);
        autosaveTimerRef.current = null;
      }
      if (hasConflictRef.current) return;
      if (typeof navigator !== 'undefined' && navigator.onLine === false) {
        return;
      }

      autosaveTimerRef.current = window.setTimeout(() => {
        autosaveTimerRef.current = null;
        void flushSaveQueueRef.current({ source: 'autosave' });
      }, delayMs);
    },
    []
  );

  // Batch-safe cell updater preserving optimistic concurrency baselines + 100-step undo/redo history
  const applyCellUpdates = useCallback(
    (
      updates: Array<{ row: number; col: number; value: string | number }>,
      options?: {
        recordHistory?: boolean;
        selectionBefore?: SelectionState;
        selectionAfter?: SelectionState;
      }
    ) => {
      const targetSheet = sheets[activeSheetName];
      if (!targetSheet || !targetSheet.id || updates.length === 0) {
        return;
      }

      const sheetId = targetSheet.id;
      const nextDirtyHeaders = { ...dirtyHeadersRef.current };
      const nextDirtyCells = { ...dirtyCellsRef.current };
      let headersChanged = false;
      let rowsChanged = false;

      const updatedHeaders = [...targetSheet.headers];
      const updatedRows = [...targetSheet.rows];
      const diffs: CellDiffEntry[] = [];

      const applySingleUpdate = (r: number, c: number, value: string | number) => {
        if (r < 0 || c < 0 || c >= MAX_COLS_LIMIT) return;

        const prevValue =
          r === 0
            ? (updatedHeaders[c] ?? '')
            : (updatedRows[r - 1]?.[c] ?? '');

        if (prevValue === value) return;

        diffs.push({ row: r, col: c, before: prevValue, after: value });

        if (r === 0) {
          headersChanged = true;
          while (updatedHeaders.length <= c) updatedHeaders.push('');
          updatedHeaders[c] = String(value);

          const headerKey = `${sheetId}:${c}`;
          const existingDirtyHeader = nextDirtyHeaders[headerKey];
          const originalHeader =
            existingDirtyHeader !== undefined
              ? existingDirtyHeader.originalHeader
              : String(prevValue ?? '');

          if (String(value) === originalHeader) {
            delete nextDirtyHeaders[headerKey];
          } else {
            const expectedHeaderVersion =
              existingDirtyHeader?.expectedHeaderVersion ?? targetSheet.headerVersion ?? 1;
            nextDirtyHeaders[headerKey] = {
              sheetId,
              sheetName: activeSheetName,
              colIndex: c,
              newHeader: String(value),
              originalHeader,
              expectedHeaderVersion
            };
          }
        } else {
          rowsChanged = true;
          const dataRowIdx = r - 1;
          while (updatedRows.length <= dataRowIdx) {
            updatedRows.push(new Array(totalCols).fill(''));
          }
          const rowCopy = [...(updatedRows[dataRowIdx] || [])];
          while (rowCopy.length <= c) {
            rowCopy.push('');
          }
          rowCopy[c] = value;
          updatedRows[dataRowIdx] = rowCopy;

          const cellKey = `${sheetId}:${dataRowIdx}:${c}`;
          const existingDirtyCell = nextDirtyCells[cellKey];
          const originalValue =
            existingDirtyCell !== undefined ? existingDirtyCell.originalValue : prevValue;

          if (value === originalValue) {
            delete nextDirtyCells[cellKey];
          } else {
            const baselineVersion =
              existingDirtyCell?.expectedCellVersion ??
              targetSheet.cellVersions?.[dataRowIdx]?.[c] ??
              0;

            nextDirtyCells[cellKey] = {
              sheetId,
              sheetName: activeSheetName,
              rowIndex: dataRowIdx,
              colIndex: c,
              value,
              originalValue,
              expectedCellVersion: baselineVersion
            };
          }
        }
      };

      for (const { row: r, col: c, value } of updates) {
        applySingleUpdate(r, c, value);
      }

      // Automatic Reception Total = Test_Fee + Medication (dynamic header index safe across column moves)
      if (workspace === 'reception' && options?.recordHistory !== false) {
        let testFeeIdx = -1;
        let medIdx = -1;
        let totalIdx = -1;
        updatedHeaders.forEach((h, idx) => {
          const norm = normalizeHeaderKey(h);
          if (norm === 'test fee' && testFeeIdx === -1) testFeeIdx = idx;
          else if (norm === 'medication' && medIdx === -1) medIdx = idx;
          else if (norm === 'total' && totalIdx === -1) totalIdx = idx;
        });

        if (totalIdx !== -1 && (testFeeIdx !== -1 || medIdx !== -1)) {
          const feeTouchedRows = new Set<number>();
          const explicitTotalRows = new Set<number>();
          for (const u of updates) {
            if (u.row >= 1) {
              if (u.col === totalIdx) explicitTotalRows.add(u.row);
              if (u.col === testFeeIdx || u.col === medIdx) feeTouchedRows.add(u.row);
            }
          }

          for (const rowIdx of feeTouchedRows) {
            if (explicitTotalRows.has(rowIdx)) continue;
            const dataRow = updatedRows[rowIdx - 1];
            const feeParsed = parseOptionalFeeValue(
              testFeeIdx !== -1 ? dataRow?.[testFeeIdx] : ''
            );
            const medParsed = parseOptionalFeeValue(
              medIdx !== -1 ? dataRow?.[medIdx] : ''
            );

            if (feeParsed.invalid || medParsed.invalid) continue;

            const nextTotal: string | number =
              feeParsed.empty && medParsed.empty
                ? ''
                : Math.round((feeParsed.num + medParsed.num) * 100) / 100;

            applySingleUpdate(rowIdx, totalIdx, nextTotal);
          }
        }
      }

      if (!headersChanged && !rowsChanged) {
        return;
      }

      if (options?.recordHistory !== false && diffs.length > 0) {
        const stackKey = sheetId || activeSheetName;
        const prevStack = undoStackBySheetRef.current[stackKey] || [];
        const selSnapshot = options?.selectionBefore ?? selection;
        const step: HistoryStep = {
          diffs,
          selectionBefore: selSnapshot,
          selectionAfter: options?.selectionAfter ?? selSnapshot
        };
        undoStackBySheetRef.current[stackKey] = [...prevStack, step].slice(-MAX_UNDO_HISTORY);
        redoStackBySheetRef.current[stackKey] = [];
      }

      if (headersChanged) {
        dirtyHeadersRef.current = nextDirtyHeaders;
        setDirtyHeaders(nextDirtyHeaders);
      }
      if (rowsChanged) {
        dirtyCellsRef.current = nextDirtyCells;
        setDirtyCells(nextDirtyCells);
      }

      setSheets((prev) => {
        const sheet = prev[activeSheetName];
        if (!sheet) return prev;
        return {
          ...prev,
          [activeSheetName]: {
            ...sheet,
            ...(headersChanged ? { headers: updatedHeaders } : {}),
            ...(rowsChanged ? { rows: updatedRows } : {})
          }
        };
      });

      setHasConflict(false);
      setSaveError(null);
      retryAttemptRef.current = 0;

      const stillHasDirty =
        Object.keys(nextDirtyCells).length > 0 ||
        Object.keys(nextDirtyHeaders).length > 0 ||
        computePendingLayouts().length > 0;

      if (stillHasDirty) {
        setSaveStatus('pending');
        setSaveMessage('');
        scheduleAutosave(1000);
      } else {
        if (autosaveTimerRef.current !== null) {
          window.clearTimeout(autosaveTimerRef.current);
          autosaveTimerRef.current = null;
        }
        setSaveStatus('idle');
        setSaveMessage('');
      }
    },
    [activeSheetName, sheets, totalCols, selection, workspace, computePendingLayouts, scheduleAutosave]
  );

  const setCellValue = useCallback(
    (
      r: number,
      c: number,
      value: string | number,
      options?: { selectionBefore?: SelectionState; selectionAfter?: SelectionState }
    ) => {
      applyCellUpdates([{ row: r, col: c, value }], options);
    },
    [applyCellUpdates]
  );

  useEffect(() => {
    if (isEditing && editInputRef.current) {
      editInputRef.current.focus();
    }
  }, [isEditing]);

  // Commit current edit, optionally stepping rowDelta / colDelta afterwards
  const commitEdit = useCallback(
    (rowDelta: number = 0, colDelta: number = 0) => {
      if (!isEditing) return;
      const currentSelection = selection;
      const prevVal = getCellValue(selectedCell.row, selectedCell.col);

      let nextSelection = currentSelection;
      if (rowDelta !== 0 || colDelta !== 0) {
        const nextRow = Math.max(0, Math.min(displayRowsCount - 1, selectedCell.row + rowDelta));
        const nextCol = Math.max(0, Math.min(totalCols - 1, selectedCell.col + colDelta));
        nextSelection = {
          mode: 'cell',
          anchorRow: nextRow,
          anchorCol: nextCol,
          focusRow: nextRow,
          focusCol: nextCol
        };
      }

      if (String(prevVal ?? '') !== editValue) {
        setCellValue(selectedCell.row, selectedCell.col, editValue, {
          selectionBefore: currentSelection,
          selectionAfter: nextSelection
        });
      }
      setIsEditing(false);

      if (rowDelta !== 0 || colDelta !== 0) {
        setSelection(nextSelection);
        ensureRowVisible(nextSelection.anchorRow);
      }
    },
    [
      isEditing,
      selection,
      getCellValue,
      selectedCell.row,
      selectedCell.col,
      editValue,
      setCellValue,
      displayRowsCount,
      ensureRowVisible,
      totalCols
    ]
  );

  const cancelEdit = useCallback(() => {
    setIsEditing(false);
  }, []);

  // Latest-state refs so memoized row callbacks & window listeners stay current without re-renders
  const isEditingRef = useRef(isEditing);
  const selectedCellRef = useRef(selectedCell);
  const selectionModeRef = useRef(selection.mode);
  const selectionBoundsRef = useRef(selectionBounds);
  const renamingSheetNameRef = useRef(renamingSheetName);
  const isTourOpenRef = useRef(isTourOpen);
  const pendingBulkClearRef = useRef(pendingBulkClear);
  const pendingConfirmActionRef = useRef(pendingConfirmAction);
  const contextMenuRef = useRef(contextMenu);
  const resizeColDialogRef = useRef(resizeColDialog);
  const isStructuralBusyRef = useRef(isStructuralBusy);
  const rowHeaderWidthRef = useRef(rowHeaderWidth);
  const totalColsRef = useRef(totalCols);
  const displayRowsCountRef = useRef(displayRowsCount);
  const activeDimKeyRef = useRef(activeDimKey);
  const totalTableWidthRef = useRef(totalTableWidth);
  const commitEditRef = useRef(commitEdit);
  const setCellValueRef = useRef(setCellValue);
  const getCellValueRef = useRef(getCellValue);
  const getColWidthRef = useRef(getColWidth);
  const getManualRowHeightRef = useRef(getManualRowHeight);
  const justLongPressedRef = useRef<boolean>(false);

  useEffect(() => {
    isEditingRef.current = isEditing;
    selectedCellRef.current = selectedCell;
    selectionModeRef.current = selection.mode;
    selectionBoundsRef.current = selectionBounds;
    renamingSheetNameRef.current = renamingSheetName;
    isTourOpenRef.current = isTourOpen;
    pendingBulkClearRef.current = pendingBulkClear;
    pendingConfirmActionRef.current = pendingConfirmAction;
    contextMenuRef.current = contextMenu;
    resizeColDialogRef.current = resizeColDialog;
    isStructuralBusyRef.current = isStructuralBusy;
    rowHeaderWidthRef.current = rowHeaderWidth;
    totalColsRef.current = totalCols;
    displayRowsCountRef.current = displayRowsCount;
    activeDimKeyRef.current = activeDimKey;
    totalTableWidthRef.current = totalTableWidth;
    commitEditRef.current = commitEdit;
    setCellValueRef.current = setCellValue;
    getCellValueRef.current = getCellValue;
    getColWidthRef.current = getColWidth;
    getManualRowHeightRef.current = getManualRowHeight;
  });

  useEffect(() => {
    const wasActive = wasActiveWorkspaceRef.current;
    wasActiveWorkspaceRef.current = isActiveWorkspace;

    if (!wasActive && isActiveWorkspace && !isLoading && activeSheetName && !isEditingRef.current) {
      const sheet = sheetsRef.current[activeSheetName];
      const latestRow = findLatestRecordGridRow(sheet);
      setSelection({
        mode: 'cell',
        anchorRow: latestRow,
        anchorCol: 0,
        focusRow: latestRow,
        focusCol: 0
      });
      setScrollPositionRequest((prev) => ({
        sheetName: activeSheetName,
        targetRow: latestRow,
        nonce: prev.nonce + 1
      }));
    }
  }, [isActiveWorkspace, isLoading, activeSheetName]);

  // Auto-launch role-aware guided tour once per user when landing workbook finishes loading
  const tourStorageKey = user ? `mk_tour_completed_${user.id}_${user.role}` : null;
  useEffect(() => {
    if (hasAutoCheckedTourRef.current) return;
    if (!isActiveWorkspace || isLoading || sheetKeys.length === 0 || !tourStorageKey) return;

    hasAutoCheckedTourRef.current = true;
    try {
      const completed = window.localStorage.getItem(tourStorageKey);
      if (completed !== '1') {
        const rafId = window.requestAnimationFrame(() => {
          setIsTourOpen(true);
        });
        return () => window.cancelAnimationFrame(rafId);
      }
    } catch {
      // Ignore storage access errors
    }
  }, [isActiveWorkspace, isLoading, sheetKeys.length, tourStorageKey]);

  const markTourSeen = useCallback(() => {
    if (!tourStorageKey) return;
    try {
      window.localStorage.setItem(tourStorageKey, '1');
    } catch {
      // Ignore storage access errors
    }
  }, [tourStorageKey]);

  const handleCloseTour = useCallback(() => {
    markTourSeen();
    setIsTourOpen(false);
    rootRef.current?.focus({ preventScroll: true });
  }, [markTourSeen]);

  const handleCompleteTour = useCallback(() => {
    markTourSeen();
    setIsTourOpen(false);
    rootRef.current?.focus({ preventScroll: true });
  }, [markTourSeen]);

  const findLastPopulatedBounds = useCallback(() => {
    let maxRow = 0;
    let maxCol = 0;

    for (let c = 0; c < currentSheet.headers.length; c++) {
      if (isNonEmptyCell(currentSheet.headers[c])) {
        maxCol = Math.max(maxCol, c);
      }
    }
    for (let r = 0; r < currentSheet.rows.length; r++) {
      const row = currentSheet.rows[r];
      if (!Array.isArray(row)) continue;
      for (let c = 0; c < row.length; c++) {
        if (isNonEmptyCell(row[c])) {
          maxRow = Math.max(maxRow, r + 1);
          maxCol = Math.max(maxCol, c);
        }
      }
    }
    return { maxRow, maxCol };
  }, [currentSheet.headers, currentSheet.rows]);

  const findRowLastPopulatedCol = useCallback(
    (gridRowIdx: number): number => {
      if (gridRowIdx === 0) {
        for (let c = currentSheet.headers.length - 1; c >= 0; c--) {
          if (isNonEmptyCell(currentSheet.headers[c])) return c;
        }
        return 0;
      }
      const row = currentSheet.rows[gridRowIdx - 1];
      if (Array.isArray(row)) {
        for (let c = row.length - 1; c >= 0; c--) {
          if (isNonEmptyCell(row[c])) return c;
        }
      }
      return Math.max(0, currentSheet.headers.length - 1);
    },
    [currentSheet.headers, currentSheet.rows]
  );

  // Undo & Redo handlers (100-step per-sheet history, preserved across autosaves with fresh concurrency baselines)
  const handleUndo = useCallback(() => {
    if (
      isEditingRef.current ||
      renamingSheetNameRef.current !== null ||
      isTourOpenRef.current ||
      pendingBulkClearRef.current !== null ||
      pendingConfirmActionRef.current !== null ||
      resizeColDialogRef.current !== null
    ) {
      return;
    }

    const activeSheet = sheetsRef.current[activeSheetName];
    if (!activeSheet?.id) return;

    const stackKey = activeSheet.id;
    const undoStack = undoStackBySheetRef.current[stackKey] || [];
    if (undoStack.length === 0) return;

    const step = undoStack[undoStack.length - 1];
    undoStackBySheetRef.current[stackKey] = undoStack.slice(0, -1);

    const reverseUpdates = step.diffs.map((d) => ({
      row: d.row,
      col: d.col,
      value: d.before
    }));
    applyCellUpdates(reverseUpdates, { recordHistory: false });

    const redoStack = redoStackBySheetRef.current[stackKey] || [];
    redoStackBySheetRef.current[stackKey] = [...redoStack, step].slice(-MAX_UNDO_HISTORY);

    setSelection(step.selectionBefore);
    ensureRowVisible(step.selectionBefore.anchorRow);
  }, [activeSheetName, applyCellUpdates, ensureRowVisible]);

  const handleRedo = useCallback(() => {
    if (
      isEditingRef.current ||
      renamingSheetNameRef.current !== null ||
      isTourOpenRef.current ||
      pendingBulkClearRef.current !== null ||
      pendingConfirmActionRef.current !== null ||
      resizeColDialogRef.current !== null
    ) {
      return;
    }

    const activeSheet = sheetsRef.current[activeSheetName];
    if (!activeSheet?.id) return;

    const stackKey = activeSheet.id;
    const redoStack = redoStackBySheetRef.current[stackKey] || [];
    if (redoStack.length === 0) return;

    const step = redoStack[redoStack.length - 1];
    redoStackBySheetRef.current[stackKey] = redoStack.slice(0, -1);

    const forwardUpdates = step.diffs.map((d) => ({
      row: d.row,
      col: d.col,
      value: d.after
    }));
    applyCellUpdates(forwardUpdates, { recordHistory: false });

    const undoStack = undoStackBySheetRef.current[stackKey] || [];
    undoStackBySheetRef.current[stackKey] = [...undoStack, step].slice(-MAX_UNDO_HISTORY);

    setSelection(step.selectionAfter);
    ensureRowVisible(step.selectionAfter.anchorRow);
  }, [activeSheetName, applyCellUpdates, ensureRowVisible]);

  const handleUndoRef = useRef(handleUndo);
  const handleRedoRef = useRef(handleRedo);
  useEffect(() => {
    handleUndoRef.current = handleUndo;
    handleRedoRef.current = handleRedo;
  });

  // Clear selected cells with clinical safeguard confirmation for bulk/row/col/all deletions
  const handleDeleteSelection = useCallback(() => {
    if (selection.mode === 'cell') {
      const cur = getCellValue(selectedCell.row, selectedCell.col);
      if (isNonEmptyCell(cur)) {
        setCellValue(selectedCell.row, selectedCell.col, '', {
          selectionBefore: selection,
          selectionAfter: selection
        });
      }
      return;
    }

    const updates: Array<{ row: number; col: number; value: string | number }> = [];
    let populatedCount = 0;
    let includesHeader = false;

    const maxEffectiveRow =
      selection.mode === 'col' || selection.mode === 'all'
        ? Math.min(selectionBounds.rMax, Math.max(currentSheet.rows.length, 0))
        : selectionBounds.rMax;
    const maxEffectiveCol =
      selection.mode === 'row' || selection.mode === 'all'
        ? Math.min(selectionBounds.cMax, Math.max(currentSheet.headers.length - 1, maxRowCols - 1, 0))
        : selectionBounds.cMax;

    for (let r = selectionBounds.rMin; r <= maxEffectiveRow; r++) {
      for (let c = selectionBounds.cMin; c <= maxEffectiveCol; c++) {
        const val = getCellValue(r, c);
        if (isNonEmptyCell(val)) {
          populatedCount++;
          if (r === 0) includesHeader = true;
          updates.push({ row: r, col: c, value: '' });
        }
      }
    }

    if (updates.length === 0) return;

    if (
      selection.mode === 'all' ||
      selection.mode === 'col' ||
      selection.mode === 'row' ||
      populatedCount > 1 ||
      includesHeader
    ) {
      setPendingBulkClear({ updates, populatedCount, includesHeader });
      return;
    }

    applyCellUpdates(updates, {
      selectionBefore: selection,
      selectionAfter: selection
    });
  }, [
    selection,
    selectedCell.row,
    selectedCell.col,
    getCellValue,
    setCellValue,
    selectionBounds,
    currentSheet.rows.length,
    currentSheet.headers.length,
    maxRowCols,
    applyCellUpdates
  ]);

  // Complete Excel-style keyboard navigation & shortcuts with strict input/rename isolation
  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (
      renamingSheetName !== null ||
      isTourOpen ||
      pendingBulkClear !== null ||
      pendingConfirmAction !== null ||
      resizeColDialog !== null
    ) {
      return;
    }

    const isMod = e.ctrlKey || e.metaKey;
    const lowerKey = e.key.toLowerCase();

    // Ctrl/Cmd + S and Undo/Redo handled globally by active-workspace window listener
    if (isMod && (lowerKey === 's' || lowerKey === 'z' || lowerKey === 'y')) {
      if (lowerKey === 's') e.preventDefault();
      return;
    }

    if (isEditing) {
      if (e.key === 'Enter') {
        e.preventDefault();
        commitEdit(e.shiftKey ? -1 : 1, 0);
      } else if (e.key === 'Tab') {
        e.preventDefault();
        commitEdit(0, e.shiftKey ? -1 : 1);
      } else if (e.key === 'Escape') {
        e.preventDefault();
        cancelEdit();
        rootRef.current?.focus({ preventScroll: true });
      }
      return;
    }

    // Never intercept typing/navigation inside other inputs, dialogs, or tab rename fields
    if (isTextInputElement(e.target)) {
      return;
    }

    // Ctrl/Cmd + A: smart select populated data region first, then entire worksheet on repeat
    if (isMod && lowerKey === 'a') {
      e.preventDefault();
      const { maxRow, maxCol } = findLastPopulatedBounds();
      const hasDataRegion = maxRow > 0 || maxCol > 0;
      const alreadySelectingDataRegion =
        selection.mode === 'range' &&
        selectionBounds.rMin === 0 &&
        selectionBounds.cMin === 0 &&
        selectionBounds.rMax === maxRow &&
        selectionBounds.cMax === maxCol;

      if (hasDataRegion && !alreadySelectingDataRegion && selection.mode !== 'all') {
        setSelection({
          mode: 'range',
          anchorRow: 0,
          anchorCol: 0,
          focusRow: maxRow,
          focusCol: maxCol
        });
      } else {
        setSelection({
          mode: 'all',
          anchorRow: 0,
          anchorCol: 0,
          focusRow: displayRowsCount - 1,
          focusCol: totalCols - 1
        });
      }
      return;
    }

    // Ctrl/Cmd + Home: first worksheet cell (A1)
    if (isMod && e.key === 'Home') {
      e.preventDefault();
      if (e.shiftKey) {
        setSelection((prev) => ({
          mode: prev.anchorRow === 0 && prev.anchorCol === 0 ? 'cell' : 'range',
          anchorRow: prev.anchorRow,
          anchorCol: prev.anchorCol,
          focusRow: 0,
          focusCol: 0
        }));
      } else {
        setSelection({
          mode: 'cell',
          anchorRow: 0,
          anchorCol: 0,
          focusRow: 0,
          focusCol: 0
        });
      }
      ensureRowVisible(0);
      if (gridContainerRef.current) {
        gridContainerRef.current.scrollLeft = 0;
      }
      return;
    }

    // Ctrl/Cmd + End: last populated worksheet cell
    if (isMod && e.key === 'End') {
      e.preventDefault();
      const { maxRow, maxCol } = findLastPopulatedBounds();
      if (e.shiftKey) {
        setSelection((prev) => ({
          mode: prev.anchorRow === maxRow && prev.anchorCol === maxCol ? 'cell' : 'range',
          anchorRow: prev.anchorRow,
          anchorCol: prev.anchorCol,
          focusRow: maxRow,
          focusCol: maxCol
        }));
      } else {
        setSelection({
          mode: 'cell',
          anchorRow: maxRow,
          anchorCol: maxCol,
          focusRow: maxRow,
          focusCol: maxCol
        });
      }
      ensureRowVisible(maxRow);
      return;
    }

    // Home / End (without Ctrl/Meta): beginning or end of current row
    if (!isMod && e.key === 'Home') {
      e.preventDefault();
      if (e.shiftKey) {
        setSelection((prev) => ({
          mode: prev.anchorRow === prev.focusRow && prev.anchorCol === 0 ? 'cell' : 'range',
          anchorRow: prev.anchorRow,
          anchorCol: prev.anchorCol,
          focusRow: prev.focusRow,
          focusCol: 0
        }));
      } else {
        setSelection({
          mode: 'cell',
          anchorRow: selectedCell.row,
          anchorCol: 0,
          focusRow: selectedCell.row,
          focusCol: 0
        });
      }
      if (gridContainerRef.current) {
        gridContainerRef.current.scrollLeft = 0;
      }
      return;
    }

    if (!isMod && e.key === 'End') {
      e.preventDefault();
      const endCol = findRowLastPopulatedCol(e.shiftKey ? selection.focusRow : selectedCell.row);
      if (e.shiftKey) {
        setSelection((prev) => ({
          mode: prev.anchorRow === prev.focusRow && prev.anchorCol === endCol ? 'cell' : 'range',
          anchorRow: prev.anchorRow,
          anchorCol: prev.anchorCol,
          focusRow: prev.focusRow,
          focusCol: endCol
        }));
      } else {
        setSelection({
          mode: 'cell',
          anchorRow: selectedCell.row,
          anchorCol: endCol,
          focusRow: selectedCell.row,
          focusCol: endCol
        });
      }
      return;
    }

    if (e.key === 'Escape') {
      e.preventDefault();
      if (contextMenu) {
        setContextMenu(null);
      } else if (selection.mode !== 'cell') {
        setSelection({
          mode: 'cell',
          anchorRow: selection.anchorRow,
          anchorCol: selection.anchorCol,
          focusRow: selection.anchorRow,
          focusCol: selection.anchorCol
        });
      }
      return;
    }

    if (e.key === 'ArrowUp' || e.key === 'ArrowDown' || e.key === 'ArrowLeft' || e.key === 'ArrowRight') {
      e.preventDefault();
      const dRow = e.key === 'ArrowUp' ? -1 : e.key === 'ArrowDown' ? 1 : 0;
      const dCol = e.key === 'ArrowLeft' ? -1 : e.key === 'ArrowRight' ? 1 : 0;

      if (e.shiftKey) {
        const nextFocusRow = Math.max(0, Math.min(displayRowsCount - 1, selection.focusRow + dRow));
        const nextFocusCol = Math.max(0, Math.min(totalCols - 1, selection.focusCol + dCol));
        setSelection((prev) => ({
          mode: nextFocusRow === prev.anchorRow && nextFocusCol === prev.anchorCol ? 'cell' : 'range',
          anchorRow: prev.anchorRow,
          anchorCol: prev.anchorCol,
          focusRow: nextFocusRow,
          focusCol: nextFocusCol
        }));
        ensureRowVisible(nextFocusRow);
      } else {
        const nextRow = Math.max(0, Math.min(displayRowsCount - 1, selectedCell.row + dRow));
        const nextCol = Math.max(0, Math.min(totalCols - 1, selectedCell.col + dCol));
        setSelection({
          mode: 'cell',
          anchorRow: nextRow,
          anchorCol: nextCol,
          focusRow: nextRow,
          focusCol: nextCol
        });
        ensureRowVisible(nextRow);
      }
      return;
    }

    if (e.key === 'Tab') {
      e.preventDefault();
      const nextCol = e.shiftKey
        ? Math.max(0, selectedCell.col - 1)
        : Math.min(totalCols - 1, selectedCell.col + 1);
      setSelection({
        mode: 'cell',
        anchorRow: selectedCell.row,
        anchorCol: nextCol,
        focusRow: selectedCell.row,
        focusCol: nextCol
      });
    } else if (e.key === 'Enter' || e.key === 'F2') {
      e.preventDefault();
      const currentVal = getCellValue(selectedCell.row, selectedCell.col);
      setEditValue(String(currentVal ?? ''));
      setIsEditing(true);
    } else if (e.key === 'Delete' || e.key === 'Backspace') {
      e.preventDefault();
      handleDeleteSelection();
    } else if (e.key.length === 1 && !isMod && !e.altKey) {
      setSelection({
        mode: 'cell',
        anchorRow: selectedCell.row,
        anchorCol: selectedCell.col,
        focusRow: selectedCell.row,
        focusCol: selectedCell.col
      });
      setEditValue(e.key);
      setIsEditing(true);
    }
  };

  // Single click selects only — double-click (or fast touch double-tap on already-selected cell) opens edit mode
  const handleCellClick = useCallback((r: number, c: number, shiftKey: boolean, detail: number = 1) => {
    if (justResizedRef.current || justLongPressedRef.current) return;
    setContextMenu(null);
    rootRef.current?.focus({ preventScroll: true });

    if (shiftKey) {
      lastTouchTapRef.current = null;
      if (isEditingRef.current) commitEditRef.current(0, 0);
      setSelection((prev) => ({
        mode: r === prev.anchorRow && c === prev.anchorCol ? 'cell' : 'range',
        anchorRow: prev.anchorRow,
        anchorCol: prev.anchorCol,
        focusRow: r,
        focusCol: c
      }));
      return;
    }

    if (isEditingRef.current) {
      if (selectedCellRef.current.row === r && selectedCellRef.current.col === c) {
        return;
      }
      commitEditRef.current(0, 0);
    }

    const now = Date.now();
    const isCoarsePointer =
      typeof window !== 'undefined' && window.matchMedia('(pointer: coarse)').matches;
    const lastTap = lastTouchTapRef.current;

    if (
      isCoarsePointer &&
      detail <= 1 &&
      lastTap &&
      lastTap.row === r &&
      lastTap.col === c &&
      selectedCellRef.current.row === r &&
      selectedCellRef.current.col === c &&
      now - lastTap.time < 380
    ) {
      lastTouchTapRef.current = null;
      const val = getCellValueRef.current(r, c);
      setEditValue(String(val ?? ''));
      setIsEditing(true);
      return;
    }

    lastTouchTapRef.current = { row: r, col: c, time: now };

    setSelection({
      mode: 'cell',
      anchorRow: r,
      anchorCol: c,
      focusRow: r,
      focusCol: c
    });
  }, []);

  // Intentional double-click opens edit mode on the clicked cell
  const handleCellDoubleClick = useCallback((r: number, c: number, detail: number) => {
    if (justResizedRef.current || detail < 2) return;
    rootRef.current?.focus({ preventScroll: true });
    if (isEditingRef.current) commitEditRef.current(0, 0);
    setSelection({
      mode: 'cell',
      anchorRow: r,
      anchorCol: c,
      focusRow: r,
      focusCol: c
    });
    const val = getCellValueRef.current(r, c);
    setEditValue(String(val ?? ''));
    setIsEditing(true);
  }, []);

  const handleRowHeaderClick = useCallback((r: number, shiftKey: boolean) => {
    if (justResizedRef.current || justLongPressedRef.current) return;
    setContextMenu(null);
    rootRef.current?.focus({ preventScroll: true });
    if (isEditingRef.current) commitEditRef.current(0, 0);

    const maxColIdx = totalColsRef.current - 1;
    if (shiftKey && selectionModeRef.current === 'row') {
      setSelection((prev) => ({
        mode: 'row',
        anchorRow: prev.anchorRow,
        anchorCol: 0,
        focusRow: r,
        focusCol: maxColIdx
      }));
    } else {
      setSelection({
        mode: 'row',
        anchorRow: r,
        anchorCol: 0,
        focusRow: r,
        focusCol: maxColIdx
      });
    }
  }, []);

  const handleColHeaderClick = useCallback((c: number, shiftKey: boolean) => {
    if (justResizedRef.current || justLongPressedRef.current) return;
    setContextMenu(null);
    rootRef.current?.focus({ preventScroll: true });
    if (isEditingRef.current) commitEditRef.current(0, 0);

    const maxRowIdx = displayRowsCountRef.current - 1;
    if (shiftKey && selectionModeRef.current === 'col') {
      setSelection((prev) => ({
        mode: 'col',
        anchorRow: 0,
        anchorCol: prev.anchorCol,
        focusRow: maxRowIdx,
        focusCol: c
      }));
    } else {
      setSelection({
        mode: 'col',
        anchorRow: 0,
        anchorCol: c,
        focusRow: maxRowIdx,
        focusCol: c
      });
    }
  }, []);

  const handleSelectAllClick = useCallback(() => {
    rootRef.current?.focus({ preventScroll: true });
    if (isEditingRef.current) commitEditRef.current(0, 0);
    setSelection({
      mode: 'all',
      anchorRow: 0,
      anchorCol: 0,
      focusRow: displayRowsCountRef.current - 1,
      focusCol: totalColsRef.current - 1
    });
  }, []);

  const setEditInputRef = useCallback((el: HTMLInputElement | null) => {
    editInputRef.current = el;
  }, []);

  const setRowElementRef = useCallback((gridRowIdx: number, el: HTMLTableRowElement | null) => {
    rowElementsRef.current[gridRowIdx] = el;
  }, []);

  const handleCommitEditBlur = useCallback(() => {
    commitEditRef.current(0, 0);
  }, []);

  const handlePaidValueChangeCb = useCallback((gridRowIdx: number, colIdx: number, val: string) => {
    if (isEditingRef.current) commitEditRef.current(0, 0);
    const cellSel: SelectionState = {
      mode: 'cell',
      anchorRow: gridRowIdx,
      anchorCol: colIdx,
      focusRow: gridRowIdx,
      focusCol: colIdx
    };
    setSelection(cellSel);
    setCellValueRef.current(gridRowIdx, colIdx, val, {
      selectionBefore: cellSel,
      selectionAfter: cellSel
    });
  }, []);

  const buildTsvForBounds = useCallback(
    (rStart: number, cStart: number, rEnd: number, cEnd: number): string => {
      const lines: string[] = [];
      for (let r = rStart; r <= rEnd; r++) {
        const rowCells: string[] = [];
        for (let c = cStart; c <= cEnd; c++) {
          const val = getCellValue(r, c);
          rowCells.push(val !== undefined && val !== null ? String(val) : '');
        }
        lines.push(rowCells.join('\t'));
      }
      return lines.join('\n');
    },
    [getCellValue]
  );

  const buildSelectionTsv = useCallback((): string => {
    const rStart = selectionBounds.rMin;
    const cStart = selectionBounds.cMin;
    const rEnd =
      selection.mode === 'col' || selection.mode === 'all'
        ? Math.min(selectionBounds.rMax, Math.max(currentSheet.rows.length, 0))
        : selectionBounds.rMax;
    const cEnd =
      selection.mode === 'row' || selection.mode === 'all'
        ? Math.min(selectionBounds.cMax, Math.max(currentSheet.headers.length - 1, maxRowCols - 1, 0))
        : selectionBounds.cMax;

    return buildTsvForBounds(rStart, cStart, rEnd, cEnd);
  }, [
    selectionBounds,
    selection.mode,
    currentSheet.rows.length,
    currentSheet.headers.length,
    maxRowCols,
    buildTsvForBounds
  ]);

  const writeToClipboardSafe = useCallback(
    async (text: string, kind: 'cells' | 'row' | 'col' = 'cells') => {
      internalClipboardRef.current = { text, kind };
      if (typeof navigator !== 'undefined' && navigator.clipboard?.writeText) {
        try {
          await navigator.clipboard.writeText(text);
        } catch {
          // Internal clipboard fallback retains text cleanly if permission is restricted
        }
      }
    },
    []
  );

  const readFromClipboardSafe = useCallback(async (): Promise<string> => {
    if (typeof navigator !== 'undefined' && navigator.clipboard?.readText) {
      try {
        const txt = await navigator.clipboard.readText();
        if (txt) return txt;
      } catch {
        // Fallback to in-app clipboard
      }
    }
    return internalClipboardRef.current?.text ?? '';
  }, []);

  const applyPasteTextToTarget = useCallback(
    (
      text: string,
      targetSelection: SelectionState,
      options?: { confirmOverwriteIfPopulated?: boolean }
    ) => {
      if (!text) return;
      const matrix = parseClipboardTsv(text);
      if (matrix.length === 0) return;

      const tBounds = (() => {
        if (targetSelection.mode === 'all') {
          return { rMin: 0, rMax: displayRowsCount - 1, cMin: 0, cMax: totalCols - 1 };
        }
        if (targetSelection.mode === 'row') {
          return {
            rMin: Math.min(targetSelection.anchorRow, targetSelection.focusRow),
            rMax: Math.max(targetSelection.anchorRow, targetSelection.focusRow),
            cMin: 0,
            cMax: totalCols - 1
          };
        }
        if (targetSelection.mode === 'col') {
          return {
            rMin: 0,
            rMax: displayRowsCount - 1,
            cMin: Math.min(targetSelection.anchorCol, targetSelection.focusCol),
            cMax: Math.max(targetSelection.anchorCol, targetSelection.focusCol)
          };
        }
        return {
          rMin: Math.min(targetSelection.anchorRow, targetSelection.focusRow),
          rMax: Math.max(targetSelection.anchorRow, targetSelection.focusRow),
          cMin: Math.min(targetSelection.anchorCol, targetSelection.focusCol),
          cMax: Math.max(targetSelection.anchorCol, targetSelection.focusCol)
        };
      })();

      const updates: Array<{ row: number; col: number; value: string | number }> = [];
      let nextSelection = targetSelection;

      if (matrix.length === 1 && matrix[0].length === 1) {
        const singleVal = matrix[0][0];
        if (targetSelection.mode === 'range') {
          for (let r = tBounds.rMin; r <= tBounds.rMax; r++) {
            for (let c = tBounds.cMin; c <= tBounds.cMax; c++) {
              updates.push({ row: r, col: c, value: singleVal });
            }
          }
        } else if (targetSelection.mode === 'row') {
          const effectiveCols = Math.max(currentSheet.headers.length, maxRowCols, 1);
          for (let r = tBounds.rMin; r <= tBounds.rMax; r++) {
            for (let c = 0; c < effectiveCols; c++) {
              updates.push({ row: r, col: c, value: singleVal });
            }
          }
        } else if (targetSelection.mode === 'col') {
          const effectiveRows = Math.max(currentSheet.rows.length, 1);
          for (let r = 0; r <= effectiveRows; r++) {
            for (let c = tBounds.cMin; c <= tBounds.cMax; c++) {
              updates.push({ row: r, col: c, value: singleVal });
            }
          }
        } else {
          updates.push({
            row: targetSelection.anchorRow,
            col: targetSelection.anchorCol,
            value: singleVal
          });
          nextSelection = {
            mode: 'cell',
            anchorRow: targetSelection.anchorRow,
            anchorCol: targetSelection.anchorCol,
            focusRow: targetSelection.anchorRow,
            focusCol: targetSelection.anchorCol
          };
        }
      } else {
        const startRow = tBounds.rMin;
        const startCol = tBounds.cMin;
        let maxRowHit = startRow;
        let maxColHit = startCol;
        for (let dr = 0; dr < matrix.length; dr++) {
          const targetRow = startRow + dr;
          if (targetRow >= 100000) break;
          const cols = matrix[dr];
          for (let dc = 0; dc < cols.length; dc++) {
            const targetCol = startCol + dc;
            if (targetCol >= MAX_COLS_LIMIT) break;
            updates.push({ row: targetRow, col: targetCol, value: cols[dc] });
            maxRowHit = Math.max(maxRowHit, targetRow);
            maxColHit = Math.max(maxColHit, targetCol);
          }
        }
        if (updates.length > 0 && targetSelection.mode !== 'row' && targetSelection.mode !== 'col') {
          nextSelection = {
            mode: maxRowHit === startRow && maxColHit === startCol ? 'cell' : 'range',
            anchorRow: startRow,
            anchorCol: startCol,
            focusRow: maxRowHit,
            focusCol: maxColHit
          };
        }
      }

      if (updates.length === 0) return;

      if (options?.confirmOverwriteIfPopulated) {
        let overwriteCount = 0;
        for (const u of updates) {
          const existing = getCellValue(u.row, u.col);
          if (isNonEmptyCell(existing) && String(existing) !== String(u.value)) {
            overwriteCount++;
          }
        }
        if (overwriteCount > 0) {
          setPendingConfirmAction({
            title: 'Overwrite Existing Cells?',
            message: `Pasting here will overwrite ${overwriteCount} populated cell${
              overwriteCount === 1 ? '' : 's'
            }. Do you want to continue?`,
            confirmLabel: 'Paste & Overwrite',
            confirmVariant: 'primary',
            run: () => {
              applyCellUpdates(updates, {
                selectionBefore: selection,
                selectionAfter: nextSelection
              });
              setSelection(nextSelection);
            }
          });
          return;
        }
      }

      applyCellUpdates(updates, {
        selectionBefore: selection,
        selectionAfter: nextSelection
      });
      setSelection(nextSelection);
    },
    [
      displayRowsCount,
      totalCols,
      currentSheet.headers.length,
      currentSheet.rows.length,
      maxRowCols,
      getCellValue,
      applyCellUpdates,
      selection
    ]
  );

  // Copy selected cell, range, row(s), column(s), or entire worksheet in Excel TSV format
  const handleCopy = (e: React.ClipboardEvent) => {
    if (
      isEditing ||
      renamingSheetName !== null ||
      isTourOpen ||
      pendingBulkClear !== null ||
      pendingConfirmAction !== null ||
      resizeColDialog !== null
    ) {
      return;
    }
    if (isTextInputElement(e.target) || isTextInputElement(document.activeElement)) {
      return;
    }

    const tsv = buildSelectionTsv();
    const clipKind = selection.mode === 'row' ? 'row' : selection.mode === 'col' ? 'col' : 'cells';
    internalClipboardRef.current = { text: tsv, kind: clipKind };
    e.clipboardData.setData('text/plain', tsv);
    e.preventDefault();
  };

  // Cut selected cell(s) to clipboard and clear with deletion safeguards + single undo step
  const handleCut = (e: React.ClipboardEvent) => {
    if (
      isEditing ||
      renamingSheetName !== null ||
      isTourOpen ||
      pendingBulkClear !== null ||
      pendingConfirmAction !== null ||
      resizeColDialog !== null
    ) {
      return;
    }
    if (isTextInputElement(e.target) || isTextInputElement(document.activeElement)) {
      return;
    }

    const tsv = buildSelectionTsv();
    const clipKind = selection.mode === 'row' ? 'row' : selection.mode === 'col' ? 'col' : 'cells';
    internalClipboardRef.current = { text: tsv, kind: clipKind };
    e.clipboardData.setData('text/plain', tsv);
    e.preventDefault();
    handleDeleteSelection();
  };

  // Paste compatible TSV/scalar values starting at selection anchor (queues debounced autosave)
  const handlePaste = (e: React.ClipboardEvent) => {
    if (
      isEditing ||
      renamingSheetName !== null ||
      isTourOpen ||
      pendingBulkClear !== null ||
      pendingConfirmAction !== null ||
      resizeColDialog !== null
    ) {
      return;
    }
    if (isTextInputElement(e.target) || isTextInputElement(document.activeElement)) {
      return;
    }

    const text = e.clipboardData.getData('text/plain') || internalClipboardRef.current?.text || '';
    if (!text) return;

    e.preventDefault();
    applyPasteTextToTarget(text, selection);
  };

  const parseJsonSafe = async (res: Response): Promise<Record<string, unknown>> => {
    const text = await res.text();
    if (!text) return {};
    try {
      return JSON.parse(text) as Record<string, unknown>;
    } catch {
      return { error: text.slice(0, 200) };
    }
  };



  const isTransientSaveFailure = (status: number): boolean => {
    return status === 0 || status === 408 || status === 429 || status >= 500;
  };

  const flushSaveQueue = useCallback(
    async (options?: {
      source?: 'autosave' | 'manual' | 'retry' | 'export' | 'structural';
    }): Promise<boolean> => {
      const source = options?.source ?? 'manual';

      if (source === 'manual' || source === 'retry' || source === 'export' || source === 'structural') {
        clearSaveTimers();
        retryAttemptRef.current = 0;
        if (isEditingRef.current) {
          commitEditRef.current(0, 0);
        }
      }

      if (inFlightSavePromiseRef.current) {
        flushAgainAfterFlightRef.current = true;
        const firstOk = await inFlightSavePromiseRef.current;
        if (hasAnyDirtyWork() && !hasConflictRef.current && (typeof navigator === 'undefined' || navigator.onLine !== false)) {
          return flushSaveQueueRef.current({ source });
        }
        return firstOk;
      }

      const pendingHeadersMap = { ...dirtyHeadersRef.current };
      const pendingCellsMap = { ...dirtyCellsRef.current };
      const headerEdits = Object.entries(pendingHeadersMap);
      const cellEdits = Object.entries(pendingCellsMap);
      const pendingLayouts = computePendingLayouts();

      if (headerEdits.length === 0 && cellEdits.length === 0 && pendingLayouts.length === 0) {
        if (saveStatusRef.current === 'error' || saveStatusRef.current === 'pending') {
          setSaveStatus('idle');
          setSaveError(null);
        }
        return true;
      }

      if (typeof navigator !== 'undefined' && navigator.onLine === false) {
        setSaveStatus('pending');
        setSaveMessage('');
        if (source === 'manual' || source === 'retry' || source === 'export' || source === 'structural') {
          setSaveError('You are offline. Changes are kept locally and will auto-save when reconnected.');
        }
        return false;
      }

      const runSave = async (): Promise<boolean> => {
        isSavingRef.current = true;
        flushAgainAfterFlightRef.current = false;
        setSaveStatus('saving');
        setSaveMessage('');
        setSaveError(null);
        setHasConflict(false);

        try {
          // 1. Save pending header edits per sheet sequentially so headerVersion stays synchronized
          const headersBySheet = new Map<string, Array<{ key: string; edit: DirtyHeaderEdit }>>();
          for (const [key, edit] of headerEdits) {
            const list = headersBySheet.get(edit.sheetId) || [];
            list.push({ key, edit });
            headersBySheet.set(edit.sheetId, list);
          }

          for (const [sheetId, entries] of headersBySheet.entries()) {
            entries.sort((a, b) => a.edit.colIndex - b.edit.colIndex);
            let runningHeaderVersion = entries[0].edit.expectedHeaderVersion;

            for (const { key, edit } of entries) {
              const res = await fetch('/api/records', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                  workspace,
                  action: 'patch_header',
                  sheetId,
                  colIndex: edit.colIndex,
                  newHeader: edit.newHeader,
                  expectedHeaderVersion: runningHeaderVersion
                })
              });

              const data = await parseJsonSafe(res);
              if (!res.ok) {
                if (res.status === 409) {
                  setHasConflict(true);
                }
                const err = new Error(
                  typeof data.error === 'string' && data.error
                    ? data.error
                    : 'Failed to save header change.'
                ) as Error & { httpStatus?: number };
                err.httpStatus = res.status;
                throw err;
              }

              runningHeaderVersion = Number(data.newHeaderVersion ?? runningHeaderVersion + 1);

              const updatedDirtyHeaders = { ...dirtyHeadersRef.current };
              const latestHeaderEdit = updatedDirtyHeaders[key];
              if (latestHeaderEdit && latestHeaderEdit.newHeader === edit.newHeader) {
                delete updatedDirtyHeaders[key];
              } else if (latestHeaderEdit) {
                // User edited this same header column again while save was in flight: preserve newer edit with fresh baseline
                updatedDirtyHeaders[key] = {
                  ...latestHeaderEdit,
                  originalHeader: edit.newHeader,
                  expectedHeaderVersion: runningHeaderVersion
                };
              }

              for (const remKey of Object.keys(updatedDirtyHeaders)) {
                if (updatedDirtyHeaders[remKey].sheetId === sheetId) {
                  updatedDirtyHeaders[remKey] = {
                    ...updatedDirtyHeaders[remKey],
                    expectedHeaderVersion: runningHeaderVersion
                  };
                }
              }
              dirtyHeadersRef.current = updatedDirtyHeaders;
              setDirtyHeaders(updatedDirtyHeaders);

              setSheets((prev) => {
                const target = prev[edit.sheetName];
                if (!target) return prev;
                let mergedHeaders = target.headers;
                if (Array.isArray(data.headers) && data.headers.length > 0) {
                  const serverHeaders = [...(data.headers as string[])];
                  for (const hEdit of Object.values(dirtyHeadersRef.current)) {
                    if (hEdit.sheetId === sheetId) {
                      while (serverHeaders.length <= hEdit.colIndex) serverHeaders.push('');
                      serverHeaders[hEdit.colIndex] = hEdit.newHeader;
                    }
                  }
                  mergedHeaders = serverHeaders;
                }
                return {
                  ...prev,
                  [edit.sheetName]: {
                    ...target,
                    headerVersion: runningHeaderVersion,
                    headers: mergedHeaders
                  }
                };
              });
            }
          }

          // 2. Save pending cell edits per sheet via patch_cell (1 edit) or patch_cells_batch (2..N edits)
          const cellsBySheet = new Map<string, Array<{ key: string; edit: DirtyCellEdit }>>();
          for (const [key, edit] of cellEdits) {
            const list = cellsBySheet.get(edit.sheetId) || [];
            list.push({ key, edit });
            cellsBySheet.set(edit.sheetId, list);
          }

          for (const [sheetId, entries] of cellsBySheet.entries()) {
            for (let offset = 0; offset < entries.length; offset += 500) {
              const chunk = entries.slice(offset, offset + 500);
              const sheetName = chunk[0].edit.sheetName;

              if (chunk.length === 1) {
                const [{ key, edit }] = chunk;
                const res = await fetch('/api/records', {
                  method: 'POST',
                  headers: { 'Content-Type': 'application/json' },
                  body: JSON.stringify({
                    workspace,
                    action: 'patch_cell',
                    sheetId,
                    rowIndex: edit.rowIndex,
                    colIndex: edit.colIndex,
                    value: edit.value,
                    expectedCellVersion: edit.expectedCellVersion
                  })
                });

                const data = await parseJsonSafe(res);
                if (!res.ok) {
                  if (res.status === 409) {
                    setHasConflict(true);
                  }
                  const err = new Error(
                    typeof data.error === 'string' && data.error
                      ? data.error
                      : 'Failed to save cell edit.'
                  ) as Error & { httpStatus?: number };
                  err.httpStatus = res.status;
                  throw err;
                }

                const newCellVer = Number(data.newCellVersion ?? edit.expectedCellVersion + 1);
                const newRowVer =
                  data.newRowVersion !== undefined ? Number(data.newRowVersion) : undefined;
                const savedAt =
                  typeof data.updatedAt === 'string' ? data.updatedAt : new Date().toISOString();

                const updatedDirtyCells = { ...dirtyCellsRef.current };
                const latestEdit = updatedDirtyCells[key];
                if (latestEdit && latestEdit.value === edit.value) {
                  delete updatedDirtyCells[key];
                } else if (latestEdit) {
                  updatedDirtyCells[key] = {
                    ...latestEdit,
                    originalValue: edit.value,
                    expectedCellVersion: newCellVer
                  };
                }
                dirtyCellsRef.current = updatedDirtyCells;
                setDirtyCells(updatedDirtyCells);

                setSheets((prev) => {
                  const target = prev[sheetName];
                  if (!target) return prev;
                  const nextCellVersions = { ...(target.cellVersions || {}) };
                  nextCellVersions[edit.rowIndex] = {
                    ...(nextCellVersions[edit.rowIndex] || {}),
                    [edit.colIndex]: newCellVer
                  };
                  const nextRowVersions = { ...(target.rowVersions || {}) };
                  if (newRowVer !== undefined) {
                    nextRowVersions[edit.rowIndex] = newRowVer;
                  }
                  const nextRowUpdatedAt = {
                    ...(target.rowUpdatedAt || {}),
                    [edit.rowIndex]: savedAt
                  };
                  return {
                    ...prev,
                    [sheetName]: {
                      ...target,
                      cellVersions: nextCellVersions,
                      rowVersions: nextRowVersions,
                      rowUpdatedAt: nextRowUpdatedAt
                    }
                  };
                });
              } else {
                const res = await fetch('/api/records', {
                  method: 'POST',
                  headers: { 'Content-Type': 'application/json' },
                  body: JSON.stringify({
                    workspace,
                    action: 'patch_cells_batch',
                    sheetId,
                    edits: chunk.map(({ edit }) => ({
                      rowIndex: edit.rowIndex,
                      colIndex: edit.colIndex,
                      value: edit.value,
                      expectedCellVersion: edit.expectedCellVersion
                    }))
                  })
                });

                const data = await parseJsonSafe(res);
                if (!res.ok) {
                  if (res.status === 409) {
                    setHasConflict(true);
                  }
                  const err = new Error(
                    typeof data.error === 'string' && data.error
                      ? data.error
                      : 'Failed to save cell edits batch.'
                  ) as Error & { httpStatus?: number };
                  err.httpStatus = res.status;
                  throw err;
                }

                const updatedCellsRes: Array<{
                  rowIndex: number;
                  colIndex: number;
                  newCellVersion: number;
                  newRowVersion?: number;
                }> = Array.isArray(data.updatedCells)
                  ? (data.updatedCells as Array<{
                      rowIndex: number;
                      colIndex: number;
                      newCellVersion: number;
                      newRowVersion?: number;
                    }>)
                  : [];
                const savedAt =
                  typeof data.updatedAt === 'string' ? data.updatedAt : new Date().toISOString();

                const serverVerMap = new Map<string, { cellVer: number; rowVer?: number }>();
                updatedCellsRes.forEach((u, idx) => {
                  const fallbackEdit = chunk[idx]?.edit;
                  const rIdx = u.rowIndex ?? fallbackEdit?.rowIndex;
                  const cIdx = u.colIndex ?? fallbackEdit?.colIndex;
                  if (rIdx === undefined || cIdx === undefined) return;
                  serverVerMap.set(`${rIdx}:${cIdx}`, {
                    cellVer: Number(
                      u.newCellVersion ?? (fallbackEdit ? fallbackEdit.expectedCellVersion + 1 : 1)
                    ),
                    rowVer: u.newRowVersion !== undefined ? Number(u.newRowVersion) : undefined
                  });
                });

                const updatedDirtyCells = { ...dirtyCellsRef.current };
                for (const { key, edit } of chunk) {
                  const serverInfo = serverVerMap.get(`${edit.rowIndex}:${edit.colIndex}`);
                  const newCellVer = serverInfo?.cellVer ?? edit.expectedCellVersion + 1;
                  const latestEdit = updatedDirtyCells[key];
                  if (latestEdit && latestEdit.value === edit.value) {
                    delete updatedDirtyCells[key];
                  } else if (latestEdit) {
                    updatedDirtyCells[key] = {
                      ...latestEdit,
                      originalValue: edit.value,
                      expectedCellVersion: newCellVer
                    };
                  }
                }
                dirtyCellsRef.current = updatedDirtyCells;
                setDirtyCells(updatedDirtyCells);

                setSheets((prev) => {
                  const target = prev[sheetName];
                  if (!target) return prev;
                  const nextCellVersions = { ...(target.cellVersions || {}) };
                  const nextRowVersions = { ...(target.rowVersions || {}) };
                  const nextRowUpdatedAt = { ...(target.rowUpdatedAt || {}) };

                  for (const { edit } of chunk) {
                    const serverInfo = serverVerMap.get(`${edit.rowIndex}:${edit.colIndex}`);
                    const newCellVer = serverInfo?.cellVer ?? edit.expectedCellVersion + 1;
                    nextCellVersions[edit.rowIndex] = {
                      ...(nextCellVersions[edit.rowIndex] || {}),
                      [edit.colIndex]: newCellVer
                    };
                    if (serverInfo?.rowVer !== undefined) {
                      nextRowVersions[edit.rowIndex] = serverInfo.rowVer;
                    }
                    nextRowUpdatedAt[edit.rowIndex] = savedAt;
                  }

                  return {
                    ...prev,
                    [sheetName]: {
                      ...target,
                      cellVersions: nextCellVersions,
                      rowVersions: nextRowVersions,
                      rowUpdatedAt: nextRowUpdatedAt
                    }
                  };
                });
              }
            }
          }

          // 3. Persist changed worksheet column widths & row heights
          for (const layoutEdit of pendingLayouts) {
            const res = await fetch('/api/records', {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({
                workspace,
                action: 'save_sheet_layout',
                sheetId: layoutEdit.sheetId,
                columnWidths: layoutEdit.columnWidths,
                rowHeights: layoutEdit.rowHeights
              })
            });

            const data = await parseJsonSafe(res);
            if (!res.ok) {
              const err = new Error(
                typeof data.error === 'string' && data.error
                  ? data.error
                  : 'Failed to save column width or row height layout.'
              ) as Error & { httpStatus?: number };
              err.httpStatus = res.status;
              throw err;
            }

            const savedCols =
              data.columnWidths && typeof data.columnWidths === 'object'
                ? (data.columnWidths as Record<number, number>)
                : layoutEdit.columnWidths;
            const savedRows =
              data.rowHeights && typeof data.rowHeights === 'object'
                ? (data.rowHeights as Record<number, number>)
                : layoutEdit.rowHeights;

            savedLayoutBySheetRef.current[layoutEdit.dimKey] = {
              cols: { ...savedCols },
              rows: { ...savedRows }
            };
            setSavedLayoutBySheet((prev) => ({
              ...prev,
              [layoutEdit.dimKey]: {
                cols: { ...savedCols },
                rows: { ...savedRows }
              }
            }));

            setSheets((prev) => {
              const target = prev[layoutEdit.sheetName];
              if (!target) return prev;
              return {
                ...prev,
                [layoutEdit.sheetName]: {
                  ...target,
                  columnWidths: { ...savedCols },
                  rowHeights: { ...savedRows }
                }
              };
            });
          }

          retryAttemptRef.current = 0;
          const remainingUnsaved = hasAnyDirtyWork();

          if (remainingUnsaved || flushAgainAfterFlightRef.current) {
            setSaveStatus('pending');
            setSaveMessage('');
            if (typeof navigator === 'undefined' || navigator.onLine !== false) {
              scheduleAutosave(250);
            }
          } else {
            setSaveStatus('saved');
            setSaveMessage('Saved');
          }
          return true;
        } catch (err: unknown) {
          const httpStatus = (err as { httpStatus?: number } | null)?.httpStatus;
          const msg = err instanceof Error ? err.message : 'Save error occurred';
          setSaveStatus('error');
          setSaveMessage('');
          setSaveError(msg);

          if (
            isTransientSaveFailure(httpStatus ?? 500) &&
            !hasConflictRef.current &&
            retryAttemptRef.current < 3 &&
            (typeof navigator === 'undefined' || navigator.onLine !== false)
          ) {
            const backoffMs = [3000, 8000, 20000][retryAttemptRef.current] ?? 20000;
            retryAttemptRef.current += 1;
            retryTimerRef.current = window.setTimeout(() => {
              retryTimerRef.current = null;
              void flushSaveQueueRef.current({ source: 'retry' });
            }, backoffMs);
          }

          return false;
        } finally {
          isSavingRef.current = false;
          inFlightSavePromiseRef.current = null;
        }
      };

      const flightPromise = runSave();
      inFlightSavePromiseRef.current = flightPromise;
      return flightPromise;
    },
    [workspace, computePendingLayouts, clearSaveTimers, hasAnyDirtyWork, scheduleAutosave]
  );

  useEffect(() => {
    flushSaveQueueRef.current = flushSaveQueue;
  }, [flushSaveQueue]);

  const handleSave = useCallback(async () => {
    await flushSaveQueue({ source: 'manual' });
  }, [flushSaveQueue]);

  const handleSaveRef = useRef(handleSave);
  useEffect(() => {
    handleSaveRef.current = handleSave;
  });

  // Track online/offline transitions so queued edits recover immediately when connection returns
  useEffect(() => {
    const syncOnline = () => {
      const online = typeof navigator === 'undefined' ? true : navigator.onLine;
      setIsOnline(online);
      if (online && hasAnyDirtyWork() && !hasConflictRef.current && !isSavingRef.current) {
        retryAttemptRef.current = 0;
        void flushSaveQueueRef.current({ source: 'autosave' });
      }
    };

    const rafId = window.requestAnimationFrame(() => {
      setIsOnline(typeof navigator === 'undefined' ? true : navigator.onLine);
    });
    const handleOffline = () => {
      setIsOnline(false);
      clearSaveTimers();
      if (hasAnyDirtyWork()) {
        setSaveStatus('pending');
        setSaveMessage('');
      }
    };

    window.addEventListener('online', syncOnline);
    window.addEventListener('offline', handleOffline);
    return () => {
      window.cancelAnimationFrame(rafId);
      window.removeEventListener('online', syncOnline);
      window.removeEventListener('offline', handleOffline);
    };
  }, [clearSaveTimers, hasAnyDirtyWork]);

  // Best-effort keepalive flush when hiding tab/navigating away + beforeunload guard while unsaved
  useEffect(() => {
    if (!isActiveWorkspace) return;

    const attemptKeepaliveFlush = () => {
      if (isSavingRef.current || hasConflictRef.current) return;
      if (typeof navigator !== 'undefined' && navigator.onLine === false) return;

      const cellEntries = Object.values(dirtyCellsRef.current);
      const headerEntries = Object.values(dirtyHeadersRef.current);
      const layoutEntries = computePendingLayouts();
      if (cellEntries.length === 0 && headerEntries.length === 0 && layoutEntries.length === 0) {
        return;
      }

      clearSaveTimers();

      // Dispatch keepalive requests for small pending batches (browsers enforce ~64KB keepalive limit)
      for (const h of headerEntries) {
        const bodyStr = JSON.stringify({
          workspace,
          action: 'patch_header',
          sheetId: h.sheetId,
          colIndex: h.colIndex,
          newHeader: h.newHeader,
          expectedHeaderVersion: h.expectedHeaderVersion
        });
        if (bodyStr.length < 48000) {
          void fetch('/api/records', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            keepalive: true,
            body: bodyStr
          }).catch(() => {});
        }
      }

      const cellsBySheet = new Map<string, DirtyCellEdit[]>();
      for (const c of cellEntries) {
        const list = cellsBySheet.get(c.sheetId) || [];
        list.push(c);
        cellsBySheet.set(c.sheetId, list);
      }
      for (const [sheetId, edits] of cellsBySheet.entries()) {
        if (edits.length === 1) {
          const c = edits[0];
          const bodyStr = JSON.stringify({
            workspace,
            action: 'patch_cell',
            sheetId,
            rowIndex: c.rowIndex,
            colIndex: c.colIndex,
            value: c.value,
            expectedCellVersion: c.expectedCellVersion
          });
          if (bodyStr.length < 48000) {
            void fetch('/api/records', {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              keepalive: true,
              body: bodyStr
            }).catch(() => {});
          }
        } else if (edits.length <= 200) {
          const bodyStr = JSON.stringify({
            workspace,
            action: 'patch_cells_batch',
            sheetId,
            edits: edits.map((c) => ({
              rowIndex: c.rowIndex,
              colIndex: c.colIndex,
              value: c.value,
              expectedCellVersion: c.expectedCellVersion
            }))
          });
          if (bodyStr.length < 48000) {
            void fetch('/api/records', {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              keepalive: true,
              body: bodyStr
            }).catch(() => {});
          }
        }
      }

      for (const l of layoutEntries) {
        const bodyStr = JSON.stringify({
          workspace,
          action: 'save_sheet_layout',
          sheetId: l.sheetId,
          columnWidths: l.columnWidths,
          rowHeights: l.rowHeights
        });
        if (bodyStr.length < 48000) {
          void fetch('/api/records', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            keepalive: true,
            body: bodyStr
          }).catch(() => {});
        }
      }
    };

    const handleVisibilityChange = () => {
      if (document.visibilityState === 'hidden') {
        attemptKeepaliveFlush();
      }
    };

    const handlePageHide = () => {
      attemptKeepaliveFlush();
    };

    const handleBeforeUnload = (e: BeforeUnloadEvent) => {
      if (hasAnyDirtyWork() || isSavingRef.current || saveStatusRef.current === 'error') {
        e.preventDefault();
        e.returnValue = '';
      }
    };

    document.addEventListener('visibilitychange', handleVisibilityChange);
    window.addEventListener('pagehide', handlePageHide);
    window.addEventListener('beforeunload', handleBeforeUnload);
    return () => {
      document.removeEventListener('visibilitychange', handleVisibilityChange);
      window.removeEventListener('pagehide', handlePageHide);
      window.removeEventListener('beforeunload', handleBeforeUnload);
    };
  }, [isActiveWorkspace, workspace, computePendingLayouts, clearSaveTimers, hasAnyDirtyWork]);

  useEffect(() => {
    if (!isActiveWorkspace) return;
    const onWindowKeyDown = (e: KeyboardEvent) => {
      const isMod = e.ctrlKey || e.metaKey;
      if (!isMod) return;
      const lowerKey = e.key.toLowerCase();

      if (lowerKey === 's') {
        e.preventDefault();
        if (
          renamingSheetNameRef.current !== null ||
          isTourOpenRef.current ||
          pendingBulkClearRef.current !== null ||
          pendingConfirmActionRef.current !== null ||
          resizeColDialogRef.current !== null
        ) {
          return;
        }
        void handleSaveRef.current();
        return;
      }

      if (
        isEditingRef.current ||
        renamingSheetNameRef.current !== null ||
        isTourOpenRef.current ||
        pendingBulkClearRef.current !== null ||
        pendingConfirmActionRef.current !== null ||
        resizeColDialogRef.current !== null ||
        isTextInputElement(e.target)
      ) {
        return;
      }

      if (lowerKey === 'z' && !e.shiftKey) {
        e.preventDefault();
        handleUndoRef.current();
      } else if (lowerKey === 'y' || (lowerKey === 'z' && e.shiftKey)) {
        e.preventDefault();
        handleRedoRef.current();
      }
    };
    window.addEventListener('keydown', onWindowKeyDown);
    return () => window.removeEventListener('keydown', onWindowKeyDown);
  }, [isActiveWorkspace]);

  const handleReloadLatest = async () => {
    if (!onReloadWorkspace) return;
    clearSaveTimers();
    setIsReloading(true);
    try {
      await onReloadWorkspace();
      setHasConflict(false);
      setSaveError(null);
      setSaveStatus('idle');
    } finally {
      setIsReloading(false);
    }
  };

  const [isExporting, setIsExporting] = useState(false);

  // Direct single-click Export button: flushes pending changes first if any, then exports workspace
  const handleDirectExport = async () => {
    if (isExporting || sheetKeys.length === 0) return;
    setIsExporting(true);
    try {
      if (hasAnyDirtyWork() || isSavingRef.current) {
        const flushed = await flushSaveQueue({ source: 'export' });
        if (!flushed || hasAnyDirtyWork()) {
          throw new Error('Could not save pending edits before export.');
        }
      }

      const targetSheets: Array<{ name: string; headers: string[]; rows: (string | number)[][] }> = [];
      for (const name of sheetKeys) {
        const s = sheets[name];
        if (!s) continue;
        const nonEmptyRows = s.rows.filter(
          (r) => r && r.some((cell) => cell !== '' && cell !== undefined && cell !== null)
        );
        targetSheets.push({
          name,
          headers: s.headers,
          rows: nonEmptyRows
        });
      }

      exportSheetsToExcel(workspace, targetSheets, 'Workspace');
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : 'Export failed.';
      setOperationError(`Export failed: ${msg}`);
    } finally {
      setIsExporting(false);
    }
  };

  // Helper to determine next available non-conflicting default sheet name
  const getNextAvailableSheetName = useCallback((existingNames: string[]): string => {
    const lowerSet = new Set(existingNames.map((n) => n.trim().toLowerCase()));
    let idx = 1;
    while (lowerSet.has(`sheet ${idx}`.toLowerCase()) || lowerSet.has(`sheet${idx}`.toLowerCase())) {
      idx++;
    }
    return `Sheet ${idx}`;
  }, []);

  // 6. Immediate Add Worksheet via '+' button (no modal)
  const handleAddSheetImmediate = async () => {
    if (isCreatingSheet) return;
    const currentKeys = sheetKeysRef.current;
    const newName = getNextAvailableSheetName(currentKeys);
    setIsCreatingSheet(true);
    setOperationError(null);

    try {
      const defaultHeaders =
        workspace === 'doctor'
          ? ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P']
          : [
              'Date',
              'Recorded_By',
              'Phone_No',
              'Examined_By',
              'Name',
              'Date_of_Birth',
              'Test_Type',
              'Test_Fee',
              'Drops_No',
              'Medication',
              'Total',
              'Glass Total',
              'Paid',
              'Glasses BAL',
              'Review Date',
              'Review Test'
            ];

      const res = await fetch('/api/records', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          workspace,
          action: 'create_sheet',
          sheetName: newName,
          headers: defaultHeaders
        })
      });

      const data = await parseJsonSafe(res);
      if (!res.ok) {
        throw new Error(
          typeof data.error === 'string' && data.error
            ? data.error
            : 'Failed to create worksheet on server.'
        );
      }

      const rawCreated = (data.sheet ?? {}) as Partial<SheetState>;
      const maxExistingOrder = currentKeys.reduce((maxOrder, k) => {
        const order = sheetsRef.current[k]?.tabOrder ?? 0;
        return Math.max(maxOrder, order);
      }, -1);

      const createdSheet: SheetState = {
        ...rawCreated,
        id: rawCreated.id,
        name: rawCreated.name || newName,
        tabOrder:
          typeof rawCreated.tabOrder === 'number' && rawCreated.tabOrder > maxExistingOrder
            ? rawCreated.tabOrder
            : maxExistingOrder + 1,
        headers: Array.isArray(rawCreated.headers) && rawCreated.headers.length > 0 ? rawCreated.headers : defaultHeaders,
        headerVersion: rawCreated.headerVersion ?? 1,
        rows: Array.isArray(rawCreated.rows) ? rawCreated.rows : [],
        rowVersions: rawCreated.rowVersions ?? {},
        cellVersions: rawCreated.cellVersions ?? {},
        rowUpdatedAt: rawCreated.rowUpdatedAt ?? {}
      };

      const resolvedName = createdSheet.name || newName;
      setSheets((prev) => ({
        ...prev,
        [resolvedName]: createdSheet
      }));

      setActiveSheetName(resolvedName);
      setSelection({
        mode: 'cell',
        anchorRow: 1,
        anchorCol: 0,
        focusRow: 1,
        focusCol: 0
      });
      setScrollPositionRequest((prev) => ({
        sheetName: resolvedName,
        targetRow: 1,
        nonce: prev.nonce + 1
      }));
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : 'Failed to create worksheet.';
      setOperationError(msg);
    } finally {
      setIsCreatingSheet(false);
    }
  };

  // 7. Worksheet Renaming via Double-Click or Context Menu (strict focus & keystroke isolation)
  const handleStartRename = useCallback((targetName: string) => {
    if (isEditingRef.current) {
      commitEditRef.current(0, 0);
    }
    renameCommitGuardRef.current = false;
    setRenamingSheetName(targetName);
    setRenameInputValue(targetName);
    setContextMenu(null);
  }, []);

  const handleCommitRename = async (oldName: string, rawNewName: string) => {
    if (renameCommitGuardRef.current) {
      return;
    }
    renameCommitGuardRef.current = true;

    const trimmed = rawNewName.trim();
    setRenamingSheetName(null);

    if (!trimmed || trimmed === oldName) {
      rootRef.current?.focus({ preventScroll: true });
      return;
    }

    const validation = validateWorksheetName(trimmed);
    if (!validation.valid) {
      setOperationError(validation.error);
      rootRef.current?.focus({ preventScroll: true });
      return;
    }

    const otherNames = sheetKeys.filter((n) => n !== oldName);
    if (otherNames.some((n) => n.toLowerCase() === trimmed.toLowerCase())) {
      setOperationError(`A worksheet named "${trimmed}" already exists.`);
      rootRef.current?.focus({ preventScroll: true });
      return;
    }

    const targetSheet = sheets[oldName];
    if (!targetSheet?.id) {
      setOperationError('Cannot rename worksheet because its server identifier is missing. Try reloading the workspace.');
      rootRef.current?.focus({ preventScroll: true });
      return;
    }

    setOperationError(null);
    setIsRenamingSheet(true);
    try {
      const res = await fetch('/api/records', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          workspace,
          action: 'rename_sheet',
          sheetId: targetSheet.id,
          newName: trimmed
        })
      });

      const data = await parseJsonSafe(res);
      if (!res.ok) {
        throw new Error(
          typeof data.error === 'string' && data.error
            ? data.error
            : 'Failed to rename worksheet.'
        );
      }

      setSheets((prev) => {
        const next: Record<string, SheetState> = {};
        for (const [k, v] of Object.entries(prev)) {
          if (k === oldName) {
            next[trimmed] = { ...v, name: trimmed };
          } else {
            next[k] = v;
          }
        }
        return next;
      });

      // Update any unsaved cell/header sheetName references without invalidating sheetId
      const updatedDirtyCells = { ...dirtyCellsRef.current };
      let cellNamesUpdated = false;
      for (const [k, edit] of Object.entries(updatedDirtyCells)) {
        if (edit.sheetId === targetSheet.id && edit.sheetName === oldName) {
          updatedDirtyCells[k] = { ...edit, sheetName: trimmed };
          cellNamesUpdated = true;
        }
      }
      if (cellNamesUpdated) {
        dirtyCellsRef.current = updatedDirtyCells;
        setDirtyCells(updatedDirtyCells);
      }

      const updatedDirtyHeaders = { ...dirtyHeadersRef.current };
      let headerNamesUpdated = false;
      for (const [k, edit] of Object.entries(updatedDirtyHeaders)) {
        if (edit.sheetId === targetSheet.id && edit.sheetName === oldName) {
          updatedDirtyHeaders[k] = { ...edit, sheetName: trimmed };
          headerNamesUpdated = true;
        }
      }
      if (headerNamesUpdated) {
        dirtyHeadersRef.current = updatedDirtyHeaders;
        setDirtyHeaders(updatedDirtyHeaders);
      }

      if (activeSheetName === oldName) {
        setActiveSheetName(trimmed);
      }
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : 'Failed to rename worksheet.';
      setOperationError(msg);
    } finally {
      setIsRenamingSheet(false);
      rootRef.current?.focus({ preventScroll: true });
    }
  };

  // 8. Immediate Worksheet Deletion from Right-Click Context Menu (NO confirmation dialog)
  const handleDeleteSheetImmediate = useCallback(
    async (targetSheetName: string) => {
      if (isDeletingSheet) return;
      const currentKeys = sheetKeysRef.current;

    if (currentKeys.length <= 1) {
      setContextMenu(null);
      setOperationError(`Cannot delete the only remaining worksheet in the ${workspace} workspace.`);
      return;
    }

    const targetSheet = sheetsRef.current[targetSheetName];
    if (!targetSheet?.id) {
      setContextMenu(null);
      setOperationError('Worksheet cannot be deleted because its server identifier is missing. Try reloading the workspace.');
      return;
    }

    setOperationError(null);
    setIsDeletingSheet(true);
    setDeletingSheetName(targetSheetName);

    try {
      const res = await fetch('/api/records', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          workspace,
          action: 'delete_sheet',
          sheetId: targetSheet.id
        })
      });

      const data = await parseJsonSafe(res);
      if (!res.ok) {
        throw new Error(
          typeof data.error === 'string' && data.error
            ? data.error
            : 'Failed to delete worksheet.'
        );
      }

      // Discard any unsaved pending edits tied to the deleted worksheet ID
      const nextDirtyCells = { ...dirtyCellsRef.current };
      let dirtyCellsChanged = false;
      for (const [k, edit] of Object.entries(nextDirtyCells)) {
        if (edit.sheetId === targetSheet.id || edit.sheetName === targetSheetName) {
          delete nextDirtyCells[k];
          dirtyCellsChanged = true;
        }
      }
      if (dirtyCellsChanged) {
        dirtyCellsRef.current = nextDirtyCells;
        setDirtyCells(nextDirtyCells);
      }

      const nextDirtyHeaders = { ...dirtyHeadersRef.current };
      let dirtyHeadersChanged = false;
      for (const [k, edit] of Object.entries(nextDirtyHeaders)) {
        if (edit.sheetId === targetSheet.id || edit.sheetName === targetSheetName) {
          delete nextDirtyHeaders[k];
          dirtyHeadersChanged = true;
        }
      }
      if (dirtyHeadersChanged) {
        dirtyHeadersRef.current = nextDirtyHeaders;
        setDirtyHeaders(nextDirtyHeaders);
      }

      const deletedIdx = currentKeys.indexOf(targetSheetName);
      const remainingKeys = currentKeys.filter((k) => k !== targetSheetName);

      setSheets((prev) => {
        const orderedRemaining = Object.keys(prev)
          .filter((k) => k !== targetSheetName)
          .sort((a, b) => (prev[a]?.tabOrder ?? 0) - (prev[b]?.tabOrder ?? 0));
        const next: Record<string, SheetState> = {};
        orderedRemaining.forEach((key, idx) => {
          next[key] = { ...prev[key], tabOrder: idx };
        });
        return next;
      });

      if (activeSheetName === targetSheetName && remainingKeys.length > 0) {
        const adjacentIdx = Math.min(deletedIdx >= 0 ? deletedIdx : 0, remainingKeys.length - 1);
        const adjacentSheetName = remainingKeys[adjacentIdx];
        const adjacentLatestRow = findLatestRecordGridRow(sheetsRef.current[adjacentSheetName]);
        setActiveSheetName(adjacentSheetName);
        setSelection({
          mode: 'cell',
          anchorRow: adjacentLatestRow,
          anchorCol: 0,
          focusRow: adjacentLatestRow,
          focusCol: 0
        });
        setScrollPositionRequest((prev) => ({
          sheetName: adjacentSheetName,
          targetRow: adjacentLatestRow,
          nonce: prev.nonce + 1
        }));
      }

      setContextMenu(null);
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : 'Failed to delete worksheet.';
      setOperationError(msg);
      setContextMenu(null);
    } finally {
      setIsDeletingSheet(false);
      setDeletingSheetName(null);
    }
  },
    [activeSheetName, isDeletingSheet, workspace]
  );

  // 9. Worksheet Reordering via Pointer Events + Server Persistence
  const commitSheetOrder = useCallback(
    async (orderedKeys: string[]) => {
      const previousSheets = sheetsRef.current;
      const orderedSheetIds = orderedKeys
        .map((k) => previousSheets[k]?.id)
        .filter((id): id is string => typeof id === 'string' && id.length > 0);

      // Rebuild sheets object in the new order with updated tabOrder indices
      setSheets((prev) => {
        const next: Record<string, SheetState> = {};
        orderedKeys.forEach((key, idx) => {
          if (prev[key]) {
            next[key] = { ...prev[key], tabOrder: idx };
          }
        });
        // Preserve any keys not in orderedKeys defensively
        for (const [k, v] of Object.entries(prev)) {
          if (!next[k]) {
            next[k] = v;
          }
        }
        return next;
      });

      if (orderedSheetIds.length === 0) return;

      setOperationError(null);
      try {
        const res = await fetch('/api/records', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            workspace,
            action: 'reorder_sheets',
            sheetIds: orderedSheetIds
          })
        });
        const data = await parseJsonSafe(res);
        if (!res.ok) {
          throw new Error(
            typeof data.error === 'string' && data.error
              ? data.error
              : 'Failed to persist worksheet order.'
          );
        }
      } catch (err: unknown) {
        // Revert optimistic tab order on server error so UI matches persisted DB state
        setSheets(previousSheets);
        const msg = err instanceof Error ? err.message : 'Failed to reorder worksheets.';
        setOperationError(msg);
      }
    },
    [workspace]
  );

  const openTabContextMenuFromRect = useCallback(
    (sheetName: string, anchorX: number, anchorY: number) => {
      if (isDeletingSheet) return;
      const currentKeys = sheetKeysRef.current;
      const idx = currentKeys.indexOf(sheetName);
      const canMoveLeft = idx > 0;
      const canMoveRight = idx !== -1 && idx < currentKeys.length - 1;
      const canDelete = currentKeys.length > 1;
      const isDeletingThis = isDeletingSheet && deletingSheetName === sheetName;

      setContextMenu({
        kind: 'tab',
        x: anchorX,
        y: anchorY,
        sheetName,
        canMoveLeft,
        canMoveRight,
        canDelete,
        isDeletingThis
      });
    },
    [isDeletingSheet, deletingSheetName]
  );

  const handleMoveSheetRelative = useCallback(
    (sheetName: string, direction: -1 | 1) => {
      const currentKeys = sheetKeysRef.current;
      const idx = currentKeys.indexOf(sheetName);
      if (idx === -1) return;
      const targetIdx = idx + direction;
      if (targetIdx < 0 || targetIdx >= currentKeys.length) return;
      const reordered = [...currentKeys];
      const [moved] = reordered.splice(idx, 1);
      reordered.splice(targetIdx, 0, moved);
      setContextMenu(null);
      void commitSheetOrder(reordered);
    },
    [commitSheetOrder]
  );

  const handleTabPointerDown = useCallback(
    (name: string, e: React.PointerEvent<HTMLDivElement>) => {
      if (e.pointerType === 'mouse' && e.button !== 0) return;
      if (renamingSheetName === name || isDeletingSheet) return;

      // On touch/pen devices, preserve native horizontal scrolling on the tab strip and trigger context menu via long press
      if (e.pointerType !== 'mouse') {
        if (tabDragCleanupRef.current) {
          tabDragCleanupRef.current();
        }
        const startX = e.clientX;
        const startY = e.clientY;
        const tabRect = e.currentTarget.getBoundingClientRect();
        let cancelled = false;

        const timerId = window.setTimeout(() => {
          if (cancelled) return;
          suppressNextTabClickRef.current = true;
          openTabContextMenuFromRect(name, tabRect.left, tabRect.top);
          window.setTimeout(() => {
            suppressNextTabClickRef.current = false;
          }, 250);
          cleanup();
        }, 460);

        const onMove = (moveEv: PointerEvent) => {
          if (
            Math.abs(moveEv.clientX - startX) > 8 ||
            Math.abs(moveEv.clientY - startY) > 8
          ) {
            cancelled = true;
            cleanup();
          }
        };

        const onEnd = () => {
          cancelled = true;
          cleanup();
        };

        const cleanup = () => {
          window.clearTimeout(timerId);
          window.removeEventListener('pointermove', onMove);
          window.removeEventListener('pointerup', onEnd);
          window.removeEventListener('pointercancel', onEnd);
          window.removeEventListener('blur', onEnd);
          if (tabDragCleanupRef.current === cleanup) {
            tabDragCleanupRef.current = null;
          }
        };

        tabDragCleanupRef.current = cleanup;
        window.addEventListener('pointermove', onMove, { passive: true });
        window.addEventListener('pointerup', onEnd, { passive: true });
        window.addEventListener('pointercancel', onEnd, { passive: true });
        window.addEventListener('blur', onEnd);
        return;
      }

      const currentKeys = sheetKeysRef.current;
      if (currentKeys.length <= 1) return;

      if (tabDragCleanupRef.current) {
        tabDragCleanupRef.current();
      }

      const sourceIdx = currentKeys.indexOf(name);
      if (sourceIdx === -1) return;

      // Snapshot initial tab bounding boxes before any visual shifting occurs
      const initialRects = currentKeys.map((k) => {
        const el = tabElementRefs.current[k];
        const rect = el?.getBoundingClientRect();
        return {
          name: k,
          left: rect?.left ?? 0,
          right: rect?.right ?? 0,
          midX: rect ? rect.left + rect.width / 2 : 0
        };
      });

      const startX = e.clientX;
      let isDragging = false;
      let currentTargetIdx = sourceIdx;

      const computeTargetIndex = (clientX: number): number => {
        const otherTabs = initialRects.filter((r) => r.name !== name);
        let insertIdx = 0;
        for (const tab of otherTabs) {
          if (clientX > tab.midX) {
            insertIdx++;
          }
        }
        return insertIdx;
      };

      const handlePointerMove = (moveEvent: PointerEvent) => {
        const deltaX = moveEvent.clientX - startX;
        if (!isDragging) {
          if (Math.abs(deltaX) < 6) return;
          isDragging = true;
          suppressNextTabClickRef.current = true;
          setContextMenu(null);
        }

        currentTargetIdx = computeTargetIndex(moveEvent.clientX);
        setDragTabState({
          sourceName: name,
          targetIndex: currentTargetIdx
        });
      };

      const finishDrag = (commit: boolean) => {
        window.removeEventListener('pointermove', handlePointerMove);
        window.removeEventListener('pointerup', onPointerUp);
        window.removeEventListener('pointercancel', onPointerCancel);
        window.removeEventListener('blur', onWindowBlur);
        tabDragCleanupRef.current = null;

        setDragTabState(null);

        if (isDragging) {
          window.setTimeout(() => {
            suppressNextTabClickRef.current = false;
          }, 60);

          if (commit && currentTargetIdx !== sourceIdx) {
            const reordered = currentKeys.filter((k) => k !== name);
            reordered.splice(currentTargetIdx, 0, name);
            void commitSheetOrder(reordered);
          }
        }
      };

      const onPointerUp = () => finishDrag(true);
      const onPointerCancel = () => finishDrag(false);
      const onWindowBlur = () => finishDrag(false);

      tabDragCleanupRef.current = () => finishDrag(false);
      window.addEventListener('pointermove', handlePointerMove);
      window.addEventListener('pointerup', onPointerUp);
      window.addEventListener('pointercancel', onPointerCancel);
      window.addEventListener('blur', onWindowBlur);
    },
    [renamingSheetName, isDeletingSheet, commitSheetOrder, openTabContextMenuFromRect]
  );

  const handleSignOutClick = async () => {
    if (isLoggingOut) return;
    setIsLoggingOut(true);
    setOperationError(null);
    try {
      await onLogout();
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : 'Failed to sign out. Please try again.';
      setOperationError(msg);
      setIsLoggingOut(false);
    }
  };

  const openCellContextMenu = useCallback(
    (row: number, col: number, x: number, y: number) => {
      if (isEditing) commitEdit(0, 0);
      setSelection({
        mode: 'cell',
        anchorRow: row,
        anchorCol: col,
        focusRow: row,
        focusCol: col
      });
      const val = getCellValue(row, col);
      const hasContent = isNonEmptyCell(val);
      const canPaste = Boolean(internalClipboardRef.current?.text);
      const label = `${getColumnLetter(col)}${row + 1}`;
      setContextMenu({
        kind: 'cell',
        x,
        y,
        row,
        col,
        label,
        hasContent,
        canPaste
      });
    },
    [isEditing, commitEdit, getCellValue]
  );

  const openRowContextMenu = useCallback(
    (gridRowIdx: number, x: number, y: number) => {
      if (isEditing) commitEdit(0, 0);
      const maxColIdx = totalCols - 1;
      const isHighlighted = isRowHeaderHighlighted(gridRowIdx);
      const rMin = isHighlighted ? selectionBounds.rMin : gridRowIdx;
      const rMax = isHighlighted ? selectionBounds.rMax : gridRowIdx;

      setSelection({
        mode: 'row',
        anchorRow: rMin,
        anchorCol: 0,
        focusRow: rMax,
        focusCol: maxColIdx
      });

      let hasContent = false;
      for (let r = rMin; r <= rMax; r++) {
        for (let c = 0; c <= maxColIdx; c++) {
          if (isNonEmptyCell(getCellValue(r, c))) {
            hasContent = true;
            break;
          }
        }
        if (hasContent) break;
      }

      const isHeaderRow = rMin === 0;
      const label = rMin === rMax ? `Row ${rMin + 1}` : `Rows ${rMin + 1}–${rMax + 1}`;
      const canPaste = Boolean(internalClipboardRef.current?.text);

      setContextMenu({
        kind: 'row',
        x,
        y,
        startRow: rMin,
        endRow: rMax,
        label,
        isHeaderRow,
        hasContent,
        canPaste
      });
    },
    [isEditing, commitEdit, totalCols, isRowHeaderHighlighted, selectionBounds, getCellValue]
  );

  const openColContextMenu = useCallback(
    (cIdx: number, x: number, y: number) => {
      if (isEditing) commitEdit(0, 0);
      const maxRowIdx = displayRowsCount - 1;
      const isHighlighted = isColHeaderHighlighted(cIdx);
      const cMin = isHighlighted ? selectionBounds.cMin : cIdx;
      const cMax = isHighlighted ? selectionBounds.cMax : cIdx;

      setSelection({
        mode: 'col',
        anchorRow: 0,
        anchorCol: cMin,
        focusRow: maxRowIdx,
        focusCol: cMax
      });

      let hasContent = false;
      for (let r = 0; r <= maxRowIdx; r++) {
        for (let c = cMin; c <= cMax; c++) {
          if (isNonEmptyCell(getCellValue(r, c))) {
            hasContent = true;
            break;
          }
        }
        if (hasContent) break;
      }

      const label = cMin === cMax ? `Column ${getColumnLetter(cMin)}` : `Columns ${getColumnLetter(cMin)}–${getColumnLetter(cMax)}`;
      const canPaste = Boolean(internalClipboardRef.current?.text);
      const canInsert = currentSheet.headers.length < MAX_COLS_LIMIT;
      const canDelete = currentSheet.headers.length > 1;

      setContextMenu({
        kind: 'col',
        x,
        y,
        col: cMin,
        endCol: cMax,
        label,
        hasContent,
        canPaste,
        canInsert,
        canDelete
      });
    },
    [isEditing, commitEdit, displayRowsCount, isColHeaderHighlighted, selectionBounds, getCellValue, currentSheet.headers.length]
  );

  const startGridLongPress = useCallback(
    (e: React.PointerEvent, openMenuFn: (x: number, y: number) => void) => {
      if (e.pointerType === 'mouse' && e.button !== 0) return;
      if (e.pointerType === 'mouse') return;

      if (gridLongPressCleanupRef.current) {
        gridLongPressCleanupRef.current();
      }

      const startX = e.clientX;
      const startY = e.clientY;
      const clientX = e.clientX;
      const clientY = e.clientY;
      let cancelled = false;

      const timerId = window.setTimeout(() => {
        if (cancelled) return;
        justLongPressedRef.current = true;
        openMenuFn(clientX, clientY);
        window.setTimeout(() => {
          justLongPressedRef.current = false;
        }, 250);
        cleanup();
      }, 480);

      const onMove = (moveEv: PointerEvent) => {
        if (Math.abs(moveEv.clientX - startX) > 10 || Math.abs(moveEv.clientY - startY) > 10) {
          cancelled = true;
          cleanup();
        }
      };

      const onEnd = () => {
        cancelled = true;
        cleanup();
      };

      const cleanup = () => {
        window.clearTimeout(timerId);
        window.removeEventListener('pointermove', onMove);
        window.removeEventListener('pointerup', onEnd);
        window.removeEventListener('pointercancel', onEnd);
        window.removeEventListener('blur', onEnd);
        if (gridLongPressCleanupRef.current === cleanup) {
          gridLongPressCleanupRef.current = null;
        }
      };

      gridLongPressCleanupRef.current = cleanup;
      window.addEventListener('pointermove', onMove, { passive: true });
      window.addEventListener('pointerup', onEnd, { passive: true });
      window.addEventListener('pointercancel', onEnd, { passive: true });
      window.addEventListener('blur', onEnd);
    },
    []
  );

  // Ref mirrors + stable row callbacks so React.memo on SpreadsheetDataRow never invalidates on cell click
  const openRowContextMenuRef = useRef(openRowContextMenu);
  const openCellContextMenuRef = useRef(openCellContextMenu);
  const startGridLongPressRef = useRef(startGridLongPress);
  useEffect(() => {
    openRowContextMenuRef.current = openRowContextMenu;
    openCellContextMenuRef.current = openCellContextMenu;
    startGridLongPressRef.current = startGridLongPress;
  });

  const handleRowHeaderContextMenuCb = useCallback((r: number, x: number, y: number) => {
    openRowContextMenuRef.current(r, x, y);
  }, []);

  const handleRowHeaderTouchStartCb = useCallback(
    (r: number, e: React.PointerEvent<HTMLTableCellElement>) => {
      startGridLongPressRef.current(e, (x, y) => openRowContextMenuRef.current(r, x, y));
    },
    []
  );

  const handleCellContextMenuCb = useCallback((r: number, c: number, x: number, y: number) => {
    openCellContextMenuRef.current(r, c, x, y);
  }, []);

  const handleCellTouchStartCb = useCallback(
    (r: number, c: number, e: React.PointerEvent<HTMLTableCellElement>) => {
      startGridLongPressRef.current(e, (x, y) => openCellContextMenuRef.current(r, c, x, y));
    },
    []
  );

  const executeStructuralRowOp = useCallback(
    async (action: 'insert_row' | 'delete_rows', startRowIdx: number, endRowIdx: number = startRowIdx) => {
      const targetSheet = sheetsRef.current[activeSheetName];
      if (!targetSheet?.id || isStructuralBusy) return;

      setIsStructuralBusy(true);
      setOperationError(null);

      try {
        const flushed = await flushSaveQueue({ source: 'structural' });
        if (!flushed || hasAnyDirtyWork()) {
          throw new Error('Please save pending edits before modifying worksheet rows.');
        }

        const bodyPayload: Record<string, unknown> = {
          workspace,
          action,
          sheetId: targetSheet.id
        };
        if (action === 'insert_row') {
          bodyPayload.rowIndex = startRowIdx;
        } else {
          bodyPayload.startRowIndex = startRowIdx;
          bodyPayload.endRowIndex = endRowIdx;
        }

        const res = await fetch('/api/records', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(bodyPayload)
        });

        const data = await parseJsonSafe(res);
        if (!res.ok) {
          throw new Error(typeof data.error === 'string' && data.error ? data.error : 'Row operation failed.');
        }

        if (onReloadWorkspace) {
          await onReloadWorkspace();
        }
        setContextMenu(null);
      } catch (err: unknown) {
        const msg = err instanceof Error ? err.message : 'Row operation failed.';
        setOperationError(msg);
      } finally {
        setIsStructuralBusy(false);
      }
    },
    [activeSheetName, isStructuralBusy, flushSaveQueue, hasAnyDirtyWork, workspace, onReloadWorkspace]
  );

  const executeStructuralColOp = useCallback(
    async (action: 'insert_column' | 'delete_column', colIdx: number) => {
      const targetSheet = sheetsRef.current[activeSheetName];
      if (!targetSheet?.id || isStructuralBusy) return;

      setIsStructuralBusy(true);
      setOperationError(null);

      try {
        const flushed = await flushSaveQueue({ source: 'structural' });
        if (!flushed || hasAnyDirtyWork()) {
          throw new Error('Please save pending edits before modifying worksheet columns.');
        }

        const res = await fetch('/api/records', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            workspace,
            action,
            sheetId: targetSheet.id,
            colIndex: colIdx
          })
        });

        const data = await parseJsonSafe(res);
        if (!res.ok) {
          throw new Error(typeof data.error === 'string' && data.error ? data.error : 'Column operation failed.');
        }

        if (onReloadWorkspace) {
          await onReloadWorkspace();
        }
        setContextMenu(null);
      } catch (err: unknown) {
        const msg = err instanceof Error ? err.message : 'Column operation failed.';
        setOperationError(msg);
      } finally {
        setIsStructuralBusy(false);
      }
    },
    [activeSheetName, isStructuralBusy, flushSaveQueue, hasAnyDirtyWork, workspace, onReloadWorkspace]
  );

  const handleContextMenuAction = useCallback(
    async (action: ContextMenuActionId) => {
      const target = contextMenu;
      if (!target) return;

      if (action === 'cell_select') {
        if (target.kind === 'cell') {
          setSelection({
            mode: 'cell',
            anchorRow: target.row,
            anchorCol: target.col,
            focusRow: target.row,
            focusCol: target.col
          });
        }
        setContextMenu(null);
      } else if (action === 'cell_copy' || action === 'row_copy' || action === 'col_copy') {
        const tsv = buildSelectionTsv();
        const kind = target.kind === 'row' ? 'row' : target.kind === 'col' ? 'col' : 'cells';
        await writeToClipboardSafe(tsv, kind);
        setContextMenu(null);
      } else if (action === 'cell_cut' || action === 'row_cut' || action === 'col_cut') {
        const tsv = buildSelectionTsv();
        const kind = target.kind === 'row' ? 'row' : target.kind === 'col' ? 'col' : 'cells';
        await writeToClipboardSafe(tsv, kind);
        handleDeleteSelection();
        setContextMenu(null);
      } else if (action === 'cell_paste' || action === 'row_paste' || action === 'col_paste') {
        const text = await readFromClipboardSafe();
        setContextMenu(null);
        applyPasteTextToTarget(text, selection, { confirmOverwriteIfPopulated: true });
      } else if (action === 'cell_delete') {
        setContextMenu(null);
        handleDeleteSelection();
      } else if (action === 'row_select') {
        if (target.kind === 'row') {
          setSelection({
            mode: 'row',
            anchorRow: target.startRow,
            anchorCol: 0,
            focusRow: target.endRow,
            focusCol: totalCols - 1
          });
        }
        setContextMenu(null);
      } else if (action === 'row_insert_above') {
        if (target.kind === 'row') {
          if (target.isHeaderRow || target.startRow === 0) {
            setOperationError('Cannot insert a row above header Row 1.');
            setContextMenu(null);
            return;
          }
          await executeStructuralRowOp('insert_row', target.startRow - 1);
        }
      } else if (action === 'row_insert_below') {
        if (target.kind === 'row') {
          await executeStructuralRowOp('insert_row', target.endRow);
        }
      } else if (action === 'row_delete') {
        if (target.kind === 'row') {
          if (target.isHeaderRow || target.startRow === 0) {
            setOperationError('Header Row 1 cannot be deleted.');
            setContextMenu(null);
            return;
          }
          let populatedCount = 0;
          for (let r = target.startRow; r <= target.endRow; r++) {
            for (let c = 0; c < totalCols; c++) {
              if (isNonEmptyCell(getCellValue(r, c))) populatedCount++;
            }
          }
          if (populatedCount > 0) {
            setPendingConfirmAction({
              title: target.endRow > target.startRow ? 'Delete Populated Rows?' : 'Delete Populated Row?',
              message: `You are about to delete ${target.endRow - target.startRow + 1} row${
                target.endRow === target.startRow ? '' : 's'
              } containing ${populatedCount} populated cell${populatedCount === 1 ? '' : 's'}. This cannot be undone.`,
              confirmLabel: 'Delete Rows',
              confirmVariant: 'danger',
              run: () => {
                void executeStructuralRowOp('delete_rows', target.startRow - 1, target.endRow - 1);
              }
            });
            setContextMenu(null);
            return;
          }
          await executeStructuralRowOp('delete_rows', target.startRow - 1, target.endRow - 1);
        }
      } else if (action === 'col_select') {
        if (target.kind === 'col') {
          setSelection({
            mode: 'col',
            anchorRow: 0,
            anchorCol: target.col,
            focusRow: displayRowsCount - 1,
            focusCol: target.endCol
          });
        }
        setContextMenu(null);
      } else if (action === 'col_insert_left') {
        if (target.kind === 'col') {
          await executeStructuralColOp('insert_column', target.col);
        }
      } else if (action === 'col_insert_right') {
        if (target.kind === 'col') {
          await executeStructuralColOp('insert_column', target.endCol + 1);
        }
      } else if (action === 'col_resize') {
        if (target.kind === 'col') {
          const dimKey = currentSheet.id || activeSheetName;
          const curWidth = colWidthsBySheet[dimKey]?.[target.col] ?? colMetaList[target.col]?.defaultWidth ?? DEFAULT_COL_WIDTH;
          const defWidth = colMetaList[target.col]?.defaultWidth ?? DEFAULT_COL_WIDTH;
          setResizeColDialog({
            colIndex: target.col,
            currentWidth: curWidth,
            defaultWidth: defWidth
          });
          setContextMenu(null);
        }
      } else if (action === 'col_delete') {
        if (target.kind === 'col') {
          if (currentSheet.headers.length <= 1) {
            setOperationError('Cannot delete the only remaining column.');
            setContextMenu(null);
            return;
          }
          let populatedCount = 0;
          for (let r = 0; r < displayRowsCount; r++) {
            if (isNonEmptyCell(getCellValue(r, target.col))) populatedCount++;
          }
          if (populatedCount > 0) {
            setPendingConfirmAction({
              title: `Delete Column ${getColumnLetter(target.col)}?`,
              message: `Column ${getColumnLetter(target.col)} contains ${populatedCount} populated cell${
                populatedCount === 1 ? '' : 's'
              }. Deleting this column will shift subsequent columns left.`,
              confirmLabel: 'Delete Column',
              confirmVariant: 'danger',
              run: () => {
                void executeStructuralColOp('delete_column', target.col);
              }
            });
            setContextMenu(null);
            return;
          }
          await executeStructuralColOp('delete_column', target.col);
        }
      } else if (action === 'tab_rename') {
        if (target.kind === 'tab') {
          handleStartRename(target.sheetName);
        }
      } else if (action === 'tab_move_left') {
        if (target.kind === 'tab') {
          handleMoveSheetRelative(target.sheetName, -1);
        }
      } else if (action === 'tab_move_right') {
        if (target.kind === 'tab') {
          handleMoveSheetRelative(target.sheetName, 1);
        }
      } else if (action === 'tab_delete') {
        if (target.kind === 'tab') {
          void handleDeleteSheetImmediate(target.sheetName);
        }
      }
    },
    [
      contextMenu,
      buildSelectionTsv,
      writeToClipboardSafe,
      handleDeleteSelection,
      readFromClipboardSafe,
      applyPasteTextToTarget,
      selection,
      totalCols,
      displayRowsCount,
      executeStructuralRowOp,
      getCellValue,
      executeStructuralColOp,
      currentSheet.id,
      currentSheet.headers.length,
      activeSheetName,
      colWidthsBySheet,
      colMetaList,
      handleStartRename,
      handleMoveSheetRelative,
      handleDeleteSheetImmediate
    ]
  );

  // Column resize via Pointer Events + rAF DOM updates (commits React state once on pointerup)
  const handleColResizeStart = useCallback(
    (colIndex: number, e: React.PointerEvent<HTMLDivElement>) => {
      if (e.pointerType === 'mouse' && e.button !== 0) return;
      e.stopPropagation();
      e.preventDefault();

      if (resizeCleanupRef.current) {
        resizeCleanupRef.current();
      }

      const targetSheetKey = activeDimKeyRef.current;
      const targetCol = colIndex;
      const startX = e.clientX;
      const startWidth = getColWidthRef.current(colIndex);
      const startTableWidth = totalTableWidthRef.current;
      let latestWidth = startWidth;
      let rafId: number | null = null;
      let pendingClientX = startX;

      const applyDomWidth = (nextWidth: number) => {
        const colEl = colElementsRef.current[targetCol];
        if (colEl) {
          colEl.style.width = `${nextWidth}px`;
        }
        const letterTh = headerLetterThRefs.current[targetCol];
        if (letterTh) {
          letterTh.style.width = `${nextWidth}px`;
          letterTh.style.minWidth = `${nextWidth}px`;
        }
        const titleTd = headerTitleTdRefs.current[targetCol];
        if (titleTd) {
          titleTd.style.width = `${nextWidth}px`;
          titleTd.style.minWidth = `${nextWidth}px`;
        }
        if (tableRef.current) {
          tableRef.current.style.width = `${startTableWidth - startWidth + nextWidth}px`;
        }
      };

      const handlePointerMove = (moveEvent: PointerEvent) => {
        pendingClientX = moveEvent.clientX;
        if (rafId === null) {
          rafId = window.requestAnimationFrame(() => {
            rafId = null;
            const delta = pendingClientX - startX;
            latestWidth = Math.min(
              MAX_COL_WIDTH,
              Math.max(MIN_COL_WIDTH, Math.round(startWidth + delta))
            );
            applyDomWidth(latestWidth);
          });
        }
      };

      const finishResize = (commitState: boolean) => {
        if (rafId !== null) {
          window.cancelAnimationFrame(rafId);
          rafId = null;
          const delta = pendingClientX - startX;
          latestWidth = Math.min(
            MAX_COL_WIDTH,
            Math.max(MIN_COL_WIDTH, Math.round(startWidth + delta))
          );
          applyDomWidth(latestWidth);
        }

        justResizedRef.current = true;
        window.setTimeout(() => {
          justResizedRef.current = false;
        }, 80);

        window.removeEventListener('pointermove', handlePointerMove);
        window.removeEventListener('pointerup', onPointerUp);
        window.removeEventListener('pointercancel', onPointerCancel);
        window.removeEventListener('blur', onWindowBlur);
        resizeCleanupRef.current = null;

        if (commitState && latestWidth !== startWidth) {
          setColWidthsBySheet((prev) => ({
            ...prev,
            [targetSheetKey]: {
              ...(prev[targetSheetKey] || {}),
              [targetCol]: latestWidth
            }
          }));
        }
      };

      const onPointerUp = () => finishResize(true);
      const onPointerCancel = () => finishResize(true);
      const onWindowBlur = () => finishResize(true);

      resizeCleanupRef.current = () => finishResize(false);
      window.addEventListener('pointermove', handlePointerMove);
      window.addEventListener('pointerup', onPointerUp);
      window.addEventListener('pointercancel', onPointerCancel);
      window.addEventListener('blur', onWindowBlur);
    },
    []
  );

  // Row height resize via Pointer Events + rAF DOM updates (commits React state once on pointerup)
  const handleRowResizeStart = useCallback(
    (gridRowIdx: number, e: React.PointerEvent<HTMLDivElement>) => {
      if (e.pointerType === 'mouse' && e.button !== 0) return;
      e.stopPropagation();
      e.preventDefault();

      if (resizeCleanupRef.current) {
        resizeCleanupRef.current();
      }

      const targetSheetKey = activeDimKeyRef.current;
      const targetRow = gridRowIdx;
      const startY = e.clientY;
      const thEl = e.currentTarget.parentElement;
      const measuredHeight = thEl ? thEl.getBoundingClientRect().height : 0;
      const fallbackHeight =
        targetRow === 0 ? DEFAULT_HEADER_ROW_HEIGHT : DEFAULT_DATA_ROW_HEIGHT;
      const startHeight =
        getManualRowHeightRef.current(gridRowIdx) ??
        (measuredHeight > 0 ? measuredHeight : fallbackHeight);
      let latestHeight = startHeight;
      let rafId: number | null = null;
      let pendingClientY = startY;

      const applyDomHeight = (nextHeight: number) => {
        if (targetRow === 0) {
          if (headerRowThRef.current) {
            headerRowThRef.current.style.height = `${nextHeight}px`;
          }
          const trEl = rowElementsRef.current[0];
          if (trEl) {
            trEl.style.height = `${nextHeight}px`;
          }
        } else {
          const trEl = rowElementsRef.current[targetRow];
          if (trEl) {
            trEl.style.height = `${nextHeight}px`;
          }
        }
      };

      const handlePointerMove = (moveEvent: PointerEvent) => {
        pendingClientY = moveEvent.clientY;
        if (rafId === null) {
          rafId = window.requestAnimationFrame(() => {
            rafId = null;
            const delta = pendingClientY - startY;
            latestHeight = Math.min(
              MAX_ROW_HEIGHT,
              Math.max(MIN_ROW_HEIGHT, Math.round(startHeight + delta))
            );
            applyDomHeight(latestHeight);
          });
        }
      };

      const finishResize = (commitState: boolean) => {
        if (rafId !== null) {
          window.cancelAnimationFrame(rafId);
          rafId = null;
          const delta = pendingClientY - startY;
          latestHeight = Math.min(
            MAX_ROW_HEIGHT,
            Math.max(MIN_ROW_HEIGHT, Math.round(startHeight + delta))
          );
          applyDomHeight(latestHeight);
        }

        justResizedRef.current = true;
        window.setTimeout(() => {
          justResizedRef.current = false;
        }, 80);

        window.removeEventListener('pointermove', handlePointerMove);
        window.removeEventListener('pointerup', onPointerUp);
        window.removeEventListener('pointercancel', onPointerCancel);
        window.removeEventListener('blur', onWindowBlur);
        resizeCleanupRef.current = null;

        if (commitState && latestHeight !== startHeight) {
          setRowHeightsBySheet((prev) => ({
            ...prev,
            [targetSheetKey]: {
              ...(prev[targetSheetKey] || {}),
              [targetRow]: latestHeight
            }
          }));
        }
      };

      const onPointerUp = () => finishResize(true);
      const onPointerCancel = () => finishResize(true);
      const onWindowBlur = () => finishResize(true);

      resizeCleanupRef.current = () => finishResize(false);
      window.addEventListener('pointermove', handlePointerMove);
      window.addEventListener('pointerup', onPointerUp);
      window.addEventListener('pointercancel', onPointerCancel);
      window.addEventListener('blur', onWindowBlur);
    },
    []
  );

  // Excel-Style Drag-and-Drop Row Rearrangement (cannot drop above Header Row 1; single atomic RPC)
  const executeReorderRows = useCallback(
    async (startDataIdx: number, endDataIdx: number, targetDataIdx: number) => {
      if (targetDataIdx >= startDataIdx && targetDataIdx <= endDataIdx + 1) return;
      const targetSheet = sheetsRef.current[activeSheetName];
      if (!targetSheet?.id || isStructuralBusyRef.current) return;

      // Do not submit empty trailing-space drags if no existing data rows fall in range
      if (startDataIdx >= targetSheet.rows.length) return;

      const sheetId = targetSheet.id;
      const dimKey = sheetId;
      setIsStructuralBusy(true);
      setOperationError(null);

      // Flush any unsaved cell/layout edits first so backend reorder operates on synced state
      try {
        if (isEditingRef.current) commitEditRef.current(0, 0);
        if (hasAnyDirtyWork() || isSavingRef.current) {
          const flushed = await flushSaveQueue({ source: 'structural' });
          if (!flushed || hasAnyDirtyWork()) {
            throw new Error('Please save pending edits before reordering rows.');
          }
        }
      } catch (err: unknown) {
        const msg = err instanceof Error ? err.message : 'Row rearrangement failed.';
        setOperationError(msg);
        setIsStructuralBusy(false);
        return;
      }

      const prevSheet = sheetsRef.current[activeSheetName];
      if (!prevSheet?.id) {
        setIsStructuralBusy(false);
        return;
      }

      const prevRowHeightsMap = { ...(rowHeightsBySheetRef.current[dimKey] || {}) };
      const prevSavedSnap = savedLayoutBySheetRef.current[dimKey]
        ? {
            cols: { ...savedLayoutBySheetRef.current[dimKey].cols },
            rows: { ...savedLayoutBySheetRef.current[dimKey].rows }
          }
        : { cols: {}, rows: {} };
      const prevSelection = selection;

      // Apply exact optimistic permutation matching server SQL & mockStore logic
      const maxAffected = Math.max(endDataIdx, targetDataIdx - 1);
      const minAffected = Math.min(startDataIdx, targetDataIdx);
      const nextRows = [...prevSheet.rows];
      const requiredLen = Math.min(100000, Math.max(nextRows.length, maxAffected + 1));
      while (nextRows.length < requiredLen) nextRows.push([]);

      const movedCount = endDataIdx - startDataIdx + 1;
      const movedSlice = nextRows.splice(startDataIdx, movedCount);
      const insertAt = targetDataIdx < startDataIdx ? targetDataIdx : targetDataIdx - movedCount;
      nextRows.splice(insertAt, 0, ...movedSlice);

      while (nextRows.length > 0) {
        const last = nextRows[nextRows.length - 1];
        if (Array.isArray(last) && last.some(isNonEmptyCell)) break;
        if (nextRows.length - 1 < minAffected) break;
        nextRows.pop();
      }

      const nextRowVersions: Record<number, number> = {};
      for (const [k, v] of Object.entries(prevSheet.rowVersions || {})) {
        nextRowVersions[remapMovedIndex(Number(k), startDataIdx, endDataIdx, targetDataIdx)] = v;
      }

      const nextCellVersions: Record<number, Record<number, number>> = {};
      for (const [k, v] of Object.entries(prevSheet.cellVersions || {})) {
        nextCellVersions[remapMovedIndex(Number(k), startDataIdx, endDataIdx, targetDataIdx)] = v;
      }

      const nextRowUpdatedAt: Record<number, string> = {};
      for (const [k, v] of Object.entries(prevSheet.rowUpdatedAt || {})) {
        nextRowUpdatedAt[remapMovedIndex(Number(k), startDataIdx, endDataIdx, targetDataIdx)] = v;
      }

      const nextRowHeights: Record<number, number> = {};
      for (const [k, v] of Object.entries(prevRowHeightsMap)) {
        const gridIdx = Number(k);
        if (gridIdx === 0) {
          nextRowHeights[0] = v;
        } else if (gridIdx >= 1) {
          const nextD = remapMovedIndex(gridIdx - 1, startDataIdx, endDataIdx, targetDataIdx);
          if (nextD >= 0 && nextD + 1 < 100000) {
            nextRowHeights[nextD + 1] = v;
          }
        }
      }

      const newStartGrid = insertAt + 1;
      const newEndGrid = insertAt + movedCount;

      undoStackBySheetRef.current[sheetId] = [];
      redoStackBySheetRef.current[sheetId] = [];

      setSheets((prev) => ({
        ...prev,
        [activeSheetName]: {
          ...prevSheet,
          rows: nextRows,
          rowVersions: nextRowVersions,
          cellVersions: nextCellVersions,
          rowUpdatedAt: nextRowUpdatedAt,
          rowHeights: { ...nextRowHeights }
        }
      }));
      setRowHeightsBySheet((prev) => ({ ...prev, [dimKey]: nextRowHeights }));
      setSavedLayoutBySheet((prev) => ({
        ...prev,
        [dimKey]: {
          cols: { ...(prev[dimKey]?.cols || {}) },
          rows: { ...nextRowHeights }
        }
      }));
      setSelection({
        mode: 'row',
        anchorRow: newStartGrid,
        anchorCol: 0,
        focusRow: newEndGrid,
        focusCol: totalColsRef.current - 1
      });

      try {
        const res = await fetch('/api/records', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            workspace,
            action: 'reorder_rows',
            sheetId,
            startRowIndex: startDataIdx,
            endRowIndex: endDataIdx,
            targetRowIndex: targetDataIdx
          })
        });
        const data = await parseJsonSafe(res);
        if (!res.ok) {
          throw new Error(
            typeof data.error === 'string' && data.error
              ? data.error
              : 'Failed to persist row order.'
          );
        }

        const syncedHeights =
          data.rowHeights && typeof data.rowHeights === 'object'
            ? (data.rowHeights as Record<number, number>)
            : nextRowHeights;

        setRowHeightsBySheet((prev) => ({ ...prev, [dimKey]: { ...syncedHeights } }));
        setSavedLayoutBySheet((prev) => ({
          ...prev,
          [dimKey]: {
            cols: { ...(prev[dimKey]?.cols || {}) },
            rows: { ...syncedHeights }
          }
        }));
        setSheets((prev) => {
          const cur = prev[activeSheetName];
          if (!cur) return prev;
          return {
            ...prev,
            [activeSheetName]: { ...cur, rowHeights: { ...syncedHeights } }
          };
        });
        setSaveStatus('saved');
        setSaveMessage('Saved');
      } catch (err: unknown) {
        // Revert cleanly on persistence failure
        setSheets((prev) => ({ ...prev, [activeSheetName]: prevSheet }));
        setRowHeightsBySheet((prev) => ({ ...prev, [dimKey]: prevRowHeightsMap }));
        setSavedLayoutBySheet((prev) => ({ ...prev, [dimKey]: prevSavedSnap }));
        setSelection(prevSelection);
        const msg = err instanceof Error ? err.message : 'Failed to persist row rearrangement.';
        setOperationError(msg);
      } finally {
        setIsStructuralBusy(false);
      }
    },
    [activeSheetName, flushSaveQueue, hasAnyDirtyWork, selection, workspace]
  );

  // Excel-Style Drag-and-Drop Column Rearrangement (single atomic RPC, moves header + cells + widths)
  const executeReorderColumns = useCallback(
    async (startCol: number, endCol: number, targetCol: number) => {
      if (targetCol >= startCol && targetCol <= endCol + 1) return;
      const targetSheet = sheetsRef.current[activeSheetName];
      if (!targetSheet?.id || isStructuralBusyRef.current) return;

      const sheetId = targetSheet.id;
      const dimKey = sheetId;
      setIsStructuralBusy(true);
      setOperationError(null);

      try {
        if (isEditingRef.current) commitEditRef.current(0, 0);
        if (hasAnyDirtyWork() || isSavingRef.current) {
          const flushed = await flushSaveQueue({ source: 'structural' });
          if (!flushed || hasAnyDirtyWork()) {
            throw new Error('Please save pending edits before reordering columns.');
          }
        }
      } catch (err: unknown) {
        const msg = err instanceof Error ? err.message : 'Column rearrangement failed.';
        setOperationError(msg);
        setIsStructuralBusy(false);
        return;
      }

      const prevSheet = sheetsRef.current[activeSheetName];
      if (!prevSheet?.id) {
        setIsStructuralBusy(false);
        return;
      }

      const prevColWidthsMap = { ...(colWidthsBySheetRef.current[dimKey] || {}) };
      const prevSavedSnap = savedLayoutBySheetRef.current[dimKey]
        ? {
            cols: { ...savedLayoutBySheetRef.current[dimKey].cols },
            rows: { ...savedLayoutBySheetRef.current[dimKey].rows }
          }
        : { cols: {}, rows: {} };
      const prevSelection = selection;

      const movedCount = endCol - startCol + 1;
      const insertAt = targetCol < startCol ? targetCol : targetCol - movedCount;

      const nextHeaders = [...prevSheet.headers];
      const reqHeaderLen = Math.min(256, Math.max(nextHeaders.length, endCol + 1, targetCol));
      while (nextHeaders.length < reqHeaderLen) nextHeaders.push('');
      const movedHeaders = nextHeaders.splice(startCol, movedCount);
      nextHeaders.splice(insertAt, 0, ...movedHeaders);

      const minAffected = Math.min(startCol, targetCol);
      const nextRows = prevSheet.rows.map((row) => {
        if (!Array.isArray(row) || row.length <= minAffected) return row;
        const copy = [...row];
        const reqLen = Math.min(256, Math.max(copy.length, endCol + 1, targetCol));
        while (copy.length < reqLen) copy.push('');
        const movedCells = copy.splice(startCol, movedCount);
        copy.splice(insertAt, 0, ...movedCells);
        return copy;
      });

      const nextRowVersions: Record<number, number> = {};
      for (let r = 0; r < prevSheet.rows.length; r++) {
        nextRowVersions[r] = (prevSheet.rowVersions?.[r] ?? 0) + 1;
      }

      const nextCellVersions: Record<number, Record<number, number>> = {};
      for (const [rKey, cvMap] of Object.entries(prevSheet.cellVersions || {})) {
        const nextRowCv: Record<number, number> = {};
        for (const [cKey, ver] of Object.entries(cvMap)) {
          const nextC = remapMovedIndex(Number(cKey), startCol, endCol, targetCol);
          if (nextC >= 0 && nextC < MAX_COLS_LIMIT) nextRowCv[nextC] = ver;
        }
        nextCellVersions[Number(rKey)] = nextRowCv;
      }

      const nextColWidths: Record<number, number> = {};
      for (const [k, v] of Object.entries(prevColWidthsMap)) {
        const nextC = remapMovedIndex(Number(k), startCol, endCol, targetCol);
        if (nextC >= 0 && nextC < MAX_COLS_LIMIT) nextColWidths[nextC] = v;
      }

      undoStackBySheetRef.current[sheetId] = [];
      redoStackBySheetRef.current[sheetId] = [];

      const nextHeaderVersion = (prevSheet.headerVersion ?? 1) + 1;

      setSheets((prev) => ({
        ...prev,
        [activeSheetName]: {
          ...prevSheet,
          headers: nextHeaders,
          headerVersion: nextHeaderVersion,
          rows: nextRows,
          rowVersions: nextRowVersions,
          cellVersions: nextCellVersions,
          columnWidths: { ...nextColWidths }
        }
      }));
      setColWidthsBySheet((prev) => ({ ...prev, [dimKey]: nextColWidths }));
      setSavedLayoutBySheet((prev) => ({
        ...prev,
        [dimKey]: {
          cols: { ...nextColWidths },
          rows: { ...(prev[dimKey]?.rows || {}) }
        }
      }));
      setSelection({
        mode: 'col',
        anchorRow: 0,
        anchorCol: insertAt,
        focusRow: displayRowsCountRef.current - 1,
        focusCol: insertAt + movedCount - 1
      });

      try {
        const res = await fetch('/api/records', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            workspace,
            action: 'reorder_columns',
            sheetId,
            startColIndex: startCol,
            endColIndex: endCol,
            targetColIndex: targetCol
          })
        });
        const data = await parseJsonSafe(res);
        if (!res.ok) {
          throw new Error(
            typeof data.error === 'string' && data.error
              ? data.error
              : 'Failed to persist column order.'
          );
        }

        const syncedHeaders = Array.isArray(data.headers) ? (data.headers as string[]) : nextHeaders;
        const syncedHeaderVer = Number(data.headerVersion ?? nextHeaderVersion);
        const syncedWidths =
          data.columnWidths && typeof data.columnWidths === 'object'
            ? (data.columnWidths as Record<number, number>)
            : nextColWidths;

        setColWidthsBySheet((prev) => ({ ...prev, [dimKey]: { ...syncedWidths } }));
        setSavedLayoutBySheet((prev) => ({
          ...prev,
          [dimKey]: {
            cols: { ...syncedWidths },
            rows: { ...(prev[dimKey]?.rows || {}) }
          }
        }));
        setSheets((prev) => {
          const cur = prev[activeSheetName];
          if (!cur) return prev;
          return {
            ...prev,
            [activeSheetName]: {
              ...cur,
              headers: syncedHeaders,
              headerVersion: syncedHeaderVer,
              columnWidths: { ...syncedWidths }
            }
          };
        });
        setSaveStatus('saved');
        setSaveMessage('Saved');
      } catch (err: unknown) {
        setSheets((prev) => ({ ...prev, [activeSheetName]: prevSheet }));
        setColWidthsBySheet((prev) => ({ ...prev, [dimKey]: prevColWidthsMap }));
        setSavedLayoutBySheet((prev) => ({ ...prev, [dimKey]: prevSavedSnap }));
        setSelection(prevSelection);
        const msg = err instanceof Error ? err.message : 'Failed to persist column rearrangement.';
        setOperationError(msg);
      } finally {
        setIsStructuralBusy(false);
      }
    },
    [activeSheetName, flushSaveQueue, hasAnyDirtyWork, selection, workspace]
  );

  const executeReorderRowsRef = useRef(executeReorderRows);
  const executeReorderColumnsRef = useRef(executeReorderColumns);
  useEffect(() => {
    executeReorderRowsRef.current = executeReorderRows;
    executeReorderColumnsRef.current = executeReorderColumns;
  });

  const handleRowDragStart = useCallback(
    (gridRowIdx: number, e: React.PointerEvent<HTMLTableCellElement>) => {
      if (gridRowIdx <= 0) return; // Header Row 1 cannot be dragged
      if (isStructuralBusyRef.current) return;
      if (e.pointerType === 'mouse' && e.button !== 0) return;

      // On touch/pen, only initiate drag when already selecting full row(s) so normal scroll/long-press stay natural
      const bounds = selectionBoundsRef.current;
      const isRowAlreadySelected =
        selectionModeRef.current === 'row' &&
        gridRowIdx >= bounds.rMin &&
        gridRowIdx <= bounds.rMax;
      if (e.pointerType !== 'mouse' && !isRowAlreadySelected) return;

      if (reorderDragCleanupRef.current) {
        reorderDragCleanupRef.current();
      }

      const startGridRow = isRowAlreadySelected ? Math.max(1, bounds.rMin) : gridRowIdx;
      const endGridRow = isRowAlreadySelected ? Math.max(startGridRow, bounds.rMax) : gridRowIdx;

      const startY = e.clientY;
      let isDragging = false;
      let currentTargetGapGrid = startGridRow;
      let rafId: number | null = null;
      let pendingClientY = startY;

      const updateDropLine = (clientY: number) => {
        const container = gridContainerRef.current;
        const wrapper = tableWrapperRef.current;
        const lineEl = rowDropLineRef.current;
        if (!container || !wrapper || !lineEl) return;

        const containerRect = container.getBoundingClientRect();
        const theadHeight = theadRef.current?.offsetHeight ?? 58;
        if (clientY < containerRect.top + theadHeight + 24) {
          container.scrollTop = Math.max(0, container.scrollTop - 14);
        } else if (clientY > containerRect.bottom - 28) {
          container.scrollTop += 14;
        }

        const totalRows = displayRowsCountRef.current;
        let gapGrid = totalRows;
        for (let g = 1; g < totalRows; g++) {
          const tr = rowElementsRef.current[g];
          if (!tr) continue;
          const trRect = tr.getBoundingClientRect();
          if (clientY < trRect.top + trRect.height / 2) {
            gapGrid = g;
            break;
          }
        }
        gapGrid = Math.max(1, gapGrid); // Never drop above Header Row 1
        currentTargetGapGrid = gapGrid;

        const wrapperRect = wrapper.getBoundingClientRect();
        let lineViewportY: number;
        if (gapGrid < totalRows && rowElementsRef.current[gapGrid]) {
          lineViewportY = rowElementsRef.current[gapGrid]!.getBoundingClientRect().top - 1;
        } else {
          const lastTr = rowElementsRef.current[totalRows - 1];
          lineViewportY = lastTr ? lastTr.getBoundingClientRect().bottom + 1 : clientY;
        }

        const topOffset = Math.max(theadHeight, lineViewportY - wrapperRect.top);
        lineEl.style.display = 'block';
        lineEl.style.top = `${topOffset}px`;
      };

      const onPointerMove = (moveEvent: PointerEvent) => {
        pendingClientY = moveEvent.clientY;
        if (!isDragging) {
          if (Math.abs(pendingClientY - startY) < 6) return;
          isDragging = true;
          if (gridLongPressCleanupRef.current) {
            gridLongPressCleanupRef.current();
          }
          setContextMenu(null);
        }
        if (rafId === null) {
          rafId = window.requestAnimationFrame(() => {
            rafId = null;
            updateDropLine(pendingClientY);
          });
        }
      };

      const finishDrag = (commit: boolean) => {
        if (rafId !== null) {
          window.cancelAnimationFrame(rafId);
          rafId = null;
        }
        if (rowDropLineRef.current) {
          rowDropLineRef.current.style.display = 'none';
        }
        window.removeEventListener('pointermove', onPointerMove);
        window.removeEventListener('pointerup', onPointerUp);
        window.removeEventListener('pointercancel', onPointerCancel);
        window.removeEventListener('blur', onBlur);
        reorderDragCleanupRef.current = null;

        if (isDragging) {
          justResizedRef.current = true;
          window.setTimeout(() => {
            justResizedRef.current = false;
          }, 80);

          if (commit) {
            const startDataIdx = startGridRow - 1;
            const endDataIdx = endGridRow - 1;
            const targetDataIdx = currentTargetGapGrid - 1;
            if (targetDataIdx < startDataIdx || targetDataIdx > endDataIdx + 1) {
              void executeReorderRowsRef.current(startDataIdx, endDataIdx, targetDataIdx);
            }
          }
        }
      };

      const onPointerUp = () => finishDrag(true);
      const onPointerCancel = () => finishDrag(false);
      const onBlur = () => finishDrag(false);

      reorderDragCleanupRef.current = () => finishDrag(false);
      window.addEventListener('pointermove', onPointerMove);
      window.addEventListener('pointerup', onPointerUp);
      window.addEventListener('pointercancel', onPointerCancel);
      window.addEventListener('blur', onBlur);
    },
    []
  );

  const handleColDragStart = useCallback(
    (colIdx: number, e: React.PointerEvent<HTMLTableCellElement>) => {
      if (isStructuralBusyRef.current) return;
      if (e.pointerType === 'mouse' && e.button !== 0) return;

      const bounds = selectionBoundsRef.current;
      const isColAlreadySelected =
        selectionModeRef.current === 'col' &&
        colIdx >= bounds.cMin &&
        colIdx <= bounds.cMax;
      if (e.pointerType !== 'mouse' && !isColAlreadySelected) return;

      if (reorderDragCleanupRef.current) {
        reorderDragCleanupRef.current();
      }

      const startCol = isColAlreadySelected ? bounds.cMin : colIdx;
      const endCol = isColAlreadySelected ? bounds.cMax : colIdx;

      const startX = e.clientX;
      let isDragging = false;
      let currentTargetGapCol = startCol;
      let rafId: number | null = null;
      let pendingClientX = startX;

      const updateDropLine = (clientX: number) => {
        const container = gridContainerRef.current;
        const wrapper = tableWrapperRef.current;
        const lineEl = colDropLineRef.current;
        if (!container || !wrapper || !lineEl) return;

        const containerRect = container.getBoundingClientRect();
        if (clientX < containerRect.left + rowHeaderWidthRef.current + 24) {
          container.scrollLeft = Math.max(0, container.scrollLeft - 18);
        } else if (clientX > containerRect.right - 28) {
          container.scrollLeft += 18;
        }

        const colsTotal = totalColsRef.current;
        let gapCol = colsTotal;
        for (let c = 0; c < colsTotal; c++) {
          const th = headerLetterThRefs.current[c];
          if (!th) continue;
          const thRect = th.getBoundingClientRect();
          if (clientX < thRect.left + thRect.width / 2) {
            gapCol = c;
            break;
          }
        }
        currentTargetGapCol = gapCol;

        const wrapperRect = wrapper.getBoundingClientRect();
        let lineViewportX: number;
        if (gapCol < colsTotal && headerLetterThRefs.current[gapCol]) {
          lineViewportX = headerLetterThRefs.current[gapCol]!.getBoundingClientRect().left - 1;
        } else {
          const lastTh = headerLetterThRefs.current[colsTotal - 1];
          lineViewportX = lastTh ? lastTh.getBoundingClientRect().right + 1 : clientX;
        }

        const leftOffset = Math.max(rowHeaderWidthRef.current, lineViewportX - wrapperRect.left);
        lineEl.style.display = 'block';
        lineEl.style.left = `${leftOffset}px`;
      };

      const onPointerMove = (moveEvent: PointerEvent) => {
        pendingClientX = moveEvent.clientX;
        if (!isDragging) {
          if (Math.abs(pendingClientX - startX) < 6) return;
          isDragging = true;
          if (gridLongPressCleanupRef.current) {
            gridLongPressCleanupRef.current();
          }
          setContextMenu(null);
        }
        if (rafId === null) {
          rafId = window.requestAnimationFrame(() => {
            rafId = null;
            updateDropLine(pendingClientX);
          });
        }
      };

      const finishDrag = (commit: boolean) => {
        if (rafId !== null) {
          window.cancelAnimationFrame(rafId);
          rafId = null;
        }
        if (colDropLineRef.current) {
          colDropLineRef.current.style.display = 'none';
        }
        window.removeEventListener('pointermove', onPointerMove);
        window.removeEventListener('pointerup', onPointerUp);
        window.removeEventListener('pointercancel', onPointerCancel);
        window.removeEventListener('blur', onBlur);
        reorderDragCleanupRef.current = null;

        if (isDragging) {
          justResizedRef.current = true;
          window.setTimeout(() => {
            justResizedRef.current = false;
          }, 80);

          if (commit && (currentTargetGapCol < startCol || currentTargetGapCol > endCol + 1)) {
            void executeReorderColumnsRef.current(startCol, endCol, currentTargetGapCol);
          }
        }
      };

      const onPointerUp = () => finishDrag(true);
      const onPointerCancel = () => finishDrag(false);
      const onBlur = () => finishDrag(false);

      reorderDragCleanupRef.current = () => finishDrag(false);
      window.addEventListener('pointermove', onPointerMove);
      window.addEventListener('pointerup', onPointerUp);
      window.addEventListener('pointercancel', onPointerCancel);
      window.addEventListener('blur', onBlur);
    },
    []
  );

  const renderedSheetKeys = useMemo(() => {
    if (!dragTabState) return sheetKeys;
    const withoutSource = sheetKeys.filter((k) => k !== dragTabState.sourceName);
    const clampedIndex = Math.max(0, Math.min(dragTabState.targetIndex, withoutSource.length));
    withoutSource.splice(clampedIndex, 0, dragTabState.sourceName);
    return withoutSource;
  }, [sheetKeys, dragTabState]);

  const visibleWorkspaces = allowedWorkspaces && allowedWorkspaces.length > 0 ? allowedWorkspaces : [workspace];
  const staffLabel = user?.name || user?.username || workspace;

  const isFullRowSelectionMode = selection.mode === 'row' || selection.mode === 'all';
  const isFullColSelectionMode = selection.mode === 'col' || selection.mode === 'all';
  const showAdminImportBtn = Boolean(user?.role === 'admin' && canAdminImport && onRequestInitialImport);
  const hasStatusFeedback = Boolean(
    isLoading ||
      saveStatus === 'saving' ||
      (saveStatus === 'saved' && !hasUnsavedChanges) ||
      hasUnsavedChanges ||
      saveStatus === 'error' ||
      hasConflict
  );
  const hasMobileSecondaryRow = visibleWorkspaces.length > 1 || showAdminImportBtn || hasStatusFeedback;

  return (
    <div
      ref={rootRef}
      className="flex flex-col mk-h-dvh w-full max-w-full bg-[var(--grid-canvas)] select-none overflow-hidden font-sans text-xs text-[var(--control-text)] outline-none"
      onKeyDown={handleKeyDown}
      onCopy={handleCopy}
      onCut={handleCut}
      onPaste={handlePaste}
      tabIndex={0}
    >
      {/* 1. Application Toolbar / Header (Responsive: single row on md+, clean compact adaptive bar below md) */}
      <div className="min-h-12 md:h-12 px-2.5 sm:px-4 py-1.5 md:py-0 bg-[var(--chrome-bg)] text-[var(--control-text)] border-b border-[var(--chrome-border)] flex flex-wrap md:flex-nowrap items-center justify-between gap-x-2 sm:gap-x-3 gap-y-1.5 shrink-0">
        {/* Top-Left: Official MK Eye Center Logo directly, NO background container */}
        <div className="flex items-center shrink-0" data-tour="brand-header">
          {/* eslint-disable-next-line @next/next/no-img-element */}
          <img
            src="/assets/aistudio/mk_long_logo_a5863dd4.png"
            alt="MK Eye Center Logo"
            className="h-6 sm:h-7 md:h-8 w-auto max-w-[104px] sm:max-w-[155px] md:max-w-none object-contain dark:invert dark:hue-rotate-180 dark:brightness-110 dark:contrast-110"
            referrerPolicy="no-referrer"
          />
        </div>

        {/* Workspace Switcher + Save Status (inline on md+, neat compact secondary row below md when present) */}
        {hasMobileSecondaryRow && (
          <div className="order-last w-full md:order-none md:w-auto md:ml-auto flex flex-wrap items-center justify-between md:justify-end gap-2 border-t border-[var(--chrome-border)] pt-1.5 md:border-t-0 md:pt-0">
            {/* Admin workspace switcher if applicable */}
            {visibleWorkspaces.length > 1 && (
              <div
                data-tour="workspace-switcher"
                className="flex items-center bg-[var(--control-bg)] p-0.5 rounded-lg border border-[var(--control-border)] text-xs shrink-0"
              >
                <button
                  type="button"
                  onClick={() => onWorkspaceSwitchRequest('reception')}
                  className={`min-h-[30px] md:min-h-0 px-2.5 sm:px-3 py-1 font-semibold text-[11px] sm:text-xs rounded-md cursor-pointer touch-manipulation ${
                    activeWorkspace === 'reception'
                      ? 'bg-blue-600 text-white shadow-xs'
                      : 'text-[var(--control-muted)] hover:text-[var(--control-text)]'
                  }`}
                >
                  Reception
                </button>
                <button
                  type="button"
                  onClick={() => onWorkspaceSwitchRequest('doctor')}
                  className={`min-h-[30px] md:min-h-0 px-2.5 sm:px-3 py-1 font-semibold text-[11px] sm:text-xs rounded-md cursor-pointer touch-manipulation ${
                    activeWorkspace === 'doctor'
                      ? 'bg-blue-600 text-white shadow-xs'
                      : 'text-[var(--control-muted)] hover:text-[var(--control-text)]'
                  }`}
                >
                  Doctor
                </button>
              </div>
            )}

            {/* Save & Conflict Status Notice */}
            {hasStatusFeedback && (
              <div
                className="flex items-center gap-2 text-[11px] ml-auto md:ml-0 md:mr-1 min-w-0"
                role="status"
                aria-live="polite"
              >
                {isLoading && <span className="text-[var(--control-muted)]">Loading...</span>}
                {!isLoading && saveStatus === 'saving' && (
                  <span className="text-[var(--control-text)] font-medium">Saving...</span>
                )}
                {!isLoading && saveStatus === 'saved' && !hasUnsavedChanges && (
                  <span className="text-emerald-600 dark:text-emerald-400 font-semibold whitespace-nowrap">
                    {saveMessage || 'Saved'}
                  </span>
                )}
                {!isLoading && hasUnsavedChanges && saveStatus !== 'saving' && (
                  <span className="text-amber-600 dark:text-amber-400 font-medium whitespace-nowrap">
                    Unsaved
                  </span>
                )}
                {!isLoading && saveStatus === 'error' && (
                  <span
                    title={saveError || 'Save failed'}
                    className="text-red-600 dark:text-red-400 font-medium max-w-[160px] sm:max-w-[200px] truncate"
                  >
                    {saveError || 'Save failed'}
                  </span>
                )}
                {!isLoading && hasConflict && onReloadWorkspace && (
                  <button
                    type="button"
                    onClick={() => void handleReloadLatest()}
                    disabled={isReloading}
                    className="min-h-[28px] md:min-h-0 px-2.5 py-0.5 bg-red-600 text-white rounded-md font-medium hover:bg-red-700 cursor-pointer disabled:opacity-50 text-[11px] shrink-0 touch-manipulation"
                  >
                    {isReloading ? 'Reloading...' : 'Reload'}
                  </button>
                )}
              </div>
            )}

            {showAdminImportBtn && (
              <button
                type="button"
                onClick={onRequestInitialImport}
                className="min-h-[30px] md:min-h-0 px-2.5 sm:px-3 py-1.5 font-semibold text-white bg-emerald-700 hover:bg-emerald-600 rounded-lg cursor-pointer text-[11px] sm:text-xs whitespace-nowrap touch-manipulation"
              >
                Import Workbooks
              </button>
            )}
          </div>
        )}

        {/* Primary Right Controls: Tour & Theme Icons, Save and Export */}
        <div className={`flex items-center gap-1.5 sm:gap-2 ${hasMobileSecondaryRow ? 'ml-auto md:ml-0' : 'ml-auto'}`}>
          <div className="flex items-center gap-0.5 sm:gap-1" data-tour="utility-controls">
            <button
              type="button"
              onClick={() => setIsTourOpen(true)}
              aria-label="Open workspace guided tour"
              title="Open guided tour"
              className="min-w-[34px] min-h-[34px] md:min-w-0 md:min-h-0 p-1.5 bg-transparent border-0 shadow-none text-slate-600 hover:text-slate-950 dark:text-slate-300 dark:hover:text-white hover:scale-110 transition-transform duration-75 focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-[var(--selection-ring)] rounded-sm cursor-pointer shrink-0 flex items-center justify-center touch-manipulation"
            >
              <HelpCircle className="w-[18px] h-[18px]" />
            </button>
            <SpreadsheetThemeToggle overrideTheme={theme} onOverrideToggle={onToggleTheme} />
          </div>

          <div className="flex items-center gap-1.5 sm:gap-2" data-tour="save-export">
            <button
              type="button"
              onClick={() => void handleSave()}
              disabled={saveStatus === 'saving' || isLoading || !currentSheet.id}
              className="min-h-[34px] md:min-h-0 px-2.5 sm:px-3.5 md:px-4 py-1.5 font-semibold text-white bg-blue-600 hover:bg-blue-500 active:bg-blue-700 rounded-lg disabled:opacity-50 cursor-pointer flex items-center gap-1.5 text-[11px] sm:text-xs whitespace-nowrap touch-manipulation"
            >
              {saveStatus === 'saving' ? 'Saving...' : 'Save'}
            </button>

            <button
              type="button"
              onClick={() => void handleDirectExport()}
              disabled={isExporting || sheetKeys.length === 0}
              className="min-h-[34px] md:min-h-0 px-2.5 sm:px-3.5 md:px-4 py-1.5 font-semibold text-[var(--control-text)] bg-[var(--control-bg)] hover:bg-[var(--control-bg-hover)] border border-[var(--control-border)] rounded-lg disabled:opacity-50 cursor-pointer text-[11px] sm:text-xs whitespace-nowrap touch-manipulation"
            >
              {isExporting ? 'Exporting...' : 'Export'}
            </button>
          </div>
        </div>
      </div>

      {/* 2. Formula & Cell Address Bar */}
      <div
        data-tour="formula-bar"
        className="min-h-10 sm:h-10 px-2 sm:px-3 py-1 sm:py-0 bg-[var(--formula-bg)] border-b border-[var(--chrome-border)] flex items-center gap-1.5 sm:gap-2 shrink-0 text-xs"
      >
        <div className="min-w-[52px] sm:min-w-[72px] max-w-[96px] sm:max-w-[140px] truncate h-8 sm:h-7 px-2 sm:px-2.5 bg-[var(--formula-address-bg)] border border-[var(--control-border)] rounded-md flex items-center justify-center font-mono font-semibold tabular-nums text-[var(--control-text)] text-[11px] sm:text-xs select-none shrink-0">
          {selectionAddressLabel}
        </div>
        <div
          className="h-8 sm:h-7 px-1.5 sm:px-2 bg-[var(--control-bg)] border border-[var(--control-border)] rounded-md flex items-center justify-center text-[var(--control-muted)] font-serif italic text-[12px] sm:text-[13px] select-none shrink-0"
          aria-hidden="true"
        >
          fx
        </div>
        <input
          type="text"
          aria-label="Cell value or formula input"
          list={isEditing && autocompleteOptions.length > 0 ? 'mk-cell-autocomplete' : undefined}
          value={isEditing ? editValue : String(getCellValue(selectedCell.row, selectedCell.col) ?? '')}
          disabled={!currentSheet.id}
          onChange={(e) => {
            if (!isEditing) {
              setIsEditing(true);
            }
            setEditValue(e.target.value);
          }}
          onKeyDown={(e) => {
            if (!(e.ctrlKey || e.metaKey) || e.key.toLowerCase() !== 's') {
              e.stopPropagation();
            }
            if (e.key === 'Enter') {
              e.preventDefault();
              commitEdit(e.shiftKey ? -1 : 1, 0);
              rootRef.current?.focus({ preventScroll: true });
            } else if (e.key === 'Tab') {
              e.preventDefault();
              commitEdit(0, e.shiftKey ? -1 : 1);
              rootRef.current?.focus({ preventScroll: true });
            } else if (e.key === 'Escape') {
              e.preventDefault();
              cancelEdit();
              rootRef.current?.focus({ preventScroll: true });
            }
          }}
          onBlur={() => {
            if (isEditing) {
              commitEdit(0, 0);
            }
          }}
          className="flex-1 min-w-0 h-8 sm:h-7 px-2.5 sm:px-3 bg-[var(--formula-input-bg)] border border-[var(--control-border)] rounded-md text-[var(--control-text)] text-base sm:text-xs font-normal outline-none focus:border-[var(--selection-ring)] focus:ring-1 focus:ring-[var(--selection-ring)] disabled:opacity-50"
        />
        {isEditing && (
          <div className="flex items-center gap-1 sm:hidden shrink-0">
            <button
              type="button"
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => {
                commitEdit(0, 0);
                rootRef.current?.focus({ preventScroll: true });
              }}
              aria-label="Confirm cell edit"
              title="Confirm edit"
              className="h-8 w-8 flex items-center justify-center rounded-md bg-blue-600 text-white hover:bg-blue-500 cursor-pointer touch-manipulation"
            >
              <Check className="w-4 h-4" />
            </button>
            <button
              type="button"
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => {
                cancelEdit();
                rootRef.current?.focus({ preventScroll: true });
              }}
              aria-label="Cancel cell edit"
              title="Cancel edit"
              className="h-8 w-8 flex items-center justify-center rounded-md bg-[var(--control-bg)] border border-[var(--control-border)] text-[var(--control-text)] hover:bg-[var(--control-bg-hover)] cursor-pointer touch-manipulation"
            >
              <X className="w-4 h-4" />
            </button>
          </div>
        )}
      </div>

      {/* 2.1 Load Error Strip (if Supabase load fails) */}
      {loadError && (
        <div className="bg-red-50 dark:bg-red-950/60 border-b border-red-200 dark:border-red-800 px-3 py-1.5 flex flex-wrap items-center justify-between gap-2 text-xs text-red-800 dark:text-red-200 shrink-0">
          <span className="min-w-0 flex-1 break-words">{loadError}</span>
          {onReloadWorkspace && (
            <button
              type="button"
              onClick={() => void handleReloadLatest()}
              disabled={isReloading}
              className="px-2.5 py-1 bg-red-600 text-white rounded-md font-medium hover:bg-red-700 cursor-pointer shrink-0 touch-manipulation"
            >
              Retry
            </button>
          )}
        </div>
      )}

      {/* 2.1b Worksheet Operation / Sign-Out Error Strip */}
      {operationError && (
        <div className="bg-red-50 dark:bg-red-950/60 border-b border-red-200 dark:border-red-800 px-3 py-1.5 flex flex-wrap items-center justify-between gap-2 text-xs text-red-800 dark:text-red-200 shrink-0" role="alert">
          <span className="min-w-0 flex-1 break-words">{operationError}</span>
          <button
            type="button"
            onClick={() => setOperationError(null)}
            className="px-2 py-0.5 text-[11px] font-medium underline cursor-pointer shrink-0 touch-manipulation"
          >
            Dismiss
          </button>
        </div>
      )}

      {/* 2.2 Admin Import Result / Error Strip */}
      {user?.role === 'admin' && importNotice && (
        <div
          className={`border-b px-3 py-1.5 flex flex-wrap items-center justify-between gap-2 text-xs shrink-0 ${
            importNotice.type === 'error'
              ? 'bg-red-50 border-red-200 text-red-800 dark:bg-red-950/60 dark:border-red-800 dark:text-red-200'
              : 'bg-emerald-50 border-emerald-200 text-emerald-900 dark:bg-emerald-950/60 dark:border-emerald-800 dark:text-emerald-200'
          }`}
        >
          <span className="min-w-0 flex-1 break-words">{importNotice.message}</span>
          {onDismissImportNotice && (
            <button
              type="button"
              onClick={onDismissImportNotice}
              className="px-2 py-0.5 text-[11px] font-medium underline cursor-pointer shrink-0 touch-manipulation"
            >
              Dismiss
            </button>
          )}
        </div>
      )}

      {/* 3. Main Spreadsheet Grid Viewport (Rounded Cell Tiles + Negative-Space Gutters) */}
      <div
        ref={gridContainerRef}
        data-tour="grid-viewport"
        className="flex-1 min-h-0 overflow-auto touch-pan-x touch-pan-y overscroll-contain bg-[var(--grid-canvas)] mk-spreadsheet-scroll relative outline-none"
      >
        {!isLoading && sheetKeys.length === 0 ? (
          <div className="h-full flex flex-col items-center justify-center text-[var(--control-muted)] p-6 text-center">
            <p className="text-sm font-medium text-[var(--control-text)]">
              No worksheets found in the {workspace === 'doctor' ? 'Doctor' : 'Reception'} workspace yet.
            </p>
            <p className="text-xs text-[var(--control-muted)] mt-1">
              Click the <span className="font-mono font-bold">+</span> button at the bottom left to create a worksheet.
            </p>
            {user?.role === 'admin' && canAdminImport && onRequestInitialImport && (
              <button
                type="button"
                onClick={onRequestInitialImport}
                className="mt-4 px-4 py-1.5 text-xs font-semibold text-white bg-blue-600 hover:bg-blue-500 rounded-lg transition cursor-pointer"
              >
                Import Initial Workbooks (Reception &amp; Doctor)
              </button>
            )}
          </div>
        ) : (
          <div ref={tableWrapperRef} className="relative inline-block min-w-full align-top">
            <div
              ref={rowDropLineRef}
              style={{ display: 'none' }}
              aria-hidden="true"
              className="pointer-events-none absolute left-0 right-0 z-40 h-[3px] -mt-[1.5px] bg-blue-600 dark:bg-sky-400"
            >
              <span className="absolute left-0 -top-[3px] w-[9px] h-[9px] rounded-full bg-blue-600 dark:bg-sky-400" />
            </div>
            <div
              ref={colDropLineRef}
              style={{ display: 'none' }}
              aria-hidden="true"
              className="pointer-events-none absolute top-0 bottom-0 z-40 w-[3px] -ml-[1.5px] bg-blue-600 dark:bg-sky-400"
            >
              <span className="absolute top-0 -left-[3px] w-[9px] h-[9px] rounded-full bg-blue-600 dark:bg-sky-400" />
            </div>
            <table
              ref={tableRef}
              style={{ width: `${totalTableWidth}px` }}
              className="border-separate border-spacing-[2px] table-fixed text-xs bg-[var(--grid-canvas)]"
            >
            <colgroup>
              <col style={{ width: `${rowHeaderWidth}px` }} />
              {Array.from({ length: totalCols }).map((_, cIdx) => (
                <col
                  key={cIdx}
                  ref={(el) => {
                    colElementsRef.current[cIdx] = el;
                  }}
                  style={{ width: `${getColWidth(cIdx)}px` }}
                />
              ))}
            </colgroup>

            <thead ref={theadRef} className="sticky top-0 z-30 bg-[var(--grid-canvas)]">
              {/* Row 0: Column Letter Tiles (top rounded corners ~8px, square bottom corners) */}
              <tr style={{ height: `${DEFAULT_COL_LETTER_ROW_HEIGHT}px` }}>
                <th
                  onClick={handleSelectAllClick}
                  title="Select entire worksheet"
                  style={{ width: `${rowHeaderWidth}px`, minWidth: `${rowHeaderWidth}px` }}
                  className={`h-[28px] sticky left-0 z-50 rounded-[6px] select-none text-center cursor-pointer touch-manipulation shadow-[-4px_0_0_0_var(--grid-canvas)] ${
                    selection.mode === 'all'
                      ? 'bg-[var(--header-tile-active)] text-[var(--header-text-active)]'
                      : 'bg-[var(--header-tile-bg)] hover:bg-[var(--header-tile-hover)] text-[var(--header-text)]'
                  }`}
                >
                  <span className="text-[10px] leading-none">◢</span>
                </th>

                {Array.from({ length: totalCols }).map((_, cIdx) => {
                  const width = getColWidth(cIdx);
                  const isColSelected = isColHeaderHighlighted(cIdx);
                  return (
                    <th
                      key={cIdx}
                      ref={(el) => {
                        headerLetterThRefs.current[cIdx] = el;
                      }}
                      style={{ width: `${width}px`, minWidth: `${width}px` }}
                      onClick={(e) => handleColHeaderClick(cIdx, e.shiftKey)}
                      onContextMenu={(e) => {
                        e.preventDefault();
                        e.stopPropagation();
                        openColContextMenu(cIdx, e.clientX, e.clientY);
                      }}
                      onPointerDown={(e) => {
                        if (e.pointerType !== 'mouse') {
                          startGridLongPress(e, (x, y) => openColContextMenu(cIdx, x, y));
                        }
                        handleColDragStart(cIdx, e);
                      }}
                      title={`Column ${getColumnLetter(cIdx)} (Click to select, drag to reorder, right-click for options)`}
                      className={`h-[28px] text-center text-[11px] tracking-wide select-none relative touch-manipulation rounded-t-[8px] rounded-b-none ${
                        isColSelected && isFullColSelectionMode
                          ? 'bg-[var(--header-tile-active)] text-[var(--header-text-active)] font-semibold cursor-grab active:cursor-grabbing'
                          : isColSelected
                            ? 'bg-[var(--header-tile-range)] text-[var(--col-title-text)] font-semibold cursor-pointer'
                            : 'bg-[var(--header-tile-bg)] hover:bg-[var(--header-tile-hover)] font-medium text-[var(--header-text)] cursor-pointer'
                      }`}
                    >
                      <span>{getColumnLetter(cIdx)}</span>
                      <div
                        onPointerDown={(e) => handleColResizeStart(cIdx, e)}
                        onClick={(e) => e.stopPropagation()}
                        onDoubleClick={(e) => e.stopPropagation()}
                        title="Drag to resize column width"
                        className="absolute right-0 top-1 bottom-0 w-3 md:w-1.5 cursor-col-resize touch-none hover:bg-[var(--resize-accent)]/70 rounded-full z-20"
                      />
                    </th>
                  );
                })}
              </tr>

              {/* Row 1: Column Title Header Tiles */}
              <tr
                ref={(el) => {
                  rowElementsRef.current[0] = el;
                }}
                style={{ height: `${getManualRowHeight(0) ?? DEFAULT_HEADER_ROW_HEIGHT}px` }}
              >
                <th
                  ref={headerRowThRef}
                  onClick={(e) => handleRowHeaderClick(0, e.shiftKey)}
                  style={{
                    width: `${rowHeaderWidth}px`,
                    minWidth: `${rowHeaderWidth}px`,
                    height: `${getManualRowHeight(0) ?? DEFAULT_HEADER_ROW_HEIGHT}px`
                  }}
                  className={`px-1 font-semibold text-center text-[11px] tabular-nums align-middle sticky left-0 z-40 select-none cursor-pointer touch-manipulation rounded-l-[8px] rounded-r-none shadow-[-4px_0_0_0_var(--grid-canvas)] ${
                    isRowHeaderHighlighted(0) && isFullRowSelectionMode
                      ? 'bg-[var(--header-tile-active)] text-[var(--header-text-active)]'
                      : isRowHeaderHighlighted(0)
                        ? 'bg-[var(--header-tile-range)] text-[var(--col-title-text)]'
                        : 'bg-[var(--header-tile-bg)] hover:bg-[var(--header-tile-hover)] text-[var(--header-text)]'
                  }`}
                >
                  <span>1</span>
                  <div
                    onPointerDown={(e) => handleRowResizeStart(0, e)}
                    onClick={(e) => e.stopPropagation()}
                    onDoubleClick={(e) => e.stopPropagation()}
                    title="Drag to resize header row height"
                    className="absolute left-1 right-0 bottom-0 h-3 md:h-1.5 cursor-row-resize touch-none hover:bg-[var(--resize-accent)]/70 rounded-full z-20"
                  />
                </th>

                {Array.from({ length: totalCols }).map((_, cIdx) => {
                  const isAnchor = selectedCell.row === 0 && selectedCell.col === cIdx;
                  const isInSel = isCellInSelection(0, cIdx);
                  const isCellEditing = isAnchor && isEditing && selection.mode === 'cell';
                  const isHeaderDirty = dirtyHeaderCols.has(cIdx);
                  const cellVal = getCellValue(0, cIdx);
                  const width = getColWidth(cIdx);
                  const manualHeight = getManualRowHeight(0);
                  const colMeta = colMetaList[cIdx];
                  const alignClass = getAlignmentClass(colMeta?.align ?? 'left');
                  const headerBgClass = getUniformHeaderClasses(
                    isInSel && !(isAnchor && selection.mode === 'cell'),
                    isHeaderDirty
                  );

                  return (
                    <td
                      key={cIdx}
                      ref={(el) => {
                        headerTitleTdRefs.current[cIdx] = el;
                      }}
                      style={{
                        width: `${width}px`,
                        minWidth: `${width}px`,
                        ...(manualHeight ? { height: `${manualHeight}px` } : {})
                      }}
                      onClick={(e) => handleCellClick(0, cIdx, e.shiftKey, e.detail)}
                      onDoubleClick={(e) => handleCellDoubleClick(0, cIdx, e.detail)}
                      className={`px-2 sm:px-3 py-2 font-semibold text-xs leading-[20px] tracking-[0.01em] align-middle relative rounded-t-[2px] rounded-b-[5px] touch-manipulation ${alignClass} ${headerBgClass} ${
                        isAnchor && !isCellEditing
                          ? 'ring-2 ring-inset ring-[var(--selection-ring)] z-[25]'
                          : 'z-20'
                      }`}
                    >
                      {isHeaderDirty && !isCellEditing && (
                        <span
                          className="pointer-events-none absolute top-1 right-1 w-1.5 h-1.5 rounded-full bg-amber-500 dark:bg-amber-300"
                          title="Unsaved header edit"
                        />
                      )}
                      {isCellEditing ? (
                        <input
                          ref={setEditInputRef}
                          type="text"
                          value={editValue}
                          onChange={(e) => setEditValue(e.target.value)}
                          onBlur={() => commitEdit(0, 0)}
                          className={`absolute inset-0 w-full h-full px-2 sm:px-3 text-base sm:text-xs font-semibold rounded-[5px] bg-[var(--formula-input-bg)] text-[var(--control-text)] border-2 border-[var(--selection-ring)] outline-none z-20 ${alignClass}`}
                        />
                      ) : (
                        <div
                          className={
                            isWrapEnabled
                              ? 'whitespace-normal break-words leading-[20px]'
                              : 'truncate whitespace-nowrap leading-[20px]'
                          }
                          style={
                            manualHeight
                              ? { maxHeight: `${Math.max(20, manualHeight - 16)}px`, overflowY: 'auto' }
                              : undefined
                          }
                        >
                          {cellVal !== undefined && cellVal !== null ? String(cellVal) : ''}
                        </div>
                      )}
                    </td>
                  );
                })}
              </tr>
            </thead>

            <tbody ref={tbodyRef}>
              {Array.from({ length: displayRowsCount - 1 }).map((_, offsetIdx) => {
                const rIdx = offsetIdx + 1;
                const manualHeight = getManualRowHeight(rIdx);
                const isRowHighlighted =
                  rIdx >= selectionBounds.rMin && rIdx <= selectionBounds.rMax;
                const anchorColForRow = selectedCell.row === rIdx ? selectedCell.col : -1;
                const isEditingRow =
                  isEditing && selection.mode === 'cell' && selectedCell.row === rIdx;

                return (
                  <SpreadsheetDataRow
                    key={rIdx}
                    gridRowIdx={rIdx}
                    rowCells={currentSheet.rows[rIdx - 1]}
                    totalCols={totalCols}
                    rowHeaderWidth={rowHeaderWidth}
                    colMetaList={colMetaList}
                    manualHeight={manualHeight}
                    isWrapEnabled={isWrapEnabled}
                    isRowHighlighted={isRowHighlighted}
                    isFullRowMode={isFullRowSelectionMode}
                    anchorCol={anchorColForRow}
                    isEditingRow={isEditingRow}
                    editValueIfEditing={isEditingRow ? editValue : ''}
                    editInputListId={
                      isEditingRow && autocompleteOptions.length > 0
                        ? 'mk-cell-autocomplete'
                        : undefined
                    }
                    selectedColStart={isRowHighlighted ? selectionBounds.cMin : -1}
                    selectedColEnd={isRowHighlighted ? selectionBounds.cMax : -1}
                    dirtyColsKey={dirtyColsByRow.get(rIdx - 1)}
                    workspace={workspace}
                    paidColIdx={receptionSpecialCols.paidColIdx}
                    onRowHeaderClick={handleRowHeaderClick}
                    onRowHeaderContextMenu={handleRowHeaderContextMenuCb}
                    onRowHeaderTouchStart={handleRowHeaderTouchStartCb}
                    onRowDragStart={handleRowDragStart}
                    onRowResizeStart={handleRowResizeStart}
                    onCellClick={handleCellClick}
                    onCellDoubleClick={handleCellDoubleClick}
                    onCellContextMenu={handleCellContextMenuCb}
                    onCellTouchStart={handleCellTouchStartCb}
                    onPaidValueChange={handlePaidValueChangeCb}
                    onEditValueChange={setEditValue}
                    onCommitEdit={handleCommitEditBlur}
                    setEditInputRef={setEditInputRef}
                    setRowElementRef={setRowElementRef}
                  />
                );
              })}
            </tbody>
            </table>
          </div>
        )}
      </div>

      {/* 4. Worksheet Tabs along the bottom with '+' button and Sign Out at bottom-right */}
      <div className="min-h-10 sm:h-10 pb-[env(safe-area-inset-bottom)] bg-[var(--tab-strip-bg)] border-t border-[var(--chrome-border)] flex items-center justify-between px-2 sm:px-3 gap-2 sm:gap-3 shrink-0">
        <div
          data-tour="worksheet-tabs"
          className="mk-tab-strip-scroll flex items-center gap-1.5 overflow-x-auto min-w-0 flex-1 py-1"
        >
          {renderedSheetKeys.map((name) => {
            const isActive = activeSheetName === name;
            const isBeingRenamed = renamingSheetName === name;
            const isBeingDragged = dragTabState?.sourceName === name;
            const isDeletingThis = isDeletingSheet && deletingSheetName === name;

            return (
              <div
                key={name}
                ref={(el) => {
                  tabElementRefs.current[name] = el;
                }}
                onPointerDown={(e) => handleTabPointerDown(name, e)}
                onContextMenu={(e) => {
                  e.preventDefault();
                  e.stopPropagation();
                  openTabContextMenuFromRect(name, e.clientX, e.clientY);
                }}
                onClick={() => {
                  if (suppressNextTabClickRef.current) return;
                  if (isEditing) commitEdit(0, 0);
                  if (!isBeingRenamed) {
                    setActiveSheetName(name);
                    const latestRow = findLatestRecordGridRow(sheets[name]);
                    setSelection({
                      mode: 'cell',
                      anchorRow: latestRow,
                      anchorCol: 0,
                      focusRow: latestRow,
                      focusCol: 0
                    });
                    setScrollPositionRequest((prev) => ({
                      sheetName: name,
                      targetRow: latestRow,
                      nonce: prev.nonce + 1
                    }));
                  }
                }}
                onDoubleClick={(e) => {
                  e.stopPropagation();
                  if (!isDeletingSheet) {
                    handleStartRename(name);
                  }
                }}
                className={`group relative min-h-[30px] sm:min-h-0 px-2.5 sm:px-3 py-1 text-xs font-medium rounded-md border cursor-pointer select-none whitespace-nowrap flex items-center gap-1 touch-pan-x shrink-0 ${
                  isBeingDragged
                    ? 'bg-blue-50 border-blue-600 text-blue-900 dark:bg-sky-900/70 dark:border-sky-400 dark:text-sky-100 shadow-md ring-2 ring-blue-500/40 dark:ring-sky-400/50 scale-[1.02] z-20'
                    : isActive
                      ? 'bg-[var(--tab-active-bg)] border-[var(--control-border)] border-b-2 border-b-[var(--tab-active-border)] text-[var(--tab-text-active)] font-semibold'
                      : 'bg-[var(--tab-bg)] border-transparent text-[var(--tab-text)] hover:bg-[var(--tab-hover)] hover:text-[var(--tab-text-active)]'
                } ${isDeletingThis ? 'opacity-50' : ''}`}
                title={`Worksheet: ${name} (Double-click to rename, right-click or hold for options, drag to reorder)`}
              >
                {isBeingDragged && (
                  <span className="w-1 h-3 rounded-full bg-blue-600 dark:bg-sky-400 shrink-0" aria-hidden="true" />
                )}
                {isBeingRenamed ? (
                  <input
                    ref={renameInputRef}
                    type="text"
                    value={renameInputValue}
                    onChange={(e) => setRenameInputValue(e.target.value)}
                    onPointerDown={(e) => e.stopPropagation()}
                    onMouseDown={(e) => e.stopPropagation()}
                    onDoubleClick={(e) => e.stopPropagation()}
                    onCopy={(e) => e.stopPropagation()}
                    onCut={(e) => e.stopPropagation()}
                    onPaste={(e) => e.stopPropagation()}
                    onKeyUp={(e) => e.stopPropagation()}
                    onKeyDown={(e) => {
                      e.stopPropagation();
                      if (e.key === 'Enter') {
                        e.preventDefault();
                        void handleCommitRename(name, renameInputValue);
                      } else if (e.key === 'Escape') {
                        e.preventDefault();
                        renameCommitGuardRef.current = true;
                        setRenamingSheetName(null);
                        rootRef.current?.focus({ preventScroll: true });
                      }
                    }}
                    onBlur={() => {
                      if (!renameCommitGuardRef.current) {
                        void handleCommitRename(name, renameInputValue);
                      }
                    }}
                    onClick={(e) => e.stopPropagation()}
                    maxLength={31}
                    aria-label="Rename worksheet"
                    className="px-1.5 py-0.5 text-sm sm:text-xs font-semibold bg-[var(--formula-input-bg)] border border-[var(--selection-ring)] rounded-xs text-[var(--control-text)] outline-none w-28 sm:w-32 select-text"
                  />
                ) : (
                  <>
                    <span>{name}</span>
                    {isActive && (
                      <button
                        type="button"
                        onPointerDown={(e) => e.stopPropagation()}
                        onClick={(e) => {
                          e.stopPropagation();
                          const r = e.currentTarget.getBoundingClientRect();
                          openTabContextMenuFromRect(name, r.left, r.top);
                        }}
                        aria-label={`Worksheet options for ${name}`}
                        title="Worksheet options (Rename, Delete, Reorder)"
                        className="md:hidden -mr-1 ml-0.5 p-1 text-[var(--control-muted)] hover:text-[var(--control-text)] rounded cursor-pointer touch-manipulation"
                      >
                        <MoreVertical className="w-3.5 h-3.5" />
                      </button>
                    )}
                  </>
                )}
              </div>
            );
          })}

          <button
            type="button"
            onClick={() => void handleAddSheetImmediate()}
            disabled={isCreatingSheet || isDeletingSheet}
            title="Create new worksheet"
            aria-label="Create new worksheet"
            className="h-8 w-8 sm:h-7 sm:w-7 flex items-center justify-center rounded-md bg-[var(--control-bg)] hover:bg-[var(--control-bg-hover)] border border-[var(--control-border)] text-[var(--control-text)] font-bold text-sm cursor-pointer ml-0.5 shrink-0 disabled:opacity-50 touch-manipulation"
          >
            {isCreatingSheet ? '...' : '+'}
          </button>
        </div>

        {/* Bottom-Right Corner: Sign Out */}
        <div className="flex items-center gap-2 sm:gap-2.5 shrink-0 pl-2 sm:pl-3 border-l border-[var(--chrome-border)]">
          <span
            className="hidden sm:inline text-[var(--control-muted)] font-mono text-[11px] max-w-[120px] md:max-w-[180px] truncate"
            title={user?.role}
          >
            {staffLabel}
          </span>
          <button
            type="button"
            onClick={() => void handleSignOutClick()}
            disabled={isLoggingOut}
            className="min-h-[34px] sm:min-h-0 text-[11px] sm:text-xs text-[var(--control-text)] hover:text-red-600 dark:hover:text-red-400 font-semibold px-2 sm:px-2.5 py-1 rounded-md hover:bg-[var(--control-bg-hover)] cursor-pointer disabled:opacity-60 disabled:cursor-not-allowed flex items-center gap-1 sm:gap-1.5 whitespace-nowrap touch-manipulation"
            title="Sign Out"
          >
            {isLoggingOut ? (
              <>
                <Loader2 className="w-3.5 h-3.5 animate-spin shrink-0" />
                <span className="hidden sm:inline">Signing out...</span>
              </>
            ) : (
              <>
                <LogOut className="w-3.5 h-3.5 shrink-0" />
                <span>Sign Out</span>
              </>
            )}
          </button>
        </div>
      </div>



      {/* 6. Safeguard Confirmation Modal for Bulk / Row / Column Clear */}
      <ConfirmDialog
        isOpen={pendingBulkClear !== null}
        title="Clear Selected Cells"
        message={
          pendingBulkClear
            ? `You are about to clear ${pendingBulkClear.populatedCount} populated cell${
                pendingBulkClear.populatedCount === 1 ? '' : 's'
              }${
                pendingBulkClear.includesHeader ? ' (including worksheet column header titles)' : ''
              }. Changes will not be saved until you click Save or press Ctrl+S.`
            : ''
        }
        confirmLabel="Clear Cells"
        confirmVariant="danger"
        onConfirm={() => {
          if (pendingBulkClear) {
            applyCellUpdates(pendingBulkClear.updates, {
              selectionBefore: selection,
              selectionAfter: selection
            });
            setPendingBulkClear(null);
            rootRef.current?.focus({ preventScroll: true });
          }
        }}
        onCancel={() => {
          setPendingBulkClear(null);
          rootRef.current?.focus({ preventScroll: true });
        }}
      />

      {/* 5b. Unified Spreadsheet Context Menu */}
      <SpreadsheetContextMenu
        menu={contextMenu}
        isBusy={isStructuralBusy}
        onAction={handleContextMenuAction}
        onClose={() => setContextMenu(null)}
      />

      {/* 5c. Column Width Resize Modal */}
      <ColumnWidthModal
        isOpen={resizeColDialog !== null}
        columnLabel={resizeColDialog ? getColumnLetter(resizeColDialog.colIndex) : ''}
        initialWidth={resizeColDialog?.currentWidth ?? DEFAULT_COL_WIDTH}
        defaultWidth={resizeColDialog?.defaultWidth ?? DEFAULT_COL_WIDTH}
        minWidth={MIN_COL_WIDTH}
        maxWidth={MAX_COL_WIDTH}
        onApply={(newWidth) => {
          if (!resizeColDialog) return;
          const targetCol = resizeColDialog.colIndex;
          const targetSheetKey = currentSheet.id || activeSheetName;
          setColWidthsBySheet((prev) => ({
            ...prev,
            [targetSheetKey]: {
              ...(prev[targetSheetKey] || {}),
              [targetCol]: newWidth
            }
          }));
          setResizeColDialog(null);
          scheduleAutosave(200);
          rootRef.current?.focus({ preventScroll: true });
        }}
        onResetDefault={() => {
          if (!resizeColDialog) return;
          const targetCol = resizeColDialog.colIndex;
          const targetSheetKey = currentSheet.id || activeSheetName;
          setColWidthsBySheet((prev) => {
            const nextMap = { ...(prev[targetSheetKey] || {}) };
            delete nextMap[targetCol];
            return {
              ...prev,
              [targetSheetKey]: nextMap
            };
          });
          setResizeColDialog(null);
          scheduleAutosave(200);
          rootRef.current?.focus({ preventScroll: true });
        }}
        onClose={() => {
          setResizeColDialog(null);
          rootRef.current?.focus({ preventScroll: true });
        }}
      />

      {/* 6b. Structural Operation / Overwrite Confirmation Modal */}
      <ConfirmDialog
        isOpen={pendingConfirmAction !== null}
        title={pendingConfirmAction?.title ?? 'Confirm Action'}
        message={pendingConfirmAction?.message ?? ''}
        confirmLabel={pendingConfirmAction?.confirmLabel ?? 'Confirm'}
        confirmVariant={pendingConfirmAction?.confirmVariant ?? 'primary'}
        onConfirm={() => {
          if (pendingConfirmAction) {
            const actionToRun = pendingConfirmAction.run;
            setPendingConfirmAction(null);
            actionToRun();
            rootRef.current?.focus({ preventScroll: true });
          }
        }}
        onCancel={() => {
          setPendingConfirmAction(null);
          rootRef.current?.focus({ preventScroll: true });
        }}
      />

      {/* Lightweight Cell Autocomplete Datalist (only populated while editing) */}
      {isEditing && autocompleteOptions.length > 0 && (
        <datalist id="mk-cell-autocomplete">
          {autocompleteOptions.map((optionVal) => (
            <option key={optionVal} value={optionVal} />
          ))}
        </datalist>
      )}

      {/* 7. Role-Aware Guided Walkthrough Tour */}
      {isActiveWorkspace && user && (
        <GuidedTour
          isOpen={isTourOpen}
          role={user.role}
          workspace={workspace}
          hasWorkspaceSwitcher={visibleWorkspaces.length > 1}
          onClose={handleCloseTour}
          onComplete={handleCompleteTour}
        />
      )}
    </div>
  );
}
