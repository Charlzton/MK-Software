'use client';

import React, { useState, useEffect, useRef, useCallback, useMemo } from 'react';
import * as XLSX from 'xlsx';
import { LogOut } from 'lucide-react';
import { User } from '@/lib/types';
import { ConfirmDialog } from '@/components/ConfirmDialog';
import { validateWorksheetName, type WorkspaceSheetState } from '@/lib/store';
import { exportSheetsToExcel } from '@/lib/excelUtils';
import {
  getColumnVisualMeta,
  ClinicalColumnCategory,
  ColumnVisualMeta
} from '@/lib/clinicalColumns';

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
}

interface DirtyCellEdit {
  sheetId: string;
  sheetName: string;
  rowIndex: number;
  colIndex: number;
  value: string | number;
  expectedCellVersion: number;
}

interface DirtyHeaderEdit {
  sheetId: string;
  sheetName: string;
  colIndex: number;
  newHeader: string;
  expectedHeaderVersion: number;
}

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
  onLogout: () => void;
  isLoading?: boolean;
  loadError?: string | null;
  defaultSheetName?: string;
  canAdminImport?: boolean;
  onRequestInitialImport?: () => void;
  importNotice?: { type: 'success' | 'info' | 'error'; message: string } | null;
  onDismissImportNotice?: () => void;
}

const ROW_HEADER_COL_WIDTH = 48;
const DEFAULT_COL_WIDTH = 110;
const MIN_COL_WIDTH = 56;
const MAX_COL_WIDTH = 600;
const DEFAULT_DATA_ROW_HEIGHT = 28;
const DEFAULT_HEADER_ROW_HEIGHT = 34;
const MIN_ROW_HEIGHT = 24;
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

function getClinicalHeaderClasses(
  category: ClinicalColumnCategory,
  isSelected: boolean,
  isDoctorWorkspace: boolean
): string {
  if (isSelected) {
    return 'bg-[#cfe0f5] text-slate-900';
  }
  if (!isDoctorWorkspace) {
    return 'bg-[#e3ebf6] text-slate-900';
  }
  switch (category) {
    case 'patient_id':
      return 'bg-sky-50 text-slate-900 border-t-2 border-t-sky-600';
    case 'prescription':
      return 'bg-indigo-50 text-slate-900 border-t-2 border-t-indigo-600';
    case 'clinical_finding':
      return 'bg-teal-50 text-slate-900 border-t-2 border-t-teal-600';
    case 'medication_treatment':
      return 'bg-emerald-50 text-slate-900 border-t-2 border-t-emerald-600';
    case 'surgery':
      return 'bg-rose-50 text-slate-900 border-t-2 border-t-rose-600';
    default:
      return 'bg-[#e3ebf6] text-slate-900';
  }
}

function getClinicalCellBgClass(
  category: ClinicalColumnCategory,
  isSelected: boolean,
  isDoctorWorkspace: boolean
): string {
  if (isSelected) {
    return 'bg-sky-100/80';
  }
  if (!isDoctorWorkspace) {
    return 'bg-white';
  }
  switch (category) {
    case 'patient_id':
      return 'bg-sky-50/25';
    case 'prescription':
      return 'bg-indigo-50/25';
    case 'clinical_finding':
      return 'bg-teal-50/25';
    case 'medication_treatment':
      return 'bg-emerald-50/25';
    case 'surgery':
      return 'bg-rose-50/25';
    default:
      return 'bg-white';
  }
}

/**
 * Determine which grid row (1 = Excel Row 2, i.e. dataRowIdx = 0) should be opened first:
 * 1. Most recently updated populated row via rowUpdatedAt timestamps when strictly distinguishable.
 * 2. Fallback to the highest populated row_index when timestamps are missing or identical.
 * 3. Fallback to 1 (first data row) when the worksheet has no populated rows.
 */
function findLatestRecordGridRow(sheet: SheetState | undefined): number {
  if (!sheet || !Array.isArray(sheet.rows) || sheet.rows.length === 0) {
    return 1;
  }

  let bestDataRowIdx = -1;
  let bestTimestamp = Number.NEGATIVE_INFINITY;

  for (let dataRowIdx = 0; dataRowIdx < sheet.rows.length; dataRowIdx++) {
    const row = sheet.rows[dataRowIdx];
    const hasData = Array.isArray(row) && row.some(isNonEmptyCell);
    if (!hasData) continue;

    const rawUpdated = sheet.rowUpdatedAt?.[dataRowIdx];
    const parsedTs = rawUpdated ? Date.parse(rawUpdated) : Number.NaN;
    const effectiveTs = Number.isFinite(parsedTs) ? parsedTs : Number.NEGATIVE_INFINITY;

    if (
      bestDataRowIdx === -1 ||
      effectiveTs > bestTimestamp ||
      (effectiveTs === bestTimestamp && dataRowIdx > bestDataRowIdx)
    ) {
      bestDataRowIdx = dataRowIdx;
      bestTimestamp = effectiveTs;
    }
  }

  return bestDataRowIdx >= 0 ? bestDataRowIdx + 1 : 1;
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

interface SpreadsheetDataRowProps {
  gridRowIdx: number;
  rowCells: (string | number)[] | undefined;
  totalCols: number;
  colMetaList: ColumnVisualMeta[];
  isDoctorWorkspace: boolean;
  manualHeight: number | undefined;
  isWrapEnabled: boolean;
  isRowHighlighted: boolean;
  isFullRowMode: boolean;
  anchorCol: number;
  isEditingRow: boolean;
  editValueIfEditing: string;
  selectedColStart: number;
  selectedColEnd: number;
  onRowHeaderClick: (gridRowIdx: number, shiftKey: boolean) => void;
  onRowResizeStart: (gridRowIdx: number, e: React.PointerEvent<HTMLDivElement>) => void;
  onCellClick: (gridRowIdx: number, colIdx: number, shiftKey: boolean) => void;
  onCellDoubleClick: (gridRowIdx: number, colIdx: number, detail: number) => void;
  onEditValueChange: (val: string) => void;
  onCommitEdit: () => void;
  setEditInputRef: (el: HTMLInputElement | null) => void;
  setRowElementRef: (gridRowIdx: number, el: HTMLTableRowElement | null) => void;
}

const SpreadsheetDataRow = React.memo(function SpreadsheetDataRow({
  gridRowIdx,
  rowCells,
  totalCols,
  colMetaList,
  isDoctorWorkspace,
  manualHeight,
  isWrapEnabled,
  isRowHighlighted,
  isFullRowMode,
  anchorCol,
  isEditingRow,
  editValueIfEditing,
  selectedColStart,
  selectedColEnd,
  onRowHeaderClick,
  onRowResizeStart,
  onCellClick,
  onCellDoubleClick,
  onEditValueChange,
  onCommitEdit,
  setEditInputRef,
  setRowElementRef
}: SpreadsheetDataRowProps) {
  const excelRowNum = gridRowIdx + 1;

  return (
    <tr
      ref={(el) => setRowElementRef(gridRowIdx, el)}
      style={
        manualHeight
          ? { height: `${manualHeight}px` }
          : !isWrapEnabled
            ? { height: `${DEFAULT_DATA_ROW_HEIGHT}px` }
            : undefined
      }
    >
      <th
        onClick={(e) => onRowHeaderClick(gridRowIdx, e.shiftKey)}
        className={`w-12 min-w-12 border-r border-b border-slate-300 text-center text-[11px] tabular-nums sticky left-0 z-[15] select-none cursor-pointer relative ${
          isRowHighlighted && isFullRowMode
            ? 'bg-slate-700 text-white font-semibold'
            : isRowHighlighted
              ? 'bg-slate-200 text-slate-900 font-medium'
              : 'bg-slate-100 hover:bg-slate-200/70 text-slate-600 font-normal'
        }`}
      >
        <span>{excelRowNum}</span>
        <div
          onPointerDown={(e) => onRowResizeStart(gridRowIdx, e)}
          onClick={(e) => e.stopPropagation()}
          onDoubleClick={(e) => e.stopPropagation()}
          title="Drag to resize row height"
          className="absolute left-0 right-0 bottom-0 h-2 md:h-1.5 cursor-row-resize touch-none hover:bg-sky-500/70 z-20"
        />
      </th>

      {Array.from({ length: totalCols }).map((_, cIdx) => {
        const isAnchor = anchorCol === cIdx;
        const isInSel =
          selectedColStart !== -1 && cIdx >= selectedColStart && cIdx <= selectedColEnd;
        const isCellEditing = isAnchor && isEditingRow;
        const cellVal = rowCells ? (rowCells[cIdx] ?? '') : '';
        const colMeta = colMetaList[cIdx];
        const category = colMeta?.category ?? 'normal';
        const alignClass = getAlignmentClass(colMeta?.align ?? 'left');
        const showRangeTint = isInSel && !(isAnchor && !isFullRowMode && selectedColStart === selectedColEnd);
        const bgClass = getClinicalCellBgClass(category, showRangeTint, isDoctorWorkspace);

        return (
          <td
            key={cIdx}
            onClick={(e) => onCellClick(gridRowIdx, cIdx, e.shiftKey)}
            onDoubleClick={(e) => onCellDoubleClick(gridRowIdx, cIdx, e.detail)}
            className={`min-h-[28px] px-2.5 py-1 border-r border-b border-slate-300 relative text-xs leading-[18px] text-slate-900 font-normal tabular-nums align-top ${alignClass} ${bgClass} ${
              isAnchor && !isCellEditing
                ? 'outline-2 outline-slate-900 outline-offset-[-2px] z-10'
                : ''
            }`}
          >
            {isCellEditing ? (
              <input
                ref={setEditInputRef}
                type="text"
                value={editValueIfEditing}
                onChange={(e) => onEditValueChange(e.target.value)}
                onBlur={onCommitEdit}
                className={`absolute inset-0 w-full h-full px-2.5 text-xs bg-white text-slate-900 border-2 border-slate-900 outline-none z-10 ${alignClass}`}
              />
            ) : (
              <div
                className={
                  isWrapEnabled
                    ? 'whitespace-normal break-words leading-[18px] min-h-[18px]'
                    : 'truncate whitespace-nowrap leading-[18px] h-[18px]'
                }
                style={
                  manualHeight
                    ? { maxHeight: `${Math.max(18, manualHeight - 8)}px`, overflowY: 'auto' }
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
    const keys = Object.keys(initialSheets);
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

  const [saveStatus, setSaveStatus] = useState<'idle' | 'saving' | 'saved' | 'error'>('idle');
  const [saveMessage, setSaveMessage] = useState<string>('');
  const [saveError, setSaveError] = useState<string | null>(null);
  const [hasConflict, setHasConflict] = useState<boolean>(false);
  const isSavingRef = useRef<boolean>(false);

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

    const keys = Object.keys(initialSheets);
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
    dirtyCellsRef.current = dirtyCells;
  }, [dirtyCells]);

  useEffect(() => {
    dirtyHeadersRef.current = dirtyHeaders;
  }, [dirtyHeaders]);

  const hasUnsavedChanges =
    Object.keys(dirtyCells).length > 0 || Object.keys(dirtyHeaders).length > 0;

  useEffect(() => {
    onUnsavedStatusChange(hasUnsavedChanges);
  }, [hasUnsavedChanges, onUnsavedStatusChange]);

  // Per-sheet column widths & manual row heights (kept cleanly in React state, never interfering with saves)
  const [colWidthsBySheet, setColWidthsBySheet] = useState<Record<string, Record<number, number>>>({});
  const [rowHeightsBySheet, setRowHeightsBySheet] = useState<Record<string, Record<number, number>>>({});
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
  const tableRef = useRef<HTMLTableElement>(null);
  const theadRef = useRef<HTMLTableSectionElement>(null);
  const tbodyRef = useRef<HTMLTableSectionElement>(null);
  const editInputRef = useRef<HTMLInputElement | null>(null);

  // Direct DOM refs for fast rAF drag resizing without per-frame React re-renders
  const colElementsRef = useRef<Record<number, HTMLTableColElement | null>>({});
  const headerLetterThRefs = useRef<Record<number, HTMLTableCellElement | null>>({});
  const headerTitleTdRefs = useRef<Record<number, HTMLTableCellElement | null>>({});
  const rowElementsRef = useRef<Record<number, HTMLTableRowElement | null>>({});
  const headerRowThRef = useRef<HTMLTableCellElement | null>(null);

  // Worksheet management state: immediate add, inline rename, right-click context menu, drag & drop reorder
  const [isCreatingSheet, setIsCreatingSheet] = useState(false);
  const [isReloading, setIsReloading] = useState(false);
  const [renamingSheetName, setRenamingSheetName] = useState<string | null>(null);
  const [renameInputValue, setRenameInputValue] = useState<string>('');
  const [, setIsRenamingSheet] = useState(false);
  const [tabContextMenu, setTabContextMenu] = useState<{
    sheetName: string;
    x: number;
    y: number;
  } | null>(null);
  const [pendingDeleteSheet, setPendingDeleteSheet] = useState<string | null>(null);
  const [isDeletingSheet, setIsDeletingSheet] = useState(false);
  const [draggedSheet, setDraggedSheet] = useState<string | null>(null);
  const [dragOverSheet, setDragOverSheet] = useState<string | null>(null);
  const renameInputRef = useRef<HTMLInputElement | null>(null);
  const [pendingBulkClear, setPendingBulkClear] = useState<{
    updates: Array<{ row: number; col: number; value: string | number }>;
    populatedCount: number;
    includesHeader: boolean;
  } | null>(null);

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

  const isDoctorWorkspace = workspace === 'doctor';

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
    let sum = ROW_HEADER_COL_WIDTH;
    for (let c = 0; c < totalCols; c++) {
      sum += getColWidth(c);
    }
    return sum;
  }, [totalCols, getColWidth]);

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

  // Auto-position scroll viewport to the latest record when a worksheet opens or becomes visible
  const isActiveWorkspace = workspace === activeWorkspace;
  useEffect(() => {
    if (!isActiveWorkspace || isLoading) return;
    if (scrollPositionRequest.nonce === 0) return;
    if (scrollPositionRequest.sheetName !== activeSheetName) return;
    if (lastExecutedScrollNonceRef.current === scrollPositionRequest.nonce) return;

    const container = gridContainerRef.current;
    if (!container || container.clientHeight === 0) return;

    lastExecutedScrollNonceRef.current = scrollPositionRequest.nonce;
    const targetRow = scrollPositionRequest.targetRow;

    const rafId = window.requestAnimationFrame(() => {
      const el = gridContainerRef.current;
      if (!el || el.clientHeight === 0) return;

      const theadHeight = theadRef.current?.offsetHeight ?? 58;
      if (targetRow <= 1) {
        el.scrollTop = 0;
        el.scrollLeft = 0;
        return;
      }

      const tbodyEl = tbodyRef.current;
      const targetTr = tbodyEl?.children[targetRow - 1] as HTMLElement | undefined;
      const viewportDataHeight = Math.max(120, el.clientHeight - theadHeight);

      if (targetTr) {
        const rowTopInTbody = Math.max(0, targetTr.offsetTop - theadHeight);
        el.scrollTop = Math.max(0, rowTopInTbody - Math.floor(viewportDataHeight * 0.45));
      } else {
        const estimatedY = (targetRow - 1) * DEFAULT_DATA_ROW_HEIGHT;
        el.scrollTop = Math.max(0, estimatedY - Math.floor(viewportDataHeight * 0.45));
      }
      el.scrollLeft = 0;
    });

    return () => window.cancelAnimationFrame(rafId);
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

  // Batch-safe cell updater preserving optimistic concurrency baselines
  const applyCellUpdates = useCallback(
    (updates: Array<{ row: number; col: number; value: string | number }>) => {
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

      for (const { row: r, col: c, value } of updates) {
        if (r < 0 || c < 0 || c >= MAX_COLS_LIMIT) continue;

        const prevValue =
          r === 0
            ? (updatedHeaders[c] ?? '')
            : (updatedRows[r - 1]?.[c] ?? '');

        if (prevValue === value) continue;

        if (r === 0) {
          headersChanged = true;
          while (updatedHeaders.length <= c) updatedHeaders.push('');
          updatedHeaders[c] = String(value);

          const headerKey = `${sheetId}:${c}`;
          const existingDirtyHeader = nextDirtyHeaders[headerKey];
          const expectedHeaderVersion =
            existingDirtyHeader?.expectedHeaderVersion ?? targetSheet.headerVersion ?? 1;

          nextDirtyHeaders[headerKey] = {
            sheetId,
            sheetName: activeSheetName,
            colIndex: c,
            newHeader: String(value),
            expectedHeaderVersion
          };
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
            expectedCellVersion: baselineVersion
          };
        }
      }

      if (!headersChanged && !rowsChanged) {
        return;
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
      setSaveStatus('idle');
    },
    [activeSheetName, sheets, totalCols]
  );

  const setCellValue = useCallback(
    (r: number, c: number, value: string | number) => {
      applyCellUpdates([{ row: r, col: c, value }]);
    },
    [applyCellUpdates]
  );

  useEffect(() => {
    if (isEditing && editInputRef.current) {
      editInputRef.current.focus();
    }
  }, [isEditing]);

  // Commit current edit, preserving original scalar type if text unchanged
  const commitEdit = useCallback(
    (moveToNextRow: boolean = false, moveToNextCol: boolean = false) => {
      if (!isEditing) return;
      const prevVal = getCellValue(selectedCell.row, selectedCell.col);
      if (String(prevVal ?? '') !== editValue) {
        setCellValue(selectedCell.row, selectedCell.col, editValue);
      }
      setIsEditing(false);

      if (moveToNextRow) {
        const nextRow = Math.min(displayRowsCount - 1, selectedCell.row + 1);
        setSelection({
          mode: 'cell',
          anchorRow: nextRow,
          anchorCol: selectedCell.col,
          focusRow: nextRow,
          focusCol: selectedCell.col
        });
        ensureRowVisible(nextRow);
      } else if (moveToNextCol) {
        const nextCol = Math.min(totalCols - 1, selectedCell.col + 1);
        setSelection({
          mode: 'cell',
          anchorRow: selectedCell.row,
          anchorCol: nextCol,
          focusRow: selectedCell.row,
          focusCol: nextCol
        });
      }
    },
    [
      isEditing,
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

  // Latest-state refs so memoized row callbacks remain strictly stable across renders
  const isEditingRef = useRef(isEditing);
  const selectedCellRef = useRef(selectedCell);
  const selectionModeRef = useRef(selection.mode);
  const totalColsRef = useRef(totalCols);
  const displayRowsCountRef = useRef(displayRowsCount);
  const activeDimKeyRef = useRef(activeDimKey);
  const totalTableWidthRef = useRef(totalTableWidth);
  const commitEditRef = useRef(commitEdit);
  const getCellValueRef = useRef(getCellValue);
  const getColWidthRef = useRef(getColWidth);
  const getManualRowHeightRef = useRef(getManualRowHeight);

  useEffect(() => {
    isEditingRef.current = isEditing;
    selectedCellRef.current = selectedCell;
    selectionModeRef.current = selection.mode;
    totalColsRef.current = totalCols;
    displayRowsCountRef.current = displayRowsCount;
    activeDimKeyRef.current = activeDimKey;
    totalTableWidthRef.current = totalTableWidth;
    commitEditRef.current = commitEdit;
    getCellValueRef.current = getCellValue;
    getColWidthRef.current = getColWidth;
    getManualRowHeightRef.current = getManualRowHeight;
  });

  // Clear selected cells with clinical safeguard confirmation for bulk/row/col/all deletions
  const handleDeleteSelection = () => {
    if (selection.mode === 'cell') {
      const cur = getCellValue(selectedCell.row, selectedCell.col);
      if (isNonEmptyCell(cur)) {
        setCellValue(selectedCell.row, selectedCell.col, '');
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

    applyCellUpdates(updates);
  };

  // Keyboard navigation, selection extension & direct typing (Ctrl+S handled by active-workspace listener)
  const handleKeyDown = (e: React.KeyboardEvent) => {
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') {
      e.preventDefault();
      return;
    }

    // Ctrl+A / Cmd+A selects entire worksheet when not editing
    if (!isEditing && (e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'a') {
      e.preventDefault();
      setSelection({
        mode: 'all',
        anchorRow: 0,
        anchorCol: 0,
        focusRow: displayRowsCount - 1,
        focusCol: totalCols - 1
      });
      return;
    }

    if (isEditing) {
      if (e.key === 'Enter') {
        e.preventDefault();
        commitEdit(true, false);
      } else if (e.key === 'Tab') {
        e.preventDefault();
        commitEdit(false, true);
      } else if (e.key === 'Escape') {
        e.preventDefault();
        cancelEdit();
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
    } else if (e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) {
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

  // Single click selects only — NEVER opens edit mode on hover or single-click
  const handleCellClick = useCallback((r: number, c: number, shiftKey: boolean) => {
    if (justResizedRef.current) return;
    rootRef.current?.focus({ preventScroll: true });

    if (shiftKey) {
      if (isEditingRef.current) commitEditRef.current(false, false);
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
      commitEditRef.current(false, false);
    }

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
    if (isEditingRef.current) commitEditRef.current(false, false);
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
    if (justResizedRef.current) return;
    rootRef.current?.focus({ preventScroll: true });
    if (isEditingRef.current) commitEditRef.current(false, false);

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
    if (justResizedRef.current) return;
    rootRef.current?.focus({ preventScroll: true });
    if (isEditingRef.current) commitEditRef.current(false, false);

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
    if (isEditingRef.current) commitEditRef.current(false, false);
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
    commitEditRef.current(false, false);
  }, []);

  // Copy selected cell, range, row(s), column(s), or entire worksheet in Excel TSV format
  const handleCopy = (e: React.ClipboardEvent) => {
    if (isEditing) return;

    const activeEl = document.activeElement;
    if (activeEl && activeEl.tagName === 'INPUT') {
      // Allow normal text copy inside formula bar or modal inputs
      return;
    }

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

    const lines: string[] = [];
    for (let r = rStart; r <= rEnd; r++) {
      const rowCells: string[] = [];
      for (let c = cStart; c <= cEnd; c++) {
        const val = getCellValue(r, c);
        rowCells.push(val !== undefined && val !== null ? String(val) : '');
      }
      lines.push(rowCells.join('\t'));
    }

    e.clipboardData.setData('text/plain', lines.join('\n'));
    e.preventDefault();
  };

  // Paste compatible TSV/scalar values starting at selection anchor without triggering unintended saves
  const handlePaste = (e: React.ClipboardEvent) => {
    if (isEditing) return;

    const activeEl = document.activeElement;
    if (activeEl && activeEl.tagName === 'INPUT') {
      return;
    }

    const text = e.clipboardData.getData('text/plain');
    if (!text) return;

    e.preventDefault();
    const matrix = parseClipboardTsv(text);
    if (matrix.length === 0) return;

    const updates: Array<{ row: number; col: number; value: string | number }> = [];

    if (matrix.length === 1 && matrix[0].length === 1) {
      const singleVal = matrix[0][0];
      if (selection.mode === 'range') {
        for (let r = selectionBounds.rMin; r <= selectionBounds.rMax; r++) {
          for (let c = selectionBounds.cMin; c <= selectionBounds.cMax; c++) {
            updates.push({ row: r, col: c, value: singleVal });
          }
        }
      } else {
        // For cell, row, col, or all: paste single value only at anchor cell to avoid accidental bulk overwrite
        updates.push({ row: selectedCell.row, col: selectedCell.col, value: singleVal });
      }
    } else {
      const startRow = selectionBounds.rMin;
      const startCol = selectionBounds.cMin;
      for (let dr = 0; dr < matrix.length; dr++) {
        const targetRow = startRow + dr;
        if (targetRow >= 100000) break;
        const cols = matrix[dr];
        for (let dc = 0; dc < cols.length; dc++) {
          const targetCol = startCol + dc;
          if (targetCol >= MAX_COLS_LIMIT) break;
          updates.push({ row: targetRow, col: targetCol, value: cols[dc] });
        }
      }
    }

    if (updates.length > 0) {
      applyCellUpdates(updates);
    }
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

  // Targeted RPC Save (Headers via patch_workspace_header, Cells via patch_workspace_cell / patch_workspace_cells_batch)
  const handleSave = async () => {
    if (isSavingRef.current) return;

    if (isEditing) {
      commitEdit();
    }

    const pendingHeadersMap = { ...dirtyHeadersRef.current };
    const pendingCellsMap = { ...dirtyCellsRef.current };

    const headerEdits = Object.entries(pendingHeadersMap);
    const cellEdits = Object.entries(pendingCellsMap);

    if (headerEdits.length === 0 && cellEdits.length === 0) {
      if (saveStatus === 'error') {
        setSaveStatus('idle');
        setSaveError(null);
      }
      return;
    }

    isSavingRef.current = true;
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
            throw new Error(
              typeof data.error === 'string' && data.error
                ? data.error
                : 'Failed to save header change.'
            );
          }

          runningHeaderVersion = Number(data.newHeaderVersion ?? runningHeaderVersion + 1);

          const updatedDirtyHeaders = { ...dirtyHeadersRef.current };
          delete updatedDirtyHeaders[key];
          dirtyHeadersRef.current = updatedDirtyHeaders;

          for (const remKey of Object.keys(dirtyHeadersRef.current)) {
            if (dirtyHeadersRef.current[remKey].sheetId === sheetId) {
              dirtyHeadersRef.current[remKey] = {
                ...dirtyHeadersRef.current[remKey],
                expectedHeaderVersion: runningHeaderVersion
              };
            }
          }
          setDirtyHeaders({ ...dirtyHeadersRef.current });

          setSheets((prev) => {
            const target = prev[edit.sheetName];
            if (!target) return prev;
            return {
              ...prev,
              [edit.sheetName]: {
                ...target,
                headerVersion: runningHeaderVersion,
                headers:
                  Array.isArray(data.headers) && data.headers.length > 0
                    ? (data.headers as string[])
                    : target.headers
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
              throw new Error(
                typeof data.error === 'string' && data.error
                  ? data.error
                  : 'Failed to save cell edit.'
              );
            }

            const newCellVer = Number(data.newCellVersion ?? edit.expectedCellVersion + 1);
            const newRowVer = data.newRowVersion !== undefined ? Number(data.newRowVersion) : undefined;
            const savedAt = typeof data.updatedAt === 'string' ? data.updatedAt : new Date().toISOString();

            const updatedDirtyCells = { ...dirtyCellsRef.current };
            delete updatedDirtyCells[key];
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
              const nextRowUpdatedAt = { ...(target.rowUpdatedAt || {}), [edit.rowIndex]: savedAt };
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
              throw new Error(
                typeof data.error === 'string' && data.error
                  ? data.error
                  : 'Failed to save cell edits batch.'
              );
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
            const savedAt = typeof data.updatedAt === 'string' ? data.updatedAt : new Date().toISOString();

            const updatedDirtyCells = { ...dirtyCellsRef.current };
            for (const { key } of chunk) {
              delete updatedDirtyCells[key];
            }
            dirtyCellsRef.current = updatedDirtyCells;
            setDirtyCells(updatedDirtyCells);

            setSheets((prev) => {
              const target = prev[sheetName];
              if (!target) return prev;
              const nextCellVersions = { ...(target.cellVersions || {}) };
              const nextRowVersions = { ...(target.rowVersions || {}) };
              const nextRowUpdatedAt = { ...(target.rowUpdatedAt || {}) };

              updatedCellsRes.forEach((u, idx) => {
                const fallbackEdit = chunk[idx]?.edit;
                const rIdx = u.rowIndex ?? fallbackEdit?.rowIndex;
                const cIdx = u.colIndex ?? fallbackEdit?.colIndex;
                if (rIdx === undefined || cIdx === undefined) return;
                nextCellVersions[rIdx] = {
                  ...(nextCellVersions[rIdx] || {}),
                  [cIdx]: Number(u.newCellVersion ?? (fallbackEdit ? fallbackEdit.expectedCellVersion + 1 : 1))
                };
                if (u.newRowVersion !== undefined) {
                  nextRowVersions[rIdx] = Number(u.newRowVersion);
                }
                nextRowUpdatedAt[rIdx] = savedAt;
              });

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

      const remainingUnsaved =
        Object.keys(dirtyCellsRef.current).length > 0 ||
        Object.keys(dirtyHeadersRef.current).length > 0;

      if (remainingUnsaved) {
        setSaveStatus('idle');
        setSaveMessage('');
      } else {
        setSaveStatus('saved');
        setSaveMessage('Saved changes');
      }
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : 'Save error occurred';
      setSaveStatus('error');
      setSaveMessage('');
      setSaveError(msg);
    } finally {
      isSavingRef.current = false;
    }
  };

  const handleSaveRef = useRef(handleSave);
  useEffect(() => {
    handleSaveRef.current = handleSave;
  });

  useEffect(() => {
    if (!isActiveWorkspace) return;
    const onWindowKeyDown = (e: KeyboardEvent) => {
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') {
        e.preventDefault();
        void handleSaveRef.current();
      }
    };
    window.addEventListener('keydown', onWindowKeyDown);
    return () => window.removeEventListener('keydown', onWindowKeyDown);
  }, [isActiveWorkspace]);

  const handleReloadLatest = async () => {
    if (!onReloadWorkspace) return;
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

  // Direct single-click Export button: saves pending changes first if any, then exports workspace
  const handleDirectExport = async () => {
    if (isExporting || sheetKeys.length === 0) return;
    setIsExporting(true);
    try {
      if (hasUnsavedChanges) {
        await handleSave();
        if (Object.keys(dirtyCellsRef.current).length > 0 || Object.keys(dirtyHeadersRef.current).length > 0) {
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
      alert(`Export failed: ${msg}`);
    } finally {
      setIsExporting(false);
    }
  };

  // Close worksheet context menu on click outside or Escape
  useEffect(() => {
    if (!tabContextMenu) return;
    const handleGlobalClick = () => setTabContextMenu(null);
    const handleGlobalKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'Escape') setTabContextMenu(null);
    };
    window.addEventListener('click', handleGlobalClick);
    window.addEventListener('keydown', handleGlobalKeyDown);
    return () => {
      window.removeEventListener('click', handleGlobalClick);
      window.removeEventListener('keydown', handleGlobalKeyDown);
    };
  }, [tabContextMenu]);

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
    const newName = getNextAvailableSheetName(sheetKeys);
    setIsCreatingSheet(true);

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

      const data = await res.json();
      if (!res.ok) {
        throw new Error(data.error || 'Failed to create worksheet on server.');
      }

      const createdSheet: SheetState = data.sheet || {
        name: newName,
        headers: defaultHeaders,
        headerVersion: 1,
        rows: [],
        rowVersions: {},
        cellVersions: {},
        rowUpdatedAt: {}
      };

      setSheets((prev) => ({
        ...prev,
        [newName]: createdSheet
      }));

      setActiveSheetName(newName);
      setSelection({
        mode: 'cell',
        anchorRow: 1,
        anchorCol: 0,
        focusRow: 1,
        focusCol: 0
      });
      setScrollPositionRequest((prev) => ({
        sheetName: newName,
        targetRow: 1,
        nonce: prev.nonce + 1
      }));
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : 'Failed to create worksheet.';
      alert(msg);
    } finally {
      setIsCreatingSheet(false);
    }
  };

  // 7. Worksheet Renaming via Double-Click or Context Menu
  const handleStartRename = (targetName: string) => {
    setRenamingSheetName(targetName);
    setRenameInputValue(targetName);
    setTabContextMenu(null);
    setTimeout(() => {
      renameInputRef.current?.focus();
      renameInputRef.current?.select();
    }, 20);
  };

  const handleCommitRename = async (oldName: string, rawNewName: string) => {
    const trimmed = rawNewName.trim();
    setRenamingSheetName(null);

    if (!trimmed || trimmed === oldName) {
      return;
    }

    const validation = validateWorksheetName(trimmed);
    if (!validation.valid) {
      alert(validation.error);
      return;
    }

    const otherNames = sheetKeys.filter((n) => n !== oldName);
    if (otherNames.some((n) => n.toLowerCase() === trimmed.toLowerCase())) {
      alert(`A worksheet named "${trimmed}" already exists.`);
      return;
    }

    const targetSheet = sheets[oldName];
    if (!targetSheet?.id) {
      alert('Cannot rename worksheet before it is saved.');
      return;
    }

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

      const data = await res.json();
      if (!res.ok) {
        throw new Error(data.error || 'Failed to rename worksheet.');
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

      if (activeSheetName === oldName) {
        setActiveSheetName(trimmed);
      }
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : 'Failed to rename worksheet.';
      alert(msg);
    } finally {
      setIsRenamingSheet(false);
    }
  };

  // 8. Worksheet Context Menu - Delete
  const handleDeleteSheet = async (targetSheetName: string) => {
    if (sheetKeys.length <= 1) {
      alert(`Cannot delete the only remaining worksheet in the ${workspace} workspace.`);
      setPendingDeleteSheet(null);
      return;
    }

    const targetSheet = sheets[targetSheetName];
    if (!targetSheet?.id) {
      alert('Worksheet cannot be deleted.');
      setPendingDeleteSheet(null);
      return;
    }

    setIsDeletingSheet(true);
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

      const data = await res.json();
      if (!res.ok) {
        throw new Error(data.error || 'Failed to delete worksheet.');
      }

      const remainingKeys = sheetKeys.filter((k) => k !== targetSheetName);
      setSheets((prev) => {
        const next = { ...prev };
        delete next[targetSheetName];
        return next;
      });

      if (activeSheetName === targetSheetName && remainingKeys.length > 0) {
        const newActive = remainingKeys[0];
        setActiveSheetName(newActive);
      }
      setPendingDeleteSheet(null);
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : 'Failed to delete worksheet.';
      alert(msg);
    } finally {
      setIsDeletingSheet(false);
    }
  };

  // 9. Worksheet Reordering via Drag and Drop
  const handleReorderSheets = async (sourceName: string, targetName: string) => {
    if (sourceName === targetName) return;
    const currentKeys = [...sheetKeys];
    const sourceIdx = currentKeys.indexOf(sourceName);
    const targetIdx = currentKeys.indexOf(targetName);

    if (sourceIdx === -1 || targetIdx === -1) return;

    currentKeys.splice(sourceIdx, 1);
    currentKeys.splice(targetIdx, 0, sourceName);

    // Immediately update local state tabOrder
    setSheets((prev) => {
      const next = { ...prev };
      currentKeys.forEach((key, idx) => {
        if (next[key]) {
          next[key] = { ...next[key], tabOrder: idx };
        }
      });
      return next;
    });

    const orderedSheetIds = currentKeys
      .map((k) => sheets[k]?.id)
      .filter((id): id is string => typeof id === 'string' && id.length > 0);

    if (orderedSheetIds.length === 0) return;

    try {
      await fetch('/api/records', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          workspace,
          action: 'reorder_sheets',
          sheetIds: orderedSheetIds
        })
      });
    } catch (err) {
      console.error('Failed to persist sheet reordering in Supabase:', err);
    }
  };

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

  const sheetKeys = Object.keys(sheets);
  const visibleWorkspaces = allowedWorkspaces && allowedWorkspaces.length > 0 ? allowedWorkspaces : [workspace];
  const staffLabel = user?.name || user?.username || workspace;

  const isFullRowSelectionMode = selection.mode === 'row' || selection.mode === 'all';
  const isFullColSelectionMode = selection.mode === 'col' || selection.mode === 'all';

  return (
    <div
      ref={rootRef}
      className="flex flex-col h-screen w-full bg-white select-none overflow-hidden font-sans text-xs text-slate-800 outline-none"
      onKeyDown={handleKeyDown}
      onCopy={handleCopy}
      onPaste={handlePaste}
      tabIndex={0}
    >
      {/* 1. Application Toolbar / Header */}
      <div className="h-12 px-4 bg-slate-900 text-white border-b border-slate-800 flex items-center justify-between gap-4 shrink-0 shadow-xs">
        {/* Top-Left: Official MK Eye Center Logo directly, NO background container */}
        <div className="flex items-center shrink-0">
          {/* eslint-disable-next-line @next/next/no-img-element */}
          <img
            src="/assets/aistudio/mk_long_logo_a5863dd4.png"
            alt="MK Eye Center Logo"
            className="h-7 sm:h-8 w-auto object-contain"
            referrerPolicy="no-referrer"
          />
        </div>

        {/* Top-Right: Save and Export neatly aligned horizontally */}
        <div className="flex items-center gap-2.5 ml-auto">
          {/* Admin workspace switcher if applicable */}
          {visibleWorkspaces.length > 1 && (
            <div className="flex items-center bg-slate-800 p-0.5 rounded-lg border border-slate-700 text-xs mr-2">
              <button
                type="button"
                onClick={() => onWorkspaceSwitchRequest('reception')}
                className={`px-3 py-1 font-semibold text-xs rounded-md transition cursor-pointer ${
                  activeWorkspace === 'reception'
                    ? 'bg-blue-600 text-white shadow-xs'
                    : 'text-slate-300 hover:text-white'
                }`}
              >
                Reception
              </button>
              <button
                type="button"
                onClick={() => onWorkspaceSwitchRequest('doctor')}
                className={`px-3 py-1 font-semibold text-xs rounded-md transition cursor-pointer ${
                  activeWorkspace === 'doctor'
                    ? 'bg-blue-600 text-white shadow-xs'
                    : 'text-slate-300 hover:text-white'
                }`}
              >
                Doctor
              </button>
            </div>
          )}

          {/* Save & Conflict Status Notice */}
          <div className="flex items-center gap-2 text-[11px] mr-1" role="status" aria-live="polite">
            {isLoading && <span className="text-slate-400">Loading...</span>}
            {!isLoading && saveStatus === 'saving' && (
              <span className="text-slate-300 font-medium">Saving...</span>
            )}
            {!isLoading && saveStatus === 'saved' && !hasUnsavedChanges && (
              <span className="text-emerald-400 font-semibold">{saveMessage || 'Saved'}</span>
            )}
            {!isLoading && hasUnsavedChanges && saveStatus !== 'saving' && (
              <span className="text-amber-400 font-medium">Unsaved</span>
            )}
            {!isLoading && saveStatus === 'error' && (
              <span title={saveError || 'Save failed'} className="text-red-400 font-medium max-w-[150px] truncate">
                {saveError || 'Save failed'}
              </span>
            )}
            {!isLoading && hasConflict && onReloadWorkspace && (
              <button
                type="button"
                onClick={() => void handleReloadLatest()}
                disabled={isReloading}
                className="px-2 py-0.5 bg-red-600 text-white rounded-md font-medium hover:bg-red-700 cursor-pointer disabled:opacity-50 text-[11px]"
              >
                {isReloading ? 'Reload' : 'Reload'}
              </button>
            )}
          </div>

          {user?.role === 'admin' && canAdminImport && onRequestInitialImport && (
            <button
              type="button"
              onClick={onRequestInitialImport}
              className="px-3 py-1.5 font-semibold text-white bg-emerald-700 hover:bg-emerald-800 rounded-lg transition cursor-pointer text-xs"
            >
              Import Workbooks
            </button>
          )}

          <button
            type="button"
            onClick={() => void handleSave()}
            disabled={saveStatus === 'saving' || isLoading || !currentSheet.id}
            className="px-4 py-1.5 font-semibold text-white bg-blue-600 hover:bg-blue-500 active:bg-blue-700 rounded-lg transition disabled:opacity-50 cursor-pointer flex items-center gap-1.5 text-xs shadow-xs"
          >
            {saveStatus === 'saving' ? 'Saving...' : 'Save'}
          </button>

          <button
            type="button"
            onClick={() => void handleDirectExport()}
            disabled={isExporting || sheetKeys.length === 0}
            className="px-4 py-1.5 font-semibold text-slate-200 bg-slate-800 hover:bg-slate-700 border border-slate-700 rounded-lg transition disabled:opacity-50 cursor-pointer text-xs"
          >
            {isExporting ? 'Exporting...' : 'Export'}
          </button>
        </div>
      </div>

      {/* 2. Formula & Cell Address Bar */}
      <div className="h-8 md:h-7 px-2.5 md:px-3 bg-white border-b border-slate-300 flex items-center gap-2 shrink-0 text-xs">
        <div className="min-w-16 h-5 px-1.5 bg-slate-50 border border-slate-300 rounded-xs flex items-center justify-center font-mono font-semibold text-slate-700 text-[11px] select-none shrink-0">
          {selectionAddressLabel}
        </div>
        <span className="text-slate-400 font-serif italic text-xs select-none">fx</span>
        <div className="h-4 w-px bg-slate-300 shrink-0" />
        <input
          type="text"
          value={isEditing ? editValue : String(getCellValue(selectedCell.row, selectedCell.col) ?? '')}
          disabled={!currentSheet.id}
          onChange={(e) => {
            if (!isEditing) {
              setIsEditing(true);
            }
            setEditValue(e.target.value);
          }}
          onKeyDown={(e) => {
            if (e.key === 'Enter') {
              e.preventDefault();
              commitEdit(true, false);
            } else if (e.key === 'Escape') {
              e.preventDefault();
              cancelEdit();
            }
          }}
          onBlur={() => {
            if (isEditing) {
              commitEdit(false, false);
            }
          }}
          className="flex-1 min-w-0 h-6 md:h-5 px-1.5 bg-transparent text-slate-900 text-xs font-normal outline-none focus:bg-slate-50 rounded-xs disabled:opacity-50"
        />
      </div>

      {/* 2.1 Load Error Strip (if Supabase load fails) */}
      {loadError && (
        <div className="bg-red-50 border-b border-red-200 px-3 py-1.5 flex items-center justify-between text-xs text-red-800 shrink-0">
          <span>{loadError}</span>
          {onReloadWorkspace && (
            <button
              type="button"
              onClick={() => void handleReloadLatest()}
              disabled={isReloading}
              className="px-2 py-0.5 bg-red-600 text-white rounded-xs font-medium hover:bg-red-700 cursor-pointer"
            >
              Retry
            </button>
          )}
        </div>
      )}

      {/* 2.2 Admin Import Result / Error Strip */}
      {user?.role === 'admin' && importNotice && (
        <div
          className={`border-b px-3 py-1.5 flex items-center justify-between text-xs shrink-0 ${
            importNotice.type === 'error'
              ? 'bg-red-50 border-red-200 text-red-800'
              : 'bg-emerald-50 border-emerald-200 text-emerald-900'
          }`}
        >
          <span>{importNotice.message}</span>
          {onDismissImportNotice && (
            <button
              type="button"
              onClick={onDismissImportNotice}
              className="px-2 py-0.5 text-[11px] font-medium underline cursor-pointer"
            >
              Dismiss
            </button>
          )}
        </div>
      )}

      {/* 3. Main Spreadsheet Grid Viewport */}
      <div
        ref={gridContainerRef}
        className="flex-1 overflow-auto touch-pan-x touch-pan-y bg-white relative outline-none"
      >
        {!isLoading && sheetKeys.length === 0 ? (
          <div className="h-full flex flex-col items-center justify-center text-slate-500 p-6 text-center">
            <p className="text-sm font-medium text-slate-700">
              No worksheets found in the {workspace === 'doctor' ? 'Doctor' : 'Reception'} workspace yet.
            </p>
            <p className="text-xs text-slate-500 mt-1">
              Click the <span className="font-mono font-bold">+</span> button at the bottom left to create a worksheet.
            </p>
            {user?.role === 'admin' && canAdminImport && onRequestInitialImport && (
              <button
                type="button"
                onClick={onRequestInitialImport}
                className="mt-4 px-4 py-1.5 text-xs font-semibold text-white bg-slate-900 hover:bg-slate-800 rounded-xs transition cursor-pointer shadow-xs"
              >
                Import Initial Workbooks (Reception &amp; Doctor)
              </button>
            )}
          </div>
        ) : (
          <table
            ref={tableRef}
            style={{ width: `${totalTableWidth}px` }}
            className="border-separate border-spacing-0 table-fixed text-xs"
          >
            <colgroup>
              <col style={{ width: `${ROW_HEADER_COL_WIDTH}px` }} />
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

            <thead ref={theadRef}>
              {/* Row 0: Excel Column Letters (sticky top-0, height 24px / h-6) */}
              <tr className="h-6">
                <th
                  onClick={handleSelectAllClick}
                  title="Select entire worksheet"
                  className={`w-12 min-w-12 h-6 border-r border-b border-slate-300 sticky top-0 left-0 z-50 select-none text-center cursor-pointer transition-colors ${
                    selection.mode === 'all'
                      ? 'bg-slate-700 text-white'
                      : 'bg-slate-200 hover:bg-slate-300 text-slate-500'
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
                      className={`h-6 border-r border-b border-slate-300 text-center text-[11px] select-none sticky top-0 z-30 cursor-pointer ${
                        isColSelected && isFullColSelectionMode
                          ? 'bg-slate-700 text-white font-semibold'
                          : isColSelected
                            ? 'bg-slate-200 text-slate-900 font-semibold'
                            : 'bg-slate-100 hover:bg-slate-200/70 font-medium text-slate-700'
                      }`}
                    >
                      <span>{getColumnLetter(cIdx)}</span>
                      <div
                        onPointerDown={(e) => handleColResizeStart(cIdx, e)}
                        onClick={(e) => e.stopPropagation()}
                        onDoubleClick={(e) => e.stopPropagation()}
                        title="Drag to resize column width"
                        className="absolute right-0 top-0 bottom-0 w-2 md:w-1.5 cursor-col-resize touch-none hover:bg-sky-500/70 z-20"
                      />
                    </th>
                  );
                })}
              </tr>

              {/* Row 1: Sticky Worksheet Column Header Titles (sticky top-6 / 24px) */}
              <tr
                ref={(el) => {
                  rowElementsRef.current[0] = el;
                }}
                style={getManualRowHeight(0) ? { height: `${getManualRowHeight(0)}px` } : undefined}
              >
                <th
                  ref={headerRowThRef}
                  onClick={(e) => handleRowHeaderClick(0, e.shiftKey)}
                  style={getManualRowHeight(0) ? { height: `${getManualRowHeight(0)}px` } : undefined}
                  className={`w-12 min-w-12 min-h-[34px] border-r border-b-2 border-slate-300 border-b-slate-400 font-bold text-center text-[11px] sticky top-6 left-0 z-40 select-none cursor-pointer shadow-[0_2px_4px_-1px_rgba(15,23,42,0.12)] ${
                    isRowHeaderHighlighted(0) && isFullRowSelectionMode
                      ? 'bg-slate-700 text-white'
                      : isRowHeaderHighlighted(0)
                        ? 'bg-[#c5d4ea] text-slate-900'
                        : 'bg-[#d5e0ef] hover:bg-[#c5d4ea] text-slate-800'
                  }`}
                >
                  <span>1</span>
                  <div
                    onPointerDown={(e) => handleRowResizeStart(0, e)}
                    onClick={(e) => e.stopPropagation()}
                    onDoubleClick={(e) => e.stopPropagation()}
                    title="Drag to resize header row height"
                    className="absolute left-0 right-0 bottom-0 h-2 md:h-1.5 cursor-row-resize touch-none hover:bg-sky-500/70 z-20"
                  />
                </th>

                {Array.from({ length: totalCols }).map((_, cIdx) => {
                  const isAnchor = selectedCell.row === 0 && selectedCell.col === cIdx;
                  const isInSel = isCellInSelection(0, cIdx);
                  const isCellEditing = isAnchor && isEditing && selection.mode === 'cell';
                  const cellVal = getCellValue(0, cIdx);
                  const width = getColWidth(cIdx);
                  const manualHeight = getManualRowHeight(0);
                  const colMeta = colMetaList[cIdx];
                  const category = colMeta?.category ?? 'normal';
                  const alignClass = getAlignmentClass(colMeta?.align ?? 'left');
                  const headerBgClass = getClinicalHeaderClasses(
                    category,
                    isInSel && !(isAnchor && selection.mode === 'cell'),
                    isDoctorWorkspace
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
                      onClick={(e) => handleCellClick(0, cIdx, e.shiftKey)}
                      onDoubleClick={(e) => handleCellDoubleClick(0, cIdx, e.detail)}
                      className={`min-h-[34px] px-2.5 py-1.5 border-r border-b-2 border-slate-300 border-b-slate-400 font-bold text-xs leading-[18px] tracking-[0.01em] sticky top-6 shadow-[0_2px_4px_-1px_rgba(15,23,42,0.1)] align-middle ${alignClass} ${headerBgClass} ${
                        isAnchor || isCellEditing ? 'z-[25]' : 'z-20'
                      } ${
                        isAnchor && !isCellEditing
                          ? 'outline-2 outline-slate-900 outline-offset-[-2px]'
                          : ''
                      }`}
                    >
                      {isCellEditing ? (
                        <input
                          ref={setEditInputRef}
                          type="text"
                          value={editValue}
                          onChange={(e) => setEditValue(e.target.value)}
                          onBlur={() => commitEdit(false, false)}
                          className={`absolute inset-0 w-full h-full px-2.5 text-xs font-semibold bg-white text-slate-900 border-2 border-slate-900 outline-none z-20 ${alignClass}`}
                        />
                      ) : (
                        <div
                          className={
                            isWrapEnabled
                              ? 'whitespace-normal break-words leading-[18px]'
                              : 'truncate whitespace-nowrap leading-[18px]'
                          }
                          style={
                            manualHeight
                              ? { maxHeight: `${Math.max(18, manualHeight - 8)}px`, overflowY: 'auto' }
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
                    colMetaList={colMetaList}
                    isDoctorWorkspace={isDoctorWorkspace}
                    manualHeight={manualHeight}
                    isWrapEnabled={isWrapEnabled}
                    isRowHighlighted={isRowHighlighted}
                    isFullRowMode={isFullRowSelectionMode}
                    anchorCol={anchorColForRow}
                    isEditingRow={isEditingRow}
                    editValueIfEditing={isEditingRow ? editValue : ''}
                    selectedColStart={isRowHighlighted ? selectionBounds.cMin : -1}
                    selectedColEnd={isRowHighlighted ? selectionBounds.cMax : -1}
                    onRowHeaderClick={handleRowHeaderClick}
                    onRowResizeStart={handleRowResizeStart}
                    onCellClick={handleCellClick}
                    onCellDoubleClick={handleCellDoubleClick}
                    onEditValueChange={setEditValue}
                    onCommitEdit={handleCommitEditBlur}
                    setEditInputRef={setEditInputRef}
                    setRowElementRef={setRowElementRef}
                  />
                );
              })}
            </tbody>
          </table>
        )}
      </div>

      {/* 4. Worksheet Tabs along the bottom with '+' button and Sign Out at bottom-right */}
      <div className="h-9 md:h-8 bg-slate-200 border-t border-slate-300 flex items-center justify-between px-2 gap-2 shrink-0">
        <div className="flex items-center gap-1 overflow-x-auto min-w-0 flex-1 py-0.5">
          {sheetKeys.map((name) => {
            const isActive = activeSheetName === name;
            const isBeingRenamed = renamingSheetName === name;
            const isDragOver = dragOverSheet === name;

            return (
              <div
                key={name}
                draggable={!isBeingRenamed}
                onDragStart={(e) => {
                  setDraggedSheet(name);
                  e.dataTransfer.effectAllowed = 'move';
                  e.dataTransfer.setData('text/plain', name);
                }}
                onDragOver={(e) => {
                  e.preventDefault();
                  e.dataTransfer.dropEffect = 'move';
                  if (draggedSheet && draggedSheet !== name) {
                    setDragOverSheet(name);
                  }
                }}
                onDragLeave={() => {
                  if (dragOverSheet === name) setDragOverSheet(null);
                }}
                onDrop={(e) => {
                  e.preventDefault();
                  const source = draggedSheet || e.dataTransfer.getData('text/plain');
                  setDraggedSheet(null);
                  setDragOverSheet(null);
                  if (source && source !== name) {
                    void handleReorderSheets(source, name);
                  }
                }}
                onDragEnd={() => {
                  setDraggedSheet(null);
                  setDragOverSheet(null);
                }}
                onContextMenu={(e) => {
                  e.preventDefault();
                  e.stopPropagation();
                  const menuWidth = 140;
                  const menuHeight = 80;
                  const x = Math.min(e.clientX, window.innerWidth - menuWidth - 12);
                  const y = Math.max(12, e.clientY - menuHeight - 12);
                  setTabContextMenu({ sheetName: name, x, y });
                }}
                onClick={() => {
                  if (isEditing) commitEdit();
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
                  handleStartRename(name);
                }}
                className={`group relative px-3 py-1 text-xs font-medium border-t-2 rounded-t-xs transition cursor-pointer select-none whitespace-nowrap flex items-center gap-1.5 ${
                  isActive
                    ? 'bg-white border-slate-900 text-slate-900 shadow-xs'
                    : 'bg-slate-200 border-transparent text-slate-600 hover:bg-slate-100 hover:text-slate-900'
                } ${isDragOver ? 'border-l-2 border-l-blue-600 bg-blue-50' : ''}`}
                title={`Worksheet: ${name} (Double-click to rename, right-click for options, drag to reorder)`}
              >
                {isBeingRenamed ? (
                  <input
                    ref={renameInputRef}
                    type="text"
                    value={renameInputValue}
                    onChange={(e) => setRenameInputValue(e.target.value)}
                    onKeyDown={(e) => {
                      if (e.key === 'Enter') {
                        e.preventDefault();
                        e.stopPropagation();
                        void handleCommitRename(name, renameInputValue);
                      } else if (e.key === 'Escape') {
                        e.preventDefault();
                        e.stopPropagation();
                        setRenamingSheetName(null);
                      }
                    }}
                    onBlur={() => {
                      void handleCommitRename(name, renameInputValue);
                    }}
                    onClick={(e) => e.stopPropagation()}
                    maxLength={31}
                    className="px-1 py-0.5 text-xs font-semibold bg-white border border-blue-600 rounded-xs text-slate-900 outline-none w-28"
                  />
                ) : (
                  <span>{name}</span>
                )}
              </div>
            );
          })}

          <button
            type="button"
            onClick={() => void handleAddSheetImmediate()}
            disabled={isCreatingSheet}
            title="Create new worksheet"
            className="h-6 w-6 flex items-center justify-center rounded-xs bg-slate-300/80 hover:bg-slate-300 active:bg-slate-400 text-slate-700 font-bold text-sm cursor-pointer ml-1 shrink-0 transition disabled:opacity-50"
          >
            {isCreatingSheet ? '...' : '+'}
          </button>
        </div>

        {/* Bottom-Right Corner: Sign Out */}
        <div className="flex items-center gap-2.5 shrink-0 pl-2.5 border-l border-slate-300">
          <span className="text-slate-600 font-mono text-[11px] max-w-[120px] sm:max-w-none truncate" title={user?.role}>
            {staffLabel}
          </span>
          <button
            type="button"
            onClick={onLogout}
            className="text-xs text-slate-700 hover:text-red-700 font-semibold px-2 py-0.5 rounded-xs hover:bg-slate-300/60 transition cursor-pointer flex items-center gap-1 whitespace-nowrap"
            title="Sign Out"
          >
            <LogOut className="w-3.5 h-3.5" />
            <span>Sign Out</span>
          </button>
        </div>
      </div>

      {/* 5. Worksheet Tab Context Menu (Right-Click) */}
      {tabContextMenu && (
        <div
          className="fixed z-50 bg-white border border-slate-300 rounded-sm shadow-xl py-1 w-36 text-xs text-slate-800"
          style={{ left: tabContextMenu.x, top: tabContextMenu.y }}
          onClick={(e) => e.stopPropagation()}
        >
          <div className="px-3 py-1 font-semibold text-[10px] uppercase tracking-wider text-slate-400 border-b border-slate-100 truncate">
            {tabContextMenu.sheetName}
          </div>
          <button
            type="button"
            onClick={() => handleStartRename(tabContextMenu.sheetName)}
            className="w-full text-left px-3 py-1.5 hover:bg-slate-100 flex items-center gap-2 cursor-pointer transition"
          >
            <span>Rename</span>
          </button>
          <button
            type="button"
            onClick={() => {
              const target = tabContextMenu.sheetName;
              setTabContextMenu(null);
              if (sheetKeys.length <= 1) {
                alert(`Cannot delete the only remaining worksheet in the ${workspace} workspace.`);
                return;
              }
              setPendingDeleteSheet(target);
            }}
            disabled={sheetKeys.length <= 1}
            className="w-full text-left px-3 py-1.5 hover:bg-red-50 text-red-600 disabled:opacity-40 disabled:hover:bg-transparent flex items-center gap-2 cursor-pointer transition border-t border-slate-100"
          >
            <span>Delete</span>
          </button>
        </div>
      )}

      {/* Worksheet Delete Confirmation Dialog */}
      <ConfirmDialog
        isOpen={pendingDeleteSheet !== null}
        title="Delete Worksheet"
        message={
          pendingDeleteSheet
            ? `Are you sure you want to delete worksheet "${pendingDeleteSheet}" and all of its records? This action cannot be undone.`
            : ''
        }
        confirmLabel={isDeletingSheet ? 'Deleting...' : 'Delete Worksheet'}
        confirmVariant="danger"
        onConfirm={() => {
          if (pendingDeleteSheet) {
            void handleDeleteSheet(pendingDeleteSheet);
          }
        }}
        onCancel={() => setPendingDeleteSheet(null)}
      />

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
            applyCellUpdates(pendingBulkClear.updates);
            setPendingBulkClear(null);
          }
        }}
        onCancel={() => setPendingBulkClear(null)}
      />


    </div>
  );
}
