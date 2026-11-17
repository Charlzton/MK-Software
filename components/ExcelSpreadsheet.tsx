'use client';

import React, { useState, useEffect, useRef, useCallback, useMemo } from 'react';
import * as XLSX from 'xlsx';
import { User } from '@/lib/types';

export interface SheetState {
  id?: string;
  name?: string;
  tabOrder?: number;
  headers: string[];
  headerVersion?: number;
  rows: (string | number)[][];
  rowVersions?: Record<number, number>;
  cellVersions?: Record<number, Record<number, number>>;
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

function getColumnLetter(colIndex: number): string {
  let temp = colIndex;
  let letter = '';
  while (temp >= 0) {
    letter = String.fromCharCode((temp % 26) + 65) + letter;
    temp = Math.floor(temp / 26) - 1;
  }
  return letter;
}

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
    if (!activeSheetName || !initialSheets[activeSheetName]) {
      const keys = Object.keys(initialSheets);
      if (defaultSheetName && initialSheets[defaultSheetName]) {
        setActiveSheetName(defaultSheetName);
      } else {
        setActiveSheetName(keys[0] || '');
      }
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

  // Notify parent of unsaved changes status
  useEffect(() => {
    onUnsavedStatusChange(hasUnsavedChanges);
  }, [hasUnsavedChanges, onUnsavedStatusChange]);

  // Active & Editing cell coordinates: row 0 = Header (Row 1), row >= 1 = Data rows (Row 2..N)
  const [selectedCell, setSelectedCell] = useState<{ row: number; col: number }>({ row: 1, col: 0 });
  const [isEditing, setIsEditing] = useState<boolean>(false);
  const [editValue, setEditValue] = useState<string>('');

  // Column widths state
  const [columnWidths, setColumnWidths] = useState<Record<number, number>>({
    0: 105,
    1: 120,
    2: 130,
    3: 130,
    4: 190,
    5: 110,
    6: 120,
    7: 95,
    8: 85,
    9: 100,
    10: 95,
    11: 100,
    12: 90,
    13: 100,
    14: 110,
    15: 110
  });

  const resizingColRef = useRef<{ col: number; startX: number; startWidth: number } | null>(null);
  const gridContainerRef = useRef<HTMLDivElement>(null);
  const editInputRef = useRef<HTMLInputElement>(null);

  // Add sheet modal state
  const [isAddSheetModalOpen, setIsAddSheetModalOpen] = useState(false);
  const [newSheetNameInput, setNewSheetNameInput] = useState('');
  const [addSheetError, setAddSheetError] = useState<string | null>(null);
  const [isCreatingSheet, setIsCreatingSheet] = useState(false);
  const [isReloading, setIsReloading] = useState(false);

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

  // Update cell value & record targeted dirty edit with baseline optimistic concurrency version
  const setCellValue = useCallback(
    (r: number, c: number, value: string | number) => {
      const targetSheet = sheets[activeSheetName];
      if (!targetSheet || !targetSheet.id) {
        return;
      }

      const prevValue =
        r === 0
          ? (targetSheet.headers[c] ?? '')
          : (targetSheet.rows[r - 1]?.[c] ?? '');

      if (prevValue === value) {
        return;
      }

      const sheetId = targetSheet.id;

      if (r === 0) {
        const headerKey = `${sheetId}:${c}`;
        const existingDirtyHeader = dirtyHeadersRef.current[headerKey];
        const expectedHeaderVersion =
          existingDirtyHeader?.expectedHeaderVersion ?? targetSheet.headerVersion ?? 1;

        const nextDirtyHeader: DirtyHeaderEdit = {
          sheetId,
          sheetName: activeSheetName,
          colIndex: c,
          newHeader: String(value),
          expectedHeaderVersion
        };
        dirtyHeadersRef.current = {
          ...dirtyHeadersRef.current,
          [headerKey]: nextDirtyHeader
        };
        setDirtyHeaders(dirtyHeadersRef.current);

        setSheets((prev) => {
          const sheet = prev[activeSheetName];
          if (!sheet) return prev;
          const newHeaders = [...sheet.headers];
          while (newHeaders.length <= c) newHeaders.push('');
          newHeaders[c] = String(value);
          return {
            ...prev,
            [activeSheetName]: {
              ...sheet,
              headers: newHeaders
            }
          };
        });
      } else {
        const dataRowIdx = r - 1;
        const cellKey = `${sheetId}:${dataRowIdx}:${c}`;
        const existingDirtyCell = dirtyCellsRef.current[cellKey];
        const baselineVersion =
          existingDirtyCell?.expectedCellVersion ??
          targetSheet.cellVersions?.[dataRowIdx]?.[c] ??
          0;

        const nextDirtyCell: DirtyCellEdit = {
          sheetId,
          sheetName: activeSheetName,
          rowIndex: dataRowIdx,
          colIndex: c,
          value,
          expectedCellVersion: baselineVersion
        };
        dirtyCellsRef.current = {
          ...dirtyCellsRef.current,
          [cellKey]: nextDirtyCell
        };
        setDirtyCells(dirtyCellsRef.current);

        setSheets((prev) => {
          const sheet = prev[activeSheetName];
          if (!sheet) return prev;

          const newRows = [...sheet.rows];
          while (newRows.length <= dataRowIdx) {
            newRows.push(new Array(totalCols).fill(''));
          }

          const updatedRow = [...(newRows[dataRowIdx] || [])];
          while (updatedRow.length <= c) {
            updatedRow.push('');
          }

          updatedRow[c] = value;
          newRows[dataRowIdx] = updatedRow;

          return {
            ...prev,
            [activeSheetName]: {
              ...sheet,
              rows: newRows
            }
          };
        });
      }

      setHasConflict(false);
      setSaveError(null);
      setSaveStatus('idle');
    },
    [activeSheetName, sheets, totalCols]
  );

  useEffect(() => {
    if (isEditing && editInputRef.current) {
      editInputRef.current.focus();
    }
  }, [isEditing]);

  // Commit current edit, preserving original scalar type if text unchanged
  const commitEdit = (moveToNextRow: boolean = false, moveToNextCol: boolean = false) => {
    if (!isEditing) return;
    const prevVal = getCellValue(selectedCell.row, selectedCell.col);
    if (String(prevVal ?? '') !== editValue) {
      setCellValue(selectedCell.row, selectedCell.col, editValue);
    }
    setIsEditing(false);

    if (moveToNextRow) {
      setSelectedCell((prev) => ({
        ...prev,
        row: Math.min(displayRowsCount - 1, prev.row + 1)
      }));
    } else if (moveToNextCol) {
      setSelectedCell((prev) => ({
        ...prev,
        col: Math.min(totalCols - 1, prev.col + 1)
      }));
    }
  };

  const cancelEdit = () => {
    setIsEditing(false);
  };

  // Keyboard navigation & direct typing (Ctrl+S handled uniformly by active-workspace window listener)
  const handleKeyDown = (e: React.KeyboardEvent) => {
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') {
      e.preventDefault();
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

    if (e.key === 'ArrowUp') {
      e.preventDefault();
      setSelectedCell((prev) => ({
        ...prev,
        row: Math.max(0, prev.row - 1)
      }));
    } else if (e.key === 'ArrowDown') {
      e.preventDefault();
      setSelectedCell((prev) => ({
        ...prev,
        row: Math.min(displayRowsCount - 1, prev.row + 1)
      }));
    } else if (e.key === 'ArrowLeft') {
      e.preventDefault();
      setSelectedCell((prev) => ({
        ...prev,
        col: Math.max(0, prev.col - 1)
      }));
    } else if (e.key === 'ArrowRight') {
      e.preventDefault();
      setSelectedCell((prev) => ({
        ...prev,
        col: Math.min(totalCols - 1, prev.col + 1)
      }));
    } else if (e.key === 'Tab') {
      e.preventDefault();
      if (e.shiftKey) {
        setSelectedCell((prev) => ({
          ...prev,
          col: Math.max(0, prev.col - 1)
        }));
      } else {
        setSelectedCell((prev) => ({
          ...prev,
          col: Math.min(totalCols - 1, prev.col + 1)
        }));
      }
    } else if (e.key === 'Enter' || e.key === 'F2') {
      e.preventDefault();
      const currentVal = getCellValue(selectedCell.row, selectedCell.col);
      setEditValue(String(currentVal ?? ''));
      setIsEditing(true);
    } else if (e.key === 'Delete' || e.key === 'Backspace') {
      e.preventDefault();
      setCellValue(selectedCell.row, selectedCell.col, '');
    } else if (e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) {
      setEditValue(e.key);
      setIsEditing(true);
    }
  };

  const handleCellClick = (r: number, c: number) => {
    if (selectedCell.row === r && selectedCell.col === c) {
      if (!isEditing) {
        const val = getCellValue(r, c);
        setEditValue(String(val ?? ''));
        setIsEditing(true);
      }
    } else {
      if (isEditing) {
        commitEdit();
      }
      setSelectedCell({ row: r, col: c });
    }
  };

  const handleCellDoubleClick = (r: number, c: number) => {
    if (isEditing) commitEdit();
    setSelectedCell({ row: r, col: c });
    const val = getCellValue(r, c);
    setEditValue(String(val ?? ''));
    setIsEditing(true);
  };

  const handleCopy = (e: React.ClipboardEvent) => {
    if (!isEditing) {
      const val = getCellValue(selectedCell.row, selectedCell.col);
      e.clipboardData.setData('text/plain', String(val ?? ''));
      e.preventDefault();
    }
  };

  const handlePaste = (e: React.ClipboardEvent) => {
    if (!isEditing) {
      const text = e.clipboardData.getData('text/plain');
      if (text) {
        e.preventDefault();
        setCellValue(selectedCell.row, selectedCell.col, text.trim());
      }
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

          // Remove saved key from dirty headers
          const updatedDirtyHeaders = { ...dirtyHeadersRef.current };
          delete updatedDirtyHeaders[key];
          dirtyHeadersRef.current = updatedDirtyHeaders;
          setDirtyHeaders(updatedDirtyHeaders);

          // Update sheet headerVersion & keep remaining dirty headers aligned
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
        // Chunk into <= 500 edits per RPC call
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
              return {
                ...prev,
                [sheetName]: {
                  ...target,
                  cellVersions: nextCellVersions,
                  rowVersions: nextRowVersions
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
              });

              return {
                ...prev,
                [sheetName]: {
                  ...target,
                  cellVersions: nextCellVersions,
                  rowVersions: nextRowVersions
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

  const isActiveWorkspace = workspace === activeWorkspace;
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

  // Export to Excel (active workspace ONLY)
  const handleExportToExcel = () => {
    try {
      const wb = XLSX.utils.book_new();
      const sheetKeys = Object.keys(sheets);

      for (const sheetName of sheetKeys) {
        const sheet = sheets[sheetName];
        if (!sheet) continue;

        const nonEmptyRows = sheet.rows.filter(
          (r) => r && r.some((cell) => cell !== '' && cell !== undefined && cell !== null)
        );

        const sheetAoa = [sheet.headers, ...nonEmptyRows];
        const ws = XLSX.utils.aoa_to_sheet(sheetAoa);
        XLSX.utils.book_append_sheet(wb, ws, sheetName);
      }

      const today = new Date().toISOString().slice(0, 10);
      const filename =
        workspace === 'doctor'
          ? `MK_Eye_Center_Doctor_Export_${today}.xlsx`
          : `MK_Eye_Center_Reception_Export_${today}.xlsx`;

      XLSX.writeFile(wb, filename);
    } catch (err) {
      console.error('Failed to export to Excel:', err);
      alert('Failed to generate Excel file.');
    }
  };

  // Add Worksheet via '+' button -> create_workspace_sheet RPC
  const handleCreateNewSheet = async (e: React.FormEvent) => {
    e.preventDefault();
    const trimmed = newSheetNameInput.trim();
    if (!trimmed) {
      setAddSheetError('Please enter a worksheet name.');
      return;
    }

    if (trimmed.length > 31) {
      setAddSheetError('Worksheet name cannot exceed 31 characters.');
      return;
    }

    if (/[\\/?*[\]:]/.test(trimmed)) {
      setAddSheetError('Name cannot contain any of: \\ / ? * [ ] :');
      return;
    }

    const existingNames = Object.keys(sheets);
    if (existingNames.some((n) => n.toLowerCase() === trimmed.toLowerCase())) {
      setAddSheetError(`A worksheet named "${trimmed}" already exists.`);
      return;
    }

    setIsCreatingSheet(true);
    setAddSheetError(null);

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
          sheetName: trimmed,
          headers: defaultHeaders
        })
      });

      const data = await res.json();
      if (!res.ok) {
        throw new Error(data.error || 'Failed to create worksheet on server.');
      }

      const createdSheet: SheetState = data.sheet || {
        name: trimmed,
        headers: defaultHeaders,
        headerVersion: 1,
        rows: [],
        rowVersions: {},
        cellVersions: {}
      };

      setSheets((prev) => ({
        ...prev,
        [trimmed]: createdSheet
      }));

      setActiveSheetName(trimmed);
      setSelectedCell({ row: 1, col: 0 });
      setIsAddSheetModalOpen(false);
      setNewSheetNameInput('');
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : 'Failed to create worksheet.';
      setAddSheetError(msg);
    } finally {
      setIsCreatingSheet(false);
    }
  };

  // Column resize mouse drag handlers
  const handleResizeStart = (colIndex: number, e: React.MouseEvent) => {
    e.stopPropagation();
    e.preventDefault();
    const startWidth = columnWidths[colIndex] || 110;
    resizingColRef.current = { col: colIndex, startX: e.clientX, startWidth };

    const handleMouseMove = (moveEvent: MouseEvent) => {
      if (!resizingColRef.current) return;
      const delta = moveEvent.clientX - resizingColRef.current.startX;
      const newWidth = Math.max(50, resizingColRef.current.startWidth + delta);
      setColumnWidths((prev) => ({
        ...prev,
        [resizingColRef.current!.col]: newWidth
      }));
    };

    const handleMouseUp = () => {
      resizingColRef.current = null;
      window.removeEventListener('mousemove', handleMouseMove);
      window.removeEventListener('mouseup', handleMouseUp);
    };

    window.addEventListener('mousemove', handleMouseMove);
    window.addEventListener('mouseup', handleMouseUp);
  };

  const sheetKeys = Object.keys(sheets);
  const visibleWorkspaces = allowedWorkspaces && allowedWorkspaces.length > 0 ? allowedWorkspaces : [workspace];
  const staffLabel = user?.name || user?.username || workspace;

  return (
    <div
      className="flex flex-col h-screen w-full bg-white select-none overflow-hidden font-sans text-xs text-slate-800"
      onKeyDown={handleKeyDown}
      onCopy={handleCopy}
      onPaste={handlePaste}
      tabIndex={0}
    >
      {/* 1. Compact Application Toolbar */}
      <div className="h-10 px-3 bg-slate-100 border-b border-slate-300 flex items-center justify-between shrink-0">
        <div className="flex items-center gap-2">
          <button
            type="button"
            onClick={() => void handleSave()}
            disabled={saveStatus === 'saving' || isLoading || !currentSheet.id}
            className="px-3 py-1 font-semibold text-white bg-slate-900 hover:bg-slate-800 active:bg-black rounded-xs transition disabled:opacity-50 cursor-pointer flex items-center gap-1.5"
          >
            {saveStatus === 'saving' ? 'Saving...' : 'Save'}
          </button>

          <button
            type="button"
            onClick={handleExportToExcel}
            disabled={sheetKeys.length === 0}
            className="px-3 py-1 font-semibold text-slate-800 bg-white hover:bg-slate-50 border border-slate-300 rounded-xs transition disabled:opacity-50 cursor-pointer"
          >
            Export to Excel
          </button>

          {user?.role === 'admin' && canAdminImport && onRequestInitialImport && (
            <button
              type="button"
              onClick={onRequestInitialImport}
              className="px-3 py-1 font-semibold text-white bg-emerald-700 hover:bg-emerald-800 rounded-xs transition cursor-pointer"
            >
              Import Initial Workbooks
            </button>
          )}

          {/* Save & Conflict Status Notice */}
          <div className="ml-2 flex items-center gap-2 text-[11px]" role="status" aria-live="polite">
            {isLoading && <span className="text-slate-500">Loading workbook from Supabase...</span>}
            {!isLoading && saveStatus === 'saving' && (
              <span className="text-slate-600 font-medium">Saving changes...</span>
            )}
            {!isLoading && saveStatus === 'saved' && !hasUnsavedChanges && (
              <span className="text-emerald-700 font-semibold">{saveMessage || 'Saved changes'}</span>
            )}
            {!isLoading && hasUnsavedChanges && saveStatus !== 'saving' && (
              <span className="text-amber-700 font-medium">Unsaved edits</span>
            )}
            {!isLoading && saveStatus === 'error' && (
              <span className="text-red-600 font-medium">{saveError || 'Save failed'}</span>
            )}
            {!isLoading && hasConflict && onReloadWorkspace && (
              <button
                type="button"
                onClick={() => void handleReloadLatest()}
                disabled={isReloading}
                className="px-2 py-0.5 bg-red-600 text-white rounded-xs font-medium hover:bg-red-700 cursor-pointer disabled:opacity-50"
              >
                {isReloading ? 'Reloading...' : 'Reload Latest'}
              </button>
            )}
          </div>
        </div>

        {/* Center/Right: Role-Authorized Workspace Tabs + Staff Identity + Sign Out */}
        <div className="flex items-center gap-4">
          {visibleWorkspaces.length > 1 ? (
            <div className="flex items-center bg-slate-200/90 p-0.5 rounded-xs border border-slate-300 text-xs">
              {visibleWorkspaces.includes('reception') && (
                <button
                  type="button"
                  onClick={() => onWorkspaceSwitchRequest('reception')}
                  className={`px-3 py-0.5 font-semibold text-xs rounded-xs transition cursor-pointer ${
                    activeWorkspace === 'reception'
                      ? 'bg-white text-slate-900 shadow-xs'
                      : 'text-slate-600 hover:text-slate-900'
                  }`}
                >
                  Reception
                </button>
              )}
              {visibleWorkspaces.includes('doctor') && (
                <button
                  type="button"
                  onClick={() => onWorkspaceSwitchRequest('doctor')}
                  className={`px-3 py-0.5 font-semibold text-xs rounded-xs transition cursor-pointer ${
                    activeWorkspace === 'doctor'
                      ? 'bg-white text-slate-900 shadow-xs'
                      : 'text-slate-600 hover:text-slate-900'
                  }`}
                >
                  Doctor
                </button>
              )}
            </div>
          ) : (
            <span className="px-2.5 py-0.5 font-semibold text-xs bg-white text-slate-900 border border-slate-300 rounded-xs">
              {workspace === 'doctor' ? 'Doctor Workspace' : 'Reception Workspace'}
            </span>
          )}

          <div className="flex items-center gap-3 border-l border-slate-300 pl-3">
            <span className="text-slate-600 font-mono text-[11px]" title={user?.role}>
              {staffLabel}
            </span>
            <button
              type="button"
              onClick={onLogout}
              className="text-slate-600 hover:text-slate-900 underline text-[11px] cursor-pointer"
            >
              Sign Out
            </button>
          </div>
        </div>
      </div>

      {/* 2. Formula & Cell Address Bar */}
      <div className="h-7 px-3 bg-white border-b border-slate-300 flex items-center gap-2 shrink-0 text-xs">
        <div className="w-14 h-5 px-1 bg-slate-50 border border-slate-300 rounded-xs flex items-center justify-center font-mono font-semibold text-slate-700 text-[11px] select-none shrink-0">
          {getColumnLetter(selectedCell.col)}
          {selectedCell.row + 1}
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
          className="flex-1 h-5 px-1.5 bg-transparent text-slate-900 text-xs font-normal outline-none focus:bg-slate-50 rounded-xs disabled:opacity-50"
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
      <div ref={gridContainerRef} className="flex-1 overflow-auto bg-white relative outline-none">
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
          <table className="border-separate border-spacing-0 table-fixed min-w-full text-xs">
            <thead>
              {/* Row 0: Excel Column Letters (sticky top-0, height 24px / h-6) */}
              <tr className="h-6">
                <th className="w-12 min-w-12 h-6 border-r border-b border-slate-300 bg-slate-200 sticky top-0 left-0 z-50 select-none text-slate-400 font-normal text-center" />

                {Array.from({ length: totalCols }).map((_, cIdx) => {
                  const width = columnWidths[cIdx] || 110;
                  return (
                    <th
                      key={cIdx}
                      style={{ width: `${width}px`, minWidth: `${width}px` }}
                      className="h-6 border-r border-b border-slate-300 bg-slate-100 font-medium text-slate-700 text-center select-none sticky top-0 z-30"
                    >
                      <span>{getColumnLetter(cIdx)}</span>
                      <div
                        onMouseDown={(e) => handleResizeStart(cIdx, e)}
                        className="absolute right-0 top-0 bottom-0 w-1.5 cursor-col-resize hover:bg-slate-400 z-20"
                      />
                    </th>
                  );
                })}
              </tr>

              {/* Row 1: Sticky Worksheet Column Header Titles (sticky top-6 / 24px, height 28px / h-7) */}
              <tr className="h-7">
                <th className="w-12 min-w-12 h-7 border-r border-b-2 border-slate-300 border-b-slate-400 bg-[#d5e0ef] text-slate-800 font-bold text-center sticky top-6 left-0 z-40 select-none shadow-[0_2px_4px_-1px_rgba(15,23,42,0.12)]">
                  1
                </th>

                {Array.from({ length: totalCols }).map((_, cIdx) => {
                  const isSelected = selectedCell.row === 0 && selectedCell.col === cIdx;
                  const isCellEditing = isSelected && isEditing;
                  const cellVal = getCellValue(0, cIdx);
                  const width = columnWidths[cIdx] || 110;

                  return (
                    <td
                      key={cIdx}
                      style={{ width: `${width}px`, minWidth: `${width}px` }}
                      onClick={() => handleCellClick(0, cIdx)}
                      onDoubleClick={() => handleCellDoubleClick(0, cIdx)}
                      className={`h-7 px-2 border-r border-b-2 border-slate-300 border-b-slate-400 bg-[#e3ebf6] font-bold text-[12px] tracking-[0.01em] text-slate-900 truncate sticky top-6 shadow-[0_2px_4px_-1px_rgba(15,23,42,0.1)] ${
                        isSelected || isCellEditing ? 'z-[25]' : 'z-20'
                      } ${
                        isSelected && !isCellEditing
                          ? 'outline-2 outline-slate-900 outline-offset-[-2px]'
                          : ''
                      }`}
                    >
                      {isCellEditing ? (
                        <input
                          ref={editInputRef}
                          type="text"
                          value={editValue}
                          onChange={(e) => setEditValue(e.target.value)}
                          onBlur={() => commitEdit(false, false)}
                          className="absolute inset-0 w-full h-full px-2 text-xs font-semibold bg-white text-slate-900 border-2 border-slate-900 outline-none z-20"
                        />
                      ) : (
                        <span>{cellVal !== undefined && cellVal !== null ? String(cellVal) : ''}</span>
                      )}
                    </td>
                  );
                })}
              </tr>
            </thead>

            <tbody>
              {Array.from({ length: displayRowsCount - 1 }).map((_, offsetIdx) => {
                const rIdx = offsetIdx + 1;
                const excelRowNum = rIdx + 1;

                return (
                  <tr key={rIdx} className="h-6">
                    <th className="w-12 min-w-12 h-6 border-r border-b border-slate-300 bg-slate-100 text-slate-600 font-normal text-center sticky left-0 z-[15] select-none">
                      {excelRowNum}
                    </th>

                    {Array.from({ length: totalCols }).map((_, cIdx) => {
                      const isSelected = selectedCell.row === rIdx && selectedCell.col === cIdx;
                      const isCellEditing = isSelected && isEditing;
                      const cellVal = getCellValue(rIdx, cIdx);

                      return (
                        <td
                          key={cIdx}
                          onClick={() => handleCellClick(rIdx, cIdx)}
                          onDoubleClick={() => handleCellDoubleClick(rIdx, cIdx)}
                          className={`h-6 px-1.5 border-r border-b border-slate-300 truncate relative text-slate-900 bg-white font-normal ${
                            isSelected && !isCellEditing
                              ? 'outline-2 outline-slate-900 outline-offset-[-2px] z-10'
                              : ''
                          }`}
                        >
                          {isCellEditing ? (
                            <input
                              ref={editInputRef}
                              type="text"
                              value={editValue}
                              onChange={(e) => setEditValue(e.target.value)}
                              onBlur={() => commitEdit(false, false)}
                              className="absolute inset-0 w-full h-full px-1.5 text-xs bg-white text-slate-900 border-2 border-slate-900 outline-none z-10"
                            />
                          ) : (
                            <span>{cellVal !== undefined && cellVal !== null ? String(cellVal) : ''}</span>
                          )}
                        </td>
                      );
                    })}
                  </tr>
                );
              })}
            </tbody>
          </table>
        )}
      </div>

      {/* 4. Worksheet Tabs along the bottom with '+' button */}
      <div className="h-8 bg-slate-200 border-t border-slate-300 flex items-center px-2 gap-1 shrink-0 overflow-x-auto">
        {sheetKeys.map((name) => {
          const isActive = activeSheetName === name;
          return (
            <button
              key={name}
              type="button"
              onClick={() => {
                if (isEditing) commitEdit();
                setActiveSheetName(name);
                setSelectedCell({ row: 1, col: 0 });
              }}
              className={`px-3 py-1 text-xs font-medium border-t-2 rounded-t-xs transition cursor-pointer whitespace-nowrap ${
                isActive
                  ? 'bg-white border-slate-900 text-slate-900 shadow-xs'
                  : 'bg-slate-200 border-transparent text-slate-600 hover:bg-slate-100 hover:text-slate-900'
              }`}
            >
              {name}
            </button>
          );
        })}

        <button
          type="button"
          onClick={() => {
            setAddSheetError(null);
            setNewSheetNameInput('');
            setIsAddSheetModalOpen(true);
          }}
          title={`New worksheet in ${workspace === 'doctor' ? 'Doctor' : 'Reception'}`}
          className="h-6 w-6 flex items-center justify-center rounded-xs bg-slate-300/80 hover:bg-slate-300 active:bg-slate-400 text-slate-700 font-bold text-sm cursor-pointer ml-1 shrink-0 transition"
        >
          +
        </button>
      </div>

      {/* 5. Modal for Adding a New Worksheet */}
      {isAddSheetModalOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/30 backdrop-blur-xs">
          <div className="w-full max-w-sm bg-white rounded-xs shadow-xl border border-slate-300 p-5 space-y-4">
            <div>
              <h3 className="text-sm font-bold text-slate-900">
                New Worksheet ({workspace === 'reception' ? 'Reception' : 'Doctor'})
              </h3>
              <p className="text-xs text-slate-500 mt-0.5">
                Enter a name for the new worksheet tab in the {workspace === 'doctor' ? 'Doctor' : 'Reception'} workbook.
              </p>
            </div>

            <form onSubmit={handleCreateNewSheet} className="space-y-3">
              <div>
                <input
                  type="text"
                  autoFocus
                  placeholder="e.g. CONSULTATIONS, Sheet3"
                  value={newSheetNameInput}
                  onChange={(e) => setNewSheetNameInput(e.target.value)}
                  maxLength={31}
                  className="w-full h-8 px-2.5 text-xs bg-white border border-slate-300 rounded-xs focus:border-slate-900 focus:outline-none"
                />
                <span className="text-[10px] text-slate-400 mt-1 block">
                  Max 31 characters. Cannot contain: \ / ? * [ ] :
                </span>
              </div>

              {addSheetError && (
                <div className="p-2 bg-red-50 border border-red-200 rounded-xs text-[11px] text-red-700">
                  {addSheetError}
                </div>
              )}

              <div className="flex items-center justify-end gap-2 pt-2 border-t border-slate-200">
                <button
                  type="button"
                  onClick={() => setIsAddSheetModalOpen(false)}
                  disabled={isCreatingSheet}
                  className="px-3 py-1.5 text-xs text-slate-700 hover:bg-slate-100 rounded-xs border border-slate-300 cursor-pointer"
                >
                  Cancel
                </button>
                <button
                  type="submit"
                  disabled={isCreatingSheet || !newSheetNameInput.trim()}
                  className="px-3 py-1.5 text-xs font-semibold text-white bg-slate-900 hover:bg-slate-800 rounded-xs disabled:opacity-50 cursor-pointer"
                >
                  {isCreatingSheet ? 'Creating...' : 'Create Worksheet'}
                </button>
              </div>
            </form>
          </div>
        </div>
      )}
    </div>
  );
}
