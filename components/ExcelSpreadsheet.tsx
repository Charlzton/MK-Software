'use client';

import React, { useState, useEffect, useRef, useCallback, useMemo } from 'react';
import * as XLSX from 'xlsx';

export interface SheetState {
  headers: string[];
  rows: (string | number)[][];
}

export interface ExcelSpreadsheetProps {
  workspace: 'reception' | 'doctor';
  initialSheets: Record<string, SheetState>;
  activeWorkspace: 'reception' | 'doctor';
  onWorkspaceSwitchRequest: (targetWorkspace: 'reception' | 'doctor') => void;
  onUnsavedStatusChange: (hasUnsaved: boolean) => void;
  onLogout: () => void;
  isConfigured?: boolean;
  configNotice?: string;
  defaultSheetName?: string;
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
  initialSheets,
  activeWorkspace,
  onWorkspaceSwitchRequest,
  onUnsavedStatusChange,
  onLogout,
  isConfigured = true,
  configNotice,
  defaultSheetName
}: ExcelSpreadsheetProps) {
  const initialTabName = useMemo(() => {
    if (defaultSheetName && initialSheets[defaultSheetName]) return defaultSheetName;
    const keys = Object.keys(initialSheets);
    if (keys.length > 0) return keys[0];
    return workspace === 'doctor' ? 'Lens Rx' : "PATIENT'S DATA";
  }, [defaultSheetName, initialSheets, workspace]);

  const [activeSheetName, setActiveSheetName] = useState<string>(initialTabName);
  const [sheets, setSheets] = useState<Record<string, SheetState>>(initialSheets);

  // Synchronize when initialSheets prop updates
  const [prevInitialSheets, setPrevInitialSheets] = useState(initialSheets);
  if (initialSheets !== prevInitialSheets) {
    setPrevInitialSheets(initialSheets);
    setSheets(initialSheets);
    if (!initialSheets[activeSheetName]) {
      const keys = Object.keys(initialSheets);
      if (keys.length > 0) {
        setActiveSheetName(keys[0]);
      }
    }
  }

  // Track unsaved modifications
  const [hasUnsavedChanges, setHasUnsavedChanges] = useState<boolean>(false);
  const [saveStatus, setSaveStatus] = useState<'idle' | 'saving' | 'saved' | 'error'>('idle');
  const [saveMessage, setSaveMessage] = useState<string>('');
  const [saveError, setSaveError] = useState<string | null>(null);

  // Notify parent of unsaved changes status
  useEffect(() => {
    onUnsavedStatusChange(hasUnsavedChanges);
  }, [hasUnsavedChanges, onUnsavedStatusChange]);

  // Active & Editing cell coordinates: { row: number, col: number }
  // row 0 is Header/Row 1 in Excel, row 1 is Row 2 in Excel...
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

  // Update cell value
  const setCellValue = useCallback(
    (r: number, c: number, value: string | number) => {
      setSheets((prev) => {
        const sheet = prev[activeSheetName];
        if (!sheet) return prev;

        if (r === 0) {
          // Modifying row 1 (headers)
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
        }

        const dataRowIdx = r - 1;
        const newRows = [...sheet.rows];

        // Pad rows if typing into blank rows below
        while (newRows.length <= dataRowIdx) {
          newRows.push(new Array(totalCols).fill(''));
        }

        const updatedRow = [...newRows[dataRowIdx]];
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

      setHasUnsavedChanges(true);
      setSaveStatus('idle');
    },
    [activeSheetName, totalCols]
  );

  // Focus input on edit mode
  useEffect(() => {
    if (isEditing && editInputRef.current) {
      editInputRef.current.focus();
    }
  }, [isEditing]);

  // Commit current edit
  const commitEdit = (moveToNextRow: boolean = false, moveToNextCol: boolean = false) => {
    if (!isEditing) return;
    setCellValue(selectedCell.row, selectedCell.col, editValue);
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

  // Keyboard navigation & direct typing
  const handleKeyDown = (e: React.KeyboardEvent) => {
    // Intercept Ctrl+S / Cmd+S for saving
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') {
      e.preventDefault();
      handleSave();
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

    // When NOT in edit mode:
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
      // Direct typing into selected cell
      setEditValue(e.key);
      setIsEditing(true);
    }
  };

  // Cell click handler
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

  // Clipboard Copy & Paste
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

  // Primary Action A: Save (Saves this workspace ONLY)
  const handleSave = async () => {
    if (isEditing) commitEdit();

    setSaveStatus('saving');
    setSaveError(null);

    try {
      // Build payload preserving both headers and data rows
      const payloadSheets: Record<string, { headers: string[]; rows: (string | number)[][] }> = {};
      for (const [name, s] of Object.entries(sheets)) {
        payloadSheets[name] = {
          headers: s.headers,
          rows: s.rows
        };
      }

      const res = await fetch('/api/records', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          workspace,
          sheets: payloadSheets
        })
      });

      const data = await res.json();
      if (!res.ok) {
        throw new Error(data.error || 'Failed to save changes.');
      }

      setSaveStatus('saved');
      setSaveMessage(data.message || (workspace === 'doctor' && !isConfigured ? 'Saved locally (Doctor workbook not connected)' : 'Saved to Google Sheets'));
      setHasUnsavedChanges(false);
      setTimeout(() => {
        setSaveStatus('idle');
      }, 4000);
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : 'Save error occurred';
      setSaveStatus('error');
      setSaveError(msg);
    }
  };

  // Primary Action B: Export to Excel (Exports active workspace ONLY)
  const handleExportToExcel = () => {
    try {
      const wb = XLSX.utils.book_new();
      const sheetKeys = Object.keys(sheets);

      for (const sheetName of sheetKeys) {
        const sheet = sheets[sheetName];
        if (!sheet) continue;

        // Clean trailing completely empty rows
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

  // Primary Action C: Add Worksheet via '+' button (Adds to this workspace ONLY)
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

    const illegalChars = /[\\/?*[\]:]/;
    if (illegalChars.test(trimmed)) {
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

      // Add to local state immediately
      setSheets((prev) => ({
        ...prev,
        [trimmed]: {
          headers: defaultHeaders,
          rows: []
        }
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
          {/* Action A: Save */}
          <button
            type="button"
            onClick={handleSave}
            disabled={saveStatus === 'saving'}
            className="px-3 py-1 font-semibold text-white bg-slate-900 hover:bg-slate-800 active:bg-black rounded-xs transition disabled:opacity-50 cursor-pointer flex items-center gap-1.5"
          >
            {saveStatus === 'saving' ? 'Saving...' : 'Save'}
          </button>

          {/* Action B: Export to Excel */}
          <button
            type="button"
            onClick={handleExportToExcel}
            className="px-3 py-1 font-semibold text-slate-800 bg-white hover:bg-slate-50 border border-slate-300 rounded-xs transition cursor-pointer"
          >
            Export to Excel
          </button>

          {/* Save Status Notice */}
          <div className="ml-2 flex items-center text-[11px]">
            {saveStatus === 'saving' && (
              <span className="text-slate-500">
                Saving to {workspace === 'reception' ? 'Google Sheets' : isConfigured ? 'Doctor Google Sheets' : 'Doctor local store'}...
              </span>
            )}
            {saveStatus === 'saved' && (
              <span className="text-emerald-700 font-medium">{saveMessage || 'Saved'}</span>
            )}
            {saveStatus === 'error' && (
              <span className="text-red-600 font-medium">{saveError || 'Save failed'}</span>
            )}
            {hasUnsavedChanges && saveStatus === 'idle' && (
              <span className="text-amber-700 font-medium">Unsaved edits</span>
            )}
          </div>
        </div>

        {/* Center/Right: Two Main Navigation Tabs (Reception & Doctor) + Admin Badge + Sign Out */}
        <div className="flex items-center gap-4">
          {/* Main Workspace Navigation Tabs */}
          <div className="flex items-center bg-slate-200/90 p-0.5 rounded-xs border border-slate-300 text-xs">
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
          </div>

          {/* Minimal Session Indicator & Logout */}
          <div className="flex items-center gap-3 border-l border-slate-300 pl-3">
            <span className="text-slate-500 font-mono text-[11px]">admin</span>
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
        {/* Name Box (Current Cell Reference, e.g. A1, E14) */}
        <div className="w-14 h-5 px-1 bg-slate-50 border border-slate-300 rounded-xs flex items-center justify-center font-mono font-semibold text-slate-700 text-[11px] select-none shrink-0">
          {getColumnLetter(selectedCell.col)}
          {selectedCell.row + 1}
        </div>
        <span className="text-slate-400 font-serif italic text-xs select-none">fx</span>
        <div className="h-4 w-px bg-slate-300 shrink-0" />
        {/* Formula / Cell value input */}
        <input
          type="text"
          value={isEditing ? editValue : String(getCellValue(selectedCell.row, selectedCell.col) ?? '')}
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
          className="flex-1 h-5 px-1.5 bg-transparent text-slate-900 text-xs font-normal outline-none focus:bg-slate-50 rounded-xs"
        />
      </div>

      {/* 2.1 Unconfigured Doctor Workbook Notice Strip */}
      {workspace === 'doctor' && !isConfigured && (
        <div className="bg-amber-50/95 border-b border-amber-300 px-3 py-1.5 flex items-center justify-between text-xs text-amber-900 shrink-0">
          <div className="flex items-center gap-2.5">
            <span className="font-bold px-1.5 py-0.5 bg-amber-200 text-amber-950 text-[10px] rounded-xs uppercase tracking-wide">
              Doctor Workbook Not Connected
            </span>
            <span className="text-[11px] text-amber-900">
              {configNotice ||
                "Doctor's Google Sheets workbook is not connected yet. Once imported into Google Sheets, configure DOCTOR_SPREADSHEET_ID in environment variables. Doctor data is saved separately in the Doctor workspace."}
            </span>
          </div>
          <span className="text-[10px] font-mono text-amber-800 bg-amber-100/80 border border-amber-300 px-1.5 py-0.5 rounded-xs shrink-0">
            Isolated Doctor Storage
          </span>
        </div>
      )}

      {/* 3. Main Spreadsheet Grid Viewport */}
      <div ref={gridContainerRef} className="flex-1 overflow-auto bg-white relative outline-none">
        <table className="border-collapse table-fixed min-w-full text-xs">
          {/* Column Header Row (Letters A, B, C...) */}
          <thead>
            <tr className="sticky top-0 z-20 bg-slate-100">
              {/* Top-left corner cell */}
              <th className="w-12 min-w-12 h-6 border-r border-b border-slate-300 bg-slate-200 sticky left-0 z-30 select-none text-slate-400 font-normal text-center" />

              {Array.from({ length: totalCols }).map((_, cIdx) => {
                const width = columnWidths[cIdx] || 110;
                return (
                  <th
                    key={cIdx}
                    style={{ width: `${width}px`, minWidth: `${width}px` }}
                    className="h-6 border-r border-b border-slate-300 bg-slate-100 font-medium text-slate-700 text-center select-none relative"
                  >
                    <span>{getColumnLetter(cIdx)}</span>
                    {/* Column Resizer Handle */}
                    <div
                      onMouseDown={(e) => handleResizeStart(cIdx, e)}
                      className="absolute right-0 top-0 bottom-0 w-1.5 cursor-col-resize hover:bg-slate-400 z-10"
                    />
                  </th>
                );
              })}
            </tr>
          </thead>

          {/* Grid Rows (Row 1 is headers/row 1, Rows 2..N are data rows) */}
          <tbody>
            {Array.from({ length: displayRowsCount }).map((_, rIdx) => {
              const excelRowNum = rIdx + 1;
              const isHeaderRow = rIdx === 0;

              return (
                <tr key={rIdx} className="h-6">
                  {/* Row Number Header down the left */}
                  <th className="w-12 min-w-12 h-6 border-r border-b border-slate-300 bg-slate-100 text-slate-600 font-normal text-center sticky left-0 z-10 select-none">
                    {excelRowNum}
                  </th>

                  {/* Row Cells */}
                  {Array.from({ length: totalCols }).map((_, cIdx) => {
                    const isSelected = selectedCell.row === rIdx && selectedCell.col === cIdx;
                    const isCellEditing = isSelected && isEditing;
                    const cellVal = getCellValue(rIdx, cIdx);

                    return (
                      <td
                        key={cIdx}
                        onClick={() => handleCellClick(rIdx, cIdx)}
                        onDoubleClick={() => handleCellDoubleClick(rIdx, cIdx)}
                        className={`h-6 px-1.5 border-r border-b border-slate-300 truncate relative text-slate-900 ${
                          isHeaderRow && cellVal !== ''
                            ? 'font-medium bg-slate-50/50'
                            : 'bg-white font-normal'
                        } ${
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
                            className="absolute inset-0 w-full h-full px-1.5 text-xs bg-white text-slate-900 border-2 border-slate-900 outline-none z-20"
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

        {/* Small '+' Button similar to Microsoft Excel */}
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

      {/* 5. Clean Modal for Adding a New Worksheet */}
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
