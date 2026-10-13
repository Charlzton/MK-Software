'use client';

import React, { useState, useEffect } from 'react';
import {
  X,
  FileSpreadsheet,
  Link,
  RefreshCw,
  Copy,
  Check,
  CheckCircle2,
  AlertTriangle,
  History,
  Columns3,
  Sliders,
  ShieldCheck
} from 'lucide-react';
import { ClinicSettings, AuditLogEntry, ColumnDefinition, User } from '@/lib/types';
import { GOOGLE_APPS_SCRIPT_TEMPLATE } from '@/lib/scriptTemplate';

interface SettingsModalProps {
  isOpen: boolean;
  currentUser: User;
  onClose: () => void;
  onRefreshData: () => void;
}

export function SettingsModal({
  isOpen,
  currentUser,
  onClose,
  onRefreshData
}: SettingsModalProps) {
  const [activeTab, setActiveTab] = useState<'sheets' | 'schema' | 'audit' | 'clinic'>('sheets');
  const [settings, setSettings] = useState<ClinicSettings | null>(null);
  const [auditLogs, setAuditLogs] = useState<AuditLogEntry[]>([]);
  const [scriptUrlInput, setScriptUrlInput] = useState('');
  const [isTesting, setIsTesting] = useState(false);
  const [testResult, setTestResult] = useState<{ success: boolean; message: string } | null>(null);
  const [isSavingUrl, setIsSavingUrl] = useState(false);
  const [copiedCode, setCopiedCode] = useState(false);

  // New Column state
  const [newColWorksheet, setNewColWorksheet] = useState<'PATIENT\'S DATA' | 'DVLA'>('PATIENT\'S DATA');
  const [newColLabel, setNewColLabel] = useState('');
  const [newColType, setNewColType] = useState<'text' | 'number' | 'date'>('text');
  const [isAddingCol, setIsAddingCol] = useState(false);
  const [colSuccessMsg, setColSuccessMsg] = useState<string | null>(null);

  // Load settings & audit logs on open
  useEffect(() => {
    if (isOpen) {
      fetch('/api/settings')
        .then((res) => res.json())
        .then((data) => {
          setSettings(data);
          setScriptUrlInput(data.googleSheetsUrl || '');
        })
        .catch(console.error);

      fetch('/api/audit')
        .then((res) => res.json())
        .then((data) => {
          if (data.logs) setAuditLogs(data.logs);
        })
        .catch(console.error);

      const timer = setTimeout(() => {
        setTestResult(null);
        setColSuccessMsg(null);
      }, 0);
      return () => clearTimeout(timer);
    }
  }, [isOpen]);

  const handleTestConnection = async () => {
    setIsTesting(true);
    setTestResult(null);
    try {
      const res = await fetch('/api/sheets/sync', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          action: 'TEST_CONNECTION',
          scriptUrl: scriptUrlInput
        })
      });
      const data = await res.json();
      setTestResult({
        success: data.success,
        message: data.message
      });
    } catch {
      setTestResult({
        success: false,
        message: 'Network error communicating with synchronization API.'
      });
    } finally {
      setIsTesting(false);
    }
  };

  const handleSaveUrl = async () => {
    setIsSavingUrl(true);
    try {
      const res = await fetch('/api/sheets/sync', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          action: 'SAVE_URL',
          scriptUrl: scriptUrlInput
        })
      });
      const data = await res.json();
      if (data.settings) {
        setSettings(data.settings);
        setTestResult({
          success: true,
          message: 'Google Sheets synchronization endpoint saved successfully.'
        });
      }
    } catch {
      setTestResult({
        success: false,
        message: 'Failed to update Google Sheets settings.'
      });
    } finally {
      setIsSavingUrl(false);
    }
  };

  const handlePullFromSheets = async () => {
    setIsTesting(true);
    try {
      const res = await fetch('/api/sheets/sync', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          action: 'PULL_FROM_SHEETS',
          scriptUrl: scriptUrlInput
        })
      });
      const data = await res.json();
      setTestResult({
        success: data.success,
        message: data.message
      });
      if (data.success) {
        onRefreshData();
      }
    } catch {
      setTestResult({
        success: false,
        message: 'Failed to pull from Google Sheets.'
      });
    } finally {
      setIsTesting(false);
    }
  };

  const handleCopyScriptCode = () => {
    navigator.clipboard.writeText(GOOGLE_APPS_SCRIPT_TEMPLATE);
    setCopiedCode(true);
    setTimeout(() => setCopiedCode(false), 2500);
  };

  const handleAddColumn = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!newColLabel.trim()) return;

    setIsAddingCol(true);
    setColSuccessMsg(null);

    try {
      const res = await fetch('/api/schema', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          worksheet: newColWorksheet,
          column: {
            label: newColLabel.trim(),
            type: newColType
          }
        })
      });
      const data = await res.json();
      if (data.success) {
        setColSuccessMsg(data.message);
        setNewColLabel('');
        onRefreshData();
      } else {
        setColSuccessMsg(data.error || 'Failed to add column');
      }
    } catch {
      setColSuccessMsg('Network error adding column');
    } finally {
      setIsAddingCol(false);
    }
  };

  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-3 sm:p-4 bg-slate-900/40 backdrop-blur-xs">
      <div className="relative w-full max-w-3xl bg-white rounded-xl shadow-2xl border border-slate-200 overflow-hidden flex flex-col max-h-[88vh]">
        {/* Header */}
        <div className="flex items-center justify-between px-6 py-4 border-b border-slate-200 bg-slate-50">
          <div>
            <h2 className="text-base font-bold text-slate-900">
              System Settings & Google Sheets Hub
            </h2>
            <p className="text-xs text-slate-500">
              MK Eye Center · Persistent Storage, Column Schemas, and Audit Records
            </p>
          </div>
          <button
            type="button"
            onClick={onClose}
            className="p-1.5 text-slate-400 hover:text-slate-700 hover:bg-slate-100 rounded-md transition"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        {/* Tab Switcher */}
        <div className="flex items-center gap-1 px-6 pt-3 border-b border-slate-200 bg-white">
          <button
            type="button"
            onClick={() => setActiveTab('sheets')}
            className={`px-3.5 py-2 text-xs font-semibold border-b-2 transition -mb-px flex items-center gap-1.5 ${
              activeTab === 'sheets'
                ? 'border-slate-900 text-slate-900'
                : 'border-transparent text-slate-500 hover:text-slate-800'
            }`}
          >
            <FileSpreadsheet className="w-3.5 h-3.5" />
            <span>Google Sheets Sync</span>
          </button>

          <button
            type="button"
            onClick={() => setActiveTab('schema')}
            className={`px-3.5 py-2 text-xs font-semibold border-b-2 transition -mb-px flex items-center gap-1.5 ${
              activeTab === 'schema'
                ? 'border-slate-900 text-slate-900'
                : 'border-transparent text-slate-500 hover:text-slate-800'
            }`}
          >
            <Columns3 className="w-3.5 h-3.5" />
            <span>Column Manager</span>
          </button>

          <button
            type="button"
            onClick={() => setActiveTab('audit')}
            className={`px-3.5 py-2 text-xs font-semibold border-b-2 transition -mb-px flex items-center gap-1.5 ${
              activeTab === 'audit'
                ? 'border-slate-900 text-slate-900'
                : 'border-transparent text-slate-500 hover:text-slate-800'
            }`}
          >
            <History className="w-3.5 h-3.5" />
            <span>Audit Trail</span>
          </button>

          <button
            type="button"
            onClick={() => setActiveTab('clinic')}
            className={`px-3.5 py-2 text-xs font-semibold border-b-2 transition -mb-px flex items-center gap-1.5 ${
              activeTab === 'clinic'
                ? 'border-slate-900 text-slate-900'
                : 'border-transparent text-slate-500 hover:text-slate-800'
            }`}
          >
            <Sliders className="w-3.5 h-3.5" />
            <span>Clinic Details</span>
          </button>
        </div>

        {/* Tab Body */}
        <div className="p-6 overflow-y-auto flex-1 space-y-6 text-xs text-slate-700">
          {/* TAB 1: Google Sheets */}
          {activeTab === 'sheets' && (
            <div className="space-y-6">
              {/* Status Banner */}
              <div className="p-4 bg-slate-50 border border-slate-200 rounded-xl space-y-2">
                <div className="flex items-center justify-between">
                  <span className="font-semibold text-slate-800 flex items-center gap-1.5">
                    <ShieldCheck className="w-4 h-4 text-teal-600" />
                    Storage Engine Status
                  </span>
                  <span className="font-mono text-[11px] px-2 py-0.5 rounded bg-white border border-slate-200 font-semibold text-slate-700">
                    {settings?.googleSheetsUrl ? 'Live Google Apps Script Sync' : 'Autonomous Server Store (Ready to Link)'}
                  </span>
                </div>
                <p className="text-slate-600 leading-relaxed text-[11px]">
                  All patient registrations and DVLA updates are safely written to the central server store in real time. Linking a Google Apps Script Web App synchronizes updates directly into your clinic Google Sheet.
                </p>
              </div>

              {/* URL Input & Controls */}
              <div className="space-y-3">
                <label className="block font-semibold text-slate-800">
                  Google Apps Script Web App URL
                </label>
                <div className="flex gap-2">
                  <input
                    type="url"
                    value={scriptUrlInput}
                    onChange={(e) => setScriptUrlInput(e.target.value)}
                    placeholder="https://script.google.com/macros/s/AKfycbx.../exec"
                    className="flex-1 px-3 py-2 text-xs bg-white border border-slate-300 rounded-lg text-slate-900 font-mono focus:ring-2 focus:ring-slate-900 focus:outline-hidden"
                  />
                  <button
                    type="button"
                    onClick={handleSaveUrl}
                    disabled={isSavingUrl}
                    className="px-4 py-2 font-semibold text-white bg-slate-900 hover:bg-slate-800 rounded-lg transition"
                  >
                    {isSavingUrl ? 'Saving...' : 'Save URL'}
                  </button>
                </div>

                <div className="flex flex-wrap items-center gap-2 pt-1">
                  <button
                    type="button"
                    onClick={handleTestConnection}
                    disabled={isTesting || !scriptUrlInput}
                    className="px-3 py-1.5 font-medium text-slate-700 bg-white border border-slate-200 rounded-lg hover:bg-slate-50 transition flex items-center gap-1.5"
                  >
                    <RefreshCw className={`w-3.5 h-3.5 ${isTesting ? 'animate-spin' : ''}`} />
                    <span>Test Ping</span>
                  </button>

                  <button
                    type="button"
                    onClick={handlePullFromSheets}
                    disabled={isTesting || !scriptUrlInput}
                    className="px-3 py-1.5 font-medium text-slate-700 bg-white border border-slate-200 rounded-lg hover:bg-slate-50 transition flex items-center gap-1.5"
                  >
                    <span>Pull All Records from Sheets</span>
                  </button>
                </div>

                {testResult && (
                  <div
                    className={`p-3 rounded-lg border flex items-start gap-2 ${
                      testResult.success
                        ? 'bg-emerald-50 border-emerald-200 text-emerald-800'
                        : 'bg-red-50 border-red-200 text-red-800'
                    }`}
                  >
                    {testResult.success ? (
                      <CheckCircle2 className="w-4 h-4 shrink-0 mt-0.5" />
                    ) : (
                      <AlertTriangle className="w-4 h-4 shrink-0 mt-0.5" />
                    )}
                    <span>{testResult.message}</span>
                  </div>
                )}
              </div>

              {/* Instructions & Copyable Code */}
              <div className="pt-4 border-t border-slate-100 space-y-3">
                <div className="flex items-center justify-between">
                  <h3 className="font-bold text-slate-900">
                    Google Apps Script Deployment Code
                  </h3>
                  <button
                    type="button"
                    onClick={handleCopyScriptCode}
                    className="px-2.5 py-1 text-[11px] font-semibold text-slate-700 bg-slate-100 hover:bg-slate-200 rounded transition flex items-center gap-1"
                  >
                    {copiedCode ? (
                      <>
                        <Check className="w-3.5 h-3.5 text-teal-600" />
                        <span>Copied!</span>
                      </>
                    ) : (
                      <>
                        <Copy className="w-3.5 h-3.5" />
                        <span>Copy Code.gs</span>
                      </>
                    )}
                  </button>
                </div>

                <div className="p-3 bg-slate-900 rounded-lg text-slate-200 font-mono text-[11px] max-h-48 overflow-y-auto leading-relaxed select-all">
                  <pre>{GOOGLE_APPS_SCRIPT_TEMPLATE}</pre>
                </div>

                <div className="p-3 bg-slate-50 rounded-lg border border-slate-200 space-y-1 text-[11px] text-slate-600">
                  <span className="font-semibold text-slate-800 block">5-Step Setup Checklist:</span>
                  <ol className="list-decimal pl-4 space-y-0.5">
                    <li>Open your MK Eye Center Google Spreadsheet.</li>
                    <li>Ensure sheet tabs exist: <code>PATIENT&apos;S DATA</code>, <code>DVLA</code>, <code>MEDS</code>.</li>
                    <li>Click <strong>Extensions &gt; Apps Script</strong> and paste the copied code into <code>Code.gs</code>.</li>
                    <li>Click <strong>Deploy &gt; New deployment</strong>, select <strong>Web app</strong>, execute as <strong>Me</strong>, set access to <strong>Anyone</strong>.</li>
                    <li>Copy the Web App URL and paste it into the box above.</li>
                  </ol>
                </div>
              </div>
            </div>
          )}

          {/* TAB 2: Column Schema Manager */}
          {activeTab === 'schema' && (
            <div className="space-y-5">
              <div className="p-4 bg-slate-50 border border-slate-200 rounded-xl space-y-1">
                <p className="font-semibold text-slate-800">
                  Safe Schema Management
                </p>
                <p className="text-slate-600 leading-relaxed text-[11px]">
                  Add custom columns to the clinic registers without breaking existing column order or overwriting legacy data.
                </p>
              </div>

              <form onSubmit={handleAddColumn} className="space-y-4">
                <div>
                  <label className="block font-semibold text-slate-700 mb-1">
                    Select Worksheet
                  </label>
                  <select
                    value={newColWorksheet}
                    onChange={(e) => setNewColWorksheet(e.target.value as 'PATIENT\'S DATA' | 'DVLA')}
                    className="w-full px-3 py-2 bg-white border border-slate-300 rounded-lg text-slate-800"
                  >
                    <option value="PATIENT'S DATA">PATIENT&apos;S DATA (16 Existing Columns)</option>
                    <option value="DVLA">DVLA (Driver Eye Examination Register)</option>
                  </select>
                </div>

                <div className="grid grid-cols-2 gap-3">
                  <div>
                    <label className="block font-semibold text-slate-700 mb-1">
                      New Column Heading
                    </label>
                    <input
                      type="text"
                      required
                      value={newColLabel}
                      onChange={(e) => setNewColLabel(e.target.value)}
                      placeholder="e.g. NHIS_Number or Insurance_Type"
                      className="w-full px-3 py-2 bg-white border border-slate-300 rounded-lg text-slate-900"
                    />
                  </div>

                  <div>
                    <label className="block font-semibold text-slate-700 mb-1">
                      Data Type
                    </label>
                    <select
                      value={newColType}
                      onChange={(e) => setNewColType(e.target.value as 'text' | 'number' | 'date')}
                      className="w-full px-3 py-2 bg-white border border-slate-300 rounded-lg text-slate-800"
                    >
                      <option value="text">Text / Free-form</option>
                      <option value="number">Numeric / Amount</option>
                      <option value="date">Date</option>
                    </select>
                  </div>
                </div>

                {colSuccessMsg && (
                  <div className="p-2.5 bg-emerald-50 border border-emerald-200 rounded text-emerald-800">
                    {colSuccessMsg}
                  </div>
                )}

                <div className="pt-2">
                  <button
                    type="submit"
                    disabled={isAddingCol || !newColLabel.trim()}
                    className="px-4 py-2 font-semibold text-white bg-slate-900 hover:bg-slate-800 disabled:opacity-40 rounded-lg transition"
                  >
                    {isAddingCol ? 'Updating Schema...' : 'Safely Append Column'}
                  </button>
                </div>
              </form>
            </div>
          )}

          {/* TAB 3: Audit Trail */}
          {activeTab === 'audit' && (
            <div className="space-y-4">
              <div className="flex items-center justify-between">
                <div>
                  <h3 className="font-bold text-slate-900">Clinical Audit Trail</h3>
                  <p className="text-slate-500 text-[11px]">
                    Chronological immutable log of staff mutations and system synchronizations
                  </p>
                </div>
                <span className="font-mono text-slate-500 text-[11px]">
                  {auditLogs.length} events logged
                </span>
              </div>

              <div className="border border-slate-200 rounded-xl overflow-hidden max-h-80 overflow-y-auto">
                <table className="w-full border-collapse text-left text-[11px]">
                  <thead className="bg-slate-50 border-b border-slate-200 text-slate-500 sticky top-0">
                    <tr>
                      <th className="p-2.5 font-semibold">Timestamp</th>
                      <th className="p-2.5 font-semibold">Staff Actor</th>
                      <th className="p-2.5 font-semibold">Action</th>
                      <th className="p-2.5 font-semibold">Worksheet</th>
                      <th className="p-2.5 font-semibold">Details</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {auditLogs.map((log) => (
                      <tr key={log.id} className="hover:bg-slate-50/50">
                        <td className="p-2.5 font-mono text-slate-500 whitespace-nowrap">
                          {new Date(log.timestamp).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', second: '2-digit' })}
                        </td>
                        <td className="p-2.5 font-semibold text-slate-800 whitespace-nowrap">
                          {log.actor}
                        </td>
                        <td className="p-2.5 whitespace-nowrap">
                          <span className={`px-1.5 py-0.5 rounded font-mono font-bold text-[10px] ${
                            log.action === 'CREATE' ? 'bg-emerald-50 text-emerald-700' :
                            log.action === 'UPDATE' ? 'bg-blue-50 text-blue-700' :
                            log.action === 'DELETE' ? 'bg-red-50 text-red-700' : 'bg-slate-100 text-slate-700'
                          }`}>
                            {log.action}
                          </span>
                        </td>
                        <td className="p-2.5 text-slate-600 whitespace-nowrap">{log.worksheet}</td>
                        <td className="p-2.5 text-slate-700">{log.details}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {/* TAB 4: Clinic Details */}
          {activeTab === 'clinic' && (
            <div className="space-y-4">
              <div className="grid grid-cols-2 gap-4">
                <div>
                  <label className="block font-semibold text-slate-700 mb-1">
                    Clinic Display Name
                  </label>
                  <input
                    type="text"
                    disabled
                    value={settings?.clinicName || 'MK Eye Center'}
                    className="w-full px-3 py-2 bg-slate-100 border border-slate-300 rounded-lg text-slate-800 font-semibold"
                  />
                </div>

                <div>
                  <label className="block font-semibold text-slate-700 mb-1">
                    Default Currency
                  </label>
                  <input
                    type="text"
                    disabled
                    value={settings?.currency || 'GH₵'}
                    className="w-full px-3 py-2 bg-slate-100 border border-slate-300 rounded-lg text-slate-800 font-semibold"
                  />
                </div>
              </div>

              <div className="p-4 bg-slate-50 border border-slate-200 rounded-xl space-y-1">
                <p className="font-semibold text-slate-800">Current Operator</p>
                <p className="text-slate-600 text-[11px]">
                  Signed in as <strong>{currentUser.name}</strong> ({currentUser.role} role).
                </p>
              </div>
            </div>
          )}
        </div>

        {/* Footer */}
        <div className="px-6 py-3.5 bg-slate-50 border-t border-slate-200 flex items-center justify-end">
          <button
            type="button"
            onClick={onClose}
            className="px-4 py-2 font-semibold text-slate-800 bg-white border border-slate-300 rounded-lg hover:bg-slate-50 transition"
          >
            Close
          </button>
        </div>
      </div>
    </div>
  );
}
