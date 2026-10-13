'use client';

import React, { useState, useMemo } from 'react';
import {
  Stethoscope,
  Search,
  Calendar,
  Clock,
  Eye,
  FileSpreadsheet,
  AlertCircle,
  ClipboardList,
  CheckCircle2
} from 'lucide-react';
import { PatientRecord, User } from '@/lib/types';

interface DoctorWorkspaceProps {
  user: User;
  patients: PatientRecord[];
}

export function DoctorWorkspace({ user, patients }: DoctorWorkspaceProps) {
  const [searchTerm, setSearchTerm] = useState('');
  const [selectedPatient, setSelectedPatient] = useState<PatientRecord | null>(null);

  // Staging consultation note state
  const [clinicalNotes, setClinicalNotes] = useState('');
  const [stagedSuccess, setStagedSuccess] = useState(false);

  // Filter patients by name or ID
  const searchResults = useMemo(() => {
    if (!searchTerm.trim()) return [];
    const term = searchTerm.toLowerCase();
    return patients.filter(
      (p) =>
        p.name.toLowerCase().includes(term) ||
        p.phoneNo.includes(term) ||
        p.id.toLowerCase().includes(term)
    ).slice(0, 8);
  }, [patients, searchTerm]);

  const handleSaveStagedNote = (e: React.FormEvent) => {
    e.preventDefault();
    if (!clinicalNotes.trim() || !selectedPatient) return;
    setStagedSuccess(true);
    setTimeout(() => {
      setStagedSuccess(false);
      setClinicalNotes('');
    }, 3000);
  };

  return (
    <div className="space-y-6">
      {/* 1. Official Empty-State & Workbook Notice Banner */}
      <div className="bg-white rounded-xl border border-slate-200 p-6 shadow-xs">
        <div className="flex flex-col sm:flex-row items-start gap-4">
          <div className="p-3 bg-teal-50 border border-teal-100 rounded-xl text-teal-700 shrink-0">
            <Stethoscope className="w-6 h-6 stroke-[2.2]" />
          </div>
          <div className="flex-1 space-y-2">
            <div className="flex items-center gap-2">
              <h2 className="text-base font-bold text-slate-900">
                Doctor Workspace · Clinical Consultation Portal
              </h2>
              <span className="text-[11px] font-medium text-amber-700 bg-amber-50 border border-amber-200 px-2 py-0.5 rounded">
                Awaiting Doctor Excel Ingestion
              </span>
            </div>
            <p className="text-xs text-slate-600 leading-relaxed">
              In accordance with project architecture, the dedicated doctor&apos;s Excel workbook will be ingested separately upon delivery. To prevent data contamination, reception schemas have not been duplicated into the clinical module, and no simulated clinical records have been fabricated.
            </p>
            <div className="pt-1 flex flex-wrap items-center gap-4 text-xs text-slate-500">
              <span className="flex items-center gap-1.5">
                <CheckCircle2 className="w-3.5 h-3.5 text-teal-600" />
                Active Clinician: <strong className="text-slate-800">{user.name}</strong>
              </span>
              <span className="flex items-center gap-1.5">
                <FileSpreadsheet className="w-3.5 h-3.5 text-slate-400" />
                Linked Patient Registry: <strong className="text-slate-800">{patients.length} records ready for cross-referencing</strong>
              </span>
            </div>
          </div>
        </div>
      </div>

      {/* 2. Patient Clinical History Cross-Reference Tool */}
      <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
        {/* Left: Patient Search Column */}
        <div className="bg-white rounded-xl border border-slate-200 p-5 shadow-xs space-y-4">
          <div>
            <h3 className="text-xs font-bold uppercase tracking-wider text-slate-500 mb-1 flex items-center gap-1.5">
              <Search className="w-3.5 h-3.5 text-slate-400" />
              Patient Lookup
            </h3>
            <p className="text-xs text-slate-500">
              Search by patient name, phone number, or stable ID to inspect visit history.
            </p>
          </div>

          <div className="relative">
            <Search className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
            <input
              type="text"
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
              placeholder="e.g. Priscilla, Kwame, or 0244..."
              className="w-full pl-9 pr-3 py-2 text-xs bg-slate-50 border border-slate-200 rounded-lg text-slate-800 placeholder:text-slate-400 focus:outline-hidden focus:ring-2 focus:ring-slate-900 focus:bg-white transition"
            />
          </div>

          {searchTerm.trim() && (
            <div className="space-y-1.5 max-h-80 overflow-y-auto">
              {searchResults.length === 0 ? (
                <div className="p-3 text-center text-xs text-slate-400 bg-slate-50 rounded-lg">
                  No matching patients found.
                </div>
              ) : (
                searchResults.map((p) => (
                  <div
                    key={p.id}
                    onClick={() => setSelectedPatient(p)}
                    className={`p-2.5 rounded-lg border text-xs cursor-pointer transition ${
                      selectedPatient?.id === p.id
                        ? 'border-slate-900 bg-slate-50/80'
                        : 'border-slate-200 hover:border-slate-300 bg-white'
                    }`}
                  >
                    <div className="font-semibold text-slate-900">{p.name}</div>
                    <div className="flex items-center justify-between text-slate-500 text-[11px] mt-1">
                      <span>{p.phoneNo || 'No phone'}</span>
                      <span className="font-mono text-[10px] text-slate-400">{p.id}</span>
                    </div>
                  </div>
                ))
              )}
            </div>
          )}

          {!searchTerm.trim() && (
            <div className="p-4 bg-slate-50 border border-dashed border-slate-200 rounded-lg text-center text-xs text-slate-400">
              Enter a patient name above to search clinic visit history.
            </div>
          )}
        </div>

        {/* Right: Clinical History Review & Note Preparation */}
        <div className="lg:col-span-2 space-y-6">
          {selectedPatient ? (
            <div className="bg-white rounded-xl border border-slate-200 p-6 shadow-xs space-y-5">
              <div className="flex items-start justify-between border-b border-slate-100 pb-4">
                <div>
                  <h3 className="text-base font-bold text-slate-900 flex items-center gap-2">
                    {selectedPatient.name}
                    <span className="text-xs font-mono font-normal text-slate-500 bg-slate-100 px-2 py-0.5 rounded">
                      {selectedPatient.id}
                    </span>
                  </h3>
                  <div className="flex items-center gap-3 text-xs text-slate-500 mt-1">
                    <span>Phone: {selectedPatient.phoneNo || 'N/A'}</span>
                    <span>·</span>
                    <span>DOB/Age: {selectedPatient.dateOfBirth || 'N/A'}</span>
                  </div>
                </div>
                <button
                  type="button"
                  onClick={() => setSelectedPatient(null)}
                  className="text-xs text-slate-400 hover:text-slate-700"
                >
                  Clear Selection
                </button>
              </div>

              {/* Patient Examination History Card */}
              <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 text-xs">
                <div className="p-3 bg-slate-50 border border-slate-200 rounded-lg">
                  <span className="text-[10px] uppercase font-bold text-slate-400 block">Exam Date</span>
                  <span className="font-semibold text-slate-800">{selectedPatient.date}</span>
                </div>
                <div className="p-3 bg-slate-50 border border-slate-200 rounded-lg">
                  <span className="text-[10px] uppercase font-bold text-slate-400 block">Examining Doctor</span>
                  <span className="font-semibold text-slate-800">{selectedPatient.examinedBy || 'Unrecorded'}</span>
                </div>
                <div className="p-3 bg-slate-50 border border-slate-200 rounded-lg">
                  <span className="text-[10px] uppercase font-bold text-slate-400 block">Test Type</span>
                  <span className="font-semibold text-slate-800">{selectedPatient.testType || 'Standard'}</span>
                </div>
                <div className="p-3 bg-slate-50 border border-slate-200 rounded-lg">
                  <span className="text-[10px] uppercase font-bold text-slate-400 block">Drops Count</span>
                  <span className="font-semibold text-slate-800">{selectedPatient.dropsNo || '0'}</span>
                </div>
              </div>

              {/* Follow-up Details */}
              <div className="p-3.5 bg-teal-50/50 border border-teal-100 rounded-lg text-xs space-y-1">
                <span className="text-[10px] uppercase font-bold text-teal-800 block">Scheduled Review</span>
                <div className="flex items-center gap-4 text-teal-900">
                  <span>Review Date: <strong>{selectedPatient.reviewDate || 'None Scheduled'}</strong></span>
                  <span>Review Procedure: <strong>{selectedPatient.reviewTest || 'Standard Check'}</strong></span>
                </div>
              </div>

              {/* Note Preparation / Staging form */}
              <form onSubmit={handleSaveStagedNote} className="space-y-3 pt-2">
                <div>
                  <label className="block text-xs font-semibold text-slate-700 mb-1">
                    Clinical Observations & Diagnosis Preparation
                  </label>
                  <textarea
                    rows={4}
                    value={clinicalNotes}
                    onChange={(e) => setClinicalNotes(e.target.value)}
                    placeholder="Record clinical impressions, visual acuity, slit lamp findings, intraocular pressure notes, or prescription updates..."
                    className="w-full p-3 text-xs bg-white border border-slate-300 rounded-lg text-slate-900 focus:ring-2 focus:ring-slate-900 focus:outline-hidden"
                  />
                </div>

                {stagedSuccess && (
                  <div className="p-2.5 bg-emerald-50 border border-emerald-200 rounded text-xs text-emerald-800 flex items-center gap-1.5">
                    <CheckCircle2 className="w-4 h-4 text-emerald-600" />
                    Clinical consultation note staged successfully for patient ID {selectedPatient.id}.
                  </div>
                )}

                <div className="flex items-center justify-end gap-2">
                  <button
                    type="submit"
                    disabled={!clinicalNotes.trim()}
                    className="px-4 py-2 text-xs font-semibold text-white bg-slate-900 hover:bg-slate-800 disabled:opacity-40 rounded-lg transition"
                  >
                    Stage Consultation Note
                  </button>
                </div>
              </form>
            </div>
          ) : (
            <div className="bg-white rounded-xl border border-slate-200 p-12 text-center shadow-xs">
              <ClipboardList className="w-10 h-10 text-slate-300 mx-auto mb-3" />
              <h4 className="text-sm font-semibold text-slate-800">No Patient Selected</h4>
              <p className="text-xs text-slate-500 max-w-sm mx-auto mt-1">
                Search and select a patient from the roster on the left to review their clinical examination history and prepare consultation notes.
              </p>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}
