'use client';

import React, { useState } from 'react';
import {
  Users,
  Car,
  Pill,
  Archive,
  AlertCircle,
  TrendingUp,
  Receipt
} from 'lucide-react';
import { PatientRecord, DVLARecord, MedsRecord, Sheet2Record, ColumnDefinition, User } from '@/lib/types';
import { PatientTable } from './PatientTable';
import { DVLATable } from './DVLATable';
import { MedsTable } from './MedsTable';
import { Sheet2Table } from './Sheet2Table';
import { PatientFormModal } from './PatientFormModal';
import { ConfirmDialog } from './ConfirmDialog';

interface ReceptionWorkspaceProps {
  currentUser: User;
  patients: PatientRecord[];
  patientColumns: ColumnDefinition[];
  dvla: DVLARecord[];
  dvlaColumns: ColumnDefinition[];
  meds: MedsRecord[];
  sheet2: Sheet2Record[];
  isLoading: boolean;
  onRefreshData: () => void;
  onSavePatient: (record: Partial<PatientRecord>) => Promise<void>;
  onDeletePatient: (id: string) => Promise<void>;
  onSaveDVLA: (record: Omit<DVLARecord, 'id'>) => Promise<void>;
  onUpdateDVLA: (id: string, updates: Partial<DVLARecord>) => Promise<void>;
  onDeleteDVLA: (id: string) => Promise<void>;
}

export function ReceptionWorkspace({
  currentUser,
  patients,
  patientColumns,
  dvla,
  dvlaColumns,
  meds,
  sheet2,
  isLoading,
  onRefreshData,
  onSavePatient,
  onDeletePatient,
  onSaveDVLA,
  onUpdateDVLA,
  onDeleteDVLA
}: ReceptionWorkspaceProps) {
  const [subTab, setSubTab] = useState<'patients' | 'dvla' | 'meds' | 'sheet2'>('patients');

  // Patient Modal state
  const [isPatientModalOpen, setIsPatientModalOpen] = useState(false);
  const [selectedPatientForEdit, setSelectedPatientForEdit] = useState<PatientRecord | null>(null);

  // Delete Confirmation state
  const [deleteTarget, setDeleteTarget] = useState<{
    id: string;
    name: string;
    type: 'patient' | 'dvla';
  } | null>(null);
  const [isDeleting, setIsDeleting] = useState(false);

  // Calculate real summaries from real data
  const totalPatientsCount = patients.length;
  const totalDVLACount = dvla.length;
  const pendingGlassBalancesCount = patients.filter((p) => Number(p.glassesBal) > 0).length;

  const handleOpenAddPatient = () => {
    setSelectedPatientForEdit(null);
    setIsPatientModalOpen(true);
  };

  const handleOpenEditPatient = (p: PatientRecord) => {
    setSelectedPatientForEdit(p);
    setIsPatientModalOpen(true);
  };

  const handleConfirmDelete = async () => {
    if (!deleteTarget) return;
    setIsDeleting(true);
    try {
      if (deleteTarget.type === 'patient') {
        await onDeletePatient(deleteTarget.id);
      } else {
        await onDeleteDVLA(deleteTarget.id);
      }
      setDeleteTarget(null);
    } catch (err) {
      console.error('Delete failed:', err);
    } finally {
      setIsDeleting(false);
    }
  };

  return (
    <div className="space-y-5">
      {/* Clinic Operational Metric Bar (Quiet, anti-slop, clean typography) */}
      <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
        <div className="bg-white p-3.5 rounded-xl border border-slate-200 shadow-xs flex items-center justify-between">
          <div>
            <span className="text-[11px] font-semibold uppercase tracking-wider text-slate-500 block">
              Patient Register Entries
            </span>
            <span className="text-xl font-bold font-mono tabular-nums text-slate-900 mt-0.5 block">
              {totalPatientsCount}
            </span>
          </div>
          <div className="p-2.5 rounded-lg bg-slate-100 text-slate-700">
            <Users className="w-4 h-4" />
          </div>
        </div>

        <div className="bg-white p-3.5 rounded-xl border border-slate-200 shadow-xs flex items-center justify-between">
          <div>
            <span className="text-[11px] font-semibold uppercase tracking-wider text-slate-500 block">
              DVLA License Exams
            </span>
            <span className="text-xl font-bold font-mono tabular-nums text-slate-900 mt-0.5 block">
              {totalDVLACount}
            </span>
          </div>
          <div className="p-2.5 rounded-lg bg-slate-100 text-slate-700">
            <Car className="w-4 h-4" />
          </div>
        </div>

        <div className="bg-white p-3.5 rounded-xl border border-slate-200 shadow-xs flex items-center justify-between">
          <div>
            <span className="text-[11px] font-semibold uppercase tracking-wider text-slate-500 block">
              Unsettled Optical Balances
            </span>
            <span className="text-xl font-bold font-mono tabular-nums text-amber-900 mt-0.5 block">
              {pendingGlassBalancesCount}
            </span>
          </div>
          <div className="p-2.5 rounded-lg bg-amber-50 text-amber-700">
            <Receipt className="w-4 h-4" />
          </div>
        </div>
      </div>

      {/* Sub-Worksheet Tabs (clean segmented buttons) */}
      <div className="flex flex-wrap items-center justify-between gap-3 border-b border-slate-200 pb-1">
        <div className="flex items-center gap-1">
          <button
            type="button"
            onClick={() => setSubTab('patients')}
            className={`px-3.5 py-2 text-xs font-semibold rounded-lg transition flex items-center gap-2 ${
              subTab === 'patients'
                ? 'bg-slate-900 text-white shadow-xs'
                : 'text-slate-600 hover:text-slate-900 hover:bg-slate-100'
            }`}
          >
            <Users className="w-3.5 h-3.5" />
            <span>PATIENT&apos;S DATA (16 Cols)</span>
          </button>

          <button
            type="button"
            onClick={() => setSubTab('dvla')}
            className={`px-3.5 py-2 text-xs font-semibold rounded-lg transition flex items-center gap-2 ${
              subTab === 'dvla'
                ? 'bg-slate-900 text-white shadow-xs'
                : 'text-slate-600 hover:text-slate-900 hover:bg-slate-100'
            }`}
          >
            <Car className="w-3.5 h-3.5" />
            <span>DVLA Register</span>
          </button>

          <button
            type="button"
            onClick={() => setSubTab('meds')}
            className={`px-3.5 py-2 text-xs font-semibold rounded-lg transition flex items-center gap-2 ${
              subTab === 'meds'
                ? 'bg-slate-900 text-white shadow-xs'
                : 'text-slate-600 hover:text-slate-900 hover:bg-slate-100'
            }`}
          >
            <Pill className="w-3.5 h-3.5" />
            <span>MEDS Reference</span>
          </button>

          <button
            type="button"
            onClick={() => setSubTab('sheet2')}
            className={`px-3.5 py-2 text-xs font-semibold rounded-lg transition flex items-center gap-2 ${
              subTab === 'sheet2'
                ? 'bg-slate-900 text-white shadow-xs'
                : 'text-slate-600 hover:text-slate-900 hover:bg-slate-100'
            }`}
          >
            <Archive className="w-3.5 h-3.5" />
            <span>Sheet2 (Unclassified)</span>
          </button>
        </div>

        <div className="text-xs text-slate-500 font-mono">
          MK Eye Center · Station: {currentUser.name}
        </div>
      </div>

      {/* Main Content Area according to active sub-tab */}
      {subTab === 'patients' && (
        <PatientTable
          records={patients}
          columns={patientColumns}
          isLoading={isLoading}
          onAddRecord={handleOpenAddPatient}
          onEditRecord={handleOpenEditPatient}
          onDeleteRecord={(id, name) => setDeleteTarget({ id, name, type: 'patient' })}
          onRefresh={onRefreshData}
        />
      )}

      {subTab === 'dvla' && (
        <DVLATable
          records={dvla}
          columns={dvlaColumns}
          isLoading={isLoading}
          defaultRecordedBy={currentUser.name}
          onAddRecord={onSaveDVLA}
          onEditRecord={onUpdateDVLA}
          onDeleteRecord={(id, name) => setDeleteTarget({ id, name, type: 'dvla' })}
          onRefresh={onRefreshData}
        />
      )}

      {subTab === 'meds' && <MedsTable records={meds} isLoading={isLoading} />}

      {subTab === 'sheet2' && <Sheet2Table records={sheet2} isLoading={isLoading} />}

      {/* Patient Add / Edit Form Modal */}
      <PatientFormModal
        isOpen={isPatientModalOpen}
        initialData={selectedPatientForEdit}
        columns={patientColumns}
        defaultRecordedBy={currentUser.name}
        onClose={() => setIsPatientModalOpen(false)}
        onSubmit={onSavePatient}
      />

      {/* Confirmation Dialog for Record Deletions */}
      <ConfirmDialog
        isOpen={!!deleteTarget}
        title={`Delete ${deleteTarget?.type === 'patient' ? 'Patient' : 'DVLA'} Record`}
        message={`Are you sure you want to permanently delete the record for "${deleteTarget?.name}"? This action will be logged in the clinic audit trail and synchronized.`}
        confirmLabel="Confirm Delete"
        confirmVariant="danger"
        isLoading={isDeleting}
        onConfirm={handleConfirmDelete}
        onCancel={() => setDeleteTarget(null)}
      />
    </div>
  );
}
