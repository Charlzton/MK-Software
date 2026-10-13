'use client';

import React, { useState, useEffect } from 'react';
import { X, Calculator } from 'lucide-react';
import { PatientRecord, ColumnDefinition } from '@/lib/types';
import { KNOWN_DOCTORS, KNOWN_TEST_TYPES } from '@/lib/data/initialData';

interface PatientFormModalProps {
  isOpen: boolean;
  initialData?: PatientRecord | null;
  columns: ColumnDefinition[];
  defaultRecordedBy: string;
  onClose: () => void;
  onSubmit: (record: Partial<PatientRecord>) => Promise<void>;
}

export function PatientFormModal({
  isOpen,
  initialData,
  columns,
  defaultRecordedBy,
  onClose,
  onSubmit
}: PatientFormModalProps) {
  const [formData, setFormData] = useState<Partial<PatientRecord>>({});
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [error, setError] = useState<string | null>(null);

  // Initialize form
  useEffect(() => {
    if (!isOpen) return;
    const timer = setTimeout(() => {
      if (initialData) {
        setFormData({ ...initialData });
      } else {
        const today = new Date();
        const day = today.getDate();
        const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
        const dateStr = `${day}-${months[today.getMonth()]}-${today.getFullYear().toString().slice(-2)}`;

        setFormData({
          date: dateStr,
          recordedBy: defaultRecordedBy,
          phoneNo: '',
          examinedBy: 'Dr. Danso',
          name: '',
          dateOfBirth: '',
          testType: 'Standard',
          testFee: 100,
          dropsNo: 0,
          medication: 0,
          total: 100,
          glassTotal: '',
          paid: '',
          glassesBal: '',
          reviewDate: '',
          reviewTest: ''
        });
      }
      setError(null);
    }, 0);
    return () => clearTimeout(timer);
  }, [initialData, defaultRecordedBy, isOpen]);

  // Recalculate Total and Balance on fee change
  const handleChange = (field: keyof PatientRecord, value: unknown) => {
    setFormData((prev) => {
      const updated = { ...prev, [field]: value };

      if (field === 'testFee' || field === 'medication') {
        const testFeeNum = parseFloat(String(field === 'testFee' ? value : updated.testFee)) || 0;
        const medFeeNum = parseFloat(String(field === 'medication' ? value : updated.medication)) || 0;
        updated.total = testFeeNum + medFeeNum;
      }

      if (field === 'glassTotal' || field === 'paid') {
        const glassTot = parseFloat(String(field === 'glassTotal' ? value : updated.glassTotal));
        const paidVal = parseFloat(String(field === 'paid' ? value : updated.paid));
        if (!isNaN(glassTot) && !isNaN(paidVal)) {
          updated.glassesBal = Math.max(0, glassTot - paidVal);
        }
      }

      return updated;
    });
  };

  const handleCustomFieldChange = (key: string, value: string | number) => {
    setFormData((prev) => ({
      ...prev,
      customFields: {
        ...(prev.customFields || {}),
        [key]: value
      }
    }));
  };

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!formData.name || formData.name.trim() === '') {
      setError("Patient's Full Name is required.");
      return;
    }

    setIsSubmitting(true);
    setError(null);

    try {
      await onSubmit(formData);
      onClose();
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : 'Failed to save patient entry';
      setError(msg);
    } finally {
      setIsSubmitting(false);
    }
  };

  if (!isOpen) return null;

  const customCols = columns.filter((c) => c.isCustom);

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-3 sm:p-4 bg-slate-900/40 backdrop-blur-xs overflow-y-auto">
      <div className="relative w-full max-w-2xl bg-white rounded-xl shadow-2xl border border-slate-200 overflow-hidden my-6">
        {/* Modal Header */}
        <div className="flex items-center justify-between px-6 py-4 border-b border-slate-100 bg-slate-50/50">
          <div>
            <h2 className="text-base font-bold text-slate-900">
              {initialData ? 'Edit Patient Record' : 'Register New Patient'}
            </h2>
            <p className="text-xs text-slate-500">
              PATIENT&apos;S DATA · Form matches central clinical register structure
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

        {/* Modal Body */}
        <form onSubmit={handleSubmit} className="p-6 space-y-5">
          {error && (
            <div className="p-3 bg-red-50 border border-red-200 rounded-lg text-xs text-red-700">
              {error}
            </div>
          )}

          {/* Section 1: Demographics */}
          <div>
            <h3 className="text-xs font-bold uppercase tracking-wider text-slate-500 mb-3">
              1. Patient Demographics & Reception
            </h3>
            <div className="grid grid-cols-1 sm:grid-cols-3 gap-3.5">
              <div className="sm:col-span-2">
                <label className="block text-xs font-semibold text-slate-700 mb-1">
                  Patient Full Name <span className="text-red-500">*</span>
                </label>
                <input
                  type="text"
                  required
                  value={formData.name || ''}
                  onChange={(e) => handleChange('name', e.target.value)}
                  placeholder="e.g. Priscilla Adjoba Cudjoe"
                  className="w-full px-3 py-2 text-sm bg-white border border-slate-300 rounded-lg text-slate-900 focus:ring-2 focus:ring-slate-900 focus:outline-hidden"
                />
              </div>

              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">
                  Phone / Contact No
                </label>
                <input
                  type="text"
                  value={formData.phoneNo || ''}
                  onChange={(e) => handleChange('phoneNo', e.target.value)}
                  placeholder="e.g. 0244123456"
                  className="w-full px-3 py-2 text-sm bg-white border border-slate-300 rounded-lg text-slate-900 focus:ring-2 focus:ring-slate-900 focus:outline-hidden"
                />
              </div>

              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">
                  Registration Date
                </label>
                <input
                  type="text"
                  value={formData.date || ''}
                  onChange={(e) => handleChange('date', e.target.value)}
                  placeholder="e.g. 13-Oct-26"
                  className="w-full px-3 py-2 text-sm bg-white border border-slate-300 rounded-lg text-slate-900 focus:ring-2 focus:ring-slate-900 focus:outline-hidden"
                />
              </div>

              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">
                  Date of Birth / Age
                </label>
                <input
                  type="text"
                  value={formData.dateOfBirth || ''}
                  onChange={(e) => handleChange('dateOfBirth', e.target.value)}
                  placeholder="e.g. 15-08-1992 or 34"
                  className="w-full px-3 py-2 text-sm bg-white border border-slate-300 rounded-lg text-slate-900 focus:ring-2 focus:ring-slate-900 focus:outline-hidden"
                />
              </div>

              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">
                  Recorded By
                </label>
                <input
                  type="text"
                  value={formData.recordedBy || ''}
                  onChange={(e) => handleChange('recordedBy', e.target.value)}
                  placeholder="Staff member"
                  className="w-full px-3 py-2 text-sm bg-white border border-slate-300 rounded-lg text-slate-900 focus:ring-2 focus:ring-slate-900 focus:outline-hidden"
                />
              </div>
            </div>
          </div>

          {/* Section 2: Clinical Details */}
          <div className="pt-2 border-t border-slate-100">
            <h3 className="text-xs font-bold uppercase tracking-wider text-slate-500 mb-3">
              2. Clinical Examination
            </h3>
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3.5">
              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">
                  Examined By (Doctor)
                </label>
                <input
                  list="doctors-list"
                  type="text"
                  value={formData.examinedBy || ''}
                  onChange={(e) => handleChange('examinedBy', e.target.value)}
                  placeholder="Select or enter doctor"
                  className="w-full px-3 py-2 text-sm bg-white border border-slate-300 rounded-lg text-slate-900 focus:ring-2 focus:ring-slate-900 focus:outline-hidden"
                />
                <datalist id="doctors-list">
                  {KNOWN_DOCTORS.map((doc) => (
                    <option key={doc} value={doc} />
                  ))}
                </datalist>
              </div>

              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">
                  Test Type
                </label>
                <input
                  list="test-types-list"
                  type="text"
                  value={formData.testType || ''}
                  onChange={(e) => handleChange('testType', e.target.value)}
                  placeholder="Select or enter test type"
                  className="w-full px-3 py-2 text-sm bg-white border border-slate-300 rounded-lg text-slate-900 focus:ring-2 focus:ring-slate-900 focus:outline-hidden"
                />
                <datalist id="test-types-list">
                  {KNOWN_TEST_TYPES.map((t) => (
                    <option key={t} value={t} />
                  ))}
                </datalist>
              </div>
            </div>
          </div>

          {/* Section 3: Fees & Medication Billing */}
          <div className="pt-2 border-t border-slate-100">
            <div className="flex items-center justify-between mb-3">
              <h3 className="text-xs font-bold uppercase tracking-wider text-slate-500">
                3. Fees & Medication (GH₵)
              </h3>
              <span className="text-[11px] text-teal-700 flex items-center gap-1 font-medium">
                <Calculator className="w-3.5 h-3.5" />
                Live calculation enabled
              </span>
            </div>
            <div className="grid grid-cols-2 sm:grid-cols-4 gap-3.5">
              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">
                  Test Fee
                </label>
                <input
                  type="number"
                  step="any"
                  value={formData.testFee !== undefined ? formData.testFee : ''}
                  onChange={(e) => handleChange('testFee', e.target.value)}
                  placeholder="100"
                  className="w-full px-3 py-2 text-sm bg-white border border-slate-300 rounded-lg text-slate-900 tabular-nums focus:ring-2 focus:ring-slate-900 focus:outline-hidden"
                />
              </div>

              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">
                  Drops No
                </label>
                <input
                  type="number"
                  step="any"
                  value={formData.dropsNo !== undefined ? formData.dropsNo : ''}
                  onChange={(e) => handleChange('dropsNo', e.target.value)}
                  placeholder="0"
                  className="w-full px-3 py-2 text-sm bg-white border border-slate-300 rounded-lg text-slate-900 tabular-nums focus:ring-2 focus:ring-slate-900 focus:outline-hidden"
                />
              </div>

              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">
                  Medication Fee
                </label>
                <input
                  type="number"
                  step="any"
                  value={formData.medication !== undefined ? formData.medication : ''}
                  onChange={(e) => handleChange('medication', e.target.value)}
                  placeholder="0"
                  className="w-full px-3 py-2 text-sm bg-white border border-slate-300 rounded-lg text-slate-900 tabular-nums focus:ring-2 focus:ring-slate-900 focus:outline-hidden"
                />
              </div>

              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">
                  Total (Exam + Meds)
                </label>
                <input
                  type="number"
                  step="any"
                  value={formData.total !== undefined ? formData.total : ''}
                  onChange={(e) => handleChange('total', e.target.value)}
                  placeholder="Total"
                  className="w-full px-3 py-2 text-sm bg-slate-100 font-semibold border border-slate-300 rounded-lg text-slate-900 tabular-nums focus:ring-2 focus:ring-slate-900 focus:outline-hidden"
                />
              </div>
            </div>
          </div>

          {/* Section 4: Optical Dispensary & Balances */}
          <div className="pt-2 border-t border-slate-100">
            <h3 className="text-xs font-bold uppercase tracking-wider text-slate-500 mb-3">
              4. Optical Dispensary & Glass Balances
            </h3>
            <div className="grid grid-cols-1 sm:grid-cols-3 gap-3.5">
              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">
                  Glass Total
                </label>
                <input
                  type="text"
                  value={formData.glassTotal !== undefined ? formData.glassTotal : ''}
                  onChange={(e) => handleChange('glassTotal', e.target.value)}
                  placeholder="e.g. 1500"
                  className="w-full px-3 py-2 text-sm bg-white border border-slate-300 rounded-lg text-slate-900 tabular-nums focus:ring-2 focus:ring-slate-900 focus:outline-hidden"
                />
              </div>

              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">
                  Paid / Deposit
                </label>
                <input
                  type="text"
                  value={formData.paid !== undefined ? formData.paid : ''}
                  onChange={(e) => handleChange('paid', e.target.value)}
                  placeholder="e.g. 1000 or Paid"
                  className="w-full px-3 py-2 text-sm bg-white border border-slate-300 rounded-lg text-slate-900 tabular-nums focus:ring-2 focus:ring-slate-900 focus:outline-hidden"
                />
              </div>

              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">
                  Glasses BAL (Remaining)
                </label>
                <input
                  type="text"
                  value={formData.glassesBal !== undefined ? formData.glassesBal : ''}
                  onChange={(e) => handleChange('glassesBal', e.target.value)}
                  placeholder="Calculated balance"
                  className="w-full px-3 py-2 text-sm bg-slate-100 border border-slate-300 rounded-lg text-slate-900 tabular-nums font-semibold focus:ring-2 focus:ring-slate-900 focus:outline-hidden"
                />
              </div>
            </div>
          </div>

          {/* Section 5: Review Follow-up */}
          <div className="pt-2 border-t border-slate-100">
            <h3 className="text-xs font-bold uppercase tracking-wider text-slate-500 mb-3">
              5. Review & Follow-up Appointment
            </h3>
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3.5">
              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">
                  Review Date
                </label>
                <input
                  type="text"
                  value={formData.reviewDate || ''}
                  onChange={(e) => handleChange('reviewDate', e.target.value)}
                  placeholder="e.g. 26-05-2025"
                  className="w-full px-3 py-2 text-sm bg-white border border-slate-300 rounded-lg text-slate-900 focus:ring-2 focus:ring-slate-900 focus:outline-hidden"
                />
              </div>

              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">
                  Review Test / Notes
                </label>
                <input
                  type="text"
                  value={formData.reviewTest || ''}
                  onChange={(e) => handleChange('reviewTest', e.target.value)}
                  placeholder="e.g. IOP or VFT"
                  className="w-full px-3 py-2 text-sm bg-white border border-slate-300 rounded-lg text-slate-900 focus:ring-2 focus:ring-slate-900 focus:outline-hidden"
                />
              </div>
            </div>
          </div>

          {/* Custom Columns if any */}
          {customCols.length > 0 && (
            <div className="pt-2 border-t border-slate-100">
              <h3 className="text-xs font-bold uppercase tracking-wider text-slate-500 mb-3">
                6. Custom Clinic Fields
              </h3>
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3.5">
                {customCols.map((col) => (
                  <div key={col.key}>
                    <label className="block text-xs font-semibold text-slate-700 mb-1">
                      {col.label}
                    </label>
                    <input
                      type="text"
                      value={formData.customFields?.[col.key] || ''}
                      onChange={(e) => handleCustomFieldChange(col.key, e.target.value)}
                      className="w-full px-3 py-2 text-sm bg-white border border-slate-300 rounded-lg text-slate-900 focus:ring-2 focus:ring-slate-900 focus:outline-hidden"
                    />
                  </div>
                ))}
              </div>
            </div>
          )}

          {/* Modal Actions */}
          <div className="pt-4 border-t border-slate-100 flex items-center justify-end gap-3">
            <button
              type="button"
              onClick={onClose}
              disabled={isSubmitting}
              className="px-4 py-2 text-xs font-semibold text-slate-700 bg-white border border-slate-300 rounded-lg hover:bg-slate-50 transition"
            >
              Cancel
            </button>
            <button
              type="submit"
              disabled={isSubmitting}
              className="px-5 py-2 text-xs font-semibold text-white bg-slate-900 hover:bg-slate-800 rounded-lg shadow-xs transition flex items-center gap-2"
            >
              {isSubmitting && (
                <span className="w-3.5 h-3.5 border-2 border-white/30 border-t-white rounded-full animate-spin" />
              )}
              <span>{initialData ? 'Save Changes' : 'Register Patient'}</span>
            </button>
          </div>
        </form>
      </div>
    </div>
  );
}
