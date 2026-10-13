'use client';

import React, { useState, useMemo } from 'react';
import {
  Search,
  Plus,
  ArrowUpDown,
  Download,
  Edit2,
  Trash2,
  X
} from 'lucide-react';
import { DVLARecord, ColumnDefinition } from '@/lib/types';
import { KNOWN_DOCTORS } from '@/lib/data/initialData';

interface DVLATableProps {
  records: DVLARecord[];
  columns: ColumnDefinition[];
  isLoading: boolean;
  defaultRecordedBy: string;
  onAddRecord: (record: Omit<DVLARecord, 'id'>) => Promise<void>;
  onEditRecord: (id: string, updates: Partial<DVLARecord>) => Promise<void>;
  onDeleteRecord: (id: string, name: string) => void;
  onRefresh: () => void;
}

export function DVLATable({
  records,
  columns,
  isLoading,
  defaultRecordedBy,
  onAddRecord,
  onEditRecord,
  onDeleteRecord,
  onRefresh
}: DVLATableProps) {
  const [searchTerm, setSearchTerm] = useState('');
  const [dicFilter, setDicFilter] = useState('');
  const [sortField, setSortField] = useState<keyof DVLARecord>('date');
  const [sortAsc, setSortAsc] = useState(false);
  const [currentPage, setCurrentPage] = useState(1);
  const [pageSize, setPageSize] = useState(25);

  // Modal state for Add/Edit
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [editingRecord, setEditingRecord] = useState<DVLARecord | null>(null);
  const [formData, setFormData] = useState<Partial<DVLARecord>>({});
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [formError, setFormError] = useState<string | null>(null);

  const openAddModal = () => {
    const today = new Date();
    const day = today.getDate();
    const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    const dateStr = `${day}-${months[today.getMonth()]}-${today.getFullYear().toString().slice(-2)}`;

    setEditingRecord(null);
    setFormData({
      date: dateStr,
      recordedBy: defaultRecordedBy,
      contact: '',
      dic: 'Dr Danso',
      patientName: '',
      newFee: 100,
      renewalFee: 0,
      notes: ''
    });
    setFormError(null);
    setIsModalOpen(true);
  };

  const openEditModal = (rec: DVLARecord) => {
    setEditingRecord(rec);
    setFormData({ ...rec });
    setFormError(null);
    setIsModalOpen(true);
  };

  const handleFormSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!formData.patientName || formData.patientName.trim() === '') {
      setFormError("Applicant / Patient's Name is required.");
      return;
    }

    setIsSubmitting(true);
    setFormError(null);

    try {
      if (editingRecord) {
        await onEditRecord(editingRecord.id, formData);
      } else {
        await onAddRecord(formData as Omit<DVLARecord, 'id'>);
      }
      setIsModalOpen(false);
    } catch (err: unknown) {
      setFormError(err instanceof Error ? err.message : 'Failed to save DVLA record');
    } finally {
      setIsSubmitting(false);
    }
  };

  const handleSort = (field: keyof DVLARecord) => {
    if (sortField === field) {
      setSortAsc(!sortAsc);
    } else {
      setSortField(field);
      setSortAsc(true);
    }
  };

  const filteredRecords = useMemo(() => {
    return records.filter((r) => {
      const matchSearch =
        !searchTerm ||
        r.patientName.toLowerCase().includes(searchTerm.toLowerCase()) ||
        r.contact.toLowerCase().includes(searchTerm.toLowerCase()) ||
        (r.recordedBy && r.recordedBy.toLowerCase().includes(searchTerm.toLowerCase()));

      const matchDic =
        !dicFilter ||
        (r.dic && r.dic.toLowerCase().includes(dicFilter.toLowerCase()));

      return matchSearch && matchDic;
    });
  }, [records, searchTerm, dicFilter]);

  const sortedRecords = useMemo(() => {
    return [...filteredRecords].sort((a, b) => {
      const aVal = a[sortField] ?? '';
      const bVal = b[sortField] ?? '';

      if (typeof aVal === 'number' && typeof bVal === 'number') {
        return sortAsc ? aVal - bVal : bVal - aVal;
      }

      const strA = String(aVal).toLowerCase();
      const strB = String(bVal).toLowerCase();
      return sortAsc ? strA.localeCompare(strB) : strB.localeCompare(strA);
    });
  }, [filteredRecords, sortField, sortAsc]);

  const totalPages = Math.ceil(sortedRecords.length / pageSize) || 1;
  const paginatedRecords = useMemo(() => {
    const start = (currentPage - 1) * pageSize;
    return sortedRecords.slice(start, start + pageSize);
  }, [sortedRecords, currentPage, pageSize]);

  const handleExportCSV = () => {
    const headers = 'Date,Recorded_By,Contact,DIC,Patient Name,New Fee,Renewal Fee,Notes';
    const rows = sortedRecords.map((r) =>
      `"${r.date}","${r.recordedBy}","${r.contact}","${r.dic}","${r.patientName}","${r.newFee}","${r.renewalFee}","${r.notes || ''}"`
    );
    const csvContent = [headers, ...rows].join('\n');
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `MK_DVLA_Register_${new Date().toISOString().slice(0, 10)}.csv`;
    link.click();
  };

  return (
    <div className="space-y-4">
      {/* Toolbar */}
      <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-xs flex flex-col lg:flex-row lg:items-center justify-between gap-3">
        <div className="flex flex-wrap items-center gap-2.5 flex-1">
          <div className="relative min-w-[240px] flex-1 sm:flex-initial">
            <Search className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
            <input
              type="text"
              value={searchTerm}
              onChange={(e) => {
                setSearchTerm(e.target.value);
                setCurrentPage(1);
              }}
              placeholder="Search DVLA applicants by name, phone, or staff..."
              className="w-full pl-9 pr-3 py-1.5 text-xs bg-slate-50 border border-slate-200 rounded-lg text-slate-800 placeholder:text-slate-400 focus:outline-hidden focus:ring-2 focus:ring-slate-900 focus:bg-white transition"
            />
          </div>

          <div className="relative">
            <select
              value={dicFilter}
              onChange={(e) => {
                setDicFilter(e.target.value);
                setCurrentPage(1);
              }}
              className="text-xs bg-slate-50 border border-slate-200 rounded-lg px-2.5 py-1.5 text-slate-700 focus:outline-hidden focus:ring-2 focus:ring-slate-900"
            >
              <option value="">All Doctors (DIC)</option>
              {KNOWN_DOCTORS.map((doc) => (
                <option key={doc} value={doc}>
                  {doc}
                </option>
              ))}
            </select>
          </div>

          {(searchTerm || dicFilter) && (
            <button
              type="button"
              onClick={() => {
                setSearchTerm('');
                setDicFilter('');
                setCurrentPage(1);
              }}
              className="text-xs text-slate-500 hover:text-slate-900 underline"
            >
              Reset
            </button>
          )}
        </div>

        <div className="flex items-center gap-2 shrink-0">
          <button
            type="button"
            onClick={onRefresh}
            className="px-2.5 py-1.5 text-xs font-medium text-slate-700 bg-white border border-slate-200 rounded-lg hover:bg-slate-50 transition"
          >
            Refresh
          </button>
          <button
            type="button"
            onClick={handleExportCSV}
            className="px-2.5 py-1.5 text-xs font-medium text-slate-700 bg-white border border-slate-200 rounded-lg hover:bg-slate-50 transition flex items-center gap-1.5"
          >
            <Download className="w-3.5 h-3.5 text-slate-500" />
            <span>Export</span>
          </button>
          <button
            type="button"
            onClick={openAddModal}
            className="px-3.5 py-1.5 text-xs font-semibold text-white bg-slate-900 hover:bg-slate-800 rounded-lg shadow-xs transition flex items-center gap-1.5"
          >
            <Plus className="w-3.5 h-3.5" />
            <span>Add DVLA Entry</span>
          </button>
        </div>
      </div>

      {/* DVLA Grid */}
      <div className="bg-white rounded-xl border border-slate-200 shadow-xs overflow-hidden">
        <div className="overflow-x-auto min-h-[360px]">
          <table className="w-full border-collapse text-left text-xs">
            <thead>
              <tr className="bg-slate-50/90 border-b border-slate-200 text-slate-600 select-none">
                <th className="py-2.5 px-3 font-semibold w-12 text-center text-slate-400">#</th>
                <th
                  onClick={() => handleSort('date')}
                  className="py-2.5 px-3 font-semibold whitespace-nowrap cursor-pointer hover:bg-slate-100/70"
                >
                  <div className="flex items-center gap-1">
                    <span>Date</span>
                    <ArrowUpDown className="w-3 h-3 text-slate-400" />
                  </div>
                </th>
                <th
                  onClick={() => handleSort('recordedBy')}
                  className="py-2.5 px-3 font-semibold whitespace-nowrap cursor-pointer hover:bg-slate-100/70"
                >
                  <div className="flex items-center gap-1">
                    <span>Recorded_By (Col 2)</span>
                    <ArrowUpDown className="w-3 h-3 text-slate-400" />
                  </div>
                </th>
                <th className="py-2.5 px-3 font-semibold whitespace-nowrap">Contact</th>
                <th
                  onClick={() => handleSort('dic')}
                  className="py-2.5 px-3 font-semibold whitespace-nowrap cursor-pointer hover:bg-slate-100/70"
                >
                  <div className="flex items-center gap-1">
                    <span>DIC</span>
                    <ArrowUpDown className="w-3 h-3 text-slate-400" />
                  </div>
                </th>
                <th
                  onClick={() => handleSort('patientName')}
                  className="py-2.5 px-3 font-semibold whitespace-nowrap cursor-pointer hover:bg-slate-100/70"
                >
                  <div className="flex items-center gap-1">
                    <span>Patient&apos;s Name</span>
                    <ArrowUpDown className="w-3 h-3 text-slate-400" />
                  </div>
                </th>
                <th className="py-2.5 px-3 font-semibold whitespace-nowrap text-right">New (GH₵)</th>
                <th className="py-2.5 px-3 font-semibold whitespace-nowrap text-right">Renewal (GH₵)</th>
                <th className="py-2.5 px-3 font-semibold whitespace-nowrap">Notes</th>
                <th className="py-2.5 px-3 font-semibold text-right pr-4">Actions</th>
              </tr>
            </thead>

            <tbody className="divide-y divide-slate-100">
              {isLoading ? (
                <tr>
                  <td colSpan={10} className="py-16 text-center text-slate-400">
                    <span className="text-xs">Loading DVLA register...</span>
                  </td>
                </tr>
              ) : paginatedRecords.length === 0 ? (
                <tr>
                  <td colSpan={10} className="py-16 text-center text-slate-500">
                    <p className="font-semibold text-slate-800">No DVLA records found</p>
                    <p className="text-xs text-slate-500">Click &quot;Add DVLA Entry&quot; to log a new record.</p>
                  </td>
                </tr>
              ) : (
                paginatedRecords.map((record, index) => {
                  const rowNumber = (currentPage - 1) * pageSize + index + 1;
                  return (
                    <tr
                      key={record.id}
                      className="hover:bg-slate-50/80 transition-colors group cursor-pointer"
                      onClick={() => openEditModal(record)}
                    >
                      <td className="py-2 px-3 text-center text-slate-400 tabular-nums font-mono text-[11px]">
                        {rowNumber}
                      </td>
                      <td className="py-2 px-3 text-slate-700 whitespace-nowrap">{record.date}</td>
                      <td className="py-2 px-3 text-slate-700 whitespace-nowrap">{record.recordedBy || '-'}</td>
                      <td className="py-2 px-3 text-slate-700 whitespace-nowrap font-mono tabular-nums">
                        {record.contact || '-'}
                      </td>
                      <td className="py-2 px-3 text-slate-800 font-medium whitespace-nowrap">{record.dic || '-'}</td>
                      <td className="py-2 px-3 font-semibold text-slate-900 whitespace-nowrap">
                        {record.patientName}
                      </td>
                      <td className="py-2 px-3 text-slate-800 font-mono tabular-nums text-right whitespace-nowrap">
                        {record.newFee !== '' ? `GH₵ ${record.newFee}` : '-'}
                      </td>
                      <td className="py-2 px-3 text-slate-800 font-mono tabular-nums text-right whitespace-nowrap">
                        {record.renewalFee !== '' ? `GH₵ ${record.renewalFee}` : '-'}
                      </td>
                      <td className="py-2 px-3 text-slate-500 whitespace-nowrap">{record.notes || '-'}</td>
                      <td
                        className="py-2 px-3 text-right pr-4 whitespace-nowrap"
                        onClick={(e) => e.stopPropagation()}
                      >
                        <div className="flex items-center justify-end gap-1 opacity-80 group-hover:opacity-100 transition">
                          <button
                            type="button"
                            onClick={() => openEditModal(record)}
                            className="p-1 text-slate-400 hover:text-slate-700 hover:bg-slate-100 rounded transition"
                            title="Edit DVLA Record"
                          >
                            <Edit2 className="w-3.5 h-3.5" />
                          </button>
                          <button
                            type="button"
                            onClick={() => onDeleteRecord(record.id, record.patientName)}
                            className="p-1 text-slate-400 hover:text-red-600 hover:bg-red-50 rounded transition"
                            title="Delete Record"
                          >
                            <Trash2 className="w-3.5 h-3.5" />
                          </button>
                        </div>
                      </td>
                    </tr>
                  );
                })
              )}
            </tbody>
          </table>
        </div>

        {/* Footer */}
        <div className="px-4 py-3 bg-slate-50/70 border-t border-slate-200 flex flex-col sm:flex-row items-center justify-between gap-3 text-xs text-slate-600">
          <div>
            Showing <strong className="text-slate-900 tabular-nums">{sortedRecords.length}</strong> DVLA exam records
          </div>
          <div className="flex items-center gap-1.5">
            <button
              type="button"
              disabled={currentPage <= 1}
              onClick={() => setCurrentPage((p) => Math.max(1, p - 1))}
              className="px-2.5 py-1 text-xs border border-slate-200 rounded bg-white hover:bg-slate-50 disabled:opacity-40 transition"
            >
              Previous
            </button>
            <span className="px-2 text-xs text-slate-500 tabular-nums">
              Page {currentPage} of {totalPages}
            </span>
            <button
              type="button"
              disabled={currentPage >= totalPages}
              onClick={() => setCurrentPage((p) => Math.min(totalPages, p + 1))}
              className="px-2.5 py-1 text-xs border border-slate-200 rounded bg-white hover:bg-slate-50 disabled:opacity-40 transition"
            >
              Next
            </button>
          </div>
        </div>
      </div>

      {/* Add / Edit Modal */}
      {isModalOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/40 backdrop-blur-xs">
          <div className="bg-white rounded-xl shadow-2xl border border-slate-200 w-full max-w-md overflow-hidden">
            <div className="px-5 py-4 border-b border-slate-100 flex items-center justify-between bg-slate-50">
              <h3 className="text-sm font-bold text-slate-900">
                {editingRecord ? 'Edit DVLA Record' : 'New DVLA License Exam Entry'}
              </h3>
              <button
                type="button"
                onClick={() => setIsModalOpen(false)}
                className="text-slate-400 hover:text-slate-700 p-1 rounded"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            <form onSubmit={handleFormSubmit} className="p-5 space-y-3.5 text-xs">
              {formError && (
                <div className="p-2.5 bg-red-50 border border-red-200 rounded text-red-700">
                  {formError}
                </div>
              )}

              <div>
                <label className="block font-semibold text-slate-700 mb-1">
                  Applicant / Patient Name <span className="text-red-500">*</span>
                </label>
                <input
                  type="text"
                  required
                  value={formData.patientName || ''}
                  onChange={(e) => setFormData({ ...formData, patientName: e.target.value })}
                  placeholder="e.g. Eric Abam"
                  className="w-full px-3 py-2 bg-white border border-slate-300 rounded-lg text-slate-900 focus:ring-2 focus:ring-slate-900 focus:outline-hidden"
                />
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block font-semibold text-slate-700 mb-1">Date</label>
                  <input
                    type="text"
                    value={formData.date || ''}
                    onChange={(e) => setFormData({ ...formData, date: e.target.value })}
                    className="w-full px-3 py-2 bg-white border border-slate-300 rounded-lg text-slate-900 focus:ring-2 focus:ring-slate-900 focus:outline-hidden"
                  />
                </div>
                <div>
                  <label className="block font-semibold text-slate-700 mb-1">Contact Phone</label>
                  <input
                    type="text"
                    value={formData.contact || ''}
                    onChange={(e) => setFormData({ ...formData, contact: e.target.value })}
                    placeholder="0543861680"
                    className="w-full px-3 py-2 bg-white border border-slate-300 rounded-lg text-slate-900 focus:ring-2 focus:ring-slate-900 focus:outline-hidden"
                  />
                </div>
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block font-semibold text-slate-700 mb-1">DIC (Doctor)</label>
                  <input
                    type="text"
                    value={formData.dic || ''}
                    onChange={(e) => setFormData({ ...formData, dic: e.target.value })}
                    placeholder="Dr. Danso"
                    className="w-full px-3 py-2 bg-white border border-slate-300 rounded-lg text-slate-900 focus:ring-2 focus:ring-slate-900 focus:outline-hidden"
                  />
                </div>
                <div>
                  <label className="block font-semibold text-slate-700 mb-1">Recorded By (Col 2)</label>
                  <input
                    type="text"
                    value={formData.recordedBy || ''}
                    onChange={(e) => setFormData({ ...formData, recordedBy: e.target.value })}
                    className="w-full px-3 py-2 bg-white border border-slate-300 rounded-lg text-slate-900 focus:ring-2 focus:ring-slate-900 focus:outline-hidden"
                  />
                </div>
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block font-semibold text-slate-700 mb-1">New Exam Fee (GH₵)</label>
                  <input
                    type="number"
                    value={formData.newFee !== undefined ? formData.newFee : ''}
                    onChange={(e) => setFormData({ ...formData, newFee: e.target.value ? Number(e.target.value) : '' })}
                    placeholder="70 or 100"
                    className="w-full px-3 py-2 bg-white border border-slate-300 rounded-lg text-slate-900 tabular-nums focus:ring-2 focus:ring-slate-900 focus:outline-hidden"
                  />
                </div>
                <div>
                  <label className="block font-semibold text-slate-700 mb-1">Renewal Fee (GH₵)</label>
                  <input
                    type="number"
                    value={formData.renewalFee !== undefined ? formData.renewalFee : ''}
                    onChange={(e) => setFormData({ ...formData, renewalFee: e.target.value ? Number(e.target.value) : '' })}
                    placeholder="50 or 80"
                    className="w-full px-3 py-2 bg-white border border-slate-300 rounded-lg text-slate-900 tabular-nums focus:ring-2 focus:ring-slate-900 focus:outline-hidden"
                  />
                </div>
              </div>

              <div>
                <label className="block font-semibold text-slate-700 mb-1">Notes / Remarks</label>
                <input
                  type="text"
                  value={formData.notes || ''}
                  onChange={(e) => setFormData({ ...formData, notes: e.target.value })}
                  placeholder="e.g. RW"
                  className="w-full px-3 py-2 bg-white border border-slate-300 rounded-lg text-slate-900 focus:ring-2 focus:ring-slate-900 focus:outline-hidden"
                />
              </div>

              <div className="pt-3 border-t border-slate-100 flex items-center justify-end gap-2">
                <button
                  type="button"
                  onClick={() => setIsModalOpen(false)}
                  disabled={isSubmitting}
                  className="px-3 py-1.5 border border-slate-300 rounded bg-white hover:bg-slate-50 transition"
                >
                  Cancel
                </button>
                <button
                  type="submit"
                  disabled={isSubmitting}
                  className="px-4 py-1.5 font-semibold text-white bg-slate-900 hover:bg-slate-800 rounded transition"
                >
                  {isSubmitting ? 'Saving...' : editingRecord ? 'Save Changes' : 'Register DVLA'}
                </button>
              </div>
            </form>
          </div>
        </div>
      )}
    </div>
  );
}
