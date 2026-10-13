'use client';

import React, { useState, useMemo } from 'react';
import {
  Search,
  Plus,
  ArrowUpDown,
  Download,
  Filter,
  Eye,
  Edit2,
  Trash2,
  SlidersHorizontal
} from 'lucide-react';
import { PatientRecord, ColumnDefinition } from '@/lib/types';
import { KNOWN_DOCTORS, KNOWN_TEST_TYPES } from '@/lib/data/initialData';

interface PatientTableProps {
  records: PatientRecord[];
  columns: ColumnDefinition[];
  isLoading: boolean;
  onAddRecord: () => void;
  onEditRecord: (record: PatientRecord) => void;
  onDeleteRecord: (id: string, name: string) => void;
  onRefresh: () => void;
}

export function PatientTable({
  records,
  columns,
  isLoading,
  onAddRecord,
  onEditRecord,
  onDeleteRecord,
  onRefresh
}: PatientTableProps) {
  // Search & Filter state
  const [searchTerm, setSearchTerm] = useState('');
  const [doctorFilter, setDoctorFilter] = useState('');
  const [testTypeFilter, setTestTypeFilter] = useState('');

  // Sorting state
  const [sortField, setSortField] = useState<keyof PatientRecord>('date');
  const [sortAsc, setSortAsc] = useState<boolean>(false);

  // Pagination state
  const [currentPage, setCurrentPage] = useState(1);
  const [pageSize, setPageSize] = useState(25);

  // Column Visibility state
  const [visibleColKeys, setVisibleColKeys] = useState<Set<string>>(() => {
    return new Set(columns.map((c) => c.key));
  });
  const [showColMenu, setShowColMenu] = useState(false);

  // Selected patient for detail inspection
  const [inspectedPatient, setInspectedPatient] = useState<PatientRecord | null>(null);

  // Toggle sort
  const handleSort = (field: keyof PatientRecord) => {
    if (sortField === field) {
      setSortAsc(!sortAsc);
    } else {
      setSortField(field);
      setSortAsc(true);
    }
  };

  // Toggle column visibility
  const toggleColumn = (key: string) => {
    setVisibleColKeys((prev) => {
      const next = new Set(prev);
      if (next.has(key)) {
        if (next.size > 2) next.delete(key); // keep at least 2 columns
      } else {
        next.add(key);
      }
      return next;
    });
  };

  // Filtered & Sorted records
  const filteredRecords = useMemo(() => {
    return records.filter((r) => {
      const matchSearch =
        !searchTerm ||
        r.name.toLowerCase().includes(searchTerm.toLowerCase()) ||
        r.phoneNo.toLowerCase().includes(searchTerm.toLowerCase()) ||
        (r.recordedBy && r.recordedBy.toLowerCase().includes(searchTerm.toLowerCase()));

      const matchDoctor =
        !doctorFilter ||
        (r.examinedBy && r.examinedBy.toLowerCase().includes(doctorFilter.toLowerCase()));

      const matchTest =
        !testTypeFilter ||
        (r.testType && r.testType.toLowerCase() === testTypeFilter.toLowerCase());

      return matchSearch && matchDoctor && matchTest;
    });
  }, [records, searchTerm, doctorFilter, testTypeFilter]);

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

  // Paginated records
  const totalPages = Math.ceil(sortedRecords.length / pageSize) || 1;
  const paginatedRecords = useMemo(() => {
    const start = (currentPage - 1) * pageSize;
    return sortedRecords.slice(start, start + pageSize);
  }, [sortedRecords, currentPage, pageSize]);

  // CSV Export handler
  const handleExportCSV = () => {
    const headers = columns.map((c) => c.label).join(',');
    const rows = sortedRecords.map((r) => {
      return columns
        .map((c) => {
          const val = r[c.key as keyof PatientRecord] ?? '';
          const escaped = String(val).replace(/"/g, '""');
          return `"${escaped}"`;
        })
        .join(',');
    });

    const csvContent = [headers, ...rows].join('\n');
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.setAttribute('href', url);
    link.setAttribute('download', `MK_Patients_Register_${new Date().toISOString().slice(0, 10)}.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  return (
    <div className="space-y-4">
      {/* Search & Action Toolbar */}
      <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-xs flex flex-col lg:flex-row lg:items-center justify-between gap-3">
        {/* Left: Search & Filters */}
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
              placeholder="Search by patient name, phone, or staff..."
              className="w-full pl-9 pr-3 py-1.5 text-xs bg-slate-50 border border-slate-200 rounded-lg text-slate-800 placeholder:text-slate-400 focus:outline-hidden focus:ring-2 focus:ring-slate-900 focus:bg-white transition"
            />
          </div>

          <div className="flex items-center gap-2">
            <div className="relative">
              <select
                value={doctorFilter}
                onChange={(e) => {
                  setDoctorFilter(e.target.value);
                  setCurrentPage(1);
                }}
                className="text-xs bg-slate-50 border border-slate-200 rounded-lg px-2.5 py-1.5 text-slate-700 focus:outline-hidden focus:ring-2 focus:ring-slate-900"
              >
                <option value="">All Doctors</option>
                {KNOWN_DOCTORS.map((doc) => (
                  <option key={doc} value={doc}>
                    {doc}
                  </option>
                ))}
              </select>
            </div>

            <div className="relative">
              <select
                value={testTypeFilter}
                onChange={(e) => {
                  setTestTypeFilter(e.target.value);
                  setCurrentPage(1);
                }}
                className="text-xs bg-slate-50 border border-slate-200 rounded-lg px-2.5 py-1.5 text-slate-700 focus:outline-hidden focus:ring-2 focus:ring-slate-900"
              >
                <option value="">All Test Types</option>
                {KNOWN_TEST_TYPES.map((t) => (
                  <option key={t} value={t}>
                    {t}
                  </option>
                ))}
              </select>
            </div>

            {(searchTerm || doctorFilter || testTypeFilter) && (
              <button
                type="button"
                onClick={() => {
                  setSearchTerm('');
                  setDoctorFilter('');
                  setTestTypeFilter('');
                  setCurrentPage(1);
                }}
                className="text-xs text-slate-500 hover:text-slate-900 underline"
              >
                Reset
              </button>
            )}
          </div>
        </div>

        {/* Right: Actions */}
        <div className="flex items-center gap-2 shrink-0">
          <button
            type="button"
            onClick={onRefresh}
            className="px-2.5 py-1.5 text-xs font-medium text-slate-700 bg-white border border-slate-200 rounded-lg hover:bg-slate-50 transition"
            title="Refresh patient records"
          >
            Refresh
          </button>

          {/* Column Visibility Dropdown */}
          <div className="relative">
            <button
              type="button"
              onClick={() => setShowColMenu(!showColMenu)}
              className="px-2.5 py-1.5 text-xs font-medium text-slate-700 bg-white border border-slate-200 rounded-lg hover:bg-slate-50 transition flex items-center gap-1.5"
            >
              <SlidersHorizontal className="w-3.5 h-3.5 text-slate-500" />
              <span>Columns</span>
            </button>

            {showColMenu && (
              <div className="absolute right-0 mt-1 w-56 bg-white border border-slate-200 rounded-lg shadow-xl p-3 z-30 max-h-80 overflow-y-auto space-y-1">
                <div className="text-[11px] font-bold uppercase tracking-wider text-slate-400 mb-2">
                  Toggle Columns
                </div>
                {columns.map((col) => (
                  <label
                    key={col.key}
                    className="flex items-center gap-2 text-xs text-slate-700 hover:bg-slate-50 p-1 rounded cursor-pointer"
                  >
                    <input
                      type="checkbox"
                      checked={visibleColKeys.has(col.key)}
                      onChange={() => toggleColumn(col.key)}
                      className="rounded border-slate-300 text-slate-900"
                    />
                    <span>{col.label}</span>
                  </label>
                ))}
              </div>
            )}
          </div>

          <button
            type="button"
            onClick={handleExportCSV}
            className="px-2.5 py-1.5 text-xs font-medium text-slate-700 bg-white border border-slate-200 rounded-lg hover:bg-slate-50 transition flex items-center gap-1.5"
            title="Export filtered records to CSV"
          >
            <Download className="w-3.5 h-3.5 text-slate-500" />
            <span>Export</span>
          </button>

          <button
            type="button"
            onClick={onAddRecord}
            className="px-3.5 py-1.5 text-xs font-semibold text-white bg-slate-900 hover:bg-slate-800 rounded-lg shadow-xs transition flex items-center gap-1.5 whitespace-nowrap shrink-0"
          >
            <Plus className="w-3.5 h-3.5" />
            <span>Add Patient</span>
          </button>
        </div>
      </div>

      {/* Spreadsheet Data Grid */}
      <div className="bg-white rounded-xl border border-slate-200 shadow-xs overflow-hidden">
        <div className="overflow-x-auto min-h-[400px]">
          <table className="w-full border-collapse text-left text-xs">
            {/* Header */}
            <thead>
              <tr className="bg-slate-50/90 border-b border-slate-200 text-slate-600 select-none">
                <th className="py-2.5 px-3 font-semibold w-12 text-center text-slate-400">#</th>
                {columns
                  .filter((c) => visibleColKeys.has(c.key))
                  .map((col) => {
                    const isSorted = sortField === col.key;
                    return (
                      <th
                        key={col.key}
                        onClick={() => handleSort(col.key as keyof PatientRecord)}
                        className="py-2.5 px-3 font-semibold whitespace-nowrap cursor-pointer hover:bg-slate-100/70 transition-colors"
                      >
                        <div className="flex items-center gap-1.5">
                          <span>{col.label}</span>
                          <ArrowUpDown
                            className={`w-3 h-3 ${isSorted ? 'text-slate-900' : 'text-slate-300'}`}
                          />
                        </div>
                      </th>
                    );
                  })}
                <th className="py-2.5 px-3 font-semibold text-right pr-4">Actions</th>
              </tr>
            </thead>

            {/* Body */}
            <tbody className="divide-y divide-slate-100">
              {isLoading ? (
                <tr>
                  <td
                    colSpan={visibleColKeys.size + 2}
                    className="py-16 text-center text-slate-400"
                  >
                    <div className="flex flex-col items-center justify-center gap-2">
                      <span className="w-5 h-5 border-2 border-slate-300 border-t-slate-800 rounded-full animate-spin" />
                      <span className="text-xs">Loading patient register...</span>
                    </div>
                  </td>
                </tr>
              ) : paginatedRecords.length === 0 ? (
                <tr>
                  <td
                    colSpan={visibleColKeys.size + 2}
                    className="py-16 text-center text-slate-500"
                  >
                    <div className="max-w-sm mx-auto space-y-2">
                      <p className="font-semibold text-slate-800">No patient records found</p>
                      <p className="text-xs text-slate-500">
                        {searchTerm || doctorFilter || testTypeFilter
                          ? 'No entries match your search criteria. Try resetting filters.'
                          : 'The register is currently empty. Click "Add Patient" to register the first entry.'}
                      </p>
                    </div>
                  </td>
                </tr>
              ) : (
                paginatedRecords.map((record, index) => {
                  const rowNumber = (currentPage - 1) * pageSize + index + 1;
                  return (
                    <tr
                      key={record.id}
                      className="hover:bg-slate-50/80 transition-colors group cursor-pointer"
                      onClick={() => setInspectedPatient(record)}
                    >
                      <td className="py-2 px-3 text-center text-slate-400 tabular-nums font-mono text-[11px]">
                        {rowNumber}
                      </td>

                      {columns
                        .filter((c) => visibleColKeys.has(c.key))
                        .map((col) => {
                          const val = record[col.key as keyof PatientRecord];
                          const isNumeric = col.type === 'number';

                          // Special styling for crucial clinical columns
                          let formatted = val !== undefined && val !== null ? String(val) : '';
                          let cellClass = 'py-2 px-3 text-slate-700 whitespace-nowrap';

                          if (col.key === 'name') {
                            cellClass += ' font-semibold text-slate-900';
                          } else if (isNumeric || col.key === 'phoneNo' || col.key === 'glassesBal') {
                            cellClass += ' tabular-nums font-mono';
                          }

                          // Highlight balance if remaining
                          if (col.key === 'glassesBal' && Number(val) > 0) {
                            cellClass += ' text-amber-700 font-semibold';
                          }

                          return (
                            <td key={col.key} className={cellClass}>
                              {formatted || <span className="text-slate-300">-</span>}
                            </td>
                          );
                        })}

                      {/* Row Actions */}
                      <td
                        className="py-2 px-3 text-right pr-4 whitespace-nowrap"
                        onClick={(e) => e.stopPropagation()}
                      >
                        <div className="flex items-center justify-end gap-1 opacity-80 group-hover:opacity-100 transition">
                          <button
                            type="button"
                            onClick={() => setInspectedPatient(record)}
                            className="p-1 text-slate-400 hover:text-slate-700 hover:bg-slate-100 rounded transition"
                            title="Inspect Patient Details"
                          >
                            <Eye className="w-3.5 h-3.5" />
                          </button>
                          <button
                            type="button"
                            onClick={() => onEditRecord(record)}
                            className="p-1 text-slate-400 hover:text-slate-700 hover:bg-slate-100 rounded transition"
                            title="Edit Record"
                          >
                            <Edit2 className="w-3.5 h-3.5" />
                          </button>
                          <button
                            type="button"
                            onClick={() => onDeleteRecord(record.id, record.name)}
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

        {/* Footer: Pagination & Counts */}
        <div className="px-4 py-3 bg-slate-50/70 border-t border-slate-200 flex flex-col sm:flex-row items-center justify-between gap-3 text-xs text-slate-600">
          <div className="flex items-center gap-2">
            <span>
              Showing <strong className="text-slate-900 tabular-nums">{sortedRecords.length}</strong> patient records
            </span>
            <span className="text-slate-300">|</span>
            <div className="flex items-center gap-1.5">
              <span>Rows per page:</span>
              <select
                value={pageSize}
                onChange={(e) => {
                  setPageSize(Number(e.target.value));
                  setCurrentPage(1);
                }}
                className="bg-white border border-slate-200 rounded px-1.5 py-0.5 text-xs text-slate-700 focus:outline-hidden"
              >
                <option value={25}>25</option>
                <option value={50}>50</option>
                <option value={100}>100</option>
              </select>
            </div>
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

      {/* Patient Inspection Drawer / Modal */}
      {inspectedPatient && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/40 backdrop-blur-xs">
          <div className="bg-white rounded-xl shadow-2xl border border-slate-200 w-full max-w-lg overflow-hidden">
            <div className="px-5 py-4 border-b border-slate-100 flex items-center justify-between bg-slate-50">
              <div>
                <h3 className="text-sm font-bold text-slate-900">{inspectedPatient.name}</h3>
                <p className="text-xs text-slate-500">ID: {inspectedPatient.id} · Registered {inspectedPatient.date}</p>
              </div>
              <button
                type="button"
                onClick={() => setInspectedPatient(null)}
                className="text-slate-400 hover:text-slate-700 text-xs px-2 py-1 rounded"
              >
                Close
              </button>
            </div>

            <div className="p-5 space-y-4 max-h-[70vh] overflow-y-auto text-xs">
              <div className="grid grid-cols-2 gap-3 p-3 bg-slate-50 rounded-lg border border-slate-100">
                <div>
                  <span className="text-slate-400 block text-[10px] uppercase font-bold">Contact</span>
                  <span className="font-mono text-slate-800">{inspectedPatient.phoneNo || 'N/A'}</span>
                </div>
                <div>
                  <span className="text-slate-400 block text-[10px] uppercase font-bold">Examined By</span>
                  <span className="font-semibold text-slate-800">{inspectedPatient.examinedBy || 'N/A'}</span>
                </div>
                <div>
                  <span className="text-slate-400 block text-[10px] uppercase font-bold">Test Type</span>
                  <span className="text-slate-800">{inspectedPatient.testType || 'Standard'}</span>
                </div>
                <div>
                  <span className="text-slate-400 block text-[10px] uppercase font-bold">Recorded By</span>
                  <span className="text-slate-800">{inspectedPatient.recordedBy || 'N/A'}</span>
                </div>
              </div>

              <div>
                <h4 className="text-[11px] font-bold uppercase tracking-wider text-slate-500 mb-2">
                  Financial & Optical Breakdown
                </h4>
                <div className="grid grid-cols-3 gap-2 text-center">
                  <div className="p-2.5 bg-slate-50 border border-slate-200 rounded">
                    <span className="text-slate-400 block text-[10px]">Test Fee</span>
                    <span className="font-mono font-semibold text-slate-800">
                      GH₵ {inspectedPatient.testFee || 0}
                    </span>
                  </div>
                  <div className="p-2.5 bg-slate-50 border border-slate-200 rounded">
                    <span className="text-slate-400 block text-[10px]">Medication</span>
                    <span className="font-mono font-semibold text-slate-800">
                      GH₵ {inspectedPatient.medication || 0}
                    </span>
                  </div>
                  <div className="p-2.5 bg-slate-100 border border-slate-300 rounded">
                    <span className="text-slate-500 block text-[10px] font-bold">Total Bill</span>
                    <span className="font-mono font-bold text-slate-900">
                      GH₵ {inspectedPatient.total || 0}
                    </span>
                  </div>
                </div>

                <div className="grid grid-cols-3 gap-2 text-center mt-2">
                  <div className="p-2.5 bg-slate-50 border border-slate-200 rounded">
                    <span className="text-slate-400 block text-[10px]">Glass Total</span>
                    <span className="font-mono text-slate-800">{inspectedPatient.glassTotal || '-'}</span>
                  </div>
                  <div className="p-2.5 bg-slate-50 border border-slate-200 rounded">
                    <span className="text-slate-400 block text-[10px]">Paid</span>
                    <span className="font-mono text-slate-800">{inspectedPatient.paid || '-'}</span>
                  </div>
                  <div className="p-2.5 bg-amber-50/60 border border-amber-200 rounded">
                    <span className="text-amber-800 block text-[10px] font-bold">Glasses BAL</span>
                    <span className="font-mono font-bold text-amber-900">
                      {inspectedPatient.glassesBal !== '' ? `GH₵ ${inspectedPatient.glassesBal}` : '-'}
                    </span>
                  </div>
                </div>
              </div>

              {(inspectedPatient.reviewDate || inspectedPatient.reviewTest) && (
                <div className="p-3 bg-teal-50/50 border border-teal-100 rounded-lg">
                  <span className="text-teal-900 block text-[10px] uppercase font-bold">
                    Follow-Up Review
                  </span>
                  <div className="flex items-center gap-3 mt-1 text-teal-800">
                    <span>Date: <strong>{inspectedPatient.reviewDate}</strong></span>
                    <span>Test: <strong>{inspectedPatient.reviewTest}</strong></span>
                  </div>
                </div>
              )}
            </div>

            <div className="px-5 py-3 bg-slate-50 border-t border-slate-100 flex items-center justify-end gap-2">
              <button
                type="button"
                onClick={() => {
                  const p = inspectedPatient;
                  setInspectedPatient(null);
                  onEditRecord(p);
                }}
                className="px-3 py-1.5 text-xs font-semibold text-slate-800 bg-white border border-slate-300 rounded hover:bg-slate-50 transition"
              >
                Edit Patient Record
              </button>
              <button
                type="button"
                onClick={() => setInspectedPatient(null)}
                className="px-3 py-1.5 text-xs font-semibold text-white bg-slate-900 rounded hover:bg-slate-800 transition"
              >
                Done
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
