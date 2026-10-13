'use client';

import React from 'react';
import { Pill, ShieldAlert } from 'lucide-react';
import { MedsRecord } from '@/lib/types';

interface MedsTableProps {
  records: MedsRecord[];
  isLoading: boolean;
}

export function MedsTable({ records, isLoading }: MedsTableProps) {
  return (
    <div className="space-y-4">
      {/* Notice Banner */}
      <div className="p-4 bg-slate-50 border border-slate-200 rounded-xl flex items-start gap-3">
        <Pill className="w-5 h-5 text-teal-700 shrink-0 mt-0.5" />
        <div className="text-xs text-slate-600 space-y-1">
          <p className="font-semibold text-slate-800">
            Medicine Reference Data (MEDS Worksheet)
          </p>
          <p>
            Preserved verbatim from the clinic&apos;s legacy Excel file in original column order. In accordance with clinic policy, column headers in this worksheet are unclassified pending formal confirmation of column semantic definitions.
          </p>
        </div>
      </div>

      {/* Grid */}
      <div className="bg-white rounded-xl border border-slate-200 shadow-xs overflow-hidden">
        <div className="overflow-x-auto">
          <table className="w-full border-collapse text-left text-xs">
            <thead>
              <tr className="bg-slate-50/90 border-b border-slate-200 text-slate-600">
                <th className="py-2.5 px-4 font-semibold w-16 text-center text-slate-400">#</th>
                <th className="py-2.5 px-4 font-semibold">Medicine Item (Col 1)</th>
                <th className="py-2.5 px-4 font-semibold text-right">Value 1 (Col 2)</th>
                <th className="py-2.5 px-4 font-semibold text-right">Value 2 (Col 3)</th>
                <th className="py-2.5 px-4 font-semibold text-right">Value 3 (Col 4)</th>
                <th className="py-2.5 px-4 font-semibold text-right">Value 4 (Col 5)</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {isLoading ? (
                <tr>
                  <td colSpan={6} className="py-12 text-center text-slate-400">
                    Loading medicine catalog...
                  </td>
                </tr>
              ) : records.length === 0 ? (
                <tr>
                  <td colSpan={6} className="py-12 text-center text-slate-400">
                    No medicine records found.
                  </td>
                </tr>
              ) : (
                records.map((r, i) => (
                  <tr key={r.id || i} className="hover:bg-slate-50/70 transition-colors">
                    <td className="py-2.5 px-4 text-center text-slate-400 tabular-nums font-mono">
                      {i + 1}
                    </td>
                    <td className="py-2.5 px-4 font-semibold text-slate-900 flex items-center gap-2">
                      <span className="w-2 h-2 rounded-full bg-teal-600 inline-block" />
                      {r.medicineName}
                    </td>
                    <td className="py-2.5 px-4 text-slate-700 font-mono tabular-nums text-right">
                      {r.val1}
                    </td>
                    <td className="py-2.5 px-4 text-slate-700 font-mono tabular-nums text-right">
                      {r.val2}
                    </td>
                    <td className="py-2.5 px-4 text-slate-700 font-mono tabular-nums text-right">
                      {r.val3}
                    </td>
                    <td className="py-2.5 px-4 text-slate-700 font-mono tabular-nums text-right">
                      {r.val4}
                    </td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}
