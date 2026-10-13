'use client';

import React from 'react';
import { Lock, FileQuestion } from 'lucide-react';
import { Sheet2Record } from '@/lib/types';

interface Sheet2TableProps {
  records: Sheet2Record[];
  isLoading: boolean;
}

export function Sheet2Table({ records, isLoading }: Sheet2TableProps) {
  return (
    <div className="space-y-4">
      {/* Safety Notice Banner */}
      <div className="p-4 bg-amber-50/70 border border-amber-200 rounded-xl flex items-start gap-3">
        <FileQuestion className="w-5 h-5 text-amber-700 shrink-0 mt-0.5" />
        <div className="text-xs text-amber-900 space-y-1">
          <p className="font-bold flex items-center gap-1.5">
            <Lock className="w-3.5 h-3.5" />
            Unclassified Source Data (Sheet2) · Isolated Environment
          </p>
          <p className="text-amber-800 leading-relaxed">
            In compliance with clinic data architecture rules, records in &quot;Sheet2&quot; are isolated and deliberately not imported into the active patient or DVLA registers until its column meanings and clinical purpose are formally classified.
          </p>
        </div>
      </div>

      {/* Grid */}
      <div className="bg-white rounded-xl border border-slate-200 shadow-xs overflow-hidden">
        <div className="overflow-x-auto">
          <table className="w-full border-collapse text-left text-xs">
            <thead>
              <tr className="bg-slate-50/90 border-b border-slate-200 text-slate-600">
                <th className="py-2.5 px-4 font-semibold w-12 text-center text-slate-400">#</th>
                <th className="py-2.5 px-4 font-semibold">Column 1 (Name)</th>
                <th className="py-2.5 px-4 font-semibold text-center">Col 2</th>
                <th className="py-2.5 px-4 font-semibold text-center">Col 3</th>
                <th className="py-2.5 px-4 font-semibold">Column 4 (Contact)</th>
                <th className="py-2.5 px-4 font-semibold text-center">Col 5</th>
                <th className="py-2.5 px-4 font-semibold">Column 6 (Date)</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {isLoading ? (
                <tr>
                  <td colSpan={7} className="py-12 text-center text-slate-400">
                    Loading unclassified worksheet records...
                  </td>
                </tr>
              ) : records.length === 0 ? (
                <tr>
                  <td colSpan={7} className="py-12 text-center text-slate-400">
                    No records found in Sheet2.
                  </td>
                </tr>
              ) : (
                records.map((r, i) => (
                  <tr key={r.id || i} className="hover:bg-slate-50/70 transition-colors">
                    <td className="py-2.5 px-4 text-center text-slate-400 tabular-nums font-mono">
                      {i + 1}
                    </td>
                    <td className="py-2.5 px-4 font-semibold text-slate-900">{r.col1}</td>
                    <td className="py-2.5 px-4 text-center text-slate-300">-</td>
                    <td className="py-2.5 px-4 text-center text-slate-300">-</td>
                    <td className="py-2.5 px-4 text-slate-700 font-mono tabular-nums">{r.contact}</td>
                    <td className="py-2.5 px-4 text-center text-slate-300">-</td>
                    <td className="py-2.5 px-4 text-slate-700">{r.date}</td>
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
