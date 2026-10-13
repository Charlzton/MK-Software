'use client';

import React, { useState, useEffect, useCallback } from 'react';
import { LoginCard } from '@/components/LoginCard';
import { ExcelSpreadsheet, SheetState } from '@/components/ExcelSpreadsheet';
import { INITIAL_WORKBOOK } from '@/lib/data/rawWorkbookData';
import { INITIAL_DOCTOR_WORKBOOK } from '@/lib/data/rawDoctorWorkbookData';
import { User } from '@/lib/types';
import { ConfirmDialog } from '@/components/ConfirmDialog';

export default function Home() {
  const [user, setUser] = useState<User | null>(null);
  const [isCheckingAuth, setIsCheckingAuth] = useState(true);

  // Active workspace: 'reception' or 'doctor'
  const [activeWorkspace, setActiveWorkspace] = useState<'reception' | 'doctor'>('reception');

  // Separate sheets state for Reception & Doctor
  const [receptionSheets, setReceptionSheets] = useState<Record<string, SheetState>>(INITIAL_WORKBOOK);
  const [doctorSheets, setDoctorSheets] = useState<Record<string, SheetState>>(INITIAL_DOCTOR_WORKBOOK);

  // Unsaved status tracking for each workspace
  const [receptionHasUnsaved, setReceptionHasUnsaved] = useState(false);
  const [doctorHasUnsaved, setDoctorHasUnsaved] = useState(false);

  // Doctor workbook configuration status
  const [isDoctorConfigured, setIsDoctorConfigured] = useState(false);
  const [doctorConfigNotice, setDoctorConfigNotice] = useState<string | undefined>(undefined);

  // Unsaved changes confirmation dialog state
  const [showSwitchWarning, setShowSwitchWarning] = useState(false);
  const [pendingWorkspaceSwitch, setPendingWorkspaceSwitch] = useState<'reception' | 'doctor' | null>(null);

  // Check existing session on load
  useEffect(() => {
    fetch('/api/auth/me')
      .then((res) => {
        if (res.ok) return res.json();
        return null;
      })
      .then((data) => {
        if (data && data.authenticated && data.user) {
          setUser(data.user);
        }
      })
      .catch(() => {
        // Unauthenticated
      })
      .finally(() => {
        setIsCheckingAuth(false);
      });
  }, []);

  // Fetch full sheets data for Reception & Doctor when user is authenticated
  useEffect(() => {
    if (user) {
      // 1. Fetch Reception sheets
      fetch('/api/records?workspace=reception')
        .then((res) => (res.ok ? res.json() : null))
        .then((data) => {
          if (data && data.sheets) {
            setReceptionSheets(data.sheets);
          }
        })
        .catch((err) => {
          console.error('Error loading Reception sheets:', err);
        });

      // 2. Fetch Doctor sheets
      fetch('/api/records?workspace=doctor')
        .then((res) => (res.ok ? res.json() : null))
        .then((data) => {
          if (data) {
            if (data.sheets) {
              setDoctorSheets(data.sheets);
            }
            setIsDoctorConfigured(Boolean(data.isConfigured));
            if (data.configMessage) {
              setDoctorConfigNotice(data.configMessage);
            }
          }
        })
        .catch((err) => {
          console.error('Error loading Doctor sheets:', err);
        });
    }
  }, [user]);

  const handleLoginSuccess = (authenticatedUser: User) => {
    setUser(authenticatedUser);
  };

  const handleLogout = async () => {
    try {
      await fetch('/api/auth/logout', { method: 'POST' });
    } catch {
      // ignore
    }
    setUser(null);
  };

  // Handle workspace tab switch request
  const handleWorkspaceSwitchRequest = useCallback(
    (targetWorkspace: 'reception' | 'doctor') => {
      if (targetWorkspace === activeWorkspace) return;

      const currentHasUnsaved = activeWorkspace === 'reception' ? receptionHasUnsaved : doctorHasUnsaved;

      if (currentHasUnsaved) {
        setPendingWorkspaceSwitch(targetWorkspace);
        setShowSwitchWarning(true);
      } else {
        setActiveWorkspace(targetWorkspace);
      }
    },
    [activeWorkspace, receptionHasUnsaved, doctorHasUnsaved]
  );

  const confirmSwitchAnyway = () => {
    if (pendingWorkspaceSwitch) {
      setActiveWorkspace(pendingWorkspaceSwitch);
      setPendingWorkspaceSwitch(null);
    }
    setShowSwitchWarning(false);
  };

  const cancelSwitch = () => {
    setPendingWorkspaceSwitch(null);
    setShowSwitchWarning(false);
  };

  if (isCheckingAuth) {
    return (
      <div className="min-h-screen flex items-center justify-center bg-white">
        <div className="flex flex-col items-center gap-2">
          <span className="w-5 h-5 border-2 border-slate-300 border-t-slate-800 rounded-full animate-spin" />
        </div>
      </div>
    );
  }

  if (!user) {
    return (
      <main className="min-h-screen flex items-center justify-center p-4 bg-white">
        <LoginCard onSuccess={handleLoginSuccess} />
      </main>
    );
  }

  return (
    <>
      {/* Reception Workspace Container */}
      <div className={activeWorkspace === 'reception' ? 'block h-screen w-full' : 'hidden'}>
        <ExcelSpreadsheet
          workspace="reception"
          initialSheets={receptionSheets}
          activeWorkspace={activeWorkspace}
          onWorkspaceSwitchRequest={handleWorkspaceSwitchRequest}
          onUnsavedStatusChange={setReceptionHasUnsaved}
          onLogout={handleLogout}
          isConfigured={true}
          defaultSheetName="PATIENT'S DATA"
        />
      </div>

      {/* Doctor Workspace Container */}
      <div className={activeWorkspace === 'doctor' ? 'block h-screen w-full' : 'hidden'}>
        <ExcelSpreadsheet
          workspace="doctor"
          initialSheets={doctorSheets}
          activeWorkspace={activeWorkspace}
          onWorkspaceSwitchRequest={handleWorkspaceSwitchRequest}
          onUnsavedStatusChange={setDoctorHasUnsaved}
          onLogout={handleLogout}
          isConfigured={isDoctorConfigured}
          configNotice={doctorConfigNotice}
          defaultSheetName="Lens Rx"
        />
      </div>

      {/* Warning Dialog when switching away from a workspace with unsaved edits */}
      <ConfirmDialog
        isOpen={showSwitchWarning}
        title="Unsaved Changes"
        message={`You have unsaved changes in the ${
          activeWorkspace === 'reception' ? 'Reception' : 'Doctor'
        } workspace. Switching without saving will leave your edits pending. Do you want to switch anyway?`}
        confirmLabel="Switch Workspace"
        confirmVariant="danger"
        onConfirm={confirmSwitchAnyway}
        onCancel={cancelSwitch}
      />
    </>
  );
}
