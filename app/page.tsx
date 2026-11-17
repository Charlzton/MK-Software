'use client';

import React, { useState, useEffect, useCallback, useMemo } from 'react';
import { LoginCard } from '@/components/LoginCard';
import { ExcelSpreadsheet, SheetState } from '@/components/ExcelSpreadsheet';
import { createClient as createBrowserSupabaseClient } from '@/lib/supabase/client';
import { User } from '@/lib/types';
import { ConfirmDialog } from '@/components/ConfirmDialog';

type WorkspaceKey = 'reception' | 'doctor';

export default function Home() {
  const [user, setUser] = useState<User | null>(null);
  const [isCheckingAuth, setIsCheckingAuth] = useState(true);

  const [activeWorkspace, setActiveWorkspace] = useState<WorkspaceKey>('reception');

  // Persisted workspace data loaded exclusively from Supabase via /api/records
  const [receptionSheets, setReceptionSheets] = useState<Record<string, SheetState>>({});
  const [doctorSheets, setDoctorSheets] = useState<Record<string, SheetState>>({});

  const [isLoadingReception, setIsLoadingReception] = useState(false);
  const [isLoadingDoctor, setIsLoadingDoctor] = useState(false);
  const [receptionLoadError, setReceptionLoadError] = useState<string | null>(null);
  const [doctorLoadError, setDoctorLoadError] = useState<string | null>(null);

  const [receptionHasUnsaved, setReceptionHasUnsaved] = useState(false);
  const [doctorHasUnsaved, setDoctorHasUnsaved] = useState(false);

  const [showSwitchWarning, setShowSwitchWarning] = useState(false);
  const [pendingWorkspaceSwitch, setPendingWorkspaceSwitch] = useState<WorkspaceKey | null>(null);
  const [sessionError, setSessionError] = useState<string | null>(null);
  const [isImportConfirmOpen, setIsImportConfirmOpen] = useState(false);
  const [isImporting, setIsImporting] = useState(false);
  const [importNotice, setImportNotice] = useState<{
    type: 'success' | 'info' | 'error';
    message: string;
  } | null>(null);

  const allowedWorkspaces = useMemo<WorkspaceKey[]>(() => {
    if (!user) return [];
    if (user.role === 'admin') return ['reception', 'doctor'];
    if (user.role === 'doctor') return ['doctor'];
    return ['reception'];
  }, [user]);

  // Verify browser Supabase client initialization cleanly when client app mounts
  useEffect(() => {
    try {
      createBrowserSupabaseClient();
    } catch (err) {
      console.error(
        'Browser Supabase client initialization failed:',
        err instanceof Error ? err.message : 'Configuration missing'
      );
    }
  }, []);

  const fetchWorkspace = useCallback(async (workspace: WorkspaceKey) => {
    const setLoading = workspace === 'reception' ? setIsLoadingReception : setIsLoadingDoctor;
    const setLoadError = workspace === 'reception' ? setReceptionLoadError : setDoctorLoadError;
    const setSheets = workspace === 'reception' ? setReceptionSheets : setDoctorSheets;

    setLoading(true);
    setLoadError(null);

    try {
      const res = await fetch(`/api/records?workspace=${workspace}`);
      if (res.status === 401) {
        setSessionError(
          'Your session could not be verified on workspace load. Please ensure cookies are enabled and sign in again.'
        );
        setUser(null);
        setReceptionSheets({});
        setDoctorSheets({});
        return;
      }
      const data = await res.json();
      if (!res.ok) {
        throw new Error(data.error || `Failed to load ${workspace} workbook.`);
      }
      setSheets(data.sheets || {});
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : `Failed to load ${workspace} workbook.`;
      setLoadError(msg);
    } finally {
      setLoading(false);
    }
  }, []);

  const loadAuthorizedWorkspaces = useCallback(
    (authenticatedUser: User) => {
      if (authenticatedUser.role === 'doctor') {
        void fetchWorkspace('doctor');
      } else if (authenticatedUser.role === 'reception') {
        void fetchWorkspace('reception');
      } else {
        void fetchWorkspace('reception');
        void fetchWorkspace('doctor');
      }
    },
    [fetchWorkspace]
  );

  // Check existing session on load
  useEffect(() => {
    fetch('/api/auth/me')
      .then((res) => (res.ok ? res.json() : null))
      .then((data) => {
        if (data && data.authenticated && data.user) {
          setUser(data.user);
          setActiveWorkspace(data.user.role === 'doctor' ? 'doctor' : 'reception');
          loadAuthorizedWorkspaces(data.user);
        }
      })
      .catch(() => {
        // Unauthenticated
      })
      .finally(() => {
        setIsCheckingAuth(false);
      });
  }, [loadAuthorizedWorkspaces]);

  const handleLoginSuccess = (authenticatedUser: User) => {
    setSessionError(null);
    setUser(authenticatedUser);
    setActiveWorkspace(authenticatedUser.role === 'doctor' ? 'doctor' : 'reception');
    loadAuthorizedWorkspaces(authenticatedUser);
  };

  const handleLogout = async () => {
    try {
      await fetch('/api/auth/logout', { method: 'POST' });
    } catch {
      // ignore
    }
    setSessionError(null);
    setImportNotice(null);
    setIsImportConfirmOpen(false);
    setUser(null);
    setReceptionSheets({});
    setDoctorSheets({});
    setReceptionHasUnsaved(false);
    setDoctorHasUnsaved(false);
  };

  const handleRequestInitialImport = useCallback(() => {
    if (user?.role !== 'admin') return;
    setImportNotice(null);
    setIsImportConfirmOpen(true);
  }, [user?.role]);

  const handleConfirmInitialImport = async () => {
    if (user?.role !== 'admin') return;
    setIsImporting(true);
    setImportNotice(null);

    try {
      const res = await fetch('/api/records/import', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ confirm: true })
      });
      const data = await res.json();

      if (!res.ok) {
        throw new Error(data.error || 'Initial workbook import failed.');
      }

      const recSheets = Number(data.receptionSheets ?? 0);
      const recRows = Number(data.receptionRows ?? 0);
      const docSheets = Number(data.doctorSheets ?? 0);
      const docRows = Number(data.doctorRows ?? 0);
      const totalSheets = Number(data.totalSheets ?? recSheets + docSheets);
      const totalRows = Number(data.totalRows ?? recRows + docRows);

      setImportNotice({
        type: data.alreadyImported ? 'info' : 'success',
        message: data.alreadyImported
          ? `Historical workbooks were already imported (${totalSheets} worksheets, ${totalRows} rows — Reception: ${recSheets} sheets / ${recRows} rows, Doctor: ${docSheets} sheets / ${docRows} rows).`
          : `Successfully imported ${totalSheets} worksheets and ${totalRows} rows into Supabase (Reception: ${recSheets} sheets / ${recRows} rows, Doctor: ${docSheets} sheets / ${docRows} rows).`
      });

      setIsImportConfirmOpen(false);
      await Promise.all([fetchWorkspace('reception'), fetchWorkspace('doctor')]);
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : 'Initial workbook import failed.';
      setImportNotice({ type: 'error', message: msg });
      setIsImportConfirmOpen(false);
    } finally {
      setIsImporting(false);
    }
  };

  const handleWorkspaceSwitchRequest = useCallback(
    (targetWorkspace: WorkspaceKey) => {
      if (targetWorkspace === activeWorkspace) return;
      if (!allowedWorkspaces.includes(targetWorkspace)) return;

      const currentHasUnsaved = activeWorkspace === 'reception' ? receptionHasUnsaved : doctorHasUnsaved;

      if (currentHasUnsaved) {
        setPendingWorkspaceSwitch(targetWorkspace);
        setShowSwitchWarning(true);
      } else {
        setActiveWorkspace(targetWorkspace);
      }
    },
    [activeWorkspace, allowedWorkspaces, receptionHasUnsaved, doctorHasUnsaved]
  );

  const confirmSwitchAnyway = () => {
    if (pendingWorkspaceSwitch && allowedWorkspaces.includes(pendingWorkspaceSwitch)) {
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
        <LoginCard onSuccess={handleLoginSuccess} externalError={sessionError} />
      </main>
    );
  }

  const canAccessReception = allowedWorkspaces.includes('reception');
  const canAccessDoctor = allowedWorkspaces.includes('doctor');
  const isAdmin = user.role === 'admin';
  const canAdminImport =
    isAdmin &&
    !isLoadingReception &&
    !isLoadingDoctor &&
    !receptionLoadError &&
    !doctorLoadError &&
    Object.keys(receptionSheets).length === 0 &&
    Object.keys(doctorSheets).length === 0;

  return (
    <>
      {canAccessReception && (
        <div className={activeWorkspace === 'reception' ? 'block h-screen w-full' : 'hidden'}>
          <ExcelSpreadsheet
            workspace="reception"
            user={user}
            allowedWorkspaces={allowedWorkspaces}
            initialSheets={receptionSheets}
            activeWorkspace={activeWorkspace}
            onWorkspaceSwitchRequest={handleWorkspaceSwitchRequest}
            onUnsavedStatusChange={setReceptionHasUnsaved}
            onReloadWorkspace={() => fetchWorkspace('reception')}
            onLogout={handleLogout}
            isLoading={isLoadingReception}
            loadError={receptionLoadError}
            defaultSheetName="PATIENT'S DATA"
            canAdminImport={canAdminImport}
            onRequestInitialImport={isAdmin ? handleRequestInitialImport : undefined}
            importNotice={isAdmin ? importNotice : null}
            onDismissImportNotice={isAdmin ? () => setImportNotice(null) : undefined}
          />
        </div>
      )}

      {canAccessDoctor && (
        <div className={activeWorkspace === 'doctor' ? 'block h-screen w-full' : 'hidden'}>
          <ExcelSpreadsheet
            workspace="doctor"
            user={user}
            allowedWorkspaces={allowedWorkspaces}
            initialSheets={doctorSheets}
            activeWorkspace={activeWorkspace}
            onWorkspaceSwitchRequest={handleWorkspaceSwitchRequest}
            onUnsavedStatusChange={setDoctorHasUnsaved}
            onReloadWorkspace={() => fetchWorkspace('doctor')}
            onLogout={handleLogout}
            isLoading={isLoadingDoctor}
            loadError={doctorLoadError}
            defaultSheetName="Lens Rx"
            canAdminImport={canAdminImport}
            onRequestInitialImport={isAdmin ? handleRequestInitialImport : undefined}
            importNotice={isAdmin ? importNotice : null}
            onDismissImportNotice={isAdmin ? () => setImportNotice(null) : undefined}
          />
        </div>
      )}

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

      {isAdmin && (
        <ConfirmDialog
          isOpen={isImportConfirmOpen}
          title="Import Initial Workbooks"
          message="This one-time administrator action imports both Reception (4 worksheets, 545 rows) and Doctor (9 worksheets, 319 rows) historical workbooks into Supabase in a single atomic transaction (13 worksheets, 864 rows total). Existing live data will never be overwritten."
          confirmLabel={isImporting ? 'Importing...' : 'Import Both Workbooks'}
          confirmVariant="primary"
          isLoading={isImporting}
          onConfirm={() => void handleConfirmInitialImport()}
          onCancel={() => {
            if (!isImporting) setIsImportConfirmOpen(false);
          }}
        />
      )}
    </>
  );
}
