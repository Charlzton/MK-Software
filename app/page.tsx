'use client';

import React, { useState, useEffect, useCallback, useMemo } from 'react';
import { LoginCard, type LoginPreloadPayload } from '@/components/LoginCard';
import { ExcelSpreadsheet, SheetState } from '@/components/ExcelSpreadsheet';
import { createClient as createBrowserSupabaseClient } from '@/lib/supabase/client';
import { User } from '@/lib/types';
import { ConfirmDialog } from '@/components/ConfirmDialog';
import { syncThemeClassToDocument } from '@/hooks/useTheme';

type WorkspaceKey = 'reception' | 'doctor';

export default function Home() {
  useEffect(() => {
    syncThemeClassToDocument();
  }, []);
  const [user, setUser] = useState<User | null>(null);
  const [isCheckingAuth, setIsCheckingAuth] = useState(true);

  const [activeWorkspace, setActiveWorkspace] = useState<WorkspaceKey>('reception');
  const [visitedWorkspaces, setVisitedWorkspaces] = useState<WorkspaceKey[]>([]);

  // Persisted workspace data loaded exclusively from Supabase via /api/records (or preloaded on auth)
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
      const client = createBrowserSupabaseClient();
      if (!client) {
        return;
      }
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
        setVisitedWorkspaces([]);
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

  const hydrateSessionWorkspaces = useCallback(
    (authenticatedUser: User, preload?: LoginPreloadPayload) => {
      const landingWorkspace: WorkspaceKey = authenticatedUser.role === 'doctor' ? 'doctor' : 'reception';
      setActiveWorkspace(landingWorkspace);
      setVisitedWorkspaces([landingWorkspace]);

      if (preload?.sheets && preload.workspace === landingWorkspace) {
        if (landingWorkspace === 'reception') {
          setReceptionSheets(preload.sheets);
          setReceptionLoadError(null);
          setIsLoadingReception(false);
        } else {
          setDoctorSheets(preload.sheets);
          setDoctorLoadError(null);
          setIsLoadingDoctor(false);
        }

        // For Admin, prefetch Doctor in the background without blocking first paint of Reception
        if (authenticatedUser.role === 'admin') {
          void fetchWorkspace('doctor');
        }
        return;
      }

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

  // Single-roundtrip session + landing workbook bootstrap on mount
  useEffect(() => {
    fetch('/api/auth/me?includeWorkbook=1')
      .then((res) => (res.ok ? res.json() : null))
      .then((data) => {
        if (data && data.authenticated && data.user) {
          setUser(data.user);
          const preload: LoginPreloadPayload | undefined =
            data.initialWorkspace && data.initialWorkbook?.sheets
              ? {
                  workspace: data.initialWorkspace,
                  sheets: data.initialWorkbook.sheets
                }
              : undefined;
          hydrateSessionWorkspaces(data.user, preload);
        }
      })
      .catch(() => {
        // Unauthenticated
      })
      .finally(() => {
        setIsCheckingAuth(false);
      });
  }, [hydrateSessionWorkspaces]);

  const handleLoginSuccess = (authenticatedUser: User, preload?: LoginPreloadPayload) => {
    setSessionError(null);
    setUser(authenticatedUser);
    hydrateSessionWorkspaces(authenticatedUser, preload);
  };

  const handleLogout = async (): Promise<void> => {
    const res = await fetch('/api/auth/logout', { method: 'POST' });
    let data: { error?: unknown } = {};
    try {
      data = (await res.json()) as { error?: unknown };
    } catch {
      // Ignore non-JSON body on success/error
    }

    if (!res.ok) {
      const msg =
        typeof data.error === 'string' && data.error
          ? data.error
          : 'Failed to sign out. Please try again.';
      throw new Error(msg);
    }

    setSessionError(null);
    setImportNotice(null);
    setIsImportConfirmOpen(false);
    setUser(null);
    setVisitedWorkspaces([]);
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

  const activateWorkspace = useCallback((targetWorkspace: WorkspaceKey) => {
    setVisitedWorkspaces((prev) =>
      prev.includes(targetWorkspace) ? prev : [...prev, targetWorkspace]
    );
    setActiveWorkspace(targetWorkspace);
  }, []);

  const handleWorkspaceSwitchRequest = useCallback(
    (targetWorkspace: WorkspaceKey) => {
      if (targetWorkspace === activeWorkspace) return;
      if (!allowedWorkspaces.includes(targetWorkspace)) return;

      const currentHasUnsaved = activeWorkspace === 'reception' ? receptionHasUnsaved : doctorHasUnsaved;

      if (currentHasUnsaved) {
        setPendingWorkspaceSwitch(targetWorkspace);
        setShowSwitchWarning(true);
      } else {
        activateWorkspace(targetWorkspace);
      }
    },
    [activeWorkspace, allowedWorkspaces, receptionHasUnsaved, doctorHasUnsaved, activateWorkspace]
  );

  const confirmSwitchAnyway = () => {
    if (pendingWorkspaceSwitch && allowedWorkspaces.includes(pendingWorkspaceSwitch)) {
      activateWorkspace(pendingWorkspaceSwitch);
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
      <div className="flex flex-col mk-h-dvh w-full mk-glass-viewport overflow-hidden">
        <div className="h-12 px-3 sm:px-4 mk-glass-chrome flex items-center justify-between shrink-0">
          <div className="h-6 w-28 sm:w-32 rounded-full bg-white/50 dark:bg-white/10 animate-pulse" />
          <div className="flex items-center gap-2">
            <div className="h-7 w-14 sm:w-16 rounded-full bg-white/50 dark:bg-white/10 animate-pulse" />
            <div className="h-7 w-14 sm:w-16 rounded-full bg-white/50 dark:bg-white/10 animate-pulse" />
          </div>
        </div>
        <div className="h-10 px-2.5 sm:px-3 mk-glass-formula-bar flex items-center gap-2 shrink-0">
          <div className="h-7 w-[56px] sm:w-[72px] rounded-full mk-glass-pill-badge animate-pulse" />
          <div className="h-7 flex-1 rounded-full mk-glass-pill-input animate-pulse" />
        </div>
        <div className="flex-1 flex items-center justify-center p-6">
          <div className="flex items-center gap-2.5 px-4 py-2.5 rounded-full mk-glass-panel text-xs font-medium text-[var(--control-muted)]">
            <span className="w-4 h-4 border-2 border-[var(--control-border)] border-t-[var(--selection-ring)] rounded-full animate-spin" />
            <span>Loading MK Eye Center workspace...</span>
          </div>
        </div>
        <div className="h-10 px-3 mk-glass-tabstrip flex items-center gap-2 shrink-0">
          <div className="h-6 w-24 rounded-[10px] mk-glass-tab animate-pulse" />
          <div className="h-6 w-20 rounded-[10px] mk-glass-tab animate-pulse" />
        </div>
      </div>
    );
  }

  if (!user) {
    return (
      <main className="relative mk-min-h-dvh w-full flex flex-col overflow-x-hidden overflow-y-auto bg-slate-300 pb-[max(1.25rem,env(safe-area-inset-bottom))] pl-[max(0.75rem,env(safe-area-inset-left))] pr-[max(0.75rem,env(safe-area-inset-right))]">
        {/* Full-bleed edge-to-edge clinic background naturally dimmed ~32% for sharp foreground legibility */}
        <div className="fixed inset-0 z-0 pointer-events-none select-none overflow-hidden">
          {/* eslint-disable-next-line @next/next/no-img-element */}
          <img
            src="/assets/aistudio/login_cf1520db.jpg"
            alt=""
            aria-hidden="true"
            className="w-full h-full object-cover object-[28%_center] sm:object-[38%_center] md:object-center brightness-[0.68] contrast-[0.98]"
          />
          <div className="absolute inset-0 bg-slate-900/10" aria-hidden="true" />
        </div>

        {/* Top-Left Brand Logo visible on tablet and desktop viewports only (hidden on mobile <640px) */}
        <header className="relative z-20 w-full hidden sm:flex items-center justify-start px-5 md:px-7 pt-6 shrink-0">
          {/* eslint-disable-next-line @next/next/no-img-element */}
          <img
            src="/assets/aistudio/mk_long_logo_a5863dd4.png"
            alt="MK Eye Center"
            className="h-10 md:h-11 w-auto max-w-[215px] object-contain select-none brightness-[1.08] contrast-[1.12] drop-shadow-[0_0_10px_rgba(255,255,255,0.9)]"
            referrerPolicy="no-referrer"
          />
        </header>

        {/* Centered Glassmorphism Login Card */}
        <div className="relative z-10 flex-1 w-full flex items-center justify-center py-4 sm:py-6 px-1 sm:px-3">
          <LoginCard onSuccess={handleLoginSuccess} externalError={sessionError} />
        </div>
      </main>
    );
  }

  const canAccessReception = allowedWorkspaces.includes('reception');
  const canAccessDoctor = allowedWorkspaces.includes('doctor');
  const shouldMountReception = canAccessReception && (visitedWorkspaces.includes('reception') || activeWorkspace === 'reception');
  const shouldMountDoctor = canAccessDoctor && (visitedWorkspaces.includes('doctor') || activeWorkspace === 'doctor');
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
      {shouldMountReception && (
        <div className={activeWorkspace === 'reception' ? 'block mk-h-dvh w-full overflow-hidden' : 'hidden'}>
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

      {shouldMountDoctor && (
        <div className={activeWorkspace === 'doctor' ? 'block mk-h-dvh w-full overflow-hidden' : 'hidden'}>
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
