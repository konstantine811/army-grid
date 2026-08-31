import { createContext, useContext } from "react";
import type {
  BackendEjournalLiveState,
  BackendEjournalLiveVersion,
  BackendEjournalPbSource,
  BackendEjournalPbState,
} from "../../api";
import type { ExcelWorkbookSnapshot } from "../../excelRoundTrip";
import type { EjoosWorkspaceTab } from "../../app/navigation";
import type {
  EjoosAnketaFillMode,
  EjoosAnketaFillTarget,
} from "./ejoosAnketaFill";
import type {
  EjoosDiffSession,
  PersonChangeDecision,
} from "./ejoosPersonDiff";

export type { EjoosWorkspaceTab } from "../../app/navigation";

export type EjoosWorkspaceContextValue = {
  tab: EjoosWorkspaceTab;
  setTab: (tab: EjoosWorkspaceTab) => void;
  live: BackendEjournalLiveState | null;
  refreshLive: () => Promise<BackendEjournalLiveState>;
  pbSources: BackendEjournalPbState | null;
  refreshPbSources: () => Promise<BackendEjournalPbState>;
  ejoosSnapshot: ExcelWorkbookSnapshot | null;
  pbSnapshot: ExcelWorkbookSnapshot | null;
  session: EjoosDiffSession | null;
  selectedPersonId: string | null;
  setSelectedPersonId: (id: string | null) => void;
  message: string;
  error: string;
  isLoading: boolean;
  seedEjoos: (file: File) => Promise<void>;
  importEjoos: (file: File) => Promise<void>;
  loadEjoosFromDb: () => Promise<void>;
  ensureEjoosLoaded: () => Promise<ExcelWorkbookSnapshot | null>;
  loadPb: (file: File) => Promise<void>;
  savePbToDb: () => Promise<BackendEjournalPbSource | null>;
  loadPbFromDb: (id?: string) => Promise<void>;
  analyzePb: (file: File) => Promise<void>;
  rebuildOperations: () => Promise<void>;
  setDecision: (
    personChangeId: string,
    decision: PersonChangeDecision,
  ) => void;
  patchOpPayload: (
    personChangeId: string,
    opId: string,
    payloadPatch: Record<string, string>,
  ) => void;
  acceptReady: () => void;
  applyAccepted: () => Promise<void>;
  acceptAndApplyPerson: (personChangeId: string) => Promise<void>;
  downloadCurrentEjoos: () => Promise<void>;
  downloadVersion: (versionId: string, fileName?: string) => Promise<void>;
  downloadVersionProtocol: (version: BackendEjournalLiveVersion) => void;
  rollback: (versionId: string) => Promise<void>;
  restoreStylesFromHistory: (sourceVersionId?: string) => Promise<void>;
  fillSheetFromAnketa: (
    target: EjoosAnketaFillTarget,
    mode?: EjoosAnketaFillMode,
  ) => Promise<void>;
};

export const EjoosWorkspaceContext =
  createContext<EjoosWorkspaceContextValue | null>(null);

export function useEjoosWorkspace() {
  const context = useContext(EjoosWorkspaceContext);
  if (!context) {
    throw new Error(
      "useEjoosWorkspace must be used within EjoosWorkspaceProvider",
    );
  }
  return context;
}
