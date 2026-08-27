import { useEffect, useMemo } from "react";
import { buildEjoosLiveView } from "./ejoosLiveViews";
import { useEjoosWorkspace } from "./EjoosWorkspaceContext";

const VIEW_TABS = new Set([
  "shpo",
  "oos",
  "excluded",
  "tempArrivals",
  "tempAbsents",
  "timesheet",
  "irrevocableLosses",
]);

/** Loads live ЕЖООС snapshot when needed and builds read-model for Phase 2 screens. */
export function useEjoosLiveView() {
  const {
    tab,
    live,
    ejoosSnapshot,
    pbSnapshot,
    session,
    ensureEjoosLoaded,
    isLoading,
  } = useEjoosWorkspace();

  useEffect(() => {
    if (!VIEW_TABS.has(tab)) return;
    if (ejoosSnapshot || !live?.current) return;
    void ensureEjoosLoaded();
  }, [tab, ejoosSnapshot, live?.current, ensureEjoosLoaded]);

  const view = useMemo(
    () =>
      buildEjoosLiveView({
        workbook: ejoosSnapshot,
        asOfDate: live?.current?.asOfDate,
        pbFileName: pbSnapshot?.fileName || session?.pbFileName,
        session,
        currentVersion: live?.current,
      }),
    [ejoosSnapshot, live?.current, pbSnapshot?.fileName, session],
  );

  return {
    view,
    isLoading,
    hasLiveFile: Boolean(live?.current),
    hasSnapshot: Boolean(ejoosSnapshot),
  };
}
