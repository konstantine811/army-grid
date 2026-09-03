import { useEffect } from "react";
import { importStaffSheetFromGoogle } from "../pages/anketa-data/staffSheetImport";
import {
  shouldAutoSyncStaffSheetRoster,
  writeStaffSheetLastSyncAt,
} from "../pages/excel-fill/staffSheet";
import { showAppToast } from "../shared/appToast";

export const STAFF_SHEET_SYNCED_EVENT = "army-grid:staff-sheet-synced";

/** Перевірка кожні 5 хв — щоб спрацювало обід/вечір, якщо вкладка відкрита. */
export const STAFF_SHEET_AUTO_SYNC_INTERVAL_MS = 5 * 60 * 1000;

export type StaffSheetAutoSyncResult = {
  synced: boolean;
  error?: string;
};

let syncInFlight: Promise<StaffSheetAutoSyncResult> | null = null;

export const runStaffSheetAutoSyncIfDue =
  async (): Promise<StaffSheetAutoSyncResult> => {
    if (!shouldAutoSyncStaffSheetRoster()) {
      return { synced: false };
    }
    if (syncInFlight) return syncInFlight;

    syncInFlight = (async () => {
      try {
        await importStaffSheetFromGoogle();
        writeStaffSheetLastSyncAt(new Date().toISOString());
        window.dispatchEvent(new CustomEvent(STAFF_SHEET_SYNCED_EVENT));
        return { synced: true };
      } catch (error) {
        return {
          synced: false,
          error:
            error instanceof Error
              ? error.message
              : "Не вдалося оновити «Штатку» з Google Sheets.",
        };
      } finally {
        syncInFlight = null;
      }
    })();

    return syncInFlight;
  };

/** Автооновлення «Штатки» з Google: ранок / обід / вечір (Europe/Kyiv). */
export const useStaffSheetAutoSync = (enabled: boolean) => {
  useEffect(() => {
    if (!enabled) return;

    let cancelled = false;

    const tick = async (options?: { quiet?: boolean }) => {
      const result = await runStaffSheetAutoSyncIfDue();
      if (cancelled || !result.synced) return;
      if (!options?.quiet) {
        showAppToast({
          title: "Штатка оновлена",
          message: "Завантажено з Google Sheets у БД.",
          variant: "INFO",
        });
      }
    };

    void tick({ quiet: true });

    const intervalId = window.setInterval(() => {
      void tick();
    }, STAFF_SHEET_AUTO_SYNC_INTERVAL_MS);

    const onVisible = () => {
      if (document.visibilityState === "visible") {
        void tick({ quiet: true });
      }
    };
    document.addEventListener("visibilitychange", onVisible);

    return () => {
      cancelled = true;
      window.clearInterval(intervalId);
      document.removeEventListener("visibilitychange", onVisible);
    };
  }, [enabled]);
};
