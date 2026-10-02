import { useState } from "react";
import { readWorkbookSnapshot } from "../../excelRoundTrip";
import {
  formatAnketaBulkMergeReport,
  mergeCachedAnketaToPersonnel,
} from "../anketa-data/anketaPersonMerge";
import { runParseVkTpvDovidkyHeavy } from "../anketa-data/runStaffSheetHeavyJobs";
import { importStaffSheetFromFile } from "../anketa-data/staffSheetImport";
import { ATTACHMENT_HEAL_SESSION_KEY } from "./healOrphanPersonnelAttachments";
import {
  anketaMergeError,
  rosterImportError,
  rosterImportStatusMessage,
  vkTpvImportError,
} from "./personnelImportStatus";
import {
  formatVkTpvDovidkyMergeReport,
  mergeVkTpvDovidkyRecords,
} from "./vkTpvDovidkyImport";

const reportProgress = (onStatus: (text: string) => void) => {
  let lastProgressAt = 0;
  return (done: number, total: number, text: string) => {
    const now = Date.now();
    if (done !== total && now - lastProgressAt < 250) return;
    lastProgressAt = now;
    onStatus(text);
  };
};

export const usePersonnelWorkbookImports = ({
  reload,
  holdStatusUntilRef,
  setIsLoading,
  setMessage,
}: {
  reload: (options?: { force?: boolean }) => Promise<unknown> | unknown;
  holdStatusUntilRef: { current: number };
  setIsLoading: (value: boolean) => void;
  setMessage: (value: string) => void;
}) => {
  const [isMergingAnketaData, setIsMergingAnketaData] = useState(false);
  const [isMergingVkTpvDovidky, setIsMergingVkTpvDovidky] = useState(false);

  const importVkTpvDovidkyWorkbook = async (file: File | undefined) => {
    if (!file) return;
    setIsMergingVkTpvDovidky(true);
    setMessage(`Читаю «${file.name}»…`);
    try {
      const snapshot = await readWorkbookSnapshot(file);
      const records = await runParseVkTpvDovidkyHeavy(snapshot);
      const progress = reportProgress(setMessage);
      const report = await mergeVkTpvDovidkyRecords(records, {
        onProgress: (done, total) =>
          progress(done, total, `ВК ТПВ ДОВІДКИ · ${done}/${total}`),
      });
      await reload();
      setMessage(`ВК ТПВ ДОВІДКИ · ${formatVkTpvDovidkyMergeReport(report)}.`);
    } catch (error) {
      setMessage(vkTpvImportError(error));
    } finally {
      setIsMergingVkTpvDovidky(false);
    }
  };

  const mergeMissingFieldsFromAnketaData = async () => {
    setIsMergingAnketaData(true);
    setMessage("Завантажую анкетні дані…");
    try {
      const progress = reportProgress(setMessage);
      const report = await mergeCachedAnketaToPersonnel({
        onProgress: (done, total) =>
          progress(done, total, `Доповнення з анкетних даних… ${done}/${total}`),
        onStatus: setMessage,
      });
      sessionStorage.removeItem(ATTACHMENT_HEAL_SESSION_KEY);
      await reload();
      holdStatusUntilRef.current = Date.now() + 20_000;
      setMessage(
        `Доповнено з анкетних даних · ${formatAnketaBulkMergeReport(report)}.`,
      );
    } catch (error) {
      setMessage(anketaMergeError(error));
    } finally {
      setIsMergingAnketaData(false);
    }
  };

  const importPersonnelRosterWorkbook = async (file: File | undefined) => {
    if (!file) return;
    setIsLoading(true);
    try {
      setMessage(`Імпортую «${file.name}» у БД персоналу…`);
      const imported = await importStaffSheetFromFile(file);
      await reload({ force: true });
      holdStatusUntilRef.current = Date.now() + 20_000;
      setMessage(
        rosterImportStatusMessage({
          rows: imported.rows,
          personCount: imported.personCount,
          personCountInRoster: imported.personCountInRoster,
          personCountInArchive: imported.personCountInArchive,
        }),
      );
    } catch (error) {
      setMessage(rosterImportError(error));
    } finally {
      setIsLoading(false);
    }
  };

  return {
    isMergingAnketaData,
    isMergingVkTpvDovidky,
    importVkTpvDovidkyWorkbook,
    mergeMissingFieldsFromAnketaData,
    importPersonnelRosterWorkbook,
  };
};
