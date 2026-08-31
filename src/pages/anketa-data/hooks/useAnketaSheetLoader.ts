import { useEffect, useState } from "react";
import {
  loadAnketaSheetFromFile,
  loadAnketaSheetPreferCache,
  type AnketaSheetSnapshot,
} from "../anketaSheet";
import {
  applyAnketaEditsToSnapshot,
  countAnketaEdits,
  loadAnketaEdits,
  type AnketaEditsMap,
} from "../anketaEdits";
import { readAnketaAppsScriptUrl } from "../anketaGaps";
import { loadAnketaMissingNameKeys } from "../anketaMissingList";
import {
  addMissingEjoosPeopleToCachedAnketa,
  formatAddMissingEjoosPeopleReport,
} from "../anketaEjoosPeople";
import {
  formatAnketaBulkMergeReport,
  mergeCachedAnketaToPersonnel,
} from "../anketaPersonMerge";
import {
  buildStaffSheetEnrichmentEntries,
  formatStaffSheetEnrichmentReport,
  pushPersonnelEnrichmentToStaffSheet,
} from "../staffSheetEnrichment";
import {
  downloadEnrichedStaffSheetExcel,
  formatStaffSheetEnrichmentReport as formatStaffExportReport,
} from "../staffSheetEnrichedExport";
import {
  formatStaffSheetImportSummary,
  importStaffSheetFromFile,
  loadStaffSheetImport,
} from "../staffSheetImport";
import { loadStaffSheetEnrichmentContext } from "../staffSheetEnrichmentContext";

export function useAnketaSheetLoader() {
  const [snapshot, setSnapshot] = useState<AnketaSheetSnapshot | null>(null);
  const [isLoading, setIsLoading] = useState(false);
  const [isSyncing, setIsSyncing] = useState(false);
  const [message, setMessage] = useState(
    "Завантажте анкетні дані з Google Sheets.",
  );
  const [dirtyCount, setDirtyCount] = useState(0);
  const [editsCount, setEditsCount] = useState(0);
  const [appsScriptUrl, setAppsScriptUrl] = useState(() =>
    readAnketaAppsScriptUrl(),
  );
  const [missingQuestionnaireNames, setMissingQuestionnaireNames] = useState<
    Set<string>
  >(() => new Set());
  const [isMergingPersonnel, setIsMergingPersonnel] = useState(false);
  const [isAddingFromEjoos, setIsAddingFromEjoos] = useState(false);
  const [isPushingStaffSheet, setIsPushingStaffSheet] = useState(false);
  const [isDownloadingStaffSheet, setIsDownloadingStaffSheet] = useState(false);
  const [staffSheetImportName, setStaffSheetImportName] = useState("");
  const [staffSheetImportedAt, setStaffSheetImportedAt] = useState("");
  const [staffSheetSource, setStaffSheetSource] = useState<"file" | "">("");
  const [staffSheetPersonCount, setStaffSheetPersonCount] = useState(0);
  const [isImportingStaffSheet, setIsImportingStaffSheet] = useState(false);

  const applyStaffSheetImportMeta = (imported: {
    fileName: string;
    importedAt: string;
    personCount?: number;
    rows?: unknown[];
  }) => {
    setStaffSheetImportName(imported.fileName);
    setStaffSheetImportedAt(imported.importedAt);
    setStaffSheetSource("file");
    setStaffSheetPersonCount(
      imported.personCount ||
        (Array.isArray(imported.rows) ? imported.rows.length : 0),
    );
  };

  const rows = snapshot?.rows ?? [];

  const mergeWithEdits = async (base: AnketaSheetSnapshot) => {
    const edits = await loadAnketaEdits();
    setEditsCount(countAnketaEdits(edits));
    return applyAnketaEditsToSnapshot(base, edits);
  };

  const persistSnapshot = (
    updater: (current: AnketaSheetSnapshot) => AnketaSheetSnapshot,
    note?: string,
  ) => {
    setSnapshot((current) => {
      if (!current) return current;
      return updater(current);
    });
    setDirtyCount((value) => value + 1);
    if (note) setMessage(note);
  };

  const resetGapFocus = () => {
    return { clearSearch: true as const };
  };

  const loadFromGoogle = async () => {
    setIsLoading(true);
    try {
      const next = await loadAnketaSheetPreferCache({
        mergeEdits: mergeWithEdits,
        onCached: (cached) => {
          setSnapshot(cached);
          setIsLoading(false);
          const origin =
            cached.source === "db"
              ? "БД"
              : cached.source === "cache"
                ? "локальний кеш"
                : "кеш";
          setMessage(
            `${origin} + правки · ${cached.rows.length} анкет · оновлюю з Google…`,
          );
        },
      });
      setSnapshot(next);
      setDirtyCount(0);
      setMessage(
        next.source === "cache" || next.source === "db"
          ? `Показано ${next.source === "db" ? "БД" : "кеш"} + правки (${next.rows.length}). Мережа Google недоступна.`
          : `Оновлено з Google + накладено правки · ${new Date(next.fetchedAt).toLocaleString("uk-UA")}.`,
      );
      return resetGapFocus();
    } catch (error) {
      setMessage(
        error instanceof Error
          ? `Не вдалося завантажити таблицю: ${error.message}`
          : "Не вдалося завантажити Google Sheets.",
      );
      return null;
    } finally {
      setIsLoading(false);
    }
  };

  const loadFromCsvFile = async (file: File | undefined) => {
    if (!file) return null;
    setIsLoading(true);
    try {
      const base = await loadAnketaSheetFromFile(file);
      const next = await mergeWithEdits(base);
      setSnapshot(next);
      setDirtyCount(0);
      setMessage(`Імпортовано файл ${file.name}: ${next.rows.length} анкет.`);
      return resetGapFocus();
    } catch (error) {
      setMessage(
        error instanceof Error
          ? error.message
          : "Не вдалося прочитати CSV файл.",
      );
      return null;
    } finally {
      setIsLoading(false);
    }
  };

  const mergeToPersonnel = async () => {
    if (!rows.length) {
      setMessage("Спочатку завантажте анкетні дані.");
      return;
    }
    setIsMergingPersonnel(true);
    try {
      let lastProgressAt = 0;
      const report = await mergeCachedAnketaToPersonnel({
        onProgress: (done, total) => {
          const now = Date.now();
          if (done !== total && now - lastProgressAt < 250) return;
          lastProgressAt = now;
          setMessage(`Злиття з особовим складом… ${done}/${total}`);
        },
      });
      setMessage(
        `Злиття з особовим складом завершено · ${formatAnketaBulkMergeReport(report)}.`,
      );
    } catch (error) {
      setMessage(
        error instanceof Error
          ? error.message
          : "Не вдалося оновити особовий склад з анкет.",
      );
    } finally {
      setIsMergingPersonnel(false);
    }
  };

  const addMissingFromEjoos = async () => {
    setIsAddingFromEjoos(true);
    try {
      let lastProgressAt = 0;
      const report = await addMissingEjoosPeopleToCachedAnketa({
        onProgress: (done, total) => {
          const now = Date.now();
          if (done !== total && now - lastProgressAt < 250) return;
          lastProgressAt = now;
          setMessage(`Додаю відсутніх з ЕЖООС… ${done}/${total}`);
        },
      });
      const refreshed = await loadAnketaSheetPreferCache({
        mergeEdits: mergeWithEdits,
      });
      setSnapshot(refreshed);
      const edits = await loadAnketaEdits();
      setEditsCount(countAnketaEdits(edits));
      setMessage(
        report.added
          ? `Додано з ЕЖООС (ООС + Виключені) · ${formatAddMissingEjoosPeopleReport(report)}.`
          : `Нових осіб немає · ${formatAddMissingEjoosPeopleReport(report)}.`,
      );
    } catch (error) {
      setMessage(
        error instanceof Error
          ? error.message
          : "Не вдалося додати осіб з ЕЖООС.",
      );
    } finally {
      setIsAddingFromEjoos(false);
    }
  };

  const importStaffSheetFile = async (file: File | undefined) => {
    if (!file) return;
    setIsImportingStaffSheet(true);
    try {
      setMessage("Імпортую «Штатку» (.xlsx) у кеш і БД персоналу…");
      const imported = await importStaffSheetFromFile(file);
      applyStaffSheetImportMeta(imported);

      setMessage("Зіставляю з анкетами…");
      const context = await loadStaffSheetEnrichmentContext();
      const entries = buildStaffSheetEnrichmentEntries({
        rosterRows: context.rosterRows,
        mergedPersonnelRows: context.mergedPersonnelRows,
        anketaRows: context.anketaRows,
        questionnaires: context.questionnaires,
      });
      const report = {
        rows: entries.length,
        anketaYes: entries.filter((entry) => entry.values[0] === "так").length,
        withBirthDate: entries.filter((entry) =>
          Boolean(String(entry.values[1] ?? "").trim()),
        ).length,
        withInn: entries.filter(
          (entry) => entry.values[4]?.replace(/\D/g, "").length === 10,
        ).length,
      };

      setMessage(
        `Імпортовано «Штатку»: ${formatStaffSheetImportSummary(imported)} · зіставлення: ${formatStaffSheetEnrichmentReport(report)}. Далі — «Скачати Штатку» або «→ Google Штатка». Цей файл також доступний для ранкового звіту.`,
      );
    } catch (error) {
      setMessage(
        error instanceof Error
          ? error.message
          : "Не вдалося імпортувати файл «Штатка».",
      );
    } finally {
      setIsImportingStaffSheet(false);
    }
  };

  const pushStaffSheetEnrichment = async () => {
    setIsPushingStaffSheet(true);
    try {
      let phase = "";
      const report = await pushPersonnelEnrichmentToStaffSheet({
        onProgress: (nextPhase) => {
          phase = nextPhase;
          setMessage(nextPhase);
        },
      });
      setMessage(
        `${phase || "Google Sheet «Штатка»"} · ${formatStaffSheetEnrichmentReport(report)}.`,
      );
    } catch (error) {
      setMessage(
        error instanceof Error
          ? error.message
          : "Не вдалося оновити Google Sheet «Штатка».",
      );
    } finally {
      setIsPushingStaffSheet(false);
    }
  };

  const downloadStaffSheetExcel = async () => {
    setIsDownloadingStaffSheet(true);
    try {
      let phase = "";
      const report = await downloadEnrichedStaffSheetExcel({
        onProgress: (nextPhase) => {
          phase = nextPhase;
          setMessage(nextPhase);
        },
      });
      setMessage(
        `Завантажено Excel · ${formatStaffExportReport(report)}${phase ? ` · ${phase}` : ""}`,
      );
    } catch (error) {
      setMessage(
        error instanceof Error
          ? error.message
          : "Не вдалося сформувати Excel «Штатка».",
      );
    } finally {
      setIsDownloadingStaffSheet(false);
    }
  };

  useEffect(() => {
    void loadFromGoogle();
    void loadStaffSheetImport().then((imported) => {
      if (!imported) return;
      applyStaffSheetImportMeta(imported);
    });
    void loadAnketaEdits().then((edits: AnketaEditsMap) => {
      setEditsCount(countAnketaEdits(edits));
    });
    void loadAnketaMissingNameKeys()
      .then(({ keys }) => {
        setMissingQuestionnaireNames(keys);
      })
      .catch(() => {
        /* пошук працює без списку відсутніх анкет */
      });
  }, []);

  return {
    snapshot,
    rows,
    isLoading,
    isSyncing,
    setIsSyncing,
    message,
    setMessage,
    dirtyCount,
    editsCount,
    setEditsCount,
    appsScriptUrl,
    setAppsScriptUrl,
    missingQuestionnaireNames,
    isMergingPersonnel,
    isAddingFromEjoos,
    isPushingStaffSheet,
    isDownloadingStaffSheet,
    isImportingStaffSheet,
    staffSheetImportName,
    staffSheetImportedAt,
    staffSheetSource,
    staffSheetPersonCount,
    mergeWithEdits,
    persistSnapshot,
    loadFromGoogle,
    loadFromCsvFile,
    mergeToPersonnel,
    addMissingFromEjoos,
    pushStaffSheetEnrichment,
    downloadStaffSheetExcel,
    importStaffSheetFile,
  };
}
