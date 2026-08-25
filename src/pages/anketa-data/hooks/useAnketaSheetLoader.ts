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
import { mergeAnketaRowsToPersonnel } from "../anketaPersonMerge";

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
      const edits = await loadAnketaEdits();
      let lastProgressAt = 0;
      const report = await mergeAnketaRowsToPersonnel({
        rows,
        edits,
        onProgress: (done, total) => {
          const now = Date.now();
          if (done !== total && now - lastProgressAt < 250) return;
          lastProgressAt = now;
          setMessage(`Злиття з особовим складом… ${done}/${total}`);
        },
      });
      const parts = [
        `оновлено осіб: ${report.updated}`,
        `полів: ${report.fieldCount}`,
        report.phonesAdded ? `телефонів: ${report.phonesAdded}` : "",
        report.skippedNoMatch ? `без збігу: ${report.skippedNoMatch}` : "",
        report.skippedNoUpdates ? `без нових даних: ${report.skippedNoUpdates}` : "",
        report.skippedNoRowId ? `лише ранковий список: ${report.skippedNoRowId}` : "",
        report.errors.length ? `помилок: ${report.errors.length}` : "",
      ].filter(Boolean);
      setMessage(`Злиття з особовим складом завершено · ${parts.join(" · ")}.`);
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

  useEffect(() => {
    void loadFromGoogle();
    void loadAnketaEdits().then((edits: AnketaEditsMap) => {
      setEditsCount(countAnketaEdits(edits));
    });
    void loadAnketaMissingNameKeys()
      .then((keys) => {
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
    mergeWithEdits,
    persistSnapshot,
    loadFromGoogle,
    loadFromCsvFile,
    mergeToPersonnel,
  };
}
