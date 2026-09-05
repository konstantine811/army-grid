import { useEffect, useRef, useState } from "react";
import {
  loadAnketaSheetFromFile,
  loadAnketaSheetPreferCache,
  persistAnketaSnapshot,
  type AnketaColumnKey,
  type AnketaSheetSnapshot,
} from "../anketaSheet";
import {
  applyAnketaEditsToSnapshot,
  bulkWriteAnketaCellEdits,
  countAnketaEdits,
  loadAnketaEdits,
  type AnketaEditsMap,
} from "../anketaEdits";
import { readAnketaAppsScriptUrl } from "../anketaGaps";
import { loadAnketaMissingNameKeys } from "../anketaMissingList";
import {
  expandAnketaNameKeySet,
  loadPersonnelIndexForAnketa,
  matchAnketaRowToPersonnel,
  normalizeAnketaExternalIdKey,
  normalizeAnketaNameKey,
} from "../anketaPersonMatch";
import {
  formatReconcileAnketaWithEjoosReport,
  loadEjoosOosAndExcludedPeople,
  reconcileAnketaSnapshotWithEjoos,
  type ReconcileAnketaWithEjoosReport,
} from "../anketaEjoosPeople";
import { readWorkbookSnapshot } from "../../../excelRoundTrip";
import {
  formatAnketaMilitaryIdMergeReport,
  mergeVkMilitaryIdToAnketa,
} from "../anketaMilitaryIdImport";
import {
  buildAnketaSourceFieldUpdates,
  buildPersonnelToAnketaFieldUpdates,
} from "../anketaPersonMerge";

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
  const [isImportingVkMilitaryIds, setIsImportingVkMilitaryIds] = useState(false);
  const lastEjoosReportRef = useRef<ReconcileAnketaWithEjoosReport | null>(
    null,
  );

  const rows = snapshot?.rows ?? [];

  const mergeWithEdits = async (base: AnketaSheetSnapshot) => {
    const edits = await loadAnketaEdits();
    setEditsCount(countAnketaEdits(edits));
    const withEdits = applyAnketaEditsToSnapshot(base, edits);
    const authority = await loadEjoosOosAndExcludedPeople();
    const reconciled = reconcileAnketaSnapshotWithEjoos(
      withEdits,
      authority,
    );
    lastEjoosReportRef.current = reconciled.report;
    return reconciled.snapshot;
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
        refreshGoogle: false,
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
      const ejoosSummary = lastEjoosReportRef.current
        ? ` · ${formatReconcileAnketaWithEjoosReport(lastEjoosReportRef.current)}`
        : "";
      setMessage(
        next.source === "cache" || next.source === "db"
          ? `Завантажено ${next.source === "db" ? "серверний snapshot" : "локальний кеш"} + правки + ЕЖООС (${next.rows.length})${ejoosSummary}.`
          : `Початкові дані взято з Google + правки + останній ЕЖООС · ${new Date(next.fetchedAt).toLocaleString("uk-UA")}${ejoosSummary}.`,
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
      await persistAnketaSnapshot(next);
      setSnapshot(next);
      setDirtyCount(0);
      setMessage(
        `Імпортовано файл ${file.name} і звірено з ЕЖООС: ${next.rows.length} анкет · ${lastEjoosReportRef.current ? formatReconcileAnketaWithEjoosReport(lastEjoosReportRef.current) : ""}.`,
      );
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

  const mergeFromPersonnel = async () => {
    if (!rows.length) {
      setMessage("Спочатку завантажте анкетні дані.");
      return;
    }
    setIsMergingPersonnel(true);
    try {
      setMessage("Завантажую актуальний Особовий склад…");
      const [personnelIndex, authority] = await Promise.all([
        loadPersonnelIndexForAnketa({ force: true }),
        loadEjoosOosAndExcludedPeople(),
      ]);
      let lastProgressAt = 0;
      let matched = 0;
      let updatedPeople = 0;
      const updates: Array<{
        rowNumber: number;
        columnId: AnketaColumnKey;
        value: string;
        externalId: string;
        fullName: string;
      }> = [];
      const updatesByRow = new Map<
        number,
        Record<string, string>
      >();

      for (let index = 0; index < rows.length; index += 1) {
        const row = rows[index]!;
        const match = matchAnketaRowToPersonnel(row, personnelIndex);
        const rowId = normalizeAnketaExternalIdKey(row.externalId);
        const rowName = normalizeAnketaNameKey(row.fullName);
        const byId = rowId
          ? authority.people.find(
              (person) =>
                normalizeAnketaExternalIdKey(person.personId) === rowId,
            )
          : undefined;
        const byName = authority.people.filter(
          (person) =>
            rowName &&
            normalizeAnketaNameKey(person.fullName) === rowName,
        );
        const authorityPerson = byId ?? (byName.length === 1 ? byName[0] : undefined);
        if (match || authorityPerson) {
          matched += 1;
          const personnelUpdates = match
            ? buildPersonnelToAnketaFieldUpdates(match.row, row).fieldUpdates
            : {};
          const authorityUpdates = authorityPerson
            ? buildAnketaSourceFieldUpdates(
                {
                  ...authorityPerson.anketaFields,
                  rank: authorityPerson.rank,
                  externalId: authorityPerson.personId,
                  positionIndex: authorityPerson.positionIndex,
                },
                row,
              ).fieldUpdates
            : {};
          const preview = {
            fieldUpdates: {
              ...personnelUpdates,
              ...authorityUpdates,
            },
          };
          const entries = Object.entries(preview.fieldUpdates);
          if (entries.length) {
            updatedPeople += 1;
            const rowUpdates: Record<string, string> = {};
            for (const [columnId, value] of entries) {
              if (!value) continue;
              rowUpdates[columnId] = value;
              updates.push({
                rowNumber: row.__rowNumber,
                columnId: columnId as AnketaColumnKey,
                value,
                externalId: row.externalId,
                fullName: row.fullName,
              });
            }
            updatesByRow.set(row.__rowNumber, rowUpdates);
          }
        }
        const done = index + 1;
        const now = Date.now();
        if (done === rows.length || now - lastProgressAt >= 250) {
          lastProgressAt = now;
          setMessage(`Особовий склад → анкети… ${done}/${rows.length}`);
        }
      }

      const result = await bulkWriteAnketaCellEdits(updates);
      persistSnapshot((current) => ({
        ...current,
        rows: current.rows.map((row) => {
          const patch = updatesByRow.get(row.__rowNumber);
          return patch ? Object.assign({ ...row }, patch) : row;
        }),
      }));
      setEditsCount(countAnketaEdits(result.edits));
      setMessage(
        `Особовий склад → анкети: зіставлено ${matched} · оновлено осіб ${updatedPeople} · полів ${updates.length}${
          result.serverSynced
            ? " · збережено в БД"
            : ` · локально: ок${result.serverError ? ` · сервер: ${result.serverError}` : ""}`
        }.`,
      );
    } catch (error) {
      setMessage(
        error instanceof Error
          ? error.message
          : "Не вдалося доповнити анкети з Особового складу.",
      );
    } finally {
      setIsMergingPersonnel(false);
    }
  };

  const addMissingFromEjoos = async () => {
    setIsAddingFromEjoos(true);
    try {
      setMessage("Звіряю анкети з останнім ЕЖООС…");
      const refreshed = await loadAnketaSheetPreferCache({
        refreshGoogle: false,
        mergeEdits: mergeWithEdits,
      });
      setSnapshot(refreshed);
      const edits = await loadAnketaEdits();
      setEditsCount(countAnketaEdits(edits));
      setMessage(
        `Анкети синхронізовано з ЕЖООС (ООС + Виключені) · ${
          lastEjoosReportRef.current
            ? formatReconcileAnketaWithEjoosReport(
                lastEjoosReportRef.current,
              )
            : refreshed.rows.length
        }.`,
      );
    } catch (error) {
      setMessage(
        error instanceof Error
          ? error.message
          : "Не вдалося синхронізувати анкети з ЕЖООС.",
      );
    } finally {
      setIsAddingFromEjoos(false);
    }
  };

  const importVkMilitaryIdsFile = async (file: File | undefined) => {
    if (!file) return;
    setIsImportingVkMilitaryIds(true);
    setMessage(`Читаю «${file.name}»…`);
    try {
      const workbook = await readWorkbookSnapshot(file);
      let lastProgressAt = 0;
      const report = await mergeVkMilitaryIdToAnketa(workbook, {
        anketaRows: rows,
        onProgress: (done, total) => {
          const now = Date.now();
          if (done !== total && now - lastProgressAt < 250) return;
          lastProgressAt = now;
          setMessage(`Військові квитки · ${done}/${total}`);
        },
      });
      const refreshed = await loadAnketaSheetPreferCache({
        refreshGoogle: false,
        mergeEdits: mergeWithEdits,
      });
      setSnapshot(refreshed);
      const edits = await loadAnketaEdits();
      setEditsCount(countAnketaEdits(edits));
      setMessage(
        `Військові квитки · ${formatAnketaMilitaryIdMergeReport(report)}.`,
      );
    } catch (error) {
      setMessage(
        error instanceof Error
          ? error.message
          : "Не вдалося імпортувати військові квитки.",
      );
    } finally {
      setIsImportingVkMilitaryIds(false);
    }
  };

  useEffect(() => {
    void loadFromGoogle();
    void loadAnketaEdits().then((edits: AnketaEditsMap) => {
      setEditsCount(countAnketaEdits(edits));
    });
    void loadAnketaMissingNameKeys()
      .then(({ keys }) => {
        setMissingQuestionnaireNames(expandAnketaNameKeySet(keys));
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
    setMissingQuestionnaireNames,
    isMergingPersonnel,
    isAddingFromEjoos,
    isImportingVkMilitaryIds,
    mergeWithEdits,
    persistSnapshot,
    loadFromGoogle,
    loadFromCsvFile,
    mergeFromPersonnel,
    addMissingFromEjoos,
    importVkMilitaryIdsFile,
  };
}
