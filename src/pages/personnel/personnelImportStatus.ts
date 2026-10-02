import { getRosterPersonName } from "./personnelRosterMerge";
import { normalizeRosterMatchText } from "./fighterStatusImport";
import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";

export const rosterImportStatusMessage = ({
  rows,
  personCount,
  personCountInRoster,
  personCountInArchive,
}: {
  rows: EjournalPreviewRow[];
  personCount: number;
  personCountInRoster?: number;
  personCountInArchive?: number;
}) => {
  const hasTsapenko = rows.some((row) =>
    normalizeRosterMatchText(getRosterPersonName(row)).includes("цапенко"),
  );
  const rosterCount = personCountInRoster ?? personCount;
  const archiveCount = personCountInArchive ?? 0;
  return `Штатку імпортовано: ${rosterCount} у штаті${
    archiveCount ? ` · ${archiveCount} архів` : ""
  } · усього ${personCount} · ЦАПЕНКО: ${
    hasTsapenko ? "знайдено у файлі" : "у файлі не знайдено"
  }.`;
};

export const vkTpvImportError = (error: unknown) =>
  error instanceof Error
    ? error.message
    : "Не вдалося імпортувати ВК ТПВ ДОВІДКИ.";

export const anketaMergeError = (error: unknown) =>
  error instanceof Error
    ? error.message
    : "Не вдалося доповнити особовий склад з анкет.";

export const rosterImportError = (error: unknown) =>
  error instanceof Error
    ? error.message
    : "Не вдалося імпортувати файл «Штатка» в БД.";
