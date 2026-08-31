import { api } from "../../api";
import {
  EJOOS_SYNC_READ_OPTIONS,
  readWorkbookSnapshot,
} from "../../excelRoundTrip";
import {
  findEjoosSheet,
  parseEjoosExcluded,
  parseEjoosOos,
} from "../ejournal/ejoosParsers";
import { base64ToFile } from "../ejournal/ejoosSyncApply";
import {
  applyAnketaEditsToRows,
  bulkWriteAnketaCellEdits,
  loadAnketaEdits,
  type AnketaEditsMap,
} from "./anketaEdits";
import {
  normalizeAnketaExternalIdKey,
  normalizeAnketaNameKey,
} from "./anketaPersonMatch";
import {
  loadAnketaSheetPreferCache,
  type AnketaColumnKey,
  type AnketaRow,
} from "./anketaSheet";

export type EjoosAnketaCandidate = {
  personId: string;
  fullName: string;
  rank: string;
  positionIndex: string;
  source: "oos" | "excluded";
};

export type AddMissingEjoosPeopleReport = {
  ejoosPeople: number;
  alreadyInAnketa: number;
  added: number;
  skippedNoName: number;
  errors: Array<{ name: string; message: string }>;
};

const isUsableEjoosName = (value: string) => {
  const key = normalizeAnketaNameKey(value);
  if (!key || key.length < 4) return false;
  if (key.includes("вакан") || key === "особа не вибрана") return false;
  return /[\p{L}]/u.test(key);
};

const personDedupeKey = (personId: string, fullName: string) => {
  const id = normalizeAnketaExternalIdKey(personId);
  if (id) return `id:${id}`;
  const name = normalizeAnketaNameKey(fullName);
  return name ? `name:${name}` : "";
};

const loadEjoosOosAndExcludedPeople = async (): Promise<
  EjoosAnketaCandidate[]
> => {
  const live = await api.getEjournalLive("1ПБ");
  if (!live.current?.id) {
    throw new Error(
      "Немає збереженого ЕЖООС у БД. Завантажте журнал на сторінці ЄЖООС.",
    );
  }

  const full = await api.getEjournalLiveFile(live.current.id, "1ПБ");
  if (!full.fileBase64) {
    throw new Error("У поточній версії ЕЖООС немає файлу.");
  }

  const file = base64ToFile(
    full.fileBase64,
    full.sourceFileName || "ЕЖООС.xlsx",
  );
  const workbook = await readWorkbookSnapshot(file, EJOOS_SYNC_READ_OPTIONS);
  const oosSheet = findEjoosSheet(workbook, /(^|[.\s])оос($|[\s])/i);
  const excludedSheet = findEjoosSheet(workbook, /виключен/i);

  if (!oosSheet && !excludedSheet) {
    throw new Error("У збереженому ЕЖООС немає аркушів «ООС» або «Виключені».");
  }

  const byKey = new Map<string, EjoosAnketaCandidate>();
  const push = (candidate: EjoosAnketaCandidate) => {
    if (!isUsableEjoosName(candidate.fullName) && !candidate.personId) return;
    if (!isUsableEjoosName(candidate.fullName)) return;
    const key = personDedupeKey(candidate.personId, candidate.fullName);
    if (!key || byKey.has(key)) return;
    byKey.set(key, candidate);
  };

  for (const row of parseEjoosOos(oosSheet)) {
    push({
      personId: row.personId.trim(),
      fullName: row.fullName.trim(),
      rank: row.rank.trim(),
      positionIndex: row.positionIndex.trim(),
      source: "oos",
    });
  }

  for (const row of parseEjoosExcluded(excludedSheet)) {
    push({
      personId: row.personId.trim(),
      fullName: row.fullName.trim(),
      rank: row.rank.trim(),
      positionIndex: row.positionIndex.trim(),
      source: "excluded",
    });
  }

  return [...byKey.values()];
};

const anketaHasPerson = (
  rows: AnketaRow[],
  person: EjoosAnketaCandidate,
) => {
  const personId = normalizeAnketaExternalIdKey(person.personId);
  const personName = normalizeAnketaNameKey(person.fullName);
  for (const row of rows) {
    const rowId = normalizeAnketaExternalIdKey(row.externalId);
    if (personId && rowId && personId === rowId) return true;
    const rowName = normalizeAnketaNameKey(row.fullName);
    if (personName && rowName && personName === rowName) return true;
  }
  return false;
};

const nextAnketaRowNumber = (rows: AnketaRow[], edits: AnketaEditsMap) => {
  let maxRow = 1;
  for (const row of rows) {
    if (row.__rowNumber > maxRow) maxRow = row.__rowNumber;
  }
  for (const edit of Object.values(edits)) {
    if (edit.rowNumber > maxRow) maxRow = edit.rowNumber;
  }
  return maxRow + 1;
};

export const formatAddMissingEjoosPeopleReport = (
  report: AddMissingEjoosPeopleReport,
) => {
  const parts = [
    `у ЕЖООС: ${report.ejoosPeople}`,
    `уже в анкетах: ${report.alreadyInAnketa}`,
    `додано: ${report.added}`,
    report.skippedNoName ? `без ПІБ: ${report.skippedNoName}` : "",
    report.errors.length ? `помилок: ${report.errors.length}` : "",
  ].filter(Boolean);
  return parts.join(" · ");
};

/** Додати в анкети лише тих, кого ще немає; джерело — ООС і Виключені збереженого ЕЖООС. */
export const addMissingEjoosPeopleToCachedAnketa = async (options?: {
  onProgress?: (done: number, total: number) => void;
}): Promise<AddMissingEjoosPeopleReport> => {
  const snapshot = await loadAnketaSheetPreferCache();
  const edits = await loadAnketaEdits();
  const rows =
    edits && Object.keys(edits).length
      ? applyAnketaEditsToRows(snapshot.rows, edits)
      : snapshot.rows;

  const people = await loadEjoosOosAndExcludedPeople();
  const report: AddMissingEjoosPeopleReport = {
    ejoosPeople: people.length,
    alreadyInAnketa: 0,
    added: 0,
    skippedNoName: 0,
    errors: [],
  };

  const toAdd: EjoosAnketaCandidate[] = [];
  for (const person of people) {
    if (!isUsableEjoosName(person.fullName)) {
      report.skippedNoName += 1;
      continue;
    }
    if (anketaHasPerson(rows, person)) {
      report.alreadyInAnketa += 1;
      continue;
    }
    toAdd.push(person);
  }

  if (!toAdd.length) return report;

  let rowNumber = nextAnketaRowNumber(rows, edits);
  const items: Array<{
    rowNumber: number;
    columnId: AnketaColumnKey;
    value: string;
    externalId?: string;
    fullName?: string;
  }> = [];

  for (let index = 0; index < toAdd.length; index += 1) {
    const person = toAdd[index]!;
    options?.onProgress?.(index + 1, toAdd.length);
    const fields: Array<[AnketaColumnKey, string]> = [
      ["fullName", person.fullName],
      ["rank", person.rank],
      ["externalId", person.personId],
      ["positionIndex", person.positionIndex],
    ];
    for (const [columnId, value] of fields) {
      if (!value) continue;
      items.push({
        rowNumber,
        columnId,
        value,
        externalId: person.personId || undefined,
        fullName: person.fullName,
      });
    }
    rowNumber += 1;
    report.added += 1;
  }

  try {
    await bulkWriteAnketaCellEdits(items);
  } catch (error) {
    report.added = 0;
    report.errors.push({
      name: "ЕЖООС → анкети",
      message:
        error instanceof Error
          ? error.message
          : "Не вдалося записати нових осіб в анкети",
    });
  }

  return report;
};
