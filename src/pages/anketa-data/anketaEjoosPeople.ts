import { api } from "../../api";
import {
  EJOOS_SYNC_READ_OPTIONS,
  readWorkbookSnapshot,
  type ExcelWorkbookSnapshot,
} from "../../excelRoundTrip";
import {
  findEjoosSheet,
  parseEjoosExcluded,
  parseEjoosOos,
} from "../ejournal/ejoosParsers";
import { base64ToFile } from "../ejournal/ejoosSyncApply";
import {
  normalizeAnketaExternalIdKey,
  normalizeAnketaNameKey,
} from "./anketaPersonMatch";
import {
  ANKETA_COLUMNS,
  createEmptyAnketaRow,
  type AnketaRow,
  type AnketaSheetSnapshot,
} from "./anketaSheet";

export type EjoosAnketaCandidate = {
  personId: string;
  fullName: string;
  rank: string;
  positionIndex: string;
  source: "oos" | "excluded";
};

export type ReconcileAnketaWithEjoosReport = {
  ejoosPeople: number;
  kept: number;
  merged: number;
  added: number;
  removed: number;
  skippedNoName: number;
  conflicts: Array<{ name: string; message: string }>;
};

export type EjoosAnketaAuthority = {
  versionId: string;
  version: number;
  sourceFileName: string;
  people: EjoosAnketaCandidate[];
};

let cachedAuthority: EjoosAnketaAuthority | null = null;

const isUsableEjoosName = (value: string) => {
  const key = normalizeAnketaNameKey(value);
  if (!key || key.length < 4) return false;
  if (key.includes("вакан") || key === "особа не вибрана") return false;
  if (/вибув|розпорядж|продовольч|підрозділ/u.test(key)) return false;
  return /[\p{L}]/u.test(key);
};

export const collectEjoosOosAndExcludedPeople = (
  workbook: ExcelWorkbookSnapshot,
) => {
  const oosSheet = findEjoosSheet(workbook, /(^|[.\s])оос($|[\s])/i);
  const excludedSheet = findEjoosSheet(workbook, /виключен/i);
  if (!oosSheet && !excludedSheet) {
    throw new Error("У збереженому ЕЖООС немає аркушів «ООС» або «Виключені».");
  }

  const people: EjoosAnketaCandidate[] = [];
  const push = (candidate: EjoosAnketaCandidate) => {
    if (!isUsableEjoosName(candidate.fullName)) return;
    const id = normalizeAnketaExternalIdKey(candidate.personId);
    const name = normalizeAnketaNameKey(candidate.fullName);
    const existing = people.find((person) => {
      const existingId = normalizeAnketaExternalIdKey(person.personId);
      const existingName = normalizeAnketaNameKey(person.fullName);
      if (id && existingId) return id === existingId;
      return Boolean(name && existingName && name === existingName);
    });
    if (!existing) {
      people.push(candidate);
      return;
    }
    existing.personId ||= candidate.personId;
    existing.fullName ||= candidate.fullName;
    existing.rank ||= candidate.rank;
    existing.positionIndex ||= candidate.positionIndex;
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
  return people;
};

export const loadEjoosOosAndExcludedPeople =
  async (): Promise<EjoosAnketaAuthority> => {
  const live = await api.getEjournalLive("1ПБ");
  if (!live.current?.id) {
    throw new Error(
      "Немає збереженого ЕЖООС у БД. Завантажте журнал на сторінці ЄЖООС.",
    );
  }
  if (cachedAuthority?.versionId === live.current.id) {
    return cachedAuthority;
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
  const people = collectEjoosOosAndExcludedPeople(workbook);

  cachedAuthority = {
    versionId: live.current.id,
    version: live.current.version,
    sourceFileName: full.sourceFileName || "ЕЖООС.xlsx",
    people,
  };
  return cachedAuthority;
};

const candidateFields = (person: EjoosAnketaCandidate) => ({
  fullName: person.fullName.trim(),
  externalId: person.personId.trim(),
  rank: person.rank.trim(),
  positionIndex: person.positionIndex.trim(),
});

const mergeCandidateIntoAnketa = (
  row: AnketaRow,
  person: EjoosAnketaCandidate,
) => {
  const next = { ...row };
  let changed = false;
  for (const [key, value] of Object.entries(candidateFields(person))) {
    const column = key as keyof ReturnType<typeof candidateFields>;
    if (!String(next[column] || "").trim() && value) {
      next[column] = value;
      changed = true;
    }
  }
  return { row: next, changed };
};

export const reconcileAnketaRowsWithEjoos = (
  rows: AnketaRow[],
  people: EjoosAnketaCandidate[],
): { rows: AnketaRow[]; report: ReconcileAnketaWithEjoosReport } => {
  const report: ReconcileAnketaWithEjoosReport = {
    ejoosPeople: people.length,
    kept: 0,
    merged: 0,
    added: 0,
    removed: 0,
    skippedNoName: 0,
    conflicts: [],
  };
  const usedRows = new Set<AnketaRow>();
  const reconciled: AnketaRow[] = [];
  let nextRowNumber =
    rows.reduce((max, row) => Math.max(max, row.__rowNumber), 1) + 1;

  for (const person of people) {
    const personId = normalizeAnketaExternalIdKey(person.personId);
    const personName = normalizeAnketaNameKey(person.fullName);
    if (!personName) {
      report.skippedNoName += 1;
      continue;
    }
    const idMatches = personId
      ? rows.filter(
          (row) =>
            !usedRows.has(row) &&
            normalizeAnketaExternalIdKey(row.externalId) === personId,
        )
      : [];
    const nameMatches = rows.filter(
      (row) =>
        !usedRows.has(row) &&
        normalizeAnketaNameKey(row.fullName) === personName,
    );
    let matched = idMatches.length === 1 ? idMatches[0] : null;

    if (idMatches.length > 1) {
      report.conflicts.push({
        name: person.fullName,
        message: `В анкетах знайдено ${idMatches.length} рядки з ID ${person.personId}`,
      });
      matched = idMatches[0] ?? null;
    }
    if (!matched && nameMatches.length === 1) {
      const byName = nameMatches[0]!;
      const rowId = normalizeAnketaExternalIdKey(byName.externalId);
      if (personId && rowId && personId !== rowId) {
        report.conflicts.push({
          name: person.fullName,
          message: `Однаковий ПІБ має різні ID: ЕЖООС ${person.personId}, анкета ${byName.externalId}`,
        });
      } else {
        matched = byName;
      }
    } else if (!matched && nameMatches.length > 1) {
      report.conflicts.push({
        name: person.fullName,
        message: `В анкетах знайдено ${nameMatches.length} рядки з однаковим ПІБ`,
      });
    }

    if (matched) {
      const matchedName = normalizeAnketaNameKey(matched.fullName);
      if (matchedName && matchedName !== personName) {
        report.conflicts.push({
          name: person.fullName,
          message: `ID ${person.personId} належить іншому ПІБ в анкеті: ${matched.fullName}`,
        });
        usedRows.add(matched);
        reconciled.push(matched);
        report.kept += 1;
        continue;
      }
      usedRows.add(matched);
      const merged = mergeCandidateIntoAnketa(matched, person);
      reconciled.push(merged.row);
      if (merged.changed) report.merged += 1;
      else report.kept += 1;
      continue;
    }

    const created = createEmptyAnketaRow(nextRowNumber);
    Object.assign(created, candidateFields(person));
    created.__rowId = `anketa-${nextRowNumber}-${created.externalId || created.fullName}`;
    reconciled.push(created);
    nextRowNumber += 1;
    report.added += 1;
  }

  report.removed = rows.length - usedRows.size;
  return { rows: reconciled, report };
};

export const reconcileAnketaSnapshotWithEjoos = (
  snapshot: AnketaSheetSnapshot,
  authority: EjoosAnketaAuthority,
) => {
  const result = reconcileAnketaRowsWithEjoos(snapshot.rows, authority.people);
  return {
    snapshot: {
      ...snapshot,
      columns: ANKETA_COLUMNS,
      rows: result.rows,
      sourceLabel: `${snapshot.sourceLabel.replace(/\s*·\s*ЕЖООС v\d+$/u, "")} · ЕЖООС v${authority.version}`,
      ejoosVersionId: authority.versionId,
      reconciledAt: new Date().toISOString(),
    } as AnketaSheetSnapshot,
    report: result.report,
  };
};

export const formatReconcileAnketaWithEjoosReport = (
  report: ReconcileAnketaWithEjoosReport,
) => {
  const parts = [
    `у ЕЖООС: ${report.ejoosPeople}`,
    `залишено: ${report.kept}`,
    `доповнено: ${report.merged}`,
    `додано: ${report.added}`,
    `прибрано: ${report.removed}`,
    report.skippedNoName ? `без ПІБ: ${report.skippedNoName}` : "",
    report.conflicts.length ? `конфліктів: ${report.conflicts.length}` : "",
  ].filter(Boolean);
  return parts.join(" · ");
};
