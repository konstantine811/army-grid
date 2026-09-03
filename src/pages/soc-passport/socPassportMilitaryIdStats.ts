import type { ExcelWorkbookSnapshot } from "../../excelRoundTrip";
import {
  extractBchsAwayPeopleFromSheet,
  hasBchsFullName,
  isBchsPersonnelGeneralListSheet,
} from "../bchs/bchsCalc";
import { ANKETA_ABSENT_QUESTIONNAIRE_VALUE } from "../anketa-data/anketaGaps";
import { ANKETA_MILITARY_ID_ABSENT_VALUE } from "../anketa-data/anketaMilitaryIdImport";
import { ANKETA_MISSING_VALUE_PRESETS } from "../anketa-data/anketaGaps";
import { normalizeAnketaNameKey } from "../anketa-data/anketaPersonMatch";
import type { AnketaRow } from "../anketa-data/anketaSheet";
import { extractMilitaryIdFromText } from "../personnel/vkTpvDovidkyImport";
import { findMorningRosterSheet } from "./socPassportParse";
import { looksLikePersonName } from "./socPassportFields";

export type AnketaFieldStatus =
  | "has_value"
  | "absent"
  | "empty"
  | "no_anketa";

export type StaffFieldSummary = {
  total: number;
  has: number;
  absent: number;
  empty: number;
  noAnketa: number;
};

export type MilitaryIdPersonRow = {
  fullName: string;
  rank: string;
  status: AnketaFieldStatus;
  militaryId: string;
};

export type DocumentPersonRow = {
  fullName: string;
  rank: string;
  status: AnketaFieldStatus;
  documentNumber: string;
  documentName: string;
};

export type AnketaStaffStatsResult = {
  year: number;
  militaryId: {
    staff: StaffFieldSummary;
    anketa: StaffFieldSummary;
    staffPeople: MilitaryIdPersonRow[];
  };
  documents: {
    staff: StaffFieldSummary;
    anketa: StaffFieldSummary;
    staffPeople: DocumentPersonRow[];
  };
};

/** @deprecated use AnketaStaffStatsResult */
export type MilitaryIdStatsResult = AnketaStaffStatsResult;

const ABSENT_VALUES = new Set<string>([
  ANKETA_MILITARY_ID_ABSENT_VALUE,
  "квиток відсутній",
  ANKETA_ABSENT_QUESTIONNAIRE_VALUE,
  "відсутній",
  "немає",
]);

const PLACEHOLDER_VALUES = new Set<string>(
  ANKETA_MISSING_VALUE_PRESETS.map((value) => value.toLocaleLowerCase("uk-UA")),
);

const isAbsentText = (value: string) =>
  ABSENT_VALUES.has(value.trim().toLocaleLowerCase("uk-UA"));

const isPlaceholderText = (value: string) =>
  PLACEHOLDER_VALUES.has(value.trim().toLocaleLowerCase("uk-UA"));

const looksLikeDocumentNumber = (value: string) => {
  const compact = value.replace(/\s+/g, "");
  if (compact.length < 4) return false;
  return /[\dA-Za-zА-ЯІЇЄҐ]/u.test(compact);
};

const looksLikeDocumentName = (value: string) => {
  const trimmed = value.trim();
  if (trimmed.length < 3) return false;
  if (isAbsentText(trimmed) || isPlaceholderText(trimmed)) return false;
  return true;
};

export const classifyAnketaMilitaryId = (
  value: unknown,
): { status: AnketaFieldStatus; militaryId: string } => {
  const raw = String(value ?? "").trim();
  const extracted = extractMilitaryIdFromText(raw);
  if (extracted) {
    return { status: "has_value", militaryId: extracted };
  }
  if (!raw) {
    return { status: "empty", militaryId: "" };
  }
  if (isAbsentText(raw)) {
    return { status: "absent", militaryId: raw };
  }
  if (isPlaceholderText(raw)) {
    return { status: "empty", militaryId: raw };
  }
  return { status: "empty", militaryId: raw };
};

export const classifyAnketaDocument = (
  documentNumber: unknown,
  documentName: unknown,
): {
  status: AnketaFieldStatus;
  documentNumber: string;
  documentName: string;
} => {
  const number = String(documentNumber ?? "").trim();
  const name = String(documentName ?? "").trim();

  if (!number && !name) {
    return { status: "empty", documentNumber: "", documentName: "" };
  }

  if (
    (number && isAbsentText(number)) ||
    (name && isAbsentText(name))
  ) {
    return {
      status: "absent",
      documentNumber: number,
      documentName: name,
    };
  }

  if (
    (number && isPlaceholderText(number) && !looksLikeDocumentName(name)) ||
    (name && isPlaceholderText(name) && !looksLikeDocumentNumber(number))
  ) {
    return { status: "empty", documentNumber: number, documentName: name };
  }

  if (looksLikeDocumentNumber(number) || looksLikeDocumentName(name)) {
    return { status: "has_value", documentNumber: number, documentName: name };
  }

  return { status: "empty", documentNumber: number, documentName: name };
};

const buildAnketaIndex = (rows: AnketaRow[]) => {
  const byName = new Map<string, AnketaRow>();
  for (const row of rows) {
    const key = normalizeAnketaNameKey(row.fullName);
    if (!key || byName.has(key)) continue;
    byName.set(key, row);
  }
  return byName;
};

const summarizeMilitary = (rows: MilitaryIdPersonRow[]): StaffFieldSummary => ({
  total: rows.length,
  has: rows.filter((row) => row.status === "has_value").length,
  absent: rows.filter((row) => row.status === "absent").length,
  empty: rows.filter((row) => row.status === "empty").length,
  noAnketa: rows.filter((row) => row.status === "no_anketa").length,
});

const summarizeDocuments = (rows: DocumentPersonRow[]): StaffFieldSummary => ({
  total: rows.length,
  has: rows.filter((row) => row.status === "has_value").length,
  absent: rows.filter((row) => row.status === "absent").length,
  empty: rows.filter((row) => row.status === "empty").length,
  noAnketa: rows.filter((row) => row.status === "no_anketa").length,
});

export const buildAnketaStaffStats = ({
  morning,
  anketaRows,
  year = new Date().getFullYear(),
}: {
  morning: ExcelWorkbookSnapshot;
  anketaRows: AnketaRow[];
  year?: number;
}): AnketaStaffStatsResult => {
  const morningSheet =
    morning.sheets.find(isBchsPersonnelGeneralListSheet) ??
    findMorningRosterSheet(morning);

  const anketaByName = buildAnketaIndex(anketaRows);
  const militaryStaffPeople: MilitaryIdPersonRow[] = [];
  const documentStaffPeople: DocumentPersonRow[] = [];

  if (morningSheet) {
    for (const person of extractBchsAwayPeopleFromSheet(morningSheet).filter(
      hasBchsFullName,
    )) {
      const fullName = String(person.fullName ?? "").trim();
      const rank = String(person.rankTitle ?? person.shpkFact ?? "").trim();
      const anketa = anketaByName.get(normalizeAnketaNameKey(fullName));
      if (!anketa) {
        militaryStaffPeople.push({
          fullName,
          rank,
          status: "no_anketa",
          militaryId: "",
        });
        documentStaffPeople.push({
          fullName,
          rank,
          status: "no_anketa",
          documentNumber: "",
          documentName: "",
        });
        continue;
      }

      const military = classifyAnketaMilitaryId(anketa.militaryId);
      militaryStaffPeople.push({
        fullName,
        rank,
        status: military.status,
        militaryId: military.militaryId,
      });

      const document = classifyAnketaDocument(
        anketa.idDocumentNumber,
        anketa.idDocumentName,
      );
      documentStaffPeople.push({
        fullName,
        rank,
        status: document.status,
        documentNumber: document.documentNumber,
        documentName: document.documentName,
      });
    }

    militaryStaffPeople.sort((a, b) => a.fullName.localeCompare(b.fullName, "uk"));
    documentStaffPeople.sort((a, b) => a.fullName.localeCompare(b.fullName, "uk"));
  }

  const militaryAnketaRows = anketaRows
    .filter((row) => looksLikePersonName(String(row.fullName ?? "")))
    .map((row) => {
      const classified = classifyAnketaMilitaryId(row.militaryId);
      return {
        fullName: row.fullName,
        rank: "",
        status: classified.status,
        militaryId: classified.militaryId,
      };
    });

  const documentAnketaRows = anketaRows
    .filter((row) => looksLikePersonName(String(row.fullName ?? "")))
    .map((row) => {
      const classified = classifyAnketaDocument(
        row.idDocumentNumber,
        row.idDocumentName,
      );
      return {
        fullName: row.fullName,
        rank: "",
        status: classified.status,
        documentNumber: classified.documentNumber,
        documentName: classified.documentName,
      };
    });

  return {
    year,
    militaryId: {
      staff: summarizeMilitary(militaryStaffPeople),
      anketa: summarizeMilitary(militaryAnketaRows),
      staffPeople: militaryStaffPeople,
    },
    documents: {
      staff: summarizeDocuments(documentStaffPeople),
      anketa: summarizeDocuments(documentAnketaRows),
      staffPeople: documentStaffPeople,
    },
  };
};

/** @deprecated use buildAnketaStaffStats */
export const buildMilitaryIdStats = buildAnketaStaffStats;

export const ANKETA_FIELD_STATUS_LABELS: Record<AnketaFieldStatus, string> = {
  has_value: "Є дані",
  absent: "Відсутній (позначено)",
  empty: "Порожньо / не заповнено",
  no_anketa: "Немає в анкетних даних",
};

/** @deprecated */
export const MILITARY_ID_STATUS_LABELS = ANKETA_FIELD_STATUS_LABELS;

export type MilitaryIdStatus = AnketaFieldStatus;

/** @deprecated */
export type MilitaryIdStatsSummary = StaffFieldSummary;
