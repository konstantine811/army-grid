import {
  createWorkbookDebugPayload,
  readWorkbookSnapshot,
  type ExcelWorkbookSnapshot,
  type ReadWorkbookOptions,
} from "../../excelRoundTrip";
import { parseExcelKSPState, type EntityKSPVKSheets } from "./ExcelKSPVKData";
import { parseExcelEJOOSState, type EntityEJOOSSheets } from "./ExcelLabEJOOS";
import {
  parseExcelLabMilitaryServiceCards,
  type EntityCardsSheets,
} from "./ExcelLabMilitaryServiceCards";
import {
  parseExcelLabState,
  type EntityStateSheets,
} from "./ExcelLabStateParser";

export type ParsedExcelLabResult = {
  snapshot: ExcelWorkbookSnapshot;
  debug: ReturnType<typeof createWorkbookDebugPayload>;
};

export type ParsedStaffSheetResult = ParsedExcelLabResult & {
  state: EntityStateSheets;
};

export type ParsedMilitaryServiceCardsResult = ParsedExcelLabResult & {
  cards: EntityCardsSheets;
};

export type ParsedEJOOSResult = ParsedExcelLabResult & {
  ejoos: EntityEJOOSSheets;
};

export type ParsedKSPVKResult = ParsedExcelLabResult & {
  ksp: EntityKSPVKSheets;
};

export type ExcelLabUploadedSources = {
  staff?: {
    fileName: string;
    state: EntityStateSheets;
  };
  militaryCards?: {
    fileName: string;
    cards: EntityCardsSheets;
  };
  ejoos?: {
    fileName: string;
    state: EntityEJOOSSheets;
  };
  ksp?: {
    fileName: string;
    state: EntityKSPVKSheets;
  };
};

export type ExcelLabProcessedData = {
  staff: EntityStateSheets;
  militaryCards: EntityCardsSheets;
  ejoos: EntityEJOOSSheets;
  ksp: EntityKSPVKSheets;
  sources: ExcelLabUploadedSources;
};

export const EXCEL_LAB_SOURCE_COUNT = 4;

/** Опції читання за замовчуванням для лабораторії Excel. */
export const DEFAULT_EXCEL_LAB_READ_OPTIONS: ReadWorkbookOptions = {
  preserveLeadingColumns: true,
};

export const countMilitaryServiceCardPeople = (cards: EntityCardsSheets) =>
  Object.keys(cards).length;

export const countEJOOSPeople = (ejoos: EntityEJOOSSheets) =>
  Object.keys(ejoos).length;

export const countKSPVKPeople = (ksp: EntityKSPVKSheets) =>
  Object.keys(ksp).length;

/** Збирає результати чотирьох окремих завантажень у один об'єкт. */
export const buildExcelLabProcessedData = (
  sources: ExcelLabUploadedSources,
): ExcelLabProcessedData => ({
  staff: sources.staff?.state ?? {},
  militaryCards: sources.militaryCards?.cards ?? {},
  ejoos: sources.ejoos?.state ?? {},
  ksp: sources.ksp?.state ?? {},
  sources,
});

export const describeExcelLabProcessedData = (data: ExcelLabProcessedData) => {
  const loaded = [
    data.sources.staff ? `Штатка: ${data.sources.staff.fileName}` : null,
    data.sources.militaryCards
      ? `Квитки: ${data.sources.militaryCards.fileName}`
      : null,
    data.sources.ejoos ? `ЄЖООС: ${data.sources.ejoos.fileName}` : null,
    data.sources.ksp ? `КСП: ${data.sources.ksp.fileName}` : null,
  ].filter(Boolean);

  return [
    `Завантажено: ${loaded.length}/${EXCEL_LAB_SOURCE_COUNT}`,
    ...loaded,
    `Особи в штатці: ${Object.keys(data.staff).length}`,
    `Особи з картками: ${Object.keys(data.militaryCards).length}`,
    `Особи в ЄЖООС: ${Object.keys(data.ejoos).length}`,
    `Особи в КСП: ${Object.keys(data.ksp).length}`,
  ].join("\n");
};

/** ВК / ТПВ / ДОВІДКИ → parseExcelLabMilitaryServiceCards. */
export const handleParsedExcelWorkbook = (
  snapshot: ExcelWorkbookSnapshot,
  _debug: ReturnType<typeof createWorkbookDebugPayload>,
) => parseExcelLabMilitaryServiceCards(snapshot.sheets);

export const parseMilitaryServiceCardsExcelFile = async (
  file: File,
  options: ReadWorkbookOptions = DEFAULT_EXCEL_LAB_READ_OPTIONS,
): Promise<ParsedMilitaryServiceCardsResult> => {
  const snapshot = await readWorkbookSnapshot(file, options);
  const debug = createWorkbookDebugPayload(snapshot);
  const cards = handleParsedExcelWorkbook(snapshot, debug);
  console.log("[excel-lab] military service cards:", cards);
  return { snapshot, debug, cards };
};

/** ЄЖООС → parseExcelEJOOSState. */
export const parseEJOOSExcelFile = async (
  file: File,
  options: ReadWorkbookOptions = DEFAULT_EXCEL_LAB_READ_OPTIONS,
): Promise<ParsedEJOOSResult> => {
  const snapshot = await readWorkbookSnapshot(file, options);
  const debug = createWorkbookDebugPayload(snapshot);
  const ejoos = parseExcelEJOOSState(snapshot.sheets);
  console.log("[excel-lab] ejoos state:", ejoos);
  return { snapshot, debug, ejoos };
};

/** Штатка → parseExcelLabState. */
export const parseStaffSheetExcelFile = async (
  file: File,
  options: ReadWorkbookOptions = DEFAULT_EXCEL_LAB_READ_OPTIONS,
): Promise<ParsedStaffSheetResult> => {
  const snapshot = await readWorkbookSnapshot(file, options);
  const debug = createWorkbookDebugPayload(snapshot);
  const state = parseExcelLabState(snapshot.sheets);
  console.log("[excel-lab] staff state:", state);
  return { snapshot, debug, state };
};

/** Військові із КСП → parseExcelKSPState (ExcelKSPVKData.ts). */
export const parseKSPVKExcelFile = async (
  file: File,
  options: ReadWorkbookOptions = DEFAULT_EXCEL_LAB_READ_OPTIONS,
): Promise<ParsedKSPVKResult> => {
  const snapshot = await readWorkbookSnapshot(file, options);
  const debug = createWorkbookDebugPayload(snapshot);
  const ksp = parseExcelKSPState(snapshot.sheets);
  return { snapshot, debug, ksp };
};
