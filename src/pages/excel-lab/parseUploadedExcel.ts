import {
  createWorkbookDebugPayload,
  readWorkbookSnapshot,
  type ExcelWorkbookSnapshot,
  type ReadWorkbookOptions,
} from "../../excelRoundTrip";
import { parseExcelEJOOSState } from "./ExcelLabEJOOS";
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

/** Опції читання за замовчуванням для лабораторії Excel. */
export const DEFAULT_EXCEL_LAB_READ_OPTIONS: ReadWorkbookOptions = {
  preserveLeadingColumns: true,
};

export const countMilitaryServiceCardPeople = (cards: EntityCardsSheets) =>
  Object.keys(cards).length;

/** ВК / ТПВ / ДОВІДКИ → parseExcelLabMilitaryServiceCards. */
export const handleParsedExcelWorkbook = (
  snapshot: ExcelWorkbookSnapshot,
  _debug: ReturnType<typeof createWorkbookDebugPayload>,
) => {
  const cards = parseExcelLabMilitaryServiceCards(snapshot.sheets);
  return cards;
};

export const parseMilitaryServiceCardsExcelFile = async (
  file: File,
  options: ReadWorkbookOptions = DEFAULT_EXCEL_LAB_READ_OPTIONS,
): Promise<ParsedMilitaryServiceCardsResult> => {
  const snapshot = await readWorkbookSnapshot(file, options);
  const debug = createWorkbookDebugPayload(snapshot);
  const cards = handleParsedExcelWorkbook(snapshot, debug);
  return { snapshot, debug, cards };
};

/** Тестовий перегляд довільного Excel — лише snapshot і debug у консолі. */
export const parseGenericTestExcelFile = async (
  file: File,
  options: ReadWorkbookOptions = DEFAULT_EXCEL_LAB_READ_OPTIONS,
): Promise<ParsedExcelLabResult> => {
  const snapshot = await readWorkbookSnapshot(file, options);
  const debug = createWorkbookDebugPayload(snapshot);
  const ejoos = parseExcelEJOOSState(snapshot.sheets);
  return { snapshot, debug };
};

/** @deprecated Використовуй parseGenericTestExcelFile або parseMilitaryServiceCardsExcelFile. */
export const parseUploadedExcelFile = parseGenericTestExcelFile;

/** Штатка → parseExcelLabState (логіка в ExcelLabStateParser.ts). */
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
