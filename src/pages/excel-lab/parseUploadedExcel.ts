import {
  createWorkbookDebugPayload,
  readWorkbookSnapshot,
  type ExcelWorkbookSnapshot,
  type ReadWorkbookOptions,
} from "../../excelRoundTrip";
import { parseExcelLabMilitaryServiceCards } from "./ExcelLabMilitaryServiceCards";
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

/** Опції читання за замовчуванням для лабораторії Excel. */
export const DEFAULT_EXCEL_LAB_READ_OPTIONS: ReadWorkbookOptions = {
  preserveLeadingColumns: true,
};

/**
 * Довільний Excel → console.log.
 * Свою логіку додавайте сюди або в окремий handler.
 */
export const handleParsedExcelWorkbook = (
  snapshot: ExcelWorkbookSnapshot,
  debug: ReturnType<typeof createWorkbookDebugPayload>,
) => {
  const cards = parseExcelLabMilitaryServiceCards(snapshot.sheets);
};

export const parseUploadedExcelFile = async (
  file: File,
  options: ReadWorkbookOptions = DEFAULT_EXCEL_LAB_READ_OPTIONS,
): Promise<ParsedExcelLabResult> => {
  const snapshot = await readWorkbookSnapshot(file, options);
  const debug = createWorkbookDebugPayload(snapshot);
  handleParsedExcelWorkbook(snapshot, debug);
  return { snapshot, debug };
};

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
