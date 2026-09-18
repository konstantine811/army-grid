/** Built-in OOXML numFmtId for `@` (Text) — Excel «Special» / текстовий формат. */
export const EJOOS_TEXT_NUM_FMT_ID = "49";

/** Built-in OOXML numFmtId for General — Excel «Загальний». */
export const EJOOS_GENERAL_NUM_FMT_ID = "0";

/** РНОКПП на ООС — колонка V. */
export const EJOOS_OOS_RNOKPP_COLUMN = 22;
/** Відмова від РНОКПП на ООС — колонка W. */
export const EJOOS_OOS_RNOKPP_REFUSE_COLUMN = 23;
/** Військовий квиток на ООС — колонка Z (після паспортних полів). */
export const EJOOS_OOS_MILITARY_ID_COLUMN = 26;
/** РНОКПП на Виключені — колонка Q. */
export const EJOOS_EXCLUDED_RNOKPP_COLUMN = 17;

/** Перший рядок даних на аркушах ШПО / Табель / Виключені. */
export const EJOOS_STAFF_INDEX_MIN_ROW = 7;

const sheetLabel = (sheetNameOrRe: string | RegExp) =>
  typeof sheetNameOrRe === "string"
    ? sheetNameOrRe
    : sheetNameOrRe.source;

/** Колонки «індекс посади» на аркушах ЄЖООС (1-based). */
export const isEjoosStaffIndexColumn = (
  sheetNameOrRe: string | RegExp,
  column: number,
  row = 0,
): boolean => {
  if (row > 0 && row < EJOOS_STAFF_INDEX_MIN_ROW) return false;
  const lower = sheetLabel(sheetNameOrRe).toLowerCase();
  if (/табель/.test(lower)) return column === 2;
  if (/шпо|штатно/.test(lower)) return column === 1;
  if (/виключен/.test(lower)) return column === 4;
  return false;
};

/** Колонка ID (C) на ООС / Виключені — завжди текст, інакше Excel показує dd.mm.yyyy. */
export const isEjoosPersonIdColumn = (
  sheetNameOrRe: string | RegExp,
  column: number,
  row = 0,
): boolean => {
  if (column !== 3) return false;
  if (row > 0 && row < EJOOS_STAFF_INDEX_MIN_ROW) return false;
  const lower = sheetLabel(sheetNameOrRe).toLowerCase();
  return /оос|виключен/.test(lower);
};

/** Військовий квиток на ООС — колонка Z; завжди текст (не дата / не число). */
export const isEjoosMilitaryIdColumn = (
  sheetNameOrRe: string | RegExp,
  column: number,
  row = 0,
): boolean => {
  if (column !== EJOOS_OOS_MILITARY_ID_COLUMN) return false;
  if (row > 0 && row < EJOOS_STAFF_INDEX_MIN_ROW) return false;
  return /оос/i.test(sheetLabel(sheetNameOrRe));
};

export const isEjoosForceTextColumn = (
  sheetNameOrRe: string | RegExp,
  column: number,
  row = 0,
): boolean =>
  isEjoosStaffIndexColumn(sheetNameOrRe, column, row) ||
  isEjoosPersonIdColumn(sheetNameOrRe, column, row) ||
  isEjoosMilitaryIdColumn(sheetNameOrRe, column, row);

/** РНОКПП: General, не Custom `@` — інакше код пишеться символами. */
export const isEjoosRnokppColumn = (
  sheetNameOrRe: string | RegExp,
  column: number,
  row = 0,
): boolean => {
  if (row > 0 && row < EJOOS_STAFF_INDEX_MIN_ROW) return false;
  const lower = sheetLabel(sheetNameOrRe).toLowerCase();
  if (/виключен/.test(lower)) return column === EJOOS_EXCLUDED_RNOKPP_COLUMN;
  if (/оос/.test(lower)) return column === EJOOS_OOS_RNOKPP_COLUMN;
  return false;
};

export const normalizeStaffIndexValue = (
  value: string | number | null | undefined,
): string => String(value ?? "").trim();

/** Запис індексу посади через xlsx-populate: завжди текст, не дата. */
export const writeStaffIndexToCell = (
  cell: unknown,
  value: string | number | null | undefined,
) => {
  const text = normalizeStaffIndexValue(value);
  if (!text) return;
  const target = cell as {
    style?: (name: string, value?: unknown) => unknown;
    value: (next: string) => unknown;
  };
  target.style?.("numberFormat", "@");
  target.value(text);
};

export const writePersonIdToCell = (
  cell: unknown,
  value: string | number | null | undefined,
) => {
  const text = normalizeStaffIndexValue(value);
  if (!text || !/^\d{1,7}$/.test(text)) return;
  writeStaffIndexToCell(cell, text);
};
