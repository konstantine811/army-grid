/** Built-in OOXML numFmtId for `@` (Text) — Excel «Special» / текстовий формат. */
export const EJOOS_TEXT_NUM_FMT_ID = "49";

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
