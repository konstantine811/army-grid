import type { AnketaRow } from "./anketaSheet";

/** Номер ID-картки України — 9 цифр; Google/Excel часто зрізає нулі спереду. */
export const ANKETA_ID_CARD_NUMBER_LENGTH = 9;

const DOCUMENT_LETTER_RE = /[A-Za-zА-Яа-яІіЇїЄєҐґ]/u;

/** Відновити провідні нулі лише в чисто цифровому номері ID-картки. Серії з літерами не чіпаємо. */
export const padAnketaIdCardDocumentNumber = (value: unknown) => {
  const trimmed = String(value ?? "")
    .replace(/\u00a0/g, " ")
    .trim();
  if (!trimmed) return trimmed;
  if (DOCUMENT_LETTER_RE.test(trimmed)) return trimmed;
  const digits = trimmed.replace(/[\s'-]/g, "");
  if (!/^\d+$/.test(digits)) return trimmed;
  if (digits.length >= ANKETA_ID_CARD_NUMBER_LENGTH) return trimmed;
  return digits.padStart(ANKETA_ID_CARD_NUMBER_LENGTH, "0");
};

export const padAnketaIdDocumentNumbersInRows = (rows: AnketaRow[]) => {
  let changed = false;
  const next = rows.map((row) => {
    const padded = padAnketaIdCardDocumentNumber(row.idDocumentNumber);
    if (padded === row.idDocumentNumber) return row;
    changed = true;
    return { ...row, idDocumentNumber: padded };
  });
  return changed ? next : rows;
};
