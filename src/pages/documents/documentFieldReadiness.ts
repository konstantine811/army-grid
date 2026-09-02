import { pickUbdBasisOrderForTaskPeriod, ubdBasisDateMatchesTaskPeriod } from "./ubdBasisOrders";

/** Типи з вибором бойового розпорядження (дата БР ↔ початок періоду завдань). */
export const DOCUMENT_TYPES_WITH_BASIS_ORDER = new Set(["ubdReport", "form6Report"]);

const DOCUMENT_PLACEHOLDER_VALUES = new Set([
  "-",
  "—",
  "не вказав",
  "не вказано",
  "немає",
  "не має",
  "н/д",
  "відсутній",
  "відсутнє",
  "невідомо",
]);

/** Обов’язкові поля для червоного підсвічування в журналі та формі. */
export const DOCUMENT_REQUIRED_INPUT_KEYS: Record<string, readonly string[]> = {
  ubdReport: [
    "fullName",
    "rank",
    "staffPosition",
    "birthDate",
    "rnokpp",
    "taskPeriod",
    "taskPlace",
    "basisNumber",
    "basisDate",
  ],
  form6Report: [
    "fullName",
    "rank",
    "staffPosition",
    "birthDate",
    "idDocument",
    "rnokpp",
    "taskPeriod",
    "taskPlace",
    "basisNumber",
    "basisDate",
  ],
  form12Report: ["commander", "fullName", "rank", "staffPosition"],
  ubdRestoreReport: [
    "commander",
    "fullName",
    "rank",
    "staffPosition",
    "certificateSeries",
    "circumstances",
    "requestText",
  ],
  serviceCharacteristic: ["rank", "lastName", "firstName", "staffPosition"],
  zhbdCertificate: [
    "rank",
    "fullName",
    "staffPosition",
    "periodFrom",
    "periodTo",
    "bodyParagraph",
  ],
  salaryPowerAttorney: ["fullName", "rnokpp", "iban", "bankName"],
  temporaryMilitaryId: ["fullName", "rank", "birthDate"],
  lostMilitaryId: ["fullName", "rank", "staffPosition", "addressee", "lossDate"],
};

const RNOKPP_DOCUMENT_TYPES = new Set(["ubdReport", "form6Report", "salaryPowerAttorney"]);

const FALLBACK_BASIS = pickUbdBasisOrderForTaskPeriod("") ?? {
  number: "4862/ОКП/1162/дск",
  date: "09.05.2026",
};

export const isBlankDocumentInput = (value: unknown) => {
  const text = String(value ?? "").trim();
  if (!text) return true;
  return DOCUMENT_PLACEHOLDER_VALUES.has(text.toLocaleLowerCase("uk-UA"));
};

/** РНОКПП: рівно 10 цифр; «не вказав» / коротший номер = порожньо. */
export const isBlankUbdRnokpp = (value: unknown) => {
  if (isBlankDocumentInput(value)) return true;
  const digits = String(value ?? "").replace(/\D/g, "");
  return digits.length !== 10;
};

/** Форма 6: «Паспорт громадянина України» без серії/номера = порожньо. */
export const isBlankForm6IdDocument = (value: unknown) => {
  if (isBlankDocumentInput(value)) return true;
  const digits = String(value ?? "").replace(/\D/g, "");
  return digits.length < 6;
};

const parseCombinedBasisParts = (value: string, withFallback: boolean) => {
  const text = String(value ?? "").trim();
  const match = text.match(
    /№\s*(\S+)\s+від\s+(\d{1,2}[.\/-]\d{1,2}[.\/-]\d{2,4})/i,
  );
  if (!match) {
    if (!withFallback) {
      return { basisNumber: "", basisDate: "" };
    }
    return {
      basisNumber: FALLBACK_BASIS.number,
      basisDate: FALLBACK_BASIS.date,
    };
  }
  return {
    basisNumber: match[1],
    basisDate: match[2].replaceAll("/", ".").replaceAll("-", "."),
  };
};

export const resolveDocumentFieldsForGapCheck = (
  type: string,
  fields: Record<string, unknown> | null | undefined,
) => {
  const record: Record<string, unknown> = { ...(fields || {}) };
  if (!DOCUMENT_TYPES_WITH_BASIS_ORDER.has(type)) return record;
  if (
    isBlankDocumentInput(record.basisNumber) ||
    isBlankDocumentInput(record.basisDate)
  ) {
    const parsed = parseCombinedBasisParts(
      String(record.basis ?? ""),
      type === "ubdReport",
    );
    if (isBlankDocumentInput(record.basisNumber) && parsed.basisNumber) {
      record.basisNumber = parsed.basisNumber;
    }
    if (isBlankDocumentInput(record.basisDate) && parsed.basisDate) {
      record.basisDate = parsed.basisDate;
    }
  }
  return record;
};

export const resolveUbdFieldsForGapCheck = (
  fields: Record<string, unknown> | null | undefined,
) => resolveDocumentFieldsForGapCheck("ubdReport", fields);

export const documentRequiredInputKeys = (
  type: string,
  fields: Record<string, unknown> | null | undefined,
): readonly string[] => {
  const base = DOCUMENT_REQUIRED_INPUT_KEYS[type] ?? [];
  if (type === "lostMilitaryId") {
    const kind = String(fields?.circumstanceKind ?? "movement");
    if (kind === "custom") return [...base, "customCircumstances"];
    return [...base, "fromLocation", "toLocation"];
  }
  if (type === "form6Report") {
    const manual =
      fields?.basisManual === true || String(fields?.basisManual) === "true";
    if (manual && !isBlankDocumentInput(fields?.basis)) {
      return base.filter((key) => key !== "basisNumber" && key !== "basisDate");
    }
  }
  return base;
};

const isRequiredValueBlank = (
  type: string,
  key: string,
  value: unknown,
) => {
  if (key === "rnokpp" && RNOKPP_DOCUMENT_TYPES.has(type)) {
    return isBlankUbdRnokpp(value);
  }
  if (type === "form6Report" && key === "idDocument") {
    return isBlankForm6IdDocument(value);
  }
  return isBlankDocumentInput(value);
};

export const documentHasEmptyInputs = (
  type: string,
  fields: Record<string, unknown> | null | undefined,
) => {
  const resolved = resolveDocumentFieldsForGapCheck(type, fields);
  return documentRequiredInputKeys(type, resolved).some((key) =>
    isRequiredValueBlank(type, key, resolved[key]),
  );
};

export const documentRequiredFieldIsBlank = (
  type: string,
  key: string,
  fields: Record<string, unknown> | null | undefined,
) => {
  const required = documentRequiredInputKeys(type, fields);
  if (!required.includes(key)) return false;
  const resolved = resolveDocumentFieldsForGapCheck(type, fields);
  return isRequiredValueBlank(type, key, resolved[key]);
};

/** Дата БР не збігається з «з» періоду завдань — жовтий рядок / поле. */
export const documentHasBasisDateMismatch = (
  type: string,
  fields: Record<string, unknown> | null | undefined,
) => {
  if (!DOCUMENT_TYPES_WITH_BASIS_ORDER.has(type)) return false;
  const resolved = resolveDocumentFieldsForGapCheck(type, fields);
  if (type === "ubdReport") {
    const flag = resolved.basisNotReady;
    if (flag === true || flag === "true") return true;
  }
  const taskPeriod = String(resolved.taskPeriod ?? "").trim();
  const basisDate = String(resolved.basisDate ?? "").trim();
  if (!taskPeriod || isBlankDocumentInput(basisDate)) return false;
  return !ubdBasisDateMatchesTaskPeriod(
    taskPeriod,
    basisDate,
    String(resolved.taskPlace ?? ""),
  );
};

export const documentBasisFieldHighlightClass = (
  type: string,
  fields: Record<string, unknown> | null | undefined,
  key: "basisNumber" | "basisDate",
) => {
  if (!DOCUMENT_TYPES_WITH_BASIS_ORDER.has(type)) return undefined;
  if (documentRequiredFieldIsBlank(type, key, fields)) {
    return "document-field-invalid";
  }
  if (documentHasBasisDateMismatch(type, fields)) {
    return "document-field-warning";
  }
  return undefined;
};

export const documentFieldLabelClass = (
  ...parts: Array<string | false | null | undefined>
) => {
  const text = parts.filter((part): part is string => Boolean(part)).join(" ");
  return text || undefined;
};

export const readDocumentSkippedDueToSzch = (
  fields: Record<string, unknown> | null | undefined,
) => {
  const value = fields?.skippedDueToSzch;
  return value === true || value === "true";
};
