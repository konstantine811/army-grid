import type {
  BackendEjournalImport,
  BackendEjournalImportSheet,
  EjournalRowActionType,
} from "../../api";
import { api } from "../../api";
import type { DbPreviewState, EjournalPreviewRow } from "../ejournal/ejournalTypes";
import {
  parseDbColumns,
  previewValueToDisplay,
} from "../ejournal/ejournalUtils";

export {
  getRowKeyByKeyPart,
  getRowValueByKeyPart,
  previewValueToDisplay,
} from "../ejournal/ejournalUtils";

export type PersonAction = {
  type: EjournalRowActionType;
  label: string;
};

export type PersonActionForm = {
  validFrom: string;
  validTo: string;
  reason: string;
  place: string;
  note: string;
  positionIndex: string;
  positionTitle: string;
  rank: string;
};

export const personActions: PersonAction[] = [
  { type: "MEDICAL", label: "Лікування" },
  { type: "LEAVE", label: "Відпустка" },
  { type: "BUSINESS_TRIP", label: "Відрядження" },
  { type: "AWOL", label: "СЗЧ" },
  { type: "CAPTIVITY", label: "Полон" },
  { type: "EXCLUSION", label: "Виключення" },
  { type: "RETURNED", label: "Повернувся" },
  { type: "POSITION_CHANGE", label: "Змінити посаду" },
  { type: "RANK_CHANGE", label: "Змінити звання" },
];

export const createDefaultActionForm = (): PersonActionForm => ({
  validFrom: new Date().toISOString().slice(0, 10),
  validTo: "",
  reason: "",
  place: "",
  note: "",
  positionIndex: "",
  positionTitle: "",
  rank: "",
});

export const getPersonExternalId = (row: EjournalPreviewRow | null) => {
  if (!row) return "";

  const exact =
    Object.keys(row).find((key) => {
      const normalized = key.toLowerCase();
      return (
        !key.startsWith("__") &&
        (normalized === "id" ||
          normalized === "ід" ||
          normalized === "externalid" ||
          normalized === "external_id")
      );
    }) ??
    Object.keys(row).find((key) => {
      const normalized = key.toLowerCase();
      return (
        !key.startsWith("__") &&
        (normalized.includes("зовнішн") || /(^|_)id$/i.test(normalized))
      );
    });

  if (!exact) return "";

  const raw = previewValueToDisplay(row[exact]).trim();
  if (!raw || raw === "[object Object]" || raw === "null" || raw === "undefined") {
    return "";
  }

  if (raw.startsWith("{") || raw.startsWith("[")) {
    try {
      const parsed = JSON.parse(raw) as Record<string, unknown> | unknown[];
      if (parsed && typeof parsed === "object" && !Array.isArray(parsed)) {
        const nested =
          parsed.text ?? parsed.value ?? parsed.id ?? parsed.result;
        if (nested != null && nested !== "") return String(nested).trim();
      }
    } catch {
      // keep empty for unreadable structured IDs
    }
    return "";
  }

  return raw;
};

/** Position indexes come from Excel as newline-separated values; show them on one line. */
export const formatPositionIndexes = (value: unknown) => {
  const text = previewValueToDisplay(value).replace(/\r/g, "").trim();
  if (!text) return "";

  const parts = text
    .split(/[\n;,+]+/)
    .map((part) => part.trim())
    .filter(Boolean);

  if (parts.length > 1) return parts.join(" · ");

  const single = parts[0] ?? text;
  // Recover already-collapsed sequences of typical 7-digit indexes.
  if (/^\d{14,}$/.test(single) && single.length % 7 === 0) {
    return single.match(/.{7}/g)?.join(" · ") ?? single;
  }

  return single;
};

export const isPositionIndexField = (parts: string[]) =>
  parts.includes("індекс") && parts.includes("посади");

export const formatMultilineText = (value: unknown) =>
  previewValueToDisplay(value).replace(/\r\n/g, "\n").replace(/\r/g, "\n").trim();

/** Excel serial dates (e.g. 33448) → uk-UA display; leave normal text dates as-is. */
export const formatExcelDateDisplay = (value: unknown) => {
  if (value == null || value === "") return "";
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return new Intl.DateTimeFormat("uk-UA").format(value);
  }

  const text = previewValueToDisplay(value).trim();
  if (!text) return "";

  const asNumber = Number(String(text).replace(",", "."));
  if (
    Number.isFinite(asNumber) &&
    asNumber > 20000 &&
    asNumber < 80000 &&
    !/[./-]/.test(text)
  ) {
    const date = new Date(Math.round((asNumber - 25569) * 86400 * 1000));
    if (!Number.isNaN(date.getTime())) {
      return new Intl.DateTimeFormat("uk-UA").format(date);
    }
  }

  return text;
};

/**
 * UA phone chunks in free text:
 * 0635963362 | 097-154-28-64 | 097 154 28 64 | +380 63 596 33 62 | (063)596-33-62 | 635963362
 * Do not use \s (newlines) inside a single match — phones are often one-per-line.
 */
const PHONE_CANDIDATE_RE =
  /(?:\+?38[ \-.]*)?(?:\(?0\)?[ \-.]*)?(?:\d[ \-./()]*){8,10}\d/g;

/** Normalize to canonical `0XXXXXXXXX`, or null if not a plausible UA number. */
export const normalizeUaPhone = (raw: string): string | null => {
  const trimmed = raw.trim();
  if (!trimmed) return null;

  let digits = trimmed.replace(/\D/g, "");
  if (!digits) return null;

  if (digits.startsWith("380")) digits = `0${digits.slice(3)}`;
  else if (digits.startsWith("38") && digits.length >= 11) {
    digits = `0${digits.slice(2)}`;
  }

  if (digits.length === 9 && /^[3-9]\d{8}$/.test(digits)) {
    digits = `0${digits}`;
  }

  // Extra trailing digit typos (e.g. 09509558658) → keep first 10 if valid prefix.
  if (digits.length > 10 && digits.startsWith("0")) {
    digits = digits.slice(0, 10);
  }

  if (!/^0\d{9}$/.test(digits)) return null;
  if (/^(\d)\1{9}$/.test(digits)) return null;

  return digits;
};

export const formatUaPhoneDisplay = (phone: string) => {
  const normalized = normalizeUaPhone(phone);
  if (!normalized) return phone.trim();
  return `${normalized.slice(0, 3)} ${normalized.slice(3, 6)} ${normalized.slice(6, 8)} ${normalized.slice(8)}`;
};

/** Parse all phones from «Додаткова інформація» (and similar free-text fields). */
export const extractPhones = (value: unknown): string[] => {
  const text = formatMultilineText(value);
  if (!text) return [];

  const phones: string[] = [];
  const push = (candidate: string) => {
    const normalized = normalizeUaPhone(candidate);
    if (normalized && !phones.includes(normalized)) phones.push(normalized);
  };

  for (const match of text.matchAll(PHONE_CANDIDATE_RE)) {
    push(match[0]);
  }

  // One phone per line / semicolon / comma — covers plain `063…\n097…`.
  for (const chunk of text.split(/[\n;,]+/)) {
    push(chunk);
  }

  return phones;
};

/** Prefer exact / shortest key match so "звання" does not hit order columns. */
export const resolvePersonFieldKey = (
  row: EjournalPreviewRow | null,
  keyParts: string[],
) => {
  if (!row || keyParts.length === 0) return "";

  const keys = Object.keys(row).filter((key) => !key.startsWith("__"));
  const parts = keyParts.map((part) => part.toLowerCase());

  if (parts.length === 1) {
    const exact = keys.find((key) => key.toLowerCase() === parts[0]);
    if (exact) return exact;
  }

  const joined = parts.join("_");
  const joinedExact = keys.find((key) => key.toLowerCase() === joined);
  if (joinedExact) return joinedExact;

  const matches = keys.filter((key) => {
    const normalized = key.toLowerCase();
    return parts.every((part) => normalized.includes(part));
  });
  if (matches.length === 0) return "";

  // Prefer keys that start with the first part (рнокпп… over відмова_від_рнокпп),
  // then the shortest remaining match.
  return (
    matches.sort((left, right) => {
      const leftStarts = left.toLowerCase().startsWith(parts[0] ?? "") ? 0 : 1;
      const rightStarts = right.toLowerCase().startsWith(parts[0] ?? "") ? 0 : 1;
      if (leftStarts !== rightStarts) return leftStarts - rightStarts;
      return left.length - right.length;
    })[0] ?? ""
  );
};

export const getPersonFieldValue = (
  row: EjournalPreviewRow | null,
  keyParts: string[],
) => {
  const key = resolvePersonFieldKey(row, keyParts);
  return key ? previewValueToDisplay(row?.[key]) : "";
};

export type PersonFieldDef = {
  label: string;
  parts: string[];
  kind?: "text" | "multiline" | "date" | "positionIndex";
  section: "identity" | "service" | "orders" | "contacts";
};

/** Full OСС (2. ООС) field map for personnel card. */
export const PERSON_CARD_FIELDS: PersonFieldDef[] = [
  { label: "ПІБ", parts: ["прізвище"], section: "identity" },
  { label: "Звання", parts: ["звання"], section: "identity" },
  { label: "ID", parts: ["id"], section: "identity", kind: "text" },
  {
    label: "Індекс посади",
    parts: ["індекс", "посади"],
    section: "identity",
    kind: "positionIndex",
  },
  {
    label: "Дата народження",
    parts: ["дата_народження"],
    section: "identity",
    kind: "date",
  },
  { label: "Місце народження", parts: ["місце_народження"], section: "identity" },
  { label: "Стать", parts: ["стать"], section: "identity" },
  // Prefer рнокпп_за_наявності — bare ["рнокпп"] wrongly hits shorter відмова_від_рнокпп.
  { label: "РНОКПП", parts: ["рнокпп_за_наявності"], section: "identity" },
  { label: "Відмова від РНОКПП", parts: ["відмова", "рнокпп"], section: "identity" },
  {
    label: "Документ, що посвідчує особу",
    parts: ["назва", "документа", "посвідчує"],
    section: "identity",
  },
  {
    label: "Серія/номер документа",
    parts: ["серія", "номер", "документа", "посвідчує"],
    section: "identity",
  },
  {
    label: "Військовий квиток",
    parts: ["військового", "квитка"],
    section: "identity",
  },

  { label: "Вид служби", parts: ["вид_служби"], section: "service" },
  { label: "Дислокація", parts: ["місце_дислокації"], section: "service" },
  { label: "Звідки прибув", parts: ["звідки", "прибув"], section: "service" },
  {
    label: "Дата укладання контракту",
    parts: ["укладання", "контракту"],
    section: "service",
    kind: "date",
  },
  {
    label: "Дата закінчення контракту / період призову",
    parts: ["закінчення", "контракту"],
    section: "service",
    kind: "date",
  },
  {
    label: "Коли призваний (прийнятий)",
    parts: ["коли", "призваний"],
    section: "service",
    kind: "date",
  },
  {
    label: "Ким призваний (прийнятий)",
    parts: ["ким", "призваний"],
    section: "service",
  },
  {
    label: "Освіта",
    parts: ["освіта"],
    section: "service",
    kind: "multiline",
  },

  {
    label: "Дати прийняття посади",
    parts: ["дати_прийняття_посади"],
    section: "orders",
    kind: "multiline",
  },
  {
    label: "Номер наказу на прийняття посади",
    parts: ["номер_наказу_на_прийняття_посади"],
    section: "orders",
    kind: "multiline",
  },
  {
    label: "Дата зарахування до списків",
    parts: ["дата_зарахування_до_списків"],
    section: "orders",
    kind: "date",
  },
  {
    label: "Наказ про зарахування — дата",
    parts: ["наказ_про_зарахування_до_списків_-_дата"],
    section: "orders",
    kind: "date",
  },
  {
    label: "Наказ про зарахування — номер",
    parts: ["наказ_про_зарахування_до_списків_-_номер"],
    section: "orders",
  },
  {
    label: "Наказ на призначення — дата",
    parts: ["наказ_на_призначення_на_посаду_-_дата"],
    section: "orders",
    kind: "date",
  },
  {
    label: "Наказ на призначення — номер",
    parts: ["наказ_на_призначення_на_посаду_-_номер"],
    section: "orders",
  },
  {
    label: "Вхідна дата наказу на призначення",
    parts: ["вхідна_дата_наказу_на_призначення_на_посаду"],
    section: "orders",
    kind: "date",
  },
  {
    label: "Вхідний номер наказу на призначення",
    parts: ["вхідний_номер_наказу_на_призначення_на_посаду"],
    section: "orders",
  },
  {
    label: "Наказ на звання — дата",
    parts: ["наказ_на_присвоєння_останнього_звання_-_дата"],
    section: "orders",
    kind: "date",
  },
  {
    label: "Наказ на звання — номер",
    parts: ["наказ_на_присвоєння_останнього_звання_-_номер"],
    section: "orders",
  },
  {
    label: "Вхідна дата наказу на звання",
    parts: ["вхідна_дата_наказу_на_присвоєння_останнього_звання"],
    section: "orders",
    kind: "date",
  },
  {
    label: "Вхідний номер наказу на звання",
    parts: ["вхідний_номер_наказу_на_присвоєння_останнього_звання"],
    section: "orders",
  },

  {
    label: "Дані про родичів",
    parts: ["дані", "родичів"],
    section: "contacts",
    kind: "multiline",
  },
  {
    label: "Додаткова інформація",
    parts: ["додаткова_інформація"],
    section: "contacts",
    kind: "multiline",
  },
];

export const formatPersonFieldValue = (
  value: unknown,
  field: Pick<PersonFieldDef, "kind" | "parts">,
) => {
  if (field.kind === "positionIndex" || isPositionIndexField(field.parts)) {
    return formatPositionIndexes(value);
  }
  if (field.kind === "date") return formatExcelDateDisplay(value);
  if (field.kind === "multiline") return formatMultilineText(value);
  return previewValueToDisplay(value).trim();
};

export const PERSON_SECTION_LABELS: Record<PersonFieldDef["section"], string> =
  {
    identity: "Особові дані",
    service: "Служба",
    orders: "Накази та посади",
    contacts: "Контакти та додаткова інформація",
  };

export const buildPersonSummary = (row: EjournalPreviewRow | null) => {
  const additionalInfo = formatMultilineText(
    getPersonFieldValue(row, ["додаткова_інформація"]),
  );
  const relatives = formatMultilineText(
    getPersonFieldValue(row, ["дані", "родичів"]),
  );
  /** Separate property: phones parsed from «Додаткова інформація». */
  const phones = extractPhones(additionalInfo);

  return {
    name: getPersonFieldValue(row, ["прізвище"]) || "Особа не вибрана",
    rank: getPersonFieldValue(row, ["звання"]),
    externalId: getPersonExternalId(row),
    positionIndex: formatPositionIndexes(
      getPersonFieldValue(row, ["індекс", "посади"]),
    ),
    serviceType: getPersonFieldValue(row, ["вид_служби"]),
    birthDate: formatExcelDateDisplay(
      getPersonFieldValue(row, ["дата_народження"]),
    ),
    birthPlace: getPersonFieldValue(row, ["місце_народження"]),
    sex: getPersonFieldValue(row, ["стать"]),
    rnokpp: getPersonFieldValue(row, ["рнокпп_за_наявності"]),
    location: getPersonFieldValue(row, ["місце_дислокації"]),
    arrivedFrom: getPersonFieldValue(row, ["звідки", "прибув"]),
    education: formatMultilineText(getPersonFieldValue(row, ["освіта"])),
    relatives,
    additionalInfo,
    phones,
    phonesDisplay: phones.map(formatUaPhoneDisplay),
    militaryId: getPersonFieldValue(row, ["військового", "квитка"]),
    contractFrom: formatExcelDateDisplay(
      getPersonFieldValue(row, ["укладання", "контракту"]),
    ),
    contractTo: formatExcelDateDisplay(
      getPersonFieldValue(row, ["закінчення", "контракту"]),
    ),
    callSign: extractPersonCallSign(
      getPersonFieldValue(row, ["прізвище"]),
      additionalInfo,
    ),
  };
};

/** Позивний з дужок / «позивний …» / імені файлу анкети (не дата народження). */
export const isLikelyBirthDateToken = (value: string) => {
  const text = String(value ?? "").trim();
  if (!text) return false;
  if (/\d{1,2}[.\/-]\d{1,2}[.\/-]\d{2,4}/.test(text)) return true;
  if (/р\.?\s*н\.?/i.test(text)) return true;
  if (/^\d{1,2}\s+\S+\s+\d{4}/.test(text)) return true;
  return false;
};

export const isLikelyCallSignToken = (value: string) => {
  const text = String(value ?? "").trim();
  if (!text || text.length < 2 || text.length > 32) return false;
  if (isLikelyBirthDateToken(text)) return false;
  if (/^\d+$/.test(text)) return false;
  // Callsigns are usually words / short phrases, not full FIO.
  const words = text.split(/\s+/).filter(Boolean);
  if (words.length > 3) return false;
  return /[a-zа-яіїєґ]/i.test(text);
};

export const extractPersonCallSign = (...sources: Array<string | undefined>) => {
  for (const source of sources) {
    const text = String(source ?? "").trim();
    if (!text) continue;

    const labeled = text.match(/позивн\w*\s*[:\-–—]?\s*([^\s,;|/)]+)/i);
    if (labeled?.[1] && isLikelyCallSignToken(labeled[1])) {
      return labeled[1].trim();
    }

    const parenMatches = [...text.matchAll(/\(([^)]+)\)/g)];
    for (const match of parenMatches) {
      const candidate = match[1]?.trim() ?? "";
      if (isLikelyCallSignToken(candidate)) return candidate;
    }
  }
  return "";
};

export const readFileAsDataUrl = (file: File) =>
  new Promise<string>((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result ?? ""));
    reader.onerror = () =>
      reject(reader.error ?? new Error("Не вдалося прочитати файл."));
    reader.readAsDataURL(file);
  });

export const dataUrlToObjectUrl = (dataUrl: string) => {
  const [meta, payload = ""] = dataUrl.split(",", 2);
  const mimeMatch = /^data:([^;]+)/.exec(meta);
  const mimeType = mimeMatch?.[1] || "application/pdf";
  const binary = atob(payload);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }
  return URL.createObjectURL(new Blob([bytes], { type: mimeType }));
};

export const dataUrlToFile = (dataUrl: string, fileName: string) => {
  const [meta, payload = ""] = dataUrl.split(",", 2);
  const mimeMatch = /^data:([^;]+)/.exec(meta);
  const mimeType = mimeMatch?.[1] || "application/pdf";
  const binary = atob(payload);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }
  return new File([bytes], fileName || "questionnaire.pdf", { type: mimeType });
};

export type PersonnelRecord = {
  row: EjournalPreviewRow;
  summary: ReturnType<typeof buildPersonSummary>;
};

export const isLikelyPersonnelRow = (row: EjournalPreviewRow) => {
  const summary = buildPersonSummary(row);
  const name = summary.name.trim();

  return Boolean(
    row.__dbRowId &&
      name &&
      name !== "Особа не вибрана" &&
      name !== "-" &&
      !/^\d+([.,]\d+)?$/.test(name) &&
      !/прізвище|піб|особа/i.test(name),
  );
};

export const findEjournalPersonnelSheet = (imports: BackendEjournalImport[]) => {
  for (const item of imports) {
    const oosSheet =
      item.sheets.find((sheet) => /2\.\s*оос/i.test(sheet.name)) ??
      item.sheets.find((sheet) => /оос/i.test(sheet.name));

    if (oosSheet) return oosSheet;
  }

  return undefined;
};

export const loadAllEjournalSheetRows = async (
  sheet: BackendEjournalImportSheet,
): Promise<DbPreviewState> => {
  const firstPage = await api.listEjournalSheetRows(sheet.id, {
    limit: 500,
    offset: 0,
  });
  const columns = parseDbColumns(firstPage.columns);
  const items = [...firstPage.items];

  for (
    let offset = firstPage.items.length;
    offset < firstPage.total;
    offset += firstPage.limit
  ) {
    const nextPage = await api.listEjournalSheetRows(sheet.id, {
      limit: 500,
      offset,
    });
    items.push(...nextPage.items);
  }

  return {
    sheet,
    columns,
    rows: items.map((row) => ({
      __dbRowId: row.id,
      __rowNumber: row.excelRowNumber,
      ...row.values,
    })),
    total: firstPage.total,
    offset: 0,
    limit: items.length,
  };
};
