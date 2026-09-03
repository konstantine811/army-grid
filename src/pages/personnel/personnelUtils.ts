import type {
  BackendEjournalImport,
  BackendEjournalImportSheet,
  BackendPersonDocument,
  EjournalRowActionType,
} from "../../api";
import { api } from "../../api";
import {
  CacheKeys,
  fetchWithCache,
  jsonChanged,
  readDataCache,
  writeDataCache,
} from "../../data/idbDataCache";
import {
  dataUrlToUint8Array,
  downloadBlob,
  sanitizeFileName,
} from "../../shared/browserExport";
import type { DbPreviewState, EjournalPreviewRow } from "../ejournal/ejournalTypes";
import { readRosterColumnValue } from "../excel-fill/rosterSourceSnapshot";
import {
  parseDbColumns,
  previewValueToDisplay,
} from "../ejournal/ejournalUtils";

const ROSTER_FIELD_PREFIX = "roster__";

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

const ROSTER_FIELD_PREFIX = "roster__";

/** Позивний з дужок / «позивний …» / імені файлу анкети (не дата народження). */
export const isLikelyBirthDateToken = (value: string) => {
  const text = String(value ?? "").trim();
  if (!text) return false;
  if (/\d{1,2}[.\/-]\d{1,2}[.\/-]\d{2,4}/.test(text)) return true;
  if (/р\.?\s*н\.?/i.test(text)) return true;
  if (/^\d{1,2}\s+\S+\s+\d{4}/.test(text)) return true;
  return false;
};

const rosterSourceKey = (key: string) =>
  key.startsWith(ROSTER_FIELD_PREFIX)
    ? key.slice(ROSTER_FIELD_PREFIX.length)
    : key;

const CUID_VALUE_RE = /^c[a-z0-9]{20,}$/i;
const UUID_VALUE_RE =
  /^[0-9a-f]{8}-[0-9a-f]{4}-[1-8][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i;

export const isUnstablePersonExternalId = (value: string) => {
  const raw = String(value ?? "").trim();
  if (
    !raw ||
    raw === "0" ||
    raw === "-" ||
    raw === "null" ||
    raw === "undefined" ||
    raw === "[object Object]"
  ) {
    return true;
  }
  if (raw.startsWith("{") || raw.startsWith("[")) return true;
  if (/^roster:/i.test(raw)) return true;
  if (CUID_VALUE_RE.test(raw) || UUID_VALUE_RE.test(raw)) return true;
  if (/^\d{4}-\d{2}-\d{2}/.test(raw)) return true;
  if (isLikelyBirthDateToken(raw)) return true;
  return false;
};

/** ID рядка staging-імпорту ЕЖООС у БД (CUID/UUID), не fingerprint чи roster-ключ. */
export const isPersistedEjournalRowId = (value: string) => {
  const raw = String(value ?? "").trim();
  if (!raw || /^roster:/i.test(raw) || /^p:/i.test(raw) || /^name:/i.test(raw)) {
    return false;
  }
  return CUID_VALUE_RE.test(raw) || UUID_VALUE_RE.test(raw);
};

const isPersonSpreadsheetIdFieldKey = (key: string) => {
  if (!key || key.startsWith("__")) return false;
  const source = rosterSourceKey(key).toLowerCase();
  if (
    source.includes("дата") ||
    source.includes("народ") ||
    source.includes("наказ")
  ) {
    return false;
  }
  if (
    source === "id" ||
    source === "ід" ||
    source === "externalid" ||
    source === "external_id"
  ) {
    return true;
  }
  return source.includes("зовнішн") || /(^|_)id$/i.test(source);
};

const personSpreadsheetIdKeyRank = (key: string) => {
  const rosterPenalty = key.startsWith(ROSTER_FIELD_PREFIX) ? 1 : 0;
  const source = rosterSourceKey(key).toLowerCase();
  const exact =
    source === "id" ||
    source === "ід" ||
    source === "externalid" ||
    source === "external_id";
  return (exact ? 0 : 2) + rosterPenalty;
};

const readPersonIdFieldValue = (value: unknown) => {
  const raw = previewValueToDisplay(value).trim();
  if (
    !raw ||
    raw === "[object Object]" ||
    raw === "null" ||
    raw === "undefined"
  ) {
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
      return "";
    }
    return "";
  }

  return raw;
};

/** Spreadsheet / OOS person number only — never a DB CUID, birth date, or roster row key. */
const personSpreadsheetIdCache = new WeakMap<EjournalPreviewRow, string>();

export const getPersonExternalId = (row: EjournalPreviewRow | null) => {
  if (!row) return "";
  const cached = personSpreadsheetIdCache.get(row);
  if (cached !== undefined) return cached;

  const keys = Object.keys(row)
    .filter(isPersonSpreadsheetIdFieldKey)
    .sort((left, right) => personSpreadsheetIdKeyRank(left) - personSpreadsheetIdKeyRank(right));

  let resolved = "";
  for (const key of keys) {
    const raw = readPersonIdFieldValue(row[key]);
    if (!raw || isUnstablePersonExternalId(raw)) continue;
    resolved = raw;
    break;
  }

  personSpreadsheetIdCache.set(row, resolved);
  return resolved;
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

const personFieldKeyCache = new WeakMap<EjournalPreviewRow, Map<string, string>>();

/** Prefer exact / shortest key match so "звання" does not hit order columns. */
export const resolvePersonFieldKey = (
  row: EjournalPreviewRow | null,
  keyParts: string[],
) => {
  if (!row || keyParts.length === 0) return "";

  const cacheKey = keyParts.join("\0");
  let cached = personFieldKeyCache.get(row);
  if (!cached) {
    cached = new Map();
    personFieldKeyCache.set(row, cached);
  }
  const hit = cached.get(cacheKey);
  if (hit !== undefined) return hit;

  const keys = Object.keys(row).filter((key) => !key.startsWith("__"));
  const parts = keyParts.map((part) => part.toLowerCase());

  let resolved = "";
  if (parts.length === 1) {
    const exact = keys.find((key) => key.toLowerCase() === parts[0]);
    if (exact) resolved = exact;
  }

  if (!resolved) {
    const joined = parts.join("_");
    const joinedExact = keys.find((key) => key.toLowerCase() === joined);
    if (joinedExact) resolved = joinedExact;
  }

  if (!resolved) {
    const matches = keys.filter((key) => {
      const normalized = key.toLowerCase();
      // «посада» must not steal «повна_посада» / roster__повна_посада.
      if (
        !parts.includes("повна") &&
        (normalized.includes("повна") || normalized.includes("повн_"))
      ) {
        return false;
      }
      return parts.every((part) => normalized.includes(part));
    });
    if (matches.length > 0) {
      // Prefer keys that start with the first part (рнокпп… over відмова_від_рнокпп),
      // then the shortest remaining match.
      resolved =
        matches.sort((left, right) => {
          const leftStarts = left.toLowerCase().startsWith(parts[0] ?? "") ? 0 : 1;
          const rightStarts = right.toLowerCase().startsWith(parts[0] ?? "")
            ? 0
            : 1;
          if (leftStarts !== rightStarts) return leftStarts - rightStarts;
          return left.length - right.length;
        })[0] ?? "";
    }
  }

  cached.set(cacheKey, resolved);
  return resolved;
};

export const getPersonFieldValue = (
  row: EjournalPreviewRow | null,
  keyParts: string[],
) => {
  const key = resolvePersonFieldKey(row, keyParts);
  return key ? previewValueToDisplay(row?.[key]) : "";
};

/** ПІБ з полів Штатки (column_14 / roster__…), коли в ЕЖООС порожньо або зіпсовано. */
export const resolvePersonDisplayNameFromRoster = (
  row: EjournalPreviewRow | null,
) => {
  if (!row) return "";
  const direct = cleanPersonDisplayName(readRosterColumnValue(row, 14));
  if (direct) return direct;

  for (const [key, value] of Object.entries(row)) {
    if (!key.startsWith(ROSTER_FIELD_PREFIX)) continue;
    const bare = key.slice(ROSTER_FIELD_PREFIX.length);
    if (
      !/^column_14$/i.test(bare) &&
      bare.toLocaleLowerCase("uk-UA") !== "піб" &&
      !bare.toLocaleLowerCase("uk-UA").includes("піб")
    ) {
      continue;
    }
    const text = cleanPersonDisplayName(previewValueToDisplay(value));
    if (text) return text;
  }
  return "";
};

export const getPersonDisplayName = (row: EjournalPreviewRow | null) => {
  const fromOos = cleanPersonDisplayName(
    getPersonFieldValue(row, ["прізвище"]) ||
      getPersonFieldValue(row, ["піб"]),
  );
  const fromRoster = resolvePersonDisplayNameFromRoster(row);
  if (
    fromOos &&
    fromOos.length >= 5 &&
    !/^(прізвище|піб|особа|№)\b/i.test(fromOos)
  ) {
    return fromOos;
  }
  return fromRoster || fromOos;
};

export type PersonFieldDef = {
  label: string;
  parts: string[];
  kind?: "text" | "multiline" | "date" | "positionIndex";
  section: "identity" | "service" | "orders" | "contacts";
};

/** Full OСС (2. ООС) field map for personnel card. */
export const PERSON_CARD_FIELDS: PersonFieldDef[] = [
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
  { label: "Місце перебування", parts: ["місце_перебування"], section: "service" },
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

const isGenericRosterColumnKey = (key: string) =>
  /^column_\d+(_\d+)?$/i.test(key.trim());

/**
 * Назви колонок «1.ОС Загальний список» за Excel-індексом (1-based).
 * У файлі частина колонок праворуч без заголовка, але з даними / списками.
 */
export const MORNING_GENERAL_LIST_COLUMN_LABELS: Record<number, string> = {
  1: "№",
  2: "Підрозділ",
  3: "Взвод",
  4: "Відділення",
  5: "Посада",
  6: "ВОС",
  7: "Повна посада",
  8: "ШПК факт",
  9: "Категорія складу",
  10: "Анкета",
  11: "Військовий квиток",
  12: "Мобілізація/контракт",
  13: "Звання",
  14: "ПІБ",
  15: "Позивний",
  16: "Дата народження",
  17: "Рік",
  18: "Повних років",
  19: "ІПН",
  20: "Група крові",
  21: "Статус",
  22: "Тип В\\С",
  23: "Статус БГ",
  24: "БЗВП/БРЕЗ",
  25: "Наявність БЗВП",
  26: "Курс БЗВП",
  27: "Відрядження (БРЕЗ)",
  28: "Обмеження",
  29: "В якому підрозділі",
  31: "Місце перебування",
  32: "Примітки",
  33: "Напрямок",
  34: "Примітка 3",
  // Колонки без заголовка в Excel (часто списки / дублікаты значень)
  37: "Статус",
  38: "Тип В\\С",
  39: "БЗВП/БРЕЗ",
  40: "Місце перебування",
  41: "Обмеження",
  42: "Статус БГ",
};

const parseGenericRosterColumnNumber = (key: string) => {
  const match = key.trim().match(/^column_(\d+)(?:_\d+)?$/i);
  if (!match) return null;
  const number = Number(match[1]);
  return Number.isFinite(number) && number > 0 ? number : null;
};

export const resolveMorningGeneralListColumnLabel = (
  sourceKey: string,
  fallback = "",
) => {
  const columnNumber = parseGenericRosterColumnNumber(sourceKey);
  if (columnNumber != null) {
    const known = MORNING_GENERAL_LIST_COLUMN_LABELS[columnNumber];
    if (known) return known;
    return fallback || `Колонка ${columnNumber}`;
  }
  return fallback;
};

/** Повна посада з ранкового (`повна_посада` / `roster__повна_посада`) або коротка посада. */
export const getPersonFullPositionTitle = (row: EjournalPreviewRow | null) => {
  if (!row) return "";

  const read = (value: unknown) => previewValueToDisplay(value).trim();

  // Явно передано з картки особи при відкритті довідки.
  const explicit = read(row.__zhbdFullPosition);
  if (explicit) return explicit;

  const keyNorm = (key: string) =>
    rosterSourceKey(key).toLocaleLowerCase("uk-UA").replace(/[\s-]+/g, "_");

  let shortFallback = "";

  for (const [key, raw] of Object.entries(row)) {
    if (key.startsWith("__") && key !== "__zhbdFullPosition") continue;
    const text = read(raw);
    if (!text) continue;
    const norm = keyNorm(key);

    if (
      norm === "повна_посада" ||
      norm.endsWith("_повна_посада") ||
      (norm.includes("повна") && norm.includes("посада"))
    ) {
      return text;
    }

    if (
      !shortFallback &&
      (norm === "посада" ||
        norm.endsWith("_посада") ||
        /^column_5(_|$)/i.test(rosterSourceKey(key))) &&
      !norm.includes("повна") &&
      !norm.includes("індекс") &&
      !norm.includes("прийняття") &&
      !norm.includes("наказу")
    ) {
      shortFallback = text;
    }
  }

  return shortFallback || getPersonFieldValue(row, ["чим", "займається"]).trim();
};

/** Значення «Повна посада» з рядка картки (як у блоці Загальний список). */
export const pickFullPositionFromPersonRow = (row: EjournalPreviewRow | null) => {
  if (!row) return "";
  const read = (value: unknown) => previewValueToDisplay(value).trim();
  for (const [key, raw] of Object.entries(row)) {
    const text = read(raw);
    if (!text) continue;
    const norm = key
      .replace(/^roster__/i, "")
      .toLocaleLowerCase("uk-UA")
      .replace(/[\s-]+/g, "_");
    if (
      norm === "повна_посада" ||
      (norm.includes("повна") && norm.includes("посада"))
    ) {
      return text;
    }
  }
  return getPersonFullPositionTitle(row);
};

/** Excel I: оф. / серж. / солд. — категорія складу, не військове звання. */
const RANK_CATEGORY_VALUE_RE = /^(оф|сер[жh]|солд)\.?$/i;
const RANK_TITLE_VALUE_RE =
  /(рекрут|солдат|матрос|сержант|старшина|прапорщик|лейтенант|капітан|майор|підполковник|полковник|генерал)/i;
const GENERIC_ENLISTED_RANK_RE = /^(солдат|матрос|рекрут)$/i;

export const isPersonnelRankCategoryValue = (value: string) =>
  RANK_CATEGORY_VALUE_RE.test(
    value.trim().toLocaleLowerCase("uk-UA").replace(/\s+/g, ""),
  );

const looksLikePersonnelRankTitle = (value: string) => {
  const text = value.trim();
  if (!text || isPersonnelRankCategoryValue(text)) return false;
  return RANK_TITLE_VALUE_RE.test(text.toLocaleLowerCase("uk-UA"));
};

const shouldSkipRankSourceKey = (sourceKey: string) => {
  const lower = sourceKey.toLocaleLowerCase("uk-UA");
  if (lower.includes("fighter_status_")) return true;
  if (lower.includes("наказ")) return true;
  if (lower.includes("шпк")) return true;
  if (lower.includes("поса")) return true;
  if (lower.includes("категор")) return true;
  return false;
};

const isPreferredRankSourceKey = (sourceKey: string) => {
  const lower = sourceKey.toLocaleLowerCase("uk-UA");
  return (
    lower.includes("звання") ||
    /^column_13(_|$)/.test(lower) ||
    lower === "m"
  );
};

const scorePersonnelRankCandidate = (
  value: string,
  sourceKey: string,
  fromRoster: boolean,
) => {
  if (!looksLikePersonnelRankTitle(value)) return Number.NEGATIVE_INFINITY;

  const lowerKey = sourceKey.toLocaleLowerCase("uk-UA");
  const lowerValue = value.trim().toLocaleLowerCase("uk-UA");
  let score = 10;

  if (lowerKey.includes("звання")) score += 25;
  // Excel M у «Загальному списку» — повна назва звання.
  if (/^column_13(_|$)/.test(lowerKey) || lowerKey === "m") score += 20;
  if (fromRoster) score += 8;
  if (/молодший|старший|головний|штаб/.test(lowerValue)) score += 8;
  if (GENERIC_ENLISTED_RANK_RE.test(lowerValue)) score -= 12;
  score += Math.min(value.trim().length, 24);
  return score;
};

const pickBestRankCandidate = (
  row: EjournalPreviewRow,
  preferredKeysOnly: boolean,
) => {
  let best = "";
  let bestScore = Number.NEGATIVE_INFINITY;

  for (const [key, raw] of Object.entries(row)) {
    if (key.startsWith("__")) continue;
    const sourceKey = rosterSourceKey(key);
    if (shouldSkipRankSourceKey(sourceKey)) continue;
    if (preferredKeysOnly && !isPreferredRankSourceKey(sourceKey)) continue;

    const value = previewValueToDisplay(raw).trim();
    if (!value) continue;

    const fromRoster =
      key.startsWith(ROSTER_FIELD_PREFIX) ||
      isGenericRosterColumnKey(sourceKey);
    const score = scorePersonnelRankCandidate(value, sourceKey, fromRoster);
    if (score > bestScore) {
      bestScore = score;
      best = value;
    }
  }

  return best;
};

/** Повна назва звання (молодший сержант), не категорія солд./серж./оф. */
export const resolvePersonRankTitle = (row: EjournalPreviewRow | null) => {
  if (!row) return "";

  const preferred = pickBestRankCandidate(row, true);
  if (preferred) return preferred;

  const fallbackTitle = pickBestRankCandidate(row, false);
  if (fallbackTitle) return fallbackTitle;

  const fallback = previewValueToDisplay(
    row[resolvePersonFieldKey(row, ["звання"])],
  ).trim();
  if (fallback && !isPersonnelRankCategoryValue(fallback)) return fallback;
  return "";
};

export const pickPreferredPersonRank = (...values: Array<string | undefined>) => {
  let best = "";
  let bestScore = Number.NEGATIVE_INFINITY;

  for (const raw of values) {
    const value = String(raw ?? "").trim();
    if (!value || isPersonnelRankCategoryValue(value)) continue;
    const score = scorePersonnelRankCandidate(value, "звання", false);
    if (score > bestScore) {
      bestScore = score;
      best = value;
    }
  }

  if (best) return best;
  return (
    values
      .map((value) => String(value ?? "").trim())
      .find((value) => value && !isPersonnelRankCategoryValue(value)) || ""
  );
};

export const extractBirthYear = (value: string) => {
  const match = value.match(/(\d{1,2})[.\/-](\d{1,2})[.\/-](\d{2,4})/);
  if (!match) {
    const yearOnly = value.trim().match(/^(19|20)\d{2}$/);
    if (yearOnly) return Number(value.trim());
    return null;
  }
  let year = Number(match[3]);
  if (year < 100) year += year >= 50 ? 1900 : 2000;
  return Number.isFinite(year) ? year : null;
};

/** Повних років на сьогодні від дати ДД.ММ.РРРР. */
export const computeFullYearsFromBirthDate = (value: string) => {
  const match = String(value ?? "")
    .trim()
    .match(/(\d{1,2})[.\/-](\d{1,2})[.\/-](\d{2,4})/);
  if (!match) return null;
  let year = Number(match[3]);
  if (year < 100) year += year >= 50 ? 1900 : 2000;
  const month = Number(match[2]);
  const day = Number(match[1]);
  if (![year, month, day].every(Number.isFinite)) return null;
  const now = new Date();
  let age = now.getFullYear() - year;
  if (
    now.getMonth() + 1 < month ||
    (now.getMonth() + 1 === month && now.getDate() < day)
  ) {
    age -= 1;
  }
  return age >= 0 && age < 120 ? age : null;
};

/** Дата народження з ООС або з «Загального списку» (напр. column_17). */
export const looksLikePersonBirthDate = (value: string) => {
  const text = String(value ?? "").trim();
  if (!text || !isLikelyBirthDateToken(text)) return false;
  const year = extractBirthYear(text);
  if (year == null) return true;
  return year >= 1955 && year <= 2008;
};

const collectRosterFieldEntries = (row: EjournalPreviewRow | null) => {
  if (!row) return [];

  return Object.entries(row)
    .filter(([key]) => key.startsWith(ROSTER_FIELD_PREFIX))
    .map(([key, raw]) => ({
      sourceKey: key.slice(ROSTER_FIELD_PREFIX.length),
      value: formatExcelDateDisplay(raw).trim(),
    }))
    .filter((entry) => entry.value);
};

export const resolvePersonBirthDate = (row: EjournalPreviewRow | null) => {
  const fromOos = formatExcelDateDisplay(
    getPersonFieldValue(row, ["дата_народження"]),
  ).trim();
  if (fromOos && looksLikePersonBirthDate(fromOos)) return fromOos;

  // Інколи дата народження помилково лежить у полі «ID».
  const fromId = formatExcelDateDisplay(
    getPersonFieldValue(row, ["id"]),
  ).trim();
  if (fromId && looksLikePersonBirthDate(fromId)) return fromId;

  const rosterEntries = collectRosterFieldEntries(row);

  const isBirthDateKey = (sourceKey: string) => {
    const lowerKey = sourceKey.toLocaleLowerCase("uk-UA");
    if (lowerKey.includes("народ")) return true;
    if (/column_16(_|$)/i.test(sourceKey)) return true;
    const mapped = resolveMorningGeneralListColumnLabel(sourceKey)
      ?.toLocaleLowerCase("uk-UA")
      .replace(/_/g, " ");
    return Boolean(mapped && mapped.includes("народ"));
  };

  const isYearKey = (sourceKey: string) => {
    const lowerKey = sourceKey.toLocaleLowerCase("uk-UA");
    if (lowerKey === "рік" || lowerKey.includes("рік")) return true;
    if (/column_17(_|$)/i.test(sourceKey)) return true;
    const mapped = resolveMorningGeneralListColumnLabel(sourceKey)
      ?.toLocaleLowerCase("uk-UA")
      .replace(/_/g, " ");
    return mapped === "рік" || mapped === "рік народження";
  };

  // 1) Дата народження зі Штатки
  for (const entry of rosterEntries) {
    if (!isBirthDateKey(entry.sourceKey)) continue;
    if (looksLikePersonBirthDate(entry.value)) return entry.value;
  }

  // 2) Якщо в «Рік» помилково лежить повна дата — підставляємо в дату народження
  for (const entry of rosterEntries) {
    if (!isYearKey(entry.sourceKey)) continue;
    if (looksLikePersonBirthDate(entry.value)) return entry.value;
  }

  // 3) Інші generic-колонки зі Штатки з валідною датою
  for (const entry of rosterEntries) {
    if (!isGenericRosterColumnKey(entry.sourceKey)) continue;
    if (looksLikePersonBirthDate(entry.value)) return entry.value;
  }

  return "";
};

/** Коротка посада зі Штатки (col 5), інакше повна. */
export const resolvePersonPositionTitle = (row: EjournalPreviewRow | null) => {
  if (!row) return "";

  const read = (value: unknown) => previewValueToDisplay(value).trim();
  const keyNorm = (key: string) =>
    rosterSourceKey(key).toLocaleLowerCase("uk-UA").replace(/[\s-]+/g, "_");

  let short = "";
  let full = "";
  for (const [key, raw] of Object.entries(row)) {
    if (key.startsWith("__")) continue;
    const text = read(raw);
    if (!text) continue;
    const norm = keyNorm(key);
    if (
      !full &&
      (norm === "повна_посада" ||
        norm.endsWith("_повна_посада") ||
        (norm.includes("повна") && norm.includes("посада")))
    ) {
      full = text;
      continue;
    }
    if (
      !short &&
      (norm === "посада" ||
        norm.endsWith("_посада") ||
        /^column_5(_|$)/i.test(rosterSourceKey(key))) &&
      !norm.includes("повна") &&
      !norm.includes("індекс") &&
      !norm.includes("прийняття") &&
      !norm.includes("наказу")
    ) {
      short = text;
    }
  }

  return (
    short ||
    full ||
    getPersonFieldValue(row, ["чим", "займається"]).trim() ||
    getPersonFullPositionTitle(row)
  );
};

/** Місце перебування зі Штатки (col 31/40) або дислокація з ООС. */
export const resolvePersonStayPlace = (row: EjournalPreviewRow | null) => {
  const fromStay =
    getPersonFieldValue(row, ["місце_перебування"]) ||
    getPersonFieldValue(row, ["перебування"]);
  if (fromStay) return fromStay;

  const rosterEntries = collectRosterFieldEntries(row);
  for (const entry of rosterEntries) {
    const lower = entry.sourceKey.toLocaleLowerCase("uk-UA").replace(/_/g, " ");
    const mapped = resolveMorningGeneralListColumnLabel(entry.sourceKey)
      ?.toLocaleLowerCase("uk-UA")
      .replace(/_/g, " ");
    if (
      lower.includes("перебуван") ||
      mapped === "місце перебування" ||
      /column_31(_|$)/i.test(entry.sourceKey) ||
      /column_40(_|$)/i.test(entry.sourceKey)
    ) {
      if (entry.value) return entry.value;
    }
  }

  return getPersonFieldValue(row, ["місце_дислокації"]);
};

/** Текст для віджета: `06.03.1975 · 51 р.` */
export const formatPersonBirthDateWithAge = (birthDate: string) => {
  const date = String(birthDate ?? "").trim();
  if (!date) return "";
  const years = computeFullYearsFromBirthDate(date);
  if (years == null) return date;
  return `${date} · ${years} р.`;
};

export const inferRosterFieldLabel = (
  sourceKey: string,
  value: string,
  rosterLabels: Record<string, string>,
) => {
  const storedLabel = rosterLabels[sourceKey]?.trim() ?? "";
  if (storedLabel && !isGenericRosterColumnKey(storedLabel)) return storedLabel;

  const fromMorningMap = resolveMorningGeneralListColumnLabel(sourceKey);
  if (fromMorningMap && !isGenericRosterColumnKey(fromMorningMap)) {
    return fromMorningMap;
  }

  const lowerKey = sourceKey.toLocaleLowerCase("uk-UA");
  if (lowerKey.includes("народ") && lowerKey.includes("дата")) {
    return "Дата народження";
  }
  if (lowerKey.includes("позив")) return "Позивний";

  const displayed = formatExcelDateDisplay(value).trim();
  if (looksLikePersonBirthDate(displayed)) return "Дата народження";

  if (fromMorningMap) return fromMorningMap;
  return storedLabel || sourceKey;
};

/** Статус зі штатки / ранкового («В строю», «Новоприбулий»), не «Статус БГ». */
export const resolvePersonRosterStatus = (
  row: EjournalPreviewRow | null,
  rosterLabels: Record<string, string> = {},
) => {
  if (!row) return "";

  for (const [key, raw] of Object.entries(row)) {
    if (key.startsWith("__") || key.includes("fighter_status_")) continue;
    const displayed = previewValueToDisplay(raw).trim();
    if (!displayed) continue;
    const sourceKey = rosterSourceKey(key);
    const label = inferRosterFieldLabel(sourceKey, displayed, rosterLabels)
      .trim()
      .toLocaleLowerCase("uk-UA")
      .replace(/_/g, " ");
    if (label === "статус") return displayed;
  }

  for (const columnNumber of [21, 37] as const) {
    const displayed = previewValueToDisplay(
      row[`column_${columnNumber}`] ?? row[`${ROSTER_FIELD_PREFIX}column_${columnNumber}`],
    ).trim();
    if (displayed) return displayed;
  }

  return "";
};

/** Категорія огляду з тексту колонки «Статус» у Штатці. */
export const classifyOverviewStatusFromRoster = (status: string) => {
  const displayed = String(status ?? "")
    .replace(/\s+/g, " ")
    .trim();
  const normalized = displayed
    .replace(/[ʼ’']/g, "")
    .toLocaleLowerCase("uk-UA");
  if (normalized.includes("відряд")) {
    return { status: "BUSINESS_TRIP", statusLabel: displayed || "Відрядження" };
  }
  if (normalized.includes("відпуст")) {
    return { status: "LEAVE", statusLabel: displayed || "Відпустка" };
  }
  if (normalized.includes("ліку") || normalized.includes("шпит")) {
    return { status: "MEDICAL", statusLabel: displayed || "Лікування" };
  }
  if (normalized.includes("сзч")) {
    return { status: "AWOL", statusLabel: displayed || "СЗЧ" };
  }
  if (normalized.includes("безв") || normalized.includes("зник")) {
    return { status: "MISSING", statusLabel: displayed || "Безвісти" };
  }
  if (
    normalized.includes("загиб") ||
    normalized.includes("помер") ||
    /(?:^|\D)200(?:\D|$)/.test(normalized)
  ) {
    return { status: "DEAD", statusLabel: displayed || "Загиблі" };
  }
  return { status: "ON_DUTY", statusLabel: displayed || "На службі" };
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
    name: getPersonDisplayName(row) || "Особа не вибрана",
    rank: resolvePersonRankTitle(row),
    externalId: resolvePersonIdentityKey(row),
    positionIndex: formatPositionIndexes(
      getPersonFieldValue(row, ["індекс", "посади"]),
    ),
    serviceType: getPersonFieldValue(row, ["вид_служби"]),
    birthDate: resolvePersonBirthDate(row),
    birthPlace: getPersonFieldValue(row, ["місце_народження"]),
    sex: getPersonFieldValue(row, ["стать"]),
    rnokpp: getPersonFieldValue(row, ["рнокпп_за_наявності"]),
    location: resolvePersonStayPlace(row),
    positionTitle: resolvePersonPositionTitle(row),
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
    callSign: resolvePersonCallSign(row),
  };
};

/** Cheap list row: name, rank, id, callsign only. Full card fields are built for the selected person. */
export const buildPersonListSummary = (
  row: EjournalPreviewRow | null,
): ReturnType<typeof buildPersonSummary> => ({
  name: getPersonDisplayName(row) || "Особа не вибрана",
  rank: resolvePersonRankTitle(row),
  externalId: resolvePersonIdentityKey(row),
  positionIndex: "",
  serviceType: "",
  birthDate: "",
  birthPlace: "",
  sex: "",
  rnokpp: "",
  location: "",
  positionTitle: "",
  arrivedFrom: "",
  education: "",
  relatives: "",
  additionalInfo: "",
  phones: [],
  phonesDisplay: [],
  militaryId: "",
  contractFrom: "",
  contractTo: "",
  callSign: resolvePersonCallSign(row),
});

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

export const collectPersonCallSignFieldValues = (
  row: EjournalPreviewRow | null,
) => {
  if (!row) return [];

  const values = new Set<string>();
  for (const [key, raw] of Object.entries(row)) {
    if (key.startsWith("__")) continue;
    if (!key.toLocaleLowerCase("uk-UA").includes("позив")) continue;
    const displayed = previewValueToDisplay(raw).trim();
    if (displayed) values.add(displayed);
  }

  return [...values];
};

const resolveDirectCallSignValue = (value: string) => {
  const text = value.trim();
  if (!text) return "";

  const extracted = extractPersonCallSign(text);
  if (extracted) return extracted;

  return isLikelyCallSignToken(text) ? text : "";
};

export const resolvePersonCallSign = (row: EjournalPreviewRow | null) => {
  for (const fieldValue of collectPersonCallSignFieldValues(row)) {
    const resolved = resolveDirectCallSignValue(fieldValue);
    if (resolved) return resolved;
  }

  if (row) {
    // Штатка col 15 — часто `column_15` / `roster__column_15` без слова «позив» у ключі.
    const fromColumn = resolveDirectCallSignValue(
      previewValueToDisplay(row.column_15).trim() ||
        previewValueToDisplay(row.roster__column_15).trim() ||
        previewValueToDisplay(row["Позивний"]).trim() ||
        previewValueToDisplay(row["позивний"]).trim() ||
        previewValueToDisplay(row["roster__Позивний"]).trim() ||
        previewValueToDisplay(row["roster__позивний"]).trim(),
    );
    if (fromColumn) return fromColumn;
  }

  const additionalInfo = formatMultilineText(
    getPersonFieldValue(row, ["додаткова_інформація"]),
  );

  return extractPersonCallSign(
    getPersonFieldValue(row, ["прізвище"]),
    additionalInfo,
  );
};

const normalizePersonIdentityText = (value: unknown) =>
  previewValueToDisplay(value)
    .replace(/[ʼ’']/g, "")
    .replace(/\([^)]*\)/g, " ")
    .replace(/[.,;:№#"/\\|()[\]{}]+/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .toLocaleLowerCase("uk-UA");

export const normalizePersonBirthKey = (value: string) => {
  const text = formatExcelDateDisplay(value).trim();
  if (!text) return "";

  const dotted = text.match(/(\d{1,2})[.\/-](\d{1,2})[.\/-](\d{2,4})/);
  if (dotted) {
    let year = Number(dotted[3]);
    if (year < 100) year += year >= 50 ? 1900 : 2000;
    const month = String(dotted[2]).padStart(2, "0");
    const day = String(dotted[1]).padStart(2, "0");
    if (!Number.isFinite(year)) return "";
    return `${year}-${month}-${day}`;
  }

  const iso = text.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) return `${iso[1]}-${iso[2]}-${iso[3]}`;
  return "";
};

export const buildPersonIdentityFingerprint = (
  name: string,
  birthDate = "",
  callSign = "",
) => {
  const nameKey = normalizePersonIdentityText(name);
  if (!nameKey || nameKey === "особа не вибрана") return "";
  const birthKey = normalizePersonBirthKey(birthDate);
  if (birthKey) return `p:${nameKey}:${birthKey}`;
  const callKey = normalizePersonIdentityText(callSign);
  if (callKey) return `p:${nameKey}:c:${callKey}`;
  return `p:${nameKey}`;
};

/** Stable key for photos / questionnaires / documents across roster re-imports. */
const personIdentityKeyCache = new WeakMap<EjournalPreviewRow, string>();

export const resolvePersonIdentityKey = (row: EjournalPreviewRow | null) => {
  if (!row) return "";
  const cached = personIdentityKeyCache.get(row);
  if (cached !== undefined) return cached;

  try {
    const spreadsheetId = getPersonExternalId(row);
    if (spreadsheetId) {
      personIdentityKeyCache.set(row, spreadsheetId);
      return spreadsheetId;
    }

    const name = getPersonDisplayName(row);
    const result = buildPersonIdentityFingerprint(
      name,
      resolvePersonBirthDate(row),
      resolvePersonCallSign(row),
    );
    personIdentityKeyCache.set(row, result);
    return result;
  } catch {
    personIdentityKeyCache.set(row, "");
    return "";
  }
};

export const collectPersonExternalIdCandidates = (
  row: EjournalPreviewRow | null,
) => {
  if (!row) return [];

  const values = new Set<string>();
  const push = (value: unknown) => {
    const text = String(value ?? "").trim();
    if (text && text !== "0") values.add(text);
  };

  for (const key of Object.keys(row)) {
    if (key === "__dbRowId" || isPersonSpreadsheetIdFieldKey(key)) {
      push(readPersonIdFieldValue(row[key]));
    }
  }
  push(row.__dbRowId);

  const name =
    getPersonFieldValue(row, ["прізвище"]) ||
    getPersonFieldValue(row, ["піб"]);
  const nameKey = normalizePersonIdentityText(name);
  const spreadsheetId = getPersonExternalId(row);
  const birthKey = normalizePersonBirthKey(resolvePersonBirthDate(row));
  const callSignKey = normalizePersonIdentityText(resolvePersonCallSign(row));
  if (nameKey) {
    if (birthKey) {
      push(`name-birth:${nameKey}:${birthKey}`);
    } else {
      push(`roster:${nameKey}`);
      push(`roster:${String(name).trim()}`);
      push(`name:${nameKey}`);
      if (callSignKey) push(`name-call:${nameKey}:${callSignKey}`);
    }
  }
  if (callSignKey && !birthKey) push(`call:${callSignKey}`);
  if (spreadsheetId) push(`roster:${spreadsheetId}`);

  const identityKey = resolvePersonIdentityKey(row);
  if (identityKey) push(identityKey);

  return [...values];
};

const capitalizePersonNamePart = (word: string) => {
  if (!word) return "";
  return word
    .split("-")
    .map((part) => {
      const lower = part.toLocaleLowerCase("uk-UA");
      return lower ? lower.charAt(0).toUpperCase() + lower.slice(1) : "";
    })
    .join("-");
};

/** ПІБ для підпису / назви файлу: ПРІЗВИЩЕ Імʼя По батькові (Позивний). */
export const formatPersonSignatureName = (
  fullName: string,
  callSign?: string | null,
) => {
  const trimmed = String(fullName ?? "").normalize("NFC").trim();
  if (!trimmed) return "";

  const withoutCallSignParen = trimmed.replace(/\s*\([^)]+\)\s*$/, "").trim();
  const parts = withoutCallSignParen.split(/\s+/).filter(Boolean);
  if (!parts.length) return "";

  const formatted = [
    parts[0].toLocaleUpperCase("uk-UA"),
    ...parts.slice(1).map(capitalizePersonNamePart),
  ].join(" ");

  const signRaw = String(
    callSign ?? extractPersonCallSign(trimmed) ?? "",
  )
    .normalize("NFC")
    .trim();
  const sign = /^(?:0|[-—]|немає|відсутн(?:ій|я)|б\/п|#n\/a|\[object object\])$/iu.test(
    signRaw,
  )
    ? ""
    : signRaw;
  return sign ? `${formatted} (${sign})` : formatted;
};

export const buildQuestionnaireExportFileName = (
  fullName: string,
  callSign?: string | null,
) => {
  const base = formatPersonSignatureName(fullName, callSign);
  return base ? `${base}.pdf` : "Анкета.pdf";
};

export const renameQuestionnaireFile = (file: File, fileName: string) =>
  file.name === fileName
    ? file
    : new File([file], fileName, { type: file.type || "application/pdf" });

export const revokeQuestionnairePreviewUrl = (url: string) => {
  if (url.startsWith("blob:")) URL.revokeObjectURL(url);
};

export const downloadQuestionnairePdf = (
  fileName: string,
  source: { file?: Blob | File; fileData?: string },
) => {
  const safeName = sanitizeFileName(fileName);
  if (source.file) {
    downloadBlob(source.file, safeName);
    return;
  }
  if (source.fileData) {
    downloadBlob(
      new Blob([dataUrlToUint8Array(source.fileData)], {
        type: "application/pdf",
      }),
      safeName,
    );
  }
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

const PERSONNEL_STATUS_AS_NAME_RE =
  /(?:^|\s)(?:вибув|відсутн|виключ|перевед|знят|загиб|зникл|тимчасово|розпоряджен|командир(?:а)?\s+в(?:ійськової)?\s*ч(?:астини)?|в\s+розпоряджен)(?:\s|$)/i;

/** Прибрати з ПІБ службові хвости на кшталт «(21.02.1979 р.н.)». */
export const cleanPersonDisplayName = (value: string) =>
  String(value ?? "")
    .replace(/\([^)]*(?:р\.?\s*н\.?|народ)[^)]*\)/gi, " ")
    .replace(/\([^)]*\d{1,2}[.\/-]\d{1,2}[.\/-]\d{2,4}[^)]*\)/g, " ")
    .replace(/\s+/g, " ")
    .trim();

/** Відсікає рядки-статуси на кшталт «ВИБУВ У РОЗПОРЯДЖЕННЯ КОМАНДИРА…». */
export const looksLikePersonnelName = (value: string) => {
  const text = cleanPersonDisplayName(value);
  if (!text || text.length < 5) return false;
  if (text === "Особа не вибрана" || text === "-") return false;
  if (/^\d+([.,]\d+)?$/.test(text)) return false;
  if (/^(прізвище|піб|особа|№)\b/i.test(text)) return false;
  if (/^(управління|рота|взвод|батальйон|група|відділення|штаб)\b/i.test(text)) {
    return false;
  }
  if (PERSONNEL_STATUS_AS_NAME_RE.test(text)) return false;
  const parts = text.split(/\s+/).filter(Boolean);
  // Подвійні/складені іноземні прізвища можуть давати 6–8 частин ПІБ
  // (наприклад «КУЕЙРОЗ ПРАТА СОТО УЛЛОА Альваро Ренан»).
  if (parts.length < 2 || parts.length > 8) return false;
  return parts.every((part) => /^[\p{L}][\p{L}'ʼ’\-]*$/u.test(part));
};

export const isLikelyPersonnelRow = (row: EjournalPreviewRow) => {
  if (!row.__dbRowId) return false;
  const fromOos = cleanPersonDisplayName(
    getPersonFieldValue(row, ["прізвище"]) ||
      getPersonFieldValue(row, ["піб"]),
  );
  const fromRoster = resolvePersonDisplayNameFromRoster(row);
  return (
    looksLikePersonnelName(fromOos) || looksLikePersonnelName(fromRoster)
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

export const sheetRowsCacheKey = (sheet: BackendEjournalImportSheet) =>
  CacheKeys.sheetRows(
    sheet.id,
    `${sheet.updatedAt}|${sheet.rowCount}|${sheet.columnCount}`,
  );

const sheetRowsToPreview = (
  sheet: BackendEjournalImportSheet,
  columns: ReturnType<typeof parseDbColumns>,
  items: Array<{
    id: string;
    excelRowNumber?: number | null;
    values: Record<string, unknown>;
  }>,
  total: number,
): DbPreviewState => ({
  sheet,
  columns,
  rows: items.map((row) => ({
    __dbRowId: row.id,
    __rowNumber: row.excelRowNumber,
    ...row.values,
  })),
  total,
  offset: 0,
  limit: items.length,
});

export const loadAllEjournalSheetRows = async (
  sheet: BackendEjournalImportSheet,
  options?: {
    onCached?: (preview: DbPreviewState) => void | Promise<void>;
    onPage?: (preview: DbPreviewState) => void | Promise<void>;
    isCancelled?: () => boolean;
  },
): Promise<DbPreviewState> => {
  const toPreview = (
    items: Array<{
      id: string;
      excelRowNumber?: number | null;
      values: Record<string, unknown>;
    }>,
    columns: ReturnType<typeof parseDbColumns>,
    total: number,
  ) => sheetRowsToPreview(sheet, columns, items, total);

  const fetchFreshParallel = async (): Promise<DbPreviewState> => {
    const pageSize = 1000;
    const firstPage = await api.listEjournalSheetRows(sheet.id, {
      limit: pageSize,
      offset: 0,
    });
    const columns = parseDbColumns(firstPage.columns);
    const items = [...firstPage.items];
    const limit = firstPage.limit || pageSize;
    const remainingOffsets: number[] = [];

    for (
      let offset = firstPage.items.length;
      offset < firstPage.total;
      offset += limit
    ) {
      remainingOffsets.push(offset);
    }

    if (remainingOffsets.length > 0) {
      const pages = await Promise.all(
        remainingOffsets.map((offset) =>
          api.listEjournalSheetRows(sheet.id, { limit, offset }),
        ),
      );
      for (const page of pages) items.push(...page.items);
    }

    return toPreview(items, columns, firstPage.total);
  };

  const fetchFreshStream = async (): Promise<DbPreviewState> => {
    const firstSize = 20;
    const nextSize = 20;
    const firstPage = await api.listEjournalSheetRows(sheet.id, {
      limit: firstSize,
      offset: 0,
    });
    const columns = parseDbColumns(firstPage.columns);
    const items = [...firstPage.items];
    const total = firstPage.total;
    await options?.onPage?.(toPreview(items, columns, total));

    let offset = firstPage.items.length;
    const pageLimit = firstPage.limit || firstSize;
    while (offset < total && !options?.isCancelled?.()) {
      const page = await api.listEjournalSheetRows(sheet.id, {
        limit: nextSize,
        offset,
      });
      if (!page.items.length) break;
      items.push(...page.items);
      const grown = page.items.length;
      offset += grown || page.limit || pageLimit;
      await options?.onPage?.(toPreview(items, columns, total));
      await new Promise<void>((resolve) => {
        window.setTimeout(resolve, 0);
      });
    }

    return toPreview(items, columns, total);
  };

  if (!options?.onPage) {
    return fetchWithCache({
      key: sheetRowsCacheKey(sheet),
      fetcher: fetchFreshParallel,
      onCached: options?.onCached,
      isChanged: jsonChanged,
    });
  }

  const cacheKey = sheetRowsCacheKey(sheet);
  const cached = await readDataCache<DbPreviewState>(cacheKey);
  if (cached) await options.onCached?.(cached);

  const fresh = await fetchFreshStream();
  const complete = fresh.total === 0 || fresh.rows.length >= fresh.total;
  if (complete && (!cached || jsonChanged(cached, fresh))) {
    await writeDataCache(cacheKey, fresh);
  }
  return fresh;
};

export * from "./personAttachments";
