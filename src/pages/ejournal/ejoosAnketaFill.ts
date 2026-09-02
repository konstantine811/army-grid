import type {
  CellValue,
  ExcelSheetSnapshot,
  ExcelWorkbookSnapshot,
} from "../../excelRoundTrip";
import { isBlankPersonValue } from "../personnel/personEnrichment";
import {
  applyAnketaEditsToRows,
  loadAnketaEdits,
} from "../anketa-data/anketaEdits";
import {
  loadPersonnelIndexForAnketa,
  normalizeAnketaNameKey,
  type AnketaPersonnelMatch,
} from "../anketa-data/anketaPersonMatch";
import {
  ANKETA_MISSING_VALUE_PRESETS,
  readAnketaGapColumns,
} from "../anketa-data/anketaGaps";
import {
  isAnketaColumnReadonly,
  loadCachedAnketaSheet,
  loadAnketaSheetPreferCache,
  type AnketaColumnKey,
  type AnketaRow,
} from "../anketa-data/anketaSheet";
import {
  getPersonExternalId,
  getPersonFieldValue,
} from "../personnel/personnelUtils";
import {
  EJOOS_PERSON_DATA_START_ROW,
  EXCLUDED_ANKETA_FIXED_COLUMNS,
  formatExcludedPositionDates,
} from "./ejoosExcludedColumns";
import {
  applyInlineStringWritesToWorkbook,
} from "./ejoosZipCellWrites";

const withTimeout = async <T>(
  promise: Promise<T>,
  timeoutMs: number,
  label: string,
): Promise<T> =>
  new Promise<T>((resolve, reject) => {
    const timer = window.setTimeout(
      () => reject(new Error(`${label}: перевищено час очікування`)),
      timeoutMs,
    );
    promise.then(
      (value) => {
        window.clearTimeout(timer);
        resolve(value);
      },
      (error) => {
        window.clearTimeout(timer);
        reject(error);
      },
    );
  });

type FillField = {
  anketaKey: AnketaColumnKey;
  /** Усі токени мають бути в заголовку колонки. */
  headerParts: string[];
  excludeParts?: string[];
};

/** Поля, які можна дописати з анкети (лише якщо в аркуші порожньо). */
const SHEET_FILL_FIELDS: FillField[] = [
  { anketaKey: "rank", headerParts: ["звання"] },
  { anketaKey: "externalId", headerParts: ["id"] },
  { anketaKey: "positionIndex", headerParts: ["індекс", "посад"] },
  { anketaKey: "birthDate", headerParts: ["дата", "народж"] },
  { anketaKey: "birthPlace", headerParts: ["місце", "народж"] },
  { anketaKey: "sex", headerParts: ["стать"] },
  { anketaKey: "rnokpp", headerParts: ["рнокпп"] },
  { anketaKey: "rnokppRefuse", headerParts: ["відмова", "рнокпп"] },
  {
    anketaKey: "idDocumentName",
    headerParts: ["назва", "документ"],
  },
  {
    anketaKey: "idDocumentNumber",
    headerParts: ["серія", "номер"],
  },
  { anketaKey: "militaryId", headerParts: ["квитк"] },
  { anketaKey: "serviceType", headerParts: ["вид", "служб"] },
  { anketaKey: "location", headerParts: ["місце", "дислок"] },
  { anketaKey: "arrivedFrom", headerParts: ["звідки", "прибув"] },
  { anketaKey: "contractFrom", headerParts: ["укладан", "контракт"] },
  { anketaKey: "contractTo", headerParts: ["закінчен", "контракт"] },
  { anketaKey: "conscriptedWhen", headerParts: ["коли", "призван"] },
  { anketaKey: "conscriptedBy", headerParts: ["ким", "призван"] },
  { anketaKey: "education", headerParts: ["освіт"] },
  {
    anketaKey: "positionDates",
    headerParts: ["дат", "прийнятт", "посад"],
  },
  {
    anketaKey: "positionOrderNumber",
    headerParts: ["номер", "наказу", "прийнятт"],
  },
  {
    anketaKey: "enlistDate",
    headerParts: ["дата", "зарахуван"],
  },
  {
    anketaKey: "enlistOrderDate",
    headerParts: ["наказ", "зарахуван"],
    excludeParts: ["номер"],
  },
  {
    anketaKey: "enlistOrderNumber",
    headerParts: ["номер", "зарахуван"],
  },
  {
    anketaKey: "appointmentOrderDate",
    headerParts: ["наказ", "призначен"],
    excludeParts: ["номер", "вхідн"],
  },
  {
    anketaKey: "appointmentOrderNumber",
    headerParts: ["номер", "призначен"],
    excludeParts: ["вхідн"],
  },
  {
    anketaKey: "appointmentInDate",
    headerParts: ["вхідн", "дата", "призначен"],
  },
  {
    anketaKey: "appointmentInNumber",
    headerParts: ["вхідн", "номер", "призначен"],
  },
  {
    anketaKey: "rankOrderDate",
    headerParts: ["наказ", "зван"],
    excludeParts: ["номер", "вхідн"],
  },
  {
    anketaKey: "rankOrderNumber",
    headerParts: ["номер", "зван"],
    excludeParts: ["вхідн"],
  },
  {
    anketaKey: "rankInDate",
    headerParts: ["вхідн", "дата", "зван"],
  },
  {
    anketaKey: "rankInNumber",
    headerParts: ["вхідн", "номер", "зван"],
  },
  { anketaKey: "relatives", headerParts: ["родич"] },
  { anketaKey: "additionalInfo", headerParts: ["додатков"] },
];

export type EjoosAnketaFillTarget = "oos" | "excluded";
export type EjoosAnketaFillMode = "fill" | "merge";

export type EjoosAnketaFillReport = {
  target: EjoosAnketaFillTarget;
  mode: EjoosAnketaFillMode;
  sheetRows: number;
  matched: number;
  matchedOos: number;
  matchedPersonnel: number;
  sourceVersions: number;
  oosSourcePersons: number;
  skippedNoMatch: number;
  skippedNoUpdates: number;
  updatedRows: number;
  fieldCount: number;
  emptyWrites: number;
  richerWrites: number;
  columnsMapped: number;
  gapColumns: AnketaColumnKey[];
  unmatchedSamples: string[];
  styledCells: number;
  errors: Array<{ name: string; message: string }>;
};

const targetLabel = (target: EjoosAnketaFillTarget) =>
  target === "excluded" ? "Виключені" : "ООС";

export const formatEjoosAnketaFillReport = (report: EjoosAnketaFillReport) => {
  const sheet = targetLabel(report.target);
  const parts = [
    `рядків ${sheet}: ${report.sheetRows}`,
    report.matchedOos ? `збігів з ООС (Excel): ${report.matchedOos}` : "",
    report.matchedPersonnel
      ? `збігів зі staging: ${report.matchedPersonnel}`
      : "",
    `збігів з анкетою: ${report.matched}`,
    report.oosSourcePersons
      ? `джерело ООС: ${report.oosSourcePersons} карток / ${report.sourceVersions} версій`
      : "",
    report.mode === "merge"
      ? `колонки пропусків: ${report.gapColumns.length}`
      : "",
    `оновлено: ${report.updatedRows}`,
    report.fieldCount ? `полів: ${report.fieldCount}` : "",
    report.mode === "merge" && report.emptyWrites
      ? `порожніх: ${report.emptyWrites}`
      : "",
    report.mode === "merge" && report.richerWrites
      ? `повніших: ${report.richerWrites}`
      : "",
    report.skippedNoMatch ? `без джерела: ${report.skippedNoMatch}` : "",
    report.unmatchedSamples.length
      ? `не знайдено: ${report.unmatchedSamples.join(", ")}`
      : "",
    report.skippedNoUpdates ? `без пропусків: ${report.skippedNoUpdates}` : "",
    report.styledCells ? `стилів: ${report.styledCells}` : "",
    report.errors.length ? `помилок: ${report.errors.length}` : "",
  ].filter(Boolean);
  return parts.join(" · ");
};

const normalizeHeader = (value: CellValue | undefined) =>
  String(value ?? "")
    .toLocaleLowerCase("uk-UA")
    .replace(/[ʼ’'`´]/g, "")
    .replace(/[^\p{L}\p{N}]+/gu, " ")
    .replace(/\s+/g, " ")
    .trim();

const cellText = (value: CellValue | undefined) => {
  if (value == null) return "";
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value.toLocaleDateString("uk-UA");
  }
  if (typeof value === "number" && Number.isFinite(value)) {
    if (value > 20_000 && value < 60_000) {
      const epoch = Date.UTC(1899, 11, 30);
      const date = new Date(epoch + value * 86_400_000);
      if (!Number.isNaN(date.getTime())) {
        return date.toLocaleDateString("uk-UA");
      }
    }
    return String(value);
  }
  return String(value).trim();
};

const MISSING_DATA_PHRASES = [
  "відсутні дані",
  "немає даних",
  "немає інформації",
  "немає інф",
  "дані відсутні",
  "н/д",
  "n/a",
];

const isPlaceholderText = (value: string) => {
  if (isBlankPersonValue(value)) return true;
  const normalized = value
    .replace(/\s+/g, " ")
    .trim()
    .toLocaleLowerCase("uk-UA");
  if (!normalized) return true;
  if (/^[-–—−‑‒―_]+$/.test(normalized)) return true;
  if (
    MISSING_DATA_PHRASES.some(
      (phrase) =>
        normalized === phrase || normalized.startsWith(`${phrase} `),
    )
  ) {
    return true;
  }
  return ANKETA_MISSING_VALUE_PRESETS.some(
    (preset) =>
      normalized === preset || normalized.startsWith(`${preset} `),
  );
};

/** Порожньо / #N/A / зламаний VLOOKUP / заглушка на кшталт «відсутні дані». */
const needsFill = (value: CellValue | undefined) => {
  const text = cellText(value);
  if (isBlankPersonValue(text)) return true;
  if (text === "[object Object]") return true;
  if (/^#(N\/A|REF!|VALUE!|NAME\?|NULL!|DIV\/0!)/i.test(text)) return true;
  if (/VLOOKUP/i.test(text)) return true;
  if (/^\w+\d+;\d+;(TRUE|FALSE)\)\s*$/i.test(text)) return true;
  if (isPlaceholderText(text)) return true;
  return false;
};

const DATE_ANKETA_KEYS = new Set<AnketaColumnKey>([
  "birthDate",
  "enlistDate",
  "enlistOrderDate",
  "appointmentOrderDate",
  "appointmentInDate",
  "rankOrderDate",
  "rankInDate",
  "contractFrom",
  "contractTo",
  "conscriptedWhen",
  "positionDates",
]);

const comparableAnketaValue = (value: string, key: AnketaColumnKey) => {
  const text = value.replace(/\s+/g, " ").trim().toLocaleLowerCase("uk-UA");
  if (key === "rnokpp") return text.replace(/\D/g, "");
  if (DATE_ANKETA_KEYS.has(key)) {
    const serial = toExcelDateSerial(value);
    return serial != null ? String(serial) : text.replace(/[./-]/g, ".");
  }
  return text.replace(/[ʼ'`´]/g, "'");
};

const tokenCount = (value: string) =>
  value.split(/[\s,;·/|]+/).filter((part) => part.length > 1).length;

/** Порожньо / заглушка, або анкета явно довша / містить поточне значення. */
const isRicherAnketaValue = (
  current: CellValue | undefined,
  next: string,
  key: AnketaColumnKey,
) => {
  const incoming = next.trim();
  if (!incoming) return false;
  const existing = cellText(current);
  if (isPlaceholderText(existing)) return true;
  if (isPlaceholderText(incoming)) return false;
  const left = comparableAnketaValue(existing, key);
  const right = comparableAnketaValue(incoming, key);
  if (!right || left === right) return false;
  if (key === "rnokpp") return right.length > left.length;
  if (DATE_ANKETA_KEYS.has(key)) {
    return incoming.replace(/\D/g, "").length > existing.replace(/\D/g, "").length;
  }
  if (right.includes(left) && right.length > left.length) return true;
  if (left.includes(right)) return false;
  return (
    incoming.length - existing.length >= 6 ||
    tokenCount(incoming) - tokenCount(existing) >= 2
  );
};

const RANK_NAME_PREFIX =
  /^(?:головний\s+майстер[-\s]?сержант|старший\s+майстер[-\s]?сержант|майстер[-\s]?сержант|штаб[-\s]?сержант|головний\s+сержант|старший\s+сержант|молодший\s+сержант|сержант|старший\s+солдат|солдат|старший\s+матрос|матрос|старший\s+лейтенант|молодший\s+лейтенант|лейтенант|капітан|підполковник|полковник|майор|бригадний\s+генерал|генерал[-\s]?лейтенант|генерал[-\s]?майор|генерал|старшина)\s+/i;

const unifyNameLetters = (value: string) =>
  value
    .replace(/[iìí]/gi, "і")
    .replace(/[eèé]/gi, "е")
    .replace(/[aàá]/gi, "а")
    .replace(/[oòó]/gi, "о")
    .replace(/[y]/gi, "у");

const anketaMatchNameKey = (value: unknown) =>
  normalizeAnketaNameKey(
    unifyNameLetters(String(value ?? "").replace(RANK_NAME_PREFIX, "")),
  );

const namesLookSamePerson = (left: string, right: string) => {
  const a = anketaMatchNameKey(left);
  const b = anketaMatchNameKey(right);
  if (!a || !b) return false;
  if (a === b) return true;
  const leftParts = a.split(" ").filter(Boolean);
  const rightParts = b.split(" ").filter(Boolean);
  if (!leftParts[0] || leftParts[0] !== rightParts[0]) return false;
  if (!leftParts[1] || !rightParts[1]) return false;
  return leftParts[1][0] === rightParts[1][0];
};

const looksLikePersonName = (value: string) =>
  /[А-Яа-яІіЇїЄєҐґA-Za-z]/.test(value) && !/^#/.test(value);

const normalizePersonId = (value: unknown) => {
  const text = String(value ?? "").trim();
  if (!text || needsFill(text)) return "";
  // Excel/CSV часто дає 10901, (10901), '10901 або 10901.0.
  const compact = text
    .replace(/[\s'`(),+-]/g, "")
    .replaceAll("[", "")
    .replaceAll("]", "");
  const numeric = compact.match(/^(\d+)(?:\.0+)?$/);
  if (numeric) return numeric[1];
  return text;
};

const findTargetSheet = (
  workbook: ExcelWorkbookSnapshot,
  target: EjoosAnketaFillTarget,
) => {
  if (target === "excluded") {
    return workbook.sheets.find(
      (sheet) =>
        /3\.?\s*виключ/i.test(sheet.sheetName) ||
        /виключен/i.test(sheet.sheetName),
    );
  }
  return workbook.sheets.find(
    (sheet) =>
      /2\.?\s*оос/i.test(sheet.sheetName) ||
      /\bоос\b/i.test(sheet.sheetName) ||
      /обл[іi]к\s*особов/i.test(sheet.sheetName),
  );
};

const findHeaderRowIndex = (sheet: ExcelSheetSnapshot) => {
  const limit = Math.min(sheet.rawRows.length, 12);
  for (let index = 0; index < limit; index += 1) {
    const headers = (sheet.rawRows[index] ?? []).map(normalizeHeader);
    const hasName = headers.some(
      (header) => header.includes("прізвищ") || header.includes("піб"),
    );
    const hasId = headers.some(
      (header) => /\bid\b/.test(header) || header === "id",
    );
    const hasRank = headers.some((header) => header.includes("звання"));
    if (hasName && (hasId || hasRank)) return index;
  }
  return -1;
};

const columnIndex = (
  headerRow: CellValue[],
  parts: string[],
  exclude: string[] = [],
) =>
  headerRow.findIndex((cell) => {
    const header = normalizeHeader(cell);
    if (!header) return false;
    if (exclude.some((part) => header.includes(part))) return false;
    return parts.every((part) => header.includes(part));
  });

type AnketaLookup = {
  byId: Map<string, AnketaRow>;
  byRnokpp: Map<string, AnketaRow>;
  byName: Map<string, AnketaRow[]>;
};

const buildAnketaLookup = (rows: AnketaRow[]): AnketaLookup => {
  const byId = new Map<string, AnketaRow>();
  const byRnokpp = new Map<string, AnketaRow>();
  const byName = new Map<string, AnketaRow[]>();

  for (const row of rows) {
    const id = normalizePersonId(row.externalId);
    if (id) byId.set(id, row);
    const rnokpp = String(row.rnokpp ?? "").replace(/\D/g, "");
    if (rnokpp.length >= 8) byRnokpp.set(rnokpp, row);
    const nameKey = anketaMatchNameKey(row.fullName);
    if (!nameKey) continue;
    const list = byName.get(nameKey) ?? [];
    list.push(row);
    byName.set(nameKey, list);
  }
  return { byId, byRnokpp, byName };
};

const matchAnketaForOos = (
  lookup: AnketaLookup,
  opts: { id: string; name: string; rnokpp: string },
): AnketaRow | null => {
  const cleanId = normalizePersonId(opts.id);
  if (cleanId && lookup.byId.has(cleanId)) {
    const byId = lookup.byId.get(cleanId)!;
    if (!opts.name || namesLookSamePerson(opts.name, byId.fullName)) {
      return byId;
    }
  }
  const rnokpp = opts.rnokpp.replace(/\D/g, "");
  if (rnokpp.length >= 8 && lookup.byRnokpp.has(rnokpp)) {
    return lookup.byRnokpp.get(rnokpp)!;
  }
  const nameKey = anketaMatchNameKey(opts.name);
  if (!nameKey || !looksLikePersonName(opts.name)) return null;
  const list = lookup.byName.get(nameKey) ?? [];
  if (list.length === 1) return list[0]!;
  if (cleanId) {
    const sameId = list.filter(
      (row) => normalizePersonId(row.externalId) === cleanId,
    );
    if (sameId.length === 1) return sameId[0]!;
  }
  return null;
};

/** Мердж: спочатку ID (якщо ПІБ збігається), інакше єдине ПІБ. */
const matchAnketaByIdOrName = (
  lookup: AnketaLookup,
  opts: { id: string; name: string },
): AnketaRow | null => {
  const cleanId = normalizePersonId(opts.id);
  if (cleanId && lookup.byId.has(cleanId)) {
    const byId = lookup.byId.get(cleanId)!;
    if (!opts.name || namesLookSamePerson(opts.name, byId.fullName)) {
      return byId;
    }
  }
  const nameKey = anketaMatchNameKey(opts.name);
  if (!nameKey || !looksLikePersonName(opts.name)) return null;
  const list = lookup.byName.get(nameKey) ?? [];
  if (list.length === 1) return list[0]!;
  if (cleanId) {
    const sameId = list.filter(
      (row) => normalizePersonId(row.externalId) === cleanId,
    );
    if (sameId.length === 1) return sameId[0]!;
  }
  const withData = list.filter((row) =>
    ["conscriptedBy", "education", "relatives", "birthPlace"].some(
      (key) => !isPlaceholderText(String(row[key as AnketaColumnKey] ?? "")),
    ),
  );
  if (withData.length === 1) return withData[0]!;
  return null;
};

const normalizeFillValue = (anketaKey: AnketaColumnKey, raw: string) => {
  const trimmed = raw.trim().replace(/\r\n/g, "\n").replace(/\r/g, "\n");
  if (!trimmed) return "";
  if (anketaKey === "rnokpp") {
    const digits = trimmed.replace(/\D/g, "");
    return digits.length >= 8 ? digits : trimmed;
  }
  if (anketaKey === "positionIndex") {
    return trimmed
      .split(/\s*[·,;]\s*|\s+/)
      .map((part) => part.trim())
      .filter(Boolean)
      .join("\n");
  }
  if (anketaKey === "positionDates") {
    return formatExcludedPositionDates(trimmed) || trimmed;
  }
  return trimmed;
};

type PlannedWrite = {
  excelRow: number;
  column: number;
  value: string | number;
  field: string;
};

const EXCEL_DATE_EPOCH = Date.UTC(1899, 11, 30);

const toExcelDateSerial = (value: CellValue | undefined) => {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return (
      (Date.UTC(value.getFullYear(), value.getMonth(), value.getDate()) -
        EXCEL_DATE_EPOCH) /
      86_400_000
    );
  }
  if (typeof value === "number" && value > 20_000 && value < 60_000) {
    return value;
  }
  const text = String(value ?? "").trim();
  if (/^\d{5}(?:\.0+)?$/.test(text)) {
    const serial = Number(text);
    if (serial > 20_000 && serial < 60_000) return serial;
  }
  const match = text.match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{4})$/);
  if (!match) return null;
  const day = Number(match[1]);
  const month = Number(match[2]);
  const year = Number(match[3]);
  const timestamp = Date.UTC(year, month - 1, day);
  const date = new Date(timestamp);
  if (
    date.getUTCFullYear() !== year ||
    date.getUTCMonth() !== month - 1 ||
    date.getUTCDate() !== day
  ) {
    return null;
  }
  return (timestamp - EXCEL_DATE_EPOCH) / 86_400_000;
};

const planAnketaWritesForExcludedRow = (
  anketa: AnketaRow,
  excludedRow: CellValue[],
  excelRow: number,
  options?: { keys?: Set<AnketaColumnKey>; richer?: boolean },
) => {
  const writes: PlannedWrite[] = [];
  for (const { column, anketaKey } of EXCLUDED_ANKETA_FIXED_COLUMNS) {
    if (isAnketaColumnReadonly(anketaKey)) continue;
    if (options?.keys && !options.keys.has(anketaKey)) continue;
    const raw = String(anketa[anketaKey] ?? "").trim();
    if (!raw) continue;
    const value = normalizeFillValue(anketaKey, raw);
    if (!value) continue;
    const current = excludedRow[column - 1];
    const replace = options?.richer
      ? isRicherAnketaValue(current, String(value), anketaKey)
      : needsFill(current) ||
        comparableAnketaValue(cellText(current), anketaKey) ===
          comparableAnketaValue(String(value), anketaKey);
    if (!replace) continue;
    writes.push({
      excelRow,
      column,
      value,
      field: anketaKey,
    });
    excludedRow[column - 1] = value;
  }
  return writes;
};

/** Staging / картка ООС → колонки Виключені (семантика заголовків). */
const ANKETA_TO_PERSONNEL_PARTS: Partial<Record<AnketaColumnKey, string[]>> = {
  rank: ["звання"],
  externalId: ["id"],
  positionIndex: ["індекс", "посади"],
  positionDates: ["дати_прийняття_посади"],
  arrivedFrom: ["звідки", "прибув"],
  enlistDate: ["дата_зарахування_до_списків"],
  enlistOrderDate: ["наказ_про_зарахування_до_списків_-_дата"],
  enlistOrderNumber: ["наказ_про_зарахування_до_списків_-_номер"],
  appointmentOrderNumber: ["наказ_на_призначення_на_посаду_-_номер"],
  appointmentOrderDate: ["наказ_на_призначення_на_посаду_-_дата"],
  serviceType: ["вид_служби"],
  contractFrom: ["укладання", "контракту"],
  contractTo: ["закінчення", "контракту"],
  rnokpp: ["рнокпп_за_наявності"],
  idDocumentNumber: ["серія", "номер", "документа", "посвідчує"],
  idDocumentName: ["назва", "документа", "посвідчує"],
  birthDate: ["дата_народження"],
  birthPlace: ["місце_народження"],
  sex: ["стать"],
  conscriptedBy: ["ким", "призваний"],
  conscriptedWhen: ["коли", "призваний"],
  education: ["освіта"],
  relatives: ["дані", "родичів"],
  additionalInfo: ["додаткова_інформація"],
  militaryId: ["військового", "квитка"],
  location: ["місце_перебування"],
};

const matchPersonnelForExcluded = (
  index: Awaited<ReturnType<typeof loadPersonnelIndexForAnketa>>,
  opts: { id: string; name: string; rnokpp: string },
): AnketaPersonnelMatch | null => {
  const cleanId = normalizePersonId(opts.id);
  if (cleanId && index.byExternalId.has(cleanId)) {
    return index.byExternalId.get(cleanId)!;
  }
  const rnokpp = opts.rnokpp.replace(/\D/g, "");
  if (rnokpp.length >= 8 && index.byRnokpp.has(rnokpp)) {
    return index.byRnokpp.get(rnokpp)!;
  }
  const nameKey = normalizeAnketaNameKey(opts.name);
  if (!nameKey || !looksLikePersonName(opts.name)) return null;
  const list = index.byName.get(nameKey) ?? [];
  if (list.length === 1) return list[0]!;
  return null;
};

/**
 * rawRows стиснуті через columnIndexes — перевести Excel-колонку (1-based) у індекс масиву.
 */
const rawIndexForExcelCol = (
  sheet: ExcelSheetSnapshot,
  excelCol1: number,
) => {
  const excel0 = excelCol1 - 1;
  const idx = sheet.columnIndexes.indexOf(excel0);
  if (idx >= 0) return idx;
  const first = sheet.columnIndexes[0] ?? 0;
  return excel0 - first;
};

/**
 * Доповнює порожні / #N/A клітинки з анкет (ООС і Виключені).
 * Мердж — лише колонки пропусків, порожнє або повніше з анкети.
 */
export async function fillEjoosSheetFromAnketa(input: {
  ejoos: ExcelWorkbookSnapshot;
  target: EjoosAnketaFillTarget;
  /** merge — лише вибрані колонки пропусків, порожнє або повніше з анкети. */
  mode?: EjoosAnketaFillMode;
  anketaKeys?: AnketaColumnKey[];
  onProgress?: (done: number, total: number) => void;
  onStatus?: (message: string) => void;
}): Promise<{ blob: Blob; report: EjoosAnketaFillReport; fileName: string }> {
  const mode: EjoosAnketaFillMode = input.mode ?? "fill";
  const gapColumns =
    mode === "merge"
      ? [...new Set(input.anketaKeys ?? readAnketaGapColumns())].filter(
          (key) => !isAnketaColumnReadonly(key),
        )
      : [];
  const gapKeySet = new Set(gapColumns);
  const label = targetLabel(input.target);
  const snapshot = await withTimeout(
    loadAnketaSheetPreferCache(),
    20_000,
    "Анкетні дані",
  ).catch(() => loadCachedAnketaSheet().catch(() => null));
  const edits = snapshot ? await loadAnketaEdits().catch(() => ({})) : {};
  const anketaRows =
    snapshot?.rows?.length && Object.keys(edits).length > 0
      ? applyAnketaEditsToRows(snapshot.rows, edits)
      : snapshot?.rows ?? [];

  if (!anketaRows.length) {
    throw new Error(
      "Немає анкетних даних. Відкрийте «Анкетні дані» або оновіть таблицю з Google.",
    );
  }
  if (mode === "merge" && !gapColumns.length) {
    throw new Error(
      "Оберіть колонки пропусків в «Анкетні дані» — саме їх мердж запише в ООС / Виключені.",
    );
  }

  const targetSheet = findTargetSheet(input.ejoos, input.target);
  if (!targetSheet) {
    throw new Error(`У ЕЖООС не знайдено аркуш «${label}».`);
  }

  const headerRowIndex = findHeaderRowIndex(targetSheet);
  if (headerRowIndex < 0) {
    throw new Error(`На аркуші «${label}» не знайдено рядок заголовків.`);
  }
  const headerRow = targetSheet.rawRows[headerRowIndex] ?? [];

  // Індекси в rawRows (0-based). Для Виключені — через columnIndexes → Excel A=0,B=1,…
  const nameCol =
    input.target === "excluded"
      ? rawIndexForExcelCol(targetSheet, 2)
      : columnIndex(headerRow, ["прізвищ"]) >= 0
        ? columnIndex(headerRow, ["прізвищ"])
        : columnIndex(headerRow, ["піб"]);
  const idCol =
    input.target === "excluded"
      ? rawIndexForExcelCol(targetSheet, 3)
      : columnIndex(headerRow, ["id"]) >= 0
        ? columnIndex(headerRow, ["id"])
        : 2;
  const rnokppCol =
    input.target === "excluded"
      ? rawIndexForExcelCol(targetSheet, 17)
      : columnIndex(headerRow, ["рнокпп"]);

  if (nameCol < 0 && idCol < 0) {
    throw new Error(`Не знайдено колонки ПІБ / ID на аркуші «${label}».`);
  }

  const fieldColumns = new Map<AnketaColumnKey, number>();
  if (input.target === "oos") {
    for (const field of SHEET_FILL_FIELDS) {
      if (isAnketaColumnReadonly(field.anketaKey)) continue;
      const index = columnIndex(
        headerRow,
        field.headerParts,
        field.excludeParts ?? [],
      );
      if (index >= 0) fieldColumns.set(field.anketaKey, index);
    }
    if (mode === "merge") {
      for (const key of [...fieldColumns.keys()]) {
        if (!gapKeySet.has(key)) fieldColumns.delete(key);
      }
    }
  }

  const anketaLookup = buildAnketaLookup(anketaRows);
  const personnelIndex =
    mode === "merge" && input.target === "excluded"
      ? await withTimeout(
          loadPersonnelIndexForAnketa({ force: false }),
          20_000,
          "Картки ООС",
        ).catch(() => null)
      : null;

  const writes: PlannedWrite[] = [];
  const writeKeys = new Set<string>();
  const addWrite = (write: PlannedWrite) => {
    const key = `${write.excelRow}:${write.column}`;
    if (writeKeys.has(key)) return false;
    writeKeys.add(key);
    writes.push(write);
    return true;
  };

  const excelCol1FromRaw = (rawIndex: number) =>
    (targetSheet.columnIndexes[rawIndex] ?? rawIndex) + 1;

  /** Рядок, вирівняний під Excel-колонки (індекс 0 = A), для мапінгу VLOOKUP. */
  const toExcelAlignedRow = (rawRow: CellValue[]) => {
    const aligned: CellValue[] = [];
    for (let i = 0; i < rawRow.length; i += 1) {
      const excel0 = targetSheet.columnIndexes[i] ?? i;
      aligned[excel0] = rawRow[i] ?? null;
    }
    return aligned;
  };

  const report: EjoosAnketaFillReport = {
    target: input.target,
    mode,
    sheetRows: 0,
    matched: 0,
    matchedOos: 0,
    matchedPersonnel: 0,
    sourceVersions: 1,
    oosSourcePersons: 0,
    skippedNoMatch: 0,
    skippedNoUpdates: 0,
    updatedRows: 0,
    fieldCount: 0,
    emptyWrites: 0,
    richerWrites: 0,
    columnsMapped:
      input.target === "excluded"
        ? mode === "merge"
          ? EXCLUDED_ANKETA_FIXED_COLUMNS.filter((item) =>
              gapKeySet.has(item.anketaKey),
            ).length
          : EXCLUDED_ANKETA_FIXED_COLUMNS.length
        : fieldColumns.size,
    gapColumns,
    unmatchedSamples: [],
    styledCells: 0,
    errors: [],
  };

  // Після заголовка ще рядок нумерації полів — дані з рядка 6
  const dataStartExcelRow = Math.max(
    headerRowIndex + 2,
    EJOOS_PERSON_DATA_START_ROW,
  );
  const dataStartIndex = dataStartExcelRow - 1;
  const dataRows = targetSheet.rawRows.slice(dataStartIndex);
  const total = dataRows.length;

  for (let offset = 0; offset < dataRows.length; offset += 1) {
    if (offset > 0 && offset % 40 === 0) {
      await new Promise<void>((resolve) => {
        window.setTimeout(resolve, 0);
      });
    }
    input.onProgress?.(offset + 1, total);
    const rawRow = dataRows[offset] ?? [];
    const row =
      input.target === "excluded" ? toExcelAlignedRow(rawRow) : [...rawRow];
    const excelRow = dataStartExcelRow + offset;
    const nameRaw =
      input.target === "excluded"
        ? cellText(row[1])
        : nameCol >= 0
          ? cellText(row[nameCol])
          : "";
    const name = looksLikePersonName(nameRaw) ? nameRaw : "";
    let id =
      input.target === "excluded"
        ? cellText(row[2])
        : idCol >= 0
          ? cellText(row[idCol])
          : "";
    if (needsFill(id) || id === "[object Object]") id = "";
    const rnokppRaw =
      input.target === "excluded"
        ? cellText(row[16])
        : rnokppCol >= 0
          ? cellText(row[rnokppCol])
          : "";
    const rnokpp = needsFill(rnokppRaw) ? "" : rnokppRaw;
    if (!name && !id) continue;
    report.sheetRows += 1;

    let rowWrites = 0;
    let matchedAny = false;
    // Анкета → staging → Excel ООС (після виключення в Excel ООС часто вже немає).
    const anketa = anketaRows.length
      ? mode === "merge"
        ? matchAnketaByIdOrName(anketaLookup, { id, name })
        : matchAnketaForOos(anketaLookup, { id, name, rnokpp })
      : null;
    if (anketa) {
      matchedAny = true;
      report.matched += 1;
      const anketaId = String(anketa.externalId ?? "").trim();
      if (!id && anketaId && !needsFill(anketaId)) id = anketaId;
    }

    const personnel =
      personnelIndex != null
        ? matchPersonnelForExcluded(personnelIndex, { id, name, rnokpp })
        : null;
    if (personnel) {
      matchedAny = true;
      report.matchedPersonnel += 1;
      const persId = getPersonExternalId(personnel.row).trim();
      if (!id && persId) id = persId;
    }

    if (anketa) {
      if (input.target === "excluded") {
        for (const write of planAnketaWritesForExcludedRow(
          anketa,
          row,
          excelRow,
          mode === "merge"
            ? { keys: gapKeySet, richer: true }
            : undefined,
        )) {
          const wasEmpty = needsFill(row[write.column - 1]);
          if (addWrite(write)) {
            row[write.column - 1] = write.value;
            rowWrites += 1;
            if (wasEmpty) report.emptyWrites += 1;
            else report.richerWrites += 1;
          }
        }
      } else {
        for (const [anketaKey, column] of fieldColumns) {
          const current = row[column];
          const raw = String(anketa[anketaKey] ?? "").trim();
          if (!raw) continue;
          const value = normalizeFillValue(anketaKey, raw);
          if (!value) continue;
          const replace =
            mode === "merge"
              ? isRicherAnketaValue(current, String(value), anketaKey)
              : needsFill(current) ||
                comparableAnketaValue(cellText(current), anketaKey) ===
                  comparableAnketaValue(String(value), anketaKey);
          if (!replace) continue;
          const wasEmpty = needsFill(current);
          if (
            addWrite({
              excelRow,
              column: excelCol1FromRaw(column),
              value,
              field: anketaKey,
            })
          ) {
            row[column] = value;
            rowWrites += 1;
            if (wasEmpty) report.emptyWrites += 1;
            else report.richerWrites += 1;
          }
        }
      }
    }

    if (personnel && mode === "merge") {
      const tryPersonnelValue = (anketaKey: AnketaColumnKey) => {
        const parts = ANKETA_TO_PERSONNEL_PARTS[anketaKey];
        if (!parts) return "";
        return normalizeFillValue(
          anketaKey,
          getPersonFieldValue(personnel.row, parts),
        );
      };
      if (input.target === "excluded") {
        for (const { column, anketaKey } of EXCLUDED_ANKETA_FIXED_COLUMNS) {
          if (isAnketaColumnReadonly(anketaKey)) continue;
          if (!gapKeySet.has(anketaKey)) continue;
          const value = tryPersonnelValue(anketaKey);
          if (!value) continue;
          if (!isRicherAnketaValue(row[column - 1], String(value), anketaKey)) {
            continue;
          }
          const wasEmpty = needsFill(row[column - 1]);
          if (
            addWrite({
              excelRow,
              column,
              value,
              field: `personnel:${anketaKey}`,
            })
          ) {
            row[column - 1] = value;
            rowWrites += 1;
            if (wasEmpty) report.emptyWrites += 1;
            else report.richerWrites += 1;
          }
        }
      } else {
        for (const [anketaKey, column] of fieldColumns) {
          const value = tryPersonnelValue(anketaKey);
          if (!value) continue;
          if (!isRicherAnketaValue(row[column], String(value), anketaKey)) {
            continue;
          }
          const wasEmpty = needsFill(row[column]);
          if (
            addWrite({
              excelRow,
              column: excelCol1FromRaw(column),
              value,
              field: `personnel:${anketaKey}`,
            })
          ) {
            row[column] = value;
            rowWrites += 1;
            if (wasEmpty) report.emptyWrites += 1;
            else report.richerWrites += 1;
          }
        }
      }
    }

    if (!matchedAny) {
      report.skippedNoMatch += 1;
      if (report.unmatchedSamples.length < 5) {
        report.unmatchedSamples.push(
          `${name || "без ПІБ"}${id ? ` [ID ${id}]` : ""}`,
        );
      }
    } else if (!rowWrites) {
      report.skippedNoUpdates += 1;
    } else {
      report.updatedRows += 1;
      report.fieldCount += rowWrites;
    }
  }

  const dateStamp = new Date().toISOString().slice(0, 10);
  const fileName =
    input.target === "excluded"
      ? `ЄЖООС_Виключені_з_ООС_анкет_${dateStamp}.xlsx`
      : `ЄЖООС_ООС_з_анкет_${dateStamp}.xlsx`;

  try {
    const sheetRef: string | RegExp =
      input.target === "excluded" ? /виключ/i : targetSheet.sheetName;

    let blob: Blob = input.ejoos.file;
    if (writes.length) {
      input.onStatus?.(
        `Записую ${writes.length} змін в «${label}»…`,
      );
      await new Promise<void>((resolve) => {
        window.setTimeout(resolve, 0);
      });
      blob = await applyInlineStringWritesToWorkbook(
        blob,
        sheetRef,
        writes.map((write) => {
          const fieldKey = String(write.field).replace(
            /^personnel:/,
            "",
          ) as AnketaColumnKey;
          const text = typeof write.value === "string" ? write.value : "";
          const dateSerial =
            typeof write.value === "string" && DATE_ANKETA_KEYS.has(fieldKey)
              ? toExcelDateSerial(write.value)
              : null;
          const numericId =
            (fieldKey === "rnokpp" || fieldKey === "externalId") &&
            /^[1-9]\d{0,9}$/.test(text)
              ? Number(text)
              : null;
          const value = dateSerial ?? numericId ?? write.value;
          return {
            row: write.excelRow,
            column: write.column,
            value,
            wrapText:
              typeof value === "string" && value.includes("\n"),
          };
        }),
        {
          onProgress: (done, total) => {
            input.onStatus?.(
              `Записую в «${label}»… ${done}/${total}`,
            );
          },
          onStatus: (text) => input.onStatus?.(text),
        },
      );
    }
    input.onProgress?.(Math.max(1, total), Math.max(1, total));
    return { blob, report, fileName };
  } catch (error) {
    report.errors.push({
      name: targetSheet.sheetName,
      message: error instanceof Error ? error.message : "zip write failed",
    });
    throw error;
  }
}

export const fillEjoosOosFromAnketa = (input: {
  ejoos: ExcelWorkbookSnapshot;
  onProgress?: (done: number, total: number) => void;
}) => fillEjoosSheetFromAnketa({ ...input, target: "oos" });

export const fillEjoosExcludedFromAnketa = (input: {
  ejoos: ExcelWorkbookSnapshot;
  onProgress?: (done: number, total: number) => void;
}) => fillEjoosSheetFromAnketa({ ...input, target: "excluded" });
