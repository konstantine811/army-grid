/** Історія посад / наказів в ООС — кожне значення з нового рядка, як у шаблоні. */

export const OOS_HISTORY_COLUMNS = [4, 5, 11, 12];
export const OOS_RELATIVES_COLUMN = 32;

const excelSerialToUaDate = (value: number) => {
  if (!Number.isFinite(value) || value <= 20000 || value >= 80000) return "";
  const utc = Date.UTC(1899, 11, 30) + Math.floor(value) * 86400000;
  const date = new Date(utc);
  const day = String(date.getUTCDate()).padStart(2, "0");
  const month = String(date.getUTCMonth() + 1).padStart(2, "0");
  return `${day}.${month}.${date.getUTCFullYear()}`;
};

export const oosNameKey = (value: string) =>
  value
    .toLocaleLowerCase("uk-UA")
    .replace(/[''`´]/g, "")
    .replace(/\([^)]*\)/g, " ")
    .replace(/[.,;]/g, " ")
    .replace(/\s+/g, " ")
    .trim();

export type OosIdentity = { personId: string; fullName: string };

export const oosIdentityFromOp = (op: {
  personId?: string;
  fullName?: string;
  payload?: Record<string, string>;
}): OosIdentity => ({
  personId: String(
    op.payload?.nextPersonId ||
      op.personId ||
      op.payload?.fromPersonId ||
      "",
  ).trim(),
  fullName: String(
    op.payload?.nextName || op.fullName || op.payload?.fromName || "",
  ).trim(),
});

export const oosIdentityAliasKeys = (identity: OosIdentity) => {
  const keys: string[] = [];
  if (identity.personId) keys.push(`id:${identity.personId}`);
  const name = oosNameKey(identity.fullName);
  if (name) keys.push(`name:${name}`);
  return keys;
};

/** ID колонки C: число 24867 — це ID, не Excel-дата. */
export const oosPersonIdText = (value: unknown) => {
  if (typeof value === "number" && Number.isFinite(value)) {
    return String(Math.trunc(value));
  }
  const text = String(value ?? "")
    .replace(/\s+/g, " ")
    .trim();
  if (!text || text === "[object Object]") return "";
  if (/^\d+$/.test(text)) return text;
  return text;
};

export const OOS_NO_EMPTY_ROW_MESSAGE =
  "Немає вільного рядка в основному блоці 2. ООС";

export const findNextEmptyOosDataRow = (
  getCell: (row: number, column: number) => unknown,
  options: {
    lastRow: number;
    startRow?: number;
    reserved?: Set<number>;
    allowAppend?: boolean;
  },
) => {
  const startRow = options.startRow ?? 6;
  const lastRow = Math.max(startRow, options.lastRow);
  let sectionStart = lastRow + 1;
  for (let row = startRow; row <= lastRow; row += 1) {
    if (isOosSectionHeaderText(cellValueToOosText(getCell(row, 2)))) {
      sectionStart = row;
      break;
    }
  }
  for (let row = startRow; row < sectionStart; row += 1) {
    if (options.reserved?.has(row)) continue;
    const name = cellValueToOosText(getCell(row, 2));
    const id = oosPersonIdText(getCell(row, 3));
    if (isOosBlankOrErrorText(name) && isOosBlankOrErrorText(id)) return row;
  }
  if (options.allowAppend) {
    for (let row = lastRow + 1; row <= lastRow + 200; row += 1) {
      if (!options.reserved?.has(row)) return row;
    }
  }
  return 0;
};

export const createOosRowResolver = (input: {
  getCell: (row: number, column: number) => unknown;
  lastRow: number;
  allocateEmpty: () => number;
}) => {
  const cache = new Map<string, number>();
  const remember = (row: number, identity: OosIdentity) => {
    if (!row) return;
    for (const key of oosIdentityAliasKeys(identity)) cache.set(key, row);
  };
  const resolve = (
    identity: OosIdentity,
    options?: { knownRow?: number; create?: boolean },
  ) => {
    const known = Number(options?.knownRow || 0);
    if (known > 0) {
      remember(known, identity);
      return known;
    }
    if (identity.personId) {
      const idHit = cache.get(`id:${identity.personId}`);
      if (idHit) return idHit;
    } else {
      const nameKey = oosNameKey(identity.fullName);
      if (nameKey) {
        const nameHit = cache.get(`name:${nameKey}`);
        if (nameHit) return nameHit;
      }
    }
    const existing = findExistingOosPersonRow(input.getCell, {
      personId: identity.personId,
      fullName: identity.fullName,
      lastRow: input.lastRow,
    });
    if (existing) {
      remember(existing, identity);
      return existing;
    }
    if (!options?.create) return 0;
    const created = input.allocateEmpty();
    remember(created, identity);
    return created;
  };
  return { resolve, remember };
};

export const findDuplicateOosById = <
  T extends { excelRow: number; personId: string; fullName: string },
>(
  rows: T[],
) => {
  const byId = new Map<string, T[]>();
  for (const row of rows) {
    const id = String(row.personId || "").trim();
    if (!id) continue;
    const list = byId.get(id) ?? [];
    list.push(row);
    byId.set(id, list);
  }
  return [...byId.values()].filter((group) => group.length > 1);
};

export const findPossibleDuplicateOosByName = <
  T extends { excelRow: number; personId: string; fullName: string },
>(
  rows: T[],
) => {
  const byName = new Map<string, T[]>();
  for (const row of rows) {
    const name = oosNameKey(row.fullName);
    if (!name) continue;
    const list = byName.get(name) ?? [];
    list.push(row);
    byName.set(name, list);
  }
  return [...byName.values()].filter(
    (group) =>
      group.length > 1 &&
      group.some((row) => !String(row.personId || "").trim()),
  );
};

/** Підзаголовок блоку ООС, не картка особи. */
export const isOosSectionHeaderText = (value: string) =>
  /вибув\s+у\s+розпорядж/i.test(value) ||
  /розпорядження\s+командира/i.test(value);

export const isOosBlankOrErrorText = (value: string) =>
  !value.trim() || /^#/u.test(value.trim()) || value.trim() === "[object Object]";

/** Перша чинна картка в основному списку ООС — не дописуємо другу. */
export const findExistingOosPersonRow = (
  getCell: (row: number, column: number) => unknown,
  options: {
    personId?: string;
    fullName?: string;
    lastRow: number;
    startRow?: number;
  },
) => {
  const wantId = String(options.personId || "").trim();
  const wantName = oosNameKey(options.fullName || "");
  if (!wantId && !wantName) return 0;
  let nameHit = 0;
  const lastRow = Math.max(6, options.lastRow);
  for (let row = options.startRow ?? 6; row <= lastRow; row += 1) {
    const name = cellValueToOosText(getCell(row, 2));
    const id = oosPersonIdText(getCell(row, 3));
    if (isOosSectionHeaderText(name)) continue;
    if (isOosBlankOrErrorText(name) && isOosBlankOrErrorText(id)) continue;
    if (wantId && id === wantId) return row;
    if (wantName && oosNameKey(name) === wantName && !nameHit) {
      if (wantId && id && id !== wantId) continue;
      nameHit = row;
    }
  }
  return nameHit;
};

export const cellValueToOosText = (value: unknown) => {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    const day = String(value.getDate()).padStart(2, "0");
    const month = String(value.getMonth() + 1).padStart(2, "0");
    return `${day}.${month}.${value.getFullYear()}`;
  }
  if (typeof value === "number" && Number.isFinite(value)) {
    return excelSerialToUaDate(value) || String(Math.trunc(value));
  }
  return String(value ?? "")
    .replace(/\r\n/g, "\n")
    .replace(/\r/g, "\n")
    .trim();
};

export const splitOosHistoryLines = (raw: string) => {
  const text = cellValueToOosText(raw);
  if (!text) return [];
  const explicit = text
    .split("\n")
    .map((line) => line.trim())
    .filter(Boolean);
  if (explicit.length > 1) {
    return explicit.flatMap((line) => splitPackedOosIndexes(line));
  }
  const one = explicit[0];
  const packed = splitPackedOosIndexes(one);
  if (packed.length > 1) return packed;
  const indexes = one.match(/\d{7}/g);
  if (
    indexes &&
    indexes.length > 1 &&
    indexes.join("") === one.replace(/\s+/g, "")
  ) {
    return indexes;
  }
  const dates = [...one.matchAll(/\d{1,2}\.\d{1,2}\.\d{4}/g)].map(
    (match) => match[0],
  );
  if (dates.length > 1 && dates.join("") === one.replace(/\s+/g, "")) {
    return dates;
  }
  const dateAndTail = one.match(/^(\d{1,2}\.\d{1,2}\.\d{4})(\d{4,})$/);
  if (dateAndTail) {
    const converted = excelSerialToUaDate(Number(dateAndTail[2]));
    return [dateAndTail[1], converted || dateAndTail[2]];
  }
  return [one];
};

export const mergeOosHistoryValue = (current: unknown, next: unknown) => {
  const values = splitOosHistoryLines(cellValueToOosText(current));
  const incoming = splitOosHistoryLines(cellValueToOosText(next));
  for (const value of incoming.reverse()) {
    if (value && !values.includes(value)) values.unshift(value);
  }
  return values.join("\n");
};

/** Склеєні штатні індекси `21107862103239` → два рядки, як у шаблоні ООС. */
export const splitPackedOosIndexes = (raw: string) => {
  const packed = raw.replace(/\s+/g, "");
  if (!/^\d{14,}$/.test(packed) || packed.length % 7 !== 0) return [raw];
  return packed.match(/\d{7}/g) ?? [raw];
};

export const isJammedOosHistory = (raw: unknown) => {
  const text = cellValueToOosText(raw);
  if (!text || text.includes("\n")) return false;
  return splitOosHistoryLines(text).length > 1;
};

/** Історія посад ООС — лише числові індекси, не «РОЗПОРЯДЖЕННЯ». */
export const filterOosStaffHistoryIndexes = (raw: string) =>
  splitOosHistoryLines(raw)
    .flatMap((line) => splitPackedOosIndexes(line))
    .map((line) => line.trim())
    .filter((line) => /^\d{5,}$/.test(line))
    .join("\n");

/** Родичі в ООС: кожна особа з нового рядка, адреса — окремим рядком. */
export const formatOosRelativesText = (raw: unknown) => {
  const text = cellValueToOosText(raw);
  if (!text) return "";
  if (text.includes("\n")) {
    return text
      .split("\n")
      .map((line) => line.trim())
      .filter(Boolean)
      .join("\n");
  }
  return text
    .replace(
      /\s*(мати|батько|дружина|чоловік|син|дочка|донька|брат|сестра|довірена\s+особа)\s*:/giu,
      "\n$1 : ",
    )
    .replace(/\s*;\s*/g, ";\n")
    .replace(/,?\s*(вул\.|вулиця)\s+/giu, "\n$1 ")
    .replace(/^\n+/, "")
    .replace(/\n{2,}/g, "\n")
    .replace(/\s+:\s+/g, " : ")
    .trim();
};

export const isOosWrapColumn = (column: number) =>
  OOS_HISTORY_COLUMNS.includes(column) || column === OOS_RELATIVES_COLUMN;

/** Значення, які дописуємо (звання / ПІБ / ID / історія посад / зарахування). */
export const OOS_RESTYLE_COLUMNS = [1, 2, 3, 4, 5, 7, 8, 9, 10, 11, 12];
/** O/P — останнє звання; у сусідів часто порожні, стиль беремо з K/L. */
export const OOS_RANK_ORDER_COLUMNS = [15, 16] as const;
export const oosStyleSourceColumn = (column: number) =>
  column === 15 ? 11 : column === 16 ? 12 : undefined;
/** Увесь рядок картки ООС — A:AG, щоб шрифт і центрування були як у сусіда. */
export const OOS_LAST_DATA_COLUMN = 33;
export const OOS_DATA_COLUMNS = Array.from(
  { length: OOS_LAST_DATA_COLUMN },
  (_, index) => index + 1,
);

/**
 * Шаблон стилю ООС — зайнятий рядок з переносом у індексі, не заголовок
 * і не рядок, який зараз переписуємо (він уже може бути зламаний).
 */
export const findOosStyleSourceRow = (
  getCell: (row: number, column: number) => unknown,
  options?: { skipRows?: Iterable<number>; lastRow?: number },
) => {
  const skip = new Set(options?.skipRows ?? []);
  const lastRow = Math.max(6, options?.lastRow ?? 40);
  let fallback = 0;
  for (let row = 7; row <= lastRow; row += 1) {
    if (skip.has(row)) continue;
    const name = cellValueToOosText(getCell(row, 2));
    const index = cellValueToOosText(getCell(row, 4));
    if (!name) continue;
    if (!fallback) fallback = row;
    if (index.includes("\n")) return row;
  }
  return fallback || 7;
};
