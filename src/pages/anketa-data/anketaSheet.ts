import { api } from "../../api";
import {
  fetchWithCache,
  jsonChanged,
  readDataCache,
  writeDataCache,
} from "../../data/idbDataCache";

/** Public Google Sheet «Анкети». */
export const ANKETA_SHEET_ID = "1aPv0wwKye-N77J7Ourko3JMnk0oOf_mqRLQEm7qcNoQ";
export const ANKETA_SHEET_GID = "0";

export const anketaSheetEditUrl = () =>
  `https://docs.google.com/spreadsheets/d/${ANKETA_SHEET_ID}/edit?gid=${ANKETA_SHEET_GID}#gid=${ANKETA_SHEET_GID}`;

/** Dev proxy path (see vite.config.ts). Falls back to direct export URL. */
export const anketaSheetCsvUrls = () => {
  const query = `format=csv&gid=${ANKETA_SHEET_GID}`;
  return [
    `/google-sheets/spreadsheets/d/${ANKETA_SHEET_ID}/export?${query}`,
    `https://docs.google.com/spreadsheets/d/${ANKETA_SHEET_ID}/export?${query}`,
  ];
};

export type AnketaColumnKey =
  | "rank"
  | "fullName"
  | "externalId"
  | "positionIndex"
  | "positionDates"
  | "positionOrderNumber"
  | "arrivedFrom"
  | "enlistDate"
  | "enlistOrderDate"
  | "enlistOrderNumber"
  | "appointmentOrderNumber"
  | "appointmentOrderDate"
  | "appointmentInDate"
  | "appointmentInNumber"
  | "rankOrderDate"
  | "rankOrderNumber"
  | "rankInDate"
  | "rankInNumber"
  | "serviceType"
  | "contractFrom"
  | "contractTo"
  | "rnokpp"
  | "rnokppRefuse"
  | "idDocumentNumber"
  | "idDocumentName"
  | "birthDate"
  | "birthPlace"
  | "sex"
  | "conscriptedWhen"
  | "conscriptedBy"
  | "education"
  | "relatives"
  | "additionalInfo"
  | "location"
  | "militaryId";

export type AnketaColumnDef = {
  key: AnketaColumnKey;
  header: string;
  size: number;
  pin?: "left";
  filterVariant?: "facet";
  /** Ідентифікуючі поля — лише з Google Sheet, не редагуються в UI. */
  readonly?: true;
};

/** Column order matches the Google Sheet header row. */
export const ANKETA_COLUMNS: AnketaColumnDef[] = [
  { key: "rank", header: "Звання", size: 150, filterVariant: "facet" },
  { key: "fullName", header: "ПІБ", size: 280, pin: "left", readonly: true },
  { key: "externalId", header: "ID", size: 110 },
  { key: "positionIndex", header: "Індекс посади", size: 150 },
  {
    key: "positionDates",
    header: "Дати прийняття посади",
    size: 160,
  },
  {
    key: "positionOrderNumber",
    header: "Наказ на посаду (номер)",
    size: 160,
  },
  { key: "arrivedFrom", header: "Звідки прибув", size: 180 },
  { key: "enlistDate", header: "Дата зарахування", size: 140 },
  {
    key: "enlistOrderDate",
    header: "Наказ зарахування — дата",
    size: 160,
  },
  {
    key: "enlistOrderNumber",
    header: "Наказ зарахування — номер",
    size: 160,
  },
  {
    key: "appointmentOrderNumber",
    header: "Наказ призначення — номер",
    size: 170,
  },
  {
    key: "appointmentOrderDate",
    header: "Наказ призначення — дата",
    size: 170,
  },
  {
    key: "appointmentInDate",
    header: "Вхідна дата наказу призначення",
    size: 180,
  },
  {
    key: "appointmentInNumber",
    header: "Вхідний номер наказу призначення",
    size: 190,
  },
  {
    key: "rankOrderDate",
    header: "Наказ звання — дата",
    size: 150,
  },
  {
    key: "rankOrderNumber",
    header: "Наказ звання — номер",
    size: 150,
  },
  {
    key: "rankInDate",
    header: "Вхідна дата наказу звання",
    size: 170,
  },
  {
    key: "rankInNumber",
    header: "Вхідний номер наказу звання",
    size: 180,
  },
  { key: "serviceType", header: "Вид служби", size: 140, filterVariant: "facet" },
  { key: "contractFrom", header: "Контракт з", size: 130 },
  { key: "contractTo", header: "Контракт до / призов", size: 150 },
  { key: "rnokpp", header: "РНОКПП", size: 130 },
  {
    key: "rnokppRefuse",
    header: "Відмова від РНОКПП",
    size: 140,
  },
  { key: "idDocumentNumber", header: "Документ (серія/номер)", size: 160 },
  {
    key: "idDocumentName",
    header: "Назва документа",
    size: 180,
  },
  { key: "birthDate", header: "Дата народження", size: 130 },
  { key: "birthPlace", header: "Місце народження", size: 200 },
  { key: "sex", header: "Стать", size: 90, filterVariant: "facet" },
  {
    key: "conscriptedWhen",
    header: "Коли призваний",
    size: 140,
  },
  {
    key: "conscriptedBy",
    header: "Ким призваний",
    size: 180,
  },
  { key: "education", header: "Освіта", size: 220 },
  { key: "relatives", header: "Дані про родичів", size: 240 },
  { key: "additionalInfo", header: "Додаткова інформація", size: 260 },
  { key: "location", header: "Місце дислокації", size: 160, filterVariant: "facet" },
  { key: "militaryId", header: "Військовий квиток", size: 160 },
];

export const isAnketaColumnReadonly = (key: AnketaColumnKey): boolean =>
  ANKETA_COLUMNS.some((column) => column.key === key && column.readonly);

export type AnketaRow = {
  __rowId: string;
  __rowNumber: number;
} & Record<AnketaColumnKey, string>;

export const createEmptyAnketaRow = (rowNumber: number): AnketaRow => {
  const record = {
    __rowId: `anketa-${rowNumber}`,
    __rowNumber: rowNumber,
  } as AnketaRow;
  for (const column of ANKETA_COLUMNS) {
    record[column.key] = "";
  }
  return record;
};

export type AnketaSheetSnapshot = {
  fetchedAt: string;
  source: "google" | "file" | "cache" | "db";
  sourceLabel: string;
  columns: AnketaColumnDef[];
  rows: AnketaRow[];
  ejoosVersionId?: string;
  reconciledAt?: string;
};

export const ANKETA_CACHE_KEY = "anketa:google-sheet:v1";
const ANKETA_SOURCE_CACHE_KEY = "anketa:google-source:v1";

const isAnketaSheetSnapshot = (
  value: unknown,
): value is AnketaSheetSnapshot => {
  if (!value || typeof value !== "object") return false;
  const record = value as Partial<AnketaSheetSnapshot>;
  return Array.isArray(record.rows) && Array.isArray(record.columns);
};

/** Persist sheet snapshot to API (shared across localhost + LAN). Best-effort. */
export const saveAnketaSnapshotToServer = async (
  snapshot: AnketaSheetSnapshot,
): Promise<boolean> => {
  try {
    await api.putAnketaSnapshot({
      payload: snapshot as unknown as Record<string, unknown>,
      source: snapshot.source === "cache" || snapshot.source === "db"
        ? "google"
        : snapshot.source,
      sourceLabel: snapshot.sourceLabel,
      fetchedAt: snapshot.fetchedAt,
      sheetId: ANKETA_SHEET_ID,
      gid: ANKETA_SHEET_GID,
    });
    return true;
  } catch {
    return false;
  }
};

export const persistAnketaSnapshot = async (
  snapshot: AnketaSheetSnapshot,
) => {
  await writeDataCache(ANKETA_CACHE_KEY, snapshot);
  return saveAnketaSnapshotToServer(snapshot);
};

export const loadAnketaSnapshotFromServer = async (): Promise<
  AnketaSheetSnapshot | null
> => {
  try {
    const record = await api.getAnketaSnapshot(
      ANKETA_SHEET_ID,
      ANKETA_SHEET_GID,
    );
    if (!record || !isAnketaSheetSnapshot(record.payload)) return null;
    const snapshot: AnketaSheetSnapshot = {
      ...record.payload,
      source: "db",
      sourceLabel:
        record.sourceLabel?.trim() ||
        record.payload.sourceLabel ||
        "БД · Анкети",
      fetchedAt: record.fetchedAt || record.payload.fetchedAt,
    };
    await writeDataCache(ANKETA_CACHE_KEY, snapshot);
    return snapshot;
  } catch {
    return null;
  }
};

const normalizeCell = (value: unknown) =>
  String(value ?? "")
    .replace(/\u00a0/g, " ")
    .replace(/\s+/g, " ")
    .trim();

const looksLikePersonName = (value: string) => {
  const text = normalizeCell(value);
  if (text.length < 5) return false;
  if (/^\d/.test(text)) return false;
  if (/облік|особового|складу|всього|разом/i.test(text)) return false;
  // Prefer Cyrillic full names with at least two words.
  const parts = text.split(/\s+/).filter(Boolean);
  if (parts.length < 2) return false;
  return /[А-Яа-яІіЇїЄєҐґ]/.test(text);
};

/** Minimal RFC4180-ish CSV parser (handles quotes + multiline cells). */
export const parseCsv = (text: string): string[][] => {
  const rows: string[][] = [];
  let row: string[] = [];
  let cell = "";
  let inQuotes = false;

  for (let i = 0; i < text.length; i += 1) {
    const ch = text[i];
    const next = text[i + 1];

    if (inQuotes) {
      if (ch === '"' && next === '"') {
        cell += '"';
        i += 1;
      } else if (ch === '"') {
        inQuotes = false;
      } else {
        cell += ch;
      }
      continue;
    }

    if (ch === '"') {
      inQuotes = true;
      continue;
    }
    if (ch === ",") {
      row.push(cell);
      cell = "";
      continue;
    }
    if (ch === "\n") {
      row.push(cell);
      rows.push(row);
      row = [];
      cell = "";
      continue;
    }
    if (ch === "\r") continue;
    cell += ch;
  }

  if (cell.length || row.length) {
    row.push(cell);
    rows.push(row);
  }

  return rows;
};

export const parseAnketaCsv = (
  csvText: string,
  meta: { source: AnketaSheetSnapshot["source"]; sourceLabel: string },
): AnketaSheetSnapshot => {
  const table = parseCsv(csvText);
  if (!table.length) {
    return {
      fetchedAt: new Date().toISOString(),
      source: meta.source,
      sourceLabel: meta.sourceLabel,
      columns: ANKETA_COLUMNS,
      rows: [],
    };
  }

  const dataRows = table.slice(1);
  const rows: AnketaRow[] = [];

  dataRows.forEach((cells, index) => {
    const values = ANKETA_COLUMNS.map((_, columnIndex) =>
      normalizeCell(cells[columnIndex] ?? ""),
    );
    const fullName = values[1] ?? "";
    if (!looksLikePersonName(fullName)) return;

    const record = {
      __rowId: `anketa-${index + 2}-${values[2] || fullName}`,
      __rowNumber: index + 2,
    } as AnketaRow;

    ANKETA_COLUMNS.forEach((column, columnIndex) => {
      record[column.key] = values[columnIndex] ?? "";
    });
    rows.push(record);
  });

  return {
    fetchedAt: new Date().toISOString(),
    source: meta.source,
    sourceLabel: meta.sourceLabel,
    columns: ANKETA_COLUMNS,
    rows,
  };
};

const fetchCsvText = async (urls: string[]) => {
  let lastError: Error | null = null;
  for (const url of urls) {
    try {
      const response = await fetch(url, { credentials: "omit" });
      if (!response.ok) {
        lastError = new Error(`HTTP ${response.status} · ${url}`);
        continue;
      }
      const text = await response.text();
      if (!text.trim() || text.trimStart().startsWith("<!DOCTYPE")) {
        lastError = new Error("Отримано HTML замість CSV (доступ обмежено).");
        continue;
      }
      return text;
    } catch (error) {
      lastError =
        error instanceof Error ? error : new Error("Невідома помилка мережі");
    }
  }
  throw lastError ?? new Error("Не вдалося завантажити Google Sheet.");
};

export const loadAnketaSheetFromGoogle = async (): Promise<AnketaSheetSnapshot> => {
  const csvText = await fetchCsvText(anketaSheetCsvUrls());
  const snapshot = parseAnketaCsv(csvText, {
    source: "google",
    sourceLabel: "Google Sheets · Анкети",
  });
  await writeDataCache(ANKETA_SOURCE_CACHE_KEY, snapshot);
  return snapshot;
};

export const loadAnketaSheetFromFile = async (
  file: File,
): Promise<AnketaSheetSnapshot> => {
  const csvText = await file.text();
  const snapshot = parseAnketaCsv(csvText, {
    source: "file",
    sourceLabel: file.name,
  });
  await writeDataCache(ANKETA_SOURCE_CACHE_KEY, snapshot);
  return snapshot;
};

export const loadCachedAnketaSheet = async () =>
  readDataCache<AnketaSheetSnapshot>(ANKETA_CACHE_KEY);

export const loadAnketaSheetPreferCache = async (options?: {
  onCached?: (snapshot: AnketaSheetSnapshot) => void | Promise<void>;
  mergeEdits?: (
    snapshot: AnketaSheetSnapshot,
  ) => AnketaSheetSnapshot | Promise<AnketaSheetSnapshot>;
  /** Не звертатися до Google, якщо спільний server/IDB snapshot уже існує. */
  refreshGoogle?: boolean;
}) => {
  const merge = async (snapshot: AnketaSheetSnapshot) =>
    options?.mergeEdits ? await options.mergeEdits(snapshot) : snapshot;
  const finalize = async (snapshot: AnketaSheetSnapshot) => {
    const merged = await merge(snapshot);
    if (options?.mergeEdits) {
      await persistAnketaSnapshot(merged);
    }
    return merged;
  };

  // 1) Shared DB snapshot (same for localhost + LAN), then local IndexedDB.
  const fromDb = await loadAnketaSnapshotFromServer();
  const cached = fromDb ?? (await loadCachedAnketaSheet());
  if (cached?.rows?.length) {
    const prepared = await merge({
      ...cached,
      source: fromDb ? "db" : "cache",
    });
    await options?.onCached?.(prepared);
    if (options?.refreshGoogle === false) {
      if (options.mergeEdits) {
        await persistAnketaSnapshot(prepared);
      }
      return prepared;
    }
  }

  try {
    const fresh = await fetchWithCache({
      key: ANKETA_SOURCE_CACHE_KEY,
      fetcher: loadAnketaSheetFromGoogle,
      isChanged: jsonChanged,
    });
    return await finalize(fresh);
  } catch (error) {
    if (cached?.rows?.length) {
      return await finalize({
        ...cached,
        source: fromDb ? ("db" as const) : ("cache" as const),
      });
    }
    throw error;
  }
};
