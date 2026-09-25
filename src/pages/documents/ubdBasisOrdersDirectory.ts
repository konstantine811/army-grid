import {
  UBD_BASIS_ORDER_OPTIONS,
  type UbdBasisOrderOption,
} from "./ubdBasisOrdersData";
import {
  basisOrderKey,
  mergeBasisOrderLists,
} from "./ubdBasisOrdersImport";

export const BASIS_ORDERS_STORAGE_KEY = "army-grid:basis-orders";
export const IMPORTED_BASIS_ORDERS_STORAGE_KEY =
  "army-grid:basis-orders-imported";
export const IMPORTED_BASIS_ORDERS_EVENT = "army-grid:basis-orders-changed";

const IMPORTED_DB_NAME = "army-grid-basis-orders";
const IMPORTED_DB_STORE = "imported";
const IMPORTED_DB_KEY = "rows";

export type UbdBasisOrderRecord = UbdBasisOrderOption & {
  id: string;
};

const isRecord = (value: unknown): value is UbdBasisOrderRecord => {
  if (!value || typeof value !== "object" || Array.isArray(value)) return false;
  const row = value as Partial<UbdBasisOrderRecord>;
  return Boolean(String(row.number ?? "").trim() && String(row.date ?? "").trim());
};

export const createBasisOrderId = () =>
  `br:${Date.now().toString(36)}:${Math.random().toString(36).slice(2, 8)}`;

export const loadCustomBasisOrders = (): UbdBasisOrderRecord[] => {
  if (typeof localStorage === "undefined") return [];
  try {
    const raw = localStorage.getItem(BASIS_ORDERS_STORAGE_KEY);
    if (!raw) return [];
    const parsed = JSON.parse(raw) as unknown;
    if (!Array.isArray(parsed)) return [];
    return parsed.filter(isRecord).map((row) => ({
      id: String(row.id || createBasisOrderId()),
      number: String(row.number ?? "").trim(),
      date: String(row.date ?? "").trim(),
      location: String(row.location ?? "").trim(),
      validFrom: String(row.validFrom ?? "").trim(),
      validTo: String(row.validTo ?? "").trim(),
      note: String(row.note ?? "").trim(),
    }));
  } catch {
    return [];
  }
};

export const saveCustomBasisOrders = (rows: UbdBasisOrderRecord[]) => {
  if (typeof localStorage === "undefined") return;
  localStorage.setItem(BASIS_ORDERS_STORAGE_KEY, JSON.stringify(rows));
};

const isImportedOrder = (value: unknown): value is UbdBasisOrderOption => {
  if (!value || typeof value !== "object" || Array.isArray(value)) return false;
  const row = value as Partial<UbdBasisOrderOption>;
  return Boolean(String(row.number ?? "").trim() && String(row.date ?? "").trim());
};

/** Лише номер і дата — без імені файлу на кожному рядку. */
export const compactImportedBasisOrders = (
  rows: UbdBasisOrderOption[],
): UbdBasisOrderOption[] =>
  rows
    .filter(isImportedOrder)
    .map((row) => ({
      number: String(row.number ?? "").trim(),
      date: String(row.date ?? "").trim(),
    }));

const parseImportedBasisOrders = (raw: string | null): UbdBasisOrderOption[] => {
  if (!raw) return [];
  try {
    const parsed = JSON.parse(raw) as unknown;
    if (!Array.isArray(parsed)) return [];
    return compactImportedBasisOrders(parsed);
  } catch {
    return [];
  }
};

const readImportedBasisOrdersFromLocalStorage = (): UbdBasisOrderOption[] => {
  if (typeof localStorage === "undefined") return [];
  return parseImportedBasisOrders(
    localStorage.getItem(IMPORTED_BASIS_ORDERS_STORAGE_KEY),
  );
};

export const writeImportedBasisOrdersToLocalStorage = (
  storage: Pick<Storage, "setItem" | "removeItem">,
  rows: UbdBasisOrderOption[],
): "saved" | "rewritten" | "quota" => {
  const payload = JSON.stringify(compactImportedBasisOrders(rows));
  try {
    storage.setItem(IMPORTED_BASIS_ORDERS_STORAGE_KEY, payload);
    return "saved";
  } catch {
    try {
      storage.removeItem(IMPORTED_BASIS_ORDERS_STORAGE_KEY);
      storage.setItem(IMPORTED_BASIS_ORDERS_STORAGE_KEY, payload);
      return "rewritten";
    } catch {
      return "quota";
    }
  }
};

let importedBasisOrdersMemory: UbdBasisOrderOption[] | null = null;
let importedBasisWriteGeneration = 0;
let importedBasisHydration: Promise<UbdBasisOrderOption[]> | null = null;

const publishImportedBasisOrders = (rows: UbdBasisOrderOption[]) => {
  importedBasisOrdersMemory = compactImportedBasisOrders(rows);
  if (typeof window !== "undefined") {
    window.dispatchEvent(new Event(IMPORTED_BASIS_ORDERS_EVENT));
  }
  return importedBasisOrdersMemory;
};

const openImportedBasisDb = () =>
  new Promise<IDBDatabase>((resolve, reject) => {
    const request = indexedDB.open(IMPORTED_DB_NAME, 1);
    request.onupgradeneeded = () => {
      const db = request.result;
      if (!db.objectStoreNames.contains(IMPORTED_DB_STORE)) {
        db.createObjectStore(IMPORTED_DB_STORE);
      }
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });

const readImportedBasisOrdersFromIdb = async (): Promise<
  UbdBasisOrderOption[] | null
> => {
  if (typeof indexedDB === "undefined") return null;
  const db = await openImportedBasisDb();
  try {
    return await new Promise((resolve, reject) => {
      const request = db
        .transaction(IMPORTED_DB_STORE, "readonly")
        .objectStore(IMPORTED_DB_STORE)
        .get(IMPORTED_DB_KEY);
      request.onsuccess = () => {
        const value = request.result;
        resolve(Array.isArray(value) ? compactImportedBasisOrders(value) : null);
      };
      request.onerror = () => reject(request.error);
    });
  } finally {
    db.close();
  }
};

const writeImportedBasisOrdersToIdb = async (rows: UbdBasisOrderOption[]) => {
  if (typeof indexedDB === "undefined") return;
  const db = await openImportedBasisDb();
  try {
    await new Promise<void>((resolve, reject) => {
      const request = db
        .transaction(IMPORTED_DB_STORE, "readwrite")
        .objectStore(IMPORTED_DB_STORE)
        .put(compactImportedBasisOrders(rows), IMPORTED_DB_KEY);
      request.onsuccess = () => resolve();
      request.onerror = () => reject(request.error);
    });
  } finally {
    db.close();
  }
};

export const loadImportedBasisOrders = (): UbdBasisOrderOption[] => {
  if (importedBasisOrdersMemory) return importedBasisOrdersMemory;
  importedBasisOrdersMemory = readImportedBasisOrdersFromLocalStorage();
  return importedBasisOrdersMemory;
};

export const ensureImportedBasisOrdersHydrated = () => {
  if (importedBasisHydration) return importedBasisHydration;
  const generationAtStart = importedBasisWriteGeneration;
  importedBasisHydration = (async () => {
    const local = readImportedBasisOrdersFromLocalStorage();
    let stored: UbdBasisOrderOption[] | null = null;
    try {
      stored = await readImportedBasisOrdersFromIdb();
    } catch {
      stored = null;
    }
    if (generationAtStart !== importedBasisWriteGeneration) {
      return loadImportedBasisOrders();
    }
    const chosen =
      (stored?.length ?? 0) >= local.length ? (stored ?? local) : local;
    if ((stored?.length ?? 0) < local.length) {
      void writeImportedBasisOrdersToIdb(local).catch(() => undefined);
    }
    return publishImportedBasisOrders(chosen);
  })().finally(() => {
    importedBasisHydration = null;
  });
  return importedBasisHydration;
};

export const saveImportedBasisOrders = (rows: UbdBasisOrderOption[]) => {
  const compact = publishImportedBasisOrders(rows);
  importedBasisWriteGeneration += 1;
  if (typeof localStorage !== "undefined") {
    writeImportedBasisOrdersToLocalStorage(localStorage, compact);
  }
  void writeImportedBasisOrdersToIdb(compact).catch(() => undefined);
  return compact;
};

export const allKnownBasisOrderKeys = () => {
  const keys = new Set<string>();
  for (const row of [
    ...loadCustomBasisOrders(),
    ...loadImportedBasisOrders(),
    ...UBD_BASIS_ORDER_OPTIONS,
  ]) {
    keys.add(basisOrderKey(row));
  }
  return keys;
};

export const importBasisOrdersFromParsed = (
  incoming: UbdBasisOrderOption[],
  _sourceNote = "",
) => {
  const existing = loadImportedBasisOrders();
  const { merged, added, skipped } = mergeBasisOrderLists(
    existing,
    compactImportedBasisOrders(incoming),
    allKnownBasisOrderKeys(),
  );
  saveImportedBasisOrders(merged);
  return { merged, added, skipped, total: merged.length };
};

/** Довідник користувача + імпортовані + полковий список. Записи з локацією мають пріоритет. */
export const allBasisOrderOptions = (): UbdBasisOrderOption[] => [
  ...loadCustomBasisOrders(),
  ...loadImportedBasisOrders(),
  ...UBD_BASIS_ORDER_OPTIONS,
];
