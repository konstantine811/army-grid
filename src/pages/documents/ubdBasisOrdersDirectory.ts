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

export const loadImportedBasisOrders = (): UbdBasisOrderOption[] => {
  if (typeof localStorage === "undefined") return [];
  try {
    const raw = localStorage.getItem(IMPORTED_BASIS_ORDERS_STORAGE_KEY);
    if (!raw) return [];
    const parsed = JSON.parse(raw) as unknown;
    if (!Array.isArray(parsed)) return [];
    return parsed.filter(isImportedOrder).map((row) => ({
      number: String(row.number ?? "").trim(),
      date: String(row.date ?? "").trim(),
      note: String(row.note ?? "").trim() || undefined,
    }));
  } catch {
    return [];
  }
};

export const saveImportedBasisOrders = (rows: UbdBasisOrderOption[]) => {
  if (typeof localStorage === "undefined") return;
  localStorage.setItem(IMPORTED_BASIS_ORDERS_STORAGE_KEY, JSON.stringify(rows));
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
  sourceNote = "",
) => {
  const existing = loadImportedBasisOrders();
  const annotated = incoming.map((row) => ({
    ...row,
    note: row.note || sourceNote || undefined,
  }));
  const { merged, added, skipped } = mergeBasisOrderLists(
    existing,
    annotated,
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
