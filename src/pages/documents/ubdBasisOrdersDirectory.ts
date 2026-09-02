import {
  UBD_BASIS_ORDER_OPTIONS,
  type UbdBasisOrderOption,
} from "./ubdBasisOrdersData";

export const BASIS_ORDERS_STORAGE_KEY = "army-grid:basis-orders";

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

/** Довідник користувача + полковий список. Записи з локацією мають пріоритет. */
export const allBasisOrderOptions = (): UbdBasisOrderOption[] => [
  ...loadCustomBasisOrders(),
  ...UBD_BASIS_ORDER_OPTIONS,
];
