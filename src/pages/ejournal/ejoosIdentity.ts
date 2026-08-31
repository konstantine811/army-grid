import type { CellValue } from "../../excelRoundTrip";

export type MovementIdentityEvent = {
  personId: string;
  fullName: string;
  type: string;
  orderDate: string;
  orderNumber: string;
  previousIndex: string;
  nextIndex: string;
};

export const normKey = (value: string) =>
  value
    .toLowerCase()
    .replace(/[''`´]/g, "")
    .replace(/\s+/g, " ")
    .trim();

/** Уточнення на кшталт «(08.02.1985 р.н.)» — не інше ПІБ, а лише позначка. */
export const canonicalName = (value: string) =>
  normKey(
    value
      .replace(/\([^)]*\)/g, " ")
      .replace(/[.,;]/g, " ")
      .replace(/\s+/g, " "),
  );

export const dateMs = (value: string) => {
  const match = String(value ?? "").match(
    /(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})/,
  );
  if (!match) return 0;
  const year =
    Number(match[3]) < 100 ? 2000 + Number(match[3]) : Number(match[3]);
  const result = Date.UTC(year, Number(match[2]) - 1, Number(match[1]));
  return Number.isFinite(result) ? result : 0;
};

/** Штатний ID 1ПБ/ЕЖООС — коротке число, не РНОКПП/ІПН з 8+ цифр. */
export const isJournalPersonId = (value: string) =>
  /^\d{1,7}$/.test(value.trim());

/** ID is never an Excel date. Reject date-shaped fallback values instead of showing them as IDs. */
export const normId = (value: CellValue | unknown) => {
  if (value instanceof Date) return "";
  const text =
    typeof value === "number" && Number.isFinite(value)
      ? String(Math.trunc(value))
      : String(value ?? "")
          .replace(/\s+/g, " ")
          .trim();
  if (/^\d{1,2}[./-]\d{1,2}[./-]\d{2,4}$/.test(text)) return "";
  return text === "0" || text === "0.0" ? "" : text;
};

export const usablePersonId = (...ids: Array<string | undefined | null>) => {
  for (const id of ids) {
    const value = String(id || "").trim();
    if (value && value !== "0") return value;
  }
  return "";
};

export const createMovementKey = (event: MovementIdentityEvent) => {
  const identity = event.personId
    ? `id:${normKey(event.personId)}`
    : `name:${normKey(event.fullName)}`;
  if (identity.endsWith(":")) return "";
  return [
    identity,
    normKey(event.type),
    normKey(event.orderDate),
    normKey(event.orderNumber),
    normKey(event.previousIndex),
    normKey(event.nextIndex),
  ].join("|");
};

export const collectProcessedMovementKeys = (
  versions: Array<{ changeProtocol?: unknown }> | null | undefined,
) => {
  const keys = new Set<string>();
  for (const version of versions ?? []) {
    const protocol = version.changeProtocol;
    if (!protocol || typeof protocol !== "object") continue;
    const ops = (protocol as { ops?: unknown }).ops;
    if (!Array.isArray(ops)) continue;
    for (const op of ops) {
      if (!op || typeof op !== "object") continue;
      const key = (op as { movementKey?: unknown }).movementKey;
      if (typeof key === "string" && key.trim()) keys.add(key.trim());
    }
  }
  return keys;
};
