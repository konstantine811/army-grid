import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import { normalizeAnketaExternalIdKey } from "../anketa-data/anketaPersonMatch";
import {
  getPersonDisplayName,
  getPersonExternalId,
  resolvePersonIdentityKey,
} from "./personnelUtils";

export const PERSONNEL_FOCUS_KEY = "army-grid:focus-personnel";

export type PersonnelFocusTarget = {
  rowId: string;
  externalId: string;
  /** Підказка для поля пошуку в Особовому складі (наприклад, ПІБ з Огляду). */
  search: string;
};

const normalizeFocusPart = (value: unknown) => String(value ?? "").trim();

/** Overview rows use synthetic ids (`roster:…`); they are not personnel `__dbRowId`. */
export const isOverviewSyntheticPersonnelRowId = (rowId: string) =>
  rowId.startsWith("roster:") || rowId.startsWith("overview:");

export const normalizePersonnelFocusTarget = (target: {
  rowId?: unknown;
  externalId?: unknown;
  search?: unknown;
}): PersonnelFocusTarget => ({
  rowId: normalizeFocusPart(target.rowId),
  externalId: normalizeFocusPart(target.externalId),
  search: normalizeFocusPart(target.search),
});

export const storePersonnelFocusTarget = (target: {
  rowId?: unknown;
  externalId?: unknown;
  search?: unknown;
}) => {
  const normalized = normalizePersonnelFocusTarget(target);
  if (!normalized.rowId && !normalized.externalId && !normalized.search) return;
  try {
    window.localStorage.setItem(PERSONNEL_FOCUS_KEY, JSON.stringify(normalized));
  } catch {
    // ignore
  }
};

export const readPersonnelFocusTarget = (): PersonnelFocusTarget => {
  const params = new URLSearchParams(window.location.search);
  const fromQuery = normalizePersonnelFocusTarget({
    rowId: params.get("rowId"),
    externalId: params.get("externalId"),
    search: params.get("q") ?? params.get("search"),
  });
  if (fromQuery.rowId || fromQuery.externalId || fromQuery.search) {
    return fromQuery;
  }

  try {
    const raw = window.localStorage.getItem(PERSONNEL_FOCUS_KEY);
    if (!raw) return { rowId: "", externalId: "", search: "" };
    return normalizePersonnelFocusTarget(JSON.parse(raw) as PersonnelFocusTarget);
  } catch {
    return { rowId: "", externalId: "", search: "" };
  }
};

export const clearPersonnelFocusTarget = () => {
  try {
    window.localStorage.removeItem(PERSONNEL_FOCUS_KEY);
  } catch {
    // ignore
  }
  const url = new URL(window.location.href);
  if (
    url.searchParams.has("rowId") ||
    url.searchParams.has("externalId") ||
    url.searchParams.has("q") ||
    url.searchParams.has("search")
  ) {
    url.searchParams.delete("rowId");
    url.searchParams.delete("externalId");
    url.searchParams.delete("q");
    url.searchParams.delete("search");
    window.history.replaceState(
      { page: "personnel" },
      "",
      `${url.pathname}${url.search}${url.hash}`,
    );
  }
};

const normalizePersonNameHint = (value: unknown) =>
  String(value ?? "")
    .replace(/\s+/g, " ")
    .trim()
    .toLocaleLowerCase("uk-UA");

export const findPersonnelRowByNameHint = (
  rows: EjournalPreviewRow[],
  nameHint: string,
) => {
  const wanted = normalizePersonNameHint(nameHint);
  if (!wanted) return null;

  const normalizedRows = rows.map((row) => ({
    row,
    name: normalizePersonNameHint(getPersonDisplayName(row)),
  }));

  const exact = normalizedRows.filter((item) => item.name === wanted);
  if (exact.length === 1) return exact[0]!.row;

  const contains = normalizedRows.filter((item) => item.name.includes(wanted));
  if (contains.length === 1) return contains[0]!.row;

  const surname = wanted.split(" ")[0] ?? "";
  if (surname.length >= 3) {
    const bySurname = normalizedRows.filter((item) =>
      item.name.startsWith(`${surname} `),
    );
    if (bySurname.length === 1) return bySurname[0]!.row;
  }

  return null;
};

const externalIdMatches = (row: EjournalPreviewRow, externalId: string) => {
  const wanted = normalizeAnketaExternalIdKey(externalId);
  if (!wanted) return false;
  const keys = [
    resolvePersonIdentityKey(row),
    getPersonExternalId(row),
  ]
    .map((value) => normalizeFocusPart(value))
    .filter(Boolean);
  return keys.some(
    (key) => key === externalId || normalizeAnketaExternalIdKey(key) === wanted,
  );
};

export const findPersonnelRowByFocusTarget = (
  rows: EjournalPreviewRow[],
  focusTarget: PersonnelFocusTarget,
) => {
  const rowId = normalizeFocusPart(focusTarget.rowId);
  const externalId = normalizeFocusPart(focusTarget.externalId);

  if (rowId && !isOverviewSyntheticPersonnelRowId(rowId)) {
    const byRowId = rows.find(
      (row) => normalizeFocusPart(row.__dbRowId) === rowId,
    );
    if (byRowId) return byRowId;
  }

  if (externalId) {
    const byExternalId =
      rows.find((row) => externalIdMatches(row, externalId)) ?? null;
    if (byExternalId) return byExternalId;

    const byDbRowId = rows.find(
      (row) => normalizeFocusPart(row.__dbRowId) === externalId,
    );
    if (byDbRowId) return byDbRowId;
  }

  const search = normalizeFocusPart(focusTarget.search);
  if (search) {
    return findPersonnelRowByNameHint(rows, search);
  }

  return null;
};

export const buildPersonnelFocusTargetFromRow = (
  row: EjournalPreviewRow,
): PersonnelFocusTarget =>
  normalizePersonnelFocusTarget({
    rowId: row.__dbRowId,
    externalId: resolvePersonIdentityKey(row) || getPersonExternalId(row),
    search: getPersonDisplayName(row),
  });
