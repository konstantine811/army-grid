import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import { normalizeAnketaExternalIdKey } from "../anketa-data/anketaPersonMatch";
import {
  getPersonExternalId,
  resolvePersonIdentityKey,
} from "./personnelUtils";

export const PERSONNEL_FOCUS_KEY = "army-grid:focus-personnel";

export type PersonnelFocusTarget = {
  rowId: string;
  externalId: string;
};

const normalizeFocusPart = (value: unknown) => String(value ?? "").trim();

export const normalizePersonnelFocusTarget = (target: {
  rowId?: unknown;
  externalId?: unknown;
}): PersonnelFocusTarget => ({
  rowId: normalizeFocusPart(target.rowId),
  externalId: normalizeFocusPart(target.externalId),
});

export const storePersonnelFocusTarget = (target: {
  rowId?: unknown;
  externalId?: unknown;
}) => {
  const normalized = normalizePersonnelFocusTarget(target);
  if (!normalized.rowId && !normalized.externalId) return;
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
  });
  if (fromQuery.rowId || fromQuery.externalId) return fromQuery;

  try {
    const raw = window.localStorage.getItem(PERSONNEL_FOCUS_KEY);
    if (!raw) return { rowId: "", externalId: "" };
    return normalizePersonnelFocusTarget(JSON.parse(raw) as PersonnelFocusTarget);
  } catch {
    return { rowId: "", externalId: "" };
  }
};

export const clearPersonnelFocusTarget = () => {
  try {
    window.localStorage.removeItem(PERSONNEL_FOCUS_KEY);
  } catch {
    // ignore
  }
  const url = new URL(window.location.href);
  if (url.searchParams.has("rowId") || url.searchParams.has("externalId")) {
    url.searchParams.delete("rowId");
    url.searchParams.delete("externalId");
    window.history.replaceState(
      { page: "personnel" },
      "",
      `${url.pathname}${url.search}${url.hash}`,
    );
  }
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

  if (rowId) {
    const byRowId = rows.find(
      (row) => normalizeFocusPart(row.__dbRowId) === rowId,
    );
    if (byRowId) return byRowId;
  }

  if (externalId) {
    return rows.find((row) => externalIdMatches(row, externalId)) ?? null;
  }

  return null;
};

export const buildPersonnelFocusTargetFromRow = (
  row: EjournalPreviewRow,
): PersonnelFocusTarget =>
  normalizePersonnelFocusTarget({
    rowId: row.__dbRowId,
    externalId: resolvePersonIdentityKey(row) || getPersonExternalId(row),
  });
