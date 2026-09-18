import {
  loadPersonnelDataset,
  rosterRowsFromDataset,
} from "../../data/personnelDataset";
import {
  anketaNameKeyVariants,
} from "../anketa-data/anketaPersonMatch";
import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import { readRosterColumnValue } from "../excel-fill/rosterSourceSnapshot";
import { overviewNameMatchesQuery } from "../overview/overviewNameSearch";
import {
  buildPersonSummary,
  resolvePersonCallSign,
  resolvePersonRankTitle,
} from "../personnel/personnelUtils";
import type { PreAnketaFormFields } from "./preAnketaFields";

export type PreAnketaStaffLookupHit = {
  callsign: string;
  rank: string;
  rosterName: string;
};

let cachedRosterRows: EjournalPreviewRow[] | null = null;

export const resetPreAnketaStaffLookupCacheForTests = () => {
  cachedRosterRows = null;
};

export const loadPreAnketaStaffRosterRows = async () => {
  if (cachedRosterRows) return cachedRosterRows;
  const dataset = await loadPersonnelDataset();
  cachedRosterRows = rosterRowsFromDataset(dataset);
  return cachedRosterRows;
};

const rosterRowKey = (row: EjournalPreviewRow) =>
  String(row.__dbRowId ?? readRosterColumnValue(row, 14).trim());

const dedupeRosterRows = (rows: EjournalPreviewRow[]) => {
  const seen = new Set<string>();
  const result: EjournalPreviewRow[] = [];
  for (const row of rows) {
    const key = rosterRowKey(row);
    if (!key || seen.has(key)) continue;
    seen.add(key);
    result.push(row);
  }
  return result;
};

const rosterNameFromRow = (row: EjournalPreviewRow) =>
  readRosterColumnValue(row, 14).trim();

const nameKeysOverlap = (left: string, right: string) => {
  const leftKeys = anketaNameKeyVariants(left);
  for (const key of anketaNameKeyVariants(right)) {
    if (leftKeys.has(key)) return true;
  }
  return false;
};

/** Знаходить однозначний рядок штатки за П.І.Б. */
export const lookupPreAnketaStaffByFullName = (
  fullName: string,
  rosterRows: EjournalPreviewRow[],
): PreAnketaStaffLookupHit | null => {
  const query = fullName.trim();
  if (query.length < 3) return null;

  const exactMatches = dedupeRosterRows(
    rosterRows.filter((row) => {
      const rosterName = rosterNameFromRow(row);
      if (rosterName.length < 3) return false;
      return nameKeysOverlap(query, rosterName);
    }),
  );
  if (exactMatches.length === 1) {
    const row = exactMatches[0]!;
    return {
      callsign: resolvePersonCallSign(row).trim(),
      rank: resolvePersonRankTitle(row).trim(),
      rosterName: buildPersonSummary(row).name.trim() || rosterNameFromRow(row),
    };
  }
  if (exactMatches.length > 1) return null;

  const fuzzyMatches = dedupeRosterRows(
    rosterRows.filter((row) => {
      const rosterName = rosterNameFromRow(row);
      return (
        rosterName.length >= 3 &&
        overviewNameMatchesQuery(rosterName, query)
      );
    }),
  );
  if (fuzzyMatches.length !== 1) return null;

  const row = fuzzyMatches[0]!;
  return {
    callsign: resolvePersonCallSign(row).trim(),
    rank: resolvePersonRankTitle(row).trim(),
    rosterName: buildPersonSummary(row).name.trim() || rosterNameFromRow(row),
  };
};

export const applyPreAnketaStaffLookupHit = (
  fields: PreAnketaFormFields,
  hit: PreAnketaStaffLookupHit,
  options?: { onlyEmpty?: boolean },
) => {
  const onlyEmpty = options?.onlyEmpty !== false;
  const next = { ...fields };
  let changed = false;

  if (hit.callsign && (!onlyEmpty || !next.callsign.trim())) {
    if (!next.callsign.trim()) {
      next.callsign = hit.callsign;
      changed = true;
    }
  }
  if (hit.rank && (!onlyEmpty || !next.rank.trim())) {
    if (!next.rank.trim()) {
      next.rank = hit.rank;
      changed = true;
    }
  }

  return { fields: next, changed };
};

export const fillPreAnketaFieldsFromStaff = async (
  fields: PreAnketaFormFields,
  options?: { onlyEmpty?: boolean },
) => {
  const fullName = fields.fullName.trim();
  if (fullName.length < 3) {
    return {
      fields,
      changed: false,
      message: "Спочатку введіть П.І.Б.",
    };
  }

  const rosterRows = await loadPreAnketaStaffRosterRows();
  const hit = lookupPreAnketaStaffByFullName(fullName, rosterRows);
  if (!hit) {
    const candidates = dedupeRosterRows(
      rosterRows.filter((row) => nameKeysOverlap(fullName, rosterNameFromRow(row))),
    );
    const message =
      candidates.length > 1
        ? "У штатці кілька осіб з таким П.І.Б. — уточніть написання."
        : "У штатці не знайдено збігу за П.І.Б.";
    return { fields, changed: false, message };
  }

  const applied = applyPreAnketaStaffLookupHit(fields, hit, options);
  if (!applied.changed) {
    return {
      fields,
      changed: false,
      message: "Позивний і звання уже заповнені або відсутні в штатці.",
    };
  }

  const parts = [
    hit.callsign ? `позивний ${hit.callsign}` : "",
    hit.rank ? `звання ${hit.rank}` : "",
  ].filter(Boolean);
  return {
    fields: applied.fields,
    changed: true,
    message: `Підтягнуто зі штатки: ${parts.join(", ")}.`,
  };
};
