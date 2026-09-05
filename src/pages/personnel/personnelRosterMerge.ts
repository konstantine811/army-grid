import type { BackendEjournalImportSheet } from "../../api";
import { valueToDisplay } from "../../excelRoundTrip";
import { readRosterColumnValue } from "../excel-fill/rosterSourceSnapshot";
import type { DbPreviewState, EjournalPreviewRow } from "../ejournal/ejournalTypes";
import { normalizeRosterMatchText } from "./fighterStatusImport";
import {
  cleanPersonDisplayName,
  getPersonDisplayName,
  getPersonExternalId,
  getPersonFieldValue,
  looksLikePersonBirthDate,
  looksLikePersonnelName,
  normalizePersonBirthKey,
  resolvePersonBirthDate,
  resolvePersonIdentityKey,
  resolvePersonRankTitle,
} from "./personnelUtils";
import {
  isPersonnelFromArchive,
  ROSTER_ARCHIVE_FLAG_KEY,
  ROSTER_ARCHIVE_SOURCE_KEY,
  ROSTER_ARCHIVE_SOURCE_VALUE,
} from "./staffSheetArchiveMarker";

export const ROSTER_FIELD_PREFIX = "roster__";
const normalizeRosterText = normalizeRosterMatchText;

export { isPersonnelFromArchive } from "./staffSheetArchiveMarker";

const withArchiveMarker = (
  target: EjournalPreviewRow,
  rosterRow: EjournalPreviewRow,
): EjournalPreviewRow => {
  const orderedTarget = {
    ...target,
    __rosterOrder:
      rosterRow.__rosterOrder ?? rosterRow.__rowNumber ?? Number.MAX_SAFE_INTEGER,
  };
  if (!isPersonnelFromArchive(rosterRow)) return orderedTarget;
  return {
    ...orderedTarget,
    [ROSTER_ARCHIVE_FLAG_KEY]: true,
    [ROSTER_ARCHIVE_SOURCE_KEY]: ROSTER_ARCHIVE_SOURCE_VALUE,
    [`${ROSTER_FIELD_PREFIX}${ROSTER_ARCHIVE_SOURCE_KEY}`]:
      ROSTER_ARCHIVE_SOURCE_VALUE,
  };
};

/** Особа зі штатки / ранкового «Загального списку», не лише з ЕЖООС. Архів — окремий таб. */
export const isPersonnelInStaffRoster = (
  row: EjournalPreviewRow | null | undefined,
) => {
  if (!row) return false;
  if (isPersonnelFromArchive(row)) return false;
  if (/^roster:/i.test(String(row.__dbRowId ?? ""))) return true;
  return Object.keys(row).some((key) => key.startsWith(ROSTER_FIELD_PREFIX));
};

export const getRosterValue = (row: EjournalPreviewRow, keyParts: string[]) => {
  const key = Object.keys(row).find((item) =>
    keyParts.every((part) => item.toLocaleLowerCase("uk-UA").includes(part)),
  );
  return key
    ? valueToDisplay(row[key] as Parameters<typeof valueToDisplay>[0]).trim()
    : "";
};

/** Підрозділ зі Штатки (col 2), не «місце перебування» / «в якому підрозділі». */
export const getRosterUnit = (row: EjournalPreviewRow) =>
  readRosterColumnValue(row, 2).trim();

/** ПІБ зі Штатки: іменована колонка, col 14 (N) або col 13, якщо макет зсунутий. */
export const getRosterPersonName = (row: EjournalPreviewRow) => {
  const candidates = [
    getRosterValue(row, ["піб"]),
    getRosterValue(row, ["прізвище"]),
    readRosterColumnValue(row, 14),
    readRosterColumnValue(row, 13),
    String(row.fullName ?? "").trim(),
  ]
    .map((value) => cleanPersonDisplayName(value))
    .filter(Boolean);
  return candidates.find((value) => looksLikePersonnelName(value)) || candidates[0] || "";
};

/** Той самий фільтр «нова», що в БЧС (колонка A / батальйон). */
export const isNovaRosterRow = (row: EjournalPreviewRow) => {
  const battalion = (
    getRosterValue(row, ["батальйон"]) ||
    String(row.column_1 ?? "").trim() ||
    readRosterColumnValue(row, 1) ||
    String(row.battalion ?? "").trim() ||
    String(row.A ?? "").trim()
  )
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase()
    .replace(/ё/g, "е");
  return battalion === "нова";
};

export const extractBirthDateFromPersonName = (name: string) => {
  const text = String(name ?? "");
  const match = text.match(
    /(\d{1,2}[.\/-]\d{1,2}[.\/-]\d{2,4})/,
  );
  if (!match || !looksLikePersonBirthDate(match[1])) return "";
  return normalizePersonBirthKey(match[1]);
};

const compactRnokpp = (value: unknown) => {
  const digits = String(value ?? "").replace(/\D/g, "");
  return digits.length >= 8 ? digits : "";
};

export const getRosterPersonBirthDate = (row: EjournalPreviewRow) => {
  const fromColumn =
    getRosterValue(row, ["дата", "народ"]) ||
    readRosterColumnValue(row, 16);
  if (fromColumn && looksLikePersonBirthDate(fromColumn)) {
    return normalizePersonBirthKey(fromColumn);
  }
  const rawName =
    getRosterValue(row, ["піб"]) ||
    getRosterValue(row, ["прізвище"]) ||
    readRosterColumnValue(row, 14);
  return extractBirthDateFromPersonName(String(rawName));
};

export const getRosterPersonRnokpp = (row: EjournalPreviewRow) =>
  compactRnokpp(
    getRosterValue(row, ["рнокпп"]) ||
      getRosterValue(row, ["іпн"]) ||
      readRosterColumnValue(row, 19),
  );

const pushIndexed = (
  map: Map<string, EjournalPreviewRow[]>,
  key: string,
  row: EjournalPreviewRow,
) => {
  if (!key) return;
  const list = map.get(key);
  if (list) list.push(row);
  else map.set(key, [row]);
};

const pickBestRow = (
  rows: EjournalPreviewRow[] | undefined,
  preferredBirth = "",
) => {
  if (!rows?.length) return undefined;
  if (rows.length === 1) return rows[0];
  if (preferredBirth) {
    const matching = rows.filter(
      (row) => getRosterPersonBirthDate(row) === preferredBirth,
    );
    if (matching.length) return matching[0];
  }
  const births = new Set(rows.map((row) => getRosterPersonBirthDate(row)));
  const rnokpps = new Set(
    rows.map((row) => getRosterPersonRnokpp(row)).filter(Boolean),
  );
  // Кілька однакових рядків штатки (той самий ІПН / дата) — беремо перший.
  if (births.size <= 1 && rnokpps.size <= 1) return rows[0];
  return undefined;
};

const birthsCompatible = (left: string, right: string) =>
  !left || !right || left === right;

const getRosterAdditions = (rosterRow: EjournalPreviewRow) =>
  Object.fromEntries(
    Object.entries(rosterRow)
      .filter(
        ([key, value]) =>
          !key.startsWith("__") &&
          valueToDisplay(value as Parameters<typeof valueToDisplay>[0]).trim(),
      )
      .map(([key, value]) => [`${ROSTER_FIELD_PREFIX}${key}`, value]),
  );

/** Прибрати попереднє збагачення зі Штатки, щоб новий імпорт реально оновлював картки. */
export const stripRosterEnrichment = (row: EjournalPreviewRow): EjournalPreviewRow =>
  Object.fromEntries(
    Object.entries(row).filter(
      ([key]) =>
        !key.startsWith(ROSTER_FIELD_PREFIX) && !key.startsWith("fighter_status_"),
    ),
  ) as EjournalPreviewRow;

export const buildRosterOnlyPersonnelRow = (rosterRow: EjournalPreviewRow) => {
  const name = getRosterPersonName(rosterRow);
  const identityKey = resolvePersonIdentityKey({
    ...rosterRow,
    прізвище: name,
    ПІБ: name,
  });
  const rowKey = identityKey || normalizeRosterText(name);
  const archive = isPersonnelFromArchive(rosterRow);

  return withArchiveMarker(
    {
      __dbRowId: archive ? `roster:archive:${rowKey}` : `roster:${rowKey}`,
      __rowNumber: rosterRow.__rowNumber,
      id: identityKey,
      прізвище: name,
      ПІБ: name,
      звання: resolvePersonRankTitle(rosterRow),
      Звання: resolvePersonRankTitle(rosterRow),
      позивний: getRosterValue(rosterRow, ["позив"]),
      Позивний: getRosterValue(rosterRow, ["позив"]),
      індекс_посади: getRosterValue(rosterRow, ["індекс", "посади"]),
      "Індекс посади": getRosterValue(rosterRow, ["індекс", "посади"]),
      місце_дислокації: getRosterValue(rosterRow, ["перебування"]),
      "Місце дислокації": getRosterValue(rosterRow, ["перебування"]),
      ...getRosterAdditions(rosterRow),
    } as EjournalPreviewRow,
    rosterRow,
  );
};

/** Як на сторінці Персонал: ООС + рядки лише з «Загального списку». */
export const mergeRosterRowsIntoPreview = (
  preview: Pick<DbPreviewState, "rows">,
  rosterRows: EjournalPreviewRow[],
): EjournalPreviewRow[] => {
  if (!rosterRows.length) return preview.rows;

  const rosterById = new Map<string, EjournalPreviewRow[]>();
  const rosterByRnokpp = new Map<string, EjournalPreviewRow[]>();
  const rosterByNameBirth = new Map<string, EjournalPreviewRow[]>();
  const rosterByName = new Map<string, EjournalPreviewRow[]>();
  const usedRosterRows = new Set<EjournalPreviewRow>();
  rosterRows.forEach((row) => {
    const id = getPersonExternalId(row);
    const name = getRosterPersonName(row);
    const nameKey = normalizeRosterText(name);
    const birth = getRosterPersonBirthDate(row);
    const rnokpp = getRosterPersonRnokpp(row);
    if (id) pushIndexed(rosterById, id, row);
    if (rnokpp) pushIndexed(rosterByRnokpp, rnokpp, row);
    if (nameKey && birth) pushIndexed(rosterByNameBirth, `${nameKey}|${birth}`, row);
    if (nameKey) pushIndexed(rosterByName, nameKey, row);
  });

  const pickRosterForPreviewRow = (base: EjournalPreviewRow) => {
    const spreadsheetId = getPersonExternalId(base);
    const birth = normalizePersonBirthKey(resolvePersonBirthDate(base));
    const nameKey = normalizeRosterText(getPersonDisplayName(base));
    const byId = pickBestRow(rosterById.get(spreadsheetId), birth);
    if (byId) return byId;

    const rnokpp = compactRnokpp(
      getPersonFieldValue(base, ["рнокпп_за_наявності"]) ||
        getPersonFieldValue(base, ["рнокпп"]),
    );
    const byRnokpp = pickBestRow(rosterByRnokpp.get(rnokpp), birth);
    if (byRnokpp) {
      const rosterNameKey = normalizeRosterText(getRosterPersonName(byRnokpp));
      // ІПН у старих картках може належати іншій людині. Не склеюємо
      // різні ПІБ лише через однаковий номер; порожній ПІБ усе ще можна
      // безпечно доповнити за ІПН.
      if (!nameKey || !rosterNameKey || nameKey === rosterNameKey) {
        return byRnokpp;
      }
    }

    if (nameKey && birth) {
      const byNameBirth = pickBestRow(
        rosterByNameBirth.get(`${nameKey}|${birth}`),
        birth,
      );
      if (byNameBirth) return byNameBirth;
    }

    const nameHits = rosterByName.get(nameKey) ?? [];
    const nameMatch = pickBestRow(nameHits, birth);
    if (
      nameMatch &&
      nameHits.length === 1 &&
      birthsCompatible(birth, getRosterPersonBirthDate(nameMatch))
    ) {
      return nameMatch;
    }
    return undefined;
  };

  const mergedRows = preview.rows.map((row) => {
    try {
      const base = stripRosterEnrichment(row);
      const rosterRow = pickRosterForPreviewRow(base);
      if (!rosterRow) return base;
      usedRosterRows.add(rosterRow);
      return withArchiveMarker(
        { ...base, ...getRosterAdditions(rosterRow) },
        rosterRow,
      );
    } catch {
      return stripRosterEnrichment(row);
    }
  });

  const usedNameKeys = new Set(
    [...usedRosterRows].map((row) => normalizeRosterText(getRosterPersonName(row))),
  );
  const usedRnokpp = new Set(
    [...usedRosterRows].map((row) => getRosterPersonRnokpp(row)).filter(Boolean),
  );

  const seenExtraKeys = new Set<string>();
  const rosterOnlyRows = rosterRows
    .filter((row) => !usedRosterRows.has(row))
    .filter((row) => getRosterPersonName(row))
    .filter((row) => {
      const nameKey = normalizeRosterText(getRosterPersonName(row));
      const rnokpp = getRosterPersonRnokpp(row);
      const birth = getRosterPersonBirthDate(row);
      if (rnokpp && usedRnokpp.has(rnokpp)) return false;
      if (usedNameKeys.has(nameKey)) {
        const usedWithName = [...usedRosterRows].filter(
          (item) => normalizeRosterText(getRosterPersonName(item)) === nameKey,
        );
        const usedBirths = usedWithName.map(getRosterPersonBirthDate);
        if (usedBirths.some((item) => birthsCompatible(item, birth))) {
          return false;
        }
      }
      const extraKey = rnokpp || `${nameKey}|${birth}`;
      if (seenExtraKeys.has(extraKey)) return false;
      seenExtraKeys.add(extraKey);
      if (rnokpp) seenExtraKeys.add(`${nameKey}|${birth}`);
      return true;
    })
    .flatMap((row) => {
      try {
        return [buildRosterOnlyPersonnelRow(row)];
      } catch {
        return [];
      }
    });

  return [...mergedRows, ...rosterOnlyRows];
};

export type PersonnelMergeRowPatch = {
  index: number;
  set: Record<string, unknown>;
  remove: string[];
};

export type PersonnelMergeDelta = {
  patches: PersonnelMergeRowPatch[];
  appended: EjournalPreviewRow[];
};

/** Compact worker result: only changed fields and roster-only additions. */
export const buildPersonnelMergeDelta = (
  preview: Pick<DbPreviewState, "rows">,
  rosterRows: EjournalPreviewRow[],
): PersonnelMergeDelta => {
  const merged = mergeRosterRowsIntoPreview(preview, rosterRows);
  const patches: PersonnelMergeRowPatch[] = [];

  for (let index = 0; index < preview.rows.length; index += 1) {
    const before = preview.rows[index];
    const after = merged[index];
    if (!after) continue;
    const set: Record<string, unknown> = {};
    const remove: string[] = [];
    for (const key of Object.keys(before)) {
      if (!(key in after)) remove.push(key);
    }
    for (const [key, value] of Object.entries(after)) {
      if (!(key in before) || !Object.is(before[key], value)) set[key] = value;
    }
    if (remove.length || Object.keys(set).length) {
      patches.push({ index, set, remove });
    }
  }

  return {
    patches,
    appended: merged.slice(preview.rows.length),
  };
};

export const applyPersonnelMergeDelta = (
  rows: EjournalPreviewRow[],
  delta: PersonnelMergeDelta | EjournalPreviewRow[],
): EjournalPreviewRow[] => {
  // A tab opened during a deployment can still receive the previous worker
  // response shape (the complete merged array). Accept it until that worker is
  // naturally replaced instead of breaking the page.
  if (Array.isArray(delta)) return delta;
  if (!delta || !Array.isArray(delta.patches) || !Array.isArray(delta.appended)) {
    return rows;
  }
  if (!delta.patches.length && !delta.appended.length) return rows;
  const patchByIndex = new Map(delta.patches.map((patch) => [patch.index, patch]));
  const merged = rows.map((row, index) => {
    const patch = patchByIndex.get(index);
    if (!patch) return row;
    const next = { ...row, ...patch.set };
    for (const key of patch.remove) delete next[key];
    return next;
  });
  return [...merged, ...delta.appended];
};

/** Доповнює рядки Штатки з БД кешем/Google-імпортом (без дублікатів за ПІБ / рядком). */
export const combineRosterRowSources = (
  primary: EjournalPreviewRow[],
  secondary: EjournalPreviewRow[],
): EjournalPreviewRow[] => {
  if (!secondary.length) return primary;

  const seenRows = new Set<number>();
  const seenNames = new Set<string>();
  const seenRnokpp = new Set<string>();

  const remember = (row: EjournalPreviewRow) => {
    const rowNum = Number(row.__rowNumber);
    if (Number.isFinite(rowNum) && rowNum > 0) seenRows.add(rowNum);
    const name = normalizeRosterText(getRosterPersonName(row));
    if (name) seenNames.add(name);
    const rnokpp = getRosterPersonRnokpp(row);
    if (rnokpp) seenRnokpp.add(rnokpp);
  };

  const merged = [...primary];
  primary.forEach(remember);

  for (const row of secondary) {
    const rowNum = Number(row.__rowNumber);
    const name = normalizeRosterText(getRosterPersonName(row));
    const rnokpp = getRosterPersonRnokpp(row);
    if (Number.isFinite(rowNum) && rowNum > 0 && seenRows.has(rowNum)) continue;
    if (name && seenNames.has(name)) continue;
    if (rnokpp && seenRnokpp.has(rnokpp)) continue;
    if (!name) continue;
    merged.push(row);
    remember(row);
  }

  return merged;
};

const ROSTER_ONLY_SHEET_STUB: BackendEjournalImportSheet = {
  id: "roster-only",
  batchId: "roster-only",
  name: "1.ОС Загальний список · Штатка",
  sheetIndex: 0,
  columnCount: 0,
  rowCount: 0,
  createdAt: "",
  updatedAt: "",
};

/** Картки лише зі Штатки, коли в БД ще немає імпорту ЕЖООС. */
export const buildRosterOnlyPreviewState = (
  rosterRows: EjournalPreviewRow[],
  sheet: BackendEjournalImportSheet | null | undefined = null,
): DbPreviewState | null => {
  if (!rosterRows.length) return null;
  const rows = mergeRosterRowsIntoPreview({ rows: [] }, rosterRows);
  if (!rows.length) return null;
  const activeSheet = sheet ?? {
    ...ROSTER_ONLY_SHEET_STUB,
    rowCount: rows.length,
  };
  return {
    sheet: activeSheet,
    columns: [],
    rows,
    total: rows.length,
    offset: 0,
    limit: rows.length,
  };
};
