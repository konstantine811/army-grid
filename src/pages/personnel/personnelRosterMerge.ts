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

export const ROSTER_FIELD_PREFIX = "roster__";
const normalizeRosterText = normalizeRosterMatchText;

/** Особа зі штатки / ранкового «Загального списку», не лише з ЕЖООС. */
export const isPersonnelInStaffRoster = (
  row: EjournalPreviewRow | null | undefined,
) => {
  if (!row) return false;
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

  return {
    __dbRowId: `roster:${rowKey}`,
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
  } as EjournalPreviewRow;
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
    const byId = pickBestRow(rosterById.get(spreadsheetId), birth);
    if (byId) return byId;

    const rnokpp = compactRnokpp(
      getPersonFieldValue(base, ["рнокпп_за_наявності"]) ||
        getPersonFieldValue(base, ["рнокпп"]),
    );
    const byRnokpp = pickBestRow(rosterByRnokpp.get(rnokpp), birth);
    if (byRnokpp) return byRnokpp;

    const nameKey = normalizeRosterText(getPersonDisplayName(base));
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
      return { ...base, ...getRosterAdditions(rosterRow) };
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
