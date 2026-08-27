import { valueToDisplay } from "../../excelRoundTrip";
import type { DbPreviewState, EjournalPreviewRow } from "../ejournal/ejournalTypes";
import { normalizeRosterMatchText } from "./fighterStatusImport";
import {
  cleanPersonDisplayName,
  getPersonDisplayName,
  getPersonExternalId,
  resolvePersonIdentityKey,
  resolvePersonRankTitle,
} from "./personnelUtils";

export const ROSTER_FIELD_PREFIX = "roster__";
const normalizeRosterText = normalizeRosterMatchText;

export const getRosterValue = (row: EjournalPreviewRow, keyParts: string[]) => {
  const key = Object.keys(row).find((item) =>
    keyParts.every((part) => item.toLocaleLowerCase("uk-UA").includes(part)),
  );
  return key
    ? valueToDisplay(row[key] as Parameters<typeof valueToDisplay>[0]).trim()
    : "";
};

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
  const rawName =
    getRosterValue(rosterRow, ["піб"]) ||
    getRosterValue(rosterRow, ["прізвище"]);
  const name = cleanPersonDisplayName(rawName);
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

  const rosterById = new Map<string, EjournalPreviewRow>();
  const rosterByName = new Map<string, EjournalPreviewRow>();
  const usedRosterRows = new Set<EjournalPreviewRow>();
  rosterRows.forEach((row) => {
    const id = getPersonExternalId(row);
    const name = getRosterValue(row, ["піб"]) || getRosterValue(row, ["прізвище"]);
    if (id) rosterById.set(id, row);
    if (name) rosterByName.set(normalizeRosterText(name), row);
  });

  const mergedRows = preview.rows.map((row) => {
    try {
      const base = stripRosterEnrichment(row);
      const spreadsheetId = getPersonExternalId(base);
      const name = getPersonDisplayName(base);
      const rosterRow =
        (spreadsheetId && rosterById.get(spreadsheetId)) ||
        rosterByName.get(normalizeRosterText(name));

      if (!rosterRow) return base;
      usedRosterRows.add(rosterRow);
      return { ...base, ...getRosterAdditions(rosterRow) };
    } catch {
      return stripRosterEnrichment(row);
    }
  });

  const rosterOnlyRows = rosterRows
    .filter((row) => !usedRosterRows.has(row))
    .filter((row) => getRosterValue(row, ["піб"]) || getRosterValue(row, ["прізвище"]))
    .flatMap((row) => {
      try {
        return [buildRosterOnlyPersonnelRow(row)];
      } catch {
        return [];
      }
    });

  return [...mergedRows, ...rosterOnlyRows];
};
