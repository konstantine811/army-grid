import type { BackendPersonnelOverviewRow } from "../../api";
import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import {
  collectPersonAttachmentLookupIds,
  parseOrphanAttachmentIdentityId,
  questionnaireFileMatchesPerson,
} from "../personnel/personAttachments";
import {
  cleanPersonDisplayName,
  getPersonDisplayName,
  getPersonExternalId,
  resolvePersonBirthDate,
  resolvePersonIdentityKey,
} from "../personnel/personnelUtils";
import { normalizeOverviewName } from "./overviewNameSearch";
import {
  findOverviewRosterRow,
  indexOverviewRosterRows,
} from "./overviewPhotos";

const birthFromDisplayName = (name: string) => {
  const match = String(name ?? "").match(
    /(\d{1,2}[.\/-]\d{1,2}[.\/-]\d{2,4})\s*(?:р\.?\s*н\.?)?/i,
  );
  return match?.[1] ?? "";
};

const toPersonnelLikeRow = (
  row: BackendPersonnelOverviewRow,
  roster?: EjournalPreviewRow,
): EjournalPreviewRow => {
  const name = cleanPersonDisplayName(row.name) || row.name;
  const birth = birthFromDisplayName(row.name);
  const rosterId = roster ? getPersonExternalId(roster) : "";
  return {
    ...(roster ?? {}),
    __dbRowId: String(roster?.__dbRowId || row.id || ""),
    id: rosterId || row.externalId,
    externalId: rosterId || row.externalId,
    ПІБ: name,
    піб: name,
    прізвище: (roster && getPersonDisplayName(roster)) || name,
    ...(birth && !(roster && resolvePersonBirthDate(roster))
      ? { дата_народження: birth }
      : {}),
  } as EjournalPreviewRow;
};

const asQuestionnaireList = (
  items: Array<{ personExternalId?: string | null; fileName?: string | null }>,
) => (Array.isArray(items) ? items : []);

const uniqueNameSource = (
  items: Array<{ personExternalId?: string | null; fileName?: string | null }>,
  row: BackendPersonnelOverviewRow,
) => {
  const personKey = normalizeOverviewName(row.name);
  const hits = items.filter((item) => {
    const id = String(item.personExternalId ?? "").trim();
    const parsed = parseOrphanAttachmentIdentityId(id);
    if (parsed?.nameKey && normalizeOverviewName(parsed.nameKey) === personKey) {
      return true;
    }
    const fileName = String(item.fileName ?? "").trim();
    return (
      Boolean(fileName) &&
      !/^questionnaire\.pdf$/i.test(fileName) &&
      questionnaireFileMatchesPerson(fileName, [row.name])
    );
  });
  const ids = [
    ...new Set(
      hits
        .map((item) => String(item.personExternalId ?? "").trim())
        .filter(Boolean),
    ),
  ];
  return ids.length === 1 ? ids[0] : "";
};

/** Same matching as Особовий склад: spreadsheet ID, fingerprints, roster keys, file name. */
export const buildOverviewQuestionnairePresence = (
  rows: BackendPersonnelOverviewRow[],
  items: Array<{ personExternalId?: string | null; fileName?: string | null }>,
  rosterRows: EjournalPreviewRow[] = [],
) => {
  const presence: Record<string, true> = {};
  const sourceIds: Record<string, string> = {};
  const list = asQuestionnaireList(items);
  const storedSet = new Set(
    list
      .map((item) => String(item.personExternalId ?? "").trim())
      .filter(Boolean),
  );
  for (const id of storedSet) {
    presence[id] = true;
    sourceIds[id] = id;
  }

  let rosterById = new Map<string, EjournalPreviewRow>();
  let rosterByName = new Map<string, EjournalPreviewRow>();
  try {
    ({ rosterById, rosterByName } = indexOverviewRosterRows(rosterRows));
  } catch {
    /* keep empty indexes — still match by overview keys */
  }

  for (const row of rows) {
    const key = String(row.externalId ?? "").trim();
    if (!key) continue;
    try {
      const personRow = toPersonnelLikeRow(
        row,
        findOverviewRosterRow(row, rosterById, rosterByName),
      );
      const identity = resolvePersonIdentityKey(personRow);
      const lookup = [
        key,
        row.id,
        identity,
        ...collectPersonAttachmentLookupIds(personRow, undefined, {
          includeLooseKeys: true,
        }),
      ].filter(Boolean);
      const source =
        lookup.find((id) => storedSet.has(id)) || uniqueNameSource(list, row);
      if (!source) continue;
      presence[key] = true;
      sourceIds[key] = source;
    } catch {
      /* one bad row must not wipe the whole column */
    }
  }

  return { presence, sourceIds };
};
