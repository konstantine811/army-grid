import { api, type BackendEjournalImport } from "../../api";
import {
  CacheKeys,
  fetchWithCache,
  jsonChanged,
  readDataCache,
  writeDataCache,
} from "../../data/idbDataCache";
import {
  mapRosterLatestToPreviewRows,
  readRosterColumnValue,
} from "../excel-fill/rosterSourceSnapshot";
import type { EjournalColumn, EjournalPreviewRow } from "../ejournal/ejournalTypes";
import { parseDbColumns } from "../ejournal/ejournalUtils";
import { mergePhonesAppendOnly } from "../personnel/personEnrichment";
import {
  upsertPersonPhonesDocument,
  readStoredPersonPhones,
  writeStoredPersonPhones,
} from "../personnel/personPhonesStore";
import {
  getRosterValue,
  mergeRosterRowsIntoPreview,
} from "../personnel/personnelRosterMerge";
import {
  buildPersonIdentityFingerprint,
  cleanPersonDisplayName,
  extractPersonCallSign,
  extractPhones,
  findEjournalPersonnelSheet,
  getPersonDisplayName,
  getPersonExternalId,
  getPersonFieldValue,
  isLikelyPersonnelRow,
  loadAllEjournalSheetRows,
  looksLikePersonnelName,
  normalizePersonBirthKey,
  resolveMorningGeneralListColumnLabel,
  resolvePersonIdentityKey,
} from "../personnel/personnelUtils";
import {
  normalizeAnketaExternalIdKey,
  normalizeAnketaNameKey,
} from "./anketaPersonMatch";
import type { AnketaRow } from "./anketaSheet";

export const DEFAULT_ANKETA_PERSONNEL_ROSTER_COLUMNS: EjournalColumn[] = [
  { key: "column_13", label: "Звання", order: 0, originalIndex: 12 },
  { key: "column_14", label: "ПІБ", order: 1, originalIndex: 13 },
  { key: "column_15", label: "Позивний", order: 2, originalIndex: 14 },
  { key: "column_16", label: "Дата народження", order: 3, originalIndex: 15 },
  { key: "column_19", label: "ІПН", order: 4, originalIndex: 18 },
  { key: "column_11", label: "Військовий квиток", order: 5, originalIndex: 10 },
  { key: "column_31", label: "Місце перебування", order: 6, originalIndex: 30 },
  { key: "column_12", label: "Мобілізація/контракт", order: 7, originalIndex: 11 },
  { key: "id", label: "ID", order: 8, originalIndex: 45 },
];

type MappedAnketaRosterFields = {
  fullName: string;
  rank: string;
  externalId: string;
  callsign: string;
  birthDate: string;
  rnokpp: string;
  militaryId: string;
  location: string;
  positionIndex: string;
  serviceType: string;
};

const compactRnokpp = (value: unknown) => {
  const digits = String(value ?? "").replace(/\D/g, "");
  return digits.length >= 8 ? digits : "";
};

const personnelNameFromAnketa = (value: unknown) =>
  cleanPersonDisplayName(String(value ?? ""))
    .replace(/\([^)]*\)/g, " ")
    .replace(/\s+/g, " ")
    .trim();

export const mapAnketaRowToRosterFields = (
  row: AnketaRow,
): MappedAnketaRosterFields => {
  const fullName = personnelNameFromAnketa(row.fullName);
  return {
    fullName,
    rank: String(row.rank ?? "").trim(),
    externalId: String(row.externalId ?? "").trim(),
    callsign: extractPersonCallSign(row.fullName, row.additionalInfo),
    birthDate: String(row.birthDate ?? "").trim(),
    rnokpp: compactRnokpp(row.rnokpp) || String(row.rnokpp ?? "").trim(),
    militaryId: String(row.militaryId ?? "").trim(),
    location: String(row.location ?? "").trim(),
    positionIndex: String(row.positionIndex ?? "")
      .trim()
      .replace(/\s*[·,;]\s*|\s+/g, "\n")
      .replace(/\n+/g, "\n")
      .trim(),
    serviceType: String(row.serviceType ?? "").trim(),
  };
};

const columnHint = (column: { key: string; label?: string }) => {
  const label =
    column.label?.trim() ||
    resolveMorningGeneralListColumnLabel(column.key) ||
    "";
  return `${column.key.replace(/_/g, " ")} ${label}`
    .toLocaleLowerCase("uk-UA")
    .replace(/[ʼ’']/g, "");
};

const valueForColumnHint = (
  hint: string,
  fields: MappedAnketaRosterFields,
): string => {
  if (/позив/.test(hint)) return fields.callsign;
  if (/(піб|прізвище)/.test(hint)) return fields.fullName;
  if (/званн/.test(hint)) return fields.rank;
  if (
    /(^| )id( |$)|ідентифікатор/.test(hint) &&
    !/посад/.test(hint) &&
    !/рядк/.test(hint)
  ) {
    return fields.externalId;
  }
  if (/народж/.test(hint)) return fields.birthDate;
  if (/іпн|рнокпп|інн/.test(hint)) return fields.rnokpp;
  if (/квитк/.test(hint)) return fields.militaryId;
  if (/перебуван|дислокац/.test(hint)) return fields.location;
  if (/індекс/.test(hint) && /посад/.test(hint)) return fields.positionIndex;
  if (
    /(мобіліз|контракт|вид.{0,16}служб)/.test(hint) &&
    !/дат/.test(hint) &&
    !/наказ/.test(hint)
  ) {
    return fields.serviceType;
  }
  if (/(^| )анкета( |$)/.test(hint) || /^column 10\b/.test(hint)) return "так";
  return "";
};

const assignIf = (
  values: Record<string, unknown>,
  key: string,
  value: string,
) => {
  if (value) values[key] = value;
};

/** Значення рядка «Загального списку» з анкети — і за колонками Штатки, і за ключами картки. */
export const buildPersonnelRosterValuesFromAnketa = (
  row: AnketaRow,
  columns: EjournalColumn[] = DEFAULT_ANKETA_PERSONNEL_ROSTER_COLUMNS,
): Record<string, unknown> => {
  const fields = mapAnketaRowToRosterFields(row);
  const values: Record<string, unknown> = {};

  for (const column of columns) {
    const next = valueForColumnHint(columnHint(column), fields);
    if (next) values[column.key] = next;
  }

  assignIf(values, "ПІБ", fields.fullName);
  assignIf(values, "піб", fields.fullName);
  assignIf(values, "прізвище", fields.fullName);
  assignIf(values, "звання", fields.rank);
  assignIf(values, "Звання", fields.rank);
  assignIf(values, "id", fields.externalId);
  assignIf(values, "ID", fields.externalId);
  assignIf(values, "позивний", fields.callsign);
  assignIf(values, "Позивний", fields.callsign);
  assignIf(values, "дата_народження", fields.birthDate);
  assignIf(values, "Дата народження", fields.birthDate);
  assignIf(values, "рнокпп_за_наявності", fields.rnokpp);
  assignIf(values, "ІПН", fields.rnokpp);
  assignIf(values, "військовий_квиток", fields.militaryId);
  assignIf(values, "місце_перебування", fields.location);
  assignIf(values, "Місце перебування", fields.location);
  assignIf(values, "місце_дислокації", fields.location);
  assignIf(values, "індекс_посади", fields.positionIndex);
  assignIf(values, "вид_служби", fields.serviceType);
  assignIf(values, "звідки_прибув", String(row.arrivedFrom ?? "").trim());
  assignIf(values, "місце_народження", String(row.birthPlace ?? "").trim());
  assignIf(values, "стать", String(row.sex ?? "").trim());
  assignIf(values, "додаткова_інформація", String(row.additionalInfo ?? "").trim());

  return values;
};

/** Картка особового складу з анкети — одразу видно в списку, навіть без Штатки. */
export const buildPersonnelRowFromAnketa = (
  row: AnketaRow,
): EjournalPreviewRow => {
  const fields = mapAnketaRowToRosterFields(row);
  const identity =
    fields.externalId ||
    buildPersonIdentityFingerprint(
      fields.fullName,
      fields.birthDate,
      fields.callsign,
    ) ||
    normalizeAnketaNameKey(fields.fullName);
  const values = buildPersonnelRosterValuesFromAnketa(row);

  return {
    ...values,
    __dbRowId: `anketa:${identity}`,
    __rowNumber: row.__rowNumber,
    прізвище: fields.fullName,
    ПІБ: fields.fullName,
    звання: fields.rank,
    Звання: fields.rank,
    id: fields.externalId || identity,
    позивний: fields.callsign,
    Позивний: fields.callsign,
    дата_народження: fields.birthDate,
    рнокпп_за_наявності: fields.rnokpp,
    місце_дислокації: fields.location,
    місце_перебування: fields.location,
    індекс_посади: fields.positionIndex,
    вид_служби: fields.serviceType,
  } as EjournalPreviewRow;
};

export const isAnketaRowCreatableInPersonnel = (row: AnketaRow) =>
  looksLikePersonnelName(personnelNameFromAnketa(row.fullName));

const personnelRowNameKey = (row: EjournalPreviewRow) =>
  normalizeAnketaNameKey(getPersonDisplayName(row));

const personnelRowBirthKey = (row: EjournalPreviewRow) =>
  normalizePersonBirthKey(
    getPersonFieldValue(row, ["дата_народження"]) ||
      getRosterValue(row, ["народжен"]),
  );

const personnelRowIdentityMarks = (row: EjournalPreviewRow) => {
  const marks = new Set<string>();
  const id =
    getPersonExternalId(row) ||
    resolvePersonIdentityKey(row) ||
    "";
  if (id && !id.startsWith("anketa:")) marks.add(`id:${normalizeAnketaExternalIdKey(id)}`);
  if (id.startsWith("anketa:")) marks.add(`anketa:${id}`);
  const name = personnelRowNameKey(row);
  const birth = personnelRowBirthKey(row);
  if (name && birth) marks.add(`nb:${name}|${birth}`);
  if (name) marks.add(`name:${name}`);
  return marks;
};

/** Додати картки з анкет у список, не дублюючи тих, хто вже є. */
export const mergeAnketaCreatedRowsIntoPreview = (
  rows: EjournalPreviewRow[],
  created: EjournalPreviewRow[],
): EjournalPreviewRow[] => {
  if (!created.length) return rows;

  const used = new Set<string>();
  for (const row of rows) {
    for (const mark of personnelRowIdentityMarks(row)) used.add(mark);
  }

  const extras: EjournalPreviewRow[] = [];
  for (const row of created) {
    if (!isLikelyPersonnelRow(row) && !looksLikePersonnelName(getPersonDisplayName(row))) {
      continue;
    }
    const marks = [...personnelRowIdentityMarks(row)];
    if (marks.some((mark) => used.has(mark))) continue;
    extras.push(row);
    for (const mark of marks) used.add(mark);
  }

  return extras.length ? [...rows, ...extras] : rows;
};

export const loadAnketaCreatedPersonnel = async () => {
  const stored = await readDataCache<EjournalPreviewRow[]>(
    CacheKeys.anketaCreatedPersonnel,
  );
  return Array.isArray(stored) ? stored : [];
};

export const saveAnketaCreatedPersonnel = async (rows: EjournalPreviewRow[]) => {
  await writeDataCache(CacheKeys.anketaCreatedPersonnel, rows);
};

/** Те, що зараз видно в Особовому складі: останній ООС + Штатка + уже додані з анкет. */
export const loadVisiblePersonnelRows = async (): Promise<EjournalPreviewRow[]> => {
  const [imports, latestRoster, created] = await Promise.all([
    fetchWithCache<BackendEjournalImport[]>({
      key: CacheKeys.ejournalImports,
      fetcher: () => api.listEjournalImports(),
      isChanged: jsonChanged,
    }).catch(() => [] as BackendEjournalImport[]),
    fetchWithCache({
      key: CacheKeys.rosterLatest,
      fetcher: () => api.getLatestPersonnelRoster(),
      isChanged: jsonChanged,
    }).catch(() => null),
    loadAnketaCreatedPersonnel(),
  ]);

  const sheet = findEjournalPersonnelSheet(imports);
  const preview = sheet
    ? await loadAllEjournalSheetRows(sheet).catch(() => null)
    : null;
  const withRoster = mergeRosterRowsIntoPreview(
    { rows: preview?.rows ?? [] },
    mapRosterLatestToPreviewRows(latestRoster),
  );
  return mergeAnketaCreatedRowsIntoPreview(withRoster, created);
};

const rosterPreviewName = (row: EjournalPreviewRow) =>
  cleanPersonDisplayName(
    readRosterColumnValue(row, 14) ||
      getRosterValue(row, ["піб"]) ||
      getRosterValue(row, ["прізвище"]) ||
      getPersonDisplayName(row),
  );

const rosterPreviewRnokpp = (row: EjournalPreviewRow) =>
  compactRnokpp(
    readRosterColumnValue(row, 19) ||
      getRosterValue(row, ["іпн"]) ||
      getRosterValue(row, ["рнокпп"]),
  );

const rosterPreviewBirth = (row: EjournalPreviewRow) =>
  readRosterColumnValue(row, 16) || getRosterValue(row, ["народжен"]);

/** Анкети без збігу в ООС/списку, з ПІБ, без дублікатів між собою і з уже наявним списком. */
export const selectAnketaRowsMissingFromPersonnel = (
  unmatched: AnketaRow[],
  existingRosterRows: EjournalPreviewRow[] = [],
): AnketaRow[] => {
  const existingIds = new Set<string>();
  const existingRnokpp = new Set<string>();
  const existingNameBirth = new Set<string>();
  const existingNames = new Set<string>();

  for (const row of existingRosterRows) {
    const id = normalizeAnketaExternalIdKey(getPersonExternalId(row));
    if (id) existingIds.add(id);
    const rnokpp = rosterPreviewRnokpp(row);
    if (rnokpp) existingRnokpp.add(rnokpp);
    const name = normalizeAnketaNameKey(rosterPreviewName(row));
    const birth = normalizePersonBirthKey(rosterPreviewBirth(row));
    if (name && birth) existingNameBirth.add(`${name}|${birth}`);
    if (name) existingNames.add(name);
  }

  const seenIds = new Set<string>();
  const seenRnokpp = new Set<string>();
  const seenNameBirth = new Set<string>();
  const seenNames = new Set<string>();
  const selected: AnketaRow[] = [];

  for (const row of unmatched) {
    if (!isAnketaRowCreatableInPersonnel(row)) continue;

    const id = normalizeAnketaExternalIdKey(row.externalId);
    const rnokpp = compactRnokpp(row.rnokpp);
    const name = normalizeAnketaNameKey(row.fullName);
    const birth = normalizePersonBirthKey(row.birthDate);

    if (id && (existingIds.has(id) || seenIds.has(id))) continue;
    if (rnokpp && (existingRnokpp.has(rnokpp) || seenRnokpp.has(rnokpp))) {
      continue;
    }
    if (name && birth) {
      const nameBirth = `${name}|${birth}`;
      if (existingNameBirth.has(nameBirth) || seenNameBirth.has(nameBirth)) {
        continue;
      }
    } else if (name && (existingNames.has(name) || seenNames.has(name))) {
      continue;
    }

    selected.push(row);
    if (id) seenIds.add(id);
    if (rnokpp) seenRnokpp.add(rnokpp);
    if (name && birth) seenNameBirth.add(`${name}|${birth}`);
    if (name) seenNames.add(name);
  }

  return selected;
};

const personPhoneKey = (row: AnketaRow) => {
  const spreadsheetId = String(row.externalId ?? "").trim();
  if (spreadsheetId) return spreadsheetId;
  return buildPersonIdentityFingerprint(
    personnelNameFromAnketa(row.fullName),
    String(row.birthDate ?? "").trim(),
    extractPersonCallSign(row.fullName, row.additionalInfo),
  );
};

const storeAnketaPhonesForCreated = async (rows: AnketaRow[]) => {
  const store = readStoredPersonPhones();
  let changed = false;

  for (const row of rows) {
    const personKey = personPhoneKey(row);
    const incoming = extractPhones(row.additionalInfo);
    if (!personKey || !incoming.length) continue;
    const before = store[personKey] ?? [];
    const merged = mergePhonesAppendOnly(before, incoming);
    if (merged.length === before.length) continue;
    store[personKey] = merged;
    changed = true;
    await upsertPersonPhonesDocument(personKey, merged).catch(() => undefined);
  }

  if (changed) writeStoredPersonPhones(store);
};

export type AppendAnketaPeopleToRosterResult = {
  created: number;
  skipped: number;
};

const asImportColumns = (columns: EjournalColumn[]) =>
  columns.map((column, index) => ({
    key: column.key,
    label: column.label,
    order: Number.isFinite(column.order) ? column.order : index,
    originalIndex: column.originalIndex,
    letter: column.letter,
  }));

const valuesFromLatestRow = (values: unknown): Record<string, unknown> =>
  values && typeof values === "object" && !Array.isArray(values)
    ? { ...(values as Record<string, unknown>) }
    : {};

const tryImportAnketaPeopleIntoRoster = async (toCreate: AnketaRow[]) => {
  const latest = await api.getLatestPersonnelRoster().catch(() => null);
  const parsedColumns = parseDbColumns(latest?.sheet?.columns);
  const columns = parsedColumns.length
    ? parsedColumns
    : DEFAULT_ANKETA_PERSONNEL_ROSTER_COLUMNS;

  const existingImportRows = (latest?.rows ?? []).map((row) => ({
    excelRowNumber: row.excelRowNumber ?? undefined,
    values: valuesFromLatestRow(row.values),
  }));
  const maxExcelRow = existingImportRows.reduce(
    (max, row) => Math.max(max, Number(row.excelRowNumber) || 0),
    0,
  );

  const newRows = toCreate.map((row, index) => ({
    excelRowNumber: maxExcelRow + index + 1,
    values: buildPersonnelRosterValuesFromAnketa(row, columns),
  }));

  await api.importPersonnelRoster({
    name: latest?.importName?.trim() || "З анкетних даних",
    sourceFileName: latest?.sourceFileName ?? undefined,
    notes: `Додано ${newRows.length} осіб з анкетних даних, яких не було в особовому складі.`,
    sheets: [
      {
        name: latest?.sheet?.name?.trim() || "Загальний список",
        sheetIndex: latest?.sheet?.sheetIndex ?? 0,
        columns: asImportColumns(columns),
        rows: [...existingImportRows, ...newRows],
      },
    ],
  });

  const fresh = await api.getLatestPersonnelRoster().catch(() => null);
  if (fresh) await writeDataCache(CacheKeys.rosterLatest, fresh);
};

/** Записати відсутніх з анкет у серверну Штатку — єдине спільне джерело. */
export const appendAnketaPeopleToPersonnelRoster = async (
  unmatched: AnketaRow[],
  alreadyVisible: EjournalPreviewRow[] = [],
): Promise<AppendAnketaPeopleToRosterResult> => {
  if (!unmatched.length) return { created: 0, skipped: 0 };

  const createdLocal = await loadAnketaCreatedPersonnel();
  const toCreate = selectAnketaRowsMissingFromPersonnel(unmatched, [
    ...alreadyVisible,
    ...createdLocal,
  ]);
  const skipped = unmatched.length - toCreate.length;
  if (!toCreate.length) return { created: 0, skipped };

  await tryImportAnketaPeopleIntoRoster(toCreate);
  await storeAnketaPhonesForCreated(toCreate);

  return { created: toCreate.length, skipped };
};
