import type {
  BackendEjournalImportSheet,
  BackendPersonnelRosterLatest,
} from "../api";
import { parseDbColumns } from "../pages/ejournal/ejournalUtils";
import type {
  DbPreviewState,
  EjournalColumn,
  EjournalPreviewRow,
} from "../pages/ejournal/ejournalTypes";
import {
  applyPersonnelMergeDelta,
  buildRosterOnlyPreviewState,
} from "../pages/personnel/personnelRosterMerge";
import {
  findEjournalPersonnelSheet,
  loadAllEjournalSheetRows,
  resolveMorningGeneralListColumnLabel,
} from "../pages/personnel/personnelUtils";
import { fillDownRosterUnitRows } from "../pages/overview/overviewRosterMerge";
import { runHeavyJob } from "../workers/runHeavyJob";
import { loadSharedEjournalImports, loadSharedRosterLatest } from "./sharedAppData";
import {
  CacheKeys,
  deleteDataCache,
  evictDataCacheMemory,
  peekDataCache,
  writeMemoryDataCache,
} from "./idbDataCache";

export type PersonnelDatasetVersion = {
  oosSheetId: string;
  oosStamp: string;
  rosterImportId: string;
  rosterSheetUpdatedAt: string;
  rosterRowCount: number;
};

export type PersonnelDataset = {
  rows: EjournalPreviewRow[];
  sheet: BackendEjournalImportSheet | null;
  columns: EjournalColumn[];
  total: number;
  rosterRows: EjournalPreviewRow[];
  rosterLabels: Record<string, string>;
  rosterColumns: EjournalColumn[];
  rosterUpdatedAt: string | null;
  version: PersonnelDatasetVersion;
  fingerprint: string;
  mergedAt: number;
  complete: boolean;
};

type LoadPersonnelDatasetOptions = {
  force?: boolean;
  signal?: AbortSignal;
  onCached?: (dataset: PersonnelDataset) => void | Promise<void>;
};

let inFlight: Promise<PersonnelDataset> | null = null;
let inFlightSignal: AbortSignal | undefined;
let removedPersistedV1 = false;

const sheetStamp = (sheet: BackendEjournalImportSheet | null | undefined) =>
  sheet
    ? `${sheet.updatedAt ?? ""}|${sheet.rowCount}|${sheet.columnCount}`
    : "";

export const buildPersonnelDatasetVersion = (
  sheet: BackendEjournalImportSheet | null | undefined,
  roster: BackendPersonnelRosterLatest | null | undefined,
): PersonnelDatasetVersion => ({
  oosSheetId: sheet?.id ?? "",
  oosStamp: sheetStamp(sheet),
  rosterImportId: roster?.importId ?? "",
  rosterSheetUpdatedAt: roster?.sheet?.updatedAt ?? "",
  rosterRowCount: roster?.sheet?.rowCount ?? roster?.rows?.length ?? 0,
});

export const personnelDatasetFingerprint = (
  version: PersonnelDatasetVersion,
) =>
  [
    version.oosSheetId,
    version.oosStamp,
    version.rosterImportId,
    version.rosterSheetUpdatedAt,
    version.rosterRowCount,
  ].join("\u0000");

export const rosterRowsFromPersonnelLatest = (
  latest: BackendPersonnelRosterLatest | null | undefined,
) => {
  if (!latest?.sheet || !Array.isArray(latest.rows)) {
    return [] as EjournalPreviewRow[];
  }
  return fillDownRosterUnitRows(
    latest.rows.map((row, rosterOrder) => ({
      __dbRowId: row.id,
      __rowNumber: row.excelRowNumber,
      __rosterOrder: rosterOrder,
      ...(row.values &&
      typeof row.values === "object" &&
      !Array.isArray(row.values)
        ? row.values
        : {}),
    })) as EjournalPreviewRow[],
  );
};

export const sortPersonnelRowsByRosterOrder = (
  rows: EjournalPreviewRow[],
): EjournalPreviewRow[] =>
  rows
    .map((row, sourceOrder) => ({ row, sourceOrder }))
    .sort((left, right) => {
      const leftOrder = Number(left.row.__rosterOrder);
      const rightOrder = Number(right.row.__rosterOrder);
      const leftInRoster = Number.isFinite(leftOrder);
      const rightInRoster = Number.isFinite(rightOrder);
      if (leftInRoster && rightInRoster) {
        return leftOrder - rightOrder || left.sourceOrder - right.sourceOrder;
      }
      if (leftInRoster) return -1;
      if (rightInRoster) return 1;
      return left.sourceOrder - right.sourceOrder;
    })
    .map(({ row }) => row);

const rosterMetadata = (
  latest: BackendPersonnelRosterLatest | null | undefined,
) => {
  const columns = parseDbColumns(latest?.sheet?.columns);
  return {
    rosterColumns: columns,
    rosterLabels: Object.fromEntries(
      columns.map((column) => [
        column.key,
        column.label?.trim() ||
          resolveMorningGeneralListColumnLabel(column.key) ||
          column.key,
      ]),
    ),
    rosterUpdatedAt:
      latest?.sheet?.updatedAt ?? latest?.createdAt ?? null,
  };
};

const buildDataset = async (
  preview: DbPreviewState | null,
  roster: BackendPersonnelRosterLatest | null,
  version: PersonnelDatasetVersion,
): Promise<PersonnelDataset> => {
  const rosterRows = rosterRowsFromPersonnelLatest(roster);
  const fallback = buildRosterOnlyPreviewState(rosterRows, roster?.sheet);
  const base = preview ?? fallback;
  const mergedRows =
    base && rosterRows.length
      ? applyPersonnelMergeDelta(
          base.rows,
          await runHeavyJob({
            type: "mergePersonnel",
            preview: base,
            rosterRows,
          }),
        )
      : base?.rows ?? [];
  const rows = sortPersonnelRowsByRosterOrder(mergedRows);
  const metadata = rosterMetadata(roster);
  return {
    rows,
    sheet: base?.sheet ?? roster?.sheet ?? null,
    columns: base?.columns ?? [],
    total: rows.length,
    rosterRows,
    ...metadata,
    version,
    fingerprint: personnelDatasetFingerprint(version),
    mergedAt: Date.now(),
    complete: true,
  };
};

const buildRosterBootstrap = (
  roster: BackendPersonnelRosterLatest,
  version: PersonnelDatasetVersion,
): PersonnelDataset | null => {
  const rosterRows = rosterRowsFromPersonnelLatest(roster);
  const preview = buildRosterOnlyPreviewState(rosterRows, roster.sheet);
  if (!preview) return null;
  return {
    rows: preview.rows,
    sheet: preview.sheet,
    columns: preview.columns,
    total: preview.rows.length,
    rosterRows,
    ...rosterMetadata(roster),
    version,
    fingerprint: personnelDatasetFingerprint(version),
    mergedAt: Date.now(),
    complete: false,
  };
};

export const personnelDatasetToPreview = (
  dataset: PersonnelDataset,
): DbPreviewState | null =>
  dataset.sheet
    ? {
        sheet: dataset.sheet,
        columns: dataset.columns,
        rows: dataset.rows,
        total: dataset.total,
        offset: 0,
        limit: dataset.rows.length,
      }
    : null;

export const loadPersonnelDataset = async (
  options: LoadPersonnelDatasetOptions = {},
): Promise<PersonnelDataset> => {
  if (!removedPersistedV1) {
    removedPersistedV1 = true;
    void deleteDataCache("personnel:dataset:v1");
  }
  if (inFlight && !options.force && !inFlightSignal?.aborted) return inFlight;
  if (inFlightSignal?.aborted) {
    inFlight = null;
    inFlightSignal = undefined;
  }

  const run = async () => {
    const cached = peekDataCache<PersonnelDataset>(
      CacheKeys.personnelDataset,
    );
    if (cached) await options.onCached?.(cached);

    const [imports, roster] = await Promise.all([
      loadSharedEjournalImports({
        force: options.force,
        signal: options.signal,
      }),
      loadSharedRosterLatest({
        force: options.force,
        signal: options.signal,
      }),
    ]);
    if (options.signal?.aborted) throw new DOMException("Aborted", "AbortError");

    const sheet = findEjournalPersonnelSheet(imports);
    const version = buildPersonnelDatasetVersion(sheet, roster);
    const fingerprint = personnelDatasetFingerprint(version);
    if (!options.force && cached?.fingerprint === fingerprint) return cached;
    if (!cached && roster) {
      const bootstrap = buildRosterBootstrap(roster, version);
      if (bootstrap) await options.onCached?.(bootstrap);
    }

    const preview = sheet
      ? await loadAllEjournalSheetRows(sheet, {
          force: options.force,
          signal: options.signal,
        })
      : null;
    if (options.signal?.aborted) throw new DOMException("Aborted", "AbortError");

    const dataset = await buildDataset(preview, roster, version);
    if (options.signal?.aborted) throw new DOMException("Aborted", "AbortError");
    writeMemoryDataCache(CacheKeys.personnelDataset, dataset);
    // The merged dataset supersedes the wide raw OOS snapshot in RAM. Its
    // persistent IndexedDB copy remains available for a cold reload.
    evictDataCacheMemory("ejournal:sheet-rows:");
    return dataset;
  };

  const promise = run();
  if (!options.force) {
    inFlight = promise;
    inFlightSignal = options.signal;
  }
  try {
    return await promise;
  } finally {
    if (inFlight === promise) {
      inFlight = null;
      inFlightSignal = undefined;
    }
  }
};
