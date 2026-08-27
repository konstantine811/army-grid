import { api } from "../../api";
import { hasRowData, readWorkbookSnapshot } from "../../excelRoundTrip";
import {
  CacheKeys,
  readDataCache,
  writeDataCache,
} from "../../data/idbDataCache";
import { cellValueToJson } from "../../shared/format";
import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import {
  buildImportColumns,
  localRowsToPreviewRows,
} from "../ejournal/ejournalUtils";
import {
  buildFighterStatusAdditions,
  FIGHTER_STATUS_FIELDS,
  findFighterStatusAddition,
  findFighterStatusSheet,
} from "../personnel/fighterStatusImport";
import { buildStaffSheetPreviewRows } from "./staffSheetPreview";

export type StaffSheetImportSnapshot = {
  fileName: string;
  importedAt: string;
  sheetName: string;
  rows: EjournalPreviewRow[];
  /** Бінарний .xlsx — для експорту «як Штатка». */
  fileData?: ArrayBuffer | null;
  source: "file";
  personCount: number;
};

const findRosterSheet = (
  sheets: Awaited<ReturnType<typeof readWorkbookSnapshot>>["sheets"],
) =>
  sheets.find((sheet) => /загальний\s*список/i.test(sheet.sheetName));

const sortRosterRows = (rows: EjournalPreviewRow[]) =>
  [...rows].sort(
    (left, right) =>
      (Number(left.__rowNumber) || 0) - (Number(right.__rowNumber) || 0),
  );

export const parseStaffSheetImportFile = async (
  file: File,
): Promise<StaffSheetImportSnapshot> => {
  const [snapshot, fileData] = await Promise.all([
    readWorkbookSnapshot(file),
    file.arrayBuffer(),
  ]);

  const rosterSheet = findRosterSheet(snapshot.sheets);
  if (!rosterSheet) {
    throw new Error(
      "У файлі не знайдено аркуш «Загальний список» (1.ОС Загальний список).",
    );
  }

  const columns = buildImportColumns(rosterSheet);
  const rows = sortRosterRows(localRowsToPreviewRows(rosterSheet.rows, columns));

  if (!rows.length) {
    throw new Error("Аркуш «Загальний список» порожній.");
  }

  return {
    fileName: snapshot.fileName || file.name,
    importedAt: new Date().toISOString(),
    sheetName: rosterSheet.sheetName,
    rows,
    fileData,
    source: "file",
    personCount: buildStaffSheetPreviewRows(rows).length,
  };
};

/**
 * Єдиний імпорт «Штатки»: .xlsx → кеш для Анкетних даних / ранкового звіту
 * і водночас у БД персоналу (Загальний список).
 */
export const importStaffSheetFromFile = async (
  file: File,
): Promise<StaffSheetImportSnapshot> => {
  const [snapshot, imported] = await Promise.all([
    readWorkbookSnapshot(file),
    parseStaffSheetImportFile(file),
  ]);

  const rosterSheet = findRosterSheet(snapshot.sheets);
  if (!rosterSheet) {
    throw new Error(
      "У файлі не знайдено аркуш «Загальний список» (1.ОС Загальний список).",
    );
  }

  const fighterStatusSheet = findFighterStatusSheet(snapshot.sheets);
  const fighterStatusAdditions = fighterStatusSheet
    ? buildFighterStatusAdditions(fighterStatusSheet)
    : new Map<string, Record<string, unknown>>();
  const rosterColumns = buildImportColumns(rosterSheet);
  const columns = [
    ...rosterColumns,
    ...FIGHTER_STATUS_FIELDS.map((field, index) => ({
      key: field.key,
      label: field.label,
      order: rosterSheet.columnCount + index,
      originalIndex: rosterSheet.columnCount + index,
      letter: "",
    })),
  ];

  const rows = rosterSheet.rows
    .filter((row) => hasRowData(row.values))
    .map((row) => {
      const values = Object.fromEntries(
        rosterColumns.map((column, index) => [
          column.key,
          cellValueToJson(row.values[index]),
        ]),
      );
      const statusAddition = findFighterStatusAddition(
        values,
        fighterStatusAdditions,
      );
      return {
        excelRowNumber: row.excelRowNumber,
        values: {
          ...values,
          ...(statusAddition ?? {}),
        },
      };
    });

  await api.importPersonnelRoster({
    name: imported.fileName.replace(/\.(xlsx|xlsm)$/i, ""),
    sourceFileName: imported.fileName,
    notes: "Імпорт «Штатки» (.xlsx) з Анкетних даних.",
    sheets: [
      {
        name: rosterSheet.sheetName,
        sheetIndex: rosterSheet.sheetIndex,
        columns,
        rows,
      },
    ],
  });

  const fresh = await api.getLatestPersonnelRoster();
  if (fresh) await writeDataCache(CacheKeys.rosterLatest, fresh);

  await saveStaffSheetImport(imported);
  return imported;
};

export const saveStaffSheetImport = async (
  snapshot: StaffSheetImportSnapshot,
) => {
  await writeDataCache(CacheKeys.staffSheetImport, snapshot);
};

export const loadStaffSheetImport = async () =>
  readDataCache<StaffSheetImportSnapshot>(CacheKeys.staffSheetImport);

export const clearStaffSheetImport = async () => {
  await writeDataCache(CacheKeys.staffSheetImport, null);
};

export const formatStaffSheetImportSummary = (
  imported: StaffSheetImportSnapshot,
) => {
  const preview = buildStaffSheetPreviewRows(imported.rows);
  const first = preview
    .slice(0, 3)
    .map((row) => row.cells[2] || "—")
    .join(" · ");
  return `${imported.personCount || preview.length} осіб · перші: ${first}`;
};
