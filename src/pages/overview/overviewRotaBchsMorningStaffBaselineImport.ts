import { readWorkbookSnapshot } from "../../excelRoundTrip";
import { buildImportColumns } from "../ejournal/ejournalUtils";
import { parseStaffSheetImportFile } from "../anketa-data/staffSheetImport";
import type { BackendPersonnelOverviewRow } from "../../api";
import {
  buildStaffOverviewRowsFromRoster,
  fillDownRosterUnitRows,
} from "./overviewRosterMerge";
import { filterBchsMorningUnitRows } from "./overviewRotaBchsMorningExport";
import { buildBchsMorningDailySnapshot } from "./overviewRotaBchsMorningStatusChanges";
import type { BchsMorningDailySnapshot } from "./overviewRotaBchsMorningSnapshot";
import { kyivIsoDateLabel } from "./overviewRotaBchsMorningSnapshot";
import { parseBchsBaselineDateFromFileName } from "./overviewRotaBchsMorningBaselineImport";

const findRosterSheet = (
  sheets: Awaited<ReturnType<typeof readWorkbookSnapshot>>["sheets"],
) =>
  sheets.find((sheet) => /загальний\s*список/i.test(sheet.sheetName)) ??
  sheets.find((sheet) => /^sh$/i.test(sheet.sheetName.trim()));

const withStaffOverviewStatus = (
  row: BackendPersonnelOverviewRow,
): BackendPersonnelOverviewRow => {
  if (!row.staffStatus) return row;
  return {
    ...row,
    status: row.staffStatus,
    statusLabel: row.staffStatusLabel || row.statusLabel,
  };
};

export const parseStaffSheetMorningBaselineFile = async (
  file: File,
  unitLabel: string,
): Promise<BchsMorningDailySnapshot> => {
  if (!unitLabel.trim()) {
    throw new Error(
      "Оберіть одну роту у фільтрі колонки «Підрозділ» перед завантаженням Штатки.",
    );
  }

  const [workbook, imported] = await Promise.all([
    readWorkbookSnapshot(file),
    parseStaffSheetImportFile(file),
  ]);
  const rosterSheet = findRosterSheet(workbook.sheets);
  const columns = rosterSheet ? buildImportColumns(rosterSheet) : undefined;
  const rosterLabels = Object.fromEntries(
    (columns ?? []).map((column) => [column.key, column.label]),
  );
  const rosterRows = fillDownRosterUnitRows(imported.rows);
  const overviewRows = filterBchsMorningUnitRows(
    buildStaffOverviewRowsFromRoster(
      rosterRows,
      rosterLabels,
      columns,
    ).map(withStaffOverviewStatus),
    unitLabel,
  );

  if (!overviewRows.length) {
    throw new Error(
      `У Штатці не знайдено осіб для роти «${unitLabel}».`,
    );
  }

  const reportDate =
    parseBchsBaselineDateFromFileName(file.name) ?? kyivIsoDateLabel();
  const [year, month, day] = reportDate.split("-").map(Number);

  return {
    ...buildBchsMorningDailySnapshot(
      overviewRows,
      unitLabel,
      rosterRows,
      new Date(Date.UTC(year, month - 1, day)),
    ),
    source: "manual",
    baselineKind: "staff",
    sourceLabel: file.name,
  };
};
