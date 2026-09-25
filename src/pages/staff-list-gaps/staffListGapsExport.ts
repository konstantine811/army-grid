import { exportWorkbookFileWithMutations } from "../../excelRoundTrip";
import { formatUkDateForExcelWrite } from "../../shared/format";
import type { StaffListColumn, StaffListRow, StaffListSnapshot } from "./staffListGapsParse";

const normalizeHeader = (value: string) =>
  value.trim().toLocaleLowerCase("uk-UA");

export const staffListCellValueForExcelWrite = (
  column: StaffListColumn,
  snapshot: StaffListSnapshot,
  raw: string,
) => {
  const text = String(raw ?? "").trim();
  if (!text) return null;
  if (column.id === snapshot.birthDateColumnId) {
    return formatUkDateForExcelWrite(text);
  }
  if (normalizeHeader(column.label) === "№" && /^\d+$/.test(text)) {
    return Number(text);
  }
  return text;
};

export const buildStaffListExportFileName = (fileName: string) => {
  const base = fileName.replace(/\.xlsx?$/i, "");
  return `${base}_filled.xlsx`;
};

export const exportStaffListExcel = async (
  snapshot: StaffListSnapshot,
  rows: StaffListRow[],
  fileData: ArrayBuffer,
) => {
  const file = new File([fileData], snapshot.fileName, {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });

  await exportWorkbookFileWithMutations(
    file,
    (workbook) => {
      const sheet = workbook.sheet(snapshot.sheetIndex);
      if (!sheet) {
        throw new Error(`Не знайдено аркуш "${snapshot.sheetName}" для експорту.`);
      }

      for (const row of rows) {
        for (const column of snapshot.columns) {
          const excelColumn =
            (snapshot.columnIndexes[column.index] ?? column.index) + 1;
          const value = staffListCellValueForExcelWrite(
            column,
            snapshot,
            row.values[column.id] ?? "",
          );
          sheet.cell(row.__excelRowNumber, excelColumn).value(value);
        }
      }
    },
    buildStaffListExportFileName(snapshot.fileName),
  );
};
