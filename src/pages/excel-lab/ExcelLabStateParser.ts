import type { ExcelSheetSnapshot } from "@/excelRoundTrip";

const EXCEL_LAB_STATE_SHEET = {
  generalList: "1.ОС Загальний список",
  archive: "Архів",
} as const;

type ExcelLabStateGeneralList = {
  alias: string;
  status: string;
  unit: string;
  countUpdate: number;
};

export type EntityStateSheets = {
  [key: string]: ExcelLabStateGeneralList;
};

const SheetValues = {
  transiter: "ТРАНЗИТЕР",
  attachedPersonal: "ПРИКОМАНДИРОВАНІ",
} as const;

export const parseExcelLabState = (
  sheets: ExcelSheetSnapshot[],
): EntityStateSheets => {
  const generalList: EntityStateSheets = {};
  sheets.map((sheet) => {
    switch (sheet.sheetName) {
      case EXCEL_LAB_STATE_SHEET.generalList:
        return parseGeneralList(sheet, generalList);
      default:
        return null;
    }
  });
  return generalList;
};

const parseGeneralList = (
  sheet: ExcelSheetSnapshot,
  generalList: EntityStateSheets,
) => {
  sheet.rows.forEach((row) => {
    if (
      row.values[28] === SheetValues.transiter ||
      row.values[1] === SheetValues.attachedPersonal
    ) {
      return;
    }
    if (row.values[13] !== null) {
      if (!generalList[row.values[13] as string]) {
        generalList[row.values[13] as string] = {
          alias: row.values[14] as string,
          status: row.values[20] as string,
          unit: row.values[1] as string,
          countUpdate: 0,
        };
      } else {
        generalList[row.values[13] as string].countUpdate++;
      }
    }
  });
  return generalList;
};
