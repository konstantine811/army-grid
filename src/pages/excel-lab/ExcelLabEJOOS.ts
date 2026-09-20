import type { ExcelSheetSnapshot } from "@/excelRoundTrip";

const EXCEL_LAB_EJOOS_SHEET = {
  oos: "2. ООС",
} as const;

type ExcelLabEJOOSList = {
  cardNumber: string;
};

export type EntityEJOOSSheets = {
  [key: string]: ExcelLabEJOOSList;
};

export const parseExcelEJOOSState = (
  sheets: ExcelSheetSnapshot[],
): EntityEJOOSSheets => {
  const generalList: EntityEJOOSSheets = {};
  sheets.map((sheet) => {
    switch (sheet.sheetName) {
      case EXCEL_LAB_EJOOS_SHEET.oos:
        return parseOOSList(sheet, generalList);
      default:
        return null;
    }
  });
  return generalList;
};

const parseOOSList = (
  sheet: ExcelSheetSnapshot,
  generalList: EntityEJOOSSheets,
) => {
  sheet.rows.forEach((row) => {
    if (row.values[1] !== null && isNaN(row.values[1] as number)) {
      generalList[row.values[1] as string] = {
        cardNumber: row.values[34] as string,
      };
    }
  });
  return generalList;
};
