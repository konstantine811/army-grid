import type { ExcelSheetSnapshot } from "@/excelRoundTrip";

const EXCEL_LAB_KSP_SHEET = {
  vk: "ВК",
} as const;

type ExcelLabKSPVKList = {
  cardType: string;
  soldierStatus: string;
  soldierOperation: string;
  note: string;
};

export type EntityKSPVKSheets = {
  [key: string]: ExcelLabKSPVKList;
};

export const parseExcelKSPState = (
  sheets: ExcelSheetSnapshot[],
): EntityKSPVKSheets => {
  const generalList: EntityKSPVKSheets = {};
  sheets.map((sheet) => {
    switch (sheet.sheetName) {
      case EXCEL_LAB_KSP_SHEET.vk:
        return parseKSPVKList(sheet, generalList);
      default:
        return null;
    }
  });
  console.log("[excel-lab] ksp vk data:", generalList);
  return generalList;
};

const parseKSPVKList = (
  sheet: ExcelSheetSnapshot,
  generalList: EntityKSPVKSheets,
) => {
  sheet.rows.forEach((row) => {
    if (row.values[3] !== null && isNaN(row.values[3] as number)) {
      generalList[row.values[3] as string] = {
        cardType: row.values[8] as string,
        soldierStatus: row.values[7] as string,
        soldierOperation: row.values[5] as string,
        note: row.values[10] as string,
      };
    }
  });
  return generalList;
};
