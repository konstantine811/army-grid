import type { ExcelSheetSnapshot } from "@/excelRoundTrip";

const EXCEL_LAB_CARDS_SHEET = {
  vk: "ВК",
  tpv: "ТПВ",
  dovidki: "ДОВІДКИ",
} as const;

const StatusString = {
  has: "Є",
  not: "відсутній",
} as const;

export type CardSheetKind =
  (typeof EXCEL_LAB_CARDS_SHEET)[keyof typeof EXCEL_LAB_CARDS_SHEET];

type CardData = {
  status: boolean;
  cardNumber: string;
  alias: string;
  whereToGet: string;
};

type TPVCardData = CardData & {
  skan: boolean;
};

export type ExcelLabStateCards = {
  [EXCEL_LAB_CARDS_SHEET.vk]?: CardData;
  [EXCEL_LAB_CARDS_SHEET.tpv]?: TPVCardData;
  [EXCEL_LAB_CARDS_SHEET.dovidki]?: CardData;
};

export type EntityCardsSheets = Record<string, ExcelLabStateCards>;

const normalizeAliasValue = (value: unknown) => {
  if (value == null) return "";
  const text = String(value).trim();
  if (!text) return "";
  if (/^#N\/A$/i.test(text)) return "";
  if (text === "[object Object]") return "";
  return text;
};

const pickAliasFromRow = (values: unknown[], indices: number[]) => {
  for (const index of indices) {
    const alias = normalizeAliasValue(values[index]);
    if (alias) return alias;
  }
  return "";
};

export const parseExcelLabMilitaryServiceCards = (
  sheets: ExcelSheetSnapshot[],
): EntityCardsSheets => {
  const cards: EntityCardsSheets = {};
  for (const sheet of sheets) {
    switch (sheet.sheetName) {
      case EXCEL_LAB_CARDS_SHEET.vk:
        parseCards(sheet, cards);
        break;
      case EXCEL_LAB_CARDS_SHEET.tpv:
        parseTPV(sheet, cards);
        break;
      case EXCEL_LAB_CARDS_SHEET.dovidki:
        parseDovidki(sheet, cards);
        break;
      default:
        break;
    }
  }
  return cards;
};

const parseTPV = (sheet: ExcelSheetSnapshot, cards: EntityCardsSheets) => {
  sheet.rows.forEach((row) => {
    const personKey = String(row.values[6] ?? "").trim();

    if (!personKey) return;

    if (!cards[personKey]) {
      cards[personKey] = {};
    }
    cards[personKey][EXCEL_LAB_CARDS_SHEET.tpv] = {
      status: StatusString.has === String(row.values[20] ?? "") ? true : false,
      alias: pickAliasFromRow(row.values, [7, 8, 9]),
      cardNumber: String(row.values[13] ?? ""),
      skan: StatusString.has === String(row.values[19] ?? "") ? true : false,
      whereToGet: String(row.values[31] ?? ""),
    };
  });
  return cards;
};

const parseDovidki = (sheet: ExcelSheetSnapshot, cards: EntityCardsSheets) => {
  sheet.rows.forEach((row) => {
    const personKey = String(row.values[6] ?? "").trim();
    if (!personKey) return;

    if (!cards[personKey]) {
      cards[personKey] = {};
    }

    cards[personKey][EXCEL_LAB_CARDS_SHEET.dovidki] = {
      status: StatusString.has === String(row.values[16] ?? "") ? true : false,
      alias: pickAliasFromRow(row.values, [7, 8, 9, 10]),
      cardNumber: String(row.values[23] ?? ""),
      whereToGet: String(row.values[31] ?? ""),
    };
  });
  return cards;
};

const parseCards = (sheet: ExcelSheetSnapshot, cards: EntityCardsSheets) => {
  sheet.rows.forEach((row) => {
    const personKey = String(row.values[7] ?? "").trim();
    if (!personKey) return;

    if (!cards[personKey]) {
      cards[personKey] = {};
    }

    cards[personKey][EXCEL_LAB_CARDS_SHEET.vk] = {
      status: StatusString.has === String(row.values[20] ?? "") ? true : false,
      alias: String(row.values[8] ?? ""),
      cardNumber: String(row.values[29] ?? ""),
      whereToGet: String(row.values[31] ?? ""),
    };
  });
  return cards;
};
