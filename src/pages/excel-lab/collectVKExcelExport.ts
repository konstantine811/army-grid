import writeXlsxFile, {
  type CellObject,
  type SheetData,
} from "write-excel-file/browser";
import type { CollectedVKData, CollectedVKPerson } from "./CollectVKData";

const HEADER_BG = "#E8EAED";
const HEADER_TEXT = "#202124";
const BORDER = "#BDBDBD";
const WHITE = "#FFFFFF";
const ZEBRA = "#F8F9FA";

const border = {
  borderColor: BORDER,
  borderStyle: "thin" as const,
};

const cell = (
  value: string | number | null | undefined,
  extra: Omit<CellObject, "value"> = {},
): CellObject => ({
  value: value == null || value === "" ? undefined : value,
  fontFamily: "Arial",
  fontSize: 10,
  alignVertical: "center",
  wrap: true,
  ...border,
  ...extra,
});

export type CollectVKExportRow = {
  pib: string;
  staffAlias: string;
  staffStatus: string;
  staffUnit: string;
  ejoosCardNumber: string;
  kspCardType: string;
  kspSoldierStatus: string;
  kspSoldierOperation: string;
  kspNote: string;
  vkStatus: string;
  vkCardNumber: string;
  vkAlias: string;
  tpvStatus: string;
  tpvCardNumber: string;
  dovidkyStatus: string;
  dovidkyCardNumber: string;
};

export type CollectVKExportColumn = {
  id: keyof CollectVKExportRow;
  label: string;
  width: number;
};

/** Колонки таблиці — редагуй під свої потреби. */
export const COLLECT_VK_EXPORT_COLUMNS: CollectVKExportColumn[] = [
  { id: "pib", label: "ПІБ", width: 34 },
  { id: "staffAlias", label: "Позивний (штатка)", width: 14 },
  { id: "staffStatus", label: "Статус (штатка)", width: 14 },
  { id: "staffUnit", label: "Підрозділ", width: 18 },
  { id: "ejoosCardNumber", label: "ЄЖООС № ВК", width: 14 },
  { id: "kspCardType", label: "КСП тип", width: 10 },
  { id: "kspSoldierStatus", label: "КСП статус", width: 14 },
  { id: "kspSoldierOperation", label: "КСП операція", width: 16 },
  { id: "kspNote", label: "КСП примітка", width: 20 },
  { id: "vkStatus", label: "ВК (картки)", width: 12 },
  { id: "vkCardNumber", label: "№ ВК", width: 14 },
  { id: "vkAlias", label: "Позивний ВК", width: 12 },
  { id: "tpvStatus", label: "ТПВ", width: 12 },
  { id: "tpvCardNumber", label: "№ ТПВ", width: 14 },
  { id: "dovidkyStatus", label: "Довідки", width: 12 },
  { id: "dovidkyCardNumber", label: "№ довідки", width: 14 },
];

const cardStatusLabel = (status?: boolean) => {
  if (status === true) return "Є";
  if (status === false) return "відсутній";
  return "";
};

/** Один рядок таблиці з CollectedVKPerson — шаблон для мапінгу полів. */
export const mapCollectedVKPersonToExportRow = (
  pib: string,
  person: CollectedVKPerson,
): CollectVKExportRow => {
  const vk = person.militaryCards?.ВК;
  const tpv = person.militaryCards?.ТПВ;
  const dovidky = person.militaryCards?.ДОВІДКИ;

  return {
    pib,
    staffAlias: person.staff?.alias ?? "",
    staffStatus: person.staff?.status ?? "",
    staffUnit: person.staff?.unit ?? "",
    ejoosCardNumber: person.ejoos?.cardNumber ?? "",
    kspCardType: person.ksp?.cardType ?? "",
    kspSoldierStatus: person.ksp?.soldierStatus ?? "",
    kspSoldierOperation: person.ksp?.soldierOperation ?? "",
    kspNote: person.ksp?.note ?? "",
    vkStatus: cardStatusLabel(vk?.status),
    vkCardNumber: vk?.cardNumber ?? "",
    vkAlias: vk?.alias ?? "",
    tpvStatus: cardStatusLabel(tpv?.status),
    tpvCardNumber: tpv?.cardNumber ?? "",
    dovidkyStatus: cardStatusLabel(dovidky?.status),
    dovidkyCardNumber: dovidky?.cardNumber ?? "",
  };
};

export const mapCollectedVKSheetsToExportRows = (
  sheets: Record<string, CollectedVKPerson>,
) =>
  Object.entries(sheets).map(([pib, person]) =>
    mapCollectedVKPersonToExportRow(pib, person),
  );

const buildSheetData = (rows: CollectVKExportRow[]): SheetData => {
  const headerRow = COLLECT_VK_EXPORT_COLUMNS.map((column) =>
    cell(column.label, {
      fontWeight: "bold",
      textColor: HEADER_TEXT,
      backgroundColor: HEADER_BG,
      align: "center",
      height: 32,
    }),
  );

  const bodyRows = rows.map((row, index) => {
    const zebra =
      index % 2 === 1 ? { backgroundColor: ZEBRA } : { backgroundColor: WHITE };
    return COLLECT_VK_EXPORT_COLUMNS.map((column) =>
      cell(String(row[column.id] ?? ""), {
        ...zebra,
        align: column.id === "pib" ? "left" : "center",
        fontWeight: column.id === "pib" ? "bold" : undefined,
      }),
    );
  });

  return [headerRow, ...bodyRows];
};

export type CollectVKExportSheet = {
  sheet: string;
  data: SheetData;
  columns: Array<{ width: number }>;
  stickyRowsCount: number;
  stickyColumnsCount: number;
  showGridLines: boolean;
  orientation: "landscape";
};

/** Аркуші workbook — кожен відповідає одному з обʼєктів collectVKData. */
export const buildCollectVKExportSheets = (
  collected: CollectedVKData,
): CollectVKExportSheet[] => {
  const sheetOptions = {
    columns: COLLECT_VK_EXPORT_COLUMNS.map((column) => ({
      width: column.width,
    })),
    stickyRowsCount: 1,
    stickyColumnsCount: 1,
    showGridLines: true,
    orientation: "landscape" as const,
  };

  const sheet = (
    name: string,
    rows: CollectVKExportRow[],
  ): CollectVKExportSheet => ({
    ...sheetOptions,
    sheet: name,
    data: buildSheetData(rows),
  });

  return [
    sheet("Усі", mapCollectedVKSheetsToExportRows(collected.all)),
    sheet("Без ВК", mapCollectedVKSheetsToExportRows(collected.notVK)),
    sheet(
      "ВК картки",
      mapCollectedVKSheetsToExportRows(collected.vkOnMilitaryCards),
    ),
    sheet("ВК КСП", mapCollectedVKSheetsToExportRows(collected.vkOnKSP)),
    sheet(
      "ВК картки+КСП",
      mapCollectedVKSheetsToExportRows(collected.vkOnMilitaryCardsAndKSP),
    ),
  ];
};

export const buildCollectVKExportFileName = (exportedAt = new Date()) => {
  const stamp = exportedAt.toISOString().slice(0, 10);
  return `excel-lab-vk-${stamp}.xlsx`;
};

/** Завантажити .xlsx з 5 аркушами за групами collectVKData. */
export const exportCollectVKExcel = async (
  collected: CollectedVKData,
  fileName?: string,
) => {
  await writeXlsxFile(buildCollectVKExportSheets(collected), {
    fontFamily: "Times New Roman",
    fontSize: 14,
  }).toFile(fileName ?? buildCollectVKExportFileName());
};
