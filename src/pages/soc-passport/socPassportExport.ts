import { exportBlankWorkbookWithMutations } from "../../excelRoundTrip";
import type { PassportTableRow, SocPassportResult, SocPerson } from "./socPassportTypes";

const HEADER_FILL = "D9E2F3";
const SECTION_FILL = "E7E6E6";
const VALUE_FILL = "FFFFFF";
const ALT_ROW_FILL = "F2F2F2";
const TEXT = "000000";
const BORDER = {
  top: { style: "thin", color: "000000" },
  bottom: { style: "thin", color: "000000" },
  left: { style: "thin", color: "000000" },
  right: { style: "thin", color: "000000" },
};

const metricValues = (row: PassportTableRow) => {
  const counts = row.counts;
  if (!counts) return ["", "", "", "", "", "", ""];
  return [
    counts.officerMobilized,
    counts.officerContract,
    counts.sergeantMobilized,
    counts.sergeantContract,
    counts.soldierMobilized,
    counts.soldierContract,
    row.total ?? 0,
  ];
};

const styleHeader = (cell: any) => {
  cell.style({
    bold: true,
    fontColor: TEXT,
    fill: HEADER_FILL,
    horizontalAlignment: "center",
    verticalAlignment: "center",
    wrapText: true,
    border: BORDER,
  });
};

const writePortraitSheet = (sheet: any, result: SocPassportResult) => {
  sheet.name("Соц.портрет");
  sheet.row(1).height(22);
  sheet.row(2).height(28);
  sheet.column(1).width(8);
  sheet.column(2).width(46);
  for (let column = 3; column <= 9; column += 1) sheet.column(column).width(16);

  sheet.cell(1, 1).value("10. Соціальний портрет частини (підрозділу)");
  sheet.range(1, 1, 1, 9).merged(true).style({
    bold: true,
    fontColor: TEXT,
    fontSize: 14,
    fill: VALUE_FILL,
  });

  sheet.cell(2, 1).value("№ з/п");
  sheet.cell(2, 2).value("Вид обліку");
  sheet.cell(2, 3).value("Офіцери");
  sheet.cell(2, 5).value("Сержанти");
  sheet.cell(2, 7).value("Солдати");
  sheet.cell(2, 9).value("Всього");
  sheet.range(2, 3, 2, 4).merged(true);
  sheet.range(2, 5, 2, 6).merged(true);
  sheet.range(2, 7, 2, 8).merged(true);
  sheet.range(2, 1, 3, 1).merged(true);
  sheet.range(2, 2, 3, 2).merged(true);
  sheet.range(2, 9, 3, 9).merged(true);
  sheet.cell(3, 3).value("Мобілізовані");
  sheet.cell(3, 4).value("Контракт");
  sheet.cell(3, 5).value("Мобілізовані");
  sheet.cell(3, 6).value("Контракт");
  sheet.cell(3, 7).value("Мобілізовані");
  sheet.cell(3, 8).value("Контракт");

  for (let column = 1; column <= 9; column += 1) {
    styleHeader(sheet.cell(2, column));
    styleHeader(sheet.cell(3, column));
  }

  result.rows.forEach((row, index) => {
    const excelRow = 4 + index;
    const isSection = row.kind === "section";
    sheet.cell(excelRow, 1).value(row.number);
    sheet.cell(excelRow, 2).value(row.label);
    metricValues(row).forEach((value, columnIndex) => {
      if (!isSection) sheet.cell(excelRow, 3 + columnIndex).value(value || 0);
    });
    sheet.range(excelRow, 1, excelRow, 9).style({
      fill: isSection ? SECTION_FILL : VALUE_FILL,
      fontColor: TEXT,
      bold: isSection,
      border: BORDER,
      verticalAlignment: "center",
      horizontalAlignment: "center",
    });
    sheet.cell(excelRow, 2).style({
      horizontalAlignment: "left",
    });
  });
};

const writePeopleSheet = (sheet: any, people: SocPerson[]) => {
  sheet.name("Розбір");
  const headers = [
    "ПІБ",
    "Посада",
    "Індекс посади",
    "Категорія",
    "Вид служби",
    "Звання",
    "Стать",
    "Вік",
    "Регіон",
    "Місце народження",
    "Національність",
    "Сімейний стан",
    "Дітей до 18",
    "3+ дітей",
    "Родичі служать",
    "Родичі за кордоном",
    "Родичі в рф/РБ",
    "УБД",
    "ВПО",
    "Звідки прибув",
    "Ким призваний",
    "Статус (ранок)",
    "Виходів",
    "ООС",
    "Ранковий",
    "Примітки розбору",
    "Дані про родичів",
  ];
  headers.forEach((header, index) => {
    sheet.cell(1, index + 1).value(header);
    styleHeader(sheet.cell(1, index + 1));
    sheet.column(index + 1).width(index === 0 || index === 9 || index >= 25 ? 36 : 18);
  });

  people.forEach((person, index) => {
    const values = [
      person.name,
      person.position,
      person.positionIndex,
      person.rankGroup,
      person.serviceType,
      person.rank,
      person.sex,
      person.age ?? "",
      person.regionLabel,
      person.birthPlace,
      person.nationality,
      person.marital,
      person.childrenUnder18,
      person.children3plus ? "так" : "",
      person.relativesServing ? "так" : "",
      person.relativesAbroad ? "так" : "",
      person.relativesHostile ? "так" : "",
      person.hasUbd ? "так" : "",
      person.isIdp ? "так" : "",
      person.arrivedFrom,
      person.calledBy,
      person.morningStatus,
      person.exitCount,
      person.match.oos ? "так" : "ні",
      person.match.morning ? "так" : "ні",
      person.parseNotes.join("; "),
      person.relativesRaw,
    ];
    const excelRow = index + 2;
    const fill = index % 2 === 0 ? VALUE_FILL : ALT_ROW_FILL;
    values.forEach((value, columnIndex) => {
      sheet.cell(excelRow, columnIndex + 1).value(value).style({
        fontColor: TEXT,
        fill,
        border: BORDER,
        verticalAlignment: "center",
      });
    });
  });
};

export const exportSocPassportWorkbook = async (result: SocPassportResult) => {
  const stamp = new Date().toISOString().slice(0, 10);
  await exportBlankWorkbookWithMutations((workbook) => {
    const portrait = workbook.sheet(0);
    writePortraitSheet(portrait, result);
    const peopleSheet = workbook.addSheet("Розбір");
    writePeopleSheet(peopleSheet, result.people);
  }, `Соц.паспорт_${stamp}.xlsx`);
};
