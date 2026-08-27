import { exportBlankWorkbookWithMutations } from "../../excelRoundTrip";
import type {
  ArrivalsMonthResult,
  CombatLossesResult,
  DeparturesResult,
  DispositionArchiveResult,
  SzchRuhResult,
} from "./socPassportDepartures";

const HEADER_FILL = "D9E2F3";
const SECTION_FILL = "E7E6E6";
const VALUE_FILL = "FFFFFF";
const TEXT = "000000";
const BORDER = {
  top: { style: "thin", color: "000000" },
  bottom: { style: "thin", color: "000000" },
  left: { style: "thin", color: "000000" },
  right: { style: "thin", color: "000000" },
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

/** Таблиця джерел прибуття з розбивкою за званням. Повертає наступний вільний рядок. */
const writeArrivalsSourceRankTable = (
  sheet: any,
  startRow: number,
  arrivals: ArrivalsMonthResult,
  title: string,
  subtitle: string,
): number => {
  let cursor = startRow;
  sheet.cell(cursor, 1).value(title);
  sheet.range(cursor, 1, cursor, 6).merged(true).style({
    bold: true,
    fontSize: 13,
  });
  cursor += 1;
  sheet.cell(cursor, 1).value(subtitle);
  sheet.range(cursor, 1, cursor, 6).merged(true);
  cursor += 2;
  sheet.cell(cursor, 1).value("№");
  sheet.cell(cursor, 2).value("За джерелом прибуття");
  sheet.cell(cursor, 3).value("Офіцери");
  sheet.cell(cursor, 4).value("Сержанти");
  sheet.cell(cursor, 5).value("Солдати");
  sheet.cell(cursor, 6).value("Разом");
  for (let c = 1; c <= 6; c += 1) styleHeader(sheet.cell(cursor, c));

  arrivals.sourceSummary.forEach((row, index) => {
    const excelRow = cursor + 1 + index;
    sheet.cell(excelRow, 1).value(String(index + 1));
    sheet.cell(excelRow, 2).value(row.label);
    sheet.cell(excelRow, 3).value(row.byRank.officer || "");
    sheet.cell(excelRow, 4).value(row.byRank.sergeant || "");
    sheet.cell(excelRow, 5).value(row.byRank.soldier || "");
    sheet.cell(excelRow, 6).value(row.count);
    sheet.range(excelRow, 1, excelRow, 6).style({
      border: BORDER,
      fill: VALUE_FILL,
      verticalAlignment: "center",
    });
    sheet.cell(excelRow, 6).style({ horizontalAlignment: "center", bold: true });
  });

  const totalRow = cursor + 1 + arrivals.sourceSummary.length;
  sheet.cell(totalRow, 1).value("");
  sheet.cell(totalRow, 2).value("Разом");
  sheet.cell(totalRow, 3).value(arrivals.byRank.officer || "");
  sheet.cell(totalRow, 4).value(arrivals.byRank.sergeant || "");
  sheet.cell(totalRow, 5).value(arrivals.byRank.soldier || "");
  sheet.cell(totalRow, 6).value(arrivals.total);
  sheet.range(totalRow, 1, totalRow, 6).style({
    border: BORDER,
    fill: SECTION_FILL,
    bold: true,
    verticalAlignment: "center",
  });
  sheet.cell(totalRow, 6).style({ horizontalAlignment: "center", bold: true });
  return totalRow + 1;
};

const writeDeparturesSummarySheet = (sheet: any, result: DeparturesResult) => {
  sheet.name("Вибули");
  sheet.column(1).width(8);
  sheet.column(2).width(52);
  sheet.column(3).width(12);
  sheet.column(4).width(12);
  sheet.column(5).width(12);
  sheet.column(6).width(12);

  sheet.cell(1, 1).value("Вибули (з аркуша ЕЖООС «Виключені»)");
  sheet.range(1, 1, 1, 6).merged(true).style({
    bold: true,
    fontSize: 13,
    fill: VALUE_FILL,
  });
  sheet.cell(2, 1).value(`Джерело: ${result.sourceSheet}`);
  sheet.range(2, 1, 2, 6).merged(true);

  sheet.cell(4, 1).value("№");
  sheet.cell(4, 2).value("Категорія");
  sheet.cell(4, 3).value("Офіцери");
  sheet.cell(4, 4).value("Сержанти");
  sheet.cell(4, 5).value("Солдати");
  sheet.cell(4, 6).value("Разом");
  for (let c = 1; c <= 6; c += 1) styleHeader(sheet.cell(4, c));

  result.summary.forEach((row, index) => {
    const excelRow = 5 + index;
    const isTransfer = row.category === "transfer";
    sheet.cell(excelRow, 1).value(isTransfer ? "—" : String(index + 1));
    sheet.cell(excelRow, 2).value(row.label);
    sheet.cell(excelRow, 3).value(row.byRank.officer);
    sheet.cell(excelRow, 4).value(row.byRank.sergeant);
    sheet.cell(excelRow, 5).value(row.byRank.soldier);
    sheet.cell(excelRow, 6).value(row.count);
    sheet.range(excelRow, 1, excelRow, 6).style({
      border: BORDER,
      fill: isTransfer ? SECTION_FILL : VALUE_FILL,
      verticalAlignment: "center",
    });
    sheet.cell(excelRow, 2).style({ horizontalAlignment: "left" });
    for (let c = 3; c <= 6; c += 1) {
      sheet.cell(excelRow, c).style({ horizontalAlignment: "center", bold: c === 6 });
    }
  });

  const totalRow = 5 + result.summary.length + 1;
  const footerRows: Array<{
    label: string;
    byRank: { officer: number; sergeant: number; soldier: number };
    total: number;
  }> = [
    {
      label: "Усього у періоді",
      byRank: result.totals.byRank,
      total: result.totals.all,
    },
    {
      label: "З них звільнення / інше (без переведень)",
      byRank: result.totals.dischargesByRank,
      total: result.totals.discharges,
    },
    {
      label: "Переведені / розпорядження",
      byRank: result.totals.transfersByRank,
      total: result.totals.transfers,
    },
  ];
  footerRows.forEach((row, index) => {
    const excelRow = totalRow + index;
    sheet.cell(excelRow, 2).value(row.label);
    sheet.cell(excelRow, 3).value(row.byRank.officer);
    sheet.cell(excelRow, 4).value(row.byRank.sergeant);
    sheet.cell(excelRow, 5).value(row.byRank.soldier);
    sheet.cell(excelRow, 6).value(row.total);
    sheet.range(excelRow, 1, excelRow, 6).style({
      border: BORDER,
      bold: true,
      fill: SECTION_FILL,
    });
    for (let c = 3; c <= 6; c += 1) {
      sheet.cell(excelRow, c).style({ horizontalAlignment: "center" });
    }
  });

  let cursor = totalRow + footerRows.length + 1;
  if (result.periodFromLabel) {
    sheet.cell(cursor, 1).value(
      `Період: ${result.periodFromLabel}. У аркуші всього ${result.totalUnfiltered}; до періоду ${result.skippedBeforePeriod}; без дати ${result.skippedNoDate}.`,
    );
    sheet.range(cursor, 1, cursor, 6).merged(true);
    cursor += 2;
  }
  sheet.cell(cursor, 1).value(
    "Класифікація за текстом підстави/типу. Переведені не входять у п.1–5 як «звільнення». Джерело: ЕЖООС «3. Виключені».",
  );
  sheet.range(cursor, 1, cursor, 6).merged(true);

  const arrivals = result.arrivalsAugust;
  if (arrivals) {
    cursor += 2;
    cursor = writeArrivalsSourceRankTable(
      sheet,
      cursor,
      arrivals,
      `Прибули до військової частини протягом місяця (${arrivals.monthLabel})`,
      `Джерело: ${arrivals.sourceSheet} · «Звідки прибув» (порожнє = ТЦК)`,
    );
  }

  const arrivalsMorning = result.arrivalsFromMorning;
  if (arrivalsMorning) {
    cursor += 2;
    cursor = writeArrivalsSourceRankTable(
      sheet,
      cursor,
      arrivalsMorning,
      `Звідки прибув · Штатка / ранковий (${arrivalsMorning.monthLabel})`,
      `База: ранковий. БРЕЗ з Штатки; решта — ООС «Звідки прибув» (порожнє = ТЦК). ${arrivalsMorning.sourceSheet}`,
    );
  }

  const arrivalsPb = result.arrivalsAugustPb;
  if (arrivalsPb) {
    cursor += 2;
    cursor = writeArrivalsSourceRankTable(
      sheet,
      cursor,
      arrivalsPb,
      `Прибули з 1ПБ · Рух (ПРИБУВ) · ${arrivalsPb.monthLabel}`,
      `Джерело: ${arrivalsPb.sourceSheet} · «Звідки прибув» (порожнє = ТЦК)`,
    );
  }

  const disposition = result.dispositionFromArchive;
  if (disposition) {
    cursor += 2;
    sheet.cell(cursor, 1).value(
      `Виведені у розпорядження · archive${
        disposition.periodFrom ? ` · з ${disposition.periodFrom}` : ""
      }`,
    );
    sheet.range(cursor, 1, cursor, 6).merged(true).style({
      bold: true,
      fontSize: 13,
    });
    cursor += 2;
    sheet.cell(cursor, 1).value("№");
    sheet.cell(cursor, 2).value("Підстава");
    sheet.cell(cursor, 3).value("Офіцери");
    sheet.cell(cursor, 4).value("Сержанти");
    sheet.cell(cursor, 5).value("Солдати");
    sheet.cell(cursor, 6).value("Разом");
    for (let c = 1; c <= 6; c += 1) styleHeader(sheet.cell(cursor, c));
    disposition.summary.forEach((row, index) => {
      const excelRow = cursor + 1 + index;
      sheet.cell(excelRow, 1).value(String(index + 1));
      sheet.cell(excelRow, 2).value(row.label);
      sheet.cell(excelRow, 3).value(row.byRank.officer);
      sheet.cell(excelRow, 4).value(row.byRank.sergeant);
      sheet.cell(excelRow, 5).value(row.byRank.soldier);
      sheet.cell(excelRow, 6).value(row.count);
      sheet.range(excelRow, 1, excelRow, 6).style({
        border: BORDER,
        fill: VALUE_FILL,
        verticalAlignment: "center",
      });
      for (let c = 3; c <= 6; c += 1) {
        sheet.cell(excelRow, c).style({
          horizontalAlignment: "center",
          bold: c === 6,
        });
      }
    });
    const totalExcelRow = cursor + 1 + disposition.summary.length;
    sheet.cell(totalExcelRow, 2).value("Разом у розпорядженні");
    sheet.cell(totalExcelRow, 3).value(disposition.totals.byRank.officer);
    sheet.cell(totalExcelRow, 4).value(disposition.totals.byRank.sergeant);
    sheet.cell(totalExcelRow, 5).value(disposition.totals.byRank.soldier);
    sheet.cell(totalExcelRow, 6).value(disposition.totals.all);
    sheet.range(totalExcelRow, 1, totalExcelRow, 6).style({
      border: BORDER,
      bold: true,
      fill: SECTION_FILL,
    });
    cursor = totalExcelRow;
  }

  const szch = result.szchFromRuh;
  if (szch) {
    cursor += 2;
    sheet.cell(cursor, 1).value(
      `СЗЧ · Рух${szch.periodFrom ? ` · з ${szch.periodFrom}` : ""}`,
    );
    sheet.range(cursor, 1, cursor, 6).merged(true).style({
      bold: true,
      fontSize: 13,
    });
    cursor += 2;
    sheet.cell(cursor, 1).value("№");
    sheet.cell(cursor, 2).value("Категорія");
    sheet.cell(cursor, 3).value("Кількість");
    for (let c = 1; c <= 3; c += 1) styleHeader(sheet.cell(cursor, c));
    const rows: Array<{ num: string; label: string; count: number }> = [
      { num: "1", label: "Офіцери", count: szch.totals.byRank.officer },
      { num: "2", label: "Сержанти", count: szch.totals.byRank.sergeant },
      { num: "3", label: "Солдати", count: szch.totals.byRank.soldier },
      { num: "", label: "Разом СЗЧ", count: szch.totals.all },
    ];
    rows.forEach((row, index) => {
      const excelRow = cursor + 1 + index;
      const isTotal = index === rows.length - 1;
      sheet.cell(excelRow, 1).value(row.num);
      sheet.cell(excelRow, 2).value(row.label);
      sheet.cell(excelRow, 3).value(row.count);
      sheet.range(excelRow, 1, excelRow, 3).style({
        border: BORDER,
        fill: isTotal ? SECTION_FILL : VALUE_FILL,
        bold: isTotal,
        verticalAlignment: "center",
      });
      sheet.cell(excelRow, 3).style({ horizontalAlignment: "center", bold: true });
    });
    cursor += rows.length + 1;
  }

  const losses = result.combatLossesFromPb;
  if (losses) {
    cursor += 2;
    sheet.cell(cursor, 1).value(
      `Втрати · Рух + archive${
        losses.periodFrom ? ` · з ${losses.periodFrom}` : ""
      }`,
    );
    sheet.range(cursor, 1, cursor, 6).merged(true).style({
      bold: true,
      fontSize: 13,
    });
    cursor += 2;
    sheet.cell(cursor, 1).value("№");
    sheet.cell(cursor, 2).value("Категорія");
    sheet.cell(cursor, 3).value("Офіцери");
    sheet.cell(cursor, 4).value("Сержанти");
    sheet.cell(cursor, 5).value("Солдати");
    sheet.cell(cursor, 6).value("Разом");
    for (let c = 1; c <= 6; c += 1) styleHeader(sheet.cell(cursor, c));
    losses.summary.forEach((row, index) => {
      const excelRow = cursor + 1 + index;
      sheet.cell(excelRow, 1).value(String(index + 1));
      sheet.cell(excelRow, 2).value(row.label);
      sheet.cell(excelRow, 3).value(row.byRank.officer);
      sheet.cell(excelRow, 4).value(row.byRank.sergeant);
      sheet.cell(excelRow, 5).value(row.byRank.soldier);
      sheet.cell(excelRow, 6).value(row.count);
      sheet.range(excelRow, 1, excelRow, 6).style({
        border: BORDER,
        fill: VALUE_FILL,
        verticalAlignment: "center",
      });
      for (let c = 3; c <= 6; c += 1) {
        sheet.cell(excelRow, c).style({
          horizontalAlignment: "center",
          bold: c === 6,
        });
      }
    });
    const totalExcelRow = cursor + 1 + losses.summary.length;
    sheet.cell(totalExcelRow, 2).value("Разом втрати");
    sheet.cell(totalExcelRow, 3).value(losses.totals.byRank.officer);
    sheet.cell(totalExcelRow, 4).value(losses.totals.byRank.sergeant);
    sheet.cell(totalExcelRow, 5).value(losses.totals.byRank.soldier);
    sheet.cell(totalExcelRow, 6).value(losses.totals.all);
    sheet.range(totalExcelRow, 1, totalExcelRow, 6).style({
      border: BORDER,
      bold: true,
      fill: SECTION_FILL,
    });
  }
};

const writeDispositionPeopleSheet = (
  sheet: any,
  disposition: DispositionArchiveResult,
) => {
  sheet.name("Розпорядження archive");
  const headers = [
    "№",
    "Підстава",
    "Група звання",
    "ПІБ",
    "Звання",
    "ID",
    "Вид вибуття",
    "Куди",
    "Дата вибуття",
    "Дата наказу",
    "Маркер",
    "Рядок Excel",
  ];
  headers.forEach((header, index) => {
    sheet.cell(1, index + 1).value(header);
    styleHeader(sheet.cell(1, index + 1));
  });
  sheet.column(1).width(6);
  sheet.column(2).width(36);
  sheet.column(3).width(14);
  sheet.column(4).width(36);
  sheet.column(5).width(16);
  sheet.column(6).width(12);
  sheet.column(7).width(28);
  sheet.column(8).width(28);
  sheet.column(9).width(14);
  sheet.column(10).width(14);
  sheet.column(11).width(20);
  sheet.column(12).width(10);

  disposition.people.forEach((person, index) => {
    const excelRow = 2 + index;
    const values = [
      index + 1,
      person.reasonLabel,
      person.rankGroupLabel,
      person.fullName,
      person.rank,
      person.personId,
      person.absenceType,
      person.place,
      person.departDate,
      person.orderDate,
      person.matchNote,
      person.excelRow,
    ];
    values.forEach((value, columnIndex) => {
      sheet.cell(excelRow, columnIndex + 1).value(value);
    });
    sheet.range(excelRow, 1, excelRow, headers.length).style({
      border: BORDER,
      fill: index % 2 === 0 ? VALUE_FILL : "F2F2F2",
      verticalAlignment: "center",
    });
  });
};

const writeSzchPeopleSheet = (sheet: any, szch: SzchRuhResult) => {
  sheet.name("СЗЧ Рух");
  const headers = [
    "№",
    "Група звання",
    "ПІБ",
    "Звання",
    "ID",
    "№ руху",
    "Тип",
    "Статус",
    "Куди",
    "Примітка",
    "Дата наказу",
    "Маркер",
    "Рядок Excel",
  ];
  headers.forEach((header, index) => {
    sheet.cell(1, index + 1).value(header);
    styleHeader(sheet.cell(1, index + 1));
  });
  sheet.column(1).width(6);
  sheet.column(2).width(14);
  sheet.column(3).width(36);
  sheet.column(4).width(16);
  sheet.column(5).width(12);
  sheet.column(6).width(10);
  sheet.column(7).width(14);
  sheet.column(8).width(16);
  sheet.column(9).width(24);
  sheet.column(10).width(28);
  sheet.column(11).width(14);
  sheet.column(12).width(24);
  sheet.column(13).width(10);

  szch.people.forEach((person, index) => {
    const excelRow = 2 + index;
    const values = [
      index + 1,
      person.rankGroupLabel,
      person.fullName,
      person.rank,
      person.personId,
      person.movementNumber,
      person.type,
      person.status,
      person.destination,
      person.note,
      person.orderDate,
      person.matchNote,
      person.excelRow,
    ];
    values.forEach((value, columnIndex) => {
      sheet.cell(excelRow, columnIndex + 1).value(value);
    });
    sheet.range(excelRow, 1, excelRow, headers.length).style({
      border: BORDER,
      fill: index % 2 === 0 ? VALUE_FILL : "F2F2F2",
      verticalAlignment: "center",
    });
  });
};

const writeCombatLossesPeopleSheet = (
  sheet: any,
  losses: CombatLossesResult,
) => {
  sheet.name("Втрати");
  const headers = [
    "№",
    "Категорія",
    "Група звання",
    "ПІБ",
    "Звання",
    "ID",
    "Джерело",
    "Тип / вид",
    "Статус",
    "Куди / примітка",
    "Дата",
    "Маркер",
    "Рядок Excel",
  ];
  headers.forEach((header, index) => {
    sheet.cell(1, index + 1).value(header);
    styleHeader(sheet.cell(1, index + 1));
  });
  sheet.column(1).width(6);
  sheet.column(2).width(18);
  sheet.column(3).width(14);
  sheet.column(4).width(36);
  sheet.column(5).width(16);
  sheet.column(6).width(12);
  sheet.column(7).width(12);
  sheet.column(8).width(20);
  sheet.column(9).width(16);
  sheet.column(10).width(28);
  sheet.column(11).width(14);
  sheet.column(12).width(20);
  sheet.column(13).width(10);

  losses.people.forEach((person, index) => {
    const excelRow = 2 + index;
    const values = [
      index + 1,
      person.reasonLabel,
      person.rankGroupLabel,
      person.fullName,
      person.rank,
      person.personId,
      person.sourceSheet,
      person.typeOrAbsence,
      person.status,
      person.placeOrNote,
      person.orderDate,
      person.matchNote,
      person.excelRow,
    ];
    values.forEach((value, columnIndex) => {
      sheet.cell(excelRow, columnIndex + 1).value(value);
    });
    sheet.range(excelRow, 1, excelRow, headers.length).style({
      border: BORDER,
      fill: index % 2 === 0 ? VALUE_FILL : "F2F2F2",
      verticalAlignment: "center",
    });
  });
};

const writeDeparturesPeopleSheet = (
  sheet: any,
  result: DeparturesResult,
  sheetName: string,
  filter?: (category: string) => boolean,
) => {
  sheet.name(sheetName);
  const headers = [
    "№",
    "Категорія",
    "Група звання",
    "ПІБ",
    "Звання",
    "ID",
    "Індекс",
    "Підстава / тип",
    "Куди",
    "Дата",
    "№ наказу",
    "Дата наказу",
    "Маркер",
    "Рядок Excel",
  ];
  headers.forEach((header, index) => {
    sheet.cell(1, index + 1).value(header);
    styleHeader(sheet.cell(1, index + 1));
  });
  sheet.column(1).width(6);
  sheet.column(2).width(42);
  sheet.column(3).width(14);
  sheet.column(4).width(36);
  sheet.column(5).width(16);
  sheet.column(6).width(12);
  sheet.column(7).width(12);
  sheet.column(8).width(36);
  sheet.column(9).width(28);
  sheet.column(10).width(12);
  sheet.column(11).width(12);
  sheet.column(12).width(12);
  sheet.column(13).width(28);
  sheet.column(14).width(10);

  const rows = filter
    ? result.people.filter((person) => filter(person.category))
    : result.people;

  rows.forEach((person, index) => {
    const excelRow = 2 + index;
    const values = [
      index + 1,
      person.categoryLabel,
      person.rankGroupLabel,
      person.fullName,
      person.rank,
      person.personId,
      person.positionIndex,
      [person.ground, person.type].filter(Boolean).join(" · "),
      person.destination,
      person.excludeDate,
      person.orderNumber,
      person.orderDate,
      person.matchNote,
      person.excelRow,
    ];
    values.forEach((value, columnIndex) => {
      sheet.cell(excelRow, columnIndex + 1).value(value);
    });
    sheet.range(excelRow, 1, excelRow, headers.length).style({
      border: BORDER,
      fill: index % 2 === 0 ? VALUE_FILL : "F2F2F2",
      verticalAlignment: "center",
    });
  });
};

const writeArrivalsPeopleSheet = (
  sheet: any,
  arrivals: ArrivalsMonthResult,
) => {
  sheet.name(`Прибули ${arrivals.monthLabel}`);
  const headers = [
    "№",
    "Категорія",
    "ПІБ",
    "Звання",
    "ID",
    "Індекс",
    "Звідки прибув",
    "Джерело",
    "Дата зарахування",
    "Рядок Excel",
  ];
  headers.forEach((header, index) => {
    sheet.cell(1, index + 1).value(header);
    styleHeader(sheet.cell(1, index + 1));
  });
  sheet.column(1).width(6);
  sheet.column(2).width(14);
  sheet.column(3).width(36);
  sheet.column(4).width(18);
  sheet.column(5).width(12);
  sheet.column(6).width(14);
  sheet.column(7).width(28);
  sheet.column(8).width(32);
  sheet.column(9).width(16);
  sheet.column(10).width(10);

  arrivals.people.forEach((person, index) => {
    const excelRow = 2 + index;
    const values = [
      index + 1,
      person.rankGroupLabel,
      person.fullName,
      person.rank,
      person.personId,
      person.positionIndex,
      person.arrivedFrom,
      person.arrivalSourceLabel,
      person.enrollDate,
      person.excelRow,
    ];
    values.forEach((value, columnIndex) => {
      sheet.cell(excelRow, columnIndex + 1).value(value);
    });
    sheet.range(excelRow, 1, excelRow, headers.length).style({
      border: BORDER,
      fill: index % 2 === 0 ? VALUE_FILL : "F2F2F2",
      verticalAlignment: "center",
    });
  });
};

export const appendEjoosMovementSheets = (
  workbook: any,
  result: DeparturesResult,
  options?: { useExistingFirstSheet?: boolean },
) => {
  const summarySheet = options?.useExistingFirstSheet
    ? workbook.sheet(0)
    : workbook.addSheet("Вибули");
  writeDeparturesSummarySheet(summarySheet, result);
  writeDeparturesPeopleSheet(
    workbook.addSheet("Звільнення 1-5"),
    result,
    "Звільнення 1-5",
    (category) => category !== "transfer",
  );
  writeDeparturesPeopleSheet(
    workbook.addSheet("Переведені"),
    result,
    "Переведені",
    (category) => category === "transfer",
  );
  writeDeparturesPeopleSheet(
    workbook.addSheet("Усі Виключені"),
    result,
    "Усі Виключені",
  );
  if (result.arrivalsAugust) {
    writeArrivalsPeopleSheet(
      workbook.addSheet("Прибули серпень ООС"),
      result.arrivalsAugust,
    );
  }
  if (result.arrivalsFromMorning) {
    writeArrivalsPeopleSheet(
      workbook.addSheet("Звідки Штатка"),
      result.arrivalsFromMorning,
    );
  }
  if (result.arrivalsAugustPb) {
    writeArrivalsPeopleSheet(
      workbook.addSheet("Прибули серпень 1ПБ"),
      result.arrivalsAugustPb,
    );
  }
  if (result.dispositionFromArchive) {
    writeDispositionPeopleSheet(
      workbook.addSheet("Розпорядження archive"),
      result.dispositionFromArchive,
    );
  }
  if (result.szchFromRuh) {
    writeSzchPeopleSheet(workbook.addSheet("СЗЧ Рух"), result.szchFromRuh);
  }
  if (result.combatLossesFromPb) {
    writeCombatLossesPeopleSheet(
      workbook.addSheet("Втрати"),
      result.combatLossesFromPb,
    );
  }
};

export const exportSocPassportDeparturesWorkbook = async (
  result: DeparturesResult,
) => {
  const stamp = new Date().toISOString().slice(0, 10);
  await exportBlankWorkbookWithMutations((workbook) => {
    appendEjoosMovementSheets(workbook, result, { useExistingFirstSheet: true });
  }, `Соц.паспорт_Вибули_${stamp}.xlsx`);
  return {
    ...result.totals,
    arrivalsTotal: result.arrivalsAugust?.total ?? 0,
    arrivalsMorningTotal: result.arrivalsFromMorning?.total ?? 0,
    arrivalsPbTotal: result.arrivalsAugustPb?.total ?? 0,
    dispositionTotal: result.dispositionFromArchive?.totals.all ?? 0,
    szchTotal: result.szchFromRuh?.totals.all ?? 0,
    combatLossesTotal: result.combatLossesFromPb?.totals.all ?? 0,
  };
};
