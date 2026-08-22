import { exportBlankWorkbookWithMutations } from "../../excelRoundTrip";
import {
  countsInExitMetrics,
  countsInNoExitsList,
  exitBandLabel,
  hasMorningStatusExitRows,
  isOfficerForLbExits,
  isUkrainianBirthPlace,
  morningStatusExitMatchLabel,
} from "./socPassportFields";
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
    "№ УБД",
    "УБД реєстр",
    "ВПО",
    "Звідки прибув",
    "Ким призваний",
    "Статус (ранок)",
    "Виходів (ранок)",
    "ЖБД (+)",
    "Виходів",
    "Бойове (джерела)",
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
      person.ubdNumber,
      person.ubdRosterStatus === "submitted"
        ? "подавалися"
        : person.ubdRosterStatus === "notSubmitted"
          ? "не подавалися"
          : "",
      person.isIdp ? "так" : "",
      person.arrivedFrom,
      person.calledBy,
      person.morningStatus,
      person.morningExitCount,
      person.jbdExitCount,
      person.exitCount,
      person.combatDutyEvidence.join("; "),
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

const writeStyledPeopleTable = (
  sheet: any,
  headers: string[],
  rows: unknown[][],
  columnWidths?: number[],
) => {
  headers.forEach((header, index) => {
    sheet.cell(1, index + 1).value(header);
    styleHeader(sheet.cell(1, index + 1));
    const width = columnWidths?.[index] ?? (index === 0 ? 36 : 18);
    sheet.column(index + 1).width(width);
  });

  rows.forEach((values, index) => {
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

const SERVICE_TYPE_LABELS: Record<string, string> = {
  mobilized: "Мобілізований",
  contract: "Контракт",
};

const rankGroupLabel = (value: string) => {
  if (value === "officer") return "Офіцер";
  if (value === "sergeant") return "Сержант";
  if (value === "soldier") return "Солдат";
  return value;
};

const serviceTypeLabel = (value: string) =>
  SERVICE_TYPE_LABELS[value] ?? value;

const STATUS_FIGHTERS_HEADER =
  "Є в «Статус бійців»";
const JBD_HEADER = "Є в ЖБД";
const BPLA_HEADER = "БПЛА";

const filterNoExitsPeople = (people: SocPerson[]) =>
  people
    .filter(
      (person) =>
        person.exitBand === "none" &&
        countsInNoExitsList(person) &&
        !person.ubdRosterStatus,
    )
    .sort((left, right) => left.name.localeCompare(right.name, "uk"));

/** Офіцери з ПІБ, без «ТРАНЗИТЕР» (кол. «В якому підрозділі»). */
const filterOfficersForLbExits = (people: SocPerson[]) =>
  people.filter(isOfficerForLbExits);

/** Усі офіцери зі списку (нова + звання). */
const writeOfficersSheet = (sheet: any, people: SocPerson[]) => {
  sheet.name("Офіцери");
  const officers = filterOfficersForLbExits(people).sort((left, right) =>
    left.name.localeCompare(right.name, "uk"),
  );

  const headers = [
    "№",
    "ПІБ",
    "Посада",
    "Індекс посади",
    "Звання",
    "ШПК (факт)",
    "Вид служби",
    "Підрозділ (ранок)",
    "Виходів",
    "Група виходів",
    "Статус (ранок)",
    "В наявності",
    STATUS_FIGHTERS_HEADER,
  ];

  const rows = officers.map((person, index) => [
    index + 1,
    person.name,
    person.position,
    person.positionIndex,
    person.rank,
    person.staffRank,
    serviceTypeLabel(person.serviceType),
    person.morningDestination,
    person.exitCount,
    exitBandLabel(person.exitBand),
    person.morningStatus,
    person.present ? "так" : "ні",
    morningStatusExitMatchLabel(person),
  ]);

  writeStyledPeopleTable(sheet, headers, rows, [
    6, 36, 28, 14, 16, 14, 14, 14, 10, 28, 18, 12, 16,
  ]);

  const withExits = officers.filter((person) => person.exitBand !== "none").length;
  sheet.cell(officers.length + 3, 1).value(
    `Всього офіцерів (ПІБ, без ТРАНЗИТЕР): ${officers.length} · з виходами: ${withExits} · «Не виконували»: ${officers.length - withExits}`,
  ).style({ bold: true, fontColor: TEXT });
};

/** Хто має виходи на ЛБЗ (не потрапив у рядок «Не виконували»). */
const writeExitsPeopleSheet = (sheet: any, people: SocPerson[]) => {
  sheet.name("Виходи ЛБЗ");
  const withExits = people
    .filter((person) => person.exitBand !== "none" && countsInExitMetrics(person))
    .sort((left, right) => right.exitCount - left.exitCount || left.name.localeCompare(right.name, "uk"));

  const headers = [
    "№",
    "ПІБ",
    "Посада",
    "Індекс посади",
    "Звання",
    "Категорія",
    "Вид служби",
    "Виходів",
    "Група (соц.портрет)",
    "Статус (ранок)",
    STATUS_FIGHTERS_HEADER,
    JBD_HEADER,
    BPLA_HEADER,
  ];

  const rows = withExits.map((person, index) => [
    index + 1,
    person.name,
    person.position,
    person.positionIndex,
    person.rank,
    rankGroupLabel(person.rankGroup),
    serviceTypeLabel(person.serviceType),
    person.exitCount,
    exitBandLabel(person.exitBand),
    person.morningStatus,
    morningStatusExitMatchLabel(person),
    person.match.jbdExits ? "так" : "ні",
    person.match.bplaExits ? "так" : "ні",
  ]);

  writeStyledPeopleTable(sheet, headers, rows, [6, 36, 28, 14, 16, 12, 12, 10, 28, 18, 16, 10, 10]);
  sheet.cell(withExits.length + 3, 1).value(`Всього: ${withExits.length}`).style({
    bold: true,
    fontColor: TEXT,
  });
};

/** Хто в рядку «Не виконували» — для звірки. */
const writeNoExitsPeopleSheet = (sheet: any, people: SocPerson[]) => {
  sheet.name("Не виконували");
  const withoutExits = filterNoExitsPeople(people);

  const headers = [
    "№",
    "ПІБ",
    "Позивний",
    "Посада",
    "Звання",
    "Категорія",
    "Вид служби",
    "Статус (ранок)",
    "Виходів (ранок)",
    "ЖБД (+)",
    "Всього виходів",
    STATUS_FIGHTERS_HEADER,
    JBD_HEADER,
    BPLA_HEADER,
    "Примітка",
  ];

  const noExitsRemark = (person: SocPerson) => {
    if (person.morningExitCount > 0 || person.jbdExitCount > 0) {
      return "перевірте підрахунок";
    }
    const parts: string[] = [];
    if (!hasMorningStatusExitRows(person) && !person.match.jbdExits && !person.match.bplaExits) {
      parts.push("немає в «Статус бійців», ЖБД і БПЛА");
    } else if (!hasMorningStatusExitRows(person)) {
      parts.push("немає в «Статус бійців»");
    } else {
      parts.push("0 унікальних дат виходу");
    }
    return parts.join("; ");
  };

  const rows = withoutExits.map((person, index) => [
    index + 1,
    person.name,
    person.callsign,
    person.position,
    person.rank,
    rankGroupLabel(person.rankGroup),
    serviceTypeLabel(person.serviceType),
    person.morningStatus,
    person.morningExitCount,
    person.jbdExitCount,
    person.exitCount,
    morningStatusExitMatchLabel(person),
    person.match.jbdExits ? "так" : "ні",
    person.match.bplaExits ? "так" : "ні",
    noExitsRemark(person),
  ]);

  writeStyledPeopleTable(sheet, headers, rows, [
    6, 36, 16, 28, 16, 12, 12, 18, 14, 10, 12, 16, 10, 10, 34,
  ]);
  sheet.cell(withoutExits.length + 3, 1).value(
    `Всього (без ТРАНЗИТЕР): ${withoutExits.length}`,
  ).style({
    bold: true,
    fontColor: TEXT,
  });
  sheet.cell(withoutExits.length + 4, 1).value(
    `«${STATUS_FIGHTERS_HEADER}» — так, якщо на аркуші ранкового «Статус бійців» є ≥1 дата виходу для цього ПІБ.`,
  ).style({ fontColor: "666666", wrapText: true });
};

/** Хто потрапив у рядок «Рф» (національність). */
const writeRussiaNationalsSheet = (sheet: any, people: SocPerson[]) => {
  sheet.name("Рф");
  const russiaPeople = people
    .filter((person) => person.nationality === "russia")
    .sort((left, right) => left.name.localeCompare(right.name, "uk"));

  const russiaReason = (person: SocPerson) => {
    const fromNotes = person.parseNotes
      .filter((note) => /національність/i.test(note))
      .join("; ");
    if (fromNotes) return fromNotes;
    if (isUkrainianBirthPlace(person.birthPlace)) {
      const snippet = [person.extraRaw, person.relativesRaw]
        .map((value) => value.trim())
        .filter(Boolean)
        .join(" · ")
        .slice(0, 240);
      return snippet || "перевірте текст анкети (родичі / додаткова інформація)";
    }
    if (person.birthPlace.trim()) return `місце народження: ${person.birthPlace.trim()}`;
    return "за текстом анкети (родичі / додаткова інформація)";
  };

  const headers = [
    "№",
    "ПІБ",
    "Посада",
    "Звання",
    "Категорія",
    "Вид служби",
    "Місце народження",
    "Підстава (Рф)",
    "Статус (ранок)",
  ];

  const rows = russiaPeople.map((person, index) => [
    index + 1,
    person.name,
    person.position,
    person.rank,
    rankGroupLabel(person.rankGroup),
    serviceTypeLabel(person.serviceType),
    person.birthPlace,
    russiaReason(person),
    person.morningStatus,
  ]);

  writeStyledPeopleTable(sheet, headers, rows, [
    6, 36, 28, 16, 12, 14, 32, 40, 18,
  ]);
  sheet.cell(russiaPeople.length + 3, 1).value(`Всього (Рф): ${russiaPeople.length}`).style({
    bold: true,
    fontColor: TEXT,
  });
};

export const exportSocPassportWorkbook = async (result: SocPassportResult) => {
  const stamp = new Date().toISOString().slice(0, 10);
  const exitsCount = result.people.filter(
    (person) => person.exitBand !== "none" && countsInExitMetrics(person),
  ).length;
  const noExitsCount = result.people.filter(
    (person) =>
      person.exitBand === "none" &&
      countsInNoExitsList(person) &&
      !person.ubdRosterStatus,
  ).length;
  const officersCount = filterOfficersForLbExits(result.people).length;
  const russiaCount = result.people.filter((person) => person.nationality === "russia").length;
  await exportBlankWorkbookWithMutations((workbook) => {
    const portrait = workbook.sheet(0);
    writePortraitSheet(portrait, result);
    const peopleSheet = workbook.addSheet("Розбір");
    writePeopleSheet(peopleSheet, result.people);
    writeOfficersSheet(workbook.addSheet("Офіцери"), result.people);
    writeRussiaNationalsSheet(workbook.addSheet("Рф"), result.people);
    writeExitsPeopleSheet(workbook.addSheet("Виходи ЛБЗ"), result.people);
    writeNoExitsPeopleSheet(workbook.addSheet("Не виконували"), result.people);
  }, `Соц.паспорт_${stamp}.xlsx`);
  return {
    exitsCount,
    noExitsCount,
    officersCount,
    russiaCount,
  };
};

export const exportSocPassportExitsWorkbook = async (result: SocPassportResult) => {
  const stamp = new Date().toISOString().slice(0, 10);
  const withExits = result.people.filter(
    (person) => person.exitBand !== "none" && countsInExitMetrics(person),
  );
  const noExits = filterNoExitsPeople(result.people);
  const officers = filterOfficersForLbExits(result.people);
  const russiaPeople = result.people.filter((person) => person.nationality === "russia");
  await exportBlankWorkbookWithMutations((workbook) => {
    writeExitsPeopleSheet(workbook.sheet(0), result.people);
    writeNoExitsPeopleSheet(workbook.addSheet("Не виконували"), result.people);
    writeOfficersSheet(workbook.addSheet("Офіцери"), result.people);
    writeRussiaNationalsSheet(workbook.addSheet("Рф"), result.people);
  }, `Соц.паспорт_виходи_ЛБЗ_${stamp}.xlsx`);
  return {
    exitsCount: withExits.length,
    noExitsCount: noExits.length,
    officersCount: officers.length,
    russiaCount: russiaPeople.length,
  };
};

export const exportSocPassportNoExitsWorkbook = async (result: SocPassportResult) => {
  const stamp = new Date().toISOString().slice(0, 10);
  const noExits = filterNoExitsPeople(result.people);
  await exportBlankWorkbookWithMutations((workbook) => {
    writeNoExitsPeopleSheet(workbook.sheet(0), result.people);
  }, `Соц.паспорт_не_виконували_${stamp}.xlsx`);
  return noExits.length;
};

export const exportSocPassportRussiaWorkbook = async (result: SocPassportResult) => {
  const stamp = new Date().toISOString().slice(0, 10);
  const russiaPeople = result.people.filter((person) => person.nationality === "russia");
  await exportBlankWorkbookWithMutations((workbook) => {
    writeRussiaNationalsSheet(workbook.sheet(0), result.people);
  }, `Соц.паспорт_Рф_${stamp}.xlsx`);
  return russiaPeople.length;
};
