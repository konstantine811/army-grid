import { morningUnitMatches } from "./overviewMorningUnit";
import type { SciDataTableExportContext } from "@/components/sci/SciDataTable";
import type { BackendPersonnelOverviewRow } from "../../api";
import { exportTemplateWorkbookWithMutations } from "../../excelRoundTrip";
import {
  isBchsDetachedStatus,
  isBchsWoundedByExcelNote,
  normalizeBchsText,
} from "../bchs/bchsCalc";
import { readRosterColumnValue } from "../excel-fill/rosterSourceSnapshot";
import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import { normalizeRosterMatchText } from "../personnel/fighterStatusImport";
import { buildOverviewPpdLocationRosterLookup } from "./overviewPpdLocationExport";
import {
  cleanPersonDisplayName,
  looksLikePersonnelName,
  looksLikePersonnelRankTitle,
} from "../personnel/personnelUtils";
import {
  isPlatoonCommanderPosition,
  isSectionCommanderPosition,
  isStaffNameColumnShifted,
  resolveRotaGudzCallsign,
  resolveRotaGudzPersonName,
  resolveSelectedOverviewUnit,
} from "./overviewRotaGudzExport";

export const PB_ROTA_BCHS_MORNING_TEMPLATE_URL =
  "/templates/pb-rota-bchs-morning-template.xlsx";

const TEMPLATE_STYLE_ROWS = {
  sectionHeader: 2,
  columnHeader: 3,
  person: 4,
  total: 7,
} as const;

/** Резерв стилів поза областю звіту; прибирається після заповнення. */
const STYLE_BACKUP_ROWS = {
  sectionHeader: 10000,
  columnHeader: 10001,
  person: 10002,
  total: 10003,
} as const;

const LEFT_BLOCK = { start: 2, end: 8, mergeEnd: "H", number: 1 } as const;
const EXTRA_BLOCK = { start: 20, end: 26, mergeEnd: "Z", number: 19 } as const;
const blockForSide = (side: BchsMorningSection["side"]) => side === "left" ? LEFT_BLOCK : side === "right" ? RIGHT_BLOCK : EXTRA_BLOCK;

const RIGHT_BLOCK = { start: 11, end: 17, mergeEnd: "Q", number: 10 } as const;

export type BchsMorningPersonRow = {
  sourceOrder: number;
  rank: string;
  name: string;
  callsign: string;
  position: string;
  status: string;
  combatReadiness: string;
  location: string;
  staffUnit: string;
};

export type BchsMorningSection = {
  id: string;
  title: string;
  side: "left" | "right" | "extra";
  people: BchsMorningPersonRow[];
};

const staffValue = (row: BackendPersonnelOverviewRow, columnNumber: number) =>
  row.staffSheetColumns?.[`staff_${columnNumber}`]?.trim() ?? "";

const normalizeText = (value: string) =>
  normalizeRosterMatchText(value).replace(/\s+/g, " ");

type MorningStaffColumnLookup = (
  row: BackendPersonnelOverviewRow,
  columnNumber: number,
) => string;

const buildMorningStaffColumnLookup = (
  rosterRows?: EjournalPreviewRow[] | null,
): MorningStaffColumnLookup => {
  const rosterLookup = buildOverviewPpdLocationRosterLookup(rosterRows);
  return (row, columnNumber) => {
    const rosterRow = rosterLookup?.(row);
    if (rosterRow) {
      return readRosterColumnValue(rosterRow, columnNumber);
    }
    return staffValue(row, columnNumber);
  };
};

const isMorningInServiceStatus = (status: string) =>
  normalizeBchsText(status).includes("в строю");

const resolveMorningStaffStatus = (
  row: BackendPersonnelOverviewRow,
  readColumn: MorningStaffColumnLookup,
) => readColumn(row, 21) || resolveReportStatus(row);

const isMorningDetachedPerson = (
  row: BackendPersonnelOverviewRow,
  readColumn: MorningStaffColumnLookup,
) => isBchsDetachedStatus(resolveMorningStaffStatus(row, readColumn));

/** Кол. 22 «Тип В\\С» = Забезпечення. */
export const isMorningSupportRoleType = (value: string) =>
  normalizeBchsText(value).includes("забезпеч");

/** Статус U: «Прикомандирований…» (не «Відком. за межі ПБ»). */
export const isMorningAttachedStatus = (status: string) => {
  const normalized = normalizeBchsText(status);
  return normalized.includes("приком") && !isBchsDetachedStatus(status);
};

const isMorningAttachedPerson = (
  row: BackendPersonnelOverviewRow,
  readColumn: MorningStaffColumnLookup,
) =>
  isMorningAttachedStatus(resolveMorningStaffStatus(row, readColumn));

/** Таблиця «Відпустка/Лікувальна відпустка» — лише кол. 21 «Статус» і 31 «Місце перебування». */
export const isMorningLeaveSectionPerson = (
  statusRaw: string,
  locationRaw: string,
) => {
  const status = normalizeBchsText(statusRaw);
  const location = normalizeBchsText(locationRaw);
  if (
    status.includes("відпуст") ||
    /^відп(?:\.|\s|$)/.test(status) ||
    /лік.*відп/.test(status)
  ) {
    return true;
  }
  return location.includes("відпустка лікув");
};

/** Таблиця «ВІДКОМАНДИРОВАНІ»: відкомандировані та прикомандировані за статусом. */
export const isMorningDetachedOrAttachedSectionPerson = (
  row: BackendPersonnelOverviewRow,
  readColumn: MorningStaffColumnLookup,
) =>
  isMorningDetachedPerson(row, readColumn) ||
  isMorningAttachedPerson(row, readColumn);

const isMorningBattleReady = (
  combatReadiness: string,
  location: string,
) => {
  const readiness = normalizeBchsText(combatReadiness);
  const place = normalizeBchsText(location);
  return readiness === "бг" || place === "бг";
};

/** Підрахунок рядків зведення з колонок Штатки (22/23/21), не з розкладки по секціях. */
export const computeBchsMorningSummaryFieldCounts = (
  rows: BackendPersonnelOverviewRow[],
  rosterRows?: EjournalPreviewRow[] | null,
) => {
  const readColumn = buildMorningStaffColumnLookup(rosterRows);
  const listed = rows.filter(isOverviewPersonRow);
  const inService = listed.filter((row) =>
    isMorningInServiceStatus(
      readColumn(row, 21) || resolveReportStatus(row),
    ),
  );
  const management = inService.filter((row) =>
    isMorningManagementPlatoon(readColumn(row, 3)),
  ).length;
  const support = inService.filter((row) =>
    isMorningSupportRoleType(readColumn(row, 22)),
  ).length;
  const battleReady = inService.filter((row) =>
    isMorningBattleReady(
      readColumn(row, 23) || readColumn(row, 42),
      readColumn(row, 31),
    ),
  ).length;
  const platoonLeaders = 0;
  const detached = listed.filter((row) =>
    isMorningDetachedPerson(row, readColumn),
  ).length;
  const attached = listed.filter((row) =>
    isMorningAttachedPerson(row, readColumn),
  ).length;

  return {
    management,
    support,
    platoonLeaders,
    detached,
    attached,
    battleReady,
    inRanks: inService.length,
  };
};

/** Кол. 3 «Взвод» = «ж» → таблиця «Управління» і рядок 19 зведення. */
export const isMorningManagementPlatoon = (platoonRaw: string) =>
  normalizeBchsText(platoonRaw) === "ж";

/** Кол. 31 «Місце перебування» = «На виконанні» / «На виконання». */
export const isMorningMissionLocation = (locationRaw: string) =>
  /^на виконан(?:ня|ні)$/.test(normalizeText(locationRaw));

/** Кол. 5 «Посада» = командир взводу. */
export const isMorningPlatoonLeaderPosition = (positionRaw: string) =>
  isPlatoonCommanderPosition(positionRaw);

/** Кол. 5 «Посада» = командир відділення, зокрема з номером. */
export const isMorningSectionCommanderPosition = (positionRaw: string) =>
  isSectionCommanderPosition(positionRaw) ||
  /командир\s+\d+\s*відділен/i.test(normalizeText(positionRaw));

const resolveMorningSectionNumber = (sectionRaw: string) =>
  sectionRaw.match(/(\d+)/)?.[1] ?? "";

/** Для командира відділення — «Командир N відділення» з кол. 4. */
export const resolveMorningDisplayPosition = (
  row: BackendPersonnelOverviewRow,
  readColumn: MorningStaffColumnLookup = staffValue,
) => {
  const position = readColumn(row, 5);
  if (!isMorningSectionCommanderPosition(position)) return position;
  if (/\d/.test(position) && /відділен/i.test(position)) return position;
  const sectionNumber = resolveMorningSectionNumber(readColumn(row, 4));
  if (!sectionNumber) return position;
  return `Командир ${sectionNumber} відділення`;
};

const platoonTableRoleKey = (person: BchsMorningPersonRow) => {
  if (isMorningPlatoonLeaderPosition(person.position)) return 0;
  if (isMorningSectionCommanderPosition(person.position)) return 1;
  return 2;
};

const platoonTableSectionKey = (position: string) => {
  const match = normalizeText(position).match(/(\d+)\s*відділен/);
  return match ? Number(match[1]) : Number.POSITIVE_INFINITY;
};

const sortPlatoonTablePeople = (people: BchsMorningPersonRow[]) =>
  [...people].sort((left, right) => {
    const roleDelta = platoonTableRoleKey(left) - platoonTableRoleKey(right);
    if (roleDelta !== 0) return roleDelta;
    if (platoonTableRoleKey(left) === 1) {
      const sectionDelta =
        platoonTableSectionKey(left.position) - platoonTableSectionKey(right.position);
      if (sectionDelta !== 0) return sectionDelta;
    }
    return left.sourceOrder - right.sourceOrder;
  });

const LEFT_SECTION_SPECS = [
  { id: "management", title: "Управління" },
  { id: "support", title: "Забезпечення" },
  { id: "platoonLeaders", title: "Взводні" },
  { id: "newcomers", title: "Новоприбулі" },
] as const;

const RIGHT_SECTION_SPECS = [
  {
    id: "hospital",
    title: "Шпиталь/Мед. рота (по хворобі/по пораненню)",
  },
  {
    id: "medPoint",
    title: "Мед. пункт (по хворобі/по пораненню)",
  },
  {
    id: "leave",
    title: "Відпустка/Лікувальна відпустка",
  },
  {
    id: "detached",
    title: "ВІДКОМАНДИРОВАНІ",
  },
  { id: "missing", title: "Зниклі безвісти 500" },
] as const;

const PLATOON_SECTION_MATCHERS = [
  { id: "platoon1", title: "1 Піхотний взвод", pattern: /^1(?:\s|$)/ },
  { id: "platoon2", title: "2 Піхотний взвод", pattern: /^2(?:\s|$)/ },
  { id: "platoon3", title: "3 Піхотний взвод", pattern: /^3(?:\s|$)/ },
  {
    id: "grenade",
    title: "Гранатометний взвод",
    pattern: /гранатомет/i,
  },
  {
    id: "machineGun",
    title: "Кулеметний взвод",
    pattern: /кулемет/i,
  },
] as const;

const isPlatoonSectionId = (sectionId: string) =>
  PLATOON_SECTION_MATCHERS.some((spec) => spec.id === sectionId);

export const resolveMorningPlatoonSectionId = (platoonRaw: string) => {
  if (isMorningManagementPlatoon(platoonRaw)) return null;
  const label = resolveMorningPlatoonLabel(platoonRaw);
  if (!label) return null;
  return (
    PLATOON_SECTION_MATCHERS.find((spec) =>
      spec.pattern.test(normalizeText(label)),
    )?.id ?? null
  );
};

export const splitPlatoonTablePeople = (people: BchsMorningPersonRow[]) => {
  const commanders = people.filter((person) =>
    isMorningPlatoonLeaderPosition(person.position),
  );
  const sectionCommanders = people.filter(
    (person) =>
      isMorningSectionCommanderPosition(person.position) &&
      !isMorningPlatoonLeaderPosition(person.position),
  );
  const rest = people.filter(
    (person) =>
      !isMorningPlatoonLeaderPosition(person.position) &&
      !isMorningSectionCommanderPosition(person.position),
  );
  return {
    commanders: sortPlatoonTablePeople(commanders),
    sectionCommanders: sortPlatoonTablePeople(sectionCommanders),
    rest: rest.sort((left, right) => left.sourceOrder - right.sourceOrder),
  };
};

const columnLetter = (columnNumber: number) => {
  if (columnNumber <= 26) {
    return String.fromCharCode(64 + columnNumber);
  }
  return `A${String.fromCharCode(64 + columnNumber - 26)}`;
};

const copyCellStyle = (sourceCell: any, targetCell: any) => {
  targetCell.style(sourceCell.style(CELL_STYLE_NAMES));
};

const CELL_STYLE_NAMES = [
    "bold",
    "italic",
    "underline",
    "strikethrough",
    "fontSize",
    "fontFamily",
    "fontColor",
    "horizontalAlignment",
    "verticalAlignment",
    "wrapText",
    "shrinkToFit",
    "fill",
    "border",
    "leftBorder",
    "rightBorder",
    "topBorder",
    "bottomBorder",
    "numberFormat",
  ];

const styleSourceColumn = (targetColumn: number) => {
  if (targetColumn === LEFT_BLOCK.number || targetColumn === RIGHT_BLOCK.number || targetColumn === EXTRA_BLOCK.number) {
    return LEFT_BLOCK.start;
  }
  if (targetColumn >= EXTRA_BLOCK.start) return targetColumn - 18;
  return targetColumn;
};

const copyRowStyle = (
  sheet: any,
  sourceRow: number,
  targetRow: number,
  startColumn: number,
  endColumn: number,
) => {
  const height = sheet.row(sourceRow).height();
  if (height != null) sheet.row(targetRow).height(Math.max(sheet.row(targetRow).height() ?? 0, height));
  for (let column = startColumn; column <= endColumn; column += 1) {
    copyCellStyle(
      sheet.cell(sourceRow, styleSourceColumn(column)),
      sheet.cell(targetRow, column),
    );
  }
};

const backupTemplateStyles = (sheet: any) => {
  copyRowStyle(
    sheet,
    TEMPLATE_STYLE_ROWS.sectionHeader,
    STYLE_BACKUP_ROWS.sectionHeader,
    LEFT_BLOCK.start,
    RIGHT_BLOCK.end,
  );
  copyRowStyle(
    sheet,
    TEMPLATE_STYLE_ROWS.columnHeader,
    STYLE_BACKUP_ROWS.columnHeader,
    LEFT_BLOCK.start,
    RIGHT_BLOCK.end,
  );
  copyRowStyle(
    sheet,
    TEMPLATE_STYLE_ROWS.person,
    STYLE_BACKUP_ROWS.person,
    LEFT_BLOCK.start,
    RIGHT_BLOCK.end,
  );
  copyRowStyle(
    sheet,
    TEMPLATE_STYLE_ROWS.total,
    STYLE_BACKUP_ROWS.total,
    LEFT_BLOCK.start,
    RIGHT_BLOCK.end,
  );
};

const styleRow = (kind: keyof typeof STYLE_BACKUP_ROWS) =>
  STYLE_BACKUP_ROWS[kind];

const clearBlockValues = (
  sheet: any,
  startRow: number,
  endRow: number,
  startColumn: number,
  endColumn: number,
) => {
  for (let rowNumber = startRow; rowNumber <= endRow; rowNumber += 1) {
    try {
      sheet
        .range(
          `${columnLetter(startColumn)}${rowNumber}`,
          `${columnLetter(endColumn)}${rowNumber}`,
        )
        .merged(false);
    } catch {
      /* ignore */
    }
    for (let column = startColumn; column <= endColumn; column += 1) {
      sheet.cell(rowNumber, column).value(null);
    }
  }
};

const cellText = (value: string | null | undefined) => {
  const text = String(value ?? "").trim();
  return text || null;
};

export const resolveBchsMorningStaffUnit = (
  unitLabel: string,
  battalionLabel = "1ПБ",
) => {
  const match = unitLabel.match(/(\d+)/);
  const rotaNumber = match?.[1] ?? "";
  const short =
    rotaNumber && /рот/i.test(unitLabel)
      ? `${rotaNumber}ПР`
      : unitLabel.replace(/\s+/g, " ").trim().slice(0, 12);
  return `${battalionLabel} ${short}`.trim();
};

export const resolveBchsMorningSheetName = (unitLabel: string) => {
  const match = unitLabel.match(/(\d+)/);
  if (match && /рот/i.test(unitLabel)) return `${match[1]}ПР`;
  return unitLabel.replace(/\s+/g, " ").trim().slice(0, 31) || "Рота";
};

/** Кол. 5 «Посада», без «Повної посади». */
const resolvePositionText = (row: BackendPersonnelOverviewRow) =>
  resolveMorningDisplayPosition(row);

const sanitizeRankCandidate = (value: string) => {
  const text = value.trim();
  if (!text || /^\d+$/.test(text)) return "";
  if (looksLikePersonnelRankTitle(text)) return text;
  if (looksLikePersonnelName(text)) return "";
  return "";
};

const resolveRankFromOverview = (row: BackendPersonnelOverviewRow) => {
  const shifted = isStaffNameColumnShifted(row);
  const staffRank = staffValue(row, 13);
  const rank = shifted
    ? sanitizeRankCandidate(staffValue(row, 12)) ||
      sanitizeRankCandidate(row.rank?.trim() || "")
    : sanitizeRankCandidate(staffRank) ||
      sanitizeRankCandidate(row.rank?.trim() || "") ||
      "";
  return rank;
};

const resolveOverviewPersonName = (row: BackendPersonnelOverviewRow) => {
  const fromResolver = resolveRotaGudzPersonName(row);
  if (fromResolver) return fromResolver;
  const direct = cleanPersonDisplayName(row.name?.trim() || "");
  if (direct && looksLikePersonnelName(direct)) return direct;
  const fromStaff = cleanPersonDisplayName(staffValue(row, 14));
  if (fromStaff && looksLikePersonnelName(fromStaff)) return fromStaff;
  return "";
};

export const isBchsMorningListedPersonRow = (
  row: BackendPersonnelOverviewRow,
) => {
  const name = resolveOverviewPersonName(row);
  if (!name || name === "Без ПІБ") return false;
  return looksLikePersonnelName(name);
};

/** Ті самі рядки, що потрапляють у таблиці БЧС: обраний підрозділ + ПІБ у списку. */
export const filterBchsMorningUnitRows = (
  rows: BackendPersonnelOverviewRow[],
  unitLabel: string,
) =>
  rows.filter(
    (row) =>
      morningUnitMatches(row.unit?.trim() || "", unitLabel) &&
      isBchsMorningListedPersonRow(row),
  );

const isOverviewPersonRow = isBchsMorningListedPersonRow;

export const resolveReportStatus = (row: BackendPersonnelOverviewRow) => {
  const raw =
    staffValue(row, 21) ||
    row.staffStatusLabel?.trim() ||
    row.statusLabel?.trim() ||
    row.staffStatus?.trim() ||
    "";
  const normalized = normalizeText(raw);
  if (/лік.*відп/.test(normalized)) return "Лік.Відп.";
  if (/відпуст|^відп(?:\s|$)/.test(normalized)) return "Відпустка";
  if (/^лік\s*пор/.test(normalized)) return "Лік.Пор";
  if (/^лік\s*хвор/.test(normalized)) return "Лік.Хвор";
  if (/відком|приком|брез/.test(normalized)) return raw;
  if (/на виход|вик.*бз|на виконан|навч|відряд|загиб|200|500|бг/.test(normalized)) return raw;
  if (normalized.includes("новоприбул")) return "Новоприбулий";
  if (row.status === "MEDICAL" || /ліку|лік\s|шпит|мед\.?\s*пункт/.test(normalized)) {
    const note = `${staffValue(row, 32)} ${staffValue(row, 31)} ${raw}`;
    return isBchsWoundedByExcelNote(note) ? "Лік.Пор" : "Лік.Хвор";
  }
  if (row.status === "LEAVE" || normalized.includes("відпуст")) {
    return "Відпустка";
  }
  if (row.status === "AWOL" || normalized.includes("сзч")) {
    return "СЗЧ";
  }
  if (row.status === "MISSING" || normalized.includes("безв")) {
    return "500";
  }
  if (row.status === "BUSINESS_TRIP" || normalized.includes("відряд")) {
    return "Відрядження";
  }
  return raw || (row.status === "ON_DUTY" ? "В строю" : "—");
};

const resolveReportLocation = (row: BackendPersonnelOverviewRow) => {
  const note = staffValue(row, 32);
  const location = staffValue(row, 31);
  const direction = staffValue(row, 33) || row.fighterDirection?.trim() || "";
  if (row.status === "MEDICAL") return note || location || direction;
  return location || direction || note;
};

export const buildBchsMorningPersonRow = (
  row: BackendPersonnelOverviewRow,
  sourceOrder: number,
  staffUnit: string,
  readColumn: MorningStaffColumnLookup = staffValue,
): BchsMorningPersonRow => ({
  sourceOrder,
  rank: resolveRankFromOverview(row),
  name: resolveOverviewPersonName(row),
  callsign: resolveRotaGudzCallsign(row),
  position: resolveMorningDisplayPosition(row, readColumn),
  status: resolveReportStatus(row),
  combatReadiness: staffValue(row, 23) || staffValue(row, 42),
  location: resolveReportLocation(row),
  staffUnit,
});

const isNewcomer = (row: BackendPersonnelOverviewRow) => {
  const status = normalizeText(
    `${staffValue(row, 21)} ${row.staffStatusLabel ?? ""} ${row.staffStatus ?? ""}`,
  );
  return status.includes("новоприбул");
};

const resolveRowStaffStatus = (row: BackendPersonnelOverviewRow) =>
  staffValue(row, 21) || resolveReportStatus(row);

const isDetachedPerson = (row: BackendPersonnelOverviewRow) =>
  isBchsDetachedStatus(resolveRowStaffStatus(row));

const isAttachedPerson = (row: BackendPersonnelOverviewRow) =>
  isMorningAttachedStatus(resolveRowStaffStatus(row));

const isDetachedOrAttachedPerson = (row: BackendPersonnelOverviewRow) =>
  isDetachedPerson(row) || isAttachedPerson(row);

const isHospitalMedical = (row: BackendPersonnelOverviewRow) => {
  const blob = normalizeText(`${staffValue(row, 31)} ${staffValue(row, 32)}`);
  return /шпит|госпітал|мед\.?\s*рот|цмкл|вмг|лікарн/.test(blob);
};

/** Кол. 3 «Взвод» — єдине джерело для таблиць взводів (не col 4 «Відділення»). */
export const resolveMorningPlatoonLabel = (
  platoonRaw: string,
) => {
  const raw = platoonRaw.trim();
  if (!raw) return "";
  const match = raw.match(/(\d+)\s*(?:піхотн\S*\s+)?взвод/i);
  if (match) return `${match[1]} Піхотний взвод`;
  if (/гранатомет/i.test(raw)) return "Гранатометний взвод";
  if (/кулемет/i.test(raw)) return "Кулеметний взвод";
  return raw;
};

/** СЗЧ — окрема таблиця, не взвод і не «На виконанні». */
export const isMorningAwolPerson = (
  row: BackendPersonnelOverviewRow,
  readColumn: MorningStaffColumnLookup = staffValue,
) => {
  const staffStatus = normalizeText(readColumn(row, 21));
  const reportStatus = normalizeText(resolveReportStatus(row));
  return (
    row.status === "AWOL" ||
    staffStatus.includes("сзч") ||
    staffStatus.includes("самовіл") ||
    reportStatus.includes("сзч")
  );
};

const classifyLeftSectionId = (
  row: BackendPersonnelOverviewRow,
  readColumn: MorningStaffColumnLookup = staffValue,
) => {
  if (isNewcomer(row)) return "newcomers";
  if (isMorningManagementPlatoon(readColumn(row, 3))) return "management";
  if (isMorningSupportRoleType(readColumn(row, 22))) return "support";
  const platoonLabel = resolveMorningPlatoonLabel(readColumn(row, 3));
  const platoonSection = PLATOON_SECTION_MATCHERS.find((spec) =>
    spec.pattern.test(normalizeText(platoonLabel)),
  );
  if (platoonSection) return platoonSection.id;
  return "platoon1";
};

const classifyRightSectionId = (
  row: BackendPersonnelOverviewRow,
  readColumn: MorningStaffColumnLookup = staffValue,
) => {
  const staffStatus = readColumn(row, 21);
  const staffLocation = readColumn(row, 31);
  if (isMorningLeaveSectionPerson(staffStatus, staffLocation)) {
    return "leave";
  }
  const reportStatus = normalizeText(resolveReportStatus(row));
  if (/500|безв/.test(reportStatus) || row.status === "MISSING") return "missing";
  if (isDetachedOrAttachedPerson(row)) return "detached";
  if (row.status === "MEDICAL") {
    return isHospitalMedical(row) ? "hospital" : "medPoint";
  }
  const status = normalizeText(
    `${staffStatus} ${row.statusLabel ?? ""} ${staffValue(row, 32)}`,
  );
  if (/ліку|лік\s|шпит|мед\.?\s*пункт/.test(status)) {
    return isHospitalMedical(row) ? "hospital" : "medPoint";
  }
  if (status.includes("безв")) return "missing";
  return null;
};

const EXTRA_SECTION_SPECS = [
  { id: "mission", title: "На виконанні" },
  { id: "awol", title: "СЗЧ" },
  { id: "statusChanges", title: "Зміна статусу" },
  { id: "training", title: "Навчання/відрядження" },
  { id: "dead", title: "Загиблі 200" },
];

const STATUS_CHANGE_HEADERS = [
  "№",
  "Звання",
  "ПІБ",
  "Позивний",
  "Посада",
  "Було",
  "Стало",
  "Зміна",
] as const;

/** Було / Стало / Зміна — довгі описи статусів і місць. */
const STATUS_CHANGE_NOTE_COLUMN_WIDTHS = [16, 20, 26] as const;

const estimateWrappedLineCount = (value: unknown, columnWidth: number) => {
  const text = String(value ?? "").trim();
  if (!text) return 1;
  const byNewline = text.split(/\n/).length;
  const wrapped = Math.ceil(text.length / Math.max(10, columnWidth));
  return Math.max(byNewline, wrapped);
};

const applyStatusChangeRowLayout = (
  sheet: any,
  rowNumber: number,
  startColumn: number,
  person: BchsMorningPersonRow,
) => {
  const noteColumns = STATUS_CHANGE_NOTE_COLUMN_WIDTHS.map(
    (width, index) => ({ column: startColumn + 4 + index, width }),
  );
  for (const { column, width } of noteColumns) {
    sheet.column(column).width(width);
    sheet.cell(rowNumber, column).style({
      wrapText: true,
      shrinkToFit: false,
      verticalAlignment: "top",
    });
  }
  sheet.cell(rowNumber, startColumn + 3).style({
    wrapText: true,
    shrinkToFit: false,
    verticalAlignment: "top",
  });
  const lineCount = Math.max(
    estimateWrappedLineCount(person.position, 24),
    ...noteColumns.map(({ width }, index) => {
      const value =
        index === 0
          ? person.status
          : index === 1
            ? person.location
            : person.staffUnit;
      return estimateWrappedLineCount(value, width);
    }),
  );
  sheet.row(rowNumber).height(Math.min(144, Math.max(22, lineCount * 15)));
};

const classifyExtraSectionId = (status: string) => {
  const text = normalizeText(status);
  if (/загиб|^200$/.test(text)) return "dead";
  if (/навч|відряд/.test(text)) return "training";
  return null;
};

export const buildBchsMorningSections = (
  rows: BackendPersonnelOverviewRow[],
  unitLabel: string,
  rosterRows?: EjournalPreviewRow[] | null,
): BchsMorningSection[] => {
  const staffUnit = resolveBchsMorningStaffUnit(unitLabel);
  const readColumn = buildMorningStaffColumnLookup(rosterRows);

  const extraBuckets = new Map(EXTRA_SECTION_SPECS.map(spec => [spec.id, [] as BchsMorningPersonRow[]]));
  const leftBuckets = new Map<string, BchsMorningPersonRow[]>();
  const rightBuckets = new Map<string, BchsMorningPersonRow[]>();

  for (const spec of LEFT_SECTION_SPECS) leftBuckets.set(spec.id, []);
  for (const spec of PLATOON_SECTION_MATCHERS) leftBuckets.set(spec.id, []);
  for (const spec of RIGHT_SECTION_SPECS) rightBuckets.set(spec.id, []);

  rows.forEach((sourceRow, index) => {
    if (!isOverviewPersonRow(sourceRow)) return;
    const reportPerson = buildBchsMorningPersonRow(sourceRow, index, staffUnit, readColumn);
    if (isMorningAwolPerson(sourceRow, readColumn)) {
      extraBuckets.get("awol")?.push(reportPerson);
      return;
    }
    const extraSection = classifyExtraSectionId(reportPerson.status);
    const isPlatoonLeader = isMorningPlatoonLeaderPosition(reportPerson.position);
    const assignCommanderToPlatoon = () => {
      if (
        !isPlatoonLeader &&
        !isMorningSectionCommanderPosition(reportPerson.position)
      ) {
        return;
      }
      const platoonId = resolveMorningPlatoonSectionId(readColumn(sourceRow, 3));
      if (!platoonId) return;
      const bucket = leftBuckets.get(platoonId);
      if (
        bucket &&
        !bucket.some((person) => person.sourceOrder === reportPerson.sourceOrder)
      ) {
        bucket.push(reportPerson);
      }
    };
    if (extraSection) {
      extraBuckets.get(extraSection)?.push(reportPerson);
      if (isPlatoonLeader && extraSection !== "dead") {
        leftBuckets.get("platoonLeaders")?.push(reportPerson);
      }
      assignCommanderToPlatoon();
      return;
    }
    const rightSection = classifyRightSectionId(sourceRow, readColumn);
    if (rightSection) {
      rightBuckets.get(rightSection)?.push(reportPerson);
      if (isPlatoonLeader) leftBuckets.get("platoonLeaders")?.push(reportPerson);
      assignCommanderToPlatoon();
      return;
    }
    const leftSection = classifyLeftSectionId(sourceRow, readColumn);
    // Mission assignment comes from «Місце перебування», not «Статус» or notes.
    // Absences retain their own section even when an old location is still set.
    // Exclusive left tables: Управління ⊃ Забезпечення ⊃ взвод.
    if (isMorningMissionLocation(readColumn(sourceRow, 31))) {
      extraBuckets.get("mission")?.push(reportPerson);
      if (leftSection === "support") {
        leftBuckets.get("support")?.push(reportPerson);
      } else if (isPlatoonSectionId(leftSection)) {
        leftBuckets.get(leftSection)?.push(reportPerson);
      }
      if (isPlatoonLeader) leftBuckets.get("platoonLeaders")?.push(reportPerson);
      assignCommanderToPlatoon();
      return;
    }
    leftBuckets.get(leftSection)?.push(reportPerson);
    if (isPlatoonLeader && leftSection !== "platoonLeaders") {
      leftBuckets.get("platoonLeaders")?.push(reportPerson);
    }
    assignCommanderToPlatoon();
  });

  const sections: BchsMorningSection[] = [];

  for (const spec of [...LEFT_SECTION_SPECS.slice(0, 3), ...PLATOON_SECTION_MATCHERS, ...LEFT_SECTION_SPECS.slice(3)]) {
    const bucket = leftBuckets.get(spec.id) ?? [];
    if (!bucket.length) continue;
    sections.push({
      id: spec.id,
      title: spec.title,
      side: "left",
      people: isPlatoonSectionId(spec.id)
        ? sortPlatoonTablePeople(bucket)
        : bucket.sort((a, b) => a.sourceOrder - b.sourceOrder),
    });
  }

  for (const spec of RIGHT_SECTION_SPECS) {
    const bucket = rightBuckets.get(spec.id) ?? [];
    if (!bucket.length) continue;
    sections.push({
      id: spec.id,
      title: spec.title,
      side: "right",
      people: bucket.sort((a, b) => a.sourceOrder - b.sourceOrder),
    });
  }

  for (const spec of EXTRA_SECTION_SPECS) {
    if (spec.id === "statusChanges") continue;
    const people = extraBuckets.get(spec.id) ?? [];
    // Keep mission and СЗЧ visible even when the current count is zero.
    if (people.length || spec.id === "mission" || spec.id === "awol") {
      sections.push({ ...spec, side: "extra", people });
    }
  }
  return sections;
};

export const appendBchsMorningStatusChangeSection = (
  sections: BchsMorningSection[],
  statusChangeSection: BchsMorningSection,
) => {
  const withoutStatusChanges = sections.filter(
    (section) => section.id !== "statusChanges",
  );
  const missionIndex = withoutStatusChanges.findIndex(
    (section) => section.id === "awol",
  );
  const insertAt = missionIndex >= 0 ? missionIndex + 1 : withoutStatusChanges.length;
  return [
    ...withoutStatusChanges.slice(0, insertAt),
    statusChangeSection,
    ...withoutStatusChanges.slice(insertAt),
  ];
};

const NUMBER_COLUMN_WIDTH = 6;

const styleNumberCell = (sheet: any, rowNumber: number, column: number) => {
  sheet.cell(rowNumber, column).style({
    numberFormat: "General",
    wrapText: false,
    shrinkToFit: false,
    horizontalAlignment: "center",
  });
};

const writePersonCells = (
  sheet: any,
  rowNumber: number,
  startColumn: number,
  person: BchsMorningPersonRow,
  index: number,
  options?: { statusChangeLayout?: boolean },
) => {
  sheet.cell(rowNumber, startColumn - 1).value(index);
  styleNumberCell(sheet, rowNumber, startColumn - 1);
  sheet.cell(rowNumber, startColumn).value(cellText(person.rank));
  sheet.cell(rowNumber, startColumn + 1).value(cellText(person.name));
  sheet.cell(rowNumber, startColumn + 2).value(cellText(person.callsign));
  sheet.cell(rowNumber, startColumn + 3).value(cellText(person.position));
  sheet.cell(rowNumber, startColumn + 4).value(cellText(person.status));
  sheet.cell(rowNumber, startColumn + 5).value(cellText(person.location));
  sheet.cell(rowNumber, startColumn + 6).value(cellText(person.staffUnit));
  if (options?.statusChangeLayout) {
    applyStatusChangeRowLayout(sheet, rowNumber, startColumn, person);
  }
};

const writeSectionHeader = (
  sheet: any,
  rowNumber: number,
  side: "left" | "right" | "extra",
  title: string,
) => {
  const block = blockForSide(side);
  copyRowStyle(sheet, styleRow("sectionHeader"), rowNumber, block.number, block.end);
  clearBlockValues(sheet, rowNumber, rowNumber, block.number, block.end);
  try {
    sheet
      .range(
        `${columnLetter(block.number)}${rowNumber}`,
        `${block.mergeEnd}${rowNumber}`,
      )
      .merged(true);
  } catch {
    /* ignore */
  }
  sheet.cell(rowNumber, block.number).value(title);
};

const writeColumnHeaderRow = (
  sheet: any,
  rowNumber: number,
  side: "left" | "right" | "extra",
  headers: readonly string[] = [
    "№",
    "Звання",
    "ПІБ",
    "Позивний",
    "Посада",
    "Статус",
    "Місце перебування",
    "ШТАТ",
  ],
) => {
  const block = blockForSide(side);
  copyRowStyle(sheet, styleRow("columnHeader"), rowNumber, block.number, block.end);
  headers.forEach((header, index) => {
    sheet.cell(rowNumber, block.number + index).value(header);
  });
  styleNumberCell(sheet, rowNumber, block.number);
};

const isAttached = (person: BchsMorningPersonRow) => /приком/.test(normalizeText(person.status));

const writeTotalRow = (sheet: any, rowNumber: number, section: BchsMorningSection) => {
  const block = blockForSide(section.side);
  // The reference uses the same bold, bordered total style in all three blocks.
  copyCellStyle(sheet.cell(styleRow("total"), LEFT_BLOCK.start), sheet.cell(rowNumber, block.number));
  for (let offset = 0; offset < 7; offset += 1) {
    copyCellStyle(sheet.cell(styleRow("total"), LEFT_BLOCK.start + offset), sheet.cell(rowNumber, block.start + offset));
  }
  clearBlockValues(sheet, rowNumber, rowNumber, block.number, block.end);
  const total = section.people.length;
  let breakdown: [string, number, string, number] | undefined;
  if (section.id === "hospital" || section.id === "medPoint") {
    const wounded = section.people.filter(person => person.status === "Лік.Пор").length;
    breakdown = ["Лікування по хворобі", total - wounded, "Лікування по пораненню", wounded];
  } else if (section.id === "leave") {
    const medicalLeave = section.people.filter(person => person.status === "Лік.Відп.").length;
    breakdown = ["Лікувальна відпустка", medicalLeave, "Відпустка", total - medicalLeave];
  } else if (section.id === "detached") {
    const attached = section.people.filter(isAttached).length;
    breakdown = ["Прикомандировані", attached, "Відкомандировані", total - attached];
  } else if (section.id === "training") {
    const training = section.people.filter(person => /навч/i.test(person.status)).length;
    breakdown = ["Навчання", training, "Відрядження", total - training];
  }
  if (breakdown) {
    [...breakdown, "Загально", total].forEach((value, index) => sheet.cell(rowNumber, block.start + 1 + index).value(value));
    sheet.range(rowNumber, block.start, rowNumber, block.end).style({ bold: true, wrapText: true });
    sheet.row(rowNumber).height(Math.max(sheet.row(rowNumber).height() ?? 0, 45));
  } else {
    sheet.cell(rowNumber, block.start).value("Всього:");
    sheet.cell(rowNumber, block.end).value(total);
  }
};

const writeSideSections = (
  sheet: any,
  sections: BchsMorningSection[],
  side: "left" | "right" | "extra",
  startRow = 2,
) => {
  let rowNumber = startRow;
  const sideSections = sections.filter((entry) => entry.side === side);


  const writePersonRows = (
    people: BchsMorningPersonRow[],
    startIndex: number,
    sectionId?: string,
  ) => {
    const block = blockForSide(side);
    const statusChangeLayout = sectionId === "statusChanges";
    for (const [index, person] of people.entries()) {
      clearBlockValues(sheet, rowNumber, rowNumber, block.number, block.end);
      copyRowStyle(sheet, styleRow("person"), rowNumber, block.number, block.end);
      writePersonCells(sheet, rowNumber, block.start, person, startIndex + index + 1, {
        statusChangeLayout,
      });
      rowNumber += 1;
    }
    return people.length;
  };

  for (const section of sideSections) {
    writeSectionHeader(sheet, rowNumber, side, section.title);
    rowNumber += 1;
    writeColumnHeaderRow(
      sheet,
      rowNumber,
      side,
      section.id === "statusChanges" ? STATUS_CHANGE_HEADERS : undefined,
    );
    rowNumber += 1;
    if (isPlatoonSectionId(section.id)) {
      const groups = splitPlatoonTablePeople(section.people);
      writeSectionHeader(sheet, rowNumber, side, "Командир взводу");
      rowNumber += 1;
      let nextNumber = writePersonRows(groups.commanders, 0, section.id);
      writeSectionHeader(sheet, rowNumber, side, "Командир відділення");
      rowNumber += 1;
      nextNumber += writePersonRows(groups.sectionCommanders, nextNumber, section.id);
      writePersonRows(groups.rest, nextNumber, section.id);
    } else {
      writePersonRows(section.people, 0, section.id);
    }
    writeTotalRow(sheet, rowNumber, section);
    rowNumber += 2;
  }

  return rowNumber;
};

export const computeBchsMorningSummaryFromSections = (
  sections: BchsMorningSection[],
  rows: BackendPersonnelOverviewRow[],
  staffCount: number,
  rosterRows?: EjournalPreviewRow[] | null,
) => {
  const people = (id: string) =>
    sections.find((section) => section.id === id)?.people ?? [];
  const count = (id: string) => people(id).length;
  const listed = new Set(
    sections.flatMap((section) =>
      section.people.map((person) => person.sourceOrder),
    ),
  ).size;
  const fieldCounts = computeBchsMorningSummaryFieldCounts(rows, rosterRows);
  const detached = fieldCounts.detached;
  const treatment = count("hospital") + count("medPoint");
  const absent =
    count("training") +
    detached +
    treatment +
    count("leave") +
    count("awol") +
    count("missing") +
    count("dead");
  const awayOrders = new Set(
    sections
      .filter((section) => section.side !== "left")
      .flatMap((section) => section.people.map((person) => person.sourceOrder)),
  );
  const remaining = sections
    .filter(
      (section) =>
        section.side === "left" &&
        !["management", "support", "platoonLeaders"].includes(section.id),
    )
    .flatMap((section) => section.people)
    .filter((person) => !awayOrders.has(person.sourceOrder));
  const readyInRemaining = remaining.filter(
    (person) => normalizeText(person.combatReadiness) === "бг",
  ).length;

  return {
    staff: staffCount,
    listed,
    trainingTrip: count("training"),
    detached,
    treatment,
    treatmentPeople: [...people("hospital"), ...people("medPoint")],
    vacation: count("leave"),
    awol: count("awol"),
    missing: count("missing"),
    killed: count("dead"),
    absent,
    management: fieldCounts.management,
    support: fieldCounts.support,
    platoon: people("platoonLeaders").filter((person) =>
      isMorningInServiceStatus(person.status),
    ).length,
    attached: fieldCounts.attached,
    onExit: count("mission"),
    battleReady: fieldCounts.battleReady,
    available: remaining.length - readyInRemaining,
    inRanks: fieldCounts.inRanks,
  };
};

export const writeMorningSummary = (
  sheet: any,
  sections: BchsMorningSection[],
  options: {
    unitLabel: string;
    rows: BackendPersonnelOverviewRow[];
    staffCount: number;
    rosterRows?: EjournalPreviewRow[] | null;
  },
) => {
  const totals = computeBchsMorningSummaryFromSections(
    sections,
    options.rows,
    options.staffCount,
    options.rosterRows,
  );
  const values: Record<number, number> = {
    6: totals.staff, 7: totals.listed, 9: totals.trainingTrip, 10: totals.detached,
    11: totals.treatment, 12: totals.vacation, 13: totals.awol,
    14: totals.missing, 16: totals.killed, 17: totals.absent,
    19: totals.management, 20: totals.support,
    21: totals.platoon,
    23: totals.onExit, 25: totals.battleReady,
    26: totals.available, 27: totals.inRanks,
  };
  const labels: Record<number, string> = {5: options.unitLabel, 6: "За штатом", 7: "Всього по списку", 8: "З них відсутні:", 9: "Навчання/відрядження", 10: "Відкомандировані", 11: "Лікування", 12: "Відпустка/Лікувальна відпустка", 13: "СЗЧ", 14: "Зниклі безвісти", 16: "Загиблі", 17: "Всього відсутніх:", 18: "З них в строю", 19: "Управління", 20: "Забезпечення", 21: "Взводні (наявність)", 23: "На виконанні", 25: "БГ", 26: "В наявності", 27: "Всього в строю:"};
  for (const [row, label] of Object.entries(labels)) sheet.cell(Number(row), 21).value(label);
  for (const [row, value] of Object.entries(values)) sheet.cell(Number(row), 22).value(value);
  sheet.cell(22, 21).value(null);
  sheet.cell(22, 22).value(null);
};

export const writeRotaBchsMorningWorkbook = (
  workbook: any,
  options: {
    unitLabel: string;
    rows: BackendPersonnelOverviewRow[];
    staffCount: number;
    rosterRows?: EjournalPreviewRow[] | null;
    sections?: BchsMorningSection[];
  },
) => {
  const sheetName = resolveBchsMorningSheetName(options.unitLabel);
  const sheet =
    workbook.sheet(sheetName) ??
    workbook.sheets().find((entry: any) => /^\d/.test(String(entry.name()))) ??
    workbook.sheet(0) ??
    workbook.sheets()[0];
  if (!sheet) {
    throw new Error("У шаблоні БЧС немає аркуша для заповнення.");
  }
  // xlsx-populate keeps scoped names when deleting a sheet; they serialize with
  // localSheetId=-1. This report rebuilds all content, so discard template names
  // before renaming/removing sheets and create only the final print area below.
  const definedNames = workbook._node.children.find((node: any) => node.name === "definedNames");
  for (const node of [...(definedNames?.children ?? [])]) {
    workbook.scopedDefinedName(node.localSheet, node.attributes.name, null);
  }
  sheet.name(sheetName);
  sheet.autoFilter(null);

  const sections =
    options.sections ??
    buildBchsMorningSections(
      options.rows,
      options.unitLabel,
      options.rosterRows,
    );
  if (!sections.some(section => section.people.length)) throw new Error("Немає жодної особи з ПІБ для заповнення БЧС.");
  const extent = Math.max(sheet.usedRange()?.endCell().rowNumber() ?? 248, options.rows.length + 100);
  const summaryStyles = Array.from({ length: 23 }, (_, index) =>
    [21, 22].map(column => sheet.cell(index + 5, column).style(CELL_STYLE_NAMES)),
  );
  backupTemplateStyles(sheet);
  // Remove template merges before placing sections at their new row positions.
  for (const address of Object.keys(sheet._mergeCells ?? {})) sheet.range(address).merged(false);
  clearBlockValues(sheet, 2, extent, 1, 29);
  sheet.range(`A2:AC${extent}`).style({ border: false, fill: "FFFFFF" });
  for (const block of [LEFT_BLOCK, RIGHT_BLOCK, EXTRA_BLOCK]) {
    sheet.column(block.number).width(NUMBER_COLUMN_WIDTH);
  }
  const leftEndRow = writeSideSections(sheet, sections, "left");
  const rightEndRow = writeSideSections(sheet, sections, "right");
  const extraEndRow = writeSideSections(sheet, sections, "extra", 34);
  summaryStyles.forEach((styles, index) => styles.forEach((style, column) => sheet.cell(index + 5, column + 21).style(style)));
  for (const rowNumber of [5, 8, 18]) sheet.range(`U${rowNumber}:V${rowNumber}`).merged(true);
  writeMorningSummary(sheet, sections, options);
  for (const rowNumber of Object.values(STYLE_BACKUP_ROWS)) delete sheet._rows[rowNumber];
  for (const otherSheet of [...workbook.sheets()]) {
    if (otherSheet !== sheet) otherSheet.delete();
  }
  sheet.definedName(
    "_xlnm.Print_Area",
    sheet.range(`A1:AA${Math.max(leftEndRow, rightEndRow, extraEndRow, 28)}`),
  );
};

export const buildRotaBchsMorningExportFileName = (
  unitLabel: string,
  reportDate = new Date(),
) => {
  const dateLabel = new Intl.DateTimeFormat("uk-UA", {
    day: "2-digit",
    month: "2-digit",
    year: "numeric",
  }).format(reportDate);
  const short = resolveBchsMorningSheetName(unitLabel);
  return `ПБ ${short} БЧС ${dateLabel}.xlsx`;
};

export const exportOverviewRotaBchsMorningReport = async (
  context: SciDataTableExportContext<BackendPersonnelOverviewRow>,
  staffCount: number,
  rosterRows?: EjournalPreviewRow[] | null,
  options?: {
    comparisonSnapshot?: import("./overviewRotaBchsMorningSnapshot").BchsMorningDailySnapshot | null;
  },
) => {
  const unitLabel = resolveSelectedOverviewUnit(context.filters);
  if (!unitLabel) {
    throw new Error(
      "Оберіть одну роту у фільтрі колонки «Підрозділ» перед експортом БЧС.",
    );
  }

  const rows = filterBchsMorningUnitRows(
    context.allRows ?? context.rows,
    unitLabel,
  );
  if (!rows.length) {
    throw new Error(`Немає осіб для роти «${unitLabel}».`);
  }

  const reportDate = new Date();
  const [
    { buildBchsMorningStatusChangeSection, buildBchsMorningDailySnapshot },
    { readBchsMorningComparisonSnapshot, saveBchsMorningDailySnapshot },
  ] = await Promise.all([
    import("./overviewRotaBchsMorningStatusChanges"),
    import("./overviewRotaBchsMorningSnapshot"),
  ]);
  const previousSnapshot =
    options?.comparisonSnapshot ??
    (await readBchsMorningComparisonSnapshot(unitLabel, reportDate));
  const baseSections = buildBchsMorningSections(rows, unitLabel, rosterRows);
  const sections = appendBchsMorningStatusChangeSection(
    baseSections,
    buildBchsMorningStatusChangeSection(
      rows,
      previousSnapshot,
      unitLabel,
      rosterRows,
    ),
  );

  await exportTemplateWorkbookWithMutations(
    PB_ROTA_BCHS_MORNING_TEMPLATE_URL,
    (workbook) => {
      writeRotaBchsMorningWorkbook(workbook, {
        unitLabel,
        rows,
        staffCount,
        rosterRows,
        sections,
      });
    },
    buildRotaBchsMorningExportFileName(unitLabel, reportDate),
  );

  await saveBchsMorningDailySnapshot(
    buildBchsMorningDailySnapshot(rows, unitLabel, rosterRows, reportDate),
  );
};
