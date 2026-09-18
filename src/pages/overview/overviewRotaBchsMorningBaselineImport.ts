import { normalizeBchsText } from "../bchs/bchsCalc";
import { normalizeRosterMatchText } from "../personnel/fighterStatusImport";
import { looksLikePersonnelName } from "../personnel/personnelUtils";
import {
  isMorningLeaveSectionPerson,
  isMorningMissionLocation,
} from "./overviewRotaBchsMorningExport";
import type { MorningReportBucket } from "./overviewRotaBchsMorningStatusChanges";
import {
  formatMorningPersonStateLabel,
  morningReportBucketLabel,
} from "./overviewRotaBchsMorningStatusChanges";
import type {
  BchsMorningDailySnapshot,
  BchsMorningSnapshotPerson,
} from "./overviewRotaBchsMorningSnapshot";
import {
  bchsMorningPersonKeyFromName,
  filterSnapshotPeopleForUnit,
  kyivIsoDateLabel,
} from "./overviewRotaBchsMorningSnapshot";

const LEFT_BLOCK = { start: 2, number: 1 } as const;
const RIGHT_BLOCK = { start: 11, number: 10 } as const;
const EXTRA_BLOCK = { start: 20, number: 19 } as const;

const BLOCKS = [LEFT_BLOCK, RIGHT_BLOCK, EXTRA_BLOCK] as const;

const normalizeText = (value: string) =>
  normalizeRosterMatchText(value).replace(/\s+/g, " ");

const cellText = (sheet: { cell: (row: number, col: number) => { value: () => unknown } }, row: number, col: number) => {
  const value = sheet.cell(row, col).value();
  if (value == null) return "";
  return String(value).trim();
};

export const parseBchsBaselineDateFromFileName = (fileName: string) => {
  const match = fileName.match(/(\d{2})\.(\d{2})\.(\d{4})/);
  if (!match) return null;
  return `${match[3]}-${match[2]}-${match[1]}`;
};

export const inferUnitLabelFromBchsSheetName = (sheetName: string) => {
  const trimmed = sheetName.trim();
  const match = trimmed.match(/^(\d+)ПР$/i);
  if (match) return `${match[1]} піхотна рота`;
  return trimmed;
};

const resolveSectionIdFromTitle = (title: string) => {
  const text = normalizeText(title);
  if (!text || text === "разом") return null;
  if (text.includes("зміна статусу")) return "statusChanges";
  if (text.includes("виконан")) return "mission";
  if (text === "сзч" || text.includes("самовіл")) return "awol";
  if (text.includes("навч") || text.includes("відряд")) return "training";
  if (text.includes("загиб") || text.includes("200")) return "dead";
  if (text.includes("відпуст")) return "leave";
  if (text.includes("шпит") || text.includes("мед. рот") || text.includes("мед рот")) {
    return "hospital";
  }
  if (text.includes("мед. пункт") || text.includes("мед пункт")) return "medPoint";
  if (text.includes("відком") || text.includes("приком")) return "detached";
  if (text.includes("зникл") || text.includes("безв")) return "missing";
  if (text.includes("новоприб")) return "newcomer";
  return "inService";
};

export const morningReportBucketPriority = (bucket: MorningReportBucket) => {
  if (bucket === "inService") return 0;
  if (bucket === "newcomer") return 1;
  return 2;
};

const bucketFromSectionId = (sectionId: string): MorningReportBucket => {
  switch (sectionId) {
    case "mission":
      return "mission";
    case "awol":
      return "awol";
    case "leave":
      return "leave";
    case "hospital":
      return "hospital";
    case "medPoint":
      return "medPoint";
    case "detached":
      return "detached";
    case "missing":
      return "missing";
    case "training":
      return "training";
    case "dead":
      return "dead";
    case "newcomer":
      return "newcomer";
    default:
      return "inService";
  }
};

export const classifyMorningReportBucketFromParsedFields = (
  sectionId: string,
  status: string,
  location: string,
): MorningReportBucket => {
  if (sectionId === "statusChanges") return "inService";
  const sectionBucket = bucketFromSectionId(sectionId);
  if (sectionBucket !== "inService") return sectionBucket;

  const statusNorm = normalizeBchsText(status);
  const locationNorm = normalizeBchsText(location);
  if (statusNorm.includes("сзч") || statusNorm.includes("самовіл")) return "awol";
  if (/загиб|^200$/.test(statusNorm)) return "dead";
  if (/навч|відряд/.test(statusNorm)) return "training";
  if (isMorningLeaveSectionPerson(status, location)) return "leave";
  if (/500|безв/.test(statusNorm) || locationNorm.includes("безв")) return "missing";
  if (statusNorm.includes("відком") || statusNorm.includes("приком")) return "detached";
  if (/шпит|госпітал|мед\.?\s*рот/.test(`${statusNorm} ${locationNorm}`)) {
    return "hospital";
  }
  if (/мед\.?\s*пункт|лік\.?пор|лік\.?відп/.test(`${statusNorm} ${locationNorm}`)) {
    return /шпит|госпітал|мед\.?\s*рот/.test(`${statusNorm} ${locationNorm}`)
      ? "hospital"
      : "medPoint";
  }
  if (isMorningMissionLocation(location) || isMorningMissionLocation(status)) return "mission";
  if (statusNorm.includes("новоприб")) return "newcomer";
  if (statusNorm.includes("в строю")) return "inService";
  return "inService";
};

type ParsedPerson = {
  rank: string;
  name: string;
  callsign: string;
  position: string;
  status: string;
  location: string;
  staffUnit: string;
  sectionId: string;
};

const isKnownHeaderLabel = (value: string) => {
  const text = normalizeText(value);
  return text === "№" || text === "номер" || text === "звання" || text === "піб";
};

const isColumnHeaderRow = (sheet: { cell: (row: number, col: number) => { value: () => unknown } }, row: number, block: typeof LEFT_BLOCK) => {
  const header = normalizeText(cellText(sheet, row, block.start + 1));
  return header === "піб" || normalizeText(cellText(sheet, row, block.start)) === "звання";
};

const resolveSectionTitleRow = (
  sheet: { cell: (row: number, col: number) => { value: () => unknown } },
  row: number,
  block: typeof LEFT_BLOCK,
) => {
  const titleCell = cellText(sheet, row, block.number);
  if (!titleCell || /^\d+$/.test(titleCell) || isKnownHeaderLabel(titleCell)) {
    return null;
  }
  if (!isColumnHeaderRow(sheet, row + 1, block)) return null;
  return resolveSectionIdFromTitle(titleCell);
};

const isTotalRow = (sheet: { cell: (row: number, col: number) => { value: () => unknown } }, row: number, block: typeof LEFT_BLOCK) => {
  const label = normalizeText(cellText(sheet, row, block.number));
  return label.includes("разом");
};

const parsePersonRow = (
  sheet: { cell: (row: number, col: number) => { value: () => unknown } },
  row: number,
  block: typeof LEFT_BLOCK,
  sectionId: string,
): ParsedPerson | null => {
  if (isTotalRow(sheet, row, block)) return null;
  const name = cellText(sheet, row, block.start + 1);
  if (!looksLikePersonnelName(name)) return null;
  return {
    rank: cellText(sheet, row, block.start),
    name,
    callsign: cellText(sheet, row, block.start + 2),
    position: cellText(sheet, row, block.start + 3),
    status: cellText(sheet, row, block.start + 4),
    location: cellText(sheet, row, block.start + 5),
    staffUnit: cellText(sheet, row, block.start + 6),
    sectionId,
  };
};

export const parseBchsMorningBaselineWorkbook = (
  workbook: {
    sheets: () => Array<{ name: () => string; usedRange?: () => { endCell: () => { rowNumber: () => number } } | null; cell: (row: number, col: number) => { value: () => unknown } }>;
    sheet: (name: string) => { usedRange?: () => { endCell: () => { rowNumber: () => number } } | null; cell: (row: number, col: number) => { value: () => unknown } } | undefined;
  },
  options: {
    unitLabel?: string;
    fileName?: string;
    reportDate?: Date;
  } = {},
): BchsMorningDailySnapshot => {
  const sheet =
    (options.unitLabel
      ? workbook.sheet(
          options.unitLabel.match(/(\d+)/) && /рот/i.test(options.unitLabel)
            ? `${options.unitLabel.match(/(\d+)/)![1]}ПР`
            : options.unitLabel.slice(0, 31),
        )
      : null) ??
    workbook.sheets().find((entry) => /^\d/.test(String(entry.name()))) ??
    workbook.sheets()[0];
  if (!sheet) {
    throw new Error("У файлі БЧС немає аркуша для читання.");
  }

  const unitLabel =
    options.unitLabel?.trim() ||
    inferUnitLabelFromBchsSheetName(String(sheet.name()));
  const fileDate =
    (options.fileName ? parseBchsBaselineDateFromFileName(options.fileName) : null) ??
    kyivIsoDateLabel(options.reportDate ?? new Date());

  const blockState = new Map<
    (typeof BLOCKS)[number],
    { sectionId: string; awaitingRows: boolean }
  >();
  for (const block of BLOCKS) {
    blockState.set(block, { sectionId: "inService", awaitingRows: false });
  }

  const parsedPeople: ParsedPerson[] = [];
  const maxRow = sheet.usedRange?.()?.endCell?.()?.rowNumber?.() ?? 500;

  for (let row = 1; row <= maxRow; row += 1) {
    for (const block of BLOCKS) {
      const state = blockState.get(block)!;
      if (isColumnHeaderRow(sheet, row, block)) {
        blockState.set(block, { ...state, awaitingRows: state.sectionId !== "statusChanges" });
        continue;
      }

      const sectionId = resolveSectionTitleRow(sheet, row, block);
      if (sectionId) {
        blockState.set(block, { sectionId, awaitingRows: false });
        continue;
      }

      if (!state.awaitingRows || state.sectionId === "statusChanges") continue;
      const person = parsePersonRow(sheet, row, block, state.sectionId);
      if (person) parsedPeople.push(person);
    }
  }

  if (!parsedPeople.length) {
    throw new Error(
      "Не знайдено жодного рядка з ПІБ. Завантажте попередній експорт «БЧС (ранковий ПБ)».",
    );
  }

  const peopleByKey = new Map<string, BchsMorningSnapshotPerson>();
  for (const person of parsedPeople) {
    const bucket = classifyMorningReportBucketFromParsedFields(
      person.sectionId,
      person.status,
      person.location,
    );
    const stateLabel = formatMorningPersonStateLabel(person, bucket);
    const entry: BchsMorningSnapshotPerson = {
      key: bchsMorningPersonKeyFromName(person.name),
      rank: person.rank,
      name: person.name,
      callsign: person.callsign,
      position: person.position,
      status: person.status || morningReportBucketLabel(bucket),
      location: person.location,
      bucket,
      stateLabel,
      unit: unitLabel,
      staffUnit: person.staffUnit,
      staffCol21: person.status,
      staffCol31: person.location,
    };
    if (!entry.key) continue;
    const existing = peopleByKey.get(entry.key);
    if (
      !existing ||
      morningReportBucketPriority(entry.bucket) >
        morningReportBucketPriority(existing.bucket)
    ) {
      peopleByKey.set(entry.key, entry);
    }
  }
  let people = [...peopleByKey.values()];
  if (unitLabel.trim()) {
    people = filterSnapshotPeopleForUnit(people, unitLabel);
  }

  return {
    date: fileDate,
    unitLabel,
    savedAt: new Date().toISOString(),
    source: "manual",
    baselineKind: "bchs",
    sourceLabel: options.fileName?.trim() || undefined,
    people,
  };
};

export const parseBchsMorningBaselineFile = async (
  file: File,
  unitLabel?: string,
) => {
  const XlsxPopulate = (await import("xlsx-populate")).default;
  const workbook = await XlsxPopulate.fromDataAsync(await file.arrayBuffer());
  return parseBchsMorningBaselineWorkbook(workbook, {
    unitLabel,
    fileName: file.name,
  });
};
