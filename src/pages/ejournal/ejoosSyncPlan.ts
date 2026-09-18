import type {
  CellValue,
  ExcelSheetSnapshot,
  ExcelWorkbookSnapshot,
} from "../../excelRoundTrip";
import {
  mapPbStatusToEjoosWithRules,
  readOperatorSettings,
} from "./ejoosStatusMap";
import type { EjoosTimesheetCode } from "./ejoosRules";
import {
  findLatestPriorOwnUnitStaffMove,
  isAbsenceOnlyMovement,
  isAmbiguousStaffTransfer,
  isDispositionAbsenceStatus,
  isDispositionToStaffPlacement,
  isInternalStaffIndexHop,
  isOutboundStaffMove,
  isOwnFirstPbDestination,
  isOwnUnitStaffMove,
  resolveOutboundTransferDestination,
} from "./ejoosMovementRules";
import {
  resolveExistingTimesheetStartDay,
  resolveTimesheetArrivalDate,
  resolveTimesheetEpisodeStart,
} from "./ejoosTimesheetEpisode";
import {
  absenceOnlyBlocksExclusion,
  excludedTimesheetWrite,
  excludedRowsToClear,
  externalTransferProcessState,
  isFalseInternalHopExclusion,
  staleExcludedRowsToClear,
  isStaleVacatedAbsence,
  isUnrecordedSameMonthTransit,
  laterReturnSupersedesOutbound,
  ownUnitMoveSupersededByOutbound,
  positionCloseWritesExcluded,
  skipExternalIfAlreadyProcessed,
} from "./ejoosExcludePolicy";
import {
  absenceSpansBeforeEpisode,
  archivePeriodTouchesJournalMonth,
  buildTimesheetAbsenceSpans,
  clipAbsenceSpansToActiveEpisode,
  currentStatusConfirmsOpenAbsence,
  encodeTimesheetAbsenceSpans,
  archiveReturnContradictsCurrentSh,
  extendDispositionSpanToReportDay,
  findTimesheetMonthHeaderCell,
  formatDispositionTimesheetDeparture,
  formatTimesheetMonthHeader,
  isTimesheetAbsenceCode,
  isTimesheetDepartureMark,
  journalDayFromDateMs,
  sameTimesheetDayMark,
  timesheetHorizonFillDays,
  timesheetMarkFromArchive,
  extractTimesheetDestinationFromPosition,
} from "./ejoosTimesheetText";
import {
  findDuplicateTimesheetExtras,
  isClosedTimesheetHistoryRow,
} from "./ejoosTimesheetDuplicates";
import {
  timesheetRowInExpectedUnitSection,
  extractUnitPhrasesFromPosition,
} from "./ejoosTimesheetUnitSections";
import { findDuplicateOosById } from "./ejoosOosText";
import {
  canonicalName,
  dateMs,
  isJournalPersonId,
  norm,
  normId,
  normKey,
} from "./ejoos/parse/cellText";
import {
  cell,
  findCol,
  findEjoosSheet,
  headerMap,
  idCell,
  journalIdCell,
} from "./ejoos/parse/sheetLookup";
import {
  collectProcessedMovementKeys,
  createMovementKey,
  isCancelledMovementRecord,
  isContractMovementType,
  isSzchCancellation,
  isTransferCancellation,
  motivationContractOverlapsWindow,
  parseContractDatesFromChangeText,
  parsePbArchive,
  parsePbMovements,
  parsePbShPeople,
  parseRankPromotion,
} from "./ejoos/parse/pb";
import {
  parseAsOfDateLabel,
  parseTimesheetDayFromPbName,
  resolveJournalTimesheetDay,
  sourceTimesheetHorizonNote,
} from "./ejoos/plan/journalDay";
import {
  MONTH_ROLLOVER_BLOCK_MESSAGE,
  planBlocksWorkbookApply,
  refreshPlanTimesheetHorizon,
  SOURCE_DATE_UNKNOWN_MESSAGE,
  TIMESHEET_MONTH_HEADER_UNKNOWN_MESSAGE,
  workbookApplyBlockMessage,
} from "./ejoos/plan/planGuards";
import { opId } from "./ejoos/plan/opId";
import { planAbsentArchiveOps } from "./ejoos/plan/absentFromArchive";
import { planSzchAbsentCloseOps } from "./ejoos/plan/szchAbsentClose";
import { planTempArrivalCloseOps } from "./ejoos/plan/tempArrivalClose";
import { planRankAndContractOps } from "./ejoos/plan/rankAndContract";
import { planAbsentCloseFromShOps } from "./ejoos/plan/absentCloseFromSh";
import { planTimesheetDayFromShOps } from "./ejoos/plan/timesheetDayFromSh";
import { planTimesheetDayFromArchiveOps } from "./ejoos/plan/timesheetDayFromArchive";
import { planDispositionMovementOps } from "./ejoos/plan/movements/disposition";
import { planExcludeTransferMovementOps } from "./ejoos/plan/movements/excludeTransfer";
import { planPositionChangeMovementOps } from "./ejoos/plan/movements/positionChange";
import { planShpoOccupantOps, planShpoReconcileOps } from "./ejoos/plan/shpoOccupant";
import { planTimesheetDayCleanup } from "./ejoos/plan/timesheetDayCleanup";
import { planManualArrivalMovementOp } from "./ejoos/plan/movements/manualArrivalMovement";
import { planTransferScopeUnclearOp } from "./ejoos/plan/movements/transferScopeUnclear";
import {
  formatTransferDestinationForTimesheetMark as formatTransferDestinationForTimesheet,
  hasActualReturn,
  isPositionIndex,
  isRankAssignmentEvent,
  positionChangeDestination,
  unitCodeFromMovement,
} from "./ejoos/plan/movementContext";
import type { PbMovement } from "./ejoos/types/pb";
import type { EjoosSyncOp, EjoosSyncPlan } from "./ejoos/types/syncOp";
import type {
  EjoosAbsentRow,
  EjoosArrivalRow,
  EjoosExcludedRow,
  EjoosOosRow,
  EjoosShpoRow,
  EjoosTimesheetPersonScan,
  EjoosTimesheetRow,
} from "./ejoos/types/workbookRows";

export { extractTimesheetDestinationFromPosition };

export type {
  EjoosOpClass,
  EjoosOpKind,
  EjoosSyncOp,
  EjoosSyncPlan,
} from "./ejoos/types/syncOp";
export type {
  PbArchivePeriod,
  PbMovement,
  PbShPerson,
} from "./ejoos/types/pb";
export type {
  EjoosAbsentRow,
  EjoosArrivalRow,
  EjoosExcludedRow,
  EjoosOosRow,
  EjoosShpoRow,
  EjoosTimesheetPersonScan,
  EjoosTimesheetRow,
  JournalTimesheetDay,
} from "./ejoos/types/workbookRows";

export {
  collectProcessedMovementKeys,
  createMovementKey,
  findEjoosSheet,
  isContractMovementType,
  motivationContractOverlapsWindow,
  parseContractDatesFromChangeText,
  parsePbArchive,
  parsePbMovements,
  parsePbShPeople,
  parseRankPromotion,
  parseAsOfDateLabel,
  parseTimesheetDayFromPbName,
  planBlocksWorkbookApply,
  refreshPlanTimesheetHorizon,
  resolveJournalTimesheetDay,
  sourceTimesheetHorizonNote,
  workbookApplyBlockMessage,
  SOURCE_DATE_UNKNOWN_MESSAGE,
  TIMESHEET_MONTH_HEADER_UNKNOWN_MESSAGE,
  MONTH_ROLLOVER_BLOCK_MESSAGE,
};

const findSheet = findEjoosSheet;

export const parseEjoosAbsents = (
  sheet: ExcelSheetSnapshot | undefined,
): EjoosAbsentRow[] => {
  if (!sheet) return [];
  const rows: EjoosAbsentRow[] = [];
  for (let i = 5; i < sheet.rawRows.length; i += 1) {
    const row = sheet.rawRows[i];
    const fullName = norm(row?.[1]);
    const positionIndex = norm(row?.[3]);
    if (!fullName && !positionIndex) continue;
    rows.push({
      excelRow: i + 1,
      personId: normId(row?.[2]),
      fullName,
      rank: norm(row?.[0]),
      positionIndex,
      ground: norm(row?.[4]),
      place: norm(row?.[5]),
      departDate: norm(row?.[6]),
      actualReturn: norm(row?.[12]),
    });
  }
  return rows;
};

export const parseEjoosTimesheetDay = (
  sheet: ExcelSheetSnapshot | undefined,
  day: number,
): EjoosTimesheetRow[] => {
  if (!sheet || day < 1 || day > 31) return [];
  const dayCol = 8 + (day - 1);
  const rows: EjoosTimesheetRow[] = [];
  for (let i = 6; i < sheet.rawRows.length; i += 1) {
    const row = sheet.rawRows[i];
    const rawIndex = norm(row?.[1]);
    const positionIndex = /^\d{5,}$/.test(rawIndex) ? rawIndex : "";
    const fullName = norm(row?.[6]);
    const personId = normId(row?.[7]);
    if (!positionIndex && !fullName && !personId) continue;
    rows.push({
      excelRow: i + 1,
      personId,
      fullName,
      rank: norm(row?.[5]),
      positionIndex,
      dayValue: norm(row?.[dayCol]),
    });
  }
  return rows;
};

export const parseEjoosTimesheetPeople = (
  sheet: ExcelSheetSnapshot | undefined,
): EjoosTimesheetPersonScan[] => {
  if (!sheet) return [];
  const rows: EjoosTimesheetPersonScan[] = [];
  for (let i = 6; i < sheet.rawRows.length; i += 1) {
    const row = sheet.rawRows[i];
    const rawIndex = norm(row?.[1]);
    const positionIndex = /^\d{5,}$/.test(rawIndex) ? rawIndex : "";
    const fullName = norm(row?.[6]);
    const personId = normId(row?.[7]);
    if (!fullName && !personId) continue;
    const plusDays: number[] = [];
    const dayCodes: string[] = [];
    let hasDepartureText = false;
    let firstDepartureDay = 0;
    let departureText = "";
    for (let day = 1; day <= 31; day += 1) {
      const value = row?.[8 + (day - 1)];
      if (isTimesheetDepartureMark(value)) {
        hasDepartureText = true;
        if (!firstDepartureDay) {
          firstDepartureDay = day;
          departureText = norm(value);
        }
        dayCodes[day] = "вибув";
        continue;
      }
      const code = norm(value);
      dayCodes[day] = code;
      if (code === "+") plusDays.push(day);
    }
    rows.push({
      excelRow: i + 1,
      personId,
      fullName,
      rank: norm(row?.[5]),
      positionIndex,
      hasDepartureText,
      firstDepartureDay,
      departureText,
      plusDays,
      dayCodes,
    });
  }
  return rows;
};

export const usablePersonId = (...ids: Array<string | undefined | null>) => {
  for (const id of ids) {
    const value = String(id || "").trim();
    if (value && value !== "0") return value;
  }
  return "";
};

/** Якщо в рядку немає ID — беремо з ШПО за ПІБ або індексом посади. */
export const personIdFromShpo = (
  rows: EjoosShpoRow[],
  input: { fullName?: string; positionIndex?: string; personId?: string },
) => {
  const known = usablePersonId(input.personId);
  if (known) return known;
  const name = canonicalName(input.fullName || "");
  const index = String(input.positionIndex || "").trim();
  const byName = name
    ? rows.find(
        (row) =>
          usablePersonId(row.personId) && canonicalName(row.fullName) === name,
      )
    : undefined;
  if (byName) return byName.personId;
  const byIndex = index
    ? rows.find(
        (row) => usablePersonId(row.personId) && row.positionIndex === index,
      )
    : undefined;
  return usablePersonId(byIndex?.personId);
};

export const parseEjoosShpo = (
  sheet: ExcelSheetSnapshot | undefined,
): EjoosShpoRow[] => {
  if (!sheet) return [];
  const rows: EjoosShpoRow[] = [];
  for (let i = 6; i < sheet.rawRows.length; i += 1) {
    const row = sheet.rawRows[i];
    const positionIndex = norm(row?.[0]);
    if (!positionIndex || !/^\d/.test(positionIndex)) continue;
    rows.push({
      excelRow: i + 1,
      positionIndex,
      rank: norm(row?.[5]),
      fullName: norm(row?.[6]),
      personId: normId(row?.[7]),
    });
  }
  return rows;
};

const operationMatchesOccupant = (op: EjoosSyncOp, occupant: EjoosShpoRow) =>
  Boolean(
    (occupant.personId && op.personId && occupant.personId === op.personId) ||
    (occupant.fullName &&
      op.fullName &&
      canonicalName(occupant.fullName) === canonicalName(op.fullName)),
  );

export const findUnvacatedTargetOccupant = (
  op: EjoosSyncOp,
  shpoRows: EjoosShpoRow[],
  allOps: EjoosSyncOp[],
): EjoosShpoRow | null => {
  if (op.kind !== "position_change") return null;
  const targetIndex = op.payload.nextIndex || op.positionIndex;
  if (!targetIndex) return null;
  const occupant =
    shpoRows.find(
      (row) =>
        row.positionIndex === targetIndex &&
        Boolean(row.personId || row.fullName),
    ) ?? null;
  if (!occupant || operationMatchesOccupant(op, occupant)) return null;

  const hasReadyVacatingOperation = allOps.some((candidate) => {
    if (candidate === op || candidate.class !== "ready") return false;
    if (!operationMatchesOccupant(candidate, occupant)) return false;
    if (candidate.kind === "exclude_transfer") {
      const sourceIndex =
        candidate.payload.occupiedPositionIndex ||
        candidate.payload.fromPositionIndex ||
        candidate.payload.previousIndex ||
        candidate.positionIndex;
      return sourceIndex === targetIndex;
    }
    if (candidate.kind === "position_change") {
      const sourceIndex =
        candidate.payload.fromPositionIndex || candidate.payload.previousIndex;
      const destinationIndex =
        candidate.payload.nextIndex || candidate.positionIndex;
      return sourceIndex === targetIndex && destinationIndex !== targetIndex;
    }
    if (candidate.kind === "move_to_disposition") {
      const sourceIndex =
        candidate.payload.fromPositionIndex ||
        candidate.payload.previousIndex ||
        candidate.positionIndex;
      return sourceIndex === targetIndex;
    }
    return false;
  });

  return hasReadyVacatingOperation ? null : occupant;
};

export const parseEjoosArrivals = (
  sheet: ExcelSheetSnapshot | undefined,
): EjoosArrivalRow[] => {
  if (!sheet) return [];
  const rows: EjoosArrivalRow[] = [];
  for (let i = 5; i < sheet.rawRows.length; i += 1) {
    const row = sheet.rawRows[i];
    const fullName = norm(row?.[1]) || norm(row?.[6]);
    // Колонка C у «Тимчасово прибулі» — часто РНОКПП, не штатний ID.
    const personId = journalIdCell(row, 2);
    if (!fullName && !personId) continue;
    rows.push({
      excelRow: i + 1,
      personId,
      fullName,
      rank: norm(row?.[0]) || norm(row?.[5]),
      positionIndex: norm(row?.[4]) || norm(row?.[3]),
      fromUnit: norm(row?.[5]) || norm(row?.[8]),
      arriveDate: norm(row?.[7]) || norm(row?.[10]),
    });
  }
  return rows;
};

export const parseEjoosOos = (
  sheet: ExcelSheetSnapshot | undefined,
): EjoosOosRow[] => {
  if (!sheet) return [];
  const rows: EjoosOosRow[] = [];
  for (let i = 5; i < sheet.rawRows.length; i += 1) {
    const row = sheet.rawRows[i];
    const fullName = norm(row?.[1]);
    const personId = normId(row?.[2]);
    if (!fullName && !personId) continue;
    if (
      /вибув\s+у\s+розпорядж/i.test(fullName) ||
      /розпорядження\s+командира/i.test(fullName)
    ) {
      continue;
    }
    if (/^#/u.test(fullName) && !personId) continue;
    rows.push({
      excelRow: i + 1,
      personId,
      fullName,
      rank: norm(row?.[0]),
      positionIndex: norm(row?.[3]),
      serviceType: norm(row?.[18]),
      contractFrom: norm(row?.[19]),
      contractTo: norm(row?.[20]),
    });
  }
  return rows;
};

export const parseEjoosExcluded = (
  sheet: ExcelSheetSnapshot | undefined,
): EjoosExcludedRow[] => {
  if (!sheet) return [];
  const rows: EjoosExcludedRow[] = [];
  for (let i = 5; i < sheet.rawRows.length; i += 1) {
    const row = sheet.rawRows[i];
    const fullName = norm(row?.[1]) || norm(row?.[6]);
    const personId = normId(row?.[2]) || normId(row?.[7]);
    if (!fullName && !personId) continue;
    rows.push({
      excelRow: i + 1,
      personId,
      fullName,
      orderDate: norm(row?.[28]) || norm(row?.[10]),
      orderNumber: norm(row?.[29]) || norm(row?.[11]),
      destination: norm(row?.[30]),
      note: norm(row?.[31]),
    });
  }
  return rows;
};


export const buildEjoosSyncPlan = (
  ejoos: ExcelWorkbookSnapshot,
  pb: ExcelWorkbookSnapshot,
  options?: {
    statusRules?: import("./ejoosRules").EjoosStatusRule[];
    processedMovementKeys?: Iterable<string>;
    sourceAsOfDate?: string;
  },
): EjoosSyncPlan => {
  const statusRules =
    options?.statusRules ?? readOperatorSettings().statusRules;
  const mapStatus = (raw: string) =>
    mapPbStatusToEjoosWithRules(raw, statusRules);
  const fromName = parseTimesheetDayFromPbName(pb.fileName);
  const fromOverride = options?.sourceAsOfDate
    ? parseAsOfDateLabel(options.sourceAsOfDate)
    : null;
  const sourceDay =
    fromOverride && !fromOverride.sourceDateUnknown ? fromOverride : fromName;
  const timesheetDay = sourceDay.day;
  const timesheetDayLabel = sourceDay.label;
  const sourceDateUnknown = Boolean(sourceDay.sourceDateUnknown);
  const shPeople = parsePbShPeople(pb);
  const archiveAll = parsePbArchive(pb);
  const movementsAll = parsePbMovements(pb);
  const activeMovementsAll = movementsAll.filter(
    (event) => !isCancelledMovementRecord(event),
  );
  const processedMovementKeys = new Set(options?.processedMovementKeys ?? []);
  const wasMovementProcessed = (event: PbMovement) => {
    const key = createMovementKey(event);
    return Boolean(key && processedMovementKeys.has(key));
  };

  const absentSheet = findSheet(ejoos, /тимчасов.*відсут/i);
  const arrivalSheet = findSheet(ejoos, /тимчасов.*прибул/i);
  const excludedSheet = findSheet(ejoos, /виключен/i);
  const timesheetSheet = findSheet(ejoos, /табель/i);
  const shpoSheet = findSheet(ejoos, /шпо|штатно.?посад/i);
  const oosSheet = findSheet(ejoos, /(^|[.\s])оос($|[\s])/i);
  const ejoosAbsents = parseEjoosAbsents(absentSheet);
  const ejoosArrivals = parseEjoosArrivals(arrivalSheet);
  const ejoosExcluded = parseEjoosExcluded(excludedSheet);
  const ejoosDays = parseEjoosTimesheetDay(timesheetSheet, timesheetDay);
  const timesheetPeople = parseEjoosTimesheetPeople(timesheetSheet);
  const ejoosShpo = parseEjoosShpo(shpoSheet);
  const ejoosOos = parseEjoosOos(oosSheet);
  const shpoByIndex = new Map(ejoosShpo.map((row) => [row.positionIndex, row]));
  const positionTitleByIndex = new Map<string, string>();
  for (const person of shPeople) {
    if (person.positionIndex && person.positionTitle) {
      positionTitleByIndex.set(person.positionIndex, person.positionTitle);
    }
  }
  for (const event of activeMovementsAll) {
    if (
      event.nextIndex &&
      event.changeText &&
      (event.type === "ПОСАДА" ||
        event.type === "ПРИБУВ" ||
        event.type === "ЗМІНИШТАТ")
    ) {
      positionTitleByIndex.set(event.nextIndex, event.changeText);
    }
  }
  if (shpoSheet) {
    for (let i = 6; i < shpoSheet.rawRows.length; i += 1) {
      const row = shpoSheet.rawRows[i];
      const index = norm(row?.[0]);
      if (!index || !/^\d/.test(index)) continue;
      for (let col = 1; col <= 4; col += 1) {
        const value = norm(row?.[col]);
        if (value.length > 10 && extractUnitPhrasesFromPosition(value).length) {
          positionTitleByIndex.set(index, value);
          break;
        }
      }
    }
  }
  const staffPositionTitleForIndex = (positionIndex: string) =>
    positionTitleByIndex.get(positionIndex) || "";

  const shIds = new Set(
    shPeople.map((person) => person.personId).filter(Boolean),
  );
  const shNames = new Set(
    shPeople.map((person) => canonicalName(person.fullName)).filter(Boolean),
  );
  const shPersonById = new Map(
    shPeople
      .filter((person) => person.personId)
      .map((person) => [person.personId, person]),
  );
  const shPersonByName = new Map(
    shPeople
      .filter((person) => person.fullName)
      .map((person) => [canonicalName(person.fullName), person]),
  );
  const shpoPersonById = new Map(
    ejoosShpo.filter((row) => row.personId).map((row) => [row.personId, row]),
  );
  const shpoPersonByName = new Map(
    ejoosShpo
      .filter((row) => row.fullName)
      .map((row) => [canonicalName(row.fullName), row]),
  );
  const oosPersonById = new Map(
    ejoosOos.filter((row) => row.personId).map((row) => [row.personId, row]),
  );
  const oosPersonByName = new Map(
    ejoosOos
      .filter((row) => row.fullName)
      .map((row) => [canonicalName(row.fullName), row]),
  );
  // Поточний облік — лише ШПО + ООС. «Виключені» і старі рядки Табеля —
  // історія завершених вибуттів, а не доказ, що людина зараз вибула.
  const ejoosOccupied = [
    ...ejoosShpo.filter((row) => row.fullName || row.personId),
    ...ejoosOos.filter((row) => row.fullName || row.personId),
  ];
  const ejoosIds = new Set(
    ejoosOccupied.map((row) => row.personId).filter(Boolean),
  );
  const ejoosNames = new Set(
    ejoosOccupied.map((row) => canonicalName(row.fullName)).filter(Boolean),
  );
  const activeArrivalIds = new Set(
    ejoosArrivals.map((row) => row.personId).filter(Boolean),
  );
  const activeArrivalNames = new Set(
    ejoosArrivals.map((row) => canonicalName(row.fullName)).filter(Boolean),
  );
  const openAbsentRows = ejoosAbsents.filter((row) => !row.actualReturn);
  const openAbsentIds = new Set(
    openAbsentRows.map((row) => row.personId).filter(Boolean),
  );
  const openAbsentNames = new Set(
    openAbsentRows.map((row) => canonicalName(row.fullName)).filter(Boolean),
  );

  // Один ID — одна людина. Написання ПІБ у джерелах різне (Миколайович /
  // Михайлович), а частина рядків ЕЖООС узагалі без ID, тому для звірки
  // тримаємо всі варіанти ПІБ, зв'язані спільним ID.
  const nameVariantsById = new Map<
    string,
    Map<string, { display: string; sources: Set<string> }>
  >();
  const addNameVariant = (
    personId: string,
    fullName: string,
    source: string,
  ) => {
    const id = normId(personId);
    const name = norm(fullName);
    const key = canonicalName(name);
    if (!id || !key) return;
    const variants =
      nameVariantsById.get(id) ??
      new Map<string, { display: string; sources: Set<string> }>();
    const variant = variants.get(key) ?? { display: name, sources: new Set() };
    variant.sources.add(source);
    variants.set(key, variant);
    nameVariantsById.set(id, variants);
  };
  for (const event of activeMovementsAll) {
    addNameVariant(event.personId, event.fullName, "Рух");
  }
  for (const period of archiveAll) {
    addNameVariant(period.personId, period.fullName, "archive");
  }
  for (const person of shPeople) {
    addNameVariant(person.personId, person.fullName, "sh");
  }
  for (const row of ejoosShpo)
    addNameVariant(row.personId, row.fullName, "ШПО");
  for (const row of ejoosOos) addNameVariant(row.personId, row.fullName, "ООС");
  for (const row of ejoosAbsents) {
    addNameVariant(row.personId, row.fullName, "Тимч. відсутні");
  }
  for (const row of ejoosArrivals) {
    addNameVariant(row.personId, row.fullName, "Тимч. прибулі");
  }

  const idsByName = new Map<string, Set<string>>();
  for (const [personId, variants] of nameVariantsById) {
    for (const key of variants.keys()) {
      const ids = idsByName.get(key) ?? new Set<string>();
      ids.add(personId);
      idsByName.set(key, ids);
    }
  }
  /** Усі написання ПІБ цієї особи: пряме, і всі зв'язані спільним ID. */
  const personNameKeys = (personId: string, fullName: string) => {
    const keys = new Set<string>();
    const name = canonicalName(norm(fullName));
    if (name) keys.add(name);
    const ids = new Set<string>();
    const id = normId(personId);
    if (id) ids.add(id);
    if (!id && name) {
      const linked = idsByName.get(name);
      // Однакове ПІБ у різних ID — не вгадуємо, працюємо лише з прямим ПІБ.
      if (linked?.size === 1) ids.add([...linked][0]);
    }
    for (const candidate of ids) {
      for (const key of nameVariantsById.get(candidate)?.keys() ?? []) {
        keys.add(key);
      }
    }
    return keys;
  };
  /** Пошук рядка за будь-яким написанням ПІБ цієї особи. */
  const byPersonName = <T>(
    map: Map<string, T>,
    personId: string,
    fullName: string,
  ) => {
    for (const key of personNameKeys(personId, fullName)) {
      const found = map.get(key);
      if (found) return found;
    }
    return null;
  };
  /** Рядки ЕЖООС часто без ID, тому порівнюємо ще й за псевдонімами ПІБ. */
  const isSamePerson = (
    left: { personId: string; fullName: string },
    right: { personId: string; fullName: string },
  ) => {
    const leftId = normId(left.personId);
    const rightId = normId(right.personId);
    if (leftId && rightId) return leftId === rightId;
    const rightName = canonicalName(norm(right.fullName));
    if (!rightName) return false;
    return personNameKeys(left.personId, left.fullName).has(rightName);
  };

  const personStillInEjoos = (personId: string, fullName: string) =>
    Boolean(
      (personId && ejoosIds.has(personId)) ||
      [...personNameKeys(personId, fullName)].some((key) =>
        ejoosNames.has(key),
      ),
    );
  const personStillInSh = (personId: string, fullName: string) =>
    Boolean(
      (personId && shIds.has(personId)) ||
      [...personNameKeys(personId, fullName)].some((key) => shNames.has(key)),
    );
  const onStaffShpo = (personId: string, fullName: string) =>
    Boolean(
      (personId && shpoPersonById.has(personId)) ||
      byPersonName(shpoPersonByName, personId, fullName),
    );
  const onStaffOos = (personId: string, fullName: string) =>
    Boolean(
      (personId && oosPersonById.has(personId)) ||
      byPersonName(oosPersonByName, personId, fullName),
    );
  const timesheetClosedFor = (
    personId: string,
    fullName: string,
    sourcePositionTitle = "",
  ) => {
    const rows = timesheetPeople.filter(
      (row) =>
        (personId && row.personId === personId) ||
        [...personNameKeys(personId, fullName)].some(
          (key) => canonicalName(row.fullName) === key,
        ),
    );
    if (!rows.length) return false;
    const title = String(sourcePositionTitle || "").trim();
    if (title && timesheetSheet) {
      const misplaced = rows.some(
        (row) =>
          !timesheetRowInExpectedUnitSection(
            timesheetSheet,
            row.excelRow,
            title,
          ),
      );
      if (misplaced) return false;
    }
    const openNamed = rows.filter(
      (row) =>
        !isClosedTimesheetHistoryRow(row) &&
        Boolean(row.personId || row.fullName),
    );
    if (openNamed.some((row) => row.hasDepartureText)) return false;
    return rows.every((row) => isClosedTimesheetHistoryRow(row));
  };
  const sourcePositionTitleForOutbound = (event: PbMovement) => {
    const fromIndex = staffPositionTitleForIndex(event.previousIndex);
    if (fromIndex) return fromIndex;
    for (const period of archiveAll) {
      if (isSamePerson(event, period) && period.positionTitle) {
        return period.positionTitle;
      }
    }
    return "";
  };
  const hasOpenAbsence = (personId: string, fullName: string) =>
    Boolean(
      (personId && openAbsentIds.has(personId)) ||
      [...personNameKeys(personId, fullName)].some((key) =>
        openAbsentNames.has(key),
      ),
    );
  /**
   * Особу вже знято з штатної посади (розпорядження / безвісти), а відкрита
   * відсутність у ЕЖООС є. Повторно ставити на вакантний індекс або писати
   * новий рядок БЕЗВІСТИ не треба — це вже проведений стан.
   */
  const alreadyVacatedForAbsence = (personId: string, fullName: string) =>
    isStaleVacatedAbsence({
      onStaffShpo: onStaffShpo(personId, fullName),
      hasOpenAbsence: hasOpenAbsence(personId, fullName),
      stillInSh: personStillInSh(personId, fullName),
    });
  const findMovementExcludedRow = (event: PbMovement) =>
    ejoosExcluded.find((row) => {
      if (!isSamePerson(event, row)) return false;
      const sameOrder = Boolean(
        event.orderNumber &&
        row.orderNumber &&
        normKey(row.orderNumber) === normKey(event.orderNumber),
      );
      const sameDate = Boolean(
        event.orderDate &&
        row.orderDate &&
        normKey(row.orderDate) === normKey(event.orderDate),
      );
      return event.orderNumber ? sameOrder : sameDate;
    }) ?? null;
  const findLatestExcludedRow = (personId: string, fullName: string) =>
    [...ejoosExcluded]
      .filter((row) => isSamePerson({ personId, fullName }, row))
      .sort((left, right) => right.excelRow - left.excelRow)[0] ?? null;
  const excludedRowsOfPerson = (personId: string, fullName: string) =>
    ejoosExcluded.filter((row) => isSamePerson({ personId, fullName }, row));
  const staleExcludedForMovement = (event: {
    personId: string;
    fullName: string;
    orderNumber: string;
    orderDate: string;
  }) =>
    staleExcludedRowsToClear({
      rows: excludedRowsOfPerson(event.personId, event.fullName),
      currentOrderNumber: event.orderNumber,
      currentOrderDate: event.orderDate,
    });
  const staleExcludedClearPayload = (
    stale: Array<{ excelRow: number }>,
  ): Pick<
    EjoosSyncOp["payload"],
    "clearExcludedExcelRow" | "clearExcludedExcelRows"
  > => ({
    clearExcludedExcelRow: stale[0] ? String(stale[0].excelRow) : "",
    clearExcludedExcelRows: stale.map((row) => row.excelRow).join(","),
  });
  const findFalseHopExcludedRow = (personId: string, fullName: string) => {
    if (!personStillInSh(personId, fullName)) return null;
    if (!personStillInEjoos(personId, fullName)) return null;
    const row = findLatestExcludedRow(personId, fullName);
    if (!row || !isFalseInternalHopExclusion(row)) return null;
    return row;
  };
  const movementPersonKey = (event: { personId: string; fullName: string }) =>
    event.personId
      ? `id:${event.personId}`
      : event.fullName
        ? `name:${normKey(event.fullName)}`
        : "";
  const movementEventTime = (event: PbMovement) =>
    dateMs(event.orderDate || event.basisDate) || event.excelRow;

  /**
   * ПЕРЕВ + СКАСУВАННЯ в тому ж місяці: переведення не чинне.
   * Один рядок Табеля; запис у «Виключені» прибираємо.
   */
  const latestTransferCancelByPerson = new Map<string, PbMovement>();
  const cancelledExternalTransferRows = new Set<number>();
  for (const event of movementsAll) {
    if (!isTransferCancellation(event)) continue;
    const key = movementPersonKey(event);
    if (!key) continue;
    const previous = latestTransferCancelByPerson.get(key);
    if (
      !previous ||
      movementEventTime(event) > movementEventTime(previous) ||
      (movementEventTime(event) === movementEventTime(previous) &&
        event.excelRow > previous.excelRow)
    ) {
      latestTransferCancelByPerson.set(key, event);
    }
  }
  for (const event of movementsAll) {
    if (event.type !== "ПЕРЕВ" || isOwnUnitStaffMove(event)) continue;
    const key = movementPersonKey(event);
    if (!key) continue;
    if (isCancelledMovementRecord(event) || isTransferCancellation(event)) {
      cancelledExternalTransferRows.add(event.excelRow);
      continue;
    }
    const cancel = latestTransferCancelByPerson.get(key);
    if (!cancel) continue;
    if (
      movementEventTime(cancel) > movementEventTime(event) ||
      (movementEventTime(cancel) === movementEventTime(event) &&
        cancel.excelRow > event.excelRow)
    ) {
      cancelledExternalTransferRows.add(event.excelRow);
    }
  }
  const isLaterMovement = (later: PbMovement, earlier: PbMovement) =>
    movementEventTime(later) > movementEventTime(earlier) ||
    (movementEventTime(later) === movementEventTime(earlier) &&
      later.excelRow > earlier.excelRow);
  const laterOutboundStaffDeparture = (
    event: PbMovement,
  ): PbMovement | null => {
    if (personStillInSh(event.personId, event.fullName)) return null;
    let found: PbMovement | null = null;
    for (const other of activeMovementsAll) {
      if (other.excelRow === event.excelRow) continue;
      if (!isSamePerson(event, other)) continue;
      if (cancelledExternalTransferRows.has(other.excelRow)) continue;
      if (!isOutboundStaffMove(other)) continue;
      if (!isLaterMovement(other, event)) continue;
      if (!found || isLaterMovement(other, found)) found = other;
    }
    return found;
  };
  const ownUnitMoveSuperseded = (event: PbMovement) =>
    ownUnitMoveSupersededByOutbound({
      stillInSh: personStillInSh(event.personId, event.fullName),
      hasLaterOutbound: Boolean(laterOutboundStaffDeparture(event)),
    });
  const transferCancelForPerson = (personId: string, fullName: string) => {
    if (personId) {
      const found = latestTransferCancelByPerson.get(`id:${personId}`);
      if (found) return found;
    }
    for (const event of latestTransferCancelByPerson.values()) {
      if (isSamePerson({ personId, fullName }, event)) return event;
    }
    if (fullName) {
      return (
        latestTransferCancelByPerson.get(`name:${normKey(fullName)}`) ?? null
      );
    }
    return null;
  };
  const transferCancelOf = (event: PbMovement) =>
    transferCancelForPerson(event.personId, event.fullName);
  const cancelledTransferOf = (event: {
    personId: string;
    fullName: string;
  }) => {
    const cancel = transferCancelForPerson(event.personId, event.fullName);
    if (!cancel) return null;
    let found: PbMovement | null = null;
    for (const movement of movementsAll) {
      if (movement.type !== "ПЕРЕВ" || isOwnUnitStaffMove(movement)) {
        continue;
      }
      if (!cancelledExternalTransferRows.has(movement.excelRow)) continue;
      if (!isSamePerson(event, movement)) continue;
      if (movementEventTime(movement) > movementEventTime(cancel)) continue;
      if (
        !found ||
        movementEventTime(movement) > movementEventTime(found) ||
        (movementEventTime(movement) === movementEventTime(found) &&
          movement.excelRow > found.excelRow)
      ) {
        found = movement;
      }
    }
    return found;
  };

  const dayByIndex = new Map<string, EjoosTimesheetRow>();
  const dayRowScore = (row: EjoosTimesheetRow) => {
    const shpo = shpoByIndex.get(row.positionIndex);
    if (!shpo) return !row.personId && !row.fullName ? 2 : 1;
    if (shpo.personId && row.personId === shpo.personId) return 4;
    if (shpo.fullName && normKey(row.fullName) === normKey(shpo.fullName)) {
      return 4;
    }
    if (!shpo.personId && !shpo.fullName && !row.personId && !row.fullName) {
      return 3;
    }
    return 1;
  };
  for (const row of ejoosDays) {
    if (!row.positionIndex) continue;
    const current = dayByIndex.get(row.positionIndex);
    if (!current || dayRowScore(row) > dayRowScore(current)) {
      dayByIndex.set(row.positionIndex, row);
    }
  }
  const timesheetScanByRow = new Map(
    timesheetPeople.map((row) => [row.excelRow, row]),
  );
  const timesheetActivityScore = (row: EjoosTimesheetRow) => {
    const scan = timesheetScanByRow.get(row.excelRow);
    if (scan?.hasDepartureText && !scan.plusDays.length) return 0;
    if (scan?.hasDepartureText) return 1 + dayRowScore(row);
    return 10 + dayRowScore(row);
  };
  const dayById = new Map<string, EjoosTimesheetRow>();
  for (const row of ejoosDays) {
    if (!row.personId) continue;
    const current = dayById.get(row.personId);
    if (
      !current ||
      timesheetActivityScore(row) > timesheetActivityScore(current)
    ) {
      dayById.set(row.personId, row);
    }
  }
  const activeTimesheetRowOf = (
    personId: string,
    fullName: string,
    staffIndex = "",
  ) => {
    const byIndex = staffIndex ? dayByIndex.get(staffIndex) : undefined;
    if (
      byIndex &&
      isSamePerson({ personId, fullName }, byIndex) &&
      timesheetActivityScore(byIndex) > 0
    ) {
      return byIndex;
    }
    if (personId) {
      const byId = dayById.get(personId);
      if (byId && timesheetActivityScore(byId) > 0) return byId;
    }
    return (
      timesheetRowsOf(personId, fullName)
        .map((scan) => ejoosDays.find((row) => row.excelRow === scan.excelRow))
        .filter((row): row is EjoosTimesheetRow => Boolean(row))
        .sort(
          (left, right) =>
            timesheetActivityScore(right) - timesheetActivityScore(left),
        )[0] ?? null
    );
  };
  const timesheetRowsOf = (personId: string, fullName: string) =>
    timesheetPeople.filter((row) => isSamePerson({ personId, fullName }, row));
  /** Табель інколи має чужий Excel-серійний ID — для РОЗПОРЯДЖ шукаємо ще за ПІБ. */
  const timesheetRowByCanonicalName = (fullName: string) => {
    const want = canonicalName(fullName);
    if (!want) return null;
    const scan = timesheetPeople
      .filter((row) => canonicalName(row.fullName) === want)
      .sort((left, right) => {
        const leftOpen = isClosedTimesheetHistoryRow(left) ? 0 : 1;
        const rightOpen = isClosedTimesheetHistoryRow(right) ? 0 : 1;
        return rightOpen - leftOpen || left.excelRow - right.excelRow;
      })[0];
    if (!scan) return null;
    return (
      ejoosDays.find((row) => row.excelRow === scan.excelRow) ?? {
        excelRow: scan.excelRow,
        personId: scan.personId,
        fullName: scan.fullName,
        rank: "",
        positionIndex: scan.positionIndex,
        dayValue: "",
      }
    );
  };
  const priorEpisodeTimesheetOf = (
    personId: string,
    fullName: string,
    activeExcelRow: number,
    staffIndex: string,
  ) =>
    timesheetRowsOf(personId, fullName).find(
      (row) =>
        row.excelRow !== activeExcelRow &&
        (row.hasDepartureText ||
          Boolean(
            staffIndex && row.positionIndex && row.positionIndex !== staffIndex,
          )),
    ) ?? null;
  const isVacantStaffRow = (
    row: { personId: string; fullName: string } | null | undefined,
  ) => Boolean(row) && !row!.personId && !row!.fullName;
  /** Рядок Табеля особи на штатному індексі — не чужий рядок dayByIndex. */
  const staffIndexTimesheetForPerson = (
    personId: string,
    fullName: string,
    positionIndex: string,
  ) => {
    const personRows = timesheetRowsOf(personId, fullName).filter(
      (row) => row.positionIndex === positionIndex,
    );
    if (personRows.length) {
      const ranked = [...personRows].sort(
        (left, right) => right.plusDays.length - left.plusDays.length,
      );
      return ranked.find((row) => !row.hasDepartureText) ?? ranked[0];
    }
    const indexRow = positionIndex ? dayByIndex.get(positionIndex) : undefined;
    if (indexRow && isSamePerson({ personId, fullName }, indexRow))
      return indexRow;
    if (isVacantStaffRow(indexRow)) return indexRow ?? null;
    return null;
  };
  const shOccupantByIndex = new Map(
    shPeople
      .filter(
        (person) =>
          person.positionIndex && isPositionIndex(person.positionIndex),
      )
      .map((person) => [person.positionIndex, person]),
  );
  /**
   * Скасоване переведення — не розбиваємо Табель на історію + новий рядок.
   * Фактичного вибуття не було; лишається один активний рядок з дати постановки.
   */
  const timesheetNeedsTransferCancelSplit = (
    _personId: string,
    _fullName: string,
    _staffIndex: string,
    _cancelDate: string,
  ) => false;

  /** Дата постановки на поточний штатний індекс (ПОСАДА / внутрішній ПЕРЕВ). */
  const staffAppointmentDateFor = (
    personId: string,
    fullName: string,
    positionIndex: string,
  ) => {
    let found: PbMovement | null = null;
    for (const event of activeMovementsAll) {
      if (!isSamePerson({ personId, fullName }, event)) continue;
      if (!eventInLeadWindow(event)) continue;
      const isInbound = isOwnUnitStaffMove(event);
      if (!isInbound) continue;
      const targetIndex = event.nextIndex || "";
      if (
        positionIndex &&
        targetIndex &&
        isPositionIndex(targetIndex) &&
        targetIndex !== positionIndex
      ) {
        continue;
      }
      if (
        !found ||
        movementEventTime(event) > movementEventTime(found) ||
        (movementEventTime(event) === movementEventTime(found) &&
          event.excelRow > found.excelRow)
      ) {
        found = event;
      }
    }
    return found?.orderDate || found?.basisDate || "";
  };

  /**
   * ПОСАДА також може бути далеко за межами хвоста РУХ. Підтягуємо останню
   * подію, яка підтверджує поточну особу та індекс у `sh`; інакше різниця
   * sh ↔ старий ШПО помилково виглядає як звичайний конфлікт зайнятості.
   */
  const currentPositionMovements = new Map<
    string,
    (typeof movementsAll)[number]
  >();
  for (const event of activeMovementsAll) {
    if (!isOwnUnitStaffMove(event) || !event.fullName) {
      continue;
    }
    const person =
      (event.personId && shPersonById.get(event.personId)) ||
      byPersonName(shPersonByName, event.personId, event.fullName) ||
      shPeople.find((candidate) => isSamePerson(candidate, event)) ||
      null;
    const currentEjoosPerson =
      ((person?.personId || event.personId) &&
        shpoPersonById.get(person?.personId || event.personId)) ||
      byPersonName(shpoPersonByName, event.personId, event.fullName) ||
      ((person?.personId || event.personId) &&
        oosPersonById.get(person?.personId || event.personId)) ||
      byPersonName(oosPersonByName, event.personId, event.fullName) ||
      null;
    if (!person?.positionIndex || !isPositionIndex(person.positionIndex)) {
      // Особу вже прибрали з поточного sh, а її стару позицію зайняв хтось
      // інший. Не губимо підтверджений внутрішній ПЕРЕВ — показуємо його
      // окремою операцією для ручної перевірки/застосування.
      // Якщо далі є зовнішнє вибуття, внутрішню постановку не проводимо:
      // кінцевий стан визначає ПЕРЕВ / sh, а не проміжна ПОСАДА.
      if (
        currentEjoosPerson &&
        isOwnUnitStaffMove(event) &&
        !ownUnitMoveSuperseded(event)
      ) {
        currentPositionMovements.set(canonicalName(event.fullName), event);
      }
      continue;
    }
    const confirmsCurrentIndex =
      event.nextIndex === person.positionIndex ||
      String(event.changeText || "").includes(person.positionIndex);
    const currentOosPerson =
      ((person.personId || event.personId) &&
        oosPersonById.get(person.personId || event.personId)) ||
      byPersonName(oosPersonByName, event.personId, event.fullName) ||
      null;
    const oosPositionIndexes = (currentOosPerson?.positionIndex || "")
      .split(/[^0-9]+/)
      .filter(Boolean);
    const oosNeedsPositionUpdate =
      !currentOosPerson || !oosPositionIndexes.includes(person.positionIndex);
    const indexActuallyChanged = Boolean(
      currentEjoosPerson?.positionIndex &&
      currentEjoosPerson.positionIndex !== person.positionIndex,
    );
    const currentShpoPerson =
      ((person.personId || event.personId) &&
        shpoPersonById.get(person.personId || event.personId)) ||
      byPersonName(shpoPersonByName, event.personId, event.fullName) ||
      null;
    // ООС уже може бути на індексі, а ШПО ще вакантний — це все одно постановка.
    const isNewPlacement = confirmsCurrentIndex && !currentShpoPerson;
    const cancel = transferCancelForPerson(
      person.personId || event.personId,
      person.fullName || event.fullName,
    );
    const needsCancelSplit = Boolean(
      cancel?.orderDate &&
      timesheetNeedsTransferCancelSplit(
        person.personId || event.personId,
        person.fullName || event.fullName,
        person.positionIndex,
        cancel.orderDate,
      ),
    );
    // Якщо ШПО вже містить цю людину на поточному індексі sh, рух проведений
    // раніше й повторно показувати/застосовувати його не можна — окрім
    // скасованого переведення, де Табель ще не розкладений на історію + новий рядок.
    if (
      !indexActuallyChanged &&
      !isNewPlacement &&
      !oosNeedsPositionUpdate &&
      !needsCancelSplit
    ) {
      continue;
    }
    currentPositionMovements.set(canonicalName(person.fullName), event);
  }

  const currentPositionMovementRows = new Set(
    [...currentPositionMovements.values()].map((event) => event.excelRow),
  );
  const internalPositionMovementRows = new Set(
    [...currentPositionMovements.values()]
      .filter((event) => event.type === "ПЕРЕВ")
      .map((event) => event.excelRow),
  );

  /**
   * По одній особі може бути кілька непроведених подій. Історію зміни посади
   * втрачати не можна, тому ланцюг ПОСАДА / внутрішній ПЕРЕВ будуємо від
   * індексу, який ще стоїть у ЕЖООС, і проводимо кроки послідовно за датою.
   */
  const positionChainByPerson = new Map<string, PbMovement[]>();
  const chainedPositionRows = new Set<number>();
  const chainStepByRow = new Map<number, { step: number; total: number }>();
  {
    const positionEventsByPerson = new Map<string, PbMovement[]>();
    for (const event of activeMovementsAll) {
      // Ланцюг ведемо лише по переходах усередині 1ПБ: вибуття в іншу частину
      // проводиться через виключення, а не через зміну штатної посади.
      const isPositionEvent = isOwnUnitStaffMove(event);
      if (!isPositionEvent) continue;
      if (!event.previousIndex || !event.nextIndex) continue;
      if (event.previousIndex === event.nextIndex) continue;
      const key = movementPersonKey(event);
      if (!key) continue;
      const events = positionEventsByPerson.get(key) ?? [];
      events.push(event);
      positionEventsByPerson.set(key, events);
    }
    const eventTime = (event: PbMovement) =>
      dateMs(event.orderDate || event.basisDate);
    const isDispositionToken = (value: string) => /розпорядж/iu.test(value);
    const indexesConnect = (from: string, to: string) =>
      Boolean(from) &&
      Boolean(to) &&
      (from === to || (isDispositionToken(from) && isDispositionToken(to)));
    // Індекс у ШПО буває порожній (рядок лише з ПІБ), тому додатково
    // орієнтуємось на штатний рядок Табеля.
    const timesheetIndexById = new Map<string, string>();
    const timesheetIndexByName = new Map<string, string>();
    for (const row of ejoosDays) {
      if (!row.positionIndex) continue;
      if (/РОЗПОРЯДЖ/iu.test(row.dayValue)) continue;
      if (row.personId && !timesheetIndexById.has(row.personId)) {
        timesheetIndexById.set(row.personId, row.positionIndex);
      }
      const key = canonicalName(row.fullName);
      if (key && !timesheetIndexByName.has(key)) {
        timesheetIndexByName.set(key, row.positionIndex);
      }
    }
    for (const [key, events] of positionEventsByPerson) {
      events.sort(
        (left, right) =>
          eventTime(left) - eventTime(right) || left.excelRow - right.excelRow,
      );
      const personId = key.startsWith("id:") ? key.slice(3) : "";
      const fullName = events[0].fullName;
      const ejoosRow =
        (personId && shpoPersonById.get(personId)) ||
        byPersonName(shpoPersonByName, personId, fullName) ||
        null;
      // Точка відліку — те, що реально стоїть у ЕЖООС. Якщо індекс уже новий,
      // подія проведена раніше й у ланцюг не потрапляє.
      let currentIndex = ejoosRow?.positionIndex || "";
      // Якщо штатного рядка в ШПО вже немає, індекс у Табелі часто лишається
      // від історії вибуття в розпорядження. З нього ланцюг посад не будуємо —
      // інакше вже проведений РОЗПОРЯДЖ знову виглядає як зміна посади.
      if (!currentIndex) {
        const leftStaff = Boolean(
          (personId && oosPersonById.has(personId)) ||
          byPersonName(oosPersonByName, personId, fullName) ||
          (personId && openAbsentIds.has(personId)) ||
          [...personNameKeys(personId, fullName)].some((key) =>
            openAbsentNames.has(key),
          ),
        );
        if (!leftStaff) {
          currentIndex =
            (personId && timesheetIndexById.get(personId)) ||
            byPersonName(timesheetIndexByName, personId, fullName) ||
            "";
        }
      }
      // Повернення з розпорядження / СЗЧ: у ШПО штатного індексу ще немає,
      // ланцюг починаємо з ПОСАДИ «розпорядження → 2103…».
      if (!currentIndex) {
        const fromDisposition =
          events.find((event) => isDispositionToStaffPlacement(event)) ||
          events.find((event) => isDispositionToken(event.previousIndex));
        if (fromDisposition && personStillInSh(personId, fullName)) {
          currentIndex = fromDisposition.previousIndex;
        }
      }
      if (!currentIndex) continue;
      const pending: PbMovement[] = [];
      const visited = [currentIndex];
      for (const event of events) {
        if (!indexesConnect(event.previousIndex, currentIndex)) continue;
        currentIndex = event.nextIndex;
        // Повернення на індекс, який уже був у маршруті (А → Б → А):
        // проміжні кроки взаємно погашені, історію переносити нікуди.
        const loopAt = visited.indexOf(event.nextIndex);
        if (loopAt >= 0) {
          pending.length = loopAt;
          visited.length = loopAt + 1;
          continue;
        }
        pending.push(event);
        visited.push(event.nextIndex);
      }
      if (!pending.length) continue;
      // `sh` — джерело істини про поточну посаду. Якщо ланцюг веде не туди,
      // це стара або суперечлива подія: обробляємо звичайним шляхом.
      const shIndex =
        (personId && shPersonById.get(personId)?.positionIndex) ||
        byPersonName(shPersonByName, personId, fullName)?.positionIndex ||
        "";
      if (isPositionIndex(shIndex) && shIndex !== currentIndex) continue;
      const lastPending = pending[pending.length - 1];
      if (ownUnitMoveSuperseded(lastPending)) continue;
      const lastRow = lastPending.excelRow;
      const hasFollowUpEvent = activeMovementsAll.some(
        (event) =>
          movementPersonKey(event) === key &&
          event.excelRow > lastRow &&
          !cancelledExternalTransferRows.has(event.excelRow) &&
          (event.type === "РОЗПОРЯДЖ" ||
            event.type === "ПЕРЕВ" ||
            event.type === "ЗВІЛЬН"),
      );
      const total = pending.length + (hasFollowUpEvent ? 1 : 0);
      pending.forEach((event, index) => {
        chainedPositionRows.add(event.excelRow);
        chainStepByRow.set(event.excelRow, { step: index + 1, total });
      });
      positionChainByPerson.set(key, pending);
    }
  }
  const isPositionChangeRow = (excelRow: number) =>
    currentPositionMovementRows.has(excelRow) ||
    chainedPositionRows.has(excelRow);

  const operationalTypes = new Set([
    "ПОСАДА",
    "ПЕРЕВ",
    "РОЗПОРЯДЖ",
    "ПРИБУВ",
    "ЗВІЛЬН",
    "СЗЧ",
    "БЕЗВІСТИ",
  ]);
  const latestOperationalRowByPerson = new Map<string, number>();
  const latestRozporadjRowByPerson = new Map<string, number>();
  for (const event of activeMovementsAll) {
    if (!operationalTypes.has(event.type)) continue;
    if (cancelledExternalTransferRows.has(event.excelRow)) continue;
    if (event.type === "СКАСУВАННЯ") continue;
    const key = movementPersonKey(event);
    if (!key) continue;
    latestOperationalRowByPerson.set(
      key,
      Math.max(latestOperationalRowByPerson.get(key) ?? 0, event.excelRow),
    );
    if (event.type === "РОЗПОРЯДЖ") {
      latestRozporadjRowByPerson.set(
        key,
        Math.max(latestRozporadjRowByPerson.get(key) ?? 0, event.excelRow),
      );
    }
  }
  const effectiveMovements = activeMovementsAll.filter((event) => {
    if (event.type === "СКАСУВАННЯ") return false;
    if (cancelledExternalTransferRows.has(event.excelRow)) return false;
    if (!operationalTypes.has(event.type)) return false;
    // Крім останньої події беремо непроведені кроки зміни посади і ту
    // ПОСАДУ, яка підтверджує поточний індекс у sh. Інакше ланцюг
    // «РОЗПОРЯДЖ → ПОСАДА → ПЕРЕВ → СКАСУВАННЯ → СЗЧ» губить постановку.
    if (chainedPositionRows.has(event.excelRow)) return true;
    if (currentPositionMovementRows.has(event.excelRow)) return true;
    const key = movementPersonKey(event);
    if (
      event.type === "РОЗПОРЯДЖ" &&
      key &&
      latestRozporadjRowByPerson.get(key) === event.excelRow
    ) {
      return true;
    }
    return Boolean(
      key && latestOperationalRowByPerson.get(key) === event.excelRow,
    );
  });

  const openAbsents = ejoosAbsents.filter((row) => !row.actualReturn);
  const openById = new Map(
    openAbsents.filter((row) => row.personId).map((row) => [row.personId, row]),
  );
  const openByName = new Map(
    openAbsents.map((row) => [normKey(row.fullName), row]),
  );
  const arrivalById = new Map(
    ejoosArrivals
      .filter((row) => row.personId)
      .map((row) => [row.personId, row]),
  );
  const arrivalByName = new Map(
    ejoosArrivals.map((row) => [normKey(row.fullName), row]),
  );
  const oosById = new Map(
    ejoosOos.filter((row) => row.personId).map((row) => [row.personId, row]),
  );
  const oosByName = new Map(
    ejoosOos
      .filter((row) => row.fullName)
      .map((row) => [canonicalName(row.fullName), row]),
  );
  const latestPositionByName = new Map<string, PbMovement>();
  for (const event of effectiveMovements) {
    if (!isPositionChangeRow(event.excelRow) || !event.fullName) {
      continue;
    }
    latestPositionByName.set(normKey(event.fullName), event);
  }
  const positionEventForShPerson = (person: PbShPerson) => {
    const event = latestPositionByName.get(normKey(person.fullName));
    if (!event) return null;
    if (wasMovementProcessed(event)) return null;
    if (
      event.nextIndex &&
      person.positionIndex &&
      event.nextIndex !== person.positionIndex
    ) {
      return null;
    }
    return event;
  };

  // Ведуться лише зміни місяця 1ПБ (для 25.08.2026 — серпень).
  // Давніші періоди не підтягуємо і не дописуємо як нові статуси.
  const reportDateMs = dateMs(timesheetDayLabel);
  const monthStartMs = (() => {
    const match = String(timesheetDayLabel || "").match(
      /(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})/,
    );
    if (!match) return 0;
    const year =
      Number(match[3]) < 100 ? 2000 + Number(match[3]) : Number(match[3]);
    return Date.UTC(year, Number(match[2]) - 1, 1);
  })();
  const journalMonthStartLabel = (() => {
    const match = String(timesheetDayLabel || "").match(
      /(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})/,
    );
    if (!match) return "";
    const year =
      Number(match[3]) < 100 ? 2000 + Number(match[3]) : Number(match[3]);
    return `01.${String(match[2]).padStart(2, "0")}.${year}`;
  })();
  const timesheetMonthHeader = formatTimesheetMonthHeader(timesheetDayLabel);
  const timesheetHeaderCell = findTimesheetMonthHeaderCell(
    timesheetSheet?.rawRows ?? [],
  );
  const journalMonth = monthStartMs
    ? new Date(monthStartMs).getUTCMonth() + 1
    : 0;
  const journalYear = monthStartMs
    ? new Date(monthStartMs).getUTCFullYear()
    : 0;
  const timesheetMonthHeaderUnknown = !timesheetHeaderCell;
  const timesheetMonthHeaderMismatch = Boolean(
    timesheetMonthHeader &&
    timesheetHeaderCell &&
    (timesheetHeaderCell.month !== journalMonth ||
      timesheetHeaderCell.year !== journalYear),
  );
  const leadWindowStart = monthStartMs;
  const leadWindowEnd = reportDateMs || monthStartMs;
  const dateInLeadWindow = (value: string) => {
    const ms = dateMs(value);
    return Boolean(
      ms && leadWindowStart && ms >= leadWindowStart && ms <= leadWindowEnd,
    );
  };
  const eventInLeadWindow = (event: {
    orderDate: string;
    basisDate: string;
  }) => {
    const ms = dateMs(event.orderDate || event.basisDate);
    if (!ms || !leadWindowStart) return true;
    return ms >= leadWindowStart && ms <= leadWindowEnd;
  };
  const inTempArrivals = (personId: string, fullName: string) =>
    Boolean(
      (personId && arrivalById.get(personId)) ||
      byPersonName(arrivalByName, personId, fullName),
    );
  // Постановка цього місяця: відкритий СЗЧ з травня треба внести разом із
  // ШПО/Табелем, а не наступним прогоном після застосування ПОСАДИ.
  const inboundStaffPlacementThisMonth = (
    personId: string,
    fullName: string,
  ) => {
    const event =
      latestPositionByName.get(normKey(fullName)) ||
      [...latestPositionByName.values()].find((item) =>
        isSamePerson({ personId, fullName }, item),
      ) ||
      null;
    if (event && eventInLeadWindow(event)) return true;
    return (
      inTempArrivals(personId, fullName) &&
      Boolean(staffAppointmentDateFor(personId, fullName, ""))
    );
  };
  const shConfirmsOpenArchive = (period: {
    personId: string;
    fullName: string;
    absenceType: string;
  }) => {
    const sh =
      (period.personId && shPersonById.get(period.personId)) ||
      byPersonName(shPersonByName, period.personId, period.fullName) ||
      null;
    if (!sh) return false;
    return currentStatusConfirmsOpenAbsence(
      mapStatus(sh.status).timesheetCode,
      mapStatus(period.absenceType).timesheetCode,
    );
  };
  const laterArchivePeriodOf = (period: {
    personId: string;
    fullName: string;
    departDate: string;
  }) => {
    const departMs = dateMs(period.departDate);
    if (!departMs) return null;
    return (
      archiveAll
        .filter(
          (other) =>
            isSamePerson(period, other) && dateMs(other.departDate) > departMs,
        )
        .sort(
          (left, right) => dateMs(left.departDate) - dateMs(right.departDate),
        )[0] ?? null
    );
  };
  const laterArchivePeriodSupersedesOpen = (period: {
    personId: string;
    fullName: string;
    departDate: string;
    returnDate: string;
  }) =>
    !hasActualReturn(period.returnDate) &&
    Boolean(laterArchivePeriodOf(period));
  const archive = archiveAll.filter((period) => {
    const stillInSh = personStillInSh(period.personId, period.fullName);
    const stillTracked =
      stillInSh ||
      onStaffOos(period.personId, period.fullName) ||
      hasOpenAbsence(period.personId, period.fullName);
    if (!stillTracked) return false;
    const departMs = dateMs(period.departDate);
    const returnMs = hasActualReturn(period.returnDate)
      ? dateMs(period.returnDate)
      : null;
    const carryOpen = stillInSh
      ? shConfirmsOpenArchive(period)
      : hasOpenAbsence(period.personId, period.fullName);
    if (
      archivePeriodTouchesJournalMonth(
        departMs,
        returnMs,
        leadWindowStart,
        leadWindowEnd,
        { carryOpen },
      )
    ) {
      return true;
    }
    // Відкритий період до місяця: особа саме зараз заходить на штат
    // (sh уже «В СТРОЮ», тож confirmOpenCarry не спрацює). Внутрішній
    // стрибок посади і пізніша закрита відпустка старий МЕДРОТА не оживають.
    if (
      !returnMs &&
      departMs &&
      departMs < leadWindowStart &&
      inboundStaffPlacementThisMonth(period.personId, period.fullName) &&
      !laterArchivePeriodSupersedesOpen(period)
    ) {
      return true;
    }
    return false;
  });
  const augustAbsenceSpansFor = (personId: string, fullName: string) => {
    const person = { personId, fullName };
    for (const event of activeMovementsAll) {
      if (!isSamePerson(person, event)) continue;
      if (event.type !== "РОЗПОРЯДЖ") continue;
      const orderMs = dateMs(event.orderDate || event.basisDate);
      if (orderMs && leadWindowStart && orderMs < leadWindowStart) {
        return [];
      }
    }
    const spans = buildTimesheetAbsenceSpans(
      archiveAll
        .filter((period) => isSamePerson({ personId, fullName }, period))
        .map((period) => ({
          departDate: period.departDate,
          returnDate: period.returnDate,
          absenceType: period.absenceType,
          excelRow: period.excelRow,
          personId: period.personId,
          fullName: period.fullName,
          departMs: dateMs(period.departDate),
          returnMs: hasActualReturn(period.returnDate)
            ? dateMs(period.returnDate)
            : null,
        })),
      {
        timesheetDay,
        monthStartMs: leadWindowStart,
        monthEndMs: leadWindowEnd,
        reportDayMs: reportDateMs,
        mapCode: (absenceType) => mapStatus(absenceType).timesheetCode || "",
        hasReturn: hasActualReturn,
        confirmOpenCarry: (period) =>
          shConfirmsOpenArchive({
            personId: period.personId || personId,
            fullName: period.fullName || fullName,
            absenceType: period.absenceType,
          }),
      },
    );
    const shPerson =
      (personId && shPersonById.get(personId)) ||
      byPersonName(shPersonByName, personId, fullName) ||
      null;
    const shCode = shPerson ? mapStatus(shPerson.status).timesheetCode : null;
    const latest = [...spans].reverse().find((span) => span.code)?.code ?? null;
    return extendDispositionSpanToReportDay(spans, {
      timesheetDay,
      shTimesheetCode: shCode,
      latestArchiveCode: latest,
    });
  };

  const inboundStaffPlacementBefore = (event: PbMovement) =>
    findLatestPriorOwnUnitStaffMove(activeMovementsAll, event, {
      samePerson: isSamePerson,
      inWindow: eventInLeadWindow,
      eventTime: movementEventTime,
    });

  const inboundStaffDateFor = (personId: string, fullName: string) => {
    let inbound: PbMovement | null = null;
    let outboundBefore: PbMovement | null = null;
    for (const event of activeMovementsAll) {
      if (!isSamePerson({ personId, fullName }, event)) continue;
      if (!eventInLeadWindow(event)) continue;
      if (isOwnUnitStaffMove(event)) {
        inbound = event;
      } else if (isOutboundStaffMove(event)) {
        outboundBefore = event;
      }
    }
    if (
      inbound &&
      outboundBefore &&
      movementEventTime(inbound) > movementEventTime(outboundBefore)
    ) {
      return inbound.orderDate || inbound.basisDate;
    }
    return "";
  };

  const timesheetEpisodeStartFor = (
    personId: string,
    fullName: string,
    staffIndex: string,
  ) => {
    const ownUnitMoves: PbMovement[] = [];
    // ПРИБУВ може бути відсіяний з effective/active рухів пізнішою ПОСАДА,
    // але його дата все одно є початком фактичного перебування в підрозділі.
    const fullMovementChain = movementsAll
      .filter(
        (event) =>
          isSamePerson({ personId, fullName }, event) &&
          eventInLeadWindow(event),
      )
      .sort(
        (left, right) => movementEventTime(left) - movementEventTime(right),
      );
    const latestOutboundAt = fullMovementChain
      .filter(isOutboundStaffMove)
      .reduce((latest, event) => Math.max(latest, movementEventTime(event)), 0);
    const explicitArrival = [...fullMovementChain]
      .reverse()
      .find(
        (event) =>
          event.type === "ПРИБУВ" &&
          movementEventTime(event) >= latestOutboundAt,
      );
    const explicitArrivalDate =
      explicitArrival?.orderDate || explicitArrival?.basisDate || "";
    let leftUnitThisMonth = false;
    for (const event of activeMovementsAll) {
      if (!isSamePerson({ personId, fullName }, event)) continue;
      if (!eventInLeadWindow(event)) continue;
      if (isOutboundStaffMove(event)) leftUnitThisMonth = true;
      if (isOwnUnitStaffMove(event)) ownUnitMoves.push(event);
    }
    const temporaryArrival =
      (personId && arrivalById.get(personId)) ||
      byPersonName(arrivalByName, personId, fullName) ||
      null;
    const temporaryArrivalDate = dateInLeadWindow(
      temporaryArrival?.arriveDate || "",
    )
      ? temporaryArrival?.arriveDate || ""
      : "";
    const externalStaffArrival = [...ownUnitMoves]
      .filter((event) => !isInternalStaffIndexHop(event))
      .sort(
        (left, right) => movementEventTime(right) - movementEventTime(left),
      )[0];
    const externalStaffArrivalDate =
      externalStaffArrival?.orderDate || externalStaffArrival?.basisDate || "";
    const inbound =
      inboundStaffDateFor(personId, fullName) || temporaryArrivalDate || "";
    if (inbound) leftUnitThisMonth = true;
    const departed = timesheetRowsOf(personId, fullName).some(
      (row) => row.hasDepartureText,
    );
    return resolveTimesheetEpisodeStart({
      monthStartLabel: journalMonthStartLabel,
      // Фактичне ПРИБУВ у РУХ починає перебування раніше за наступний наказ
      // ПОСАДА. ШЕВЧУК: прибув 02.08, посада 06.08, archive з 03.08.
      appointmentDate: resolveTimesheetArrivalDate({
        explicitArrivalDate,
        externalStaffArrivalDate,
        temporaryArrivalDate,
        staffAppointmentDate: staffAppointmentDateFor(
          personId,
          fullName,
          staffIndex,
        ),
      }),
      inboundDate: inbound,
      hasMonthStartAbsence: augustAbsenceSpansFor(personId, fullName).some(
        (span) => span.fromDay === 1,
      ),
      hasDepartureEvidence: departed,
      leftUnitThisMonth,
      wasTemporaryArrival: Boolean(temporaryArrivalDate),
      ownUnitMoves,
    });
  };

  const staffEpisodePaintPayload = (
    personId: string,
    fullName: string,
    staffIndex: string,
    activeExcelRow: number,
  ) => {
    let appointment = timesheetEpisodeStartFor(personId, fullName, staffIndex);
    const allSpans = augustAbsenceSpansFor(personId, fullName);
    const carryFromMonthStart = allSpans.some((span) => span.fromDay === 1);
    let activeFromDay = carryFromMonthStart
      ? 1
      : journalDayFromDateMs(dateMs(appointment), leadWindowStart);
    const existingScan = timesheetScanByRow.get(activeExcelRow);
    const firstPlusDay = existingScan?.plusDays.length
      ? Math.min(...existingScan.plusDays)
      : 0;
    const hasInactivePrefix =
      firstPlusDay > 1 &&
      Array.from({ length: firstPlusDay - 1 }, (_, index) => index + 1).every(
        (day) => sameTimesheetDayMark(existingScan?.dayCodes[day] || "", "-"),
      );
    const resolvedActiveFromDay = resolveExistingTimesheetStartDay({
      calculatedDay: activeFromDay,
      firstPlusDay,
      hasInactivePrefix,
    });
    if (resolvedActiveFromDay !== activeFromDay) {
      activeFromDay = resolvedActiveFromDay;
      const monthMatch = journalMonthStartLabel.match(
        /^\d{2}\.(\d{2})\.(\d{4})$/,
      );
      if (monthMatch) {
        appointment = `${String(activeFromDay).padStart(2, "0")}.${monthMatch[1]}.${monthMatch[2]}`;
      }
    }
    const activeSpans =
      activeFromDay > 1
        ? clipAbsenceSpansToActiveEpisode(allSpans, activeFromDay)
        : allSpans;
    const historySpans =
      !carryFromMonthStart && activeFromDay > 1
        ? absenceSpansBeforeEpisode(allSpans, activeFromDay)
        : [];
    const history = priorEpisodeTimesheetOf(
      personId,
      fullName,
      activeExcelRow,
      staffIndex,
    );
    return {
      timesheetActiveFrom: appointment,
      timesheetPreserveHistory:
        !carryFromMonthStart && activeFromDay > 1 ? "1" : "",
      timesheetAbsenceSpans: encodeTimesheetAbsenceSpans(activeSpans),
      historyTimesheetExcelRow:
        history && historySpans.length ? String(history.excelRow) : "",
      historyTimesheetAbsenceSpans: encodeTimesheetAbsenceSpans(historySpans),
    };
  };

  const ops: EjoosSyncOp[] = [];
  const szchAbsent = planSzchAbsentCloseOps({
    activeMovementsAll,
    shPeople,
    ejoosDays,
    timesheetDayLabel,
    openById,
    openByName,
    dayByIndex,
    eventInLeadWindow,
    activeTimesheetRowOf,
  });
  ops.push(...szchAbsent.ops);
  const absenceRowsClosedByMovement = szchAbsent.absenceRowsClosedByMovement;
  // SHPO / Табель occupant identity from sh (by position index)
  ops.push(
    ...planShpoOccupantOps({
      shPeople,
      existingOps: ops,
      shpoByIndex,
      dayByIndex,
      arrivalById,
      arrivalByName,
      oosPersonById,
      oosPersonByName,
      positionEventForShPerson,
      alreadyVacatedForAbsence,
      byPersonName,
      staffIndexTimesheetForPerson,
      findLatestExcludedRow,
      personStillInEjoos,
      transferCancelForPerson,
      staffEpisodePaintPayload,
      isSamePerson,
    }),
  );

  ops.push(
    ...planAbsentCloseFromShOps({
      shPeople,
      archive,
      openById,
      openByName,
      arrivalById,
      arrivalByName,
      absenceRowsClosedByMovement,
      timesheetDayLabel,
      mapStatus,
      dateInLeadWindow,
      alreadyVacatedForAbsence,
      isSamePerson,
      byPersonName,
      activeTimesheetRowOf,
      augustAbsenceSpansFor,
    }),
  );
  ops.push(
    ...planTimesheetDayFromShOps({
      shPeople,
      ejoosDays,
      archiveAll,
      timesheetDay,
      timesheetDayLabel,
      dayByIndex,
      shpoByIndex,
      timesheetScanByRow,
      arrivalById,
      arrivalByName,
      mapStatus,
      alreadyVacatedForAbsence,
      positionEventForShPerson,
      isSamePerson,
      byPersonName,
      timesheetRowsOf,
      activeTimesheetRowOf,
      augustAbsenceSpansFor,
      staffEpisodePaintPayload,
      findFalseHopExcludedRow,
    }),
  );


  ops.push(
    ...planAbsentArchiveOps({
      archive,
      ejoosAbsents,
      openById,
      openByName,
      shPersonById,
      shPersonByName,
      absenceRowsClosedByMovement,
      existingOps: ops,
      mapStatus,
      isSamePerson,
      byPersonName,
      laterArchivePeriodOf,
      alreadyVacatedForAbsence,
      inboundStaffPlacementThisMonth,
      activeTimesheetRowOf,
      staffEpisodePaintPayload,
    }),
  );

  ops.push(
    ...planTimesheetDayFromArchiveOps({
      shPeople,
      existingOps: ops,
      timesheetDay,
      leadWindowStart,
      timesheetScanByRow,
      personStillInSh,
      positionEventForShPerson,
      isSamePerson,
      augustAbsenceSpansFor,
      staffAppointmentDateFor,
      inboundStaffDateFor,
      activeTimesheetRowOf,
      staffEpisodePaintPayload,
    }),
  );

  const seenTransferCancelReview = new Set<string>();
  for (const event of movementsAll) {
    if (!isTransferCancellation(event) || !eventInLeadWindow(event)) continue;
    const key = movementPersonKey(event);
    if (!key || seenTransferCancelReview.has(key)) continue;
    seenTransferCancelReview.add(key);
    const cancelled = cancelledTransferOf(event);
    const excludedRow = cancelled
      ? findMovementExcludedRow(cancelled)
      : findMovementExcludedRow(event);
    const tsRows = timesheetRowsOf(event.personId, event.fullName);
    const timesheetHasDepart = tsRows.some((row) => row.hasDepartureText);
    const evidence = Boolean(excludedRow || timesheetHasDepart);
    const shPerson =
      (event.personId && shPersonById.get(event.personId)) ||
      byPersonName(shPersonByName, event.personId, event.fullName) ||
      null;
    const inCurrentSh = personStillInSh(event.personId, event.fullName);
    const missingFromCurrentSh = !inCurrentSh;
    const dest = cancelled
      ? formatTransferDestinationForTimesheet(
          [cancelled.destination, cancelled.changeText]
            .filter(Boolean)
            .join(" "),
        ) || cancelled.destination
      : "";
    const staffIndex =
      shPerson?.positionIndex ||
      cancelled?.previousIndex ||
      event.nextIndex ||
      event.previousIndex ||
      "";
    const staffTs =
      staffIndexTimesheetForPerson(
        event.personId,
        event.fullName,
        staffIndex,
      ) ||
      [...tsRows]
        .filter((row) => !row.hasDepartureText)
        .sort(
          (left, right) => right.plusDays.length - left.plusDays.length,
        )[0] ||
      tsRows.find((row) => row.hasDepartureText) ||
      null;
    const historyTs =
      tsRows.find(
        (row) => row.hasDepartureText && row.excelRow !== staffTs?.excelRow,
      ) ?? null;
    const staffTsScan = staffTs
      ? tsRows.find((row) => row.excelRow === staffTs.excelRow)
      : null;
    const staffTsHasDepart = Boolean(staffTsScan?.hasDepartureText);
    const staffRowVacant = isVacantStaffRow(staffTs);
    const hasActiveStaffTs = tsRows.some((row) => !row.hasDepartureText);
    const restoreTargetTs =
      staffTsHasDepart || staffRowVacant
        ? staffTs
        : tsRows.find((row) => row.hasDepartureText) || staffTs;
    const restoreTimesheet =
      inCurrentSh &&
      (staffTsHasDepart ||
        staffRowVacant ||
        (timesheetHasDepart && !hasActiveStaffTs) ||
        (!tsRows.length && Boolean(staffTs)));
    const shpoAtIndex = staffIndex
      ? (shpoByIndex.get(staffIndex) ?? null)
      : null;
    const shpoPerson =
      shpoAtIndex ||
      (event.personId && shpoPersonById.get(event.personId)) ||
      byPersonName(shpoPersonByName, event.personId, event.fullName) ||
      null;
    const shpoVacant = Boolean(
      shpoAtIndex && !shpoAtIndex.fullName && !shpoAtIndex.personId,
    );
    const restoreShpo = inCurrentSh && shpoVacant;
    const oosPerson =
      (event.personId && oosPersonById.get(event.personId)) ||
      byPersonName(oosPersonByName, event.personId, event.fullName) ||
      null;
    const restoreOos = inCurrentSh && !oosPerson && Boolean(excludedRow);
    const timesheetActiveFrom =
      staffAppointmentDateFor(event.personId, event.fullName, staffIndex) ||
      event.orderDate ||
      "";
    const hasWorkbookWrite =
      Boolean(excludedRow) ||
      restoreTimesheet ||
      restoreShpo ||
      restoreOos ||
      Boolean(historyTs);
    // Після успішного rollback скасування лишається в РУХ, але писати вже нічого.
    // Інакше особа вічно крутиться в «До застосування» з порожнім «Табель».
    if (!missingFromCurrentSh && !hasWorkbookWrite) continue;
    ops.push({
      id: opId([
        "transfer-cancel",
        event.personId || event.fullName,
        event.orderNumber || String(event.excelRow),
      ]),
      kind: "other_manual",
      class: missingFromCurrentSh ? "needs_input" : "ready",
      sheet: missingFromCurrentSh
        ? "Дані джерел / 3. Виключені"
        : restoreTimesheet && !excludedRow
          ? "6. Табель"
          : "3. Виключені",
      personId: shPerson?.personId || event.personId,
      fullName: shPerson?.fullName || event.fullName,
      positionIndex:
        shPerson?.positionIndex || event.nextIndex || event.previousIndex,
      rank: shPerson?.rank || event.rank,
      before: missingFromCurrentSh
        ? `РУХ: ПЕРЕВ №${cancelled?.orderNumber || "?"} → ${dest || "?"} скасовано №${event.orderNumber || "?"}; sh: немає; ЕЖООС: Виключені${excludedRow ? ` R${excludedRow.excelRow}` : ""} / Табель закритий`
        : excludedRow
          ? `Виключені R${excludedRow.excelRow}: №${cancelled?.orderNumber || "?"} від ${cancelled?.orderDate || "?"}`
          : staffTsHasDepart
            ? `Табель R${staffTs?.excelRow}: ще стоїть «вибув» після скасованого ПЕРЕВ`
            : "рядка виключення за скасованим ПЕРЕВ немає",
      after: missingFromCurrentSh
        ? "NEEDS_REVIEW — скасування є, але в актуальній sh людини немає; ШПО / ООС / Табель автоматично не відновлюємо"
        : evidence
          ? `REMOVE_CANCELLED_EXCLUSION — прибрати запис №${cancelled?.orderNumber || "?"} від ${cancelled?.orderDate || "?"} (скасовано №${event.orderNumber || "?"})`
          : "REMOVE_CANCELLED_EXCLUSION — нового рядка у Виключених не створюємо",
      sourceRef: `Рух!R${event.excelRow} №${event.orderNumber || "?"} · скасування`,
      why: missingFromCurrentSh
        ? `Скасування переведення №${event.orderNumber || "?"} від ${event.orderDate || "?"} (ПЕРЕВ №${cancelled?.orderNumber || "?"} від ${cancelled?.orderDate || "?"} → ${dest || "?"}). Рух очікує залишення в 1 ПБ, але в актуальній sh особи немає${/нема в sh/i.test(event.status) ? " (у рядку скасування теж «Нема в sh»)" : ""}. Це не AUTO_RESTORE і не NO_ACTION: поки sh не підтвердить перебування, ШПО/ООС/новий Табель не відновлюємо.`
        : evidence
          ? `Переведення №${cancelled?.orderNumber || "?"} від ${cancelled?.orderDate || "?"} скасовано №${event.orderNumber || "?"} від ${event.orderDate || "?"}. Фактичного виключення не було — запис у «Виключені» прибираємо, Табель лишається одним рядком.`
          : `Скасування №${event.orderNumber || "?"} від ${event.orderDate || "?"}: рядка у Виключених і «вибув» у Табелі немає. Новий рядок не створюємо, особу виключеною не вважаємо.`,
      confidence: missingFromCurrentSh ? "manual" : "high",
      payload: {
        type: "TRANSFER_CANCELLED",
        reviewReason: missingFromCurrentSh
          ? "CANCEL_TRANSFER_BUT_NOT_IN_CURRENT_SH"
          : "",
        excludedExcelRow: excludedRow ? String(excludedRow.excelRow) : "",
        cancelledTransferOrder: cancelled?.orderNumber || "",
        cancelledTransferDate: cancelled?.orderDate || "",
        cancelledDestination: dest,
        transferCancelOrder: event.orderNumber,
        transferCancelDate: event.orderDate,
        actualExclusionEvidence: evidence ? "1" : "",
        inCurrentSh: inCurrentSh ? "1" : "",
        previousIndex: staffIndex,
        restoreTimesheet: restoreTimesheet ? "1" : "",
        timesheetExcelRow:
          restoreTimesheet && restoreTargetTs
            ? String(restoreTargetTs.excelRow)
            : "",
        timesheetActiveFrom: restoreTimesheet ? timesheetActiveFrom : "",
        timesheetAbsenceSpans: restoreTimesheet
          ? encodeTimesheetAbsenceSpans(
              augustAbsenceSpansFor(event.personId, event.fullName),
            )
          : "",
        restoreShpo: restoreShpo ? "1" : "",
        shpoExcelRow:
          (restoreShpo || Boolean(excludedRow)) && shpoPerson
            ? String(shpoPerson.excelRow)
            : "",
        restoreOos: restoreOos ? "1" : "",
        historyTimesheetExcelRow: historyTs ? String(historyTs.excelRow) : "",
      },
      movementKey: createMovementKey(event),
      checkedDefault: !missingFromCurrentSh,
    });
  }

  /**
   * `norm` віддає числа з діапазону Excel-дат як дату, тому ID (напр. 22814)
   * у тексті рядка губиться. ID збираємо окремо через `normId`.
   */
  type SheetScanRow = { excelRow: number; text: string; ids: Set<string> };
  const scanRows = (
    sheet: ExcelSheetSnapshot | undefined,
    pattern: RegExp,
  ): SheetScanRow[] =>
    (sheet?.rawRows ?? [])
      .map((row, index) => ({
        excelRow: index + 1,
        text: canonicalName(row.map(norm).join(" ")),
        ids: new Set(row.map((value) => normId(value)).filter(Boolean)),
      }))
      .filter((row) => pattern.test(row.text));
  const DISPOSITION_RE = /РОЗПОРЯДЖ/iu;
  const ABSENCE_STATUS_RE =
    /СЗЧ|САМОВІЛ|БЕЗВІСТ|(?:^|[^А-ЯІЇЄҐа-яіїєґ])ЗБ(?:$|[^А-ЯІЇЄҐа-яіїєґ])/iu;
  const shpoDispositionRows = scanRows(shpoSheet, DISPOSITION_RE);
  const timesheetDispositionRows = scanRows(timesheetSheet, DISPOSITION_RE);
  const shpoSzchRows = scanRows(shpoSheet, ABSENCE_STATUS_RE);
  const timesheetSzchRows = scanRows(timesheetSheet, ABSENCE_STATUS_RE);
  /** Штатний рядок Табеля не є блоком розпорядження, навіть якщо в клітинці дня є «вибув у розпорядження». */
  const isTimesheetStaffPositionRow = (excelRow: number) => {
    const scan = timesheetScanByRow.get(excelRow);
    return Boolean(scan?.positionIndex && isPositionIndex(scan.positionIndex));
  };
  const timesheetDispositionStaffRows = timesheetDispositionRows.filter(
    (row) => !isTimesheetStaffPositionRow(row.excelRow),
  );
  const personInTextRows = (event: PbMovement, rows: SheetScanRow[]) => {
    const id = normId(event.personId);
    if (id && rows.some((row) => row.ids.has(id))) return true;
    // Рядки ЕЖООС часто без ID, а написання ПІБ у джерелах різне.
    const names = [...personNameKeys(event.personId, event.fullName)];
    return names.some((name) => rows.some((row) => row.text.includes(name)));
  };
  const findNonStaffOccupantExcelRow = (
    sheet: ExcelSheetSnapshot | undefined,
    personId: string,
    fullName: string,
    cols: { index: number; name: number; id: number },
  ) => {
    if (!sheet) return 0;
    const names = [...personNameKeys(personId, fullName)];
    const id = normId(personId);
    for (let i = 6; i < sheet.rawRows.length; i += 1) {
      const row = sheet.rawRows[i];
      const index = norm(row?.[cols.index]);
      if (index && /^\d/.test(index)) continue;
      const rowId = normId(row?.[cols.id]);
      const rowName = canonicalName(norm(row?.[cols.name]));
      if (id && rowId && id === rowId) return i + 1;
      if (
        rowName &&
        names.some((name) => name === rowName || rowName.includes(name))
      ) {
        return i + 1;
      }
    }
    return 0;
  };
  const openAbsenceOf = (personId: string, fullName: string) =>
    (personId && openById.get(personId)) ||
    byPersonName(openByName, personId, fullName) ||
    openByName.get(normKey(fullName)) ||
    null;
  const arrivalOf = (personId: string, fullName: string) =>
    ejoosArrivals.find((row) => isSamePerson({ personId, fullName }, row)) ||
    (personId && arrivalById.get(personId)) ||
    byPersonName(arrivalByName, personId, fullName) ||
    null;
  const ownUnitIndexHistoryOf = (personId: string, fullName: string) => {
    const seen = new Set<string>();
    const entries: Array<{ index: string; date: string; order: string }> = [];
    const events = activeMovementsAll
      .filter(
        (item) =>
          isSamePerson({ personId, fullName }, item) &&
          eventInLeadWindow(item) &&
          isOwnUnitStaffMove(item) &&
          isPositionIndex(item.nextIndex),
      )
      .sort(
        (left, right) =>
          movementEventTime(left) - movementEventTime(right) ||
          left.excelRow - right.excelRow,
      );
    for (const item of events) {
      if (seen.has(item.nextIndex)) continue;
      seen.add(item.nextIndex);
      entries.push({
        index: item.nextIndex,
        date: item.orderDate || item.basisDate || "",
        order: item.orderNumber || "",
      });
    }
    return entries;
  };
  const shpoByPositionIndex = new Map(
    ejoosShpo.map((row) => [row.positionIndex, row]),
  );
  /**
   * У ШПО трапляються рядки без індексу — лише з ПІБ (залишок після
   * попередніх правок). Вони не є ні штатною посадою, ні блоком
   * розпорядження, але їх треба чистити разом із виведенням особи.
   */
  const shpoStrayRowByName = new Map<string, EjoosShpoRow>();
  {
    const dispositionExcelRows = new Set(
      shpoDispositionRows.map((row) => row.excelRow),
    );
    (shpoSheet?.rawRows ?? []).forEach((row, index) => {
      const excelRow = index + 1;
      if (excelRow < 7 || dispositionExcelRows.has(excelRow)) return;
      if (norm(row?.[0])) return;
      const fullName = norm(row?.[6]);
      if (!fullName) return;
      const key = canonicalName(fullName);
      if (!key || shpoStrayRowByName.has(key)) return;
      shpoStrayRowByName.set(key, {
        excelRow,
        positionIndex: "",
        rank: norm(row?.[5]),
        fullName,
        personId: normId(row?.[7]),
      });
    });
  }
  const activeSzchPersonIds = new Set<string>();
  const activeSzchPersonNames = new Set<string>();
  for (const period of archiveAll) {
    if (
      !isDispositionAbsenceStatus(period.absenceType) ||
      hasActualReturn(period.returnDate)
    ) {
      continue;
    }
    if (period.personId) activeSzchPersonIds.add(period.personId);
    if (period.fullName) {
      activeSzchPersonNames.add(canonicalName(period.fullName));
    }
  }
  const reportDate = dateMs(timesheetDayLabel);
  const staleMovementCutoff = reportDate ? reportDate - 90 * 86400000 : 0;

  const pendingRankByPerson = new Map<string, EjoosSyncOp>();
  let latestRankEventByPerson = new Map<string, PbMovement>();
  const personTraceableForContract = (personId: string, fullName: string) =>
    personStillInEjoos(personId, fullName) ||
    personStillInSh(personId, fullName) ||
    timesheetPeople.some((row) => isSamePerson({ personId, fullName }, row));
  const contractEventInLeadWindow = (event: PbMovement) =>
    eventInLeadWindow(event) ||
    (isContractMovementType(event.type) &&
      motivationContractOverlapsWindow(
        event,
        leadWindowStart || 0,
        leadWindowEnd || 0,
      ));

  const rankAndContract = planRankAndContractOps({
    activeMovementsAll,
    ejoosDays,
    eventInLeadWindow,
    contractEventInLeadWindow,
    movementPersonKey,
    movementEventTime,
    personStillInEjoos,
    personTraceableForContract,
    isSamePerson,
    byPersonName,
    shpoPersonById,
    shpoPersonByName,
    oosPersonById,
    oosPersonByName,
    dayById,
    staffIndexTimesheetForPerson,
    isContractMovementType,
  });
  ops.push(...rankAndContract.ops);
  for (const [key, value] of rankAndContract.pendingRankByPerson) {
    pendingRankByPerson.set(key, value);
  }
  latestRankEventByPerson = rankAndContract.latestRankEventByPerson;

  const latestAbsenceOnlyOf = (personId: string, fullName: string) => {
    let found: PbMovement | null = null;
    for (const event of activeMovementsAll) {
      if (!isSamePerson({ personId, fullName }, event)) continue;
      if (!eventInLeadWindow(event)) continue;
      if (!isAbsenceOnlyMovement(event)) continue;
      if (
        !found ||
        movementEventTime(event) > movementEventTime(found) ||
        (movementEventTime(event) === movementEventTime(found) &&
          event.excelRow > found.excelRow)
      ) {
        found = event;
      }
    }
    return found;
  };
  const openDispositionAbsenceMs = (personId: string, fullName: string) => {
    let latest = 0;
    for (const period of archiveAll) {
      if (!isSamePerson({ personId, fullName }, period)) continue;
      if (!isDispositionAbsenceStatus(period.absenceType)) continue;
      if (hasActualReturn(period.returnDate)) continue;
      latest = Math.max(latest, dateMs(period.departDate) || 0);
    }
    for (const row of ejoosAbsents) {
      if (!isSamePerson({ personId, fullName }, row)) continue;
      if (!isDispositionAbsenceStatus(row.ground)) continue;
      if (row.actualReturn) continue;
      latest = Math.max(latest, dateMs(row.departDate) || 0);
    }
    const movement = latestAbsenceOnlyOf(personId, fullName);
    if (movement) latest = Math.max(latest, movementEventTime(movement));
    return latest;
  };

  const priorMonthDispositionMonthLabel = (orderDate: string) => {
    const ms = dateMs(orderDate);
    if (!ms) return orderDate || "минулого місяця";
    const monthNames = [
      "січень",
      "лютий",
      "березень",
      "квітень",
      "травень",
      "червень",
      "липень",
      "серпень",
      "вересень",
      "жовтень",
      "листопад",
      "грудень",
    ];
    const date = new Date(ms);
    return `${monthNames[date.getUTCMonth()]} ${date.getUTCFullYear()}`;
  };
  const isPriorMonthPendingDisposition = (event: PbMovement) => {
    if (event.type !== "РОЗПОРЯДЖ" || wasMovementProcessed(event)) {
      return false;
    }
    const movementDate = dateMs(event.orderDate || event.basisDate);
    if (!leadWindowStart || !movementDate || movementDate >= leadWindowStart) {
      return false;
    }
    return (
      personStillInSh(event.personId, event.fullName) ||
      onStaffShpo(event.personId, event.fullName)
    );
  };

  const considerMovement = (event: PbMovement) => {
    const movementDate = dateMs(event.orderDate || event.basisDate);
    if (
      leadWindowStart &&
      movementDate &&
      movementDate < leadWindowStart &&
      !isPriorMonthPendingDisposition(event)
    ) {
      return;
    }
    if (isAmbiguousStaffTransfer(event)) {
      const scopeOp = planTransferScopeUnclearOp(event);
      if (scopeOp) ops.push(scopeOp);
      return;
    }
    const hasCurrentTrace = Boolean(
      personStillInSh(event.personId, event.fullName) ||
      personStillInEjoos(event.personId, event.fullName) ||
      (event.personId &&
        (activeArrivalIds.has(event.personId) ||
          openAbsentIds.has(event.personId))) ||
      (event.fullName &&
        (activeArrivalNames.has(normKey(event.fullName)) ||
          openAbsentNames.has(normKey(event.fullName)))),
    );
    if (
      staleMovementCutoff &&
      movementDate &&
      movementDate < staleMovementCutoff &&
      !hasCurrentTrace
    ) {
      return;
    }

    if (isAbsenceOnlyMovement(event)) return;

    if (event.type === "РОЗПОРЯДЖ") {
      const disposition = planDispositionMovementOps({
        event,
        activeMovementsAll,
        activeSzchPersonIds,
        activeSzchPersonNames,
        activeTimesheetRowOf,
        alreadyVacatedForAbsence,
        archiveAll,
        augustAbsenceSpansFor,
        absenceSpansBeforeEpisode,
        byPersonName,
        canonicalName,
        createMovementKey,
        dateMs,
        dayByIndex,
        ejoosAbsents,
        ejoosDays,
        encodeTimesheetAbsenceSpans,
        eventInLeadWindow,
        formatDispositionTimesheetDeparture,
        hasActualReturn,
        isDispositionAbsenceStatus,
        isOwnUnitStaffMove,
        isSamePerson,
        isTimesheetStaffPositionRow,
        isVacantStaffRow,
        journalDayFromDateMs,
        leadWindowStart,
        mapStatus,
        movementEventTime,
        movementPersonKey,
        onStaffShpo,
        oosPersonById,
        oosPersonByName,
        opId,
        personInTextRows,
        personNameKeys,
        personStillInSh,
        positionChainByPerson,
        priorMonthDispositionMonthLabel,
        shpoByPositionIndex,
        shpoDispositionRows,
        shpoPersonById,
        shpoPersonByName,
        shpoStrayRowByName,
        shpoSzchRows,
        staffIndexTimesheetForPerson,
        staffPositionTitleForIndex,
        timesheetDay,
        timesheetDayLabel,
        timesheetDispositionStaffRows,
        timesheetRowByCanonicalName,
        timesheetSzchRows,
        norm,
      });
      ops.push(...disposition.ops);
      if (disposition.handled) return;
    }

    const excludeTransfer = planExcludeTransferMovementOps({
      event,
      existingOps: ops,
      absenceOnlyBlocksExclusion,
      archiveAll,
      arrivalOf,
      byPersonName,
      cancelledExternalTransferRows,
      createMovementKey,
      dayByIndex,
      ejoosShpo,
      excludedTimesheetWrite,
      externalTransferProcessState,
      extractTimesheetDestinationFromPosition,
      findMovementExcludedRow,
      formatTransferDestinationForTimesheet,
      inboundStaffDateFor,
      inboundStaffPlacementBefore,
      internalPositionMovementRows,
      isOutboundStaffMove,
      isOwnFirstPbDestination,
      isPositionChangeRow,
      isPositionIndex,
      isSamePerson,
      isUnrecordedSameMonthTransit,
      isVacantStaffRow,
      latestRankEventByPerson,
      movementEventTime,
      movementPersonKey,
      onStaffOos,
      onStaffShpo,
      oosPersonById,
      oosPersonByName,
      openDispositionAbsenceMs,
      opId,
      parseRankPromotion,
      pendingRankByPerson,
      personIdFromShpo,
      personStillInEjoos,
      personStillInSh,
      positionChangeDestination,
      shpoByIndex,
      skipExternalIfAlreadyProcessed,
      sourcePositionTitleForOutbound,
      staffEpisodePaintPayload,
      staleExcludedForMovement,
      staleExcludedClearPayload,
      timesheetClosedFor,
      timesheetRowInExpectedUnitSection,
      timesheetRowsOf,
            timesheetSheet,
      transferCancelOf,
      unitCodeFromMovement,
      wasMovementProcessed,
      laterReturnSupersedesOutbound,
    });
    ops.push(...excludeTransfer.ops);
    if (excludeTransfer.handled) return;

    const positionChange = planPositionChangeMovementOps({
      event,
      absenceSpansBeforeEpisode,
      alreadyVacatedForAbsence,
      archiveAll,
      arrivalOf,
      augustAbsenceSpansFor,
      byPersonName,
      cancelledTransferOf,
      chainedPositionRows,
      clipAbsenceSpansToActiveEpisode,
      createMovementKey,
      dateMs,
      dayById,
      dayByIndex,
      ejoosDays,
      ejoosOos,
      encodeTimesheetAbsenceSpans,
      extractTimesheetDestinationFromPosition,
      findLatestExcludedRow,
      findMovementExcludedRow,
      findNonStaffOccupantExcelRow,
      inboundStaffDateFor,
      isInternalStaffIndexHop,
      isOwnUnitStaffMove,
      isPositionIndex,
      isPositionChangeRow,
      isSamePerson,
      isVacantStaffRow,
      journalDayFromDateMs,
      journalMonthStartLabel,
      latestPositionByName,
      leadWindowStart,
      movementPersonKey,
      normKey,
      oosById,
      oosByName,
      openAbsenceOf,
      opId,
      ownUnitIndexHistoryOf,
      ownUnitMoveSuperseded,
      positionChainByPerson,
      positionChangeDestination,
      positionCloseWritesExcluded,
      priorEpisodeTimesheetOf,
      shPersonById,
      shPersonByName,
      shpoByIndex,
      shpoPersonById,
      shpoPersonByName,
        shpoSheet,
      staffAppointmentDateFor,
      staffIndexTimesheetForPerson,
      timesheetEpisodeStartFor,
      timesheetNeedsTransferCancelSplit,
      timesheetPeople,
      timesheetRowsOf,
        timesheetSheet,
      transferCancelOf,
      wasMovementProcessed,
    });
    ops.push(...positionChange.ops);
    if (positionChange.handled) return;

    const manualArrivalOp = planManualArrivalMovementOp({
      event,
      latestPositionByName,
      personStillInEjoos,
      wasMovementProcessed,
    });
    if (manualArrivalOp) ops.push(manualArrivalOp);
  };

  effectiveMovements.forEach(considerMovement);

  const finalExternalTransferByPerson = new Map<string, PbMovement>();
  for (const event of movementsAll) {
    if (event.type !== "ПЕРЕВ" || isOwnUnitStaffMove(event)) continue;
    if (cancelledExternalTransferRows.has(event.excelRow)) continue;
    if (isCancelledMovementRecord(event) || isTransferCancellation(event)) {
      continue;
    }
    const key = movementPersonKey(event);
    if (!key) continue;
    const previous = finalExternalTransferByPerson.get(key);
    if (
      !previous ||
      movementEventTime(event) > movementEventTime(previous) ||
      (movementEventTime(event) === movementEventTime(previous) &&
        event.excelRow > previous.excelRow)
    ) {
      finalExternalTransferByPerson.set(key, event);
    }
  }
  const finalExternalTransferFor = (personId: string, fullName: string) => {
    if (personId) {
      const byId = finalExternalTransferByPerson.get(`id:${personId}`);
      if (byId) return byId;
    }
    for (const event of finalExternalTransferByPerson.values()) {
      if (isSamePerson({ personId, fullName }, event)) return event;
    }
    return fullName
      ? (finalExternalTransferByPerson.get(`name:${normKey(fullName)}`) ?? null)
      : null;
  };
  const excludeAlreadyPlanned = (personId: string, fullName: string) =>
    ops.some(
      (op) =>
        op.kind === "exclude_transfer" &&
        isSamePerson({ personId, fullName }, op),
    );
  // Людина ще в ШПО/ООС, в sh уже немає, останній ПЕРЕВ — зовнішній.
  // Статус РУХ «В СТРОЮ» ігноруємо: вирішальні ТИП + в/ч А#### + нема в sh.
  {
    const seen = new Set<number>();
    for (const row of ejoosOccupied) {
      if (personStillInSh(row.personId, row.fullName)) continue;
      if (excludeAlreadyPlanned(row.personId, row.fullName)) continue;
      const transfer = finalExternalTransferFor(row.personId, row.fullName);
      if (
        absenceOnlyBlocksExclusion({
          absenceAt: openDispositionAbsenceMs(row.personId, row.fullName),
          outboundAt: transfer ? movementEventTime(transfer) : 0,
        })
      ) {
        continue;
      }
      if (!transfer) continue;
      if (!eventInLeadWindow(transfer)) continue;
      if (cancelledExternalTransferRows.has(transfer.excelRow)) continue;
      if (seen.has(transfer.excelRow)) continue;
      seen.add(transfer.excelRow);
      considerMovement(transfer);
    }
  }
  const latestOutboundTransferOf = (personId: string, fullName: string) =>
    [...activeMovementsAll]
      .filter(
        (event) =>
          event.type === "ПЕРЕВ" &&
          isOutboundStaffMove(event) &&
          !cancelledExternalTransferRows.has(event.excelRow) &&
          isSamePerson({ personId, fullName }, event),
      )
      .sort(
        (left, right) => movementEventTime(right) - movementEventTime(left),
      )[0] ?? null;
  planTimesheetDayCleanup({
    ops,
    timesheetPeople,
    ejoosExcluded,
    dayByIndex,
    timesheetScanByRow,
    shPeople,
    shPersonById,
    shPersonByName,
    excludeAlreadyPlanned,
    finalExternalTransferFor,
    considerMovement,
    personStillInSh,
    transferCancelForPerson,
    eventInLeadWindow,
    absenceOnlyBlocksExclusion,
    openDispositionAbsenceMs,
    movementEventTime,
    timesheetRowsOf,
    excludedRowsToClear,
    staleExcludedForMovement,
    staleExcludedClearPayload,
    latestOutboundTransferOf,
    activeTimesheetRowOf,
    byPersonName,
    personStillInEjoos,
    isSamePerson,
    movementPersonKey,
  });

  for (const [personId, variants] of nameVariantsById) {
    if (variants.size < 2) continue;
    const displayNames = [...variants.values()].map(
      (variant) => `${variant.display} (${[...variant.sources].join(", ")})`,
    );
    const relevant =
      personStillInSh(personId, "") ||
      personStillInEjoos(personId, "") ||
      activeArrivalIds.has(personId) ||
      openAbsentIds.has(personId);
    if (!relevant) continue;
    const surnames = new Set(
      [...variants.keys()].map((key) => key.split(" ")[0]).filter(Boolean),
    );
    const differentPeople = surnames.size > 1;
    ops.push({
      id: opId(["data-mismatch", personId]),
      kind: "data_mismatch",
      class: "needs_input",
      sheet: "Дані джерел",
      personId,
      fullName:
        shPersonById.get(personId)?.fullName ||
        oosPersonById.get(personId)?.fullName ||
        shpoPersonById.get(personId)?.fullName ||
        "",
      rank: shPersonById.get(personId)?.rank || "",
      positionIndex: shPersonById.get(personId)?.positionIndex || "",
      before: differentPeople
        ? `ID ${personId}: один ID у різних осіб`
        : `ID ${personId}: різне написання ПІБ`,
      after: displayNames.join(" · "),
      sourceRef: `ID ${personId}`,
      why: differentPeople
        ? "Один ID зустрічається у різних прізвищ — перевірте, де ID вказано помилково. Жодних змін у ЕЖООС ця позначка не виконує."
        : "Один ID має різне написання ПІБ у джерелах. Людину зв'язано по ID; написання треба уніфікувати вручну. Жодних змін у ЕЖООС ця позначка не виконує.",
      confidence: "manual",
      payload: {
        type: "DATA_MISMATCH",
        personId,
        mismatchKind: differentPeople ? "ID_COLLISION" : "NAME_SPELLING",
        nameVariants: displayNames.join(" | "),
      },
      checkedDefault: false,
    });
  }

  const rankVariantsById = new Map<
    string,
    Map<string, { display: string; sources: Set<string> }>
  >();
  const addRankVariant = (personId: string, rank: string, source: string) => {
    const id = normId(personId);
    const display = norm(rank);
    const key = normKey(display);
    if (!id || !key) return;
    const variants =
      rankVariantsById.get(id) ??
      new Map<string, { display: string; sources: Set<string> }>();
    const variant = variants.get(key) ?? { display, sources: new Set() };
    variant.sources.add(source);
    variants.set(key, variant);
    rankVariantsById.set(id, variants);
  };
  for (const person of shPeople) {
    addRankVariant(person.personId, person.rank, "sh");
  }
  const latestRankByPerson = new Map<string, string>();
  for (const event of activeMovementsAll) {
    if (!event.personId || !event.rank) continue;
    if (isRankAssignmentEvent(event)) {
      latestRankByPerson.set(event.personId, event.rank);
      continue;
    }
    if (!latestRankByPerson.has(event.personId)) {
      latestRankByPerson.set(event.personId, event.rank);
    }
  }
  // Актуальна sh — канон на дату знімка; колонка «звання» в рядках ПОСАДА
  // часто лишається від старого присвоєння і не має давати «різне звання».
  for (const person of shPeople) {
    if (!person.personId || !person.rank) continue;
    latestRankByPerson.set(person.personId, person.rank);
  }
  for (const [personId, rank] of latestRankByPerson) {
    addRankVariant(personId, rank, "Рух");
  }
  for (const row of ejoosShpo) addRankVariant(row.personId, row.rank, "ШПО");
  for (const row of ejoosOos) addRankVariant(row.personId, row.rank, "ООС");

  for (const [personId, variants] of rankVariantsById) {
    if (variants.size < 2) continue;
    const relevant =
      personStillInSh(personId, "") ||
      personStillInEjoos(personId, "") ||
      activeArrivalIds.has(personId) ||
      openAbsentIds.has(personId);
    if (!relevant) continue;
    if (
      ops.some(
        (op) =>
          op.kind === "rank_change" &&
          (op.personId === personId ||
            isSamePerson({ personId, fullName: "" }, op)),
      )
    ) {
      continue;
    }
    const authoritativeRankKeys = new Set<string>();
    for (const [key, variant] of variants) {
      if ([...variant.sources].some((source) => source !== "Рух")) {
        authoritativeRankKeys.add(key);
      }
    }
    if (authoritativeRankKeys.size === 1) continue;
    const displayRanks = [...variants.values()].map(
      (variant) => `${variant.display} (${[...variant.sources].join(", ")})`,
    );
    ops.push({
      id: opId(["rank-mismatch", personId]),
      kind: "data_mismatch",
      class: "needs_input",
      sheet: "Дані джерел",
      personId,
      fullName:
        shPersonById.get(personId)?.fullName ||
        oosPersonById.get(personId)?.fullName ||
        shpoPersonById.get(personId)?.fullName ||
        "",
      rank: shPersonById.get(personId)?.rank || "",
      positionIndex: shPersonById.get(personId)?.positionIndex || "",
      before: `ID ${personId}: різне звання`,
      after: displayRanks.join(" · "),
      sourceRef: `ID ${personId}`,
      why: "Звання в sh/Рух не збігається зі званням в ООС/ШПО. Автоматично з sh не підставляємо — перевірте наказ про присвоєння. Жодних змін у ЕЖООС ця позначка не виконує.",
      confidence: "manual",
      payload: {
        type: "DATA_MISMATCH",
        personId,
        mismatchKind: "RANK",
        rankVariants: displayRanks.join(" | "),
      },
      checkedDefault: false,
    });
  }
  ops.push(
    ...planTempArrivalCloseOps({
      shPeople,
      movementsAll,
      shpoByIndex,
      ejoosOos,
      dayByIndex,
      oosById,
      oosByName,
      journalMonthStartLabel,
      existingOps: ops,
      arrivalOf,
      isSamePerson,
      eventInLeadWindow,
      movementEventTime,
      isPositionIndex,
      byPersonName,
      staffIndexTimesheetForPerson,
      staffAppointmentDateFor,
      timesheetEpisodeStartFor,
      inboundStaffDateFor,
      staffEpisodePaintPayload,
    }),
  );

  ops.push(
    ...planShpoReconcileOps({
      shOccupantByIndex,
      existingOps: ops,
      activeMovementsAll,
      shpoByIndex,
      dayByIndex,
      arrivalById,
      arrivalByName,
      alreadyVacatedForAbsence,
      byPersonName,
      staffIndexTimesheetForPerson,
      eventInLeadWindow,
      movementEventTime,
      staffEpisodePaintPayload,
      isSamePerson,
    }),
  );

  for (const person of shPeople) {
    if (!personStillInSh(person.personId, person.fullName)) continue;
    if (mapStatus(person.status).ruleId !== "absent_archive") continue;
    const foundInArchive = archiveAll.some((period) =>
      isSamePerson(person, period),
    );
    if (foundInArchive) continue;
    ops.push({
      id: opId(["archive_missing", person.personId || person.fullName]),
      kind: "data_mismatch",
      class: "needs_input",
      sheet: "Дані джерел / archive",
      personId: person.personId,
      fullName: person.fullName,
      positionIndex: person.positionIndex,
      rank: person.rank,
      before: person.status || "ВІДСУТНІЙ в АРХІВІ",
      after: "ARCHIVE_REFERENCE_MISSING → перевірити archive",
      sourceRef: `sh!R${person.excelRow} СТАТУС=«${person.status}»`,
      why: "sh каже «ВІДСУТНІЙ в АРХІВІ», але в archive немає запису за ПІБ чи ID. ЛІК / ВІД / СЗЧ / БЕЗВІСТИ з цього статусу не вигадуємо.",
      confidence: "review",
      payload: {
        type: "ARCHIVE_REFERENCE_MISSING",
        mismatchKind: "ARCHIVE_REFERENCE_MISSING",
        statusRaw: person.status,
      },
      checkedDefault: false,
    });
  }

  for (const op of ops) {
    if (op.class !== "ready" || op.kind !== "position_change") continue;
    const occupant = findUnvacatedTargetOccupant(op, ejoosShpo, ops);
    if (!occupant) continue;
    const targetIndex = op.payload.nextIndex || op.positionIndex;
    op.class = "conflict";
    op.checkedDefault = false;
    op.confidence = "manual";
    op.before = `${op.before} · ціль ${targetIndex} зайнята: ${occupant.rank || "?"} ${occupant.fullName || "без ПІБ"} (ID ${occupant.personId || "—"})`;
    op.why =
      `КОНФЛІКТ ЗАЙНЯТОСТІ: індекс ${targetIndex} уже займає ${occupant.fullName || `ID ${occupant.personId}`}. ` +
      "Для цієї особи немає готової операції вибуття, переміщення з посади або переведення у розпорядження. Спочатку потрібно подати й провести її звільнення з індексу.";
    op.payload.targetOccupancyConflict = "1";
    op.payload.targetOccupantName = occupant.fullName;
    op.payload.targetOccupantId = occupant.personId;
    op.payload.targetOccupantShpoRow = String(occupant.excelRow);
  }

  for (const group of findDuplicateOosById(ejoosOos)) {
    const id = group[0]?.personId || "";
    ops.push({
      id: opId(["dup-oos", id]),
      kind: "data_mismatch",
      class: "conflict",
      sheet: "2. ООС",
      personId: id,
      fullName: group[0]?.fullName || "",
      positionIndex: group[0]?.positionIndex || "",
      rank: group[0]?.rank || "",
      before: group.map((row) => `R${row.excelRow}`).join(", "),
      after: "DUPLICATE_OOS",
      sourceRef: `2. ООС ID ${id}`,
      why: `ID ${id} має більше ніж одну активну картку ООС (${group
        .map((row) => `R${row.excelRow}`)
        .join(", ")}). Apply заблоковано, доки дубль не прибрано.`,
      confidence: "review",
      payload: {
        type: "DUPLICATE_OOS",
        mismatchKind: "DUPLICATE_OOS_ID",
      },
      checkedDefault: false,
    });
  }

  const summary = {
    ready: ops.filter((op) => op.class === "ready").length,
    needsInput: ops.filter((op) => op.class === "needs_input").length,
    conflict: ops.filter((op) => op.class === "conflict").length,
  };

  return {
    ejoosName: ejoos.fileName,
    pbName: pb.fileName,
    timesheetDay,
    timesheetDayLabel,
    ops,
    summary,
    monthRolloverRequired: timesheetMonthHeaderMismatch,
    sourceDateUnknown,
    timesheetMonthHeaderUnknown,
    ejoosTimesheetMonthLabel: timesheetHeaderCell?.matched || "",
    limitsNote:
      `${
        sourceDateUnknown
          ? SOURCE_DATE_UNKNOWN_MESSAGE
          : sourceTimesheetHorizonNote(timesheetDayLabel)
      } ` +
      `РУХ: перевірено всі ${movementsAll.length} рядків, активних ` +
      `${activeMovementsAll.length}, останніх кадрових подій ` +
      `${effectiveMovements.length}` +
      (cancelledExternalTransferRows.size
        ? `, скасованих переведень ${cancelledExternalTransferRows.size}`
        : "") +
      `; archive — ${archive.length} періодів цього місяця ` +
      `з ${archiveAll.length}.` +
      (timesheetMonthHeaderUnknown
        ? ` ${TIMESHEET_MONTH_HEADER_UNKNOWN_MESSAGE}`
        : "") +
      (timesheetMonthHeaderMismatch && timesheetMonthHeader
        ? ` MONTH_ROLLOVER: у шаблоні «${timesheetHeaderCell?.matched}», 1ПБ — «${timesheetMonthHeader}». Місяць беремо з 1ПБ, заголовок не перейменовуємо.`
        : ""),
  };
};

export const buildConfirmSummary = (ops: EjoosSyncOp[]) => {
  const byKind: Record<string, number> = {};
  const bySheet: Record<string, number> = {};
  ops.forEach((op) => {
    byKind[op.kind] = (byKind[op.kind] ?? 0) + 1;
    bySheet[op.sheet] = (bySheet[op.sheet] ?? 0) + 1;
  });
  const kindLabels: Record<string, string> = {
    timesheet_day: "Табель (день)",
    shpo_occupant: "ШПО / зайнятість",
    absent_upsert: "Тимч. відсутні (новий/оновлення)",
    absent_close: "Тимч. відсутні (закриття)",
    exclude_transfer: "Виключені",
    move_to_disposition: "Переміщення у розпорядження",
    data_mismatch: "Помилка даних (ПІБ / ID / звання)",
    position_change: "Зміна посади",
    rank_change: "Присвоєння звання",
    arrival: "Прибуття",
    other_manual: "Інше",
  };
  return {
    total: ops.length,
    byKind: Object.entries(byKind).map(([kind, count]) => ({
      kind,
      label: kindLabels[kind] || kind,
      count,
    })),
    bySheet: Object.entries(bySheet).map(([sheet, count]) => ({
      sheet,
      count,
    })),
    names: ops.map((op) => op.fullName || op.personId || "—").filter(Boolean),
  };
};

export const buildProtocolText = (
  plan: EjoosSyncPlan,
  applied: EjoosSyncOp[],
  meta: { version?: number; actor?: string; at?: string },
) => {
  const confirm = buildConfirmSummary(applied);
  const lines = [
    `Протокол змін ЕЖООС`,
    `ЕЖООС: ${plan.ejoosName}`,
    `1ПБ: ${plan.pbName}`,
    `День табеля: ${plan.timesheetDayLabel} (день ${plan.timesheetDay})`,
    `Версія після застосування: ${meta.version ?? "—"}`,
    `Хто: ${meta.actor ?? "—"}`,
    `Коли: ${meta.at ?? new Date().toLocaleString("uk-UA")}`,
    `Застосовано змін: ${applied.length}`,
    ``,
    `Підсумок за типом:`,
    ...confirm.byKind.map((item) => `  - ${item.label}: ${item.count}`),
    ``,
    `Підсумок за аркушем:`,
    ...confirm.bySheet.map((item) => `  - ${item.sheet}: ${item.count}`),
    ``,
    `Деталі:`,
    ...applied.map(
      (op, index) =>
        `${index + 1}. [${op.sheet}] ${op.fullName || "—"} (ID ${op.personId || "—"}, індекс ${op.positionIndex || "—"})\n` +
        `   Було: ${op.before}\n` +
        `   Стане: ${op.after}\n` +
        `   Джерело: ${op.sourceRef}\n` +
        `   Чому: ${op.why}`,
    ),
  ];
  return lines.join("\n");
};

export type { EjoosTimesheetCode };
