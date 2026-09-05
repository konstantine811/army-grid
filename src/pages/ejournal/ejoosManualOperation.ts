import type { ExcelSheetSnapshot, ExcelWorkbookSnapshot } from "../../excelRoundTrip";
import type { BackendEjournalManualOperation } from "../../api";
import {
  findUnvacatedTargetOccupant,
  parseEjoosOos,
  parseEjoosShpo,
  parseEjoosTimesheetDay,
  parseEjoosTimesheetPeople,
  type EjoosSyncOp,
} from "./ejoosSyncPlan";
import { canonicalName } from "./ejoosIdentity";

export type ManualEjoosOperationType =
  | "exclude_transfer"
  | "dismissal"
  | "position_change"
  | "rank_change";

export type ManualEjoosPerson = {
  key: string;
  personId: string;
  fullName: string;
  rank: string;
  positionIndex: string;
};

export type ManualEjoosOperationInput = {
  type: ManualEjoosOperationType;
  personKey: string;
  orderNumber: string;
  orderDate: string;
  destination?: string;
  nextPositionIndex?: string;
  nextRank?: string;
};

export const manualInputFromBackend = (
  draft: BackendEjournalManualOperation,
): ManualEjoosOperationInput | null => {
  const input = draft.input as Partial<ManualEjoosOperationInput>;
  if (
    !input ||
    !["exclude_transfer", "dismissal", "position_change", "rank_change"].includes(
      String(input.type),
    ) ||
    !input.personKey
  ) {
    return null;
  }
  return {
    type: input.type as ManualEjoosOperationType,
    personKey: String(input.personKey),
    orderNumber: String(input.orderNumber || ""),
    orderDate: String(input.orderDate || ""),
    destination: String(input.destination || ""),
    nextPositionIndex: String(input.nextPositionIndex || ""),
    nextRank: String(input.nextRank || ""),
  };
};

export const hydrateManualEjoosOperation = (input: {
  draft: BackendEjournalManualOperation;
  ejoos: ExcelWorkbookSnapshot;
  timesheetDay: number;
  existingOps?: EjoosSyncOp[];
  currentVersionId?: string;
}) => {
  const values = manualInputFromBackend(input.draft);
  if (!values || input.draft.status !== "draft") return null;
  const op = buildManualEjoosOperation({
    ejoos: input.ejoos,
    timesheetDay: input.timesheetDay,
    values,
    existingOps: input.existingOps,
  });
  const stale = Boolean(
    input.currentVersionId &&
      input.draft.baseVersionId !== input.currentVersionId,
  );
  return {
    ...op,
    id: input.draft.id,
    class: stale ? ("needs_input" as const) : op.class,
    checkedDefault: stale ? false : op.checkedDefault,
    why: stale
      ? `${op.why}. Чернетку створено для попередньої версії ЕЖООС; відкрийте редагування і збережіть її повторно`
      : op.why,
    payload: {
      ...op.payload,
      manualDraftId: input.draft.id,
      manualDecision: input.draft.decision,
      manualBaseVersionId: input.draft.baseVersionId,
      manualCreatedBy:
        input.draft.createdByDisplayName ||
        input.draft.createdByEmail ||
        "",
      manualUpdatedAt: input.draft.updatedAt,
    },
  };
};

const findSheet = (
  workbook: ExcelWorkbookSnapshot,
  pattern: RegExp,
): ExcelSheetSnapshot | undefined =>
  workbook.sheets.find((sheet) => pattern.test(sheet.sheetName));

const samePerson = (
  person: Pick<ManualEjoosPerson, "personId" | "fullName">,
  row: { personId: string; fullName: string },
) =>
  Boolean(
    (person.personId && row.personId && person.personId === row.personId) ||
      (person.fullName &&
        row.fullName &&
        canonicalName(person.fullName) === canonicalName(row.fullName)),
  );

const manualId = (input: ManualEjoosOperationInput, person: ManualEjoosPerson) =>
  [
    "manual",
    input.type,
    person.personId || canonicalName(person.fullName),
    input.orderNumber.trim(),
    input.orderDate.trim(),
    Date.now().toString(36),
  ].join(":");

const ukDate = (value: string) => {
  const match = value.trim().match(/^(\d{4})-(\d{2})-(\d{2})$/);
  return match ? `${match[3]}.${match[2]}.${match[1]}` : value.trim();
};

export const collectManualEjoosPeople = (
  ejoos: ExcelWorkbookSnapshot,
): ManualEjoosPerson[] => {
  const shpo = parseEjoosShpo(findSheet(ejoos, /шпо|штатно.?посад/i));
  const oos = parseEjoosOos(findSheet(ejoos, /(^|[.\s])оос($|[\s])/i));
  const result = new Map<string, ManualEjoosPerson>();
  for (const row of [...oos, ...shpo]) {
    if (!row.fullName && !row.personId) continue;
    const key = row.personId
      ? `id:${row.personId}`
      : `name:${canonicalName(row.fullName)}`;
    const current = result.get(key);
    result.set(key, {
      key,
      personId: row.personId || current?.personId || "",
      fullName: row.fullName || current?.fullName || "",
      rank: row.rank || current?.rank || "",
      positionIndex: row.positionIndex || current?.positionIndex || "",
    });
  }
  return [...result.values()].sort((left, right) =>
    left.fullName.localeCompare(right.fullName, "uk"),
  );
};

export const buildManualEjoosOperation = (input: {
  ejoos: ExcelWorkbookSnapshot;
  timesheetDay: number;
  values: ManualEjoosOperationInput;
  existingOps?: EjoosSyncOp[];
}): EjoosSyncOp => {
  const { ejoos, timesheetDay, values } = input;
  const people = collectManualEjoosPeople(ejoos);
  const person = people.find((item) => item.key === values.personKey);
  if (!person) throw new Error("Оберіть службовця з поточного ЕЖООС");

  const orderNumber = values.orderNumber.trim().replace(/^№\s*/u, "");
  const orderDate = ukDate(values.orderDate);
  const shpoRows = parseEjoosShpo(findSheet(ejoos, /шпо|штатно.?посад/i));
  const oosRows = parseEjoosOos(findSheet(ejoos, /(^|[.\s])оос($|[\s])/i));
  const timesheetSheet = findSheet(ejoos, /табель/i);
  const timesheetRows = parseEjoosTimesheetPeople(timesheetSheet);
  const timesheetDayRows = parseEjoosTimesheetDay(timesheetSheet, timesheetDay);
  const shpo = shpoRows.find((row) => samePerson(person, row)) ?? null;
  const oos = oosRows.find((row) => samePerson(person, row)) ?? null;
  const activeTimesheet =
    timesheetRows.find(
      (row) => samePerson(person, row) && !row.hasDepartureText,
    ) ?? null;
  const currentPosition =
    shpo?.positionIndex ||
    oos?.positionIndex ||
    activeTimesheet?.positionIndex ||
    person.positionIndex;
  const commonReady = Boolean(orderNumber && orderDate);
  const base: EjoosSyncOp = {
    id: manualId(values, person),
    kind: values.type === "dismissal" ? "exclude_transfer" : values.type,
    class: "needs_input",
    sheet: "",
    personId: shpo?.personId || oos?.personId || person.personId,
    fullName: shpo?.fullName || oos?.fullName || person.fullName,
    positionIndex: currentPosition,
    rank: shpo?.rank || oos?.rank || person.rank,
    before: "",
    after: "",
    sourceRef: "Введено вручну оператором",
    why: "",
    confidence: "manual",
    payload: {
      manualOperation: "1",
      orderNumber,
      orderDate,
    },
    checkedDefault: false,
  };

  if (values.type === "exclude_transfer" || values.type === "dismissal") {
    const isDismissal = values.type === "dismissal";
    const destination = values.destination?.trim() || "";
    const ready = Boolean(
      commonReady &&
        destination &&
        activeTimesheet &&
        (shpo || oos) &&
        currentPosition,
    );
    return {
      ...base,
      class: ready ? "ready" : activeTimesheet ? "needs_input" : "conflict",
      sheet: "3. Виключені → 6. Табель → 1. ШПО / 2. ООС",
      before: `${base.rank || "?"} ${base.fullName} · інд. ${currentPosition || "—"}`,
      after: `виключити: ${isDismissal ? "ЗВІЛЬН" : "ПЕРЕВ"} → ${destination || "(вкажіть підставу / куди)"} · дата ${orderDate || "?"}`,
      why: !activeTimesheet
        ? "Ручний ПЕРЕВ заблоковано: не знайдено активного рядка особи в Табелі"
        : ready
          ? `Ручний ${isDismissal ? "ЗВІЛЬН" : "ПЕРЕВ"}: перенести в «Виключені», створити історію Табеля та очистити ШПО/ООС`
          : "Заповніть підставу/місце вибуття, номер і дату наказу",
      payload: {
        ...base.payload,
        type: isDismissal ? "ЗВІЛЬН" : "ПЕРЕВ",
        destination,
        documentsDest: destination,
        timesheetDestination: destination,
        excludeDate: orderDate,
        exclusionReason: isDismissal ? "ЗВІЛЬНЕННЯ" : "ПЕРЕВЕДЕННЯ",
        fromRank: base.rank,
        fromName: base.fullName,
        fromPersonId: base.personId,
        fromPositionIndex: currentPosition,
        shpoExcelRow: shpo ? String(shpo.excelRow) : "",
        oosExcelRow: oos ? String(oos.excelRow) : "",
        timesheetExcelRow: activeTimesheet
          ? String(activeTimesheet.excelRow)
          : "",
        timesheetAction: activeTimesheet ? "MOVE_TO_HISTORY" : "",
        timesheetSourceIndex: currentPosition,
        timesheetSourceRow: activeTimesheet
          ? `R${activeTimesheet.excelRow}`
          : "",
      },
      checkedDefault: ready,
    };
  }

  if (values.type === "rank_change") {
    const nextRank = values.nextRank?.trim() || "";
    const ready = Boolean(
      commonReady && nextRank && (shpo || oos || activeTimesheet),
    );
    return {
      ...base,
      class: ready ? "ready" : "needs_input",
      sheet: "1. ШПО / 2. ООС / 6. Табель",
      before: base.rank || "—",
      after: `${nextRank || "(вкажіть звання)"} · наказ №${orderNumber || "?"} від ${orderDate || "?"}`,
      why: ready
        ? "Ручне присвоєння звання в усіх активних аркушах ЕЖООС"
        : "Вкажіть нове звання, номер і дату наказу",
      rank: nextRank || base.rank,
      payload: {
        ...base.payload,
        previousRank: base.rank,
        nextRank,
        shpoExcelRow: shpo ? String(shpo.excelRow) : "",
        oosExcelRow: oos ? String(oos.excelRow) : "",
        timesheetExcelRow: activeTimesheet
          ? String(activeTimesheet.excelRow)
          : "",
      },
      checkedDefault: ready,
    };
  }

  const nextIndex = values.nextPositionIndex?.trim() || "";
  const targetShpo =
    shpoRows.find((row) => row.positionIndex === nextIndex) ?? null;
  const targetTimesheet =
    timesheetDayRows.find((row) => row.positionIndex === nextIndex) ?? null;
  const ready = Boolean(
    commonReady &&
      currentPosition &&
      nextIndex &&
      nextIndex !== currentPosition &&
      targetShpo &&
      targetTimesheet,
  );
  const op: EjoosSyncOp = {
    ...base,
    class: ready ? "ready" : "needs_input",
    sheet: "1. ШПО / 2. ООС / 6. Табель",
    positionIndex: nextIndex || currentPosition,
    before: `штатна посада ${currentPosition || "—"}`,
    after: `штатна посада ${nextIndex || "(вкажіть індекс)"}`,
    why: ready
      ? `Ручна зміна посади ${currentPosition} → ${nextIndex}`
      : "Вкажіть інший чинний індекс ШПО, номер і дату наказу",
    payload: {
      ...base.payload,
      previousIndex: currentPosition,
      fromPositionIndex: currentPosition,
      nextIndex,
      nextName: base.fullName,
      nextRank: base.rank,
      nextPersonId: base.personId,
      oosExcelRow: oos ? String(oos.excelRow) : "",
      shpoExcelRow: targetShpo ? String(targetShpo.excelRow) : "",
      timesheetExcelRow: targetTimesheet
        ? String(targetTimesheet.excelRow)
        : "",
      closeOldPosition: currentPosition !== nextIndex ? "1" : "",
      internalStaffHop: "1",
      previousShpoExcelRow: shpo ? String(shpo.excelRow) : "",
      previousIndexTimesheetExcelRow: activeTimesheet
        ? String(activeTimesheet.excelRow)
        : "",
      timesheetActiveFrom: orderDate,
      timesheetSkipHistory: "1",
    },
    checkedDefault: ready,
  };
  const occupant = findUnvacatedTargetOccupant(op, shpoRows, [
    ...(input.existingOps ?? []),
    op,
  ]);
  if (occupant) {
    op.class = "conflict";
    op.checkedDefault = false;
    op.why = `КОНФЛІКТ: індекс ${nextIndex} займає ${occupant.fullName || `ID ${occupant.personId}`}; немає готової операції його вибуття`;
    op.payload.targetOccupancyConflict = "1";
    op.payload.targetOccupantName = occupant.fullName;
    op.payload.targetOccupantId = occupant.personId;
  }
  return op;
};
