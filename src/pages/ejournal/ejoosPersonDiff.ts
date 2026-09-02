import type { EjoosOpClass, EjoosOpKind, EjoosSyncOp, EjoosSyncPlan } from "./ejoosSyncPlan";
import { parsePbShPeople } from "./ejoosSyncPlan";
import { mapPbStatusToEjoosWithRules, readOperatorSettings } from "./ejoosStatusMap";
import { formatTimesheetTransferMark } from "./ejoosExcludedColumns";
import {
  dayFromOrderLabel,
  parseTimesheetAbsenceSpans,
  timesheetMarkFromArchive,
  timesheetTransferMarkForDay,
} from "./ejoosTimesheetText";
import { excludeWritePlan, positionCloseWritesExcluded } from "./ejoosExcludePolicy";
import { personOpsBlockApply } from "./ejoosOpRequirements";
import type { ExcelWorkbookSnapshot } from "../../excelRoundTrip";

const formatArchiveSpanSummary = (raw: string) =>
  parseTimesheetAbsenceSpans(raw)
    .map((span) => `${span.fromDay}–${span.toDay}:${span.code}`)
    .join(", ");

const staffEpisodeTimesheetDetail = (
  payload: Record<string, string>,
  positionIndex: string,
) => {
  const from = payload.timesheetActiveFrom;
  const index = payload.nextIndex || positionIndex;
  const history = payload.historyTimesheetExcelRow
    ? `; коди відсутності на історичному рядку R${payload.historyTimesheetExcelRow}`
    : "";
  if (!from) return "";
  const twoEpisodes =
    payload.returningToStaffIndex === "1" ||
    Boolean(payload.historyTimesheetExcelRow) ||
    Boolean(payload.previousTimesheetExcelRow);
  return twoEpisodes
    ? `Два епізоди Табеля: історичний «вибув» лишаємо; новий ${index || "штатний рядок"} — до ${from} «-», з ${from} «+»${history}`
    : `Табель ${index || "штатний рядок"}: до ${from} «-», з ${from} «+». Коди відсутності на цей епізод не переносимо${history}`;
};

export type PersonChangeCategory =
  | "status"
  | "position"
  | "arrival"
  | "data"
  | "error"
  | "mixed";

export type PersonChangeDecision = "pending" | "accepted" | "rejected";

export type PersonSheetAction = {
  sheet: string;
  kind: EjoosOpKind;
  before: string;
  after: string;
  why: string;
  opClass: EjoosOpClass;
  opId: string;
};

/** Візуальний ефект застосування на аркуші ЕЖООС. */
export type SheetImpactEffect =
  | "append"
  | "clear"
  | "edit"
  | "history"
  | "untouched"
  | "skip";

export type SheetImpactItem = {
  sheetKey: string;
  sheetLabel: string;
  effect: SheetImpactEffect;
  effectLabel: string;
  detail: string;
  rowHint?: string;
};

export type PersonSourceKind = "ruh" | "archive" | "sh" | "ejoos";

export type PersonSourceInfluence = {
  source: PersonSourceKind;
  sourceLabel: string;
  ref: string;
  event: string;
  effect: string;
};

export type PersonTimesheetPreview = {
  lastDay: number;
  days: Array<{ day: number; mark: string; title?: string }>;
  runs: Array<{ from: number; to: number; mark: string }>;
  note?: string;
  /** День вибуття в Табелі (1–31). */
  departDay?: number;
  /** Точний текст клітинки: «вибув до МЕХАНІЗОВАНОГО ВІДДІЛЕННЯ …». */
  departPhrase?: string;
};

export type PersonChange = {
  id: string;
  personId: string;
  fullName: string;
  rank: string;
  positionIndex: string;
  category: PersonChangeCategory;
  severity: EjoosOpClass; // ready | needs_input | conflict
  summaryBefore: string;
  summaryAfter: string;
  ejoosWillDo: string[];
  sheetActions: PersonSheetAction[];
  /** Карта впливу на канонічні аркуші ЕЖООС. */
  sheetImpacts: SheetImpactItem[];
  /** Які події 1ПБ (Рух / archive / sh) спричиняють ці ops. */
  sourceInfluences: PersonSourceInfluence[];
  /** Як виглядатиме рядок Табеля після застосування. */
  timesheetPreview: PersonTimesheetPreview | null;
  ops: EjoosSyncOp[];
  decision: PersonChangeDecision;
};

export type EjoosDiffSession = {
  plan: EjoosSyncPlan;
  people: PersonChange[];
  counters: {
    oosLike: number;
    onDuty: number;
    changes: number;
    newcomers: number;
    errors: number;
    autoReady: number;
    needsReview: number;
  };
  pbFileName: string;
  analyzedAt: string;
};

const normKey = (value: string) =>
  value
    .toLowerCase()
    .replace(/[''`´]/g, "")
    .replace(/\s+/g, " ")
    .trim();

const personKey = (op: EjoosSyncOp) => {
  // Дубль Табеля групуємо за ПІБ: у хвості аркуша кілька осіб можуть мати
  // один битий Excel-ID, і тоді всі ops зліпаються в одну картку.
  if (op.payload.type === "DUPLICATE_TAB_ROW" && op.fullName.trim()) {
    return `name:${normKey(op.fullName)}`;
  }
  // ID є первинним ключем: ПІБ у РУХ, archive та ЕЖООС може відрізнятися
  // через помилку або інший варіант написання по батькові.
  if (op.personId.trim()) return `id:${op.personId.trim()}`;
  if (op.fullName.trim()) return `name:${normKey(op.fullName)}`;
  if (op.positionIndex.trim()) return `idx:${op.positionIndex.trim()}`;
  return `op:${op.id}`;
};

/**
 * Позначка помилки даних / очікування наступного кроку лише інформує
 * і не повинна блокувати застосування реальних операцій (ПОСАДА тощо).
 */
const actionableOps = (ops: EjoosSyncOp[]) => {
  const actionable = ops.filter((op) => !isInformationalOp(op));
  return actionable.length ? actionable : ops;
};

const categoryForOps = (allOps: EjoosSyncOp[]): PersonChangeCategory => {
  const ops = actionableOps(allOps);
  const kinds = new Set(ops.map((op) => op.kind));
  if (kinds.has("arrival")) return "arrival";
  if (ops.some((op) => op.class === "conflict")) {
    return "error";
  }
  const statusKinds: EjoosOpKind[] = [
    "timesheet_day",
    "absent_upsert",
    "absent_close",
  ];
  const positionKinds: EjoosOpKind[] = [
    "position_change",
    "shpo_occupant",
    "move_to_disposition",
    "rank_change",
  ];
  const hasStatus = statusKinds.some((k) => kinds.has(k));
  const hasPos = positionKinds.some((k) => kinds.has(k));
  const hasExclude = kinds.has("exclude_transfer");
  if (
    ops.some(
      (op) =>
        op.payload.reviewReason === "CANCEL_TRANSFER_BUT_NOT_IN_CURRENT_SH",
    )
  ) {
    return "mixed";
  }
  if (ops.every((op) => op.class === "needs_input" || op.class === "conflict")) {
    if (!hasStatus && !hasPos && hasExclude) return "error";
  }
  if (hasStatus && !hasPos && !hasExclude) return "status";
  if (hasPos && !hasStatus) return "position";
  if (hasStatus || hasPos || hasExclude) return "mixed";
  return "data";
};

const severityForOps = (allOps: EjoosSyncOp[]): EjoosOpClass => {
  const ops = actionableOps(allOps);
  if (ops.some((op) => op.class === "conflict")) return "conflict";
  if (ops.some((op) => op.class === "needs_input")) return "needs_input";
  return "ready";
};

const toSheetActions = (ops: EjoosSyncOp[]): PersonSheetAction[] =>
  ops.map((op) => ({
    sheet: op.sheet,
    kind: op.kind,
    before: op.before,
    after: op.after,
    why: op.why,
    opClass: op.class,
    opId: op.id,
  }));

const CANONICAL_SHEETS: Array<{ key: string; label: string }> = [
  { key: "shpo", label: "1. ШПО" },
  { key: "oos", label: "2. ООС" },
  { key: "excluded", label: "3. Виключені" },
  { key: "arrivals", label: "4. Тимч. прибулі" },
  { key: "absents", label: "5. Тимч. відсутні" },
  { key: "timesheet", label: "6. Табель" },
  { key: "history", label: "10. Історія змін" },
];

const effectLabelOf = (effect: SheetImpactEffect): string => {
  if (effect === "append") return "Додати рядок";
  if (effect === "clear") return "Очистити";
  if (effect === "edit") return "Змінити";
  if (effect === "history") return "Історія + очистка";
  if (effect === "skip") return "Не чіпаємо";
  return "Без змін";
};

const setImpact = (
  map: Map<string, SheetImpactItem>,
  key: string,
  effect: SheetImpactEffect,
  detail: string,
  rowHint?: string,
) => {
  const base = CANONICAL_SHEETS.find((sheet) => sheet.key === key);
  if (!base) return;
  const prev = map.get(key);
  // skip/untouched не перебивають реальну дію; clear дубля не має ховати edit основного рядка
  if (
    prev &&
    prev.effect !== "untouched" &&
    prev.effect !== "skip" &&
    (effect === "untouched" || effect === "skip")
  ) {
    return;
  }
  if (prev && prev.effect === "edit" && effect === "clear") {
    map.set(key, {
      ...prev,
      detail: `${prev.detail}; ${detail}`,
      rowHint: rowHint || prev.rowHint,
    });
    return;
  }
  map.set(key, {
    sheetKey: key,
    sheetLabel: base.label,
    effect,
    effectLabel: effectLabelOf(effect),
    detail: prev && prev.effect !== "untouched" && prev.effect !== "skip"
      ? `${prev.detail}; ${detail}`
      : detail,
    rowHint: rowHint || prev?.rowHint,
  });
};

/** Будує карту «що станеться на кожному аркуші» для ops бійця. */
export const buildSheetImpacts = (ops: EjoosSyncOp[]): SheetImpactItem[] => {
  const map = new Map<string, SheetImpactItem>();
  const transferCancelOwnsTimesheet = (op: EjoosSyncOp) =>
    ops.some(
      (other) =>
        other.payload.type === "TRANSFER_CANCELLED" &&
        other.payload.restoreTimesheet === "1" &&
        other.payload.reviewReason !==
          "CANCEL_TRANSFER_BUT_NOT_IN_CURRENT_SH" &&
        (other.personId === op.personId ||
          other.fullName.toLocaleLowerCase("uk-UA") ===
            op.fullName.toLocaleLowerCase("uk-UA")),
    );
  CANONICAL_SHEETS.forEach((sheet) => {
    map.set(sheet.key, {
      sheetKey: sheet.key,
      sheetLabel: sheet.label,
      effect: "untouched",
      effectLabel: effectLabelOf("untouched"),
      detail: "Не зачіпається цією зміною",
    });
  });

  const rankChange = ops.find((op) => op.kind === "rank_change");
  if (rankChange) {
    const nextRank = rankChange.payload.nextRank || rankChange.rank;
    const previousRank = rankChange.payload.previousRank || rankChange.before;
    const order = [
      rankChange.payload.orderNumber
        ? `№${rankChange.payload.orderNumber}`
        : "",
      rankChange.payload.orderDate
        ? `від ${rankChange.payload.orderDate}`
        : "",
    ]
      .filter(Boolean)
      .join(" ");
    setImpact(
      map,
      "shpo",
      "edit",
      `${previousRank || "звання"} → ${nextRank}${order ? ` · ${order}` : ""}`,
      rankChange.payload.shpoExcelRow
        ? `R${rankChange.payload.shpoExcelRow}`
        : undefined,
    );
    setImpact(
      map,
      "oos",
      "edit",
      `Звання ${nextRank}${order ? `; останнє присвоєння ${order}` : ""}`,
      rankChange.payload.oosExcelRow
        ? `R${rankChange.payload.oosExcelRow}`
        : undefined,
    );
    setImpact(
      map,
      "timesheet",
      "edit",
      `Звання ${nextRank}`,
      rankChange.payload.timesheetExcelRow
        ? `R${rankChange.payload.timesheetExcelRow}`
        : undefined,
    );
  }

  const exclude = ops.find((op) => op.kind === "exclude_transfer");
  if (exclude) {
    const dest =
      exclude.payload.documentsDest || exclude.payload.destination || "(куди?)";
    const date =
      exclude.payload.excludeDate || exclude.payload.orderDate || "?";
    const departPhrase = formatTimesheetTransferMark(exclude.payload);
    const who = [
      exclude.payload.fromRank || exclude.rank,
      exclude.payload.fromName || exclude.fullName,
      exclude.payload.fromPersonId || exclude.personId,
    ]
      .filter(Boolean)
      .join(" ");
    const idx =
      exclude.payload.fromPositionIndex ||
      exclude.payload.previousIndex ||
      exclude.positionIndex ||
      "—";
    const occupied =
      exclude.payload.occupiedPositionIndex || idx;

    setImpact(
      map,
      "excluded",
      exclude.payload.excludedExcelRow ? "skip" : "append",
      exclude.payload.excludedExcelRow
        ? `Рядок уже є (R${exclude.payload.excludedExcelRow}) — не дублюємо`
        : `Новий рядок: ${who || "—"} · інд. ${idx} → ${dest} · дата ${date}`,
    );
    if (exclude.payload.clearExcludedExcelRow || exclude.payload.clearExcludedExcelRows) {
      const stale = String(
        exclude.payload.clearExcludedExcelRows ||
          exclude.payload.clearExcludedExcelRow ||
          "",
      );
      setImpact(
        map,
        "excluded",
        "clear",
        `Прибрати попередні рядки Виключені R${stale.replaceAll(",", ", R")} — лишаємо чинне ПЕРЕВ`,
        stale ? `R${stale.split(",")[0]}` : undefined,
      );
    }
    setImpact(
      map,
      "timesheet",
      excludeWritePlan(exclude.payload).replaceInPlace
        ? "edit"
        : "history",
      excludeWritePlan(exclude.payload).replaceInPlace
        ? `Замінити наявний рядок Табеля (R${exclude.payload.timesheetExcelRow || "—"}) — новий не створювати`
        : excludeWritePlan(exclude.payload).createTimesheetHistory
        ? `Новий історичний рядок: «+» з ${exclude.payload.timesheetActiveFrom || "дати постановки"} до вибуття ${date}; у день вибуття «${departPhrase}» (посаду прибрано, до/у за підрозділом), далі −; чужий рядок ${idx} не чіпаємо`
        : `Копія рядка вниз у межах роти з R${exclude.payload.timesheetExcelRow || "—"}: +${exclude.payload.timesheetActiveFrom ? ` з ${exclude.payload.timesheetActiveFrom}` : ""} до вибуття ${date}; у день вибуття «${departPhrase}»; штат лишає індекс / ВОС / тариф, особу й дні прибираємо`,
      exclude.payload.timesheetExcelRow
        ? `R${exclude.payload.timesheetExcelRow}`
        : undefined,
    );
    setImpact(
      map,
      "shpo",
      exclude.payload.shpoExcelRow ? "clear" : "skip",
      exclude.payload.shpoExcelRow
        ? occupied && occupied !== idx
          ? `Прибрати звання/ПІБ/ID з індексу ${occupied}, посаду лишити; внутрішню постановку на ${idx} не проводимо`
          : `Прибрати звання/ПІБ/ID на індексі ${idx}, посаду лишити`
        : excludeWritePlan(exclude.payload).transitSameMonth
          ? `У ШПО не був; індекс ${idx} уже зайнятий — не ставимо і не чистимо`
          : `Прибрати звання/ПІБ/ID на індексі ${idx}, посаду лишити`,
      exclude.payload.shpoExcelRow
        ? `R${exclude.payload.shpoExcelRow}`
        : undefined,
    );
    setImpact(
      map,
      "oos",
      exclude.payload.oosExcelRow ? "clear" : "skip",
      exclude.payload.oosExcelRow
        ? `Видалити рядок анкети${exclude.personId ? ` (ID ${exclude.personId})` : ""}`
        : "В активному ООС немає — рядок не видаляємо",
    );
    setImpact(map, "absents", "skip", "Тимч. відсутні не чіпаємо при ПЕРЕВ");
    if (exclude.payload.arrivalExcelRow) {
      setImpact(
        map,
        "arrivals",
        "edit",
        `Історичний рядок лишити й закрити: вибуття ${exclude.payload.arrivalDepartDate || exclude.payload.appointmentOrderDate || "?"} · наказ №${exclude.payload.arrivalDepartOrderNumber || exclude.payload.appointmentOrderNumber || "?"} від ${exclude.payload.arrivalDepartOrderDate || exclude.payload.appointmentOrderDate || "?"}`,
        `R${exclude.payload.arrivalExcelRow}`,
      );
    } else {
      setImpact(map, "arrivals", "skip", "Тимч. прибулі не чіпаємо при ПЕРЕВ");
    }
    setImpact(
      map,
      "history",
      "append",
      "Запис у протокол застосування версії",
    );
  }

  ops.forEach((op) => {
    if (op.kind === "exclude_transfer" || op.kind === "rank_change") return;
    if (exclude && op.kind === "position_change") return;
    if (op.kind === "timesheet_day") {
      if (
        transferCancelOwnsTimesheet(op) &&
        op.payload.type !== "DUPLICATE_TAB_AFTER_CANCEL"
      ) {
        return;
      }
      setImpact(
        map,
        "timesheet",
        "edit",
        op.payload.clearStalePerson === "1"
          ? op.payload.type === "DUPLICATE_TAB_ROW"
            ? `Прибрати дубль R${op.payload.excelRow || "—"}${
                op.payload.keepTimesheetExcelRow
                  ? `; лишається R${op.payload.keepTimesheetExcelRow}`
                  : ""
              }`
          : op.payload.type === "DUPLICATE_TAB_AFTER_CANCEL"
            ? `Прибрати дубль Табеля після скасованого переведення (R${op.payload.excelRow || "—"})`
            : `Прибрати ПІБ/ID зі штатного рядка ${op.positionIndex || "—"}; історичний рядок з вибуттям лишити`
          : op.payload.restorePerson === "1"
          ? `Немає в Табелі — відновити на індексі ${op.positionIndex || "—"}${
              op.payload.excelRow ? "" : " (рядок не знайдено)"
            }`
          : op.payload.type === "PAINT_ARCHIVE"
            ? staffEpisodeTimesheetDetail(op.payload, op.positionIndex) ||
              `Фактичні коди з archive на активному рядку${
                op.payload.timesheetActiveFrom
                  ? ` з ${op.payload.timesheetActiveFrom}`
                  : ""
              }`
          : `День: ${op.before || "—"} → ${op.after || op.payload.timesheetCode || "—"}`,
        op.payload.excelRow
          ? `R${op.payload.excelRow}`
          : op.payload.timesheetExcelRow
            ? `R${op.payload.timesheetExcelRow}`
            : undefined,
      );
      if (op.payload.clearExcludedExcelRow || op.payload.clearExcludedExcelRows) {
        const stale = String(
          op.payload.clearExcludedExcelRows ||
            op.payload.clearExcludedExcelRow ||
            "",
        );
        setImpact(
          map,
          "excluded",
          "clear",
          op.payload.type === "CLEAR_STALE_EXCLUSION_DUPLICATE"
            ? `Прибрати застарілі рядки Виключені R${stale.replaceAll(",", ", R")} — лишаємо чинне ПЕРЕВ`
            : `Прибрати хибний рядок внутрішньої ПОСАДИ 1ПБ R${stale.replaceAll(",", ", R")}`,
          stale ? `R${stale.split(",")[0]}` : undefined,
        );
      }
    } else if (op.kind === "shpo_occupant") {
      setImpact(
        map,
        "shpo",
        "edit",
        `${op.before || "—"} → ${op.after || "—"}`,
        op.payload.shpoExcelRow ? `R${op.payload.shpoExcelRow}` : undefined,
      );
      if (op.payload.timesheetExcelRow) {
        if (!transferCancelOwnsTimesheet(op)) {
          setImpact(
            map,
            "timesheet",
            "edit",
            op.payload.timesheetAbsenceSpans
              ? `Відновити Табель з archive (не писати «СКАСОВАНО» в день)`
              : `Підтягнути особу в табель`,
            `R${op.payload.timesheetExcelRow}`,
          );
        }
      }
      if (op.payload.excludedSourceExcelRow) {
        setImpact(
          map,
          "oos",
          "append",
          `Додати активну облікову картку з рядка «Виключені» R${op.payload.excludedSourceExcelRow}`,
        );
      }
      if (op.payload.clearExcludedExcelRow) {
        setImpact(
          map,
          "excluded",
          "clear",
          `Прибрати застарілий рядок R${op.payload.clearExcludedExcelRow} — особа знову в штаті`,
          `R${op.payload.clearExcludedExcelRow}`,
        );
      }
    } else if (op.kind === "position_change") {
      const closesOldPosition = op.payload.closeOldPosition === "1";
      const writesExcluded = positionCloseWritesExcluded(op.payload);
      if (writesExcluded) {
        setImpact(
          map,
          "excluded",
          "append",
          `Закрити посаду ${op.payload.previousIndex || "—"}: ${op.payload.documentsDest || "нова посада"} · ${op.payload.excludeDate || op.payload.orderDate || "?"} · ${op.payload.exclusionReason || "ПЕРЕВЕДЕННЯ 1 ПБ"}`,
        );
      } else if (op.payload.clearExcludedExcelRow) {
        setImpact(
          map,
          "excluded",
          "clear",
          `Прибрати застарілий рядок R${op.payload.clearExcludedExcelRow} — особа знову в штаті`,
          `R${op.payload.clearExcludedExcelRow}`,
        );
      }
      setImpact(
        map,
        "shpo",
        "edit",
        op.payload.returningFromDisposition === "1"
          ? `Прибрати з блоку «у розпорядженні» і поставити на ${op.payload.nextIndex || op.positionIndex}`
          : closesOldPosition
            ? `Звільнити ${op.payload.previousIndex} і поставити ${op.rank} ${op.fullName} на ${op.payload.nextIndex || op.positionIndex}`
            : `Поставити ${op.rank} ${op.fullName} на індекс ${op.payload.nextIndex || op.positionIndex}`,
        op.payload.shpoExcelRow ? `R${op.payload.shpoExcelRow}` : undefined,
      );
      setImpact(
        map,
        "oos",
        op.payload.oosExcelRow ? "edit" : "append",
        op.payload.oosExcelRow
          ? `Оновити історію посад: ${
              (op.payload.oosHistoryIndexes || "")
                .split("\n")
                .filter(Boolean)
                .join(" → ") ||
              `${op.payload.previousIndex || "—"} → ${op.payload.nextIndex || op.positionIndex}`
            }`
          : op.payload.excludedSourceExcelRow
            ? `Додати активну облікову картку з рядка «Виключені» R${op.payload.excludedSourceExcelRow}`
            : "Додати активну облікову картку та реквізити постановки",
        op.payload.oosExcelRow ? `R${op.payload.oosExcelRow}` : undefined,
      );
      if (!transferCancelOwnsTimesheet(op)) {
        setImpact(
          map,
          "timesheet",
          writesExcluded ? "history" : "edit",
          op.payload.returningFromDisposition === "1"
            ? `Штатний рядок ${op.payload.nextIndex || op.positionIndex}: на початку місяця коди відсутності (СЗЧ), з фактичного повернення «+»${
                op.payload.timesheetAbsenceSpans
                  ? ` (${formatArchiveSpanSummary(op.payload.timesheetAbsenceSpans)})`
                  : ""
              }`
            : writesExcluded
            ? `Закрити рядок ${op.payload.previousIndex} з датою вибуття та поставити на ${op.payload.nextIndex || op.positionIndex} з ${op.payload.orderDate || "дати наказу"}`
            : closesOldPosition
              ? `Один рядок ${op.payload.nextIndex || op.positionIndex}; старий індекс ${op.payload.previousIndex} звільнити без «вибув» у Виключених`
            : op.payload.timesheetPreserveHistory === "1"
              ? `Старий рядок ${op.payload.nextIndex || op.positionIndex} з вибуттям лишити як історію; новий активний з ${op.payload.timesheetActiveFrom || op.payload.orderDate || "дати наказу"}`
              : op.payload.transferCancelOrder || op.payload.transferCancelDate
                ? `Один рядок ${op.payload.nextIndex || op.positionIndex} з ${op.payload.timesheetActiveFrom || op.payload.orderDate || "дати наказу"}; другий рядок після скасування не створюємо`
                : `Поставити на штатну позицію ${op.payload.nextIndex || op.positionIndex} з ${op.payload.orderDate || "дати наказу"}`,
          op.payload.timesheetExcelRow
            ? `R${op.payload.timesheetExcelRow}`
            : undefined,
        );
      }
      if (op.payload.arrivalExcelRow) {
        setImpact(
          map,
          "arrivals",
          "edit",
          `Історичний рядок лишити й закрити: вибуття ${op.payload.arrivalDepartDate || op.payload.orderDate || "?"} · наказ №${op.payload.arrivalDepartOrderNumber || op.payload.orderNumber || "?"} від ${op.payload.arrivalDepartOrderDate || op.payload.orderDate || "?"}`,
          `R${op.payload.arrivalExcelRow}`,
        );
      }
      if (op.payload.transferCancelOrder || op.payload.transferCancelDate) {
        setImpact(
          map,
          "excluded",
          "skip",
          `TRANSFER_CANCELLED №${op.payload.transferCancelOrder || "?"} від ${op.payload.transferCancelDate || "?"}: прибрати запис №${op.payload.cancelledTransferOrder || "?"} з «Виключені»; один рядок Табеля`,
        );
      }
    } else if (op.kind === "absent_upsert" || op.kind === "absent_close") {
      if (op.payload.mismatchKind === "ARCHIVE_RETURN_SH_STILL_ABSENT") {
        setImpact(
          map,
          "absents",
          "skip",
          `Перевірте: archive повернення ${op.payload.returnDate || "?"} vs sh «${op.payload.statusRaw || "СЗЧ"}»`,
        );
        setImpact(
          map,
          "timesheet",
          "skip",
          "Табель не фарбуємо: не відомо, чи archive рано закритий, чи sh застарів",
        );
        return;
      }
      if (op.kind === "absent_upsert" || op.payload.excelRow) {
        setImpact(
          map,
          "absents",
          op.kind === "absent_upsert" ? "append" : "edit",
          `${op.before || "—"} → ${op.after || "—"}`,
          op.payload.existingExcelRow
            ? `R${op.payload.existingExcelRow}`
            : op.payload.excelRow
              ? `R${op.payload.excelRow}`
              : undefined,
        );
      }
      if (op.kind === "absent_close" && op.payload.timesheetExcelRow) {
        setImpact(
          map,
          "timesheet",
          "edit",
          `Закрити відсутність і виставити фактичні коди з ${op.payload.returnDate || op.payload.returnDay || "дати повернення"}`,
          `R${op.payload.timesheetExcelRow}`,
        );
      }
      if (
        op.kind === "absent_upsert" &&
        (op.payload.timesheetAbsenceSpans ||
          (op.payload.timesheetCode && op.payload.departDate))
      ) {
        if (!transferCancelOwnsTimesheet(op)) {
          const episodeDetail = staffEpisodeTimesheetDetail(
            op.payload,
            op.positionIndex,
          );
          setImpact(
            map,
            "timesheet",
            "edit",
            episodeDetail ||
              (op.payload.historyTimesheetExcelRow
                ? `Коди відсутності (${op.payload.timesheetCode || "СЗЧ/лік/…"}) на історичному рядку R${op.payload.historyTimesheetExcelRow}; новий штатний рядок не чіпаємо кодами відсутності`
                : `Період ${op.payload.absenceType || op.payload.timesheetCode || "відсутності"} з ${op.payload.departDate || "?"} закрити в «Тимч. відсутні»; на новий штатний рядок коди не переносимо`),
            op.payload.timesheetExcelRow
              ? `R${op.payload.timesheetExcelRow}`
              : undefined,
          );
        }
      }
    } else if (op.kind === "arrival") {
      setImpact(map, "oos", "append", `Прибуття: ${op.after || "—"}`);
      setImpact(map, "shpo", "edit", "Зайняти посаду після підтвердження");
    } else if (op.kind === "move_to_disposition") {
      setImpact(
        map,
        "shpo",
        "edit",
        op.payload.skipShpoDisposition === "1"
          ? `Звільнити позицію ${op.payload.previousIndex || op.positionIndex}; у блоці розпорядження запис уже є`
          : `Звільнити позицію ${op.payload.previousIndex || op.positionIndex} і додати у блок розпорядження`,
      );
      setImpact(
        map,
        "oos",
        "skip",
        op.payload.remainsInOos === "true"
          ? "Залишити активний запис без змін"
          : "Активного запису немає — не додаємо і не видаляємо",
      );
      setImpact(
        map,
        "absents",
        op.payload.needsAbsenceRecord === "1" ? "append" : "skip",
        op.payload.needsAbsenceRecord === "1"
          ? `Додати ${op.payload.absenceType || "відсутність"} · ${op.payload.absenceDate || "дата з archive"}`
          : `Чинний запис уже є: ${op.payload.absenceType || "відсутність"}${op.personId ? `; гарантувати ID ${op.personId}` : ""}`,
      );
      setImpact(
        map,
        "timesheet",
        op.payload.timesheetFound === "true" ? "edit" : "skip",
        op.payload.keepOpenSzchTimesheet === "1"
          ? `Один рядок, 01–зріз ${op.payload.absenceCode || "СЗЧ"}; 10.08 не писати «вибув у розпорядження»`
          : op.payload.timesheetFound === "true"
            ? "Закрити штатний рядок і зберегти історію розпорядження"
            : "Штатного рядка немає — Табель не змінюємо",
      );
    } else if (op.kind === "data_mismatch" || op.payload.mismatchKind === "ARCHIVE_RETURN_SH_STILL_ABSENT") {
      const rankIssue = op.payload.mismatchKind === "RANK";
      const archiveMissing =
        op.payload.mismatchKind === "ARCHIVE_REFERENCE_MISSING";
      const archiveReturnVsSh =
        op.payload.mismatchKind === "ARCHIVE_RETURN_SH_STILL_ABSENT";
      setImpact(
        map,
        "shpo",
        "skip",
        archiveReturnVsSh
          ? "Не закриваємо штат автоматично: archive каже повернувся, sh досі відсутній"
          : archiveMissing
          ? "Не вигадуємо відсутність у ШПО зі статусу «ВІДСУТНІЙ в АРХІВІ»"
          : rankIssue
            ? "Звання в ШПО не змінюємо автоматично — потрібен наказ про присвоєння"
            : "Помилка написання ПІБ / ID; ШПО не змінюємо",
      );
      setImpact(
        map,
        "oos",
        "skip",
        archiveReturnVsSh
          ? "ООС не чіпаємо, доки не звірите archive і sh"
          : archiveMissing
          ? "Не вигадуємо ЛІК / ВІД / СЗЧ / БЕЗВІСТИ в ООС"
          : rankIssue
            ? "Звання в ООС не змінюємо автоматично — потрібен наказ про присвоєння"
            : "Помилка написання ПІБ / ID; ООС не змінюємо",
      );
      setImpact(map, "excluded", "skip", "У Виключені не додаємо");
      setImpact(
        map,
        "timesheet",
        "skip",
        archiveReturnVsSh
          ? "Табель не фарбуємо: не відомо, чи archive рано закритий, чи sh застарів"
          : archiveMissing
          ? "Немає рядка archive — табельні коди відсутності не підставляємо"
          : rankIssue
            ? "Звання в Табелі не змінюємо — лише інформаційна перевірка"
            : "Помилка даних — Табель не змінюємо",
      );
      if (archiveReturnVsSh) {
        setImpact(
          map,
          "absents",
          "skip",
          `Перевірте: archive повернення ${op.payload.returnDate || "?"} vs sh «${op.payload.statusRaw || "СЗЧ"}»`,
        );
      }
    } else if (op.kind === "other_manual") {
      if (op.payload.chainWaiting === "1") {
        setImpact(
          map,
          "shpo",
          "skip",
          `Спочатку крок 1: зміна посади ${op.payload.awaitFromIndex || "—"} → ${op.payload.awaitToIndex || "—"}${op.payload.awaitOrderNumber ? ` (наказ №${op.payload.awaitOrderNumber})` : ""}`,
        );
        setImpact(
          map,
          "excluded",
          "skip",
          "Рядок закриття старої посади створює крок зміни посади",
        );
        setImpact(
          map,
          "timesheet",
          "skip",
          "Розпорядження в Табелі проводимо після зміни посади",
        );
      } else if (op.payload.type === "TRANSFER_CANCELLED") {
        const needsShReview =
          op.payload.reviewReason === "CANCEL_TRANSFER_BUT_NOT_IN_CURRENT_SH";
        const staffIndex = op.payload.previousIndex || op.positionIndex || "—";
        setImpact(
          map,
          "excluded",
          needsShReview ? "skip" : op.payload.excludedExcelRow ? "clear" : "skip",
          needsShReview
            ? `NEEDS_REVIEW / CANCEL_TRANSFER_BUT_NOT_IN_CURRENT_SH: рядок ${op.payload.excludedExcelRow ? `R${op.payload.excludedExcelRow}` : "№" + (op.payload.cancelledTransferOrder || "?")} лишаємо; чинним остаточним виключенням не вважаємо`
            : op.payload.excludedExcelRow
              ? `REMOVE_CANCELLED_EXCLUSION: прибрати запис №${op.payload.cancelledTransferOrder || "?"} від ${op.payload.cancelledTransferDate || "?"} (скасовано №${op.payload.transferCancelOrder || "?"})`
              : "REMOVE_CANCELLED_EXCLUSION: нового рядка не створюємо",
          op.payload.excludedExcelRow
            ? `R${op.payload.excludedExcelRow}`
            : undefined,
        );
        setImpact(
          map,
          "shpo",
          needsShReview ? "skip" : op.payload.restoreShpo === "1" ? "edit" : "skip",
          needsShReview
            ? `Не відновлюємо ${staffIndex} автоматично — в актуальній sh людини немає`
            : op.payload.restoreShpo === "1"
              ? `Повернути на штатну посаду ${staffIndex}`
              : "ШПО вже містить особу на штатному індексі",
        );
        setImpact(
          map,
          "oos",
          needsShReview ? "skip" : op.payload.restoreOos === "1" ? "edit" : "skip",
          needsShReview
            ? "Активну картку ООС не відновлюємо, поки sh не підтвердить перебування в 1 ПБ"
            : op.payload.restoreOos === "1"
              ? "Відновити активну картку з «Виключені»"
              : "Активний ООС уже є — не чіпаємо",
        );
        const timesheetWrite =
          op.payload.restoreTimesheet === "1" ||
          Boolean(op.payload.historyTimesheetExcelRow) ||
          Boolean(op.payload.duplicateTimesheetExcelRow);
        setImpact(
          map,
          "timesheet",
          needsShReview ? "skip" : timesheetWrite ? "edit" : "skip",
          needsShReview
            ? "Історичне вибуття лишаємо; новий активний рядок не створюємо автоматично"
            : op.payload.restoreTimesheet === "1"
              ? (() => {
                  const archiveSummary = formatArchiveSpanSummary(
                    op.payload.timesheetAbsenceSpans || "",
                  );
                  const archivePart = archiveSummary
                    ? `; archive: ${archiveSummary}`
                    : "";
                  const dupPart = op.payload.duplicateTimesheetExcelRow
                    ? `; очистити лише дубль R${op.payload.duplicateTimesheetExcelRow}`
                    : "";
                  return op.payload.timesheetCancelKeepGap === "1"
                    ? `R${op.payload.timesheetExcelRow || "?"}: «+» до ПЕРЕВ, «−» між ПЕРЕВ і скасуванням; примітка в ШПО R / історії${archivePart}${dupPart}`
                    : `R${op.payload.timesheetExcelRow || "?"}: «+» з ${op.payload.timesheetActiveFrom || "дати постановки"}; скасування №${op.payload.transferCancelOrder || "?"} — у ШПО R / історії, не в денній клітинці${archivePart}${dupPart}`;
                })()
              : op.payload.historyTimesheetExcelRow
                ? `Очистити історичний рядок R${op.payload.historyTimesheetExcelRow} з «вибув»`
                : op.payload.duplicateTimesheetExcelRow
                  ? `Очистити дубль R${op.payload.duplicateTimesheetExcelRow}`
                  : "Табель уже активний на штатній позиції — не чіпаємо",
          timesheetWrite && op.payload.timesheetExcelRow
            ? `R${op.payload.timesheetExcelRow}`
            : undefined,
        );
      } else if (op.payload.type === "РОЗПОРЯДЖ") {
        setImpact(
          map,
          "shpo",
          "edit",
          op.payload.oldPositionFreed === "true"
            ? "Стара штатна посада вже звільнена; перевірити блок розпорядження"
            : `Звільнити стару штатну посаду ${op.payload.previousIndex || op.positionIndex}`,
        );
        setImpact(
          map,
          "oos",
          "edit",
          op.payload.remainsInOos === "true"
            ? "Активний запис збережено"
            : "Відновити активний запис; особу не виключати",
        );
        if (op.payload.hasSzchContext === "true") {
          setImpact(
            map,
            "absents",
            "edit",
            op.payload.szchRemains === "true"
              ? "Чинний запис СЗЧ збережено"
              : "Перевірити або відновити чинний запис СЗЧ",
          );
        }
        setImpact(
          map,
          "timesheet",
          "edit",
          op.payload.dispositionInTimesheet === "true"
            ? "Історія вибуття у розпорядження присутня"
            : "Перевірити історію вибуття у розпорядження",
        );
      } else {
        setImpact(
          map,
          "excluded",
          "append",
          `Ручне: ${op.after || op.why || "—"}`,
        );
      }
    }
    if (
      op.payload.duplicateTimesheetExcelRow &&
      op.payload.type !== "TRANSFER_CANCELLED"
    ) {
      setImpact(
        map,
        "timesheet",
        "clear",
        `Очистити дубль після скасованого переведення (R${op.payload.duplicateTimesheetExcelRow})`,
        `R${op.payload.duplicateTimesheetExcelRow}`,
      );
    }
  });

  return CANONICAL_SHEETS.map(
    (sheet) =>
      map.get(sheet.key) || {
        sheetKey: sheet.key,
        sheetLabel: sheet.label,
        effect: "untouched" as const,
        effectLabel: effectLabelOf("untouched"),
        detail: "Не зачіпається",
      },
  );
};

const describeWillDo = (ops: EjoosSyncOp[]): string[] => {
  const cancelReview = ops.find(
    (op) =>
      op.payload.reviewReason === "CANCEL_TRANSFER_BUT_NOT_IN_CURRENT_SH",
  );
  if (cancelReview) {
    const dest =
      cancelReview.payload.cancelledDestination ||
      cancelReview.payload.destination ||
      "?";
    return [
      `РУХ: ПЕРЕВ №${cancelReview.payload.cancelledTransferOrder || "?"} від ${cancelReview.payload.cancelledTransferDate || "?"} → ${dest} скасовано №${cancelReview.payload.transferCancelOrder || "?"} від ${cancelReview.payload.transferCancelDate || "?"}`,
      "Актуальна sh: немає — це не AUTO_RESTORE і не NO_ACTION",
      `1. ШПО / 2. ООС: позицію ${cancelReview.payload.previousIndex || cancelReview.positionIndex || "—"} автоматично не відновлюємо`,
      `3. Виключені: ${cancelReview.payload.excludedExcelRow ? `R${cancelReview.payload.excludedExcelRow}` : "запис №" + (cancelReview.payload.cancelledTransferOrder || "?")} лишаємо; другий рядок не створюємо`,
      "6. Табель: історичне вибуття лишаємо; новий активний рядок з дати скасування не створюємо",
      "4 / 5. Тимч. прибулі / відсутні — не чіпати",
    ];
  }
  const disposition = ops.find(
    (op) => op.kind === "move_to_disposition",
  );
  if (disposition) {
    const payload = disposition.payload;
    return [
      `1. ШПО: звільнити посаду ${payload.previousIndex || disposition.positionIndex}`,
      payload.skipShpoDisposition === "1"
        ? "1. ШПО: у блоці розпорядження запис уже є"
        : "1. ШПО: додати особу до блоку розпорядження",
      payload.remainsInOos === "true"
        ? "2. ООС: залишити активний запис"
        : "2. ООС: без змін (активного запису немає)",
      payload.needsAbsenceRecord === "1"
        ? `5. Тимчасово відсутні: додати ${payload.absenceType || "відсутність"}${payload.absenceDate ? ` з ${payload.absenceDate}` : ""}`
        : "5. Тимчасово відсутні: чинний запис уже є",
      payload.keepOpenSzchTimesheet === "1"
        ? `6. Табель: один рядок 01–зріз ${payload.absenceCode || "СЗЧ"}; не писати «вибув у розпорядження»`
        : payload.timesheetFound === "true"
          ? "6. Табель: закрити штатний рядок і записати розпорядження"
          : "6. Табель: без змін (штатного рядка немає)",
      "3. Виключені: без змін",
    ];
  }
  const rankChange = ops.find((op) => op.kind === "rank_change");
  const exclude = ops.find((op) => op.kind === "exclude_transfer");
  if (exclude) {
    const nextRank = rankChange?.payload.nextRank || rankChange?.rank;
    const chainLine = exclude.payload.priorPlacementDate
      ? `РУХ: ${exclude.payload.priorPlacementType || "ПОСАДА"} ${exclude.payload.priorPlacementFromIndex || "?"} → ${exclude.payload.priorPlacementToIndex || "?"} · №${exclude.payload.priorPlacementOrder || "?"} від ${exclude.payload.priorPlacementDate} (лишився б у ООС) → ${exclude.payload.type || "ПЕРЕВ"} №${exclude.payload.orderNumber || "?"} від ${exclude.payload.excludeDate || "?"} — вибув з 1ПБ`
      : "";
    return [
      ...(chainLine ? [chainLine] : []),
      ...(rankChange
        ? [
            `1. ШПО / ООС / Табель: спочатку звання ${rankChange.payload.previousRank || rankChange.before} → ${nextRank} · №${rankChange.payload.orderNumber || "?"} від ${rankChange.payload.orderDate || "?"}`,
            `3. Виключені: перенести картку вже як «${nextRank}» → ${exclude.payload.destination || exclude.payload.documentsDest || "(куди?)"} · №${exclude.payload.orderNumber || "?"} від ${exclude.payload.excludeDate || "?"}`,
          ]
        : excludeWritePlan(exclude.payload).transitSameMonth
          ? [
              `РУХ: ${exclude.payload.appointmentOrderDate || "постановка"} на ${exclude.payload.fromPositionIndex || exclude.positionIndex || "штаб"} → вибуття №${exclude.payload.orderNumber || "?"} від ${exclude.payload.excludeDate || "?"}`,
              `3. Виключені: новий рядок · інд. ${exclude.payload.fromPositionIndex || exclude.positionIndex || "—"} → ${exclude.payload.destination || exclude.payload.documentsDest || "(куди?)"} · №${exclude.payload.orderNumber || "?"} від ${exclude.payload.excludeDate || "?"}`,
            ]
          : [
              chainLine
                ? `3. Виключені: інд. ${exclude.payload.fromPositionIndex || exclude.positionIndex || "—"} → ${exclude.payload.destination || exclude.payload.documentsDest || "(куди?)"} · №${exclude.payload.orderNumber || "?"} від ${exclude.payload.excludeDate || "?"}`
                : "1. Рух: ПЕРЕВ/РОЗПОРЯДЖ → куди",
              chainLine ? "" : "3. Виключені: дата / підстава / куди",
            ].filter(Boolean)),
      excludeWritePlan(exclude.payload).createTimesheetHistory
        ? `6. Табель: новий історичний рядок ${exclude.payload.timesheetActiveFrom || "з постановки"}–${exclude.payload.excludeDate || "вибуття"} (чужий штатний рядок не чіпаємо)`
        : excludeWritePlan(exclude.payload).replaceInPlace
          ? `6. Табель: закрити активний епізод вибуттям ${exclude.payload.excludeDate || "?"} (+ … −)`
          : `6. Табель: перенести рядок вниз у межах роти (вибуття ${exclude.payload.excludeDate || "?"}); на штаті лишаються індекс, ВОС, тарифний план`,
      exclude.payload.shpoExcelRow
        ? exclude.payload.occupiedPositionIndex &&
          exclude.payload.occupiedPositionIndex !==
            (exclude.payload.fromPositionIndex || exclude.positionIndex)
          ? `1. ШПО: прибрати з індексу ${exclude.payload.occupiedPositionIndex}; на ${exclude.payload.fromPositionIndex || exclude.positionIndex} не ставимо`
          : "1. ШПО: прибрати звання/ПІБ/ID, посаду лишити"
        : "1. ШПО: не чіпаємо — особи там немає, індекс зайнятий іншим",
      exclude.payload.oosExcelRow
        ? "2. ООС: видалити рядок після перенесення"
        : "2. ООС: не чіпаємо — активного запису немає",
      exclude.payload.arrivalExcelRow
        ? `4. Тимч. прибулі: зберегти прибуття і закрити ${exclude.payload.arrivalDepartDate || "?"} наказом №${exclude.payload.arrivalDepartOrderNumber || "?"} від ${exclude.payload.arrivalDepartOrderDate || "?"}`
        : "4 / 5. Тимч. прибулі / відсутні — не чіпати",
      exclude.payload.arrivalExcelRow ? "5. Тимч. відсутні — не чіпати" : "",
    ].filter(Boolean);
  }
  if (rankChange) {
    return [
      `1. ШПО: ${rankChange.payload.previousRank || rankChange.before} → ${rankChange.payload.nextRank || rankChange.rank}`,
      `2. ООС: звання ${rankChange.payload.nextRank || rankChange.rank}, останнє присвоєння №${rankChange.payload.orderNumber || "?"} від ${rankChange.payload.orderDate || "?"}`,
      `6. Табель: звання ${rankChange.payload.nextRank || rankChange.rank}`,
      "3 / 4 / 5. Виключені, тимч. прибулі / відсутні — без змін",
    ];
  }
  const placement = ops.find((op) => op.kind === "position_change");
  if (placement && positionCloseWritesExcluded(placement.payload)) {
    const step =
      placement.payload.chainStep && placement.payload.chainTotal
        ? `Крок ${placement.payload.chainStep} з ${placement.payload.chainTotal}: `
        : "";
    return [
      `${step}3. Виключені: закрити посаду ${placement.payload.previousIndex} · ${placement.payload.documentsDest || "нова посада"} · наказ №${placement.payload.orderNumber || "?"} від ${placement.payload.orderDate || "?"}`,
      `6. Табель: закрити рядок ${placement.payload.previousIndex} і поставити «+» на ${placement.payload.nextIndex} з дати наказу`,
      `1. ШПО: звільнити ${placement.payload.previousIndex}, поставити на ${placement.payload.nextIndex}`,
      "2. ООС: залишити активний запис, дописати історію посад",
      "5. Тимч. відсутні / 4. Тимч. прибулі: без змін",
    ];
  }
  if (placement) {
    const restoreAfterCancel = Boolean(
      placement.payload.cancelledTransferOrder ||
        placement.payload.transferCancelOrder,
    ) && placement.payload.timesheetPreserveHistory === "1";
    return [
      `1. ШПО: ${
        placement.payload.excludedSourceExcelRow ? "відновити" : "поставити"
      } на індекс ${placement.payload.nextIndex || placement.positionIndex}`,
      placement.payload.oosExcelRow
        ? "2. ООС: оновити активну картку та історію посад"
        : placement.payload.excludedSourceExcelRow
          ? `2. ООС: додати активну картку з «Виключені» R${placement.payload.excludedSourceExcelRow}`
          : "2. ООС: додати активну картку",
      restoreAfterCancel
        ? `6. Табель: лишити вибуття ${placement.payload.cancelledTransferDate || "04.08"} як історію; новий активний з ${placement.payload.timesheetActiveFrom || placement.payload.transferCancelDate || "дати скасування"}`
        : placement.payload.returningFromDisposition === "1"
          ? `6. Табель: поставити на штатний рядок ${placement.payload.nextIndex || placement.positionIndex} з ${placement.payload.timesheetActiveFrom || placement.payload.orderDate || "дати наказу"}; другий рядок не створюємо`
        : placement.payload.isTempArrivalPlacement === "1"
          ? `6. Табель: старий рядок з вибуттям лишити як історію; новий активний з ${placement.payload.timesheetActiveFrom || placement.payload.orderDate || "дати наказу"}`
        : placement.payload.cancelledTransferOrder
          ? `6. Табель: один рядок ${placement.payload.nextIndex || placement.positionIndex} з ${placement.payload.timesheetActiveFrom || placement.payload.orderDate || "дати постановки"}; прибрати дубль з «вибув»`
          : "6. Табель: поставити на штатний рядок від дати наказу",
      placement.payload.arrivalExcelRow
        ? `4. Тимч. прибулі: зберегти прибуття і закрити ${placement.payload.arrivalDepartDate || placement.payload.orderDate || "?"} наказом №${placement.payload.arrivalDepartOrderNumber || placement.payload.orderNumber || "?"} від ${placement.payload.arrivalDepartOrderDate || placement.payload.orderDate || "?"}`
        : "4. Тимч. прибулі: без змін",
      ops.some((op) => op.kind === "absent_upsert")
        ? "5. Тимч. відсутні: додати період з archive в цьому ж застосуванні"
        : "5. Тимч. відсутні: без змін",
      placement.payload.transferCancelOrder || placement.payload.cancelledTransferOrder
        ? `3. Виключені: прибрати запис №${placement.payload.cancelledTransferOrder || "?"} (скасовано №${placement.payload.transferCancelOrder || "?"})`
        : "3. Виключені: без змін",
      ops.some((op) => op.payload.mismatchKind === "ARCHIVE_REFERENCE_MISSING")
        ? "archive: немає запису при «ВІДСУТНІЙ в АРХІВІ» — перевірити окремо, коди відсутності не вигадуємо"
        : "",
    ].filter(Boolean);
  }
  const transferCancelled = ops.find(
    (op) =>
      op.payload.type === "TRANSFER_CANCELLED" &&
      op.payload.reviewReason !== "CANCEL_TRANSFER_BUT_NOT_IN_CURRENT_SH",
  );
  if (transferCancelled) {
    const staffIndex =
      transferCancelled.payload.previousIndex ||
      transferCancelled.positionIndex ||
      "—";
    return [
      `РУХ: ПЕРЕВ №${transferCancelled.payload.cancelledTransferOrder || "?"} скасовано №${transferCancelled.payload.transferCancelOrder || "?"} — rollback, не новий рух`,
      transferCancelled.payload.restoreShpo === "1"
        ? `1. ШПО: повернути на ${staffIndex}`
        : "1. ШПО: вже на місці",
      transferCancelled.payload.restoreOos === "1"
        ? "2. ООС: відновити з «Виключені»"
        : "2. ООС: активна картка вже є",
      transferCancelled.payload.excludedExcelRow
        ? `3. Виключені: прибрати запис №${transferCancelled.payload.cancelledTransferOrder || "?"} (скасовано №${transferCancelled.payload.transferCancelOrder || "?"})`
        : "3. Виключені: рядка немає — новий не створюємо",
      transferCancelled.payload.restoreTimesheet === "1"
        ? (() => {
            const archiveSummary = formatArchiveSpanSummary(
              transferCancelled.payload.timesheetAbsenceSpans || "",
            );
            const archivePart = archiveSummary
              ? `; archive: ${archiveSummary}`
              : "";
            return transferCancelled.payload.timesheetCancelKeepGap === "1"
              ? `6. Табель: R${transferCancelled.payload.timesheetExcelRow || staffIndex} — «+»/«−»; скасування в ШПО R / історії${archivePart}`
              : `6. Табель: R${transferCancelled.payload.timesheetExcelRow || staffIndex} — «+» з ${transferCancelled.payload.timesheetActiveFrom || "дати постановки"}; день скасування теж «+»${archivePart}`;
          })()
        : transferCancelled.payload.historyTimesheetExcelRow ||
            transferCancelled.payload.duplicateTimesheetExcelRow
          ? `6. Табель: очистити ${
              transferCancelled.payload.historyTimesheetExcelRow
                ? `історичний R${transferCancelled.payload.historyTimesheetExcelRow}`
                : `дубль R${transferCancelled.payload.duplicateTimesheetExcelRow}`
            }`
          : `6. Табель: уже активний на ${staffIndex}`,
      transferCancelled.payload.historyTimesheetExcelRow
        ? `6. Табель: очистити історичний рядок R${transferCancelled.payload.historyTimesheetExcelRow} з «вибув»`
        : transferCancelled.payload.duplicateTimesheetExcelRow
          ? `6. Табель: очистити лише дубль R${transferCancelled.payload.duplicateTimesheetExcelRow} (основний R${transferCancelled.payload.timesheetExcelRow || staffIndex} не чіпаємо)`
          : null,
      ops.some((op) => op.kind === "absent_upsert")
        ? "5. Тимч. відсутні: оновити періоди з archive"
        : "5. Тимч. відсутні: без змін",
    ].filter((line): line is string => Boolean(line));
  }
  return toSheetActions(ops)
    .slice(0, 8)
    .map((action) => `${action.sheet}: ${action.before} → ${action.after}`);
};

const SOURCE_LABEL: Record<PersonSourceKind, string> = {
  ruh: "Рух",
  archive: "archive",
  sh: "sh",
  ejoos: "ЕЖООС",
};

const collapseTimesheetRuns = (
  days: Array<{ day: number; mark: string }>,
): PersonTimesheetPreview["runs"] => {
  const runs: PersonTimesheetPreview["runs"] = [];
  for (const cell of days) {
    const last = runs[runs.length - 1];
    if (last && last.mark === cell.mark && last.to === cell.day - 1) {
      last.to = cell.day;
    } else {
      runs.push({ from: cell.day, to: cell.day, mark: cell.mark });
    }
  }
  return runs;
};

export const buildTimesheetPreview = (
  ops: EjoosSyncOp[],
  timesheetDay: number,
): PersonTimesheetPreview | null => {
  const excludeOp = [...ops]
    .reverse()
    .find((op) => op.kind === "exclude_transfer");
  const spanOp = [...ops]
    .reverse()
    .find((op) => op.payload.timesheetAbsenceSpans?.trim());
  const spans = parseTimesheetAbsenceSpans(
    spanOp?.payload.timesheetAbsenceSpans ||
      excludeOp?.payload.timesheetAbsenceSpans ||
      "",
  );
  const activeFrom =
    dayFromOrderLabel(
      excludeOp?.payload.timesheetActiveFrom ||
        spanOp?.payload.timesheetActiveFrom ||
        ops.find((op) => op.payload.timesheetActiveFrom)?.payload
          .timesheetActiveFrom ||
        "",
    ) || 1;
  const departDay = dayFromOrderLabel(
    excludeOp?.payload.excludeDate || excludeOp?.payload.orderDate || "",
  );
  const departMark = excludeOp
    ? formatTimesheetTransferMark(excludeOp.payload)
    : "";
  const dayOps = ops.filter(
    (op) =>
      op.kind === "timesheet_day" &&
      op.payload.timesheetCode?.trim() &&
      Number(op.payload.day || 0) > 0,
  );
  const rankOnly =
    !spans.length &&
    !dayOps.length &&
    activeFrom <= 1 &&
    !departDay &&
    ops.some((op) => op.kind === "rank_change");
  if (!spans.length && !dayOps.length && activeFrom <= 1 && !departDay && !rankOnly) {
    return null;
  }

  const spanEnd = spans.reduce((max, span) => Math.max(max, span.toDay), 0);
  const lastDay = Math.min(
    31,
    Math.max(timesheetDay, spanEnd, departDay, 1),
  );
  const days: PersonTimesheetPreview["days"] = [];
  for (let day = 1; day <= lastDay; day += 1) {
    if (departDay) {
      const mark = timesheetTransferMarkForDay({
        day,
        departDay,
        lastDay,
        activeFromDay: activeFrom,
        absenceSpans: spans,
        departMark: "ПЕРЕВ",
      });
      if (mark) {
        days.push({
          day,
          mark,
          title: mark === "ПЕРЕВ" ? departMark || "ПЕРЕВЕДЕННЯ" : undefined,
        });
      }
      continue;
    }

    const fromArchive = timesheetMarkFromArchive(day, {
      activeFromDay: activeFrom,
      lastDay,
      spans,
      fillBeforeActive: activeFrom > 1,
    });
    const fromDayOp = dayOps.find((op) => Number(op.payload.day) === day);
    const mark =
      fromArchive ||
      fromDayOp?.payload.timesheetCode ||
      (day <= timesheetDay ? "+" : "");
    if (!mark) continue;
    days.push({ day, mark });
  }
  if (!days.length) return null;
  return {
    lastDay,
    days,
    runs: collapseTimesheetRuns(days),
    note: rankOnly
      ? "Дні без змін — оновлюється лише звання."
      : ops.some(
            (op) =>
              op.kind === "position_change" &&
              (op.payload.returningToStaffIndex === "1" ||
                Boolean(op.payload.historyTimesheetExcelRow) ||
                Boolean(op.payload.previousTimesheetExcelRow)),
          )
        ? "Нижче — новий активний рядок. Історичний епізод з «вибув» лишається окремим рядком."
        : undefined,
    departDay: departDay || undefined,
    departPhrase: departDay && departMark ? departMark : undefined,
  };
};

export const buildSourceInfluences = (
  ops: EjoosSyncOp[],
): PersonSourceInfluence[] => {
  const items: PersonSourceInfluence[] = [];
  const push = (item: PersonSourceInfluence) => {
    const key = `${item.source}|${item.event}`;
    if (items.some((existing) => `${existing.source}|${existing.event}` === key)) {
      return;
    }
    items.push(item);
  };

  for (const op of ops) {
    const ref = op.sourceRef || "";
    if (op.kind === "rank_change") {
      push({
        source: "ruh",
        sourceLabel: SOURCE_LABEL.ruh,
        ref,
        event: `ЗВАННЯ: ${op.payload.previousRank || op.before || "—"} → ${op.payload.nextRank || op.rank || "—"} · наказ №${op.payload.orderNumber || "?"} від ${op.payload.orderDate || "?"}`,
        effect: "ШПО / ООС / Табель: лише звання; дні присутності не змінює",
      });
      continue;
    }
    if (op.kind === "absent_upsert") {
      const code =
        op.payload.timesheetCode ||
        formatArchiveSpanSummary(op.payload.timesheetAbsenceSpans || "") ||
        "коди archive";
      push({
        source: "archive",
        sourceLabel: SOURCE_LABEL.archive,
        ref,
        event: `${op.payload.absenceType || "відсутність"} з ${op.payload.departDate || "?"} → ${op.payload.place || "?"}${
          op.payload.orderNumber
            ? ` · наказ №${op.payload.orderNumber} від ${op.payload.orderDate || "?"}`
            : ""
        }`,
        effect: op.payload.timesheetActiveFrom
          ? `Тимч. відсутні: закрити період${op.payload.returnDate ? ` поверненням ${op.payload.returnDate}` : ""}. Табель нового епізоду коди відсутності не отримує`
          : `Тимч. відсутні + Табель: ${code}`,
      });
      continue;
    }
    if (op.kind === "absent_close") {
      push({
        source: /archive/i.test(ref) ? "archive" : "sh",
        sourceLabel: /archive/i.test(ref) ? SOURCE_LABEL.archive : SOURCE_LABEL.sh,
        ref,
        event: `Повернення ${op.payload.returnDate || op.payload.returnDay || "?"}`,
        effect: "Закрити «Тимч. відсутні» і оновити Табель з дати повернення",
      });
      continue;
    }
    if (op.kind === "exclude_transfer") {
      if (op.payload.priorPlacementDate) {
        push({
          source: "ruh",
          sourceLabel: SOURCE_LABEL.ruh,
          ref: op.payload.priorPlacementRow
            ? `Рух!R${op.payload.priorPlacementRow}`
            : ref,
          event: `${op.payload.priorPlacementType || "ПОСАДА"} №${op.payload.priorPlacementOrder || "?"} від ${op.payload.priorPlacementDate} · ${op.payload.priorPlacementFromIndex || "?"} → ${op.payload.priorPlacementToIndex || "?"} (1 ПБ → 1 ПБ)`,
          effect:
            "Внутрішня зміна посади: лишився б у ООС; штат не займаємо, бо далі зовнішнє вибуття",
        });
      }
      push({
        source: "ruh",
        sourceLabel: SOURCE_LABEL.ruh,
        ref,
        event: `ПЕРЕВ №${op.payload.orderNumber || "?"} від ${op.payload.orderDate || op.payload.excludeDate || "?"} → ${op.payload.destination || "?"}`,
        effect: "Виключені + Табель (вибув) + зняти з ШПО/ООС",
      });
      continue;
    }
    if (op.kind === "position_change") {
      push({
        source: "ruh",
        sourceLabel: SOURCE_LABEL.ruh,
        ref,
        event: op.payload.isTempArrivalPlacement === "1"
          ? `ПОСАДА №${op.payload.orderNumber || "?"} від ${op.payload.orderDate || "?"} · ${op.payload.arrivedFrom || "звідки?"} → штат ${op.payload.nextIndex || op.positionIndex || "?"}`
          : `ПОСАДА / внутрішній ПЕРЕВ → ${op.payload.nextIndex || op.positionIndex || "?"}`,
        effect: op.payload.isTempArrivalPlacement === "1"
          ? "Закрити тимч. прибулих, додати ШПО/ООС, Табель на штатний індекс від дати наказу"
          : "ШПО / ООС / Табель: зміна штатної позиції",
      });
      continue;
    }
    if (op.payload.mismatchKind === "ARCHIVE_RETURN_SH_STILL_ABSENT") {
      push({
        source: "archive",
        sourceLabel: SOURCE_LABEL.archive,
        ref,
        event: `archive повернення ${op.payload.returnDate || "?"} vs sh «${op.payload.statusRaw || "СЗЧ"}»`,
        effect: "NEEDS_REVIEW: не закриваємо період і не фарбуємо Табель",
      });
      continue;
    }
    if (op.payload.mismatchKind === "ARCHIVE_REFERENCE_MISSING") {
      push({
        source: "sh",
        sourceLabel: SOURCE_LABEL.sh,
        ref,
        event: `СТАТУС «${op.payload.statusRaw || "ВІДСУТНІЙ в АРХІВІ"}» без рядка archive`,
        effect: "NEEDS_REVIEW: не вигадуємо ЛІК / ВІД / СЗЧ / БЕЗВІСТИ",
      });
      continue;
    }
    if (op.payload.type === "TRANSFER_CANCELLED") {
      push({
        source: "ruh",
        sourceLabel: SOURCE_LABEL.ruh,
        ref,
        event: `Скасовано ПЕРЕВ №${op.payload.cancelledTransferOrder || "?"} наказом №${op.payload.transferCancelOrder || "?"} від ${op.payload.transferCancelDate || "?"}`,
        effect: "Rollback: Виключені / ШПО / ООС / один рядок Табеля",
      });
      continue;
    }
    if (op.kind === "timesheet_day" || op.kind === "shpo_occupant") {
      const source: PersonSourceKind = /archive/i.test(ref)
        ? "archive"
        : /^рух/i.test(ref)
          ? "ruh"
          : /sh!/i.test(ref)
            ? "sh"
            : "ejoos";
      push({
        source,
        sourceLabel: SOURCE_LABEL[source],
        ref,
        event: op.after || op.why || op.kind,
        effect: op.why || `${op.before || "—"} → ${op.after || "—"}`,
      });
    }
  }
  return items;
};

export const patchPersonOpPayload = (
  session: EjoosDiffSession,
  personChangeId: string,
  opId: string,
  payloadPatch: Record<string, string>,
): EjoosDiffSession => {
  const people = session.people.map((person) => {
    if (person.id !== personChangeId) return person;
    const ops = person.ops.map((op) => {
      if (op.id !== opId) return op;
      const payload = { ...op.payload, ...payloadPatch };
      if (op.kind === "timesheet_day") {
        if (payload.clearStalePerson === "1") {
          return {
            ...op,
            payload,
            class: payload.excelRow ? ("ready" as const) : ("needs_input" as const),
            confidence: payload.excelRow ? ("high" as const) : ("manual" as const),
            checkedDefault: Boolean(payload.excelRow),
          };
        }
        const timesheetCode = payload.timesheetCode?.trim() || "";
        const ready = Boolean(timesheetCode && payload.excelRow);
        return {
          ...op,
          payload,
          after: timesheetCode || "(оберіть код)",
          class: ready ? ("ready" as const) : ("needs_input" as const),
          confidence: ready ? ("high" as const) : ("manual" as const),
          checkedDefault: ready,
        };
      }
      if (op.kind !== "exclude_transfer") {
        return { ...op, payload };
      }
      const destination = payload.destination || "";
      const excludeDate = payload.excludeDate || "";
      const ready =
        Boolean(destination) &&
        Boolean(excludeDate) &&
        Boolean(payload.orderNumber) &&
        Boolean(payload.orderDate) &&
        Boolean(payload.fromName || op.fullName) &&
        op.class !== "conflict";
      return {
        ...op,
        payload,
        after: `виключити: ${payload.type || "ПЕРЕВ"} → ${destination || "(куди?)"} · дата ${excludeDate || "?"}`,
        class: ready ? ("ready" as const) : ("needs_input" as const),
        checkedDefault: ready,
      };
    });
    return {
      ...person,
      ops,
      sheetActions: toSheetActions(ops),
      sheetImpacts: buildSheetImpacts(ops),
      ejoosWillDo: describeWillDo(ops),
      sourceInfluences: buildSourceInfluences(ops),
      timesheetPreview: buildTimesheetPreview(ops, session.plan.timesheetDay),
      severity: severityForOps(ops),
      summaryAfter:
        ops.find((item) => item.kind === "exclude_transfer")?.after ||
        ops.find((item) => item.kind === "rank_change")?.after ||
        ops.find((item) => item.kind === "timesheet_day")?.after ||
        person.summaryAfter,
    };
  });
  const updatedOp = people
    .find((person) => person.id === personChangeId)
    ?.ops.find((op) => op.id === opId);
  return {
    ...session,
    people,
    plan: {
      ...session.plan,
      ops: session.plan.ops.map((op) =>
        op.id === opId && updatedOp ? updatedOp : op,
      ),
    },
  };
};

export const personChangesFromOps = (
  ops: EjoosSyncOp[],
  timesheetDay: number,
  options?: { decision?: PersonChangeDecision },
): PersonChange[] => {
  const groups = new Map<string, EjoosSyncOp[]>();
  ops.forEach((op) => {
    const key = personKey(op);
    const list = groups.get(key) ?? [];
    list.push(op);
    groups.set(key, list);
  });
  for (const [nameKey, nameOps] of [...groups.entries()]) {
    if (!nameKey.startsWith("name:")) continue;
    const name = nameOps[0]?.fullName;
    if (!name) continue;
    const idKey = [...groups.keys()].find((key) => {
      if (!key.startsWith("id:")) return false;
      const idOps = groups.get(key);
      if (!idOps?.length) return false;
      const namesInId = new Set(
        idOps.map((op) => normKey(op.fullName)).filter(Boolean),
      );
      if (namesInId.size > 1) return false;
      return idOps.some(
        (op) => op.fullName && normKey(op.fullName) === normKey(name),
      );
    });
    if (!idKey) continue;
    groups.set(idKey, [...(groups.get(idKey) ?? []), ...nameOps]);
    groups.delete(nameKey);
  }

  const people: PersonChange[] = [...groups.entries()].map(([key, grouped]) => {
    const head = grouped[0];
    const category = categoryForOps(grouped);
    const severity = severityForOps(grouped);
    const ordered = [...grouped].sort((left, right) => {
      const order = (kind: string) =>
        kind === "rank_change" ? 0 : kind === "exclude_transfer" ? 1 : 2;
      return order(left.kind) - order(right.kind);
    });
    const statusOp = ordered.find((op) => op.kind === "timesheet_day");
    const rankOp = ordered.find((op) => op.kind === "rank_change");
    const excludeOp = ordered.find((op) => op.kind === "exclude_transfer");
    const summaryBefore =
      rankOp?.before ||
      statusOp?.before ||
      ordered.map((op) => op.before).find((v) => v && v !== "—") ||
      "—";
    const summaryAfter =
      excludeOp?.after ||
      rankOp?.after ||
      statusOp?.after ||
      ordered.map((op) => op.after).find((v) => v && v !== "—") ||
      "—";

    return {
      id: key,
      personId: head.personId,
      fullName: head.fullName || "Без ПІБ",
      rank:
        rankOp?.rank ||
        excludeOp?.rank ||
        excludeOp?.payload.fromRank ||
        head.rank,
      positionIndex: head.positionIndex,
      category:
        category === "error" && severity === "conflict" ? "error" : category,
      severity,
      summaryBefore,
      summaryAfter,
      ejoosWillDo: describeWillDo(ordered),
      sheetActions: toSheetActions(ordered),
      sheetImpacts: buildSheetImpacts(ordered),
      sourceInfluences: buildSourceInfluences(ordered),
      timesheetPreview: buildTimesheetPreview(ordered, timesheetDay),
      ops: ordered,
      decision: options?.decision ?? "pending",
    };
  });

  people.sort((a, b) => {
    const rank = { conflict: 0, needs_input: 1, ready: 2 };
    const d = rank[a.severity] - rank[b.severity];
    if (d !== 0) return d;
    return a.fullName.localeCompare(b.fullName, "uk");
  });
  return people;
};

export const groupOpsIntoPersonChanges = (
  plan: EjoosSyncPlan,
  pb?: ExcelWorkbookSnapshot | null,
  statusRules = readOperatorSettings().statusRules,
): EjoosDiffSession => {
  const people = personChangesFromOps(plan.ops, plan.timesheetDay);

  const shPeople = pb ? parsePbShPeople(pb) : [];
  let onDuty = 0;
  shPeople.forEach((person) => {
    const mapped = mapPbStatusToEjoosWithRules(person.status, statusRules);
    if (mapped.timesheetCode === "+") onDuty += 1;
  });

  const newcomers = people.filter((p) => p.category === "arrival").length;
  const errors = people.filter(
    (p) => p.severity === "conflict" || p.category === "error",
  ).length;
  const autoReady = people.filter((p) => p.severity === "ready").length;
  const needsReview = people.filter((p) => p.severity !== "ready").length;

  return {
    plan,
    people,
    counters: {
      oosLike: shPeople.length || people.length,
      onDuty,
      changes: people.length,
      newcomers,
      errors,
      autoReady,
      needsReview,
    },
    pbFileName: plan.pbName,
    analyzedAt: new Date().toISOString(),
  };
};

export const setPersonDecision = (
  session: EjoosDiffSession,
  personId: string,
  decision: PersonChangeDecision,
): EjoosDiffSession => ({
  ...session,
  people: session.people.map((person) =>
    person.id === personId ? { ...person, decision } : person,
  ),
});

export const dismissPersonFromSession = (
  session: EjoosDiffSession,
  personId: string,
): EjoosDiffSession => ({
  ...session,
  people: session.people.filter((person) => person.id !== personId),
  counters: {
    ...session.counters,
    changes: Math.max(0, session.counters.changes - 1),
  },
});

export const setPersonDecisions = (
  session: EjoosDiffSession,
  personIds: Iterable<string>,
  decision: PersonChangeDecision,
): EjoosDiffSession => {
  const ids = new Set(personIds);
  if (!ids.size) return session;
  return {
    ...session,
    people: session.people.map((person) =>
      ids.has(person.id) ? { ...person, decision } : person,
    ),
  };
};

export const collectedAcceptedOps = (session: EjoosDiffSession): EjoosSyncOp[] =>
  session.people
    .filter((person) => person.decision === "accepted")
    .flatMap((person) => person.ops)
    .filter((op) => op.class !== "conflict");

/** Скасування ПЕРЕВ пише книгу лише якщо є рядок/прапорець відновлення. */
export const transferCancelledHasWorkbookWrites = (op: EjoosSyncOp) => {
  if (op.payload.type !== "TRANSFER_CANCELLED") return false;
  if (op.payload.reviewReason === "CANCEL_TRANSFER_BUT_NOT_IN_CURRENT_SH") {
    return false;
  }
  return Boolean(
    op.payload.excludedExcelRow ||
      op.payload.restoreTimesheet === "1" ||
      op.payload.restoreShpo === "1" ||
      op.payload.restoreOos === "1" ||
      op.payload.historyTimesheetExcelRow ||
      op.payload.duplicateTimesheetExcelRow,
  );
};

/** ПІБ / ID / звання, ПРИБУВ, «зачекати крок» або перегляд без запису в книгу. */
export const isInformationalOp = (op: EjoosSyncOp) => {
  if (op.kind === "data_mismatch" || op.kind === "arrival") return true;
  if (op.kind !== "other_manual") return false;
  if (
    op.payload.type === "TRANSFER_SCOPE_UNCLEAR" ||
    op.payload.transferScope === "unclear"
  ) {
    return false;
  }
  return !transferCancelledHasWorkbookWrites(op);
};

const WORKBOOK_APPLY_KINDS = new Set<EjoosOpKind>([
  "timesheet_day",
  "shpo_occupant",
  "absent_upsert",
  "absent_close",
  "exclude_transfer",
  "move_to_disposition",
  "position_change",
  "rank_change",
]);

/** Є реальний handler у apply; arrival / generic other_manual — ні. */
export const isWorkbookApplyOp = (op: EjoosSyncOp) => {
  if (op.class === "conflict") return false;
  if (isInformationalOp(op)) return false;
  if (WORKBOOK_APPLY_KINDS.has(op.kind)) return true;
  return (
    op.kind === "other_manual" && transferCancelledHasWorkbookWrites(op)
  );
};

export const writableOps = (ops: EjoosSyncOp[]) =>
  ops.filter((op) => op.class !== "conflict" && !isInformationalOp(op));

export const personHasWorkbookApplyOps = (ops: EjoosSyncOp[]) =>
  ops.some(isWorkbookApplyOp);

export const personIsInformationalOnly = (ops: EjoosSyncOp[]) =>
  ops.length > 0 && writableOps(ops).length === 0;

export const personCanEnterApplyQueue = (person: PersonChange) =>
  person.severity === "ready" &&
  personHasWorkbookApplyOps(person.ops) &&
  !personOpsBlockApply(person.ops);

export const acceptAllReady = (session: EjoosDiffSession): EjoosDiffSession => ({
  ...session,
  people: session.people.map((person) =>
    personCanEnterApplyQueue(person)
      ? { ...person, decision: "accepted" as const }
      : person,
  ),
});

export const collectedWritableAcceptedOps = (
  session: EjoosDiffSession,
): EjoosSyncOp[] => writableOps(collectedAcceptedOps(session));

/** Після перебудови плану зберегти ручні «прийняти/відхилити» для інших людей. */
export const mergePersonDecisions = (
  rebuilt: EjoosDiffSession,
  previous: EjoosDiffSession,
): EjoosDiffSession => ({
  ...rebuilt,
  people: rebuilt.people.map((person) => {
    const prev = previous.people.find(
      (item) =>
        item.id === person.id ||
        (item.personId && item.personId === person.personId),
    );
    if (!prev || prev.decision === "pending") return person;
    return { ...person, decision: prev.decision };
  }),
});
