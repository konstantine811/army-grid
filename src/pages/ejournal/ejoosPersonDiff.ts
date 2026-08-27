import type { EjoosOpClass, EjoosOpKind, EjoosSyncOp, EjoosSyncPlan } from "./ejoosSyncPlan";
import { parsePbShPeople } from "./ejoosSyncPlan";
import { mapPbStatusToEjoosWithRules, readOperatorSettings } from "./ejoosStatusMap";
import type { ExcelWorkbookSnapshot } from "../../excelRoundTrip";

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
  if (op.personId.trim()) return `id:${op.personId.trim()}`;
  if (op.fullName.trim()) return `name:${normKey(op.fullName)}`;
  if (op.positionIndex.trim()) return `idx:${op.positionIndex.trim()}`;
  return `op:${op.id}`;
};

const categoryForOps = (ops: EjoosSyncOp[]): PersonChangeCategory => {
  const kinds = new Set(ops.map((op) => op.kind));
  if (kinds.has("arrival")) return "arrival";
  if (kinds.has("conflict") || ops.some((op) => op.class === "conflict")) {
    return "error";
  }
  const statusKinds: EjoosOpKind[] = [
    "timesheet_day",
    "absent_upsert",
    "absent_close",
  ];
  const positionKinds: EjoosOpKind[] = ["position_change", "shpo_occupant"];
  const hasStatus = statusKinds.some((k) => kinds.has(k));
  const hasPos = positionKinds.some((k) => kinds.has(k));
  const hasExclude = kinds.has("exclude_transfer") || kinds.has("other_manual");
  if (ops.every((op) => op.class === "needs_input" || op.class === "conflict")) {
    if (!hasStatus && !hasPos && hasExclude) return "error";
  }
  if (hasStatus && !hasPos && !hasExclude) return "status";
  if (hasPos && !hasStatus) return "position";
  if (hasStatus || hasPos || hasExclude) return "mixed";
  return "data";
};

const severityForOps = (ops: EjoosSyncOp[]): EjoosOpClass => {
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
  // skip/untouched не перебивають реальну дію
  if (
    prev &&
    prev.effect !== "untouched" &&
    prev.effect !== "skip" &&
    (effect === "untouched" || effect === "skip")
  ) {
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
  CANONICAL_SHEETS.forEach((sheet) => {
    map.set(sheet.key, {
      sheetKey: sheet.key,
      sheetLabel: sheet.label,
      effect: "untouched",
      effectLabel: effectLabelOf("untouched"),
      detail: "Не зачіпається цією зміною",
    });
  });

  const exclude = ops.find((op) => op.kind === "exclude_transfer");
  if (exclude) {
    const dest =
      exclude.payload.destination || exclude.payload.documentsDest || "(куди?)";
    const date =
      exclude.payload.excludeDate || exclude.payload.orderDate || "?";
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

    setImpact(
      map,
      "excluded",
      "append",
      `Новий рядок: ${who || "—"} · інд. ${idx} → ${dest} · дата ${date}`,
    );
    setImpact(
      map,
      "timesheet",
      "history",
      `Копія рядка з + до дня вибуття, далі −; на посаді ${idx} прибрати особу`,
      exclude.payload.timesheetExcelRow
        ? `R${exclude.payload.timesheetExcelRow}`
        : undefined,
    );
    setImpact(
      map,
      "shpo",
      "clear",
      `Прибрати звання/ПІБ/ID на індексі ${idx}, посаду лишити`,
      exclude.payload.shpoExcelRow
        ? `R${exclude.payload.shpoExcelRow}`
        : undefined,
    );
    setImpact(
      map,
      "oos",
      "clear",
      `Видалити рядок анкети${exclude.personId ? ` (ID ${exclude.personId})` : ""}`,
    );
    setImpact(map, "absents", "skip", "Тимч. відсутні не чіпаємо при ПЕРЕВ");
    setImpact(map, "arrivals", "skip", "Тимч. прибулі не чіпаємо при ПЕРЕВ");
    setImpact(
      map,
      "history",
      "append",
      "Запис у протокол застосування версії",
    );
  }

  ops.forEach((op) => {
    if (op.kind === "exclude_transfer") return;
    if (op.kind === "timesheet_day") {
      setImpact(
        map,
        "timesheet",
        "edit",
        `День: ${op.before || "—"} → ${op.after || op.payload.timesheetCode || "—"}`,
        op.payload.timesheetExcelRow
          ? `R${op.payload.timesheetExcelRow}`
          : undefined,
      );
    } else if (op.kind === "shpo_occupant") {
      setImpact(
        map,
        "shpo",
        "edit",
        `${op.before || "—"} → ${op.after || "—"}`,
        op.payload.shpoExcelRow ? `R${op.payload.shpoExcelRow}` : undefined,
      );
      if (op.payload.timesheetExcelRow) {
        setImpact(
          map,
          "timesheet",
          "edit",
          `Підтягнути особу в табель`,
          `R${op.payload.timesheetExcelRow}`,
        );
      }
    } else if (op.kind === "position_change") {
      setImpact(
        map,
        "shpo",
        "edit",
        `Індекс/посада: ${op.before || "—"} → ${op.after || "—"}`,
      );
      setImpact(map, "oos", "edit", "Оновити індекси посад в ООС (за потреби)");
    } else if (op.kind === "absent_upsert" || op.kind === "absent_close") {
      setImpact(
        map,
        "absents",
        op.kind === "absent_upsert" ? "append" : "edit",
        `${op.before || "—"} → ${op.after || "—"}`,
        op.payload.existingExcelRow
          ? `R${op.payload.existingExcelRow}`
          : undefined,
      );
    } else if (op.kind === "arrival") {
      setImpact(map, "oos", "append", `Прибуття: ${op.after || "—"}`);
      setImpact(map, "shpo", "edit", "Зайняти посаду після підтвердження");
    } else if (op.kind === "other_manual") {
      setImpact(
        map,
        "excluded",
        "append",
        `Ручне: ${op.after || op.why || "—"}`,
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
  const exclude = ops.find((op) => op.kind === "exclude_transfer");
  if (exclude) {
    return [
      "1. Рух: ПЕРЕВ/РОЗПОРЯДЖ → куди",
      "2. ШПО → дані у Виключені",
      "3. Виключені: дата / підстава / куди",
      "4. Табель: історія (+ … −) + очистка особи",
      "5. ШПО: прибрати звання/ПІБ/ID",
      "6. ООС: видалити рядок",
      "7. Тимч. відсутні/прибулі — не чіпати",
    ];
  }
  return toSheetActions(ops)
    .slice(0, 8)
    .map((action) => `${action.sheet}: ${action.before} → ${action.after}`);
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
      const destination = payload.destination || "";
      const excludeDate = payload.excludeDate || payload.orderDate || "";
      const ready =
        Boolean(destination) &&
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
      severity: severityForOps(ops),
      summaryAfter:
        ops.find((item) => item.kind === "exclude_transfer")?.after ||
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

export const groupOpsIntoPersonChanges = (
  plan: EjoosSyncPlan,
  pb?: ExcelWorkbookSnapshot | null,
): EjoosDiffSession => {
  const groups = new Map<string, EjoosSyncOp[]>();
  plan.ops.forEach((op) => {
    const key = personKey(op);
    const list = groups.get(key) ?? [];
    list.push(op);
    groups.set(key, list);
  });

  const people: PersonChange[] = [...groups.entries()].map(([key, ops]) => {
    const head = ops[0];
    const category = categoryForOps(ops);
    const severity = severityForOps(ops);
    const statusOp = ops.find((op) => op.kind === "timesheet_day");
    const summaryBefore =
      statusOp?.before ||
      ops.map((op) => op.before).find((v) => v && v !== "—") ||
      "—";
    const summaryAfter =
      statusOp?.after ||
      ops.map((op) => op.after).find((v) => v && v !== "—") ||
      "—";

    return {
      id: key,
      personId: head.personId,
      fullName: head.fullName || "Без ПІБ",
      rank: head.rank,
      positionIndex: head.positionIndex,
      category:
        category === "error" && severity === "conflict" ? "error" : category,
      severity,
      summaryBefore,
      summaryAfter,
      ejoosWillDo: describeWillDo(ops),
      sheetActions: toSheetActions(ops),
      sheetImpacts: buildSheetImpacts(ops),
      ops,
      decision: severity === "ready" ? "accepted" : "pending",
    };
  });

  people.sort((a, b) => {
    const rank = { conflict: 0, needs_input: 1, ready: 2 };
    const d = rank[a.severity] - rank[b.severity];
    if (d !== 0) return d;
    return a.fullName.localeCompare(b.fullName, "uk");
  });

  const shPeople = pb ? parsePbShPeople(pb) : [];
  const statusRules = readOperatorSettings().statusRules;
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

export const acceptAllReady = (session: EjoosDiffSession): EjoosDiffSession => ({
  ...session,
  people: session.people.map((person) =>
    person.severity === "ready"
      ? { ...person, decision: "accepted" }
      : person,
  ),
});

export const collectedAcceptedOps = (session: EjoosDiffSession): EjoosSyncOp[] =>
  session.people
    .filter((person) => person.decision === "accepted")
    .flatMap((person) => person.ops)
    .filter((op) => op.class !== "conflict");
