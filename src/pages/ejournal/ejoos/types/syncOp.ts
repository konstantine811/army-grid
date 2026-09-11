export type EjoosOpClass = "ready" | "needs_input" | "conflict";

export type EjoosOpKind =
  | "timesheet_day"
  | "shpo_occupant"
  | "absent_upsert"
  | "absent_close"
  | "exclude_transfer"
  | "move_to_disposition"
  | "data_mismatch"
  | "position_change"
  | "rank_change"
  | "contract_update"
  | "arrival"
  | "other_manual";

export type EjoosSyncOp = {
  id: string;
  kind: EjoosOpKind;
  class: EjoosOpClass;
  sheet: string;
  personId: string;
  fullName: string;
  positionIndex: string;
  rank: string;
  before: string;
  after: string;
  sourceRef: string;
  why: string;
  confidence: "high" | "review" | "manual";
  /** Extra payload for apply */
  payload: Record<string, string>;
  /** Stable source-event key persisted after a successful DB version write. */
  movementKey?: string;
  checkedDefault: boolean;
};

export type EjoosSyncPlan = {
  ejoosName: string;
  pbName: string;
  timesheetDay: number;
  timesheetDayLabel: string;
  /** Місяць Табеля ЕЖООС ≠ місяць 1ПБ — apply заборонено. */
  monthRolloverRequired?: boolean;
  /** Назва 1ПБ без дати — горизонт Табеля невідомий. */
  sourceDateUnknown?: boolean;
  /** Заголовок місяця «6. Табель» не знайдено — apply заборонено. */
  timesheetMonthHeaderUnknown?: boolean;
  ejoosTimesheetMonthLabel?: string;
  ops: EjoosSyncOp[];
  summary: {
    ready: number;
    needsInput: number;
    conflict: number;
  };
  limitsNote?: string;
};

export const EJOOS_OP_KINDS: readonly EjoosOpKind[] = [
  "timesheet_day",
  "shpo_occupant",
  "absent_upsert",
  "absent_close",
  "exclude_transfer",
  "move_to_disposition",
  "data_mismatch",
  "position_change",
  "rank_change",
  "contract_update",
  "arrival",
  "other_manual",
] as const;
