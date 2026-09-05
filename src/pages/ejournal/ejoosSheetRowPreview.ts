import type { EjoosSyncOp } from "./ejoosSyncPlan";
import {
  formatExcludedDestination,
  formatExcludedListBasis,
  formatTimesheetTransferMark,
} from "./ejoosExcludedColumns";
import { excludeWritePlan, positionCloseWritesExcluded } from "./ejoosExcludePolicy";
import {
  dayFromOrderLabel,
  parseTimesheetAbsenceSpans,
  timesheetTransferMarkForDay,
} from "./ejoosTimesheetText";

export type SheetRowPreviewCellKind = "value" | "empty" | "keep" | "depart";

export type SheetRowPreviewCell = {
  letter: string;
  header: string;
  value: string;
  kind: SheetRowPreviewCellKind;
};

export type SheetRowPreview = {
  sheetKey: string;
  sheetLabel: string;
  role: string;
  note?: string;
  cells: SheetRowPreviewCell[];
};

const excelColLetter = (column1: number) => {
  let n = column1;
  let out = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
};

const previewCell = (
  column: number,
  header: string,
  value: string,
  kind: SheetRowPreviewCellKind = value ? "value" : "empty",
): SheetRowPreviewCell => ({
  letter: excelColLetter(column),
  header,
  value:
    kind === "empty"
      ? "порожньо"
      : kind === "keep"
        ? value || "лишається"
        : value || "—",
  kind,
});

const dayHeader = (from: number, to: number) =>
  from === to
    ? String(from).padStart(2, "0")
    : `${String(from).padStart(2, "0")}–${String(to).padStart(2, "0")}`;

const timesheetDayCells = (
  op: EjoosSyncOp,
  timesheetDay: number,
  departPhrase: string,
): SheetRowPreviewCell[] => {
  const departDay = dayFromOrderLabel(
    op.payload.excludeDate || op.payload.orderDate || "",
  );
  if (departDay < 1) return [];
  const activeFromDay = dayFromOrderLabel(op.payload.timesheetActiveFrom);
  const absenceSpans = parseTimesheetAbsenceSpans(
    op.payload.timesheetAbsenceSpans || "",
  );
  const lastDay = Math.min(31, Math.max(departDay, timesheetDay, 1));
  const days: Array<{ day: number; mark: string }> = [];
  for (let day = 1; day <= lastDay; day += 1) {
    const mark = timesheetTransferMarkForDay({
      day,
      departDay,
      lastDay,
      activeFromDay,
      absenceSpans,
      departMark: departPhrase,
    });
    if (mark) days.push({ day, mark });
  }
  const cells: SheetRowPreviewCell[] = [];
  for (const cell of days) {
    const last = cells[cells.length - 1];
    const isDepart = cell.mark === departPhrase;
    if (
      last &&
      last.value === cell.mark &&
      !isDepart &&
      last.kind !== "depart"
    ) {
      const from = Number(last.header.split("–")[0]);
      last.header = dayHeader(from, cell.day);
      last.letter = `${excelColLetter(8 + from)}–${excelColLetter(8 + cell.day)}`;
      continue;
    }
    cells.push({
      letter: excelColLetter(8 + cell.day),
      header: dayHeader(cell.day, cell.day),
      value: cell.mark,
      kind: isDepart ? "depart" : cell.mark === "-" ? "empty" : "value",
    });
  }
  return cells;
};

const excludePreviews = (
  op: EjoosSyncOp,
  timesheetDay: number,
): SheetRowPreview[] => {
  const rank = op.payload.fromRank || op.rank || "";
  const name = op.payload.fromName || op.fullName || "";
  const personId = op.payload.fromPersonId || op.personId || "";
  const index =
    op.payload.fromPositionIndex ||
    op.payload.previousIndex ||
    op.positionIndex ||
    "";
  const dest = formatExcludedDestination(
    op.payload.documentsDest || op.payload.changeText,
  );
  const basis = formatExcludedListBasis(op.payload);
  const departPhrase = formatTimesheetTransferMark(op.payload);
  const rows: SheetRowPreview[] = [];

  if (op.payload.shpoExcelRow) {
    rows.push({
      sheetKey: "shpo",
      sheetLabel: "1. ШПО",
      role: `Очистити особу · R${op.payload.shpoExcelRow}`,
      note: "Індекс і назва посади лишаються вакантними. Звання / ПІБ / ID прибираємо.",
      cells: [
        previewCell(1, "індекс", index, "keep"),
        previewCell(6, "звання", "", "empty"),
        previewCell(7, "ПІБ", "", "empty"),
        previewCell(8, "ID", "", "empty"),
      ],
    });
  }

  if (op.payload.oosExcelRow) {
    rows.push({
      sheetKey: "oos",
      sheetLabel: "2. ООС",
      role: `Очистити анкету · R${op.payload.oosExcelRow}`,
      note: `Зараз: ${[rank, name, personId, index].filter(Boolean).join(" · ") || "—"}. Після застосування колонки A–AN цього рядка порожні.`,
      cells: [
        previewCell(1, "звання", "", "empty"),
        previewCell(2, "ПІБ", "", "empty"),
        previewCell(3, "ID", "", "empty"),
        previewCell(4, "індекс", "", "empty"),
      ],
    });
  }

  if (!op.payload.excludedExcelRow) {
    rows.push({
    sheetKey: "excluded",
    sheetLabel: "3. Виключені",
    role: "Новий рядок (після останнього)",
    note: op.payload.oosExcelRow
      ? `A–AA копіюються з ООС R${op.payload.oosExcelRow}; AB–AF дописуємо з РУХ.`
      : "Картки в ООС немає — звання/ПІБ/ID/індекс і поля вибуття пишемо з РУХ.",
    cells: [
      previewCell(1, "звання", rank),
      previewCell(2, "ПІБ", name),
      previewCell(3, "ID", personId),
      previewCell(4, "індекс", index),
      ...(op.payload.arrivedFrom
        ? [previewCell(6, "звідки", op.payload.arrivedFrom)]
        : []),
      previewCell(28, "дата вибуття", op.payload.excludeDate || ""),
      previewCell(29, "дата наказу", op.payload.orderDate || ""),
      previewCell(30, "№ наказу", op.payload.orderNumber || ""),
      previewCell(31, "куди вибув", dest),
      previewCell(32, "підстава", basis),
    ],
  });
  }

  const historyDays = timesheetDayCells(op, timesheetDay, departPhrase);
  const timesheetPlan = excludeWritePlan(op.payload);
  if (timesheetPlan.createTimesheetHistory || op.payload.timesheetExcelRow) {
    rows.push({
      sheetKey: "timesheet-history",
      sheetLabel: "6. Табель",
      role: timesheetPlan.replaceInPlace
        ? `Замінити наявний рядок · R${op.payload.timesheetExcelRow || "—"}`
        : timesheetPlan.createTimesheetHistory
        ? "Новий історичний рядок"
        : `Історична копія · з R${op.payload.timesheetExcelRow || "—"}`,
      note: "У день вибуття — повна фраза без назви посади, прийменник до/у за підрозділом.",
      cells: [
        previewCell(6, "звання", rank),
        previewCell(7, "ПІБ", name),
        previewCell(8, "ID", personId),
        ...historyDays,
      ],
    });
  }

  if (op.payload.timesheetExcelRow && !timesheetPlan.replaceInPlace) {
    rows.push({
      sheetKey: "timesheet-active",
      sheetLabel: "6. Табель",
      role: `Штатний рядок · R${op.payload.timesheetExcelRow}`,
      note: "Особу й дні прибираємо; індекс посади, ВОС і тарифний план лишаються.",
      cells: [
        previewCell(2, "індекс", index, "keep"),
        previewCell(6, "звання", "", "empty"),
        previewCell(7, "ПІБ", "", "empty"),
        previewCell(8, "ID", "", "empty"),
      ],
    });
  }

  return rows.filter((row) => row.cells.length || row.note);
};

const positionPreviews = (op: EjoosSyncOp): SheetRowPreview[] => {
  const index = op.payload.nextIndex || op.positionIndex || "";
  const prev = op.payload.previousIndex || "";
  const rows: SheetRowPreview[] = [];
  if (positionCloseWritesExcluded(op.payload)) {
    rows.push({
      sheetKey: "excluded",
      sheetLabel: "3. Виключені",
      role: "Новий рядок закриття старої посади",
      cells: [
        previewCell(1, "звання", op.rank),
        previewCell(2, "ПІБ", op.fullName),
        previewCell(3, "ID", op.personId),
        previewCell(4, "індекс", prev),
        previewCell(
          31,
          "куди вибув",
          formatExcludedDestination(
            op.payload.documentsDest || op.payload.changeText,
          ),
        ),
        previewCell(32, "підстава", formatExcludedListBasis(op.payload)),
      ],
    });
  }
  rows.push({
    sheetKey: "shpo",
    sheetLabel: "1. ШПО",
    role: op.payload.shpoExcelRow
      ? `Поставити на індекс · R${op.payload.shpoExcelRow}`
      : "Поставити на індекс",
    cells: [
      previewCell(1, "індекс", index, "keep"),
      previewCell(6, "звання", op.rank),
      previewCell(7, "ПІБ", op.fullName),
      previewCell(8, "ID", op.personId),
    ],
  });
  rows.push({
    sheetKey: "oos",
    sheetLabel: "2. ООС",
    role: op.payload.oosExcelRow
      ? `Оновити картку · R${op.payload.oosExcelRow}`
      : op.payload.excludedSourceExcelRow
        ? `Новий рядок з Виключених R${op.payload.excludedSourceExcelRow}`
        : "Нова активна картка",
    cells: [
      previewCell(1, "звання", op.rank),
      previewCell(2, "ПІБ", op.fullName),
      previewCell(3, "ID", op.personId),
      previewCell(4, "індекс", op.payload.oosHistoryIndexes || index),
    ],
  });
  return rows;
};

const rankPreviews = (op: EjoosSyncOp): SheetRowPreview[] => {
  const next = op.payload.nextRank || op.rank;
  return [
    {
      sheetKey: "shpo",
      sheetLabel: "1. ШПО",
      role: op.payload.shpoExcelRow
        ? `Змінити звання · R${op.payload.shpoExcelRow}`
        : "Змінити звання",
      cells: [previewCell(6, "звання", next)],
    },
    {
      sheetKey: "oos",
      sheetLabel: "2. ООС",
      role: op.payload.oosExcelRow
        ? `Змінити звання · R${op.payload.oosExcelRow}`
        : "Змінити звання",
      cells: [previewCell(1, "звання", next)],
    },
  ];
};

const contractPreviews = (op: EjoosSyncOp): SheetRowPreview[] => [
  {
    sheetKey: "oos",
    sheetLabel: "2. ООС",
    role: op.payload.oosExcelRow
      ? `Заповнити контракт · R${op.payload.oosExcelRow}`
      : "Заповнити контракт",
    cells: [
      previewCell(19, "вид служби", op.payload.serviceType || "контракт"),
      previewCell(20, "дата укладання", op.payload.contractFrom),
      previewCell(21, "дата закінчення", op.payload.contractTo),
    ],
  },
];

/** Як виглядатимуть рядки Excel на аркушах, які ця особа змінює. */
export const buildSheetRowPreviews = (
  ops: EjoosSyncOp[],
  timesheetDay = 31,
): SheetRowPreview[] => {
  const exclude = ops.find((op) => op.kind === "exclude_transfer");
  if (exclude) return excludePreviews(exclude, timesheetDay);

  const position = ops.find((op) => op.kind === "position_change");
  if (position) return positionPreviews(position);

  const occupant = ops.find((op) => op.kind === "shpo_occupant");
  if (occupant) {
    return [
      {
        sheetKey: "shpo",
        sheetLabel: "1. ШПО",
        role: occupant.payload.shpoExcelRow
          ? `R${occupant.payload.shpoExcelRow}`
          : "Зайняти посаду",
        cells: [
          previewCell(1, "індекс", occupant.positionIndex, "keep"),
          previewCell(6, "звання", occupant.rank),
          previewCell(7, "ПІБ", occupant.fullName),
          previewCell(8, "ID", occupant.personId),
        ],
      },
    ];
  }

  const rank = ops.find((op) => op.kind === "rank_change");
  const contract = ops.find((op) => op.kind === "contract_update");
  if (rank || contract) {
    return [
      ...(rank ? rankPreviews(rank) : []),
      ...(contract ? contractPreviews(contract) : []),
    ];
  }

  return [];
};
