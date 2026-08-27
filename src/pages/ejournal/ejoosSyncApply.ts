import type { ExcelWorkbookSnapshot } from "../../excelRoundTrip";
import type { EjoosSyncOp, EjoosSyncPlan } from "./ejoosSyncPlan";
import { buildProtocolText } from "./ejoosSyncPlan";

const col = (letter: string) => {
  let index = 0;
  for (const char of letter.toUpperCase()) {
    index = index * 26 + char.charCodeAt(0) - 64;
  }
  return index; // 1-based
};

/**
 * Apply only confirmed ops to a copy of the live ЕЖООС workbook.
 * Does not write anything until the caller uploads the returned blob as a new DB version.
 */
export async function applyConfirmedEjoosOps(input: {
  ejoos: ExcelWorkbookSnapshot;
  plan: EjoosSyncPlan;
  ops: EjoosSyncOp[];
  actor?: string;
}): Promise<{
  blob: Blob;
  fileName: string;
  protocolText: string;
  changeProtocol: Record<string, unknown>;
}> {
  const { ejoos, plan, ops, actor } = input;
  if (!ops.length) {
    throw new Error("Немає підтверджених змін для застосування");
  }

  const fileName = `ЄЖООС_станом_на_${plan.timesheetDayLabel.replaceAll(".", "-")}.xlsx`;
  const blob = await mutateToBlob(ejoos, ops, plan, actor);
  const protocolText = buildProtocolText(plan, ops, {
    actor,
    at: new Date().toLocaleString("uk-UA"),
  });
  const changeProtocol = {
    kind: "apply",
    pbFileName: plan.pbName,
    timesheetDay: plan.timesheetDay,
    timesheetDayLabel: plan.timesheetDayLabel,
    appliedCount: ops.length,
    ops: ops.map((op) => ({
      id: op.id,
      kind: op.kind,
      sheet: op.sheet,
      personId: op.personId,
      fullName: op.fullName,
      positionIndex: op.positionIndex,
      before: op.before,
      after: op.after,
      sourceRef: op.sourceRef,
      why: op.why,
    })),
    protocolText,
  };

  return { blob, fileName, protocolText, changeProtocol };
}

async function mutateToBlob(
  ejoos: ExcelWorkbookSnapshot,
  ops: EjoosSyncOp[],
  plan: EjoosSyncPlan,
  actor?: string,
) {
  const module = await import("xlsx-populate/browser/xlsx-populate-no-encryption");
  const XlsxPopulate = module.default;
  const workbook = await XlsxPopulate.fromDataAsync(ejoos.file);

  const timesheet = workbook.sheet("6. Табель");
  const absent = workbook.sheet("5. Тимчасово відсутні");
  const excluded = workbook.sheet("3. Виключені");
  const oos = workbook.sheet("2. ООС");
  const shpo = workbook.sheet("1. ШПО");
  const history = workbook.sheet("10. Історія змін");

  ops.forEach((op) => {
    if (op.kind === "shpo_occupant") {
      const shpoRow = Number(op.payload.shpoExcelRow || 0);
      const tsRow = Number(op.payload.timesheetExcelRow || 0);
      const name = op.payload.nextName || op.fullName;
      const rank = op.payload.nextRank || op.rank;
      const personId = op.payload.nextPersonId || op.personId;
      if (shpo && shpoRow > 0) {
        if (rank) shpo.cell(shpoRow, col("F")).value(rank);
        if (name) shpo.cell(shpoRow, col("G")).value(name);
        if (personId) shpo.cell(shpoRow, col("H")).value(personId);
      }
      if (timesheet && tsRow > 0) {
        if (rank) timesheet.cell(tsRow, col("F")).value(rank);
        if (name) timesheet.cell(tsRow, col("G")).value(name);
        if (personId) timesheet.cell(tsRow, col("H")).value(personId);
      }
      return;
    }
    if (op.kind === "timesheet_day" && timesheet) {
      const rowNumber = Number(op.payload.excelRow || 0);
      const day = Number(op.payload.day || plan.timesheetDay);
      const code = op.payload.timesheetCode || String(op.after || "").trim();
      if (
        rowNumber > 0 &&
        day >= 1 &&
        day <= 31 &&
        code &&
        code !== "(оберіть код)"
      ) {
        timesheet.cell(rowNumber, col("I") + day - 1).value(code);
      }
      return;
    }
    if (op.kind === "absent_close" && absent) {
      const rowNumber = Number(op.payload.excelRow || 0);
      if (rowNumber > 0) {
        absent
          .cell(rowNumber, col("M"))
          .value(op.payload.returnDate || plan.timesheetDayLabel);
      }
      return;
    }
    if (op.kind === "absent_upsert" && absent) {
      const existing = Number(op.payload.existingExcelRow || 0);
      const targetRow = existing > 0 ? existing : nextEmptyAbsentRow(absent);
      if (op.rank) absent.cell(targetRow, col("A")).value(op.rank);
      if (op.fullName) absent.cell(targetRow, col("B")).value(op.fullName);
      if (op.personId) absent.cell(targetRow, col("C")).value(op.personId);
      if (op.positionIndex) {
        absent.cell(targetRow, col("D")).value(op.positionIndex);
      }
      if (op.payload.absenceType) {
        absent.cell(targetRow, col("E")).value(op.payload.absenceType);
      }
      if (op.payload.place) absent.cell(targetRow, col("F")).value(op.payload.place);
      if (op.payload.departDate) {
        absent.cell(targetRow, col("G")).value(op.payload.departDate);
      }
      const orderText = [op.payload.orderDate, op.payload.orderNumber]
        .filter(Boolean)
        .join(" ");
      if (orderText) absent.cell(targetRow, col("I")).value(orderText);
      if (op.payload.plannedReturn) {
        absent.cell(targetRow, col("L")).value(op.payload.plannedReturn);
      }
      absent
        .cell(targetRow, col("R"))
        .value(`archive №${op.payload.periodNumber || "—"}`);
      return;
    }
    if (op.kind === "exclude_transfer") {
      applyExcludeTransfer({
        op,
        plan,
        excluded,
        timesheet,
        shpo,
        oos,
      });
      return;
    }
  });

  if (history) {
    const start = nextAppendRow(history, 2);
    history.cell(start, 1).value(new Date().toISOString());
    history.cell(start, 2).value(actor || "user");
    history
      .cell(start, 3)
      .value(`Застосовано ${ops.length} змін з ${plan.pbName}`);
    history.cell(start, 4).value(plan.timesheetDayLabel);
  }

  return workbook.outputAsync("blob") as Promise<Blob>;
}

type SheetLike = {
  usedRange: () => { endCell: () => { rowNumber: () => number } } | undefined;
  cell: (
    r: number,
    c: number,
  ) => {
    value: (v?: unknown) => unknown;
  };
  row?: (r: number) => { height: (h?: number) => number };
};

/** ПЕРЕВ з обліку: Виключені → Табель → ШПО/ООС. Тимч. відсутні/прибулі не чіпаємо. */
function applyExcludeTransfer(input: {
  op: EjoosSyncOp;
  plan: EjoosSyncPlan;
  excluded: SheetLike | undefined;
  timesheet: SheetLike | undefined;
  shpo: SheetLike | undefined;
  oos: SheetLike | undefined;
}) {
  const { op, plan, excluded, timesheet, shpo, oos } = input;
  const personId = op.payload.fromPersonId || op.personId;
  const fullName = op.payload.fromName || op.fullName;
  const rank = op.payload.fromRank || op.rank;
  const positionIndex =
    op.payload.fromPositionIndex ||
    op.payload.previousIndex ||
    op.positionIndex;
  const destination =
    op.payload.destination || op.payload.documentsDest || "";
  const excludeDateLabel =
    op.payload.excludeDate ||
    op.payload.orderDate ||
    plan.timesheetDayLabel;
  const excludeDateValue = toExcelDateValue(excludeDateLabel);
  const departDay = dayFromLabel(excludeDateLabel) || plan.timesheetDay;

  let shpoRow = Number(op.payload.shpoExcelRow || 0);
  if (shpo && shpoRow <= 0) {
    shpoRow = findPersonRowFlexible(shpo, {
      personId,
      fullName,
      positionIndex,
      startRow: 7,
      nameCol: col("G"),
      idCol: col("H"),
      indexCol: col("A"),
    });
  }

  let timesheetRow = Number(op.payload.timesheetExcelRow || 0);
  if (timesheet && timesheetRow <= 0) {
    timesheetRow = findPersonRowFlexible(timesheet, {
      personId,
      fullName,
      positionIndex,
      startRow: 7,
      nameCol: col("G"),
      idCol: col("H"),
      indexCol: col("B"),
    });
  }

  // 1) Виключені — дані зі ШПО + службові з Рух
  if (excluded) {
    const targetRow = nextAppendRow(excluded, 6);
    if (shpo && shpoRow > 0) {
      // звання → ПІБ → ID → індекс + базові поля з ШПО
      copyCell(shpo, shpoRow, col("F"), excluded, targetRow, col("A"));
      copyCell(shpo, shpoRow, col("G"), excluded, targetRow, col("B"));
      copyCell(shpo, shpoRow, col("H"), excluded, targetRow, col("C"));
      copyCell(shpo, shpoRow, col("A"), excluded, targetRow, col("D"));
    } else {
      if (rank) excluded.cell(targetRow, col("A")).value(rank);
      if (fullName) excluded.cell(targetRow, col("B")).value(fullName);
      if (personId) excluded.cell(targetRow, col("C")).value(personId);
      if (positionIndex) excluded.cell(targetRow, col("D")).value(positionIndex);
    }
    excluded.cell(targetRow, col("AB")).value(excludeDateValue);
    excluded.cell(targetRow, col("AC")).value(excludeDateValue);
    excluded.cell(targetRow, col("AD")).value(op.payload.orderNumber || null);
    excluded.cell(targetRow, col("AE")).value(destination || null);
    excluded
      .cell(targetRow, col("AF"))
      .value(op.payload.type === "РОЗПОРЯДЖ" ? "Розпорядження" : "Переведення");
    excluded
      .cell(targetRow, col("AG"))
      .value(`Рух №${op.payload.movementNumber || "—"}`);
  }

  // 2) Табель — історія: копія рядка + + до дня вибуття, далі −; старий рядок без особи
  if (timesheet && timesheetRow > 0) {
    const historyRow = findTimesheetAppendNear(timesheet, timesheetRow);
    copySheetRow(timesheet, timesheetRow, historyRow, col("AN"));
    writeTimesheetTransferHistory(timesheet, historyRow, {
      departDay,
      destination,
      lastDay: Math.min(31, Math.max(departDay, plan.timesheetDay)),
    });
    // очистити персональну частину на старій посаді, лишити індекс/структуру
    timesheet.cell(timesheetRow, col("F")).value(null);
    timesheet.cell(timesheetRow, col("G")).value(null);
    timesheet.cell(timesheetRow, col("H")).value(null);
    for (let day = 1; day <= 31; day += 1) {
      timesheet.cell(timesheetRow, col("I") + day - 1).value(null);
    }
  }

  // 3) ШПО — прибрати звання/ПІБ/ID, лишити посаду
  if (shpo && shpoRow > 0) {
    shpo.cell(shpoRow, col("F")).value(null);
    shpo.cell(shpoRow, col("G")).value(null);
    shpo.cell(shpoRow, col("H")).value(null);
    shpo.cell(shpoRow, col("R")).value(null);
  }

  // 4) ООС — видалити рядок
  if (oos && personId) {
    const oosRow = findPersonIdRow(oos, personId, 6, col("C"));
    if (oosRow) clearOosRow(oos, oosRow);
  } else if (oos && fullName) {
    const oosRow = findPersonRowFlexible(oos, {
      personId: "",
      fullName,
      positionIndex: "",
      startRow: 6,
      nameCol: col("B"),
      idCol: col("C"),
      indexCol: col("D"),
    });
    if (oosRow) clearOosRow(oos, oosRow);
  }
}

function copyCell(
  from: SheetLike,
  fromRow: number,
  fromCol: number,
  to: SheetLike,
  toRow: number,
  toCol: number,
) {
  const value = from.cell(fromRow, fromCol).value();
  if (value !== undefined && value !== null && String(value).trim() !== "") {
    to.cell(toRow, toCol).value(value);
  }
}

function copySheetRow(
  sheet: SheetLike,
  sourceRow: number,
  targetRow: number,
  endCol: number,
) {
  for (let c = 1; c <= endCol; c += 1) {
    sheet.cell(targetRow, c).value(sheet.cell(sourceRow, c).value());
  }
}

function writeTimesheetTransferHistory(
  sheet: SheetLike,
  rowNumber: number,
  opts: { departDay: number; destination: string; lastDay: number },
) {
  const departureText = opts.destination
    ? `вибув до ${opts.destination}`
    : "вибув";
  let presentDays = 0;
  for (let day = 1; day <= 31; day += 1) {
    const cell = sheet.cell(rowNumber, col("I") + day - 1);
    if (day < opts.departDay) {
      cell.value("+");
      presentDays += 1;
    } else if (day === opts.departDay) {
      cell.value(departureText);
    } else if (day <= opts.lastDay) {
      cell.value("-");
    } else {
      cell.value(null);
    }
  }
  sheet.cell(rowNumber, col("AN")).value(presentDays);
}

function findTimesheetAppendNear(sheet: SheetLike, sourceRow: number) {
  const end = sheet.usedRange()?.endCell().rowNumber() ?? sourceRow;
  for (let row = sourceRow + 1; row <= end + 30; row += 1) {
    const index = String(sheet.cell(row, col("B")).value() ?? "").trim();
    const name = String(sheet.cell(row, col("G")).value() ?? "").trim();
    const id = String(sheet.cell(row, col("H")).value() ?? "").trim();
    if (!index && !name && !id) return row;
  }
  return nextAppendRow(sheet, sourceRow + 1);
}

function findPersonRowFlexible(
  sheet: SheetLike,
  opts: {
    personId: string;
    fullName: string;
    positionIndex: string;
    startRow: number;
    nameCol: number;
    idCol: number;
    indexCol: number;
  },
) {
  const end = sheet.usedRange()?.endCell().rowNumber() ?? opts.startRow;
  const nameKey = opts.fullName
    .trim()
    .toLowerCase()
    .replace(/[''`´]/g, "")
    .replace(/\s+/g, " ");
  const idKey = String(opts.personId ?? "").trim();
  let indexMatch = 0;
  for (let row = opts.startRow; row <= end; row += 1) {
    const id = String(sheet.cell(row, opts.idCol).value() ?? "").trim();
    const name = String(sheet.cell(row, opts.nameCol).value() ?? "")
      .trim()
      .toLowerCase()
      .replace(/[''`´']/g, "")
      .replace(/\s+/g, " ");
    const index = String(sheet.cell(row, opts.indexCol).value() ?? "").trim();
    if (idKey && id === idKey) return row;
    if (nameKey && name === nameKey) return row;
    if (
      !indexMatch &&
      opts.positionIndex &&
      index === opts.positionIndex &&
      (id || name)
    ) {
      indexMatch = row;
    }
  }
  // Якщо ПІБ/ID не збіглися (розбіжності написання) — посада з людиною на індексі.
  return indexMatch;
}

function dayFromLabel(label: string) {
  const match = String(label ?? "").match(/(\d{1,2})[./-]/);
  if (!match) return 0;
  const day = Number(match[1]);
  return day >= 1 && day <= 31 ? day : 0;
}

function toExcelDateValue(label: string) {
  const match = String(label ?? "").match(
    /(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})/,
  );
  if (!match) return label || null;
  const day = Number(match[1]);
  const month = Number(match[2]);
  const yearRaw = Number(match[3]);
  const year = yearRaw < 100 ? 2000 + yearRaw : yearRaw;
  return new Date(year, month - 1, day);
}

function nextAppendRow(
  sheet: {
    usedRange: () => { endCell: () => { rowNumber: () => number } } | undefined;
  },
  minRow: number,
) {
  const used = sheet.usedRange();
  return Math.max(minRow, (used?.endCell().rowNumber() ?? minRow - 1) + 1);
}

function nextEmptyAbsentRow(sheet: {
  usedRange: () => { endCell: () => { rowNumber: () => number } } | undefined;
  cell: (r: number, c: number) => { value: () => unknown };
}) {
  const end = sheet.usedRange()?.endCell().rowNumber() ?? 6;
  for (let row = 6; row <= end + 5; row += 1) {
    const name = String(sheet.cell(row, 2).value() ?? "").trim();
    if (!name) return row;
  }
  return end + 1;
}

function findPersonIdRow(
  sheet: {
    usedRange: () => { endCell: () => { rowNumber: () => number } } | undefined;
    cell: (r: number, c: number) => { value: () => unknown };
  },
  personId: string,
  startRow: number,
  idCol: number,
) {
  const end = sheet.usedRange()?.endCell().rowNumber() ?? startRow;
  for (let row = startRow; row <= end; row += 1) {
    if (String(sheet.cell(row, idCol).value() ?? "").trim() === personId) {
      return row;
    }
  }
  return 0;
}

function clearOosRow(
  sheet: {
    cell: (r: number, c: number) => { value: (v?: unknown) => unknown };
  },
  row: number,
) {
  for (let c = 1; c <= 40; c += 1) {
    sheet.cell(row, c).value(null);
  }
}

export async function blobToBase64(blob: Blob): Promise<string> {
  const buffer = await blob.arrayBuffer();
  const bytes = new Uint8Array(buffer);
  let binary = "";
  const chunk = 0x8000;
  for (let i = 0; i < bytes.length; i += chunk) {
    binary += String.fromCharCode(...bytes.subarray(i, i + chunk));
  }
  return btoa(binary);
}

export function downloadTextFile(fileName: string, text: string) {
  const blob = new Blob([text], { type: "text/plain;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = fileName;
  anchor.click();
  URL.revokeObjectURL(url);
}

export function downloadBlobFile(fileName: string, blob: Blob) {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = fileName;
  anchor.click();
  URL.revokeObjectURL(url);
}

export async function fileToBase64(file: File): Promise<string> {
  return blobToBase64(file);
}

export function base64ToFile(base64: string, fileName: string): File {
  const raw = base64.includes(",") ? base64.split(",")[1] : base64;
  const binary = atob(raw);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i += 1) bytes[i] = binary.charCodeAt(i);
  return new File(
    [bytes],
    fileName,
    {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    },
  );
}
