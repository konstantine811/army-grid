import JSZip from "jszip";
import type {
  CellValue,
  ExcelSheetSnapshot,
  ExcelWorkbookSnapshot,
} from "../../excelRoundTrip";
import { EXCLUDED_TO_OOS_BASE } from "./ejoosExcludedColumns";
import {
  cellValueToOosText,
  findOosStyleSourceRow,
  formatOosRelativesText,
  isJammedOosHistory,
  isOosWrapColumn,
  mergeOosHistoryValue,
  OOS_DATA_COLUMNS,
  OOS_HISTORY_COLUMNS,
  OOS_RANK_ORDER_COLUMNS,
  OOS_RELATIVES_COLUMN,
  OOS_RESTYLE_COLUMNS,
  oosStyleSourceColumn,
  splitOosHistoryLines,
} from "./ejoosOosText";
import {
  findEjoosSheet,
  parseEjoosOos,
  type EjoosSyncOp,
  type EjoosSyncPlan,
} from "./ejoosSyncPlan";
import {
  applyInlineStringWritesToWorkbook,
  resolveSheetPath,
  type ZipCellWrite,
} from "./ejoosZipCellWrites";

const cellAt = (
  sheet: ExcelSheetSnapshot,
  row: number,
  column: number,
): CellValue | undefined => sheet.rawRows[row - 1]?.[column - 1];

const excludedColumnForOos = (oosColumn: number) =>
  EXCLUDED_TO_OOS_BASE.find(([, toCol]) => toCol === oosColumn)?.[0];

const nextEmptyOosRow = (
  sheet: ExcelSheetSnapshot,
  reserved: Set<number>,
) => {
  for (let row = 6; row <= sheet.rawRows.length + 30; row += 1) {
    if (reserved.has(row)) continue;
    const name = cellValueToOosText(cellAt(sheet, row, 2));
    const id = cellValueToOosText(cellAt(sheet, row, 3));
    if (!name && !id) {
      reserved.add(row);
      return row;
    }
  }
  const row = sheet.rawRows.length + reserved.size + 1;
  reserved.add(row);
  return row;
};

const oosIdentityNumber = (column: number, value: string | null) =>
  (column === 3 || column === 15) && value && /^\d+$/.test(value.trim())
    ? Number(value)
    : null;

const oosCellValue = (column: number, value: string | null) => {
  if (!value) return null;
  const asNumber = oosIdentityNumber(column, value);
  if (asNumber != null) return asNumber;
  if (column === OOS_RELATIVES_COLUMN) return value;
  if (isOosWrapColumn(column)) return splitOosHistoryLines(value).join("\n");
  return value;
};

const oosStyledWrite = (
  row: number,
  column: number,
  value: string | null,
  styleSourceRow: number,
): ZipCellWrite => ({
  row,
  column,
  value: oosCellValue(column, value),
  copyNeighborStyle: true,
  styleSourceRow,
  styleSourceColumn: oosStyleSourceColumn(column),
  wrapText: true,
});

const oosStyleOnlyWrite = (
  row: number,
  column: number,
  styleSourceRow: number,
): ZipCellWrite => ({
  row,
  column,
  value: null,
  styleOnly: true,
  copyNeighborStyle: true,
  styleSourceRow,
  styleSourceColumn: oosStyleSourceColumn(column),
  wrapText: true,
});

const personTouchKey = (op: EjoosSyncOp) =>
  String(op.personId || op.fullName || "")
    .trim()
    .toLocaleLowerCase("uk-UA");

const shouldTouchOos = (op: EjoosSyncOp, excludedKeys: Set<string>) => {
  if (op.kind === "exclude_transfer") return false;
  const key = personTouchKey(op);
  if (key && excludedKeys.has(key)) return false;
  return (
    op.kind === "position_change" ||
    op.kind === "rank_change" ||
    (op.kind === "shpo_occupant" &&
      Number(op.payload.excludedSourceExcelRow || 0) > 0) ||
    Number(op.payload.oosExcelRow || 0) > 0
  );
};

/** Рядки ООС, де вже є ПІБ, але Excel зняв `s=` — індекси злипаються. */
const findUnstyledOosDataRows = async (file: Blob, sheetName: string) => {
  const zip = await JSZip.loadAsync(await file.arrayBuffer());
  const path = await resolveSheetPath(zip, sheetName);
  if (!path) return [];
  const xml = await zip.file(path)?.async("string");
  if (!xml) return [];
  const rows: number[] = [];
  for (const rowMatch of xml.matchAll(
    /<row\b[^>]*\br="(\d+)"[^>]*(?:\/>|>[\s\S]*?<\/row>)/gi,
  )) {
    const row = Number(rowMatch[1]);
    if (row < 7) continue;
    const rowXml = rowMatch[0];
    const nameCell = rowXml.match(
      /<c\b([^<>]*\br="B\d+"[^<>]*)(?:\/>|>[\s\S]*?<\/c>)/i,
    );
    if (!nameCell || !/<(?:v|is)\b/i.test(nameCell[0])) continue;
    const missingStyle = [...rowXml.matchAll(/<c\b([^<>]*)(?:\/>|>[\s\S]*?<\/c>)/gi)].some(
      (cell) =>
        /<(?:v|is)\b/i.test(cell[0]) && !/\bs="\d+"/i.test(cell[1]),
    );
    if (missingStyle) rows.push(row);
  }
  return rows;
};

/**
 * xlsx-populate псує ООС. Пишемо OOXML-ом як Excel: shared string + &#10;.
 * inlineStr Excel лагодить («Repaired Records: Cell information»).
 */
export async function applyOosHistoryPresentation(input: {
  file: Blob;
  ejoos: ExcelWorkbookSnapshot;
  ops: EjoosSyncOp[];
  plan: EjoosSyncPlan;
}): Promise<Blob> {
  const oos = findEjoosSheet(input.ejoos, /(^|[.\s])оос($|[\s])/i);
  const excluded = findEjoosSheet(input.ejoos, /виключен/i);
  if (!oos) return input.file;

  const excludedKeys = new Set(
    input.ops
      .filter((op) => op.kind === "exclude_transfer")
      .map(personTouchKey)
      .filter(Boolean),
  );
  const touchOps = input.ops.filter((op) => shouldTouchOos(op, excludedKeys));

  const oosPeople = parseEjoosOos(oos);
  const reserved = new Set<number>();
  const rowByPersonKey = new Map<string, number>();
  const writes: ZipCellWrite[] = [];

  const resolveRow = (op: EjoosSyncOp) => {
    const key = personTouchKey(op);
    if (key && rowByPersonKey.has(key)) return rowByPersonKey.get(key)!;

    const known = Number(op.payload.oosExcelRow || 0);
    if (known > 0) {
      if (key) rowByPersonKey.set(key, known);
      return known;
    }
    const personId = op.payload.nextPersonId || op.personId;
    const fullName = (op.payload.nextName || op.fullName).toLocaleLowerCase(
      "uk-UA",
    );
    const found =
      (personId && oosPeople.find((row) => row.personId === personId)) ||
      oosPeople.find(
        (row) => row.fullName.toLocaleLowerCase("uk-UA") === fullName,
      );
    if (found) {
      if (key) rowByPersonKey.set(key, found.excelRow);
      return found.excelRow;
    }
    if (
      op.payload.excludedSourceExcelRow ||
      op.kind === "position_change" ||
      op.payload.isTempArrivalPlacement === "1"
    ) {
      const row = nextEmptyOosRow(oos, reserved);
      if (key) rowByPersonKey.set(key, row);
      return row;
    }
    return 0;
  };

  const rowByOpId = new Map<string, number>();
  const targetRows = new Set<number>();
  for (const op of touchOps) {
    const row = resolveRow(op);
    if (!row) continue;
    rowByOpId.set(op.id, row);
    targetRows.add(row);
  }
  const styleSourceRow = findOosStyleSourceRow(
    (row, column) => cellAt(oos, row, column),
    { skipRows: targetRows, lastRow: oos.rawRows.length },
  );

  const sourceValue = (
    op: EjoosSyncOp,
    row: number,
    column: number,
  ): CellValue | undefined => {
    const excludedRow = Number(op.payload.excludedSourceExcelRow || 0);
    const isNewCard = !Number(op.payload.oosExcelRow || 0) && excludedRow > 0;
    if (isNewCard && excluded) {
      const fromCol = excludedColumnForOos(column);
      if (fromCol) return cellAt(excluded, excludedRow, fromCol);
    }
    return cellAt(oos, row, column);
  };

  const push = (write: ZipCellWrite | null) => {
    if (write) writes.push(write);
  };

  for (const op of touchOps) {
    const row = rowByOpId.get(op.id) || 0;
    if (!row) continue;
    const rank = op.payload.nextRank || op.rank;
    const fullName = op.payload.nextName || op.fullName;
    const personId = op.payload.nextPersonId || op.personId;
    const nextIndex = op.payload.nextIndex || op.positionIndex;
    const appointment = op.payload.orderDate || input.plan.timesheetDayLabel;
    const excludedSourceRow = Number(op.payload.excludedSourceExcelRow || 0);
    const fromExcludedCard =
      excludedSourceRow > 0 && !Number(op.payload.oosExcelRow || 0);
    const isNewCard =
      fromExcludedCard ||
      (!Number(op.payload.oosExcelRow || 0) &&
        (op.kind === "position_change" ||
          op.payload.isTempArrivalPlacement === "1"));
    const appendHistory = op.kind === "position_change" || isNewCard;
    const isRankChange = op.kind === "rank_change";

    if (fromExcludedCard && excluded) {
      for (const [fromCol, toCol] of EXCLUDED_TO_OOS_BASE) {
        const raw = cellAt(excluded, excludedSourceRow, fromCol);
        const text = cellValueToOosText(raw);
        if (!text) continue;
        const value =
          toCol === OOS_RELATIVES_COLUMN
            ? formatOosRelativesText(raw) || null
            : text;
        push(oosStyledWrite(row, toCol, value, styleSourceRow));
      }
    }

    const valueFor = (column: number): string | null => {
      const current = cellValueToOosText(sourceValue(op, row, column));
      if (column === 1) {
        return rank || current || null;
      }
      if (isRankChange && column === 15) {
        return op.payload.orderDate || null;
      }
      if (isRankChange && column === 16) {
        return op.payload.orderNumber || null;
      }
      if (column === 2) {
        return fullName || current || null;
      }
      if (column === 3) {
        return personId || current || null;
      }
      if (column === 4) {
        return mergeOosHistoryValue(
          sourceValue(op, row, 4),
          appendHistory
            ? op.payload.oosHistoryIndexes || nextIndex
            : "",
        );
      }
      if (column === 5) {
        const historyDates = appendHistory
          ? op.payload.oosHistoryDates ||
            (op.payload.oosHistoryIndexes || nextIndex || "")
              .split("\n")
              .filter(Boolean)
              .map(() => appointment)
              .join("\n") || appointment
          : "";
        return mergeOosHistoryValue(sourceValue(op, row, 5), historyDates);
      }
      if (column === 7) {
        return current || op.payload.arrivedFrom || null;
      }
      if (column === 8) {
        return current || op.payload.enlistDate || op.payload.orderDate || null;
      }
      if (column === 9) {
        return (
          current ||
          op.payload.enlistOrderDate ||
          op.payload.orderDate ||
          null
        );
      }
      if (column === 10) {
        return (
          current ||
          op.payload.enlistOrderNumber ||
          op.payload.orderNumber ||
          null
        );
      }
      if (column === 11) {
        return mergeOosHistoryValue(
          sourceValue(op, row, 11),
          appendHistory && op.kind === "position_change"
            ? op.payload.orderNumber || ""
            : "",
        );
      }
      if (column === 12) {
        return mergeOosHistoryValue(
          sourceValue(op, row, 12),
          appendHistory && op.kind === "position_change" ? appointment : "",
        );
      }
      return current || null;
    };

    const rankColumns = isRankChange ? [...OOS_RANK_ORDER_COLUMNS] : [];
    if (fromExcludedCard && op.kind === "position_change") {
      const historyUpdates: Array<[number, string]> = [
        [4, nextIndex],
        [5, appointment],
        [11, op.payload.orderNumber || ""],
        [12, appointment],
      ];
      for (const [column, extra] of historyUpdates) {
        if (!extra) continue;
        const fromCol = excludedColumnForOos(column);
        const current =
          fromCol && excluded
            ? cellAt(excluded, excludedSourceRow, fromCol)
            : undefined;
        const merged = mergeOosHistoryValue(current, extra);
        if (!merged) continue;
        push(oosStyledWrite(row, column, merged, styleSourceRow));
      }
      if (rank) push(oosStyledWrite(row, 1, rank, styleSourceRow));
      if (fullName) push(oosStyledWrite(row, 2, fullName, styleSourceRow));
      if (personId) push(oosStyledWrite(row, 3, personId, styleSourceRow));
    } else if (!fromExcludedCard && !isRankChange) {
      for (const column of [...OOS_RESTYLE_COLUMNS, ...rankColumns]) {
        const value = valueFor(column);
        if (!value) continue;
        push(oosStyledWrite(row, column, value, styleSourceRow));
      }

      if (isNewCard) {
        const relatives = formatOosRelativesText(
          sourceValue(op, row, OOS_RELATIVES_COLUMN),
        );
        push(
          oosStyledWrite(
            row,
            OOS_RELATIVES_COLUMN,
            relatives || null,
            styleSourceRow,
          ),
        );
      }
    } else if (isRankChange) {
      for (const column of [1, ...rankColumns]) {
        const value = valueFor(column);
        if (!value) continue;
        push(oosStyledWrite(row, column, value, styleSourceRow));
      }
    }

    for (const column of OOS_DATA_COLUMNS) {
      if (
        writes.some(
          (write) =>
            write.row === row && write.column === column && !write.styleOnly,
        )
      ) {
        continue;
      }
      push(oosStyleOnlyWrite(row, column, styleSourceRow));
    }
  }

  for (let i = 5; i < oos.rawRows.length; i += 1) {
    const row = i + 1;
    if (targetRows.has(row)) continue;
    for (const column of OOS_HISTORY_COLUMNS) {
      const raw = cellAt(oos, row, column);
      if (!isJammedOosHistory(raw)) continue;
      const value = splitOosHistoryLines(cellValueToOosText(raw)).join("\n");
      push(oosStyledWrite(row, column, value, styleSourceRow));
    }
  }

  for (const row of await findUnstyledOosDataRows(input.file, oos.sheetName)) {
    if (targetRows.has(row)) continue;
    for (const column of [...OOS_RESTYLE_COLUMNS, ...OOS_RANK_ORDER_COLUMNS]) {
      const value = cellValueToOosText(cellAt(oos, row, column)) || null;
      if (!value) continue;
      push(oosStyledWrite(row, column, value, styleSourceRow));
    }
    for (const column of OOS_DATA_COLUMNS) {
      if (
        writes.some(
          (write) =>
            write.row === row && write.column === column && !write.styleOnly,
        )
      ) {
        continue;
      }
      push(oosStyleOnlyWrite(row, column, styleSourceRow));
    }
  }

  if (!writes.length) return input.file;
  return applyInlineStringWritesToWorkbook(input.file, oos.sheetName, writes);
}

/** Звання в ШПО/Табелі: лише текст, власний `s=` рядка не чіпаємо. */
export async function applyRankLabelsWithZip(input: {
  file: Blob;
  ops: EjoosSyncOp[];
}): Promise<Blob> {
  const rankOps = input.ops.filter((op) => op.kind === "rank_change");
  if (!rankOps.length) return input.file;

  const shpoWrites: ZipCellWrite[] = [];
  const timesheetWrites: ZipCellWrite[] = [];
  for (const op of rankOps) {
    const rank = op.payload.nextRank || op.rank;
    if (!rank) continue;
    const shpoRow = Number(op.payload.shpoExcelRow || 0);
    const timesheetRow = Number(op.payload.timesheetExcelRow || 0);
    if (shpoRow > 0) {
      shpoWrites.push({
        row: shpoRow,
        column: 6,
        value: rank,
        copyNeighborStyle: true,
      });
    }
    if (timesheetRow > 0) {
      timesheetWrites.push({
        row: timesheetRow,
        column: 6,
        value: rank,
        copyNeighborStyle: true,
      });
    }
  }

  let file = input.file;
  if (shpoWrites.length) {
    file = await applyInlineStringWritesToWorkbook(
      file,
      /шпо|штатно.?посад/i,
      shpoWrites,
    );
  }
  if (timesheetWrites.length) {
    file = await applyInlineStringWritesToWorkbook(
      file,
      /табель/i,
      timesheetWrites,
    );
  }
  return file;
}
