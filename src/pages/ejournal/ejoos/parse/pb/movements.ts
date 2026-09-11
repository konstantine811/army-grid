import type { ExcelWorkbookSnapshot } from "../../../../../excelRoundTrip";
import { resolveOutboundTransferDestination } from "../../../ejoosMovementRules";
import type { PbMovement } from "../../types/pb";
import { norm, normKey } from "../cellText";
import { cell, findCol, findEjoosSheet, headerMap, idCell } from "../sheetLookup";
import {
  isCancelledMovementRecord,
  resolveMovementDestination,
} from "./movementHelpers";

export { isCancelledMovementRecord } from "./movementHelpers";

export const createMovementKey = (event: PbMovement) => {
  const identity = event.personId
    ? `id:${normKey(event.personId)}`
    : `name:${normKey(event.fullName)}`;
  if (identity.endsWith(":")) return "";
  return [
    identity,
    normKey(event.type),
    normKey(event.orderDate),
    normKey(event.orderNumber),
    normKey(event.previousIndex),
    normKey(event.nextIndex),
  ].join("|");
};

export const collectProcessedMovementKeys = (
  versions: Array<{ changeProtocol?: unknown }> | null | undefined,
) => {
  const keys = new Set<string>();
  for (const version of versions ?? []) {
    const protocol = version.changeProtocol;
    if (!protocol || typeof protocol !== "object") continue;
    const ops = (protocol as { ops?: unknown }).ops;
    if (!Array.isArray(ops)) continue;
    for (const op of ops) {
      if (!op || typeof op !== "object") continue;
      const key = (op as { movementKey?: unknown }).movementKey;
      if (typeof key === "string" && key.trim()) keys.add(key.trim());
    }
  }
  return keys;
};

const normalizeRankLabel = (value: string) =>
  value
    .replace(/\s+/g, " ")
    .trim()
    .replace(/^(?:звання|військов(?:е|ое)\s+звання)\s+/i, "")
    .toLocaleLowerCase("uk-UA");

/** «солдат → СТАРШИЙ СОЛДАТ» або лише нове звання в «Яка зміна». */
export const parseRankPromotion = (event: PbMovement) => {
  const text = [event.changeText, event.note].filter(Boolean).join(" ");
  const match = text.match(
    /([А-ЯІЇЄҐа-яіїєґ'’.\-\s]{3,}?)\s*(?:→|->|=>|—)\s*([А-ЯІЇЄҐа-яіїєґ'’.\-\s]{3,})/u,
  );
  if (match) {
    return {
      previousRank: normalizeRankLabel(match[1]),
      nextRank: normalizeRankLabel(match[2]),
    };
  }
  const nextRank =
    normalizeRankLabel(event.changeText) || normalizeRankLabel(event.rank);
  const previousRank = normalizeRankLabel(event.rank);
  return {
    previousRank: previousRank === nextRank ? "" : previousRank,
    nextRank,
  };
};

export const isContractMovementType = (type: string) =>
  type === "КОНТРАКТ" || /МОТИВАЦ.*КОНТР/iu.test(type);

const normalizeContractDate = (value: string) => {
  const match = norm(value).match(
    /^(\d{1,2})[./-](\d{1,2})[./-](\d{2}|\d{4})$/,
  );
  if (!match) return norm(value);
  const year = match[3].length === 2 ? `20${match[3]}` : match[3];
  return `${match[1].padStart(2, "0")}.${match[2].padStart(2, "0")}.${year}`;
};

export const parseContractDatesFromChangeText = (value: string) => {
  const match = norm(value).match(
    /(?:^|\s)(?:з|із)\s+(\d{1,2}[./-]\d{1,2}[./-](?:\d{2}|\d{4}))\s+до\s+(\d{1,2}[./-]\d{1,2}[./-](?:\d{2}|\d{4}))(?:$|[\s,.;])/iu,
  );
  return {
    contractFrom: match ? normalizeContractDate(match[1]) : "",
    contractTo: match ? normalizeContractDate(match[2]) : "",
  };
};

export const parsePbMovements = (
  workbook: ExcelWorkbookSnapshot,
): PbMovement[] => {
  const sheet = findEjoosSheet(workbook, /^рух$/i);
  if (!sheet) return [];
  const headers = headerMap(sheet.rawRows[0] ?? []);
  const numCol = findCol(headers, /^№$/);
  const typeCol = findCol(headers, /^тип$/);
  const idCol = findCol(headers, /^id$/);
  const nameCol = findCol(headers, /^піб$/);
  const rankCol = findCol(headers, /зван/);
  const statusCol = findCol(headers, /^статус$/);
  const prevIdxCol = findCol(headers, /індекс.*попер|попер.*індекс/);
  const nextIdxCol = findCol(
    headers,
    /індекс.*як|яка зміна.*індекс|індекси посад \(яка/,
  );
  const changeCol = (() => {
    for (const [key, index] of headers.entries()) {
      if (/яка зміна/.test(key) && !/індекс/.test(key)) return index;
    }
    return findCol(headers, /попер/);
  })();
  const destCol = findCol(headers, /^куди(?:\s|$)/, /куди.*(?:перев|виб)/);
  const fromCol = findCol(headers, /звідки.*прибув|^звідки$/);
  const noteCol = findCol(headers, /^примітка$/);
  const orderNumCol = findCol(headers, /^наказ$/);
  const orderDateCol = findCol(headers, /^дата$/);
  const movements: PbMovement[] = [];

  const normalizeType = (value: string) => {
    const text = value.toUpperCase();
    if (text.includes("ПОСАД")) return "ПОСАДА";
    if (text.includes("ПРИБ")) return "ПРИБУВ";
    if (text.includes("ЗМІНИШТАТ") || text.includes("ЗМІНИ ШТАТ")) {
      return "ПОСАДА";
    }
    if (text.includes("РОЗПОР")) return "РОЗПОРЯДЖ";
    if (text.includes("СКАС")) {
      if (text.includes("СЗЧ") || text.includes("САМОВІЛ")) return "СЗЧ";
      return "СКАСУВАННЯ";
    }
    if (text.includes("ПЕРЕВ")) return "ПЕРЕВ";
    if (text.includes("ЗВАН")) return "ЗВАННЯ";
    if (text.includes("ЗВІЛ")) return "ЗВІЛЬН";
    if (text.includes("СЗЧ") || text.includes("САМОВІЛ")) return "СЗЧ";
    if (
      text.includes("БЕЗВІСТ") ||
      /(?:^|[^А-ЯІЇЄҐ])ЗБ(?:$|[^А-ЯІЇЄҐ])/.test(text)
    ) {
      return "БЕЗВІСТИ";
    }
    return text || "—";
  };

  sheet.rawRows.slice(1).forEach((row, offset) => {
    const type = normalizeType(cell(row, typeCol >= 0 ? typeCol : 4));
    const fullName = cell(row, nameCol >= 0 ? nameCol : 6);
    const movementNumber = cell(row, numCol >= 0 ? numCol : 1);
    if (!fullName && !movementNumber) return;
    if (type === "—" && !fullName) return;
    const note = cell(row, noteCol >= 0 ? noteCol : 18);
    const rawDest = cell(row, destCol);
    movements.push({
      excelRow: offset + 2,
      movementNumber,
      type,
      personId: idCell(row, idCol),
      fullName,
      rank: cell(row, rankCol),
      previousIndex: cell(row, prevIdxCol),
      nextIndex: cell(row, nextIdxCol),
      destination:
        type === "ПЕРЕВ"
          ? resolveOutboundTransferDestination(rawDest, note)
          : resolveMovementDestination(rawDest, note),
      orderNumber: cell(row, orderNumCol),
      orderDate: cell(row, orderDateCol),
      basisNumber: cell(row, 14),
      basisDate: cell(row, 15),
      changeText: cell(row, changeCol),
      status: cell(row, statusCol >= 0 ? statusCol : 2),
      note,
      arrivedFrom: cell(row, fromCol),
    });
  });
  return movements;
};
