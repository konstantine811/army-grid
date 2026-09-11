import type { PbMovement } from "../types/pb";
import { dateMs, norm } from "../parse/cellText";

const formatTransferDestinationForTimesheet = (value: string) => {
  const codes = [...value.matchAll(/[АA]\s*(\d{4})(?!\d)/giu)].map(
    (match) => `А${match[1]}`,
  );
  const unique = [...new Set(codes)];
  return unique.length ? unique.join(" / ") : value;
};

/**
 * «2103791 Старший кухар → 2103179 Стрілець 3 піхотного відділення …» —
 * для «куди вибув» потрібна лише нова посада без службового індексу.
 */
export const positionChangeDestination = (event: PbMovement) => {
  const text = norm(event.changeText);
  if (!text) return norm(event.destination);
  const tail =
    text
      .split(/→|->|=>/)
      .pop()
      ?.trim() || text;
  return tail.replace(/^\d{5,}[\s.:;-]*/, "").trim() || tail;
};

/**
 * У колонці повернення archive часто стоїть «до окремого розпорядження», «0»
 * або «-». Період вважаємо закритим лише за фактичною датою.
 */
export const hasActualReturn = (value: string) => Boolean(dateMs(value));

/** У `sh` замість індексу бувають маркери «ВИВЕДЕНО», «#N/A» тощо. */
export const isPositionIndex = (value: string) => /^\d{5,}$/.test(value.trim());

export const isRankAssignmentEvent = (event: PbMovement) => {
  if (event.type === "ЗВАННЯ" || event.type.startsWith("ЗВАН")) return true;
  return false;
};

export const unitCodeFromMovement = (event: PbMovement) =>
  formatTransferDestinationForTimesheet(
    [event.destination, event.changeText, event.note].join(" "),
  );

export const formatTransferDestinationForTimesheetMark = formatTransferDestinationForTimesheet;
