import { createMovementKey } from "../../parse/pb";
import { normKey } from "../../parse/cellText";
import type { PbMovement } from "../../types/pb";
import type { EjoosSyncOp } from "../../types/syncOp";
import { opId } from "../opId";

export type ManualArrivalMovementInput = {
  event: PbMovement;
  latestPositionByName: Map<string, PbMovement>;
  personStillInEjoos: (personId: string, fullName: string) => boolean;
  wasMovementProcessed: (event: PbMovement) => boolean;
};

/** ПРИБУВ / ЗВІЛЬН без автоматичного apply — лише картка для оператора. */
export const planManualArrivalMovementOp = (
  input: ManualArrivalMovementInput,
): EjoosSyncOp | null => {
  const { event } = input;
  if (event.type !== "ПРИБУВ" && event.type !== "ЗВІЛЬН") return null;
  if (
    (event.type === "ПРИБУВ" &&
      input.personStillInEjoos(event.personId, event.fullName)) ||
    (event.type === "ЗВІЛЬН" &&
      !input.personStillInEjoos(event.personId, event.fullName))
  ) {
    return null;
  }
  const laterPlacement = input.latestPositionByName.get(normKey(event.fullName));
  if (
    event.type === "ПРИБУВ" &&
    laterPlacement &&
    laterPlacement.excelRow > event.excelRow
  ) {
    return null;
  }
  const processedBefore = input.wasMovementProcessed(event);
  return {
    id: opId([
      "mov",
      event.type,
      event.movementNumber || String(event.excelRow),
    ]),
    kind: event.type === "ПРИБУВ" ? "arrival" : "other_manual",
    class: "needs_input",
    sheet: event.type === "ПРИБУВ" ? "2. ООС" : "3. Виключені",
    personId: event.personId,
    fullName: event.fullName,
    positionIndex: event.nextIndex || event.previousIndex,
    rank: event.rank,
    before: "—",
    after: `${event.type}: ${event.destination || event.changeText || "потрібні реквізити"}`,
    sourceRef: `Рух!R${event.excelRow} №${event.movementNumber}`,
    why: processedBefore
      ? "Подію вже застосовували, але поточний стан ЕЖООС їй не відповідає"
      : `${event.type} не застосовується автоматично — заповніть накази/звідки/куди`,
    confidence: "manual",
    payload: {
      movementNumber: event.movementNumber,
      type: event.type,
      destination: event.destination,
      orderNumber: event.orderNumber,
      orderDate: event.orderDate,
    },
    movementKey: createMovementKey(event),
    checkedDefault: false,
  };
};
