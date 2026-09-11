import { createMovementKey } from "../../parse/pb";
import { isAmbiguousStaffTransfer } from "../../../ejoosMovementRules";
import type { PbMovement } from "../../types/pb";
import type { EjoosSyncOp } from "../../types/syncOp";
import { opId } from "../opId";

/** Невизначене внутрішнє/зовнішнє переведення — блокуємо автозастосування. */
export const planTransferScopeUnclearOp = (
  event: PbMovement,
): EjoosSyncOp | null => {
  if (!isAmbiguousStaffTransfer(event)) return null;
  return {
    id: opId(["scope", event.personId || event.fullName, String(event.excelRow)]),
    kind: "other_manual",
    class: "needs_input",
    sheet: "Рух / Виключені або ШПО",
    personId: event.personId,
    fullName: event.fullName,
    positionIndex: event.nextIndex || event.previousIndex,
    rank: event.rank,
    before: event.destination || event.changeText || "ПЕРЕВ",
    after: "уточнити: внутрішня зміна посади чи вибуття до іншої в/ч",
    sourceRef: `Рух!R${event.excelRow} №${event.movementNumber}`,
    why: "Не визначено, чи це переведення всередині 1ПБ, чи вибуття до іншої військової частини. Без «куди вибув» / в/ч А#### масово не застосовуємо.",
    confidence: "manual",
    payload: {
      type: "TRANSFER_SCOPE_UNCLEAR",
      transferScope: "unclear",
      destination: event.destination,
      note: event.note,
      changeText: event.changeText,
      orderNumber: event.orderNumber,
      orderDate: event.orderDate,
      previousIndex: event.previousIndex,
      nextIndex: event.nextIndex,
    },
    movementKey: createMovementKey(event),
    checkedDefault: false,
  };
};
