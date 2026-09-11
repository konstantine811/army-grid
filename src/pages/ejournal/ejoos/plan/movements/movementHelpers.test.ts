import { describe, expect, it } from "vitest";
import type { PbMovement } from "../../types/pb";
import { planManualArrivalMovementOp } from "./manualArrivalMovement";
import { planTransferScopeUnclearOp } from "./transferScopeUnclear";

describe("movement plan helpers", () => {
  it("planTransferScopeUnclearOp flags ambiguous internal/external transfer", () => {
    const event = {
      excelRow: 5,
      movementNumber: "12",
      type: "ПЕРЕВ",
      personId: "100",
      fullName: "Тестов Т.",
      rank: "солдат",
      previousIndex: "2103001",
      nextIndex: "",
      destination: "",
      orderNumber: "1",
      orderDate: "01.08.2026",
      basisNumber: "",
      basisDate: "",
      changeText: "",
      status: "",
      note: "",
      arrivedFrom: "",
    } satisfies PbMovement;

    const op = planTransferScopeUnclearOp(event);
    expect(op?.payload.transferScope).toBe("unclear");
    expect(op?.kind).toBe("other_manual");
  });

  it("planManualArrivalMovementOp creates arrival card for ПРИБУВ", () => {
    const op = planManualArrivalMovementOp({
      event: {
        excelRow: 3,
        movementNumber: "2",
        type: "ПРИБУВ",
        personId: "200",
        fullName: "Новий Н.",
        rank: "солдат",
        previousIndex: "БРЕЗ",
        nextIndex: "2103100",
        destination: "БРЕЗ",
        orderNumber: "10",
        orderDate: "05.08.2026",
        basisNumber: "",
        basisDate: "",
        changeText: "",
        status: "",
        note: "",
        arrivedFrom: "",
      },
      latestPositionByName: new Map(),
      personStillInEjoos: () => false,
      wasMovementProcessed: () => false,
    });
    expect(op?.kind).toBe("arrival");
    expect(op?.sheet).toBe("2. ООС");
  });
});
