import { describe, expect, it } from "vitest";
import {
  findUnvacatedTargetOccupant,
  type EjoosShpoRow,
  type EjoosSyncOp,
} from "./ejoosSyncPlan";

const positionOp = (
  overrides: Partial<EjoosSyncOp> = {},
): EjoosSyncOp => ({
  id: "position-new",
  kind: "position_change",
  class: "ready",
  sheet: "1. ШПО / 2. ООС / 6. Табель",
  personId: "1961288",
  fullName: "КІЯНЕНКО Андрій Олександрович",
  positionIndex: "2103119",
  rank: "старший лейтенант",
  before: "попередня посада",
  after: "штатна посада 2103119",
  sourceRef: "Рух!R1",
  why: "ПОСАДА",
  confidence: "high",
  payload: { previousIndex: "2103117", nextIndex: "2103119" },
  checkedDefault: true,
  ...overrides,
});

const incumbent: EjoosShpoRow = {
  excelRow: 42,
  positionIndex: "2103119",
  personId: "incumbent-id",
  fullName: "СИДОРЕНКО Єгор Олегович",
  rank: "солдат",
};

describe("position target occupancy conflict", () => {
  it("blocks appointment when the SHPO occupant has no vacating operation", () => {
    const op = positionOp();

    expect(findUnvacatedTargetOccupant(op, [incumbent], [op])).toEqual(
      incumbent,
    );
  });

  it("allows appointment after a ready departure of the incumbent", () => {
    const op = positionOp();
    const departure = positionOp({
      id: "incumbent-departure",
      kind: "exclude_transfer",
      personId: incumbent.personId,
      fullName: incumbent.fullName,
      positionIndex: incumbent.positionIndex,
      payload: { fromPositionIndex: incumbent.positionIndex },
    });

    expect(
      findUnvacatedTargetOccupant(op, [incumbent], [op, departure]),
    ).toBeNull();
  });
});
