import { describe, expect, it } from "vitest";
import { payloadChanged, payloadFingerprint } from "./payloadFingerprint";

describe("payloadFingerprint", () => {
  it("treats identical row payloads as unchanged without full stringify", () => {
    const left = {
      importId: "imp-1",
      rows: [
        { id: "r1", excelRowNumber: 2, values: { column_14: "КОВАЛЬ" } },
        { id: "r2", excelRowNumber: 3, values: { column_14: "ПЕТРЕНКО" } },
      ],
    };
    const right = {
      importId: "imp-1",
      rows: [
        { id: "r1", excelRowNumber: 2, values: { column_14: "КОВАЛЬ" } },
        { id: "r2", excelRowNumber: 3, values: { column_14: "ПЕТРЕНКО" } },
      ],
    };

    expect(payloadFingerprint(left)).toBe(payloadFingerprint(right));
    expect(payloadChanged(left, right)).toBe(false);
  });

  it("detects a row identity change", () => {
    const left = { rows: [{ id: "r1", values: { column_14: "A" } }] };
    const right = { rows: [{ id: "r2", values: { column_14: "A" } }] };
    expect(payloadChanged(left, right)).toBe(true);
  });
});
