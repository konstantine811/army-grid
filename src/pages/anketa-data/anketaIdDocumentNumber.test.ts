import { describe, expect, it } from "vitest";
import { createEmptyAnketaRow } from "./anketaSheet";
import {
  padAnketaIdCardDocumentNumber,
  padAnketaIdDocumentNumbersInRows,
} from "./anketaIdDocumentNumber";

describe("padAnketaIdCardDocumentNumber", () => {
  it("pads purely numeric ID-card numbers to 9 digits", () => {
    expect(padAnketaIdCardDocumentNumber("3266545")).toBe("003266545");
    expect(padAnketaIdCardDocumentNumber("7352002")).toBe("007352002");
    expect(padAnketaIdCardDocumentNumber("7874995")).toBe("007874995");
    expect(padAnketaIdCardDocumentNumber("011481028")).toBe("011481028");
    expect(padAnketaIdCardDocumentNumber("001991068")).toBe("001991068");
  });

  it("does not change passport numbers that contain letters", () => {
    expect(padAnketaIdCardDocumentNumber("FZ303340")).toBe("FZ303340");
    expect(padAnketaIdCardDocumentNumber("MP 143963")).toBe("MP 143963");
    expect(padAnketaIdCardDocumentNumber("HB 762078")).toBe("HB 762078");
    expect(padAnketaIdCardDocumentNumber("CM 875104")).toBe("CM 875104");
    expect(padAnketaIdCardDocumentNumber("АА 123456")).toBe("АА 123456");
  });

  it("leaves empty and non-numeric markers unchanged", () => {
    expect(padAnketaIdCardDocumentNumber("")).toBe("");
    expect(padAnketaIdCardDocumentNumber("дані відсутні")).toBe("дані відсутні");
    expect(padAnketaIdCardDocumentNumber("1234567890")).toBe("1234567890");
  });
});

describe("padAnketaIdDocumentNumbersInRows", () => {
  it("pads only the document-number column", () => {
    const row = createEmptyAnketaRow(2);
    row.idDocumentNumber = "3266545";
    row.militaryId = "АА 987654";
    const [next] = padAnketaIdDocumentNumbersInRows([row]);
    expect(next?.idDocumentNumber).toBe("003266545");
    expect(next?.militaryId).toBe("АА 987654");
  });
});
