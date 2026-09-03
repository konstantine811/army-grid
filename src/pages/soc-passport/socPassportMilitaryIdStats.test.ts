import { describe, expect, it } from "vitest";
import { createEmptyAnketaRow } from "../anketa-data/anketaSheet";
import {
  buildAnketaStaffStats,
  classifyAnketaDocument,
  classifyAnketaMilitaryId,
} from "./socPassportMilitaryIdStats";

describe("classifyAnketaMilitaryId", () => {
  it("detects valid military id", () => {
    expect(classifyAnketaMilitaryId("АГ 1234567").status).toBe("has_value");
  });

  it("marks absent preset", () => {
    expect(classifyAnketaMilitaryId("відсутній").status).toBe("absent");
  });

  it("treats placeholder as empty", () => {
    expect(classifyAnketaMilitaryId("забув").status).toBe("empty");
  });
});

describe("classifyAnketaDocument", () => {
  it("detects document number", () => {
    expect(classifyAnketaDocument("АА 123456", "").status).toBe("has_value");
  });

  it("marks empty when both fields blank", () => {
    expect(classifyAnketaDocument("", "").status).toBe("empty");
  });

  it("marks absent preset", () => {
    expect(classifyAnketaDocument("дані відсутні", "").status).toBe("absent");
  });
});

describe("buildAnketaStaffStats", () => {
  it("summarizes military id and documents from anketa rows", () => {
    const anketaRows = [
      {
        ...createEmptyAnketaRow(2),
        fullName: "ІВАНОВ Іван Іванович",
        militaryId: "АГ 1111111",
        idDocumentNumber: "АА 123456",
      },
      {
        ...createEmptyAnketaRow(3),
        fullName: "ПЕТРОВ Петро Петрович",
        militaryId: "відсутній",
        idDocumentNumber: "",
      },
    ];

    const stats = buildAnketaStaffStats({
      morning: { fileName: "morning.xlsx", sheets: [] },
      anketaRows,
      year: 2026,
    });

    expect(stats.militaryId.anketa.total).toBe(2);
    expect(stats.militaryId.anketa.has).toBe(1);
    expect(stats.militaryId.anketa.absent).toBe(1);
    expect(stats.documents.anketa.has).toBe(1);
    expect(stats.documents.anketa.empty).toBe(1);
    expect(stats.militaryId.staff.total).toBe(0);
  });
});
