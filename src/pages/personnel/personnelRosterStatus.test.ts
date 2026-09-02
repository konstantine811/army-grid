import { describe, expect, it } from "vitest";
import {
  classifyOverviewStatusFromRoster,
  resolvePersonRosterStatus,
} from "./personnelUtils";
import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";

const row = (values: Record<string, unknown>) => values as EjournalPreviewRow;

describe("resolvePersonRosterStatus", () => {
  it("reads raw Штатка «Статус», not «Статус БГ»", () => {
    expect(
      resolvePersonRosterStatus(
        row({
          статус_бг: "БГ",
          статус: "Лікування",
        }),
      ),
    ).toBe("Лікування");
  });

  it("reads roster__ status and column_21 fallback", () => {
    expect(
      resolvePersonRosterStatus(
        row({
          roster__статус: "Відпустка",
          roster__статус_бг: "БГ",
        }),
      ),
    ).toBe("Відпустка");
    expect(
      resolvePersonRosterStatus(
        row({
          column_21: "лікування після поранення",
          column_23: "БГ",
        }),
      ),
    ).toBe("лікування після поранення");
  });

  it("uses sheet column labels when the status column is not 21", () => {
    expect(
      resolvePersonRosterStatus(
        row({
          column_15: "Лікування",
          column_21: "",
        }),
        { column_15: "Статус" },
      ),
    ).toBe("Лікування");
  });
});

describe("classifyOverviewStatusFromRoster", () => {
  it("maps лікування and відпустка from Штатка text", () => {
    expect(classifyOverviewStatusFromRoster("Лікування").status).toBe("MEDICAL");
    expect(
      classifyOverviewStatusFromRoster("лікування після поранення").status,
    ).toBe("MEDICAL");
    expect(classifyOverviewStatusFromRoster("Відпустка").status).toBe("LEAVE");
    expect(classifyOverviewStatusFromRoster("в строю").status).toBe("ON_DUTY");
  });
});
