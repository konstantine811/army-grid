import { describe, expect, it } from "vitest";
import { FIGHTER_STATUS_FALLBACK_HEADERS } from "../excel-fill/staffSheet";
import { MORNING_GENERAL_LIST_COLUMN_LABELS } from "../personnel/personnelUtils";
import { OVERVIEW_STAFF_COLUMN_HEADERS } from "./overviewStaffColumns";

describe("OVERVIEW_STAFF_COLUMN_HEADERS", () => {
  it("matches «1.ОС Загальний список» column labels", () => {
    expect(OVERVIEW_STAFF_COLUMN_HEADERS.name).toBe(
      MORNING_GENERAL_LIST_COLUMN_LABELS[14],
    );
    expect(OVERVIEW_STAFF_COLUMN_HEADERS.unit).toBe(
      MORNING_GENERAL_LIST_COLUMN_LABELS[2],
    );
    expect(OVERVIEW_STAFF_COLUMN_HEADERS.rank).toBe(
      MORNING_GENERAL_LIST_COLUMN_LABELS[13],
    );
    expect(OVERVIEW_STAFF_COLUMN_HEADERS.positionTitle).toBe(
      MORNING_GENERAL_LIST_COLUMN_LABELS[7],
    );
    expect(OVERVIEW_STAFF_COLUMN_HEADERS.status).toBe(
      MORNING_GENERAL_LIST_COLUMN_LABELS[21],
    );
    expect(OVERVIEW_STAFF_COLUMN_HEADERS.questionnaire).toBe(
      MORNING_GENERAL_LIST_COLUMN_LABELS[10],
    );
    expect(OVERVIEW_STAFF_COLUMN_HEADERS.fighterDirection).toBe(
      MORNING_GENERAL_LIST_COLUMN_LABELS[33],
    );
  });

  it("matches «Статус бійців» fallback headers", () => {
    expect(OVERVIEW_STAFF_COLUMN_HEADERS.fighterExitDate).toBe(
      FIGHTER_STATUS_FALLBACK_HEADERS[6],
    );
    expect(OVERVIEW_STAFF_COLUMN_HEADERS.fighterReturnDate).toBe(
      FIGHTER_STATUS_FALLBACK_HEADERS[7],
    );
    expect(OVERVIEW_STAFF_COLUMN_HEADERS.fighterTotalDays).toBe(
      FIGHTER_STATUS_FALLBACK_HEADERS[8],
    );
    expect(OVERVIEW_STAFF_COLUMN_HEADERS.fighterStatus).toBe(
      "Статус 200/300/500",
    );
  });
});
