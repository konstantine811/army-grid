import { describe, expect, it } from "vitest";
import type { BackendPersonnelOverviewRow } from "../../api";
import {
  buildOverviewCriticalFromRows,
  buildOverviewTodayStats,
  filterOverviewCriticalForRows,
} from "./overviewSideStats";

const row = (
  overrides: Partial<BackendPersonnelOverviewRow> = {},
): BackendPersonnelOverviewRow => ({
  id: "r1",
  externalId: "e1",
  name: "ЧЕНДЕЙ Едуард Михайлович",
  rank: "солдат",
  unit: "1 рота",
  status: "AWOL",
  statusLabel: "СЗЧ",
  validFrom: null,
  days: 109,
  plannedReturn: null,
  place: "",
  updatedAt: "",
  ...overrides,
});

describe("overview side stats follow selected rows", () => {
  it("keeps critical items only for people in the current selection", () => {
    const scoped = filterOverviewCriticalForRows(
      [
        {
          id: "r1",
          severity: "danger",
          text: "ЧЕНДЕЙ Едуард Михайлович: СЗЧ · 109 дн.",
          status: "AWOL",
          days: 109,
          daysToReturn: null,
        },
        {
          id: "other",
          severity: "warning",
          text: "Інший Іван: Лікування · 80 дн.",
          status: "MEDICAL",
          days: 80,
          daysToReturn: null,
        },
      ],
      [row()],
    );

    expect(scoped).toHaveLength(1);
    expect(scoped[0].id).toBe("r1");
  });

  it("builds critical terms from Штатка statuses when backend list is empty", () => {
    const items = buildOverviewCriticalFromRows([
      row({ status: "ON_DUTY", statusLabel: "В строю", days: null }),
      row({
        id: "m1",
        name: "ШАГОВКА Олександр",
        status: "MEDICAL",
        statusLabel: "Лікування",
        days: 137,
        plannedReturn: "2020-01-01",
      }),
    ]);

    expect(items).toHaveLength(1);
    expect(items[0].text).toContain("Лікування");
    expect(items[0].text).toContain("строк повернення минув");
  });

  it("counts today changes only among selected rows", () => {
    const now = new Date("2026-09-02T12:00:00");
    const stats = buildOverviewTodayStats(
      [
        row({
          id: "t1",
          status: "LEAVE",
          statusLabel: "Відпустка",
          updatedAt: "2026-09-02T08:00:00",
        }),
        row({
          id: "t2",
          status: "MEDICAL",
          statusLabel: "Лікування",
          updatedAt: "2026-08-01T08:00:00",
        }),
      ],
      now,
    );

    expect(stats.todayUpdates).toBe(1);
    expect(stats.todayChanges.leave).toBe(1);
    expect(stats.todayChanges.medical).toBe(0);
    expect(stats.todayChanges.total).toBe(1);
  });
});
