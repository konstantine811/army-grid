import { describe, expect, it } from "vitest";
import type { BackendPersonnelOverviewRow } from "../../api";
import {
  appendBchsMorningStatusChangeSection,
  buildBchsMorningSections,
} from "./overviewRotaBchsMorningExport";
import {
  buildBchsMorningDailySnapshot,
  buildBchsMorningStatusChangeRows,
  buildBchsMorningStatusChangeSection,
  describeMorningPersonChanges,
  describeMorningStatusChange,
} from "./overviewRotaBchsMorningStatusChanges";
import type { BchsMorningDailySnapshot } from "./overviewRotaBchsMorningSnapshot";

const row = (
  overrides: Partial<BackendPersonnelOverviewRow> = {},
): BackendPersonnelOverviewRow => ({
  id: "1",
  externalId: "1",
  validFrom: null,
  days: null,
  plannedReturn: null,
  place: "",
  updatedAt: "",
  name: "БІЛИЦЬКИЙ Сергій Володимирович",
  unit: "3 піхотна рота",
  rank: "солдат",
  status: "ON_DUTY",
  statusLabel: "В строю",
  staffStatus: "В строю",
  staffStatusLabel: "В строю",
  staffSheetColumns: {
    staff_3: "1 піхотний взвод",
    staff_5: "стрілець",
    staff_13: "солдат",
    staff_21: "В строю",
    staff_31: "ППД",
  },
  ...overrides,
});

describe("overviewRotaBchsMorningStatusChanges", () => {
  it("keeps only two confirmed daily transitions among medical wording changes and missing rows", () => {
    const person = (id: string, status: string, location: string) => row({
      externalId: id, name: `ПЕТРЕНКО Іван ${id === "1" ? "Іванович" : id === "2" ? "Петрович" : id === "3" ? "Миколайович" : "Сергійович"}`,
      staffSheetColumns: { ...row().staffSheetColumns, staff_21: status, staff_31: location },
    });
    const previous = {
      ...buildBchsMorningDailySnapshot([
        person("1", "В строю", "ППД"), person("2", "Відпустка", ""),
        person("3", "Лік.Пор", "по пораненню"), person("4", "В строю", "ППД"),
      ], "3 піхотна рота"),
      baselineKind: "staff" as const,
    };
    const changes = buildBchsMorningStatusChangeRows([
      person("1", "В строю", "На виконанні"), person("2", "Лікування", "Шпиталь"),
      person("3", "Лікування", "Шпиталь"),
    ], previous, "3 піхотна рота");
    expect(changes).toHaveLength(2);
    expect(changes.map((change) => change.changeNote)).toEqual(expect.arrayContaining([
      "Пішов на виконання", "Відпустка/Лік. відпустка → Шпиталь/Мед. рота",
    ]));
  });

  it("reports each AWOL transition when the unit total stays at one", () => {
    const second = row({ externalId: "2", name: "КУЗЬМЕНКО Руслан Олегович" });
    const previous = buildBchsMorningDailySnapshot([
      row({ status: "AWOL", staffSheetColumns: { ...row().staffSheetColumns, staff_21: "СЗЧ" } }),
      second,
    ], "3 піхотна рота");
    const changes = buildBchsMorningStatusChangeRows([
      row(),
      { ...second, status: "AWOL", staffSheetColumns: { ...second.staffSheetColumns, staff_21: "СЗЧ" } },
      row({ externalId: "3", name: "ІВАНЕНКО Іван Іванович", unit: "2 піхотна рота", status: "AWOL" }),
    ], previous, "3 піхотна рота");
    expect(changes).toHaveLength(2);
    expect(changes.find((change) => change.name === second.name)).toMatchObject({
      previousState: "В строю · ППД", newState: "СЗЧ · ППД",
    });
    expect(changes.find((change) => change.name === row().name)?.changeNote).toContain("Повернувся після СЗЧ");
  });

  it("displays AWOL from fresh roster data even when overview fields still say in service", () => {
    const previous = buildBchsMorningDailySnapshot([row()], "3 піхотна рота");
    const changes = buildBchsMorningStatusChangeRows([row()], previous, "3 піхотна рота", [
      { externalId: "1", column_21: "СЗЧ", column_31: "ППД" },
    ]);
    expect(changes).toHaveLength(1);
    expect(changes[0]?.newState).toBe("СЗЧ · ППД");
    expect(changes[0]?.changeNote).toContain("В строю → СЗЧ");
  });

  it.each([
    ["ШУРЧКОВ Микола Вікторович", "СІДОЙ"],
    ["ЩЕЛКУНОВ Денис Дмитрович", "ГЛУХОЙ"],
  ])("shows mission to in-service transition for %s even with a generic staff status", (name, callsign) => {
    const previous = buildBchsMorningDailySnapshot([row({ name })], "3 піхотна рота");
    previous.people[0] = {
      ...previous.people[0]!, callsign, bucket: "mission",
      status: "В строю", location: "ППД Петропалівка (13-й)",
    };
    const changes = buildBchsMorningStatusChangeRows([
      row({ name, staffSheetColumns: { ...row().staffSheetColumns, staff_31: "ППД Вишневе" } }),
    ], previous, "3 піхотна рота");
    expect(changes).toHaveLength(1);
    expect(changes[0]?.previousState).toBe("На виконанні · ППД Петропалівка (13-й)");
    expect(changes[0]?.newState).toBe("В строю · ППД Вишневе");
    expect(changes[0]?.changeNote).toContain("Прийшов з виконання");
    const section = buildBchsMorningStatusChangeSection([
      row({ name, staffSheetColumns: { ...row().staffSheetColumns, staff_31: "ППД Вишневе" } }),
    ], previous, "3 піхотна рота");
    expect(section.people[0]?.location).toBe("В строю · ППД Вишневе");
    expect(section.people[0]?.staffUnit).toContain("Прийшов з виконання");
    expect(changes[0]?.location).toBe("В строю · ППД Вишневе");
  });

  it("scopes both snapshots and changes to the selected unit, including long shared prefixes", () => {
    const unit = "1 піхотний взвод 3 піхотної роти";
    const other = row({ externalId: "2", name: "КУЗЬМЕНКО Руслан Олегович", unit: `${unit} резерв` });
    const own = row({ unit });
    const baseline = buildBchsMorningDailySnapshot([own, other], unit);
    expect(baseline.people.map((person) => person.key)).toEqual(["id:1"]);
    const mission = (person: BackendPersonnelOverviewRow) => ({
      ...person,
      staffSheetColumns: { ...person.staffSheetColumns, staff_31: "На виконанні" },
    });
    const changes = buildBchsMorningStatusChangeRows([mission(own), mission(other)], baseline, unit);
    expect(changes).toHaveLength(1);
    expect(changes[0]?.name).toBe(own.name);
  });

  it("does not infer a status transition from missing current rows", () => {
    const previous = buildBchsMorningDailySnapshot(
      [
        row({
          staffSheetColumns: {
            ...row().staffSheetColumns,
            staff_21: "Лік.Хвор",
            staff_31: "Шпиталь",
          },
        }),
      ],
      "3 піхотна рота",
    );
    const changes = buildBchsMorningStatusChangeRows([], previous, "3 піхотна рота");
    expect(changes).toEqual([]);
  });

  it("does not report position-only edits or new people as status transitions", () => {
    const previous = buildBchsMorningDailySnapshot([row()], "3 піхотна рота");
    expect(buildBchsMorningStatusChangeRows([
      row({ staffSheetColumns: { ...row().staffSheetColumns, staff_5: "командир" } }),
      row({ externalId: "9", name: "КУЗЬМЕНКО Руслан Олегович" }),
    ], previous, "3 піхотна рота")).toEqual([]);
  });

  it("matches names when imported row ids have been reassigned", () => {
    const previous = buildBchsMorningDailySnapshot([
      row({ staffSheetColumns: { ...row().staffSheetColumns, staff_31: "На виконанні" } }),
      row({ externalId: "2", name: "КУЗЬМЕНКО Руслан Олегович" }),
    ], "3 піхотна рота");
    const changes = buildBchsMorningStatusChangeRows([
      row({ externalId: "2" }),
      row({ externalId: "1", name: "КУЗЬМЕНКО Руслан Олегович" }),
    ], previous, "3 піхотна рота");
    expect(changes).toHaveLength(1);
    expect(changes[0]?.changeNote).toContain("Прийшов з виконання");
  });

  it("recognizes return from a hospital named in the location", () => {
    const previous = buildBchsMorningDailySnapshot([
      row({ staffSheetColumns: { ...row().staffSheetColumns, staff_21: "Лік.Хвор", staff_31: "Лікарня" } }),
    ], "3 піхотна рота");
    expect(previous.people[0]?.bucket).toBe("hospital");
    const changes = buildBchsMorningStatusChangeRows([row()], previous, "3 піхотна рота");
    expect(changes[0]?.changeNote).toContain("Прийшов з лікування");
  });

  it("describes common transitions in Ukrainian", () => {
    expect(describeMorningStatusChange("mission", "inService")).toBe(
      "Прийшов з виконання",
    );
    expect(describeMorningStatusChange("hospital", "inService")).toBe(
      "Прийшов з лікування",
    );
    expect(describeMorningStatusChange("inService", "mission")).toBe(
      "Пішов на виконання",
    );
    expect(describeMorningStatusChange("inService", "hospital")).toBe(
      "Пішов на лікування",
    );
  });

  it("matches uploaded baseline people by name when ids differ", () => {
    const previous: BchsMorningDailySnapshot = {
      date: "2026-09-12",
      unitLabel: "3 піхотна рота",
      savedAt: "2026-09-12T06:00:00.000Z",
      source: "manual",
      people: [
        {
          key: "name:білуцький сергій володимирович",
          rank: "солдат",
          name: "БІЛИЦЬКИЙ Сергій Володимирович",
          callsign: "БІЛ",
          position: "стрілець",
          status: "В строю",
          location: "На виконанні",
          bucket: "mission",
          stateLabel: "В строю · На виконанні",
          unit: "3 піхотна рота",
        },
      ],
    };
    const changes = buildBchsMorningStatusChangeRows(
      [
        row({
          unit: "3 піхотна рота",
          staffSheetColumns: {
            ...row().staffSheetColumns,
            staff_21: "В строю",
            staff_31: "ППД",
          },
        }),
      ],
      previous,
      "3 піхотна рота",
    );
    expect(changes).toHaveLength(1);
    expect(changes[0]?.changeNote).toContain("Прийшов з виконання");
  });

  it("detects people who returned from mission or medical", () => {
    const previous: BchsMorningDailySnapshot = {
      date: "2026-09-12",
      unitLabel: "3 піхотна рота",
      savedAt: "2026-09-12T06:00:00.000Z",
      people: [
        {
          key: "id:1",
          rank: "солдат",
          name: "БІЛИЦЬКИЙ Сергій Володимирович",
          callsign: "БІЛ",
          position: "стрілець",
          status: "В строю",
          location: "На виконанні",
          bucket: "mission",
          stateLabel: "В строю · На виконанні",
          unit: "3 піхотна рота",
        },
        {
          key: "id:2",
          rank: "солдат",
          name: "КУЗЬМЕНКО Руслан Олегович",
          callsign: "КУЗ",
          position: "стрілець",
          status: "Лік.Пор",
          location: "Шпиталь",
          bucket: "hospital",
          stateLabel: "Лік.Пор · Шпиталь",
          unit: "3 піхотна рота",
        },
      ],
    };
    const currentRows = [
      row({
        id: "1",
        externalId: "1",
        unit: "3 піхотна рота",
        staffSheetColumns: {
          ...row().staffSheetColumns,
          staff_21: "В строю",
          staff_31: "ППД",
        },
      }),
      row({
        id: "2",
        externalId: "2",
        name: "КУЗЬМЕНКО Руслан Олегович",
        unit: "3 піхотна рота",
        status: "ON_DUTY",
        staffSheetColumns: {
          ...row().staffSheetColumns,
          staff_21: "В строю",
          staff_31: "ППД",
        },
      }),
    ];
    const changes = buildBchsMorningStatusChangeRows(
      currentRows,
      previous,
      "3 піхотна рота",
    );
    expect(changes.map((item) => item.changeNote)).toEqual(
      expect.arrayContaining([
        expect.stringContaining("Прийшов з виконання"),
        expect.stringContaining("Прийшов з лікування"),
      ]),
    );
  });

  it("ignores EJOOS row.status drift when comparing against a staff baseline", () => {
    const previous = {
      ...buildBchsMorningDailySnapshot([row()], "3 піхотна рота"),
      baselineKind: "staff" as const,
    };
    expect(
      buildBchsMorningStatusChangeRows(
        [
          row({
            status: "MEDICAL",
            statusLabel: "Лік.Хвор",
            staffSheetColumns: {
              ...row().staffSheetColumns,
              staff_21: "В строю",
              staff_31: "ППД",
            },
          }),
        ],
        previous,
        "3 піхотна рота",
      ),
    ).toEqual([]);
  });

  it("ignores archival AWOL and detached staff rows when reporting departures", () => {
    const previous = {
      ...buildBchsMorningDailySnapshot(
        [
          row({
            externalId: "awol",
            name: "БЕВЗ Артем Сергійович",
            status: "AWOL",
            staffSheetColumns: {
              ...row().staffSheetColumns,
              staff_21: "СЗЧ",
              staff_31: "29.04.2026 на КСП",
            },
          }),
          row({
            externalId: "det",
            name: "АНТОНЮК Володимир Олександрович",
            staffSheetColumns: {
              ...row().staffSheetColumns,
              staff_21: "Відком. за межі ПБ",
              staff_31: "списання",
            },
          }),
          row({
            externalId: "hosp",
            name: "КРОШКА Іван Григорович",
            staffSheetColumns: {
              ...row().staffSheetColumns,
              staff_21: "Лік.Пор",
              staff_31: "Шпиталь",
            },
          }),
        ],
        "3 піхотна рота",
      ),
      baselineKind: "staff" as const,
    };
    const changes = buildBchsMorningStatusChangeRows(
      [row()],
      previous,
      "3 піхотна рота",
    );
    expect(changes).toEqual([]);
  });

  it("detects staff column edits for a staff baseline", () => {
    const previous = {
      ...buildBchsMorningDailySnapshot(
        [
          row({
            staffSheetColumns: {
              ...row().staffSheetColumns,
              staff_21: "Лік.Відп.",
              staff_31: "Відпустка лікувальна",
            },
          }),
        ],
        "3 піхотна рота",
      ),
      baselineKind: "staff" as const,
    };
    const changes = buildBchsMorningStatusChangeRows(
      [
        row({
          staffSheetColumns: {
            ...row().staffSheetColumns,
            staff_21: "Лік.Хвор",
            staff_31: "Шпиталь",
          },
        }),
      ],
      previous,
      "3 піхотна рота",
    );
    expect(changes).toHaveLength(1);
    expect(changes[0]?.name).toBe("БІЛИЦЬКИЙ Сергій Володимирович");
    expect(changes[0]?.changeNote).toMatch(/відпуст|лікуван|шпиталь/i);
  });

  it("ignores location or status text edits within the same bucket", () => {
    const previous: BchsMorningDailySnapshot = {
      date: "2026-09-12",
      unitLabel: "3 піхотна рота",
      savedAt: "2026-09-12T06:00:00.000Z",
      people: [
        {
          key: "id:1",
          rank: "солдат",
          name: "БІЛИЦЬКИЙ Сергій Володимирович",
          callsign: "БІЛ",
          position: "стрілець",
          status: "В строю",
          location: "ППД",
          bucket: "inService",
          stateLabel: "В строю · ППД",
          unit: "3 піхотна рота",
        },
      ],
    };
    const changes = buildBchsMorningStatusChangeRows(
      [
        row({
          unit: "3 піхотна рота",
          staffSheetColumns: {
            ...row().staffSheetColumns,
            staff_21: "В строю",
            staff_31: "Полігон",
          },
        }),
      ],
      previous,
      "3 піхотна рота",
    );
    expect(changes).toEqual([]);
    expect(
      describeMorningPersonChanges(previous.people[0]!, {
        ...previous.people[0]!,
        location: "Полігон",
      }),
    ).toBe("");
  });

  it("ignores people from other rotas in the baseline snapshot", () => {
    const previous: BchsMorningDailySnapshot = {
      date: "2026-09-12",
      unitLabel: "штатка",
      savedAt: "2026-09-12T06:00:00.000Z",
      people: [
        {
          key: "id:1",
          rank: "солдат",
          name: "БІЛИЦЬКИЙ Сергій Володимирович",
          callsign: "БІЛ",
          position: "стрілець",
          status: "В строю",
          location: "На виконанні",
          bucket: "mission",
          stateLabel: "В строю · На виконанні",
          unit: "3 піхотна рота",
          staffUnit: "1ПБ 3ПР",
        },
        {
          key: "id:9",
          rank: "солдат",
          name: "ЧУЖИЙ Іван Іванович",
          callsign: "ЧУЖ",
          position: "стрілець",
          status: "В строю",
          location: "ППД",
          bucket: "inService",
          stateLabel: "В строю",
          unit: "2 піхотна рота",
          staffUnit: "1ПБ 2ПР",
        },
      ],
    };
    const changes = buildBchsMorningStatusChangeRows(
      [
        row({
          staffSheetColumns: {
            ...row().staffSheetColumns,
            staff_21: "В строю",
            staff_31: "ППД",
          },
        }),
      ],
      previous,
      "3 піхотна рота",
    );
    expect(changes).toHaveLength(1);
    expect(changes[0]?.name).toBe("БІЛИЦЬКИЙ Сергій Володимирович");
    expect(changes.some((change) => change.name.includes("ЧУЖИЙ"))).toBe(false);
  });

  it("inserts the status change section after СЗЧ in extra block", () => {
    const previous = buildBchsMorningDailySnapshot(
      [
        row({
          staffSheetColumns: {
            ...row().staffSheetColumns,
            staff_21: "В строю",
            staff_31: "На виконанні",
          },
        }),
      ],
      "3 піхотна рота",
      null,
      new Date("2026-09-12"),
    );
    const currentRows = [
      row({
        unit: "3 піхотна рота",
        staffSheetColumns: {
          ...row().staffSheetColumns,
          staff_21: "В строю",
          staff_31: "ППД",
        },
      }),
    ];
    const base = buildBchsMorningSections(currentRows, "3 піхотна рота");
    const merged = appendBchsMorningStatusChangeSection(
      base,
      {
        id: "statusChanges",
        title: "Зміна статусу (відносно 12.09.2026)",
        side: "extra",
        people: buildBchsMorningStatusChangeRows(
          currentRows,
          previous,
          "3 піхотна рота",
        ).map((change) => ({
          ...change,
          status: change.previousState,
          location: change.newState,
          staffUnit: change.changeNote,
        })),
      },
    );
    const extraIds = merged
      .filter((section) => section.side === "extra")
      .map((section) => section.id);
    expect(extraIds).toEqual(["mission", "awol", "statusChanges"]);
  });
});
