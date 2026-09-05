import { describe, expect, it } from "vitest";
import {
  buildSheetImpacts,
  buildTimesheetPreview,
  isInformationalOp,
  isWorkbookApplyOp,
  personChangesFromOps,
  writableOps,
  type PersonChange,
} from "./ejoosPersonDiff";
import type { EjoosSyncOp } from "./ejoosSyncPlan";

const op = (partial: Partial<EjoosSyncOp> & Pick<EjoosSyncOp, "kind">): EjoosSyncOp => ({
  id: "op-1",
  class: "needs_input",
  sheet: "тест",
  personId: "1",
  fullName: "ТЕСТ",
  positionIndex: "",
  rank: "",
  before: "",
  after: "",
  sourceRef: "",
  why: "",
  confidence: "manual",
  payload: {},
  checkedDefault: false,
  ...partial,
});

describe("buildTimesheetPreview", () => {
  it("shows dashes before an external arrival even without absence spans", () => {
    const preview = buildTimesheetPreview(
      [
        op({
          kind: "position_change",
          payload: {
            orderDate: "17.08.2026",
            timesheetActiveFrom: "17.08.2026",
            arrivedFrom: "_2 ШБ",
          },
        }),
      ],
      25,
    );

    expect(preview?.runs).toEqual([
      { from: 1, to: 16, mark: "-" },
      { from: 17, to: 25, mark: "+" },
    ]);
  });

  it("shows ШЕВЧУК as inactive on day 1, present from day 2, and ЗБ from day 8", () => {
    const preview = buildTimesheetPreview(
      [
        op({
          kind: "timesheet_day",
          payload: {
            timesheetActiveFrom: "02.08.2026",
            timesheetAbsenceSpans: "8-17:ЗБ",
          },
        }),
        op({
          kind: "move_to_disposition",
          payload: {
            keepOpenSzchTimesheet: "1",
            absenceCode: "ЗБ",
            orderDate: "18.08.2026",
            destination: "у розпорядження командира військової частини",
            orderNumber: "239",
          },
        }),
      ],
      25,
    );

    expect(preview?.runs).toEqual([
      { from: 1, to: 1, mark: "-" },
      { from: 2, to: 7, mark: "+" },
      { from: 8, to: 17, mark: "ЗБ" },
      { from: 18, to: 18, mark: "ПЕРЕВ" },
      { from: 19, to: 25, mark: "-" },
    ]);
  });
});

describe("isWorkbookApplyOp", () => {
  it("rejects standalone ПРИБУВ — there is no apply handler", () => {
    expect(isWorkbookApplyOp(op({ kind: "arrival" }))).toBe(false);
  });

  it("rejects generic other_manual", () => {
    expect(
      isWorkbookApplyOp(op({ kind: "other_manual", payload: { type: "ЗВІЛЬН" } })),
    ).toBe(false);
  });

  it("treats CANCEL_TRANSFER_BUT_NOT_IN_CURRENT_SH as review-only", () => {
    const cancel = op({
      kind: "other_manual",
      payload: {
        type: "TRANSFER_CANCELLED",
        reviewReason: "CANCEL_TRANSFER_BUT_NOT_IN_CURRENT_SH",
      },
    });
    expect(isInformationalOp(cancel)).toBe(true);
    expect(isWorkbookApplyOp(cancel)).toBe(false);
  });

  it("does not apply TRANSFER_CANCELLED when nothing is left to write", () => {
    const cancel = op({
      kind: "other_manual",
      class: "ready",
      payload: { type: "TRANSFER_CANCELLED" },
    });
    expect(isInformationalOp(cancel)).toBe(true);
    expect(isWorkbookApplyOp(cancel)).toBe(false);
  });

  it("applies TRANSFER_CANCELLED when timesheet or excluded still needs a write", () => {
    const withExcluded = op({
      kind: "other_manual",
      class: "ready",
      payload: { type: "TRANSFER_CANCELLED", excludedExcelRow: "6" },
    });
    const withTimesheet = op({
      kind: "other_manual",
      class: "ready",
      payload: { type: "TRANSFER_CANCELLED", restoreTimesheet: "1" },
    });
    expect(isWorkbookApplyOp(withExcluded)).toBe(true);
    expect(isWorkbookApplyOp(withTimesheet)).toBe(true);
  });

  it("skips disposition-after-position note so the ПОСАДА step can apply", () => {
    const note = op({
      kind: "other_manual",
      payload: {
        type: "РОЗПОРЯДЖ",
        chainWaiting: "1",
        awaitFromIndex: "2103791",
        awaitToIndex: "2103179",
      },
    });
    const place = op({
      kind: "position_change",
      class: "ready",
    });
    expect(isInformationalOp(note)).toBe(true);
    expect(isWorkbookApplyOp(note)).toBe(false);
    expect(isWorkbookApplyOp(place)).toBe(true);
    expect(writableOps([note, place]).map((item) => item.kind)).toEqual([
      "position_change",
    ]);
  });

  it("does not treat unclear transfer as an informational skip", () => {
    expect(
      isInformationalOp(
        op({
          kind: "other_manual",
          payload: { type: "TRANSFER_SCOPE_UNCLEAR", transferScope: "unclear" },
        }),
      ),
    ).toBe(false);
  });

  it("allows exclude_transfer", () => {
    expect(
      isWorkbookApplyOp(op({ kind: "exclude_transfer", class: "ready" })),
    ).toBe(true);
  });
});

describe("personHasWorkbookApplyOps typing smoke", () => {
  it("arrival-only person has no workbook apply", () => {
    const person = {
      ops: [op({ kind: "arrival" })],
    } as Pick<PersonChange, "ops">;
    expect(person.ops.some(isWorkbookApplyOp)).toBe(false);
  });
});

describe("БЕЗВІСТИ → РОЗПОРЯДЖ person card", () => {
  it("keeps OOS and open БЕЗВІСТИ, paints timesheet with ЗБ without вибув", () => {
    const people = personChangesFromOps(
      [
        op({
          kind: "move_to_disposition",
          class: "ready",
          personId: "21810",
          fullName: "АЛЬОХІН Віктор Олександрович",
          positionIndex: "2103455",
          payload: {
            previousIndex: "2103455",
            destination:
              ", який знаходиться у розпорядженні командира військової частини А4862",
            orderNumber: "241",
            orderDate: "12.08.2026",
            remainsInOos: "true",
            timesheetFound: "true",
            keepOpenSzchTimesheet: "1",
            absenceCode: "ЗБ",
            absenceType: "БЕЗВІСТИ",
          },
        }),
      ],
      25,
    );
    expect(people[0].ejoosWillDo.join("\n")).toMatch(/01–зріз ЗБ/i);
    expect(people[0].ejoosWillDo.join("\n")).toMatch(/БЕЗВІСТИ лишається відкритим/i);
    expect(people[0].ejoosWillDo.join("\n")).toMatch(/Виключені: без змін/);
    expect(
      people[0].sheetImpacts.find((item) => item.sheetKey === "timesheet")
        ?.detail,
    ).toMatch(/12\.08\.2026 записати «вибув у розпорядження/i);
    expect(people[0].ops[0]?.payload.keepOpenSzchTimesheet).toBe("1");
    expect(people[0].ops[0]?.payload.absenceCode).toBe("ЗБ");
    expect(people[0].timesheetPreview?.runs).toEqual([
      { from: 1, to: 11, mark: "ЗБ" },
      { from: 12, to: 12, mark: "ПЕРЕВ" },
      { from: 13, to: 25, mark: "-" },
    ]);
    expect(people[0].timesheetPreview?.departPhrase).toBe(
      "вибув у розпорядження командира військової частини А4862 наказ №241 від 12.08.2026",
    );
  });
});

describe("СЗЧ → РОЗПОРЯДЖ person card", () => {
  it("keeps OOS and open СЗЧ, paints timesheet without вибув", () => {
    const people = personChangesFromOps(
      [
        op({
          kind: "move_to_disposition",
          class: "ready",
          personId: "11040",
          fullName: "ДІДЕНКО Ілля Андрійович",
          positionIndex: "2103520",
          payload: {
            previousIndex: "2103520",
            destination:
              ", який знаходиться у розпорядженні командира військової частини А4862",
            orderNumber: "706-РС",
            orderDate: "12.08.2026",
            remainsInOos: "true",
            timesheetFound: "true",
            keepOpenSzchTimesheet: "1",
            absenceCode: "СЗЧ",
            absenceType: "СЗЧ",
          },
        }),
      ],
      25,
    );
    expect(people[0].ejoosWillDo.join("\n")).toMatch(/блок «ВИБУВ У РОЗПОРЯДЖЕННЯ…» · 01–зріз СЗЧ/i);
    expect(people[0].ejoosWillDo.join("\n")).toMatch(/Виключені: без змін/);
    expect(
      people[0].sheetImpacts.find((item) => item.sheetKey === "timesheet")
        ?.detail,
    ).toMatch(/наказ №706-РС від 12\.08\.2026/i);
    expect(people[0].timesheetPreview?.runs).toEqual([
      { from: 1, to: 11, mark: "СЗЧ" },
      { from: 12, to: 12, mark: "ПЕРЕВ" },
      { from: 13, to: 25, mark: "-" },
    ]);
  });
});

describe("duplicate timesheet cards", () => {
  it("does not dump other people's dupes onto one shared-ID card", () => {
    const people = personChangesFromOps(
      [
        op({
          kind: "timesheet_day",
          class: "ready",
          personId: "22895",
          fullName: "ЛІСОВСЬКИЙ Микола Сергійович",
          positionIndex: "2103607",
          payload: {
            type: "DUPLICATE_TAB_ROW",
            clearStalePerson: "1",
            excelRow: "681",
            keepTimesheetExcelRow: "264",
          },
        }),
        op({
          kind: "timesheet_day",
          class: "ready",
          personId: "22895",
          fullName: "ЛЕБЕДЕВ Дмитро Сергійович",
          positionIndex: "379",
          payload: {
            type: "DUPLICATE_TAB_ROW",
            clearStalePerson: "1",
            excelRow: "1173",
            keepTimesheetExcelRow: "1172",
          },
        }),
      ],
      25,
    );
    expect(people).toHaveLength(2);
    const lisovskyi = people.find((item) => /лісовськ/i.test(item.fullName));
    expect(lisovskyi?.ops).toHaveLength(1);
    expect(
      lisovskyi?.sheetImpacts.find((item) => item.sheetKey === "timesheet")
        ?.detail,
    ).toMatch(/Прибрати дубль R681/);
    expect(
      lisovskyi?.sheetImpacts.find((item) => item.sheetKey === "timesheet")
        ?.detail,
    ).not.toMatch(/1173/);
  });
});

describe("internal position hop does not touch Виключені", () => {
  it("skips an Excluded append when closeOldPosition is an in-battalion hop", () => {
    const impacts = buildSheetImpacts([
      op({
        kind: "position_change",
        class: "ready",
        fullName: "ЛІСЮТІН Микола Олександрович",
        personId: "19374",
        payload: {
          closeOldPosition: "1",
          internalStaffHop: "1",
          previousIndex: "2103535",
          nextIndex: "2103207",
        },
      }),
    ]);
    expect(impacts.find((item) => item.sheetKey === "excluded")?.effect).toBe(
      "untouched",
    );
  });
});
