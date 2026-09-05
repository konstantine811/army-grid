import { describe, expect, it } from "vitest";
import { DEFAULT_STATUS_RULES } from "./ejoosRules";
import { mapPbStatusToEjoosWithRules } from "./ejoosStatusMap";

describe("mapPbStatusToEjoosWithRules", () => {
  it("does not map «НЕ В СТРОЮ - СЗЧ» to on_duty", () => {
    const mapped = mapPbStatusToEjoosWithRules(
      "НЕ В СТРОЮ - СЗЧ",
      DEFAULT_STATUS_RULES,
    );
    expect(mapped.ruleId).toBe("awol");
    expect(mapped.timesheetCode).toBe("СЗЧ");
  });

  it("does not map «НЕ ПРИСУТНІЙ» to on_duty", () => {
    const mapped = mapPbStatusToEjoosWithRules(
      "НЕ ПРИСУТНІЙ",
      DEFAULT_STATUS_RULES,
    );
    expect(mapped.timesheetCode).not.toBe("+");
  });

  it("maps exact «В СТРОЮ» to +", () => {
    const mapped = mapPbStatusToEjoosWithRules("В СТРОЮ", DEFAULT_STATUS_RULES);
    expect(mapped.ruleId).toBe("on_duty");
    expect(mapped.timesheetCode).toBe("+");
  });

  it("maps ВЛК to a dedicated absence and timesheet code", () => {
    const mapped = mapPbStatusToEjoosWithRules("ВЛК", DEFAULT_STATUS_RULES);

    expect(mapped.ruleId).toBe("medical_board");
    expect(mapped.absenceGround).toBe("ВЛК");
    expect(mapped.timesheetCode).toBe("ВЛК");
  });

  it("keeps МЕД РОТА as лікування in the timesheet", () => {
    const mapped = mapPbStatusToEjoosWithRules(
      "МЕД РОТА",
      DEFAULT_STATUS_RULES,
    );

    expect(mapped.ruleId).toBe("med");
    expect(mapped.timesheetCode).toBe("лік");
  });
});
