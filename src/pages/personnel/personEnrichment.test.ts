import { describe, expect, it } from "vitest";
import { buildEmptyPersonFieldUpdates } from "./personEnrichment";

describe("buildEmptyPersonFieldUpdates", () => {
  it("does not overwrite a filled personnel field by default", () => {
    expect(
      buildEmptyPersonFieldUpdates(
        { __dbRowId: "row-1", рнокпп_за_наявності: "втрачено" },
        { rnokpp: "3142223156" },
      ),
    ).toEqual({});
  });

  it("writes a valid manual RNOKPP back to the personnel source", () => {
    expect(
      buildEmptyPersonFieldUpdates(
        { __dbRowId: "row-1", рнокпп_за_наявності: "втрачено" },
        { rnokpp: "3142223156" },
        { overwriteFields: true },
      ),
    ).toEqual({ рнокпп_за_наявності: "3142223156" });
  });

  it("never writes an invalid RNOKPP into the shared personnel source", () => {
    expect(
      buildEmptyPersonFieldUpdates(
        { __dbRowId: "row-1", рнокпп_за_наявності: "втрачено" },
        { rnokpp: "31422" },
        { overwriteFields: true },
      ),
    ).toEqual({});
  });
});
