import { describe, expect, it } from "vitest";
import {
  bchsMorningManualBaselineCacheKey,
  extractBchsMorningRotaNumber,
  filterSnapshotPeopleForUnit,
  snapshotPersonMatchesUnit,
} from "./overviewRotaBchsMorningSnapshot";

describe("overviewRotaBchsMorningSnapshot", () => {
  it("does not share an uploaded baseline between units with the same first number", () => {
    expect(bchsMorningManualBaselineCacheKey("1 піхотний взвод"))
      .not.toBe(bchsMorningManualBaselineCacheKey("1 піхотна рота"));
    expect(bchsMorningManualBaselineCacheKey("1 піхотний взвод 2 піхотної роти"))
      .not.toBe(bchsMorningManualBaselineCacheKey("1 піхотний взвод 3 піхотної роти"));
  });
  it("uses the same manual baseline key for equivalent rota labels", () => {
    expect(extractBchsMorningRotaNumber("3 піхотна рота")).toBe("3");
    expect(extractBchsMorningRotaNumber("3 рота")).toBe("3");
    expect(extractBchsMorningRotaNumber("3ПР")).toBe("3");
    expect(bchsMorningManualBaselineCacheKey("3 піхотна рота")).toBe(
      bchsMorningManualBaselineCacheKey("3 рота"),
    );
  });

  it("matches snapshot people by roster unit label", () => {
    expect(
      snapshotPersonMatchesUnit(
        {
          key: "1",
          rank: "",
          name: "A",
          callsign: "",
          position: "",
          status: "",
          location: "",
          bucket: "inService",
          stateLabel: "",
          unit: "3 піхотна рота",
        },
        "3 піхотна рота",
      ),
    ).toBe(true);
    expect(
      snapshotPersonMatchesUnit(
        {
          key: "2",
          rank: "",
          name: "B",
          callsign: "",
          position: "",
          status: "",
          location: "",
          bucket: "inService",
          stateLabel: "",
          unit: "2 піхотна рота",
        },
        "3 піхотна рота",
      ),
    ).toBe(false);
  });

  it("filters baseline people to the selected rota", () => {
    const filtered = filterSnapshotPeopleForUnit(
      [
        {
          key: "1",
          rank: "",
          name: "A",
          callsign: "",
          position: "",
          status: "",
          location: "",
          bucket: "inService",
          stateLabel: "",
          unit: "3 піхотна рота",
        },
        {
          key: "2",
          rank: "",
          name: "B",
          callsign: "",
          position: "",
          status: "",
          location: "",
          bucket: "inService",
          stateLabel: "",
          unit: "2 піхотна рота",
        },
      ],
      "3 піхотна рота",
    );
    expect(filtered.map((person) => person.key)).toEqual(["1"]);
  });
});
