import { describe, expect, it } from "vitest";
import { personNameFromSyntheticDocumentId } from "./documentPersonIdentity";

describe("personNameFromSyntheticDocumentId", () => {
  it("restores a display name from a legacy person identity", () => {
    expect(
      personNameFromSyntheticDocumentId(
        "p:давиденко олександр володимирович:1985-03-08",
      ),
    ).toBe("ДАВИДЕНКО Олександр Володимирович");
  });

  it("does not treat a numeric database ID as a name", () => {
    expect(personNameFromSyntheticDocumentId("10034")).toBe("");
  });
});
