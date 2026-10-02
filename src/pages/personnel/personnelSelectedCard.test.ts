import { describe, expect, it } from "vitest";
import { collectOrphanAttachmentIds } from "./healOrphanPersonnelAttachments";
import { rosterImportStatusMessage } from "./personnelImportStatus";
import {
  mergedDocumentPhones,
  personEditDraft,
  shouldCopyFoundQuestionnaire,
} from "./personnelSelectedCard";

describe("selected card helpers", () => {
  it("copies the visible field text into the draft", () => {
    const draft = personEditDraft(
      [
        {
          key: "rank",
          label: "Звання",
          parts: ["звання"],
          section: "identity",
        },
      ],
      { rank: "солдат" } as never,
    );
    expect(draft.rank).toBe("солдат");
  });

  it("copies a questionnaire only onto a different non-callsign id", () => {
    expect(shouldCopyFoundQuestionnaire(true, true, "pid_1", "4246")).toBe(true);
    expect(shouldCopyFoundQuestionnaire(true, true, "pid_1", "p:федерко:c:музикант")).toBe(
      false,
    );
    expect(shouldCopyFoundQuestionnaire(false, true, "pid_1", "4246")).toBe(false);
  });

  it("keeps the phone list when the document adds nothing new", () => {
    expect(mergedDocumentPhones(["0631111111"], ["063 111 11 11"])).toBeNull();
    expect(mergedDocumentPhones(["0631111111"], ["0672222222"])).toEqual([
      "0631111111",
      "0672222222",
    ]);
  });
});

describe("orphan attachment ids", () => {
  it("keeps ids that are not on the current list", () => {
    const orphans = collectOrphanAttachmentIds({
      currentIds: new Set(["4246"]),
      photos: [{ personExternalId: "old-photo" }],
      questionnaires: [{ personExternalId: "4246" }],
      storedPhones: { "old-phone": ["0631111111"], "4246": ["0672222222"] },
    });
    expect([...orphans].sort()).toEqual(["old-phone", "old-photo"]);
  });
});

describe("roster import status", () => {
  it("says when Цапенко is absent", () => {
    expect(
      rosterImportStatusMessage({
        rows: [],
        personCount: 662,
        personCountInRoster: 662,
        personCountInArchive: 0,
      }),
    ).toContain("у файлі не знайдено");
  });
});
