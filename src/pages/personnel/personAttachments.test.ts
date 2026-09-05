import { describe, expect, it, vi } from "vitest";
import { api, type BackendPersonDocument } from "../../api";
import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import {
  buildOrphanAttachmentMigrationPairs,
  buildQuestionnairePresenceMap,
  collectPersonAttachmentLookupIds,
  loadPersonDocumentsForRow,
  matchOrphanIdToPersonnelRow,
  parseOrphanAttachmentIdentityId,
  personNameMatchesOrphanNameKey,
  questionnaireFileMatchesPerson,
} from "./personAttachments";
import {
  buildPersonIdentityFingerprint,
  resolvePersonIdentityKey,
} from "./personnelUtils";

const document = (
  id: string,
  personExternalId: string,
  personName: string,
): BackendPersonDocument =>
  ({
    id,
    personExternalId,
    personName,
    type: "form6Report",
    title: "Форма 6",
    status: "draft",
    updatedAt: "2026-09-05T10:00:00.000Z",
    createdAt: "2026-09-05T10:00:00.000Z",
  }) as BackendPersonDocument;

const personRow = (
  name: string,
  extra: Record<string, unknown> = {},
): EjournalPreviewRow =>
  ({
    __dbRowId: String(extra.__dbRowId ?? `row:${name}`),
    прізвище: name,
    дата_народження: extra.дата_народження ?? extra.birthDate ?? "",
    id: extra.id ?? "",
    ...extra,
  }) as EjournalPreviewRow;

describe("collectPersonAttachmentLookupIds", () => {
  it("keeps the old fingerprint after anketa merge adds ID and по батькові", () => {
    const oldKey = buildPersonIdentityFingerprint(
      "Іванов Іван",
      "01.01.1990",
    );
    const afterMerge = personRow("Іванов Іван Петрович", {
      birthDate: "01.01.1990",
      id: "2163435",
    });

    expect(oldKey).toMatch(/^p:/);
    expect(resolvePersonIdentityKey(afterMerge)).toBe("2163435");
    expect(collectPersonAttachmentLookupIds(afterMerge)).toContain(oldKey);
  });

  it("does not share name-only keys between two Шевченки with different birth dates", () => {
    const first = personRow("ШЕВЧЕНКО Олександр Володимирович", {
      birthDate: "07.09.1985",
      id: "2103004",
      __dbRowId: "row:shevchenko-1985",
    });
    const second = personRow("ШЕВЧЕНКО Олександр Володимирович", {
      birthDate: "11.05.1981",
      id: "2103005",
      __dbRowId: "row:shevchenko-1981",
    });

    const firstIds = collectPersonAttachmentLookupIds(first);
    const secondIds = collectPersonAttachmentLookupIds(second);
    const shared = firstIds.filter((id) => secondIds.includes(id));

    expect(shared).toEqual([]);
    expect(firstIds).not.toContain(
      buildPersonIdentityFingerprint("ШЕВЧЕНКО Олександр Володимирович"),
    );
    expect(
      collectPersonAttachmentLookupIds(first, undefined, {
        includeLooseKeys: true,
      }),
    ).toContain(
      buildPersonIdentityFingerprint("ШЕВЧЕНКО Олександр Володимирович"),
    );
  });
});

describe("buildQuestionnairePresenceMap", () => {
  it("marks the current ID when the PDF still lives under the previous key", () => {
    const oldKey = buildPersonIdentityFingerprint(
      "Іванов Іван",
      "21.03.1977",
    );
    const row = personRow("Іванов Іван Петрович", {
      birthDate: "21.03.1977",
      id: "2163435",
    });

    const map = buildQuestionnairePresenceMap(
      [row],
      [{ personExternalId: oldKey }],
    );

    expect(map["2163435"]).toBe(true);
    expect(map[oldKey]).toBe(true);
  });
});

describe("loadPersonDocumentsForRow", () => {
  it("finds Davydenko documents stored under a legacy person ID", async () => {
    const row = personRow("ДАВИДЕНКО Олександр Володимирович", {
      id: "current-id",
      birthDate: "08.03.1985",
    });
    vi.spyOn(api, "listPersonDocuments").mockResolvedValue([
      document("current-doc", "current-id", "ДАВИДЕНКО Олександр Володимирович"),
    ]);
    vi.spyOn(api, "listAllPersonDocuments").mockResolvedValue([
      document(
        "legacy-doc",
        "p:давиденко олександр володимирович:1985-03-08",
        "ДАВИДЕНКО Олександр Володимирович",
      ),
    ]);

    const result = await loadPersonDocumentsForRow(row);

    expect(result.map((item) => item.id).sort()).toEqual([
      "current-doc",
      "legacy-doc",
    ]);
    vi.restoreAllMocks();
  });
});

describe("buildOrphanAttachmentMigrationPairs", () => {
  it("copies a manual questionnaire from the old fingerprint onto the new spreadsheet ID", () => {
    const oldKey = buildPersonIdentityFingerprint(
      "Шевченко Тарас",
      "09.03.1984",
    );
    const row = personRow("Шевченко Тарас Григорович", {
      birthDate: "09.03.1984",
      id: "100500",
    });

    expect(
      buildOrphanAttachmentMigrationPairs([row], new Set([oldKey])),
    ).toEqual([
      expect.objectContaining({
        fromExternalId: oldKey,
        toExternalId: "100500",
      }),
    ]);
  });

  it("does not attach a shared short-name orphan to two people", () => {
    const shortKey = buildPersonIdentityFingerprint("Іванов Іван");
    const first = personRow("Іванов Іван Петрович", {
      birthDate: "01.01.1990",
      id: "1",
    });
    const second = personRow("Іванов Іван Сергійович", {
      birthDate: "02.02.1992",
      id: "2",
    });

    const pairs = buildOrphanAttachmentMigrationPairs(
      [first, second],
      new Set([shortKey]),
    );

    expect(pairs.filter((pair) => pair.fromExternalId === shortKey)).toEqual([]);
  });

  it("pairs many orphan IDs without scanning the whole list for each one", () => {
    const suffixes = "абвгдежзиклмнпрстуфхцчшщюя";
    const rows = Array.from({ length: suffixes.length }, (_, index) =>
      personRow(`Тестовий${suffixes[index]} Іван Петрович`, {
        birthDate: "01.01.1990",
        id: String(3000 + index),
      }),
    );
    const orphans = new Set(
      Array.from({ length: 400 }, (_, index) => `orphan:${index}`),
    );
    const target = rows[7];
    const oldKey = buildPersonIdentityFingerprint(
      `Тестовий${suffixes[7]} Іван`,
      "01.01.1990",
    );
    orphans.add(oldKey);

    const started = Date.now();
    const pairs = buildOrphanAttachmentMigrationPairs(rows, orphans);
    expect(Date.now() - started).toBeLessThan(250);
    expect(pairs).toEqual([
      expect.objectContaining({
        fromExternalId: oldKey,
        toExternalId: resolvePersonIdentityKey(target),
      }),
    ]);
  });
});

describe("matchOrphanIdToPersonnelRow", () => {
  it("matches p:name:birth even when the card name grew a patronymic", () => {
    const oldKey = buildPersonIdentityFingerprint(
      "Коваль Іван",
      "01.01.1990",
    );
    const row = personRow("Коваль Іван Петрович", {
      birthDate: "01.01.1990",
      id: "77",
    });

    expect(matchOrphanIdToPersonnelRow(oldKey, [row])).toBe(row);
  });

  it("does not match another person with the same surname and first name", () => {
    const oldKey = buildPersonIdentityFingerprint(
      "Коваль Іван",
      "01.01.1990",
    );
    const target = personRow("Коваль Іван Петрович", {
      birthDate: "01.01.1990",
      id: "77",
    });
    const other = personRow("Коваль Іван Сергійович", {
      birthDate: "02.02.1992",
      id: "78",
    });

    expect(matchOrphanIdToPersonnelRow(oldKey, [target, other])).toBe(target);
  });
});

describe("parseOrphanAttachmentIdentityId", () => {
  it("parses fingerprint and name-birth keys", () => {
    expect(
      parseOrphanAttachmentIdentityId("p:коваль іван:1990-01-01"),
    ).toEqual({
      nameKey: "коваль іван",
      birthKey: "1990-01-01",
    });
    expect(
      parseOrphanAttachmentIdentityId("name-birth:коваль іван:1990-01-01"),
    ).toEqual({
      nameKey: "коваль іван",
      birthKey: "1990-01-01",
    });
  });
});

describe("personNameMatchesOrphanNameKey", () => {
  it("treats a shorter saved FIO as the same person", () => {
    expect(
      personNameMatchesOrphanNameKey(
        "Коваль Іван Петрович",
        "коваль іван",
      ),
    ).toBe(true);
    expect(
      personNameMatchesOrphanNameKey(
        "Коваль Іван Петрович",
        "коваль степан",
      ),
    ).toBe(false);
  });
});

describe("questionnaireFileMatchesPerson", () => {
  it("matches a shorter filename without по батькові", () => {
    expect(
      questionnaireFileMatchesPerson("Шевцов Дмитро.pdf", [
        "ШЕВЦОВ ДМИТРО СЕРГІЙОВИЧ",
      ]),
    ).toBe(true);
    expect(
      questionnaireFileMatchesPerson(
        "Анкета (PDF) КОКИЦЬ Сергій Анатолійович.pdf",
        ["КОКИЦЬ СЕРГІЙ АНАТОЛІЙОВИЧ"],
      ),
    ).toBe(true);
    expect(
      questionnaireFileMatchesPerson("Шевцов Степан.pdf", [
        "ШЕВЦОВ ДМИТРО СЕРГІЙОВИЧ",
      ]),
    ).toBe(false);
  });

  it("does not mark a shared FIO questionnaire for two namesakes", () => {
    const first = personRow("ШЕВЧЕНКО Олександр Володимирович", {
      birthDate: "07.09.1985",
      id: "2103004",
    });
    const second = personRow("ШЕВЧЕНКО Олександр Володимирович", {
      birthDate: "11.05.1981",
      id: "2103005",
    });
    const map = buildQuestionnairePresenceMap(
      [first, second],
      [
        {
          personExternalId: "old-moto-key",
          fileName: "ШЕВЧЕНКО Олександр Володимирович (мото).pdf",
        },
      ],
    );

    expect(map["2103004"]).toBeUndefined();
    expect(map["2103005"]).toBeUndefined();
  });

  it("reattaches Шевцов when the PDF is stored under an unknown ID but the file name matches", () => {
    const row = personRow("ШЕВЦОВ ДМИТРО СЕРГІЙОВИЧ", {
      birthDate: "27.02.1986",
      id: "2103825",
    });
    const map = buildQuestionnairePresenceMap(
      [row],
      [{ personExternalId: "old-manual-key", fileName: "Шевцов Дмитро.pdf" }],
    );
    expect(map["2103825"]).toBe(true);

    expect(
      buildOrphanAttachmentMigrationPairs(
        [row],
        new Set(["old-manual-key"]),
        [{ personExternalId: "old-manual-key", fileName: "Шевцов Дмитро.pdf" }],
      ),
    ).toEqual([
      expect.objectContaining({
        fromExternalId: "old-manual-key",
        toExternalId: "2103825",
      }),
    ]);
  });
});
