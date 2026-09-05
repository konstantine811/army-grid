import { describe, expect, it, vi } from "vitest";
import type { DiskQuestionnaireSearchResult } from "../../api";
import {
  searchQuestionnairesOnDiskInBatches,
  type QuestionnaireDiskSearchInput,
} from "./questionnaireDiskSearchBatch";

describe("searchQuestionnairesOnDiskInBatches", () => {
  it("splits a large search and refreshes the disk index only once", async () => {
    const people = Array.from({ length: 145 }, (_, index) => ({
      externalId: `person-${index}`,
      fullName: `ПРІЗВИЩЕ Ім'я ${index}`,
    }));
    const search = vi.fn(async (payload: {
      people: QuestionnaireDiskSearchInput[];
      refreshIndex?: boolean;
    }) => {
      return {
        root: "/disk",
        scannedFiles: 500,
        matchedPeople: 0,
        people: payload.people.map((person) => ({
          ...person,
          callSign: "",
          missingQuestionnaire: true,
          missingPhoto: false,
          matches: [],
        })),
      } satisfies DiskQuestionnaireSearchResult;
    });
    const progress: number[] = [];

    const result = await searchQuestionnairesOnDiskInBatches({
      people,
      search,
      batchSize: 60,
      pauseMs: 0,
      refreshIndex: true,
      onProgress: (done) => progress.push(done),
    });

    expect(search).toHaveBeenCalledTimes(3);
    expect(search.mock.calls.map(([payload]) => payload.people.length)).toEqual([
      60, 60, 25,
    ]);
    expect(search.mock.calls.map(([payload]) => payload.refreshIndex)).toEqual([
      true, false, false,
    ]);
    expect(progress).toEqual([60, 120, 145]);
    expect(result.people).toHaveLength(145);
  });

  it("stops before the next batch when cancelled", async () => {
    let cancelled = false;
    const search = vi.fn(async (payload: {
      people: QuestionnaireDiskSearchInput[];
      refreshIndex?: boolean;
    }) => {
      cancelled = true;
      return {
        root: "/disk",
        scannedFiles: 1,
        matchedPeople: 0,
        people: payload.people.map((person) => ({
          ...person,
          callSign: "",
          missingQuestionnaire: true,
          missingPhoto: false,
          matches: [],
        })),
      } satisfies DiskQuestionnaireSearchResult;
    });

    const result = await searchQuestionnairesOnDiskInBatches({
      people: Array.from({ length: 100 }, (_, index) => ({
        externalId: String(index),
        fullName: `ЛЮДИНА ${index}`,
      })),
      search,
      batchSize: 20,
      pauseMs: 0,
      isCancelled: () => cancelled,
    });

    expect(search).toHaveBeenCalledTimes(1);
    expect(result.people).toHaveLength(20);
  });
});
