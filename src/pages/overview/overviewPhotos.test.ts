import { afterEach, describe, expect, it, vi } from "vitest";
import { api, type BackendPersonnelOverviewRow } from "../../api";
import { buildPersonIdentityFingerprint } from "../personnel/personnelUtils";
import {
  buildOverviewPhotoMap,
  fillMissingOverviewPhotos,
  indexOverviewRosterRows,
  overviewPhotoPreviewUrl,
  resolveOverviewPhoto,
} from "./overviewPhotos";
import { clearAvailablePersonPhotoIdsCache } from "../personnel/personAttachments";

afterEach(() => {
  clearAvailablePersonPhotoIdsCache();
  vi.restoreAllMocks();
});

const overviewRow = (
  name: string,
  extra: Partial<BackendPersonnelOverviewRow> = {},
): BackendPersonnelOverviewRow => ({
  id: extra.id || `overview:${name}`,
  externalId: extra.externalId || `roster:${name}`,
  name,
  rank: "солдат",
  unit: "1 рота",
  status: "ON_DUTY",
  statusLabel: "На службі",
  validFrom: null,
  days: null,
  plannedReturn: null,
  place: "",
  updatedAt: "",
  ...extra,
});

describe("buildOverviewPhotoMap", () => {
  it("attaches a photo stored under a name fingerprint to the roster key", () => {
    const name = "ПОЛЬОВИЙ Олексій Геннадійович";
    const photoId = buildPersonIdentityFingerprint(name);
    const row = overviewRow(name, {
      externalId: "roster:польовий олексій геннадійович",
    });
    const { photos } = buildOverviewPhotoMap(
      [{ personExternalId: photoId, photoData: "data:image/jpeg;base64,abc" }],
      [row],
      [],
    );

    expect(resolveOverviewPhoto(row, photos)).toBe("data:image/jpeg;base64,abc");
  });

  it("builds a thumbnail file url rather than fetching image bytes", async () => {
    const row = overviewRow("ПОЛЬОВИЙ Олексій Геннадійович", {
      externalId: "person-1",
    });
    const urlSpy = vi
      .spyOn(api, "personPhotoFileUrl")
      .mockReturnValue("http://test/photo?thumbnail=1");
    vi.spyOn(api, "listPersonPhotos").mockResolvedValue([
      {
        personExternalId: "person-1",
        photoData: "",
        hasThumbnail: true,
      },
    ]);
    const thumbnail = vi.spyOn(api, "getPersonPhotoThumbnail");
    const fullPhoto = vi.spyOn(api, "getPersonPhoto");

    const photos = await fillMissingOverviewPhotos([row], [], {});

    expect(urlSpy).toHaveBeenCalledWith("person-1", { thumbnail: true });
    expect(thumbnail).not.toHaveBeenCalled();
    expect(fullPhoto).not.toHaveBeenCalled();
    expect(resolveOverviewPhoto(row, photos)).toBe(
      "http://test/photo?thumbnail=1",
    );
  });

  it("maps photo index entries without inline data to thumbnail urls", () => {
    vi.spyOn(api, "personPhotoFileUrl").mockReturnValue(
      "http://test/photo?thumbnail=1&access_token=abc",
    );
    const { photos } = buildOverviewPhotoMap(
      [{ personExternalId: "person-1", hasThumbnail: true }],
      [],
      [],
    );

    expect(photos["person-1"]).toBe(
      "http://test/photo?thumbnail=1&access_token=abc",
    );
  });

  it("does not build thumbnail urls for a person absent from the photo index", async () => {
    const row = overviewRow("БЕЗ ФОТО", { externalId: "missing-person" });
    vi.spyOn(api, "listPersonPhotos").mockResolvedValue([]);
    const urlSpy = vi.spyOn(api, "personPhotoFileUrl");
    const thumbnail = vi.spyOn(api, "getPersonPhotoThumbnail");

    await fillMissingOverviewPhotos([row], [], {});

    expect(urlSpy).not.toHaveBeenCalled();
    expect(thumbnail).not.toHaveBeenCalled();
  });

  it("does not assign a name fallback when the roster name is ambiguous", () => {
    const name = "ОДНАКОВИЙ Олексій Петрович";
    const { rosterByName } = indexOverviewRosterRows([
      { __dbRowId: "one", ПІБ: name },
      { __dbRowId: "two", ПІБ: name },
    ]);

    expect(rosterByName.size).toBe(0);
  });

  it("does not use a different person's longer fingerprint as a prefix match", () => {
    const row = overviewRow("ПОЛЬОВИЙ Олексій Геннадійович");
    const otherKey = `${buildPersonIdentityFingerprint(row.name)}:1990-01-01`;

    expect(
      resolveOverviewPhoto(row, {
        [otherKey]: "data:image/jpeg;base64,wrong",
      }),
    ).toBe("");
  });
});

describe("overviewPhotoPreviewUrl", () => {
  it("removes thumbnail flag for hover preview", () => {
    expect(
      overviewPhotoPreviewUrl(
        "http://test/photo?thumbnail=1&access_token=abc",
      ),
    ).toBe("http://test/photo?access_token=abc");
  });

  it("keeps data urls unchanged", () => {
    const dataUrl = "data:image/jpeg;base64,abc";
    expect(overviewPhotoPreviewUrl(dataUrl)).toBe(dataUrl);
  });
});

describe("overviewPhotoPreviewUrl", () => {
  it("removes thumbnail flag for hover preview", () => {
    expect(
      overviewPhotoPreviewUrl(
        "http://test/photo?thumbnail=1&access_token=abc",
      ),
    ).toBe("http://test/photo?access_token=abc");
  });

  it("keeps data urls unchanged", () => {
    const dataUrl = "data:image/jpeg;base64,abc";
    expect(overviewPhotoPreviewUrl(dataUrl)).toBe(dataUrl);
  });
});
