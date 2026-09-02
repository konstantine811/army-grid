import { describe, expect, it } from "vitest";
import type { BackendPersonnelOverviewRow } from "../../api";
import { buildPersonIdentityFingerprint } from "../personnel/personnelUtils";
import { buildOverviewPhotoMap, resolveOverviewPhoto } from "./overviewPhotos";

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
});
