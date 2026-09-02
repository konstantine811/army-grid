import { describe, expect, it } from "vitest";
import { rememberMountedPage } from "./pageKeepAlive";

describe("rememberMountedPage", () => {
  it("keeps the active page when evicting older ones", () => {
    const mounted = rememberMountedPage(
      new Set(["overview", "personnel", "ejournal", "excelFill", "bchs"]),
      "documents",
      4,
    );

    expect([...mounted]).toEqual(["ejournal", "excelFill", "bchs", "documents"]);
    expect(mounted.has("overview")).toBe(false);
    expect(mounted.has("documents")).toBe(true);
  });

  it("returns the same set when the page is already kept under the limit", () => {
    const current = new Set(["overview", "personnel"] as const);
    expect(rememberMountedPage(current, "overview", 4)).toBe(current);
  });
});
