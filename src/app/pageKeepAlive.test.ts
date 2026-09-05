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
    const current = new Set(["profile", "analytics"] as const);
    expect(rememberMountedPage(current, "analytics", 4)).toBe(current);
  });

  it("treats a revisited page as the most recently used", () => {
    const revisited = rememberMountedPage(
      new Set(["analytics", "profile", "workTasks"]),
      "analytics",
      3,
    );
    expect([...revisited]).toEqual(["profile", "workTasks", "analytics"]);

    const withNextPage = rememberMountedPage(revisited, "bchs", 3);
    expect([...withNextPage]).toEqual(["workTasks", "analytics", "bchs"]);
    expect(withNextPage.has("analytics")).toBe(true);
  });

  it("unmounts every previous page with the default limit", () => {
    const personnel = rememberMountedPage(new Set(), "personnel");
    const ejournal = rememberMountedPage(personnel, "ejournal");
    const excelFill = rememberMountedPage(ejournal, "excelFill");
    const profile = rememberMountedPage(excelFill, "profile");

    expect([...personnel]).toEqual(["personnel"]);
    expect([...ejournal]).toEqual(["ejournal"]);
    expect([...excelFill]).toEqual(["excelFill"]);
    expect([...profile]).toEqual(["profile"]);
  });
});
