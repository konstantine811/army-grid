import { afterEach, describe, expect, it, vi } from "vitest";
import {
  clearPerformanceSnapshot,
  getPerformanceSnapshot,
  measuredFetch,
  sanitizePerformancePath,
} from "./performanceMonitor";

afterEach(() => {
  clearPerformanceSnapshot();
  vi.unstubAllGlobals();
});

describe("performance monitor", () => {
  it("removes identifiers and query values from recorded paths", () => {
    expect(
      sanitizePerformancePath(
        "https://example.test/api/person/12345678901234567890/file?thumbnail=1&token=secret",
      ),
    ).toBe("/api/person/:id/file?thumbnail&token");
  });

  it("records API duration and response size without response content", async () => {
    vi.stubGlobal(
      "fetch",
      vi.fn(async () =>
        new Response("{}", {
          status: 200,
          headers: { "content-length": "2048" },
        }),
      ),
    );

    await measuredFetch("https://example.test/api/overview?force=1");

    expect(getPerformanceSnapshot()).toEqual([
      expect.objectContaining({
        type: "api",
        method: "GET",
        path: "/api/overview?force",
        status: 200,
        responseBytes: 2048,
      }),
    ]);
  });
});
