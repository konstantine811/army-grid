import { afterEach, describe, expect, it, vi } from "vitest";
import {
  fetchWithCache,
  resetDataCacheMemory,
  writeDataCache,
} from "./idbDataCache";

afterEach(() => {
  resetDataCacheMemory();
});

describe("fetchWithCache shared memory", () => {
  it("deduplicates concurrent fetches for the same key", async () => {
    let calls = 0;
    const fetcher = vi.fn(async () => {
      calls += 1;
      await new Promise((resolve) => setTimeout(resolve, 20));
      return { n: calls };
    });

    const [left, right] = await Promise.all([
      fetchWithCache({ key: "test:dup", fetcher }),
      fetchWithCache({ key: "test:dup", fetcher }),
    ]);

    expect(fetcher).toHaveBeenCalledTimes(1);
    expect(left).toEqual({ n: 1 });
    expect(right).toEqual({ n: 1 });
  });

  it("skips the network while the shared copy is still fresh", async () => {
    const fetcher = vi.fn(async () => ({ v: 1 }));
    await writeDataCache("test:ttl", { v: 1 });

    const value = await fetchWithCache({
      key: "test:ttl",
      fetcher,
      ttlMs: 60_000,
    });

    expect(value).toEqual({ v: 1 });
    expect(fetcher).not.toHaveBeenCalled();
  });

  it("refetches when force is set", async () => {
    const fetcher = vi.fn(async () => ({ v: 2 }));
    await writeDataCache("test:force", { v: 1 });

    const value = await fetchWithCache({
      key: "test:force",
      fetcher,
      ttlMs: 60_000,
      force: true,
    });

    expect(value).toEqual({ v: 2 });
    expect(fetcher).toHaveBeenCalledTimes(1);
  });
});
