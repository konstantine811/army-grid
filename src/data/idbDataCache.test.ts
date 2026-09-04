import { afterEach, describe, expect, it, vi } from "vitest";
import {
  fetchWithCache,
  invalidateDataCache,
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

  it("does not restore a stale request after invalidation", async () => {
    let resolveFirst!: (value: { v: number }) => void;
    const firstResponse = new Promise<{ v: number }>((resolve) => {
      resolveFirst = resolve;
    });
    const fetcher = vi
      .fn<() => Promise<{ v: number }>>()
      .mockReturnValueOnce(firstResponse)
      .mockResolvedValue({ v: 2 });

    const first = fetchWithCache({
      key: "test:race",
      fetcher,
      force: true,
    });
    await vi.waitFor(() => expect(fetcher).toHaveBeenCalledTimes(1));

    await invalidateDataCache("test:race");
    const second = fetchWithCache({
      key: "test:race",
      fetcher,
      force: true,
    });
    resolveFirst({ v: 1 });

    await expect(first).resolves.toEqual({ v: 2 });
    await expect(second).resolves.toEqual({ v: 2 });
    expect(fetcher).toHaveBeenCalledTimes(2);
  });
});
