import { describe, expect, it } from "vitest";
import { ApiRequestPool } from "./apiRequestPool";

const deferred = () => {
  let resolve!: () => void;
  const promise = new Promise<void>((done) => {
    resolve = done;
  });
  return { promise, resolve };
};

describe("ApiRequestPool", () => {
  it("never exceeds its concurrency limit", async () => {
    const pool = new ApiRequestPool(2);
    let active = 0;
    let maxActive = 0;

    await Promise.all(
      Array.from({ length: 8 }, () =>
        pool.run(async () => {
          active += 1;
          maxActive = Math.max(maxActive, active);
          await new Promise((resolve) => setTimeout(resolve, 2));
          active -= 1;
        }),
      ),
    );

    expect(maxActive).toBe(2);
  });

  it("runs queued writes before queued background reads", async () => {
    const pool = new ApiRequestPool(1);
    const gate = deferred();
    const order: string[] = [];
    const first = pool.run(async () => {
      order.push("active-read");
      await gate.promise;
    });
    const normal = pool.run(async () => {
      order.push("queued-read");
    });
    const high = pool.run(
      async () => {
        order.push("write");
      },
      { priority: "high" },
    );

    gate.resolve();
    await Promise.all([first, normal, high]);

    expect(order).toEqual(["active-read", "write", "queued-read"]);
  });

  it("cancels a request while it is still queued", async () => {
    const pool = new ApiRequestPool(1);
    const gate = deferred();
    const first = pool.run(() => gate.promise);
    const controller = new AbortController();
    const queued = pool.run(async () => undefined, {
      signal: controller.signal,
    });

    controller.abort();
    await expect(queued).rejects.toMatchObject({ name: "AbortError" });
    gate.resolve();
    await first;
  });
});
