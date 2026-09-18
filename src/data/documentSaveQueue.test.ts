import { describe, expect, it } from "vitest";
import { enqueueDocumentSave, waitForDocumentSaves } from "./documentSaveQueue";

describe("document save ordering", () => {
  it("finishes the old write before starting the new one and waits before reading", async () => {
    let release!: () => void;
    let stored = "";
    const first = enqueueDocumentSave("a", async () => {
      await new Promise<void>((resolve) => { release = resolve; });
      stored = "old";
    });
    const second = enqueueDocumentSave("a", async () => { stored = "new"; });
    let ready = false;
    const read = waitForDocumentSaves("a").then(() => { ready = true; });
    await new Promise((resolve) => setTimeout(resolve, 0));
    expect(stored).toBe("");
    expect(ready).toBe(false);
    release();
    await Promise.all([first, second, read]);
    expect(stored).toBe("new");
    expect(ready).toBe(true);
  });

  it("allows other documents to save and recovers after a failed write", async () => {
    const failed = enqueueDocumentSave("b", async () => { throw new Error("offline"); });
    const rejection = expect(failed).rejects.toThrow("offline");
    const next = enqueueDocumentSave("b", async () => "recovered");
    expect(await enqueueDocumentSave("c", async () => "independent")).toBe("independent");
    await rejection;
    expect(await next).toBe("recovered");
    await waitForDocumentSaves("b");
  });
});
