export type ApiRequestPriority = "high" | "normal";

type QueueEntry<T> = {
  task: () => Promise<T>;
  signal?: AbortSignal;
  resolve: (value: T) => void;
  reject: (reason?: unknown) => void;
  cancelled: boolean;
};

const abortError = () =>
  typeof DOMException === "function"
    ? new DOMException("The operation was aborted.", "AbortError")
    : Object.assign(new Error("The operation was aborted."), {
        name: "AbortError",
      });

export class ApiRequestPool {
  private active = 0;
  private readonly highQueue: QueueEntry<unknown>[] = [];
  private readonly normalQueue: QueueEntry<unknown>[] = [];
  private readonly concurrency: number;

  constructor(concurrency = 4) {
    this.concurrency = concurrency;
  }

  run<T>(
    task: () => Promise<T>,
    options: { priority?: ApiRequestPriority; signal?: AbortSignal } = {},
  ): Promise<T> {
    if (options.signal?.aborted) return Promise.reject(abortError());

    return new Promise<T>((resolve, reject) => {
      const entry: QueueEntry<T> = {
        task,
        signal: options.signal,
        resolve,
        reject,
        cancelled: false,
      };
      const queue =
        options.priority === "high" ? this.highQueue : this.normalQueue;
      queue.push(entry as QueueEntry<unknown>);

      const onAbort = () => {
        if (entry.cancelled) return;
        entry.cancelled = true;
        reject(abortError());
        this.drain();
      };
      options.signal?.addEventListener("abort", onAbort, { once: true });
      this.drain();
    });
  }

  private drain() {
    while (this.active < this.concurrency) {
      const entry = this.highQueue.shift() ?? this.normalQueue.shift();
      if (!entry) return;
      if (entry.cancelled || entry.signal?.aborted) {
        if (!entry.cancelled) entry.reject(abortError());
        continue;
      }

      this.active += 1;
      void entry
        .task()
        .then(entry.resolve, entry.reject)
        .finally(() => {
          this.active -= 1;
          this.drain();
        });
    }
  }
}

export const apiRequestPool = new ApiRequestPool(4);

