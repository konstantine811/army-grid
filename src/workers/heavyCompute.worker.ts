/// <reference lib="webworker" />

import { runHeavyJobMaybeAsync, type HeavyJob } from "./heavyJobs";

type WorkerRequest = {
  id: number;
  job: HeavyJob;
};

self.onmessage = (event: MessageEvent<WorkerRequest>) => {
  const { id, job } = event.data;
  void runHeavyJobMaybeAsync(job)
    .then((result) => {
      self.postMessage({ id, ok: true, result });
    })
    .catch((error) => {
      self.postMessage({
        id,
        ok: false,
        error: error instanceof Error ? error.message : String(error),
      });
    });
};
