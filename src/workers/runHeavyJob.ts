import { runHeavyJobMaybeAsync, type HeavyJob, type HeavyJobResult } from "./heavyJobs";
import type { ExcelWorkbookSnapshot } from "../excelRoundTrip";

type WorkerResponse = {
  id: number;
  ok: boolean;
  result?: unknown;
  error?: string;
};

let worker: Worker | null = null;
let seq = 0;
let activeJobs = 0;
let terminateTimer: ReturnType<typeof setTimeout> | null = null;

const scheduleWorkerRelease = () => {
  if (activeJobs > 0 || !worker) return;
  if (terminateTimer) clearTimeout(terminateTimer);
  terminateTimer = setTimeout(() => {
    terminateTimer = null;
    if (activeJobs > 0) return;
    worker?.terminate();
    worker = null;
  }, 1_500);
};

const dropWorkbookFile = (workbook: ExcelWorkbookSnapshot): ExcelWorkbookSnapshot => ({
  ...workbook,
  file: undefined as unknown as File,
});

const toWorkerJob = <T extends HeavyJob>(job: T): T => {
  if (job.type === "ejoosSyncPlan" || job.type === "ejoosSession") {
    return {
      ...job,
      ejoos: dropWorkbookFile(job.ejoos),
      pb: dropWorkbookFile(job.pb),
    };
  }
  if (job.type === "staffSheetVkIndex" || job.type === "parseVkTpvDovidky") {
    return {
      ...job,
      snapshot: dropWorkbookFile(job.snapshot),
    };
  }
  return job;
};

const getWorker = () => {
  if (typeof Worker === "undefined") return null;
  if (terminateTimer) {
    clearTimeout(terminateTimer);
    terminateTimer = null;
  }
  if (worker) return worker;
  try {
    worker = new Worker(new URL("./heavyCompute.worker.ts", import.meta.url), {
      type: "module",
    });
    worker.addEventListener("error", () => {
      worker?.terminate();
      worker = null;
    });
    return worker;
  } catch {
    worker = null;
    return null;
  }
};

/** Run CPU-heavy roster / ЕЖООС / БЧС work off the UI thread when Worker exists. */
export const runHeavyJob = <T extends HeavyJob>(
  job: T,
): Promise<HeavyJobResult[T["type"]]> => {
  const target = getWorker();
  if (!target) {
    return runHeavyJobMaybeAsync(job) as Promise<HeavyJobResult[T["type"]]>;
  }

  return new Promise((resolve, reject) => {
    activeJobs += 1;
    const id = (seq += 1);
    let settled = false;
    const finish = () => {
      if (settled) return;
      settled = true;
      activeJobs = Math.max(0, activeJobs - 1);
      scheduleWorkerRelease();
    };
    const onMessage = (event: MessageEvent<WorkerResponse>) => {
      if (event.data.id !== id) return;
      target.removeEventListener("message", onMessage);
      target.removeEventListener("error", onError);
      finish();
      if (event.data.ok) {
        resolve(event.data.result as HeavyJobResult[T["type"]]);
        return;
      }
      reject(new Error(event.data.error || "Worker job failed"));
    };
    const onError = () => {
      target.removeEventListener("message", onMessage);
      target.removeEventListener("error", onError);
      finish();
      reject(new Error("Heavy worker failed"));
    };
    target.addEventListener("message", onMessage);
    target.addEventListener("error", onError);
    try {
      target.postMessage({ id, job: toWorkerJob(job) });
    } catch {
      target.removeEventListener("message", onMessage);
      target.removeEventListener("error", onError);
      finish();
      void runHeavyJobMaybeAsync(job)
        .then((result) => resolve(result as HeavyJobResult[T["type"]]))
        .catch((error) =>
          reject(error instanceof Error ? error : new Error(String(error))),
        );
    }
  });
};
