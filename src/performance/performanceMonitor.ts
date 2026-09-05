export type AppPerformanceEvent =
  | {
      type: "api";
      at: number;
      method: string;
      path: string;
      status: number;
      durationMs: number;
      responseBytes: number | null;
      requestBytes: number | null;
    }
  | {
      type: "long-task";
      at: number;
      durationMs: number;
    }
  | {
      type: "memory";
      at: number;
      usedBytes: number;
      totalBytes: number;
      limitBytes: number;
    };

const STORAGE_KEY = "army-grid:performance:v1";
const MAX_EVENTS = 200;
const SLOW_API_MS = 3_000;
const LARGE_RESPONSE_BYTES = 2 * 1024 * 1024;
let events: AppPerformanceEvent[] = [];
let flushTimer: number | null = null;
let started = false;

const readStoredEvents = () => {
  try {
    const parsed = JSON.parse(localStorage.getItem(STORAGE_KEY) || "[]");
    return Array.isArray(parsed) ? parsed.slice(-MAX_EVENTS) : [];
  } catch {
    return [];
  }
};

const flush = () => {
  flushTimer = null;
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(events.slice(-MAX_EVENTS)));
  } catch {
    // Diagnostics must never affect application behavior.
  }
};

const scheduleFlush = () => {
  if (flushTimer != null || typeof window === "undefined") return;
  flushTimer = window.setTimeout(flush, 2_000);
};

const record = (event: AppPerformanceEvent) => {
  events.push(event);
  if (events.length > MAX_EVENTS) events = events.slice(-MAX_EVENTS);
  scheduleFlush();
};

/** Removes query values and person/document identifiers from diagnostics. */
export const sanitizePerformancePath = (input: string) => {
  try {
    const base =
      typeof window === "undefined" ? "http://army-grid.local" : window.location.origin;
    const url = new URL(input, base);
    const path = url.pathname
      .split("/")
      .map((part) =>
        /^\d+$/.test(part) ||
        /^[0-9a-f]{8}-[0-9a-f-]{27,}$/i.test(part) ||
        /^[a-z0-9:_-]{20,}$/i.test(part)
          ? ":id"
          : part,
      )
      .join("/");
    const queryKeys = [...url.searchParams.keys()].sort();
    return queryKeys.length ? `${path}?${queryKeys.join("&")}` : path;
  } catch {
    return String(input).split("?")[0];
  }
};

const bodySize = (body: BodyInit | null | undefined): number | null => {
  if (typeof body === "string") return new TextEncoder().encode(body).byteLength;
  if (body instanceof Blob) return body.size;
  if (body instanceof ArrayBuffer) return body.byteLength;
  if (ArrayBuffer.isView(body)) return body.byteLength;
  return null;
};

export const measuredFetch = async (
  input: RequestInfo | URL,
  init?: RequestInit,
) => {
  const startedAt = performance.now();
  const method =
    init?.method || (input instanceof Request ? input.method : "GET");
  const rawUrl =
    input instanceof Request ? input.url : String(input);
  let status = 0;
  let responseBytes: number | null = null;
  try {
    const response = await fetch(input, init);
    status = response.status;
    const contentLength = Number(response.headers.get("content-length"));
    responseBytes =
      Number.isFinite(contentLength) && contentLength >= 0 ? contentLength : null;
    return response;
  } finally {
    const durationMs = performance.now() - startedAt;
    const event: AppPerformanceEvent = {
      type: "api",
      at: Date.now(),
      method: method.toUpperCase(),
      path: sanitizePerformancePath(rawUrl),
      status,
      durationMs: Math.round(durationMs),
      responseBytes,
      requestBytes: bodySize(init?.body),
    };
    record(event);
    if (
      durationMs >= SLOW_API_MS ||
      (responseBytes ?? 0) >= LARGE_RESPONSE_BYTES
    ) {
      console.warn("[Performance] Heavy API request", event);
    }
  }
};

const sampleMemory = () => {
  const memory = (
    performance as Performance & {
      memory?: {
        usedJSHeapSize: number;
        totalJSHeapSize: number;
        jsHeapSizeLimit: number;
      };
    }
  ).memory;
  if (!memory) return;
  record({
    type: "memory",
    at: Date.now(),
    usedBytes: memory.usedJSHeapSize,
    totalBytes: memory.totalJSHeapSize,
    limitBytes: memory.jsHeapSizeLimit,
  });
};

export const startPerformanceMonitoring = () => {
  if (started || typeof window === "undefined") return;
  started = true;
  events = readStoredEvents();
  if (typeof PerformanceObserver !== "undefined") {
    try {
      const observer = new PerformanceObserver((list) => {
        for (const entry of list.getEntries()) {
          record({
            type: "long-task",
            at: Date.now(),
            durationMs: Math.round(entry.duration),
          });
        }
      });
      observer.observe({ type: "longtask", buffered: true });
    } catch {
      // Long Tasks API is not available in every browser.
    }
  }
  sampleMemory();
  window.setInterval(sampleMemory, 30_000);
  (
    window as Window & {
      armyGridPerformance?: {
        snapshot: typeof getPerformanceSnapshot;
        clear: typeof clearPerformanceSnapshot;
      };
    }
  ).armyGridPerformance = {
    snapshot: getPerformanceSnapshot,
    clear: clearPerformanceSnapshot,
  };
};

export const getPerformanceSnapshot = () => [...events];

export const clearPerformanceSnapshot = () => {
  events = [];
  try {
    localStorage.removeItem(STORAGE_KEY);
  } catch {
    // Ignore private mode restrictions.
  }
};
