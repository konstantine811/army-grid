import {
  api,
  type BackendPersonnelBootstrap,
  type BackendPersonnelVersionProbe,
  type PersonnelBootstrapInclude,
} from "../api";

const BOOTSTRAP_TTL_MS = 30_000;

let versionInFlight: Promise<BackendPersonnelVersionProbe | null> | null = null;
let bootstrapInFlight: Promise<BackendPersonnelBootstrap | null> | null = null;
let cachedBootstrap: BackendPersonnelBootstrap | null = null;
let cachedBootstrapAt = 0;

/** Lightweight server fingerprint probe — replaces imports + roster version for staleness. */
export const loadPersonnelVersionProbe = async (options?: {
  force?: boolean;
  signal?: AbortSignal;
}): Promise<BackendPersonnelVersionProbe | null> => {
  if (versionInFlight && !options?.force) return versionInFlight;

  const promise = api.getPersonnelVersion({ signal: options?.signal });
  versionInFlight = promise;
  try {
    return await promise;
  } finally {
    if (versionInFlight === promise) versionInFlight = null;
  }
};

/** Snapshot availability bundle for deciding GET vs rebuild. */
export const loadPersonnelBootstrapMeta = async (options?: {
  force?: boolean;
  signal?: AbortSignal;
  include?: PersonnelBootstrapInclude[];
}): Promise<BackendPersonnelBootstrap | null> => {
  const wantsPayload = Boolean(options?.include?.length);
  if (
    !options?.force &&
    !wantsPayload &&
    cachedBootstrap &&
    Date.now() - cachedBootstrapAt < BOOTSTRAP_TTL_MS
  ) {
    return cachedBootstrap;
  }

  if (bootstrapInFlight && !options?.force && !wantsPayload) {
    return bootstrapInFlight;
  }

  const promise = api
    .getPersonnelBootstrap({
      signal: options?.signal,
      include: options?.include,
    })
    .then(
    (result) => {
      if (result && !wantsPayload) {
        cachedBootstrap = result;
        cachedBootstrapAt = Date.now();
      }
      return result;
    },
  );
  bootstrapInFlight = promise;
  try {
    return await promise;
  } finally {
    if (bootstrapInFlight === promise) bootstrapInFlight = null;
  }
};

export const peekPersonnelBootstrapMeta = () => cachedBootstrap;

export const clearPersonnelBootstrapMeta = () => {
  cachedBootstrap = null;
  cachedBootstrapAt = 0;
};
