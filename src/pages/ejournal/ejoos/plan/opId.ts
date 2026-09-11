export const buildEjoosOpId = (parts: string[]) =>
  parts
    .map((part) => part.replace(/\s+/g, "_").slice(0, 40))
    .filter(Boolean)
    .join("__");

/** @deprecated Use buildEjoosOpId — alias for incremental migration from ejoosSyncPlan. */
export const opId = buildEjoosOpId;
