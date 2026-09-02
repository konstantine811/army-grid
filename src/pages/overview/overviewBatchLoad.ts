import { api, type BackendPersonnelOverview } from "../../api";

export const OVERVIEW_PAGE_SIZE = 20;

export const mergeOverviewPages = (
  current: BackendPersonnelOverview | null,
  page: BackendPersonnelOverview,
): BackendPersonnelOverview => {
  if (!current) return page;

  const byId = new Map(current.rows.map((row) => [row.id, row]));
  for (const row of page.rows) byId.set(row.id, row);
  const rows = [...byId.values()];
  const units = [
    ...new Set([...current.units, ...page.units].filter(Boolean)),
  ].sort((left, right) =>
    left.localeCompare(right, "uk", { numeric: true, sensitivity: "base" }),
  );
  const metrics =
    page.metrics.total >= current.metrics.total ? page.metrics : current.metrics;

  return {
    ...page,
    rows,
    units,
    metrics,
  };
};

export const overviewHasMorePages = (
  page: BackendPersonnelOverview,
  loadedCount: number,
  requestedLimit: number,
) => {
  const total = Math.max(page.metrics?.total ?? 0, loadedCount);
  if (loadedCount >= total) return false;
  if (page.rows.length === 0) return false;
  if (page.rows.length >= total) return false;
  if (page.rows.length < requestedLimit) return false;
  return true;
};

export const loadPersonnelOverviewInBatches = async (input: {
  pageSize?: number;
  isCancelled?: () => boolean;
  onPage?: (
    overview: BackendPersonnelOverview,
    meta: { done: number; total: number; complete: boolean },
  ) => void | Promise<void>;
}): Promise<BackendPersonnelOverview> => {
  const pageSize = input.pageSize ?? OVERVIEW_PAGE_SIZE;
  let merged: BackendPersonnelOverview | null = null;
  let offset = 0;

  const fetchPage = async (limit: number, pageOffset: number) => {
    try {
      return await api.getPersonnelOverview({ limit, offset: pageOffset });
    } catch (error) {
      if (pageOffset === 0) return api.getPersonnelOverview();
      throw error;
    }
  };

  while (!input.isCancelled?.()) {
    const previousCount = merged?.rows.length ?? 0;
    const page = await fetchPage(pageSize, offset);
    merged = mergeOverviewPages(merged, page);
    const done = merged.rows.length;
    const total = Math.max(merged.metrics.total, done);
    const stalled = offset > 0 && done === previousCount;
    const more = !stalled && overviewHasMorePages(page, done, pageSize);
    await input.onPage?.(merged, {
      done,
      total,
      complete: !more,
    });
    if (!more) return merged;
    offset = done;
    await new Promise<void>((resolve) => {
      window.setTimeout(resolve, 0);
    });
  }

  if (!merged) {
    throw new Error("Завантаження огляду скасовано");
  }
  return merged;
};
