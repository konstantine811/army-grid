import type { AppPage } from "./navigation";

/** Keep recently visited pages mounted so sidebar clicks do not reload them. */
export const PAGE_KEEP_ALIVE_LIMIT = 6;

export const rememberMountedPage = (
  current: ReadonlySet<AppPage>,
  page: AppPage,
  limit = PAGE_KEEP_ALIVE_LIMIT,
): Set<AppPage> => {
  if (current.has(page) && current.size <= limit) {
    return current instanceof Set ? current : new Set(current);
  }

  const next = new Set(current);
  next.delete(page);
  next.add(page);
  for (const item of next) {
    if (next.size <= limit) break;
    if (item !== page) next.delete(item);
  }
  return next;
};
