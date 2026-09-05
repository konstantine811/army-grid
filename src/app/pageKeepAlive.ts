import type { AppPage } from "./navigation";

/** Only the active page stays mounted; navigation releases the previous page. */
export const PAGE_KEEP_ALIVE_LIMIT = 1;

export const rememberMountedPage = (
  current: ReadonlySet<AppPage>,
  page: AppPage,
  limit = PAGE_KEEP_ALIVE_LIMIT,
): Set<AppPage> => {
  const safeLimit = Math.max(1, Math.floor(limit));
  const currentOrder = [...current];
  if (
    current.has(page) &&
    current.size <= safeLimit &&
    currentOrder[currentOrder.length - 1] === page
  ) {
    return current instanceof Set ? current : new Set(current);
  }

  const next = new Set(current);
  // Set insertion order is our lightweight LRU queue. Revisiting a page moves
  // it to the end, so a genuinely old background page is evicted next.
  next.delete(page);
  next.add(page);

  for (const item of next) {
    if (next.size <= safeLimit) break;
    if (item !== page) next.delete(item);
  }
  return next;
};
