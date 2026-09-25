import { CacheKeys, readDataCache, writeDataCache } from "../../data/idbDataCache";
import type { StaffListSnapshot } from "./staffListGapsParse";

export const STAFF_LIST_GAPS_CACHE_KEY = CacheKeys.staffListGapsSession;

export type StaffListGapsSession = {
  savedAt: string;
  snapshot: StaffListSnapshot;
  gapColumnIds: string[];
  fileData: ArrayBuffer;
};

export const loadStaffListGapsSession = async () =>
  readDataCache<StaffListGapsSession>(STAFF_LIST_GAPS_CACHE_KEY);

export const saveStaffListGapsSession = async (session: StaffListGapsSession) => {
  await writeDataCache(STAFF_LIST_GAPS_CACHE_KEY, session);
};

export const clearStaffListGapsSession = async () => {
  await writeDataCache(STAFF_LIST_GAPS_CACHE_KEY, null);
};

export const buildStaffListGapsSession = (
  snapshot: StaffListSnapshot,
  rows: StaffListSnapshot["rows"],
  gapColumnIds: string[],
  fileData: ArrayBuffer,
): StaffListGapsSession => ({
  savedAt: new Date().toISOString(),
  snapshot: { ...snapshot, rows },
  gapColumnIds,
  fileData,
});
