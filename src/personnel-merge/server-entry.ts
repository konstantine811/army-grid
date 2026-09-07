/** Node/CJS bundle entry — same merge logic as the browser worker. */
export { buildPersonnelDatasetSync, findEjournalPersonnelSheet } from "../data/personnelDatasetCore";
export { mergeRosterRowsIntoOverview } from "../pages/overview/overviewRosterMerge";
export { overviewMergeFingerprint } from "../pages/overview/overviewMergeCache";
