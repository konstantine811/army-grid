import type { BackendEjournalImportSheet } from "../../api";

export type EjournalColumn = {
  key: string;
  label: string;
  order: number;
  originalIndex?: number;
  letter?: string;
};

export type EjournalPreviewRow = {
  __dbRowId?: string;
  __rowNumber?: number | null;
  [key: string]: unknown;
};

export type DbPreviewState = {
  sheet: BackendEjournalImportSheet;
  columns: EjournalColumn[];
  rows: EjournalPreviewRow[];
  total: number;
  offset: number;
  limit: number;
};
