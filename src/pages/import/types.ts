import type { Dispatch, SetStateAction } from "react";
import type { ExcelRow, ExcelWorkbookSnapshot } from "../../excelRoundTrip";

export type DataSourceFile = {
  id: string;
  name: string;
  size: string;
  rows: number;
  columns: number;
  uploadedAt: string;
  role: "Основний файл" | "Файл для merge";
};

export type ExcelRoundTripPanelProps = {
  snapshot: ExcelWorkbookSnapshot | null;
  rows: ExcelRow[];
  setRows: Dispatch<SetStateAction<ExcelRow[]>>;
  activeSheetIndex: number;
  onActiveSheetChange: (sheetIndex: number) => void;
  isBusy: boolean;
  message: string;
  mergeSummary: string;
  onExport: () => void;
};

export type ExcelCellEditorProps = {
  value: string;
  wrapText: boolean;
  onCommit: (value: string) => void;
};
