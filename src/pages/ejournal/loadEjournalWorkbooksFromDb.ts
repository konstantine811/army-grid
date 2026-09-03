import { api } from "../../api";
import {
  EJOOS_SYNC_READ_OPTIONS,
  readWorkbookSnapshot,
  type ExcelWorkbookSnapshot,
} from "../../excelRoundTrip";
import { base64ToFile } from "./ejoosSyncApply";
import { readEjoosWorkbookSnapshot } from "./ejoosTimesheetPersonRows";
import { assertPbWorkbook } from "./ejoosWorkbookKind";

const UNIT_LABEL = "1ПБ";

export const loadEjoosWorkbookFromDb = async (): Promise<ExcelWorkbookSnapshot> => {
  const live = await api.getEjournalLive(UNIT_LABEL);
  if (!live.current?.id) {
    throw new Error(
      "Немає збереженого ЕЖООС у БД. Завантажте журнал на сторінці ЄЖООС.",
    );
  }

  const full = await api.getEjournalLiveFile(live.current.id, UNIT_LABEL);
  if (!full.fileBase64) {
    throw new Error("У поточній версії ЕЖООС немає файлу.");
  }

  const file = base64ToFile(
    full.fileBase64,
    full.sourceFileName || `ЕЖООС_v${full.version}.xlsx`,
  );
  return readEjoosWorkbookSnapshot(file);
};

export const loadPbWorkbookFromDb = async (): Promise<ExcelWorkbookSnapshot | null> => {
  const sources = await api.getEjournalPbSources(UNIT_LABEL);
  const sourceId = sources.current?.id;
  if (!sourceId) return null;

  const remote = await api.getEjournalPbFile(sourceId, UNIT_LABEL);
  if (!remote.fileBase64) return null;

  const file = base64ToFile(
    remote.fileBase64,
    remote.sourceFileName || "1PB.xlsx",
  );
  const snapshot = await readWorkbookSnapshot(file, EJOOS_SYNC_READ_OPTIONS);
  assertPbWorkbook(snapshot);
  return snapshot;
};
