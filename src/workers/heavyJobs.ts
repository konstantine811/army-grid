import type { BackendPersonnelOverview, BackendPersonQuestionnaireMeta } from "../api";
import type { AnketaRow } from "../pages/anketa-data/anketaSheet";
import type { BchsPersonnelAwayPerson } from "../pages/bchs/bchsTypes";
import {
  extractBchsAwayPeopleFromDbRows,
  filterBchsNovaPeople,
} from "../pages/bchs/bchsCalc";
import type { ExcelWorkbookSnapshot } from "../excelRoundTrip";
import type { EjournalPreviewRow } from "../pages/ejournal/ejournalTypes";
import type { DbPreviewState } from "../pages/ejournal/ejournalTypes";
import { buildEjoosSyncPlan, type EjoosSyncPlan } from "../pages/ejournal/ejoosSyncPlan";
import { groupOpsIntoPersonChanges } from "../pages/ejournal/ejoosPersonDiff";
import type { EjoosDiffSession } from "../pages/ejournal/ejoosPersonDiff";
import type { EjoosStatusRule } from "../pages/ejournal/ejoosRules";
import { mergeRosterRowsIntoOverview } from "../pages/overview/overviewRosterMerge";
import { mergeRosterRowsIntoPreview } from "../pages/personnel/personnelRosterMerge";
import {
  buildStaffSheetEnrichmentEntries,
  type StaffSheetEnrichmentEntry,
} from "../pages/anketa-data/staffSheetEnrichment";
import { writeStaffSheetExportWorkbook } from "../pages/anketa-data/staffSheetExportWorkbook";
import {
  buildStaffSheetRosterImportPayload,
  type StaffSheetRosterImportPayload,
} from "../pages/excel-fill/staffSheet";
import {
  deserializePersonnelIndex,
  vkEntriesToIndex,
  type SerializedPersonnelIndex,
} from "./staffSheetWorkerSerialize";
import {
  buildVkTpvDovidkyNameIndex,
  parseVkTpvDovidkyWorkbook,
  type VkTpvDovidkyNameEntry,
  type VkTpvDovidkyRecord,
} from "../pages/personnel/vkTpvDovidkyImport";

export type HeavyJob =
  | {
      type: "bchsExtractPeople";
      rows: Array<Record<string, unknown>>;
      columns?: Array<{ key: string; letter?: string; originalIndex?: number }>;
    }
  | {
      type: "mergeOverview";
      overview: BackendPersonnelOverview;
      rosterRows: EjournalPreviewRow[];
      rosterLabels?: Record<string, string>;
      columns?: Array<{ key: string; letter?: string; originalIndex?: number }>;
    }
  | {
      type: "mergePersonnel";
      preview: Pick<DbPreviewState, "rows">;
      rosterRows: EjournalPreviewRow[];
    }
  | {
      type: "ejoosSyncPlan";
      ejoos: ExcelWorkbookSnapshot;
      pb: ExcelWorkbookSnapshot;
      statusRules: EjoosStatusRule[];
      processedMovementKeys?: string[];
      sourceAsOfDate?: string;
    }
  | {
      type: "ejoosSession";
      ejoos: ExcelWorkbookSnapshot;
      pb: ExcelWorkbookSnapshot;
      statusRules: EjoosStatusRule[];
      processedMovementKeys?: string[];
      sourceAsOfDate?: string;
    }
  | {
      type: "staffSheetEnrichment";
      rosterRows: EjournalPreviewRow[];
      personnelIndex: SerializedPersonnelIndex;
      anketaRows: AnketaRow[];
      questionnaires: BackendPersonQuestionnaireMeta[];
      vkEntries: VkTpvDovidkyNameEntry[];
    }
  | {
      type: "staffSheetVkIndex";
      snapshot: ExcelWorkbookSnapshot;
    }
  | {
      type: "parseVkTpvDovidky";
      snapshot: ExcelWorkbookSnapshot;
    }
  | {
      type: "staffSheetRosterImport";
      table: string[][];
      meta: {
        source: "apps-script" | "gviz";
        sourceLabel: string;
        fighterStatusTable?: string[][] | null;
        includeAllRows?: boolean;
      };
    }
  | {
      type: "staffSheetExportWorkbook";
      rosterRows: EjournalPreviewRow[];
      entries: StaffSheetEnrichmentEntry[];
      templateData: ArrayBuffer;
    };

export type HeavyJobResult = {
  bchsExtractPeople: {
    people: BchsPersonnelAwayPerson[];
    novaCount: number;
  };
  mergeOverview: BackendPersonnelOverview;
  mergePersonnel: EjournalPreviewRow[];
  ejoosSyncPlan: EjoosSyncPlan;
  ejoosSession: EjoosDiffSession;
  staffSheetEnrichment: StaffSheetEnrichmentEntry[];
  staffSheetVkIndex: VkTpvDovidkyNameEntry[];
  parseVkTpvDovidky: VkTpvDovidkyRecord[];
  staffSheetRosterImport: StaffSheetRosterImportPayload;
  staffSheetExportWorkbook: ArrayBuffer;
};

export const runHeavyJobSync = <T extends HeavyJob>(
  job: T,
): HeavyJobResult[T["type"]] => {
  switch (job.type) {
    case "bchsExtractPeople": {
      const people = extractBchsAwayPeopleFromDbRows(job.rows, job.columns);
      return {
        people,
        novaCount: filterBchsNovaPeople(people).length,
      } as HeavyJobResult[T["type"]];
    }
    case "mergeOverview":
      return mergeRosterRowsIntoOverview(
        job.overview,
        job.rosterRows,
        job.rosterLabels,
        job.columns,
      ) as HeavyJobResult[T["type"]];
    case "mergePersonnel":
      return mergeRosterRowsIntoPreview(
        job.preview,
        job.rosterRows,
      ) as HeavyJobResult[T["type"]];
    case "ejoosSyncPlan":
      return buildEjoosSyncPlan(job.ejoos, job.pb, {
        statusRules: job.statusRules,
        processedMovementKeys: job.processedMovementKeys,
        sourceAsOfDate: job.sourceAsOfDate,
      }) as HeavyJobResult[T["type"]];
    case "ejoosSession": {
      const plan = buildEjoosSyncPlan(job.ejoos, job.pb, {
        statusRules: job.statusRules,
        processedMovementKeys: job.processedMovementKeys,
        sourceAsOfDate: job.sourceAsOfDate,
      });
      return groupOpsIntoPersonChanges(
        plan,
        job.pb,
        job.statusRules,
      ) as HeavyJobResult[T["type"]];
    }
    case "staffSheetEnrichment":
      return buildStaffSheetEnrichmentEntries({
        rosterRows: job.rosterRows,
        personnelIndex: deserializePersonnelIndex(job.personnelIndex),
        anketaRows: job.anketaRows,
        questionnaires: job.questionnaires,
        vkIndex: vkEntriesToIndex(job.vkEntries),
      }) as HeavyJobResult[T["type"]];
    case "staffSheetVkIndex":
      return [...buildVkTpvDovidkyNameIndex(job.snapshot).values()] as HeavyJobResult[T["type"]];
    case "parseVkTpvDovidky":
      return parseVkTpvDovidkyWorkbook(job.snapshot) as HeavyJobResult[T["type"]];
    case "staffSheetRosterImport":
      return buildStaffSheetRosterImportPayload(job.table, job.meta) as HeavyJobResult[T["type"]];
    default: {
      const neverJob: never = job;
      throw new Error(`Unknown heavy job: ${(neverJob as HeavyJob).type}`);
    }
  }
};

export const runHeavyJobMaybeAsync = async (
  job: HeavyJob,
): Promise<unknown> => {
  if (job.type === "staffSheetExportWorkbook") {
    return writeStaffSheetExportWorkbook(job.rosterRows, job.entries, {
      templateData: job.templateData,
      download: false,
    });
  }
  return runHeavyJobSync(job);
};
