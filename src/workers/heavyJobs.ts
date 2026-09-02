import type {
  BackendPersonnelOverview,
} from "../api";
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
    default: {
      const neverJob: never = job;
      throw new Error(`Unknown heavy job: ${(neverJob as HeavyJob).type}`);
    }
  }
};
