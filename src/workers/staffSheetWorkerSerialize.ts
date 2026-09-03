import type {
  AnketaPersonnelIndex,
  AnketaPersonnelMatch,
} from "../pages/anketa-data/anketaPersonMatch";
import type { BackendPersonQuestionnaireMeta } from "../api";
import type { AnketaRow } from "../pages/anketa-data/anketaSheet";
import type { EjournalPreviewRow } from "../pages/ejournal/ejournalTypes";
import type { VkTpvDovidkyNameEntry } from "../pages/personnel/vkTpvDovidkyImport";

export type SerializedPersonnelIndex = {
  byExternalId: Array<[string, AnketaPersonnelMatch]>;
  byRnokpp: Array<[string, AnketaPersonnelMatch]>;
  byNameBirth: Array<[string, AnketaPersonnelMatch]>;
  byName: Array<[string, AnketaPersonnelMatch[]]>;
  byShortName: Array<[string, AnketaPersonnelMatch[]]>;
};

export const serializePersonnelIndex = (
  index: AnketaPersonnelIndex,
): SerializedPersonnelIndex => ({
  byExternalId: [...index.byExternalId.entries()],
  byRnokpp: [...index.byRnokpp.entries()],
  byNameBirth: [...index.byNameBirth.entries()],
  byName: [...index.byName.entries()],
  byShortName: [...index.byShortName.entries()],
});

export const deserializePersonnelIndex = (
  index: SerializedPersonnelIndex,
): AnketaPersonnelIndex => ({
  byExternalId: new Map(index.byExternalId),
  byRnokpp: new Map(index.byRnokpp),
  byNameBirth: new Map(index.byNameBirth),
  byName: new Map(index.byName),
  byShortName: new Map(index.byShortName),
});

export const vkIndexToEntries = (
  index: Map<string, VkTpvDovidkyNameEntry>,
): VkTpvDovidkyNameEntry[] => [...index.values()];

export const vkEntriesToIndex = (
  entries: VkTpvDovidkyNameEntry[],
): Map<string, VkTpvDovidkyNameEntry> =>
  new Map(entries.map((entry) => [entry.nameKey, entry]));

export type StaffSheetEnrichmentWorkerInput = {
  rosterRows: EjournalPreviewRow[];
  personnelIndex: AnketaPersonnelIndex;
  anketaRows: AnketaRow[];
  questionnaires: BackendPersonQuestionnaireMeta[];
  vkIndex: Map<string, VkTpvDovidkyNameEntry>;
};

export const toStaffSheetEnrichmentJobPayload = (
  input: StaffSheetEnrichmentWorkerInput,
) => ({
  rosterRows: input.rosterRows,
  personnelIndex: serializePersonnelIndex(input.personnelIndex),
  anketaRows: input.anketaRows,
  questionnaires: input.questionnaires,
  vkEntries: vkIndexToEntries(input.vkIndex),
});
