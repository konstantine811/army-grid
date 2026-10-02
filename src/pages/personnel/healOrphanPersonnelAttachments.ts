import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import {
  buildOrphanAttachmentMigrationPairs,
  migratePersonAttachmentsBetweenIds,
} from "./personAttachments";
import {
  migrateStoredPersonPhones,
  readStoredPersonPhones,
} from "./personPhonesStore";
import { migrateStoredPersonSignatures } from "./personSignatureStore";
import {
  isLikelyPersonnelRow,
  resolvePersonIdentityKey,
} from "./personnelUtils";

export const ATTACHMENT_HEAL_SESSION_KEY = "army-grid:attachments-healed-v2";

export const collectPersonnelIdentityIds = (rows: EjournalPreviewRow[]) => {
  const currentIds = new Set<string>();
  for (const row of rows) {
    if (!isLikelyPersonnelRow(row)) continue;
    const id = resolvePersonIdentityKey(row);
    if (id) currentIds.add(id);
  }
  return currentIds;
};

export const collectOrphanAttachmentIds = ({
  currentIds,
  photos,
  questionnaires,
  storedPhones,
}: {
  currentIds: Set<string>;
  photos: Array<{ personExternalId?: string | null }>;
  questionnaires: Array<{ personExternalId?: string | null }>;
  storedPhones: Record<string, string[]>;
}) => {
  const orphanIds = new Set<string>();
  for (const photo of photos) {
    const id = photo.personExternalId?.trim();
    if (id && !currentIds.has(id)) orphanIds.add(id);
  }
  for (const item of questionnaires) {
    const id = item.personExternalId?.trim();
    if (id && !currentIds.has(id)) orphanIds.add(id);
  }
  for (const id of Object.keys(storedPhones)) {
    if (id && !currentIds.has(id) && storedPhones[id]?.length) orphanIds.add(id);
  }
  return orphanIds;
};

export const healOrphanPersonnelAttachments = async ({
  rows,
  isCancelled,
  photosPromise,
  questionnairesPromise,
  onPhonesReplaced,
  reloadAttachments,
}: {
  rows: EjournalPreviewRow[];
  isCancelled?: () => boolean;
  photosPromise: Promise<Array<{ personExternalId: string; photoData: string }>>;
  questionnairesPromise: Promise<Array<{ personExternalId: string }>>;
  onPhonesReplaced: (phones: Record<string, string[]>) => void;
  reloadAttachments: () => Promise<void>;
}) => {
  try {
    if (sessionStorage.getItem(ATTACHMENT_HEAL_SESSION_KEY) === "1") return;
    await new Promise<void>((resolve) => {
      window.setTimeout(resolve, 0);
    });
    if (isCancelled?.()) return;

    const [photos, questionnaires] = await Promise.all([
      photosPromise,
      questionnairesPromise,
    ]);
    if (isCancelled?.()) return;

    const currentIds = collectPersonnelIdentityIds(rows);
    const orphanIds = collectOrphanAttachmentIds({
      currentIds,
      photos,
      questionnaires,
      storedPhones: readStoredPersonPhones(),
    });
    const pairs = buildOrphanAttachmentMigrationPairs(
      rows,
      orphanIds,
      questionnaires,
    );
    if (!isCancelled?.()) {
      onPhonesReplaced(migrateStoredPersonPhones(pairs));
      migrateStoredPersonSignatures(pairs);
    }
    if (!orphanIds.size) {
      sessionStorage.setItem(ATTACHMENT_HEAL_SESSION_KEY, "1");
      return;
    }

    const migrated = await migratePersonAttachmentsBetweenIds(pairs, {
      includeDocuments: false,
      photos,
      questionnaires,
    });
    sessionStorage.setItem(ATTACHMENT_HEAL_SESSION_KEY, "1");
    if (migrated > 0 && !isCancelled?.()) {
      await reloadAttachments();
    }
  } catch {
    // Background heal must never block the personnel list.
  }
};
