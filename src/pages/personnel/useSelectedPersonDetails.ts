import { useEffect } from "react";
import {
  api,
  type BackendPersonDocument,
  type BackendPersonQuestionnaire,
} from "../../api";
import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import {
  loadPersonDocumentsForRow,
  loadPersonQuestionnaireForRow,
} from "./personAttachments";
import {
  extractPhonesFromDocuments,
  readStoredPersonPhones,
  uniqueNormalizedPhones,
  upsertPersonPhonesDocument,
  writeStoredPersonPhones,
} from "./personPhonesStore";
import { samePersonNameCount } from "./personnelQuestionnaireOpen";
import {
  mergedDocumentPhones,
  shouldCopyFoundQuestionnaire,
} from "./personnelSelectedCard";

export const useSelectedPersonDetails = ({
  enabled,
  canEdit,
  personnelRows,
  selectedRow,
  externalId,
  personName,
  setQuestionnaire,
  setDocuments,
  markQuestionnairePresent,
  setPhoneDocument,
  setPhonesByExternalId,
}: {
  enabled: boolean;
  canEdit: boolean;
  personnelRows: Array<{ summary: { name: string } }>;
  selectedRow: EjournalPreviewRow | null;
  externalId: string;
  personName: string;
  setQuestionnaire: (value: BackendPersonQuestionnaire | null) => void;
  setDocuments: (value: BackendPersonDocument[]) => void;
  markQuestionnairePresent: (externalId: string) => void;
  setPhoneDocument: (externalId: string, document: BackendPersonDocument) => void;
  setPhonesByExternalId: (
    update: (
      current: Record<string, string[]>,
    ) => Record<string, string[]>,
  ) => void;
}) => {
  useEffect(() => {
    if (!enabled) return;
    if (!externalId) {
      setQuestionnaire(null);
      setDocuments([]);
      return;
    }
    let isCancelled = false;
    const nameIsAmbiguous =
      samePersonNameCount(
        personnelRows.map((item) => item.summary.name),
        personName,
      ) > 1;
    setQuestionnaire(null);
    void loadPersonQuestionnaireForRow(selectedRow, undefined, {
      nameIsAmbiguous,
    })
      .then(async ({ questionnaire: next, resolvedExternalId }) => {
        if (isCancelled) return;
        if (
          shouldCopyFoundQuestionnaire(
            canEdit,
            Boolean(next),
            resolvedExternalId ?? "",
            externalId,
          )
        ) {
          try {
            const copied = await api.copyPersonQuestionnaire(
              externalId,
              resolvedExternalId ?? "",
            );
            if (isCancelled) return;
            markQuestionnairePresent(externalId);
            setQuestionnaire(copied ?? next);
            return;
          } catch {
            // Show the PDF found under the previous identity even if copy fails.
          }
        }
        setQuestionnaire(next);
      })
      .catch(() => {
        if (!isCancelled) setQuestionnaire(null);
      });
    void loadPersonDocumentsForRow(
      selectedRow,
      { anketaFullName: personName },
      { nameIsAmbiguous },
    )
      .then((documents) => {
        if (isCancelled) return;
        setDocuments(documents);
        const { document: phoneDocument, phones } =
          extractPhonesFromDocuments(documents);
        if (phoneDocument) setPhoneDocument(externalId, phoneDocument);
        const localPhones = uniqueNormalizedPhones([
          ...(readStoredPersonPhones()[externalId] ?? []),
          ...phones,
        ]);
        if (phones.length) {
          setPhonesByExternalId((current) => {
            const merged = mergedDocumentPhones(current[externalId], phones);
            if (!merged) return current;
            const next = { ...current, [externalId]: merged };
            writeStoredPersonPhones(next);
            return next;
          });
        }
        if (localPhones.length && !phoneDocument) {
          void upsertPersonPhonesDocument(externalId, localPhones, null)
            .then((saved) => {
              if (isCancelled || !saved) return;
              setPhoneDocument(externalId, saved);
            })
            .catch(() => {
              // Local numbers still remain if the backend rejects this document type.
            });
        }
      })
      .catch(() => {
        if (!isCancelled) setDocuments([]);
      });

    return () => {
      isCancelled = true;
    };
  }, [
    canEdit,
    enabled,
    externalId,
    personName,
    personnelRows,
    selectedRow,
  ]);
};
