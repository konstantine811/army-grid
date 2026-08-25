import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { api, type BackendPersonQuestionnaire } from "../../../api";
import { subscribePersonnelAttachmentChanges } from "../../../shared/personnelAttachmentSync";
import {
  buildQuestionnaireExportFileName,
  downloadQuestionnairePdf,
  loadPersonPhotoForRow,
  loadPersonQuestionnaireForRow,
  revokeQuestionnairePreviewUrl,
} from "../../personnel/personnelUtils";
import type { QuestionnairePdfSource } from "../../personnel/questionnaireShare";
import {
  loadPersonnelIndexForAnketa,
  matchAnketaRowToPersonnel,
  type AnketaPersonnelMatch,
} from "../anketaPersonMatch";
import {
  previewAnketaRowPersonnelMerge,
  syncAnketaRowToPersonnel,
} from "../anketaPersonMerge";
import type { AnketaRow } from "../anketaSheet";

export function useAnketaPersonPanel(
  anketaRow: AnketaRow | null,
  onMessage?: (message: string) => void,
) {
  const [match, setMatch] = useState<AnketaPersonnelMatch | null>(null);
  const [matchStatus, setMatchStatus] = useState<"idle" | "loading" | "ready">(
    "idle",
  );
  const [isMerging, setIsMerging] = useState(false);
  const [attachmentExternalId, setAttachmentExternalId] = useState("");
  const [photoData, setPhotoData] = useState("");
  const [questionnaire, setQuestionnaire] =
    useState<BackendPersonQuestionnaire | null>(null);
  const [isLoadingAttachments, setIsLoadingAttachments] = useState(false);
  const [previewOpen, setPreviewOpen] = useState(false);
  const [previewUrl, setPreviewUrl] = useState("");
  const [previewTitle, setPreviewTitle] = useState("");
  const attachmentsEpochRef = useRef(0);
  const prevAnketaRowIdRef = useRef("");
  const shouldAutoOpenPreviewRef = useRef(false);

  const reloadAttachments = useCallback(
    async (personMatch: AnketaPersonnelMatch, row: AnketaRow) => {
      const epoch = ++attachmentsEpochRef.current;
      setIsLoadingAttachments(true);
      try {
        const hints = {
          anketaExternalId: row.externalId,
          anketaFullName: row.fullName,
          anketaBirthDate: row.birthDate,
        };
        const [photoResult, questionnaireResult] = await Promise.all([
          loadPersonPhotoForRow(personMatch.row, hints),
          loadPersonQuestionnaireForRow(personMatch.row, hints),
        ]);
        if (epoch !== attachmentsEpochRef.current) return null;
        setPhotoData(photoResult.photoData);
        setQuestionnaire(questionnaireResult.questionnaire);
        const resolvedExternalId =
          questionnaireResult.resolvedExternalId ||
          photoResult.resolvedExternalId ||
          personMatch.summary.externalId;
        setAttachmentExternalId(resolvedExternalId);
        return {
          questionnaire: questionnaireResult.questionnaire,
          resolvedExternalId,
        };
      } finally {
        if (epoch === attachmentsEpochRef.current) {
          setIsLoadingAttachments(false);
        }
      }
    },
    [],
  );

  const openPreviewForExternalId = useCallback(
    async (externalId: string, name: string, callSign: string) => {
      const fileName = buildQuestionnaireExportFileName(name, callSign);
      try {
        const nextUrl = await api.getPersonQuestionnaireObjectUrl(
          externalId,
          fileName,
        );
        setPreviewTitle(`Анкета · ${name} · ${fileName}`);
        setPreviewUrl((current) => {
          if (current) revokeQuestionnairePreviewUrl(current);
          return nextUrl;
        });
        setPreviewOpen(true);
      } catch {
        /* keep panel usable */
      }
    },
    [],
  );

  const mergePreview = useMemo(
    () =>
      anketaRow && match
        ? previewAnketaRowPersonnelMerge(anketaRow, match)
        : null,
    [anketaRow, match],
  );

  const personnelExternalId =
    attachmentExternalId || match?.summary.externalId || "";
  const displayName =
    anketaRow?.fullName?.trim() || match?.summary.name || "Службовець";
  const exportFileName = useMemo(
    () =>
      buildQuestionnaireExportFileName(
        displayName,
        match?.summary.callSign ?? "",
      ),
    [displayName, match?.summary.callSign],
  );

  const shareSource = useMemo<QuestionnairePdfSource | null>(() => {
    if (!questionnaire?.fileData) return null;
    return { fileData: questionnaire.fileData };
  }, [questionnaire]);

  const mergeToPersonnel = async () => {
    if (!anketaRow || !match) return;
    setIsMerging(true);
    try {
      const result = await syncAnketaRowToPersonnel({ anketaRow, match });
      if (result.skippedReason === "no-row-id") {
        onMessage?.(
          "Немає рядка ООС у БД — лише ранковий список. Спочатку імпортуйте EЖООС.",
        );
        return;
      }
      if (result.skippedReason === "no-updates") {
        onMessage?.("У особовому складі уже є всі дані з цієї анкети.");
        return;
      }
      const fieldCount = Object.keys(result.fieldUpdates).length;
      onMessage?.(
        `Перенесено в особовий склад · ${displayName} · полів: ${fieldCount}${
          result.phonesAdded.length
            ? ` · телефонів: ${result.phonesAdded.length}`
            : ""
        }.`,
      );
    } catch (error) {
      onMessage?.(
        error instanceof Error
          ? error.message
          : "Не вдалося перенести дані в особовий склад.",
      );
    } finally {
      setIsMerging(false);
    }
  };

  useEffect(() => {
    let cancelled = false;
    if (!anketaRow) {
      setMatch(null);
      setMatchStatus("idle");
      prevAnketaRowIdRef.current = "";
      return;
    }
    setMatchStatus("loading");
    void loadPersonnelIndexForAnketa({ force: true })
      .then((index) => {
        if (cancelled) return;
        setMatch(matchAnketaRowToPersonnel(anketaRow, index));
        setMatchStatus("ready");
      })
      .catch(() => {
        if (cancelled) return;
        setMatch(null);
        setMatchStatus("ready");
      });
    return () => {
      cancelled = true;
    };
  }, [anketaRow?.__rowId]);

  useEffect(() => {
    if (!anketaRow) {
      setPhotoData("");
      setQuestionnaire(null);
      setAttachmentExternalId("");
      setPreviewOpen(false);
      setPreviewUrl((current) => {
        if (current) revokeQuestionnairePreviewUrl(current);
        return "";
      });
      return;
    }

    const rowId = anketaRow.__rowId;
    if (prevAnketaRowIdRef.current !== rowId) {
      if (prevAnketaRowIdRef.current) {
        shouldAutoOpenPreviewRef.current = true;
      }
      prevAnketaRowIdRef.current = rowId;
      setPhotoData("");
      setQuestionnaire(null);
      setAttachmentExternalId("");
      setPreviewOpen(false);
      setPreviewUrl((current) => {
        if (current) revokeQuestionnairePreviewUrl(current);
        return "";
      });
    }
  }, [anketaRow?.__rowId]);

  useEffect(() => {
    if (!anketaRow || !match) return;

    void reloadAttachments(match, anketaRow).then((result) => {
      if (
        !shouldAutoOpenPreviewRef.current ||
        !result?.questionnaire ||
        !result.resolvedExternalId
      ) {
        return;
      }
      shouldAutoOpenPreviewRef.current = false;
      void openPreviewForExternalId(
        result.resolvedExternalId,
        anketaRow.fullName?.trim() || match.summary.name,
        match.summary.callSign ?? "",
      );
    });
  }, [anketaRow?.__rowId, match, reloadAttachments, openPreviewForExternalId]);

  useEffect(() => {
    if (!personnelExternalId || !match || !anketaRow) return;

    return subscribePersonnelAttachmentChanges((change) => {
      if (change.externalId !== personnelExternalId) return;
      void reloadAttachments(match, anketaRow);
    });
  }, [anketaRow, match, personnelExternalId, reloadAttachments]);

  useEffect(() => {
    if (!personnelExternalId || !match || !anketaRow) return;

    const refreshOnFocus = () => {
      if (document.visibilityState !== "visible") return;
      void reloadAttachments(match, anketaRow);
    };

    window.addEventListener("focus", refreshOnFocus);
    document.addEventListener("visibilitychange", refreshOnFocus);
    return () => {
      window.removeEventListener("focus", refreshOnFocus);
      document.removeEventListener("visibilitychange", refreshOnFocus);
    };
  }, [anketaRow, match, personnelExternalId, reloadAttachments]);

  useEffect(() => {
    return () => {
      setPreviewUrl((current) => {
        if (current) revokeQuestionnairePreviewUrl(current);
        return "";
      });
    };
  }, []);

  const openQuestionnairePreview = async () => {
    if (!questionnaire || !personnelExternalId) return;
    await openPreviewForExternalId(
      personnelExternalId,
      displayName,
      match?.summary.callSign ?? "",
    );
  };

  const openQuestionnaireTab = async () => {
    if (!questionnaire || !personnelExternalId) return;
    try {
      const url = await api.getPersonQuestionnaireObjectUrl(
        personnelExternalId,
        exportFileName,
      );
      window.open(url, "_blank", "noopener,noreferrer");
      window.setTimeout(() => URL.revokeObjectURL(url), 60_000);
    } catch {
      /* ignore */
    }
  };

  const downloadQuestionnaire = async () => {
    if (!questionnaire || !personnelExternalId) return;
    if (questionnaire.fileData) {
      downloadQuestionnairePdf(exportFileName, {
        fileData: questionnaire.fileData,
      });
      return;
    }
    try {
      const blob = await api.fetchPersonQuestionnaireFile(
        personnelExternalId,
        exportFileName,
        true,
      );
      const url = URL.createObjectURL(blob);
      const anchor = document.createElement("a");
      anchor.href = url;
      anchor.download = exportFileName;
      anchor.click();
      URL.revokeObjectURL(url);
    } catch {
      /* ignore */
    }
  };

  const closePreview = () => {
    setPreviewOpen(false);
    setPreviewUrl((current) => {
      if (current) revokeQuestionnairePreviewUrl(current);
      return "";
    });
  };

  return {
    match,
    matchStatus,
    isMerging,
    photoData,
    questionnaire,
    isLoadingAttachments,
    previewOpen,
    previewUrl,
    previewTitle,
    mergePreview,
    personnelExternalId,
    displayName,
    exportFileName,
    shareSource,
    mergeToPersonnel,
    openQuestionnairePreview,
    openQuestionnaireTab,
    downloadQuestionnaire,
    closePreview,
  };
}
