import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { api, type BackendPersonQuestionnaire } from "../../../api";
import { subscribePersonnelAttachmentChanges } from "../../../shared/personnelAttachmentSync";
import {
  loadPersonPhotoForRow,
  loadPersonQuestionnaireForRow,
  questionnaireFileMatchesPerson,
} from "../../personnel/personAttachments";
import {
  buildQuestionnaireExportFileName,
  downloadQuestionnairePdf,
  revokeQuestionnairePreviewUrl,
} from "../../personnel/personnelUtils";
import type { QuestionnairePdfSource } from "../../personnel/questionnaireShare";
import {
  loadPersonnelIndexForAnketa,
  matchAnketaRowToPersonnelDetailed,
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
  const anketaRowId = anketaRow?.__rowId ?? "";
  const [match, setMatch] = useState<AnketaPersonnelMatch | null>(null);
  const [ambiguousMatches, setAmbiguousMatches] = useState<AnketaPersonnelMatch[]>(
    [],
  );
  const [matchStatus, setMatchStatus] = useState<"idle" | "loading" | "ready">(
    "idle",
  );
  /** Row id for which `match` was computed — blocks stale match after gap-search jump. */
  const [matchedAnketaRowId, setMatchedAnketaRowId] = useState("");
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
  const previewEpochRef = useRef(0);
  const previewUrlRef = useRef("");
  const prevAnketaRowIdRef = useRef("");
  const shouldAutoOpenPreviewRef = useRef(false);

  // Синхронний скид при зміні особи (до effects) — інакше один кадр лишається
  // match попереднього і авто-PDF відкриває чужу анкету.
  const [panelRowId, setPanelRowId] = useState(anketaRowId);
  if (panelRowId !== anketaRowId) {
    setPanelRowId(anketaRowId);
    setMatch(null);
    setAmbiguousMatches([]);
    setMatchedAnketaRowId("");
    setMatchStatus(anketaRowId ? "loading" : "idle");
    setAttachmentExternalId("");
    setPhotoData("");
    setQuestionnaire(null);
    setIsLoadingAttachments(false);
    setPreviewOpen(false);
    previewEpochRef.current += 1;
    attachmentsEpochRef.current += 1;
    if (previewUrlRef.current) {
      revokeQuestionnairePreviewUrl(previewUrlRef.current);
      previewUrlRef.current = "";
    }
    setPreviewUrl("");
    setPreviewTitle("");
    shouldAutoOpenPreviewRef.current = Boolean(anketaRowId);
    prevAnketaRowIdRef.current = anketaRowId;
  }

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
    async (
      externalId: string,
      name: string,
      callSign: string,
      storedFileName?: string | null,
    ) => {
      const epoch = ++previewEpochRef.current;
      const expectedNames = [name].filter(Boolean);
      if (
        storedFileName &&
        !questionnaireFileMatchesPerson(storedFileName, expectedNames)
      ) {
        return;
      }
      const fileName =
        String(storedFileName ?? "").trim() ||
        buildQuestionnaireExportFileName(name, callSign);
      try {
        const nextUrl = await api.createPersonQuestionnairePreviewUrl(
          externalId,
          fileName,
        );
        if (epoch !== previewEpochRef.current) {
          revokeQuestionnairePreviewUrl(nextUrl);
          return;
        }
        setPreviewTitle(`Анкета · ${name} · ${fileName}`);
        setPreviewUrl((current) => {
          if (current) revokeQuestionnairePreviewUrl(current);
          previewUrlRef.current = nextUrl;
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
  const exportFileName = useMemo(() => {
    const stored = String(questionnaire?.fileName ?? "").trim();
    if (
      stored &&
      questionnaireFileMatchesPerson(stored, [displayName])
    ) {
      return stored;
    }
    return buildQuestionnaireExportFileName(
      displayName,
      match?.summary.callSign ?? "",
    );
  }, [displayName, match?.summary.callSign, questionnaire?.fileName]);

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
      setAmbiguousMatches([]);
      setMatchedAnketaRowId("");
      setMatchStatus("idle");
      return;
    }
    const rowId = anketaRow.__rowId;
    setMatch(null);
    setAmbiguousMatches([]);
    setMatchedAnketaRowId("");
    setMatchStatus("loading");
    void loadPersonnelIndexForAnketa({ force: true })
      .then((index) => {
        if (cancelled) return;
        const result = matchAnketaRowToPersonnelDetailed(anketaRow, index);
        setMatch(result.match);
        setAmbiguousMatches(result.ambiguous);
        setMatchedAnketaRowId(rowId);
        setMatchStatus("ready");
      })
      .catch(() => {
        if (cancelled) return;
        setMatch(null);
        setAmbiguousMatches([]);
        setMatchedAnketaRowId(rowId);
        setMatchStatus("ready");
      });
    return () => {
      cancelled = true;
    };
  }, [anketaRowId]);

  useEffect(() => {
    if (!anketaRow) {
      setPhotoData("");
      setQuestionnaire(null);
      setAttachmentExternalId("");
      setPreviewOpen(false);
      setPreviewUrl((current) => {
        if (current) revokeQuestionnairePreviewUrl(current);
        previewUrlRef.current = "";
        return "";
      });
      shouldAutoOpenPreviewRef.current = false;
      return;
    }

    // Row-change reset already handled during render; keep auto-open flag for first focus.
    if (prevAnketaRowIdRef.current !== anketaRowId) {
      shouldAutoOpenPreviewRef.current = true;
      prevAnketaRowIdRef.current = anketaRowId;
    }
  }, [anketaRow, anketaRowId]);

  useEffect(() => {
    if (!anketaRow || !match || matchStatus !== "ready") return;
    if (matchedAnketaRowId !== anketaRowId) return;

    const rowId = anketaRowId;
    void reloadAttachments(match, anketaRow).then((result) => {
      if (prevAnketaRowIdRef.current !== rowId) return;
      if (matchedAnketaRowId !== rowId) return;
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
        result.questionnaire?.fileName,
      );
    });
  }, [
    anketaRow,
    anketaRowId,
    match,
    matchedAnketaRowId,
    matchStatus,
    reloadAttachments,
    openPreviewForExternalId,
  ]);

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
      previewEpochRef.current += 1;
      if (previewUrlRef.current) {
        revokeQuestionnairePreviewUrl(previewUrlRef.current);
        previewUrlRef.current = "";
      }
    };
  }, []);

  const openQuestionnairePreview = async () => {
    if (!questionnaire || !personnelExternalId) return;
    const stored = String(questionnaire.fileName ?? "").trim();
    if (
      stored &&
      !questionnaireFileMatchesPerson(stored, [displayName])
    ) {
      return;
    }
    await openPreviewForExternalId(
      personnelExternalId,
      displayName,
      match?.summary.callSign ?? "",
      questionnaire.fileName,
    );
  };

  const openQuestionnaireTab = async () => {
    if (!questionnaire || !personnelExternalId) return;
    try {
      const url = await api.createPersonQuestionnairePreviewUrl(
        personnelExternalId,
        exportFileName,
      );
      window.open(url, "_blank", "noopener,noreferrer");
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
    previewEpochRef.current += 1;
    setPreviewOpen(false);
    setPreviewUrl((current) => {
      if (current) revokeQuestionnairePreviewUrl(current);
      previewUrlRef.current = "";
      return "";
    });
  };

  return {
    match,
    ambiguousMatches,
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
