import { useEffect, useMemo, useRef, useState } from "react";
import {
  api,
  type BackendPersonQuestionnaire,
} from "../../api";
import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import { notifyPersonnelAttachmentChanged } from "../../shared/personnelAttachmentSync";
import { loadPersonQuestionnaireForRow } from "./personAttachments";
import {
  exportPersonnelQuestionnaire,
  questionnaireExportError,
  questionnaireRevealError,
  revealPersonnelQuestionnaire,
} from "./personnelQuestionnaireExport";
import { questionnaireCropFile, downloadVisibleQuestionnaire } from "./personnelQuestionnaireFile";
import {
  isCallsignPersonId,
  loadQuestionnairePreviewObject,
  previewFromQuestionnaireDataUrl,
  questionnaireFileRejection,
  questionnaireOpenError,
  questionnairePreviewHeading,
  questionnairePreviewPersonId,
  questionnaireReviewHeading,
  samePersonNameCount,
} from "./personnelQuestionnaireOpen";
import {
  questionnaireSaveError,
  questionnaireSaveMissingIdMessage,
  questionnaireSavedMessage,
  saveQuestionnaireFile,
} from "./personnelQuestionnaireSave";
import {
  dataUrlToObjectUrl,
  revokeQuestionnairePreviewUrl,
} from "./personnelUtils";
import type { QuestionnairePdfSource } from "./questionnaireShare";

export const usePersonnelQuestionnairePreview = ({
  questionnaire,
  setQuestionnaire,
  selectedRow,
  personName,
  externalId,
  callSign,
  exportFileName,
  personNames,
  setMessage,
  markQuestionnaireInDb,
}: {
  questionnaire: BackendPersonQuestionnaire | null;
  setQuestionnaire: (value: BackendPersonQuestionnaire | null) => void;
  selectedRow: EjournalPreviewRow | null;
  personName: string;
  externalId: string;
  callSign: string;
  exportFileName: string;
  personNames: string[];
  setMessage: (value: string) => void;
  markQuestionnaireInDb: (externalId: string) => void;
}) => {
  const [pendingQuestionnaireFile, setPendingQuestionnaireFile] =
    useState<File | null>(null);
  const [questionnairePreviewUrl, setQuestionnairePreviewUrl] = useState("");
  const [questionnairePreviewTitle, setQuestionnairePreviewTitle] =
    useState("");
  const [isQuestionnairePreviewOpen, setIsQuestionnairePreviewOpen] =
    useState(false);
  const [isDiskFloatingPreview, setIsDiskFloatingPreview] = useState(false);
  const [diskPreviewFile, setDiskPreviewFile] = useState<File | null>(null);
  const [questionnairePreviewFile, setQuestionnairePreviewFile] =
    useState<File | null>(null);
  const [isLoadingQuestionnairePreview, setIsLoadingQuestionnairePreview] =
    useState(false);
  const [isUploadingQuestionnaire, setIsUploadingQuestionnaire] =
    useState(false);
  const questionnairePreviewCacheRef = useRef<{
    externalId: string;
    url: string;
    file: File;
  } | null>(null);
  const questionnairePreviewLoadSeqRef = useRef(0);

  const releaseQuestionnairePreviewCache = (exceptExternalId?: string) => {
    const cached = questionnairePreviewCacheRef.current;
    if (!cached) return;
    if (exceptExternalId && cached.externalId === exceptExternalId) return;
    revokeQuestionnairePreviewUrl(cached.url);
    questionnairePreviewCacheRef.current = null;
  };

  useEffect(() => {
    return () => {
      const cached = questionnairePreviewCacheRef.current;
      if (!cached) return;
      revokeQuestionnairePreviewUrl(cached.url);
      questionnairePreviewCacheRef.current = null;
    };
  }, []);

  const downloadCurrentQuestionnaire = () => {
    downloadVisibleQuestionnaire(exportFileName, {
      pendingFile: pendingQuestionnaireFile,
      diskFile: diskPreviewFile,
      fileData: questionnaire?.fileData,
    });
  };

  const exportCurrentQuestionnaire = async () => {
    try {
      const message = await exportPersonnelQuestionnaire({
        externalId,
        storedPersonId: String(
          questionnaire?.personExternalId || externalId,
        ).trim(),
        fileName: exportFileName,
        fileData: questionnaire?.fileData,
        pendingFile: pendingQuestionnaireFile,
        diskFile: diskPreviewFile,
      });
      if (message) setMessage(message);
    } catch (error) {
      setMessage(questionnaireExportError(error));
    }
  };

  const revealCurrentQuestionnaireInFinder = async () => {
    if (!externalId || !questionnaire) return;
    setMessage("Шукаю оригінал анкети на диску…");
    try {
      setMessage(
        await revealPersonnelQuestionnaire({
          externalId,
          fullName: personName,
          callSign,
          storedFileName: String(questionnaire.fileName ?? ""),
        }),
      );
    } catch (error) {
      setMessage(questionnaireRevealError(error));
    }
  };

  const closeQuestionnairePreview = () => {
    questionnairePreviewLoadSeqRef.current += 1;
    setIsLoadingQuestionnairePreview(false);
    setIsQuestionnairePreviewOpen(false);
    setIsDiskFloatingPreview(false);
    setDiskPreviewFile(null);
    setPendingQuestionnaireFile(null);
    setQuestionnairePreviewTitle("");
    setQuestionnairePreviewUrl("");
    setQuestionnairePreviewFile(null);
  };

  const openDiskQuestionnairePreview = (file: File, title: string) => {
    const nextUrl = URL.createObjectURL(file);
    setPendingQuestionnaireFile(null);
    setQuestionnairePreviewFile(null);
    setDiskPreviewFile(file);
    setQuestionnairePreviewTitle(title);
    setIsDiskFloatingPreview(true);
    setQuestionnairePreviewUrl((current) => {
      if (current) revokeQuestionnairePreviewUrl(current);
      return nextUrl;
    });
    setIsQuestionnairePreviewOpen(true);
    setMessage(title);
  };

  const openQuestionnairePreview = async (
    fileData = questionnaire?.fileData,
  ) => {
    let previewId = questionnairePreviewPersonId(
      questionnaire?.personExternalId,
      externalId,
    );
    const previewTitle = questionnairePreviewHeading(personName, exportFileName);

    if (fileData?.trim()) {
      releaseQuestionnairePreviewCache();
      const preview = previewFromQuestionnaireDataUrl(fileData, exportFileName);
      setPendingQuestionnaireFile(null);
      setDiskPreviewFile(null);
      setQuestionnairePreviewFile(preview.file);
      setIsDiskFloatingPreview(false);
      setQuestionnairePreviewTitle(previewTitle);
      setQuestionnairePreviewUrl(preview.url);
      setIsQuestionnairePreviewOpen(true);
      return;
    }

    if (!previewId || pendingQuestionnaireFile || diskPreviewFile) return;

    if (isCallsignPersonId(previewId)) {
      const found = await loadPersonQuestionnaireForRow(selectedRow, undefined, {
        nameIsAmbiguous: samePersonNameCount(personNames, personName) > 1,
      });
      const resolved = String(
        found.resolvedExternalId || found.questionnaire?.personExternalId || "",
      ).trim();
      if (resolved) previewId = resolved;
      if (found.questionnaire) setQuestionnaire(found.questionnaire);
    }

    const cached = questionnairePreviewCacheRef.current;
    if (cached?.externalId === previewId) {
      setPendingQuestionnaireFile(null);
      setDiskPreviewFile(null);
      setQuestionnairePreviewFile(cached.file);
      setIsDiskFloatingPreview(false);
      setQuestionnairePreviewTitle(previewTitle);
      setQuestionnairePreviewUrl(cached.url);
      setIsQuestionnairePreviewOpen(true);
      return;
    }

    const requestSeq = questionnairePreviewLoadSeqRef.current + 1;
    questionnairePreviewLoadSeqRef.current = requestSeq;
    releaseQuestionnairePreviewCache();
    setPendingQuestionnaireFile(null);
    setDiskPreviewFile(null);
    setQuestionnairePreviewFile(null);
    setIsDiskFloatingPreview(false);
    setQuestionnairePreviewTitle(previewTitle);
    setQuestionnairePreviewUrl("");
    setIsLoadingQuestionnairePreview(true);
    setIsQuestionnairePreviewOpen(true);

    try {
      const preview = await loadQuestionnairePreviewObject(
        previewId,
        exportFileName,
      );
      if (requestSeq !== questionnairePreviewLoadSeqRef.current) return;
      questionnairePreviewCacheRef.current = {
        externalId: previewId,
        url: preview.url,
        file: preview.file,
      };
      setQuestionnairePreviewFile(preview.file);
      setQuestionnairePreviewUrl(preview.url);
    } catch (error) {
      if (requestSeq !== questionnairePreviewLoadSeqRef.current) return;
      setIsQuestionnairePreviewOpen(false);
      setMessage(questionnaireOpenError(error));
    } finally {
      if (requestSeq === questionnairePreviewLoadSeqRef.current) {
        setIsLoadingQuestionnairePreview(false);
      }
    }
  };

  const openQuestionnaireInNewTab = async () => {
    let previewId = questionnairePreviewPersonId(
      questionnaire?.personExternalId,
      externalId,
    );
    if (questionnaire?.fileData) {
      window.open(
        dataUrlToObjectUrl(questionnaire.fileData),
        "_blank",
        "noopener,noreferrer",
      );
      return;
    }
    if (previewId && !pendingQuestionnaireFile && !diskPreviewFile) {
      try {
        if (isCallsignPersonId(previewId)) {
          const found = await loadPersonQuestionnaireForRow(selectedRow);
          const resolved = String(
            found.resolvedExternalId ||
              found.questionnaire?.personExternalId ||
              "",
          ).trim();
          if (resolved) previewId = resolved;
        }
        const url = await api.createPersonQuestionnairePreviewUrl(
          previewId,
          exportFileName,
        );
        window.open(url, "_blank", "noopener,noreferrer");
        return;
      } catch (error) {
        setMessage(questionnaireOpenError(error));
        return;
      }
    }
    downloadCurrentQuestionnaire();
  };

  const beginQuestionnaireReview = (file: File | undefined) => {
    if (!file) return;
    const rejection = questionnaireFileRejection(file, externalId);
    if (rejection) {
      setMessage(rejection);
      return;
    }
    const nextUrl = URL.createObjectURL(file);
    setPendingQuestionnaireFile(file);
    setDiskPreviewFile(null);
    setQuestionnairePreviewFile(null);
    setIsDiskFloatingPreview(false);
    setQuestionnairePreviewTitle(
      questionnaireReviewHeading(personName, file.name),
    );
    setQuestionnairePreviewUrl((current) => {
      if (current) revokeQuestionnairePreviewUrl(current);
      return nextUrl;
    });
    setIsQuestionnairePreviewOpen(true);
  };

  const uploadQuestionnaire = async (file: File) => {
    if (!externalId) {
      setMessage(questionnaireSaveMissingIdMessage);
      return;
    }
    setIsUploadingQuestionnaire(true);
    try {
      const { saved, exportFileName: savedName } = await saveQuestionnaireFile({
        row: selectedRow,
        externalId,
        name: personName,
        callSign,
        file,
      });
      setQuestionnaire(saved);
      markQuestionnaireInDb(externalId);
      notifyPersonnelAttachmentChanged(externalId, "questionnaire");
      releaseQuestionnairePreviewCache();
      setMessage(questionnaireSavedMessage(personName, savedName));
      setPendingQuestionnaireFile(null);
    } catch (error) {
      setMessage(questionnaireSaveError(error));
    } finally {
      setIsUploadingQuestionnaire(false);
    }
  };

  const confirmPendingQuestionnaire = async () => {
    if (!pendingQuestionnaireFile) return;
    await uploadQuestionnaire(pendingQuestionnaireFile);
    closeQuestionnairePreview();
  };

  const shareSource = useMemo((): QuestionnairePdfSource | null => {
    if (pendingQuestionnaireFile) return { file: pendingQuestionnaireFile };
    if (diskPreviewFile) return { file: diskPreviewFile };
    if (questionnairePreviewFile) return { file: questionnairePreviewFile };
    if (questionnaire?.fileData) return { fileData: questionnaire.fileData };
    return null;
  }, [
    diskPreviewFile,
    pendingQuestionnaireFile,
    questionnaire?.fileData,
    questionnairePreviewFile,
  ]);

  const cropFile = useMemo(
    () =>
      questionnaireCropFile({
        diskPreviewFile,
        pendingQuestionnaireFile,
        questionnairePreviewFile,
        fileData: questionnaire?.fileData,
        fileName: questionnaire?.fileName,
      }),
    [
      diskPreviewFile,
      pendingQuestionnaireFile,
      questionnaire?.fileData,
      questionnaire?.fileName,
      questionnairePreviewFile,
    ],
  );

  return {
    pendingQuestionnaireFile,
    diskPreviewFile,
    isQuestionnairePreviewOpen,
    isDiskFloatingPreview,
    isUploadingQuestionnaire,
    isLoadingQuestionnairePreview,
    questionnairePreviewTitle,
    questionnairePreviewUrl,
    shareSource,
    cropFile,
    releaseQuestionnairePreviewCache,
    closeQuestionnairePreview,
    openDiskQuestionnairePreview,
    openQuestionnairePreview,
    openQuestionnaireInNewTab,
    beginQuestionnaireReview,
    exportCurrentQuestionnaire,
    revealCurrentQuestionnaireInFinder,
    confirmPendingQuestionnaire,
  };
};
