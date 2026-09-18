import { useCallback, useRef, useState } from "react";
import {
  api,
  isApiHttpError,
  type AiQuestionnaireOcrResult,
  type BackendPersonQuestionnaire,
} from "../../../api";
import { visitPdfPagesAsImageDataUrls } from "../../personnel/PhotoCropDialog";
import { dataUrlToFile } from "../../personnel/personnelUtils";
import { prepareOcrImageDataUrl } from "../anketaOcrImage";
import {
  buildAnketaEmptyGapColumnDescriptors,
  buildAnketaOcrProposals,
  buildAnketaOcrRequestHints,
  type AnketaOcrProposal,
} from "../anketaQuestionnaireOcr";
import type { AnketaColumnKey, AnketaRow } from "../anketaSheet";

type RunOcrInput = {
  questionnaire: BackendPersonQuestionnaire | null;
  personnelExternalId: string;
  exportFileName: string;
  anketaRow: AnketaRow;
  gapColumnKeys: readonly AnketaColumnKey[];
  focusedColumnId?: AnketaColumnKey;
  forceRefresh?: boolean;
};

type OcrCacheEntry = {
  key: string;
  fields: Array<{
    key: string;
    label: string;
    value: string;
    confidence: string;
    source?: string;
  }>;
};

let sharedOcrCache: OcrCacheEntry | null = null;

const MAX_OCR_PAGES = 10;
const PDF_DATA_URL_MAX_CHARS = 8_000_000;
const PAGE_IMAGES_MAX_CHARS = 8_000_000;

let supportsPdfData: boolean | null = null;
let supportsPageImages: boolean | null = null;
let supportsOcrColumnHints: boolean | null = null;

export const isFatalOcrError = (error: unknown) =>
  isApiHttpError(error) && (error.status === 401 || error.status === 403);

export const isSkippableOcrPageError = (error: unknown) =>
  !isFatalOcrError(error);

const buildOcrCacheKey = (
  questionnaire: BackendPersonQuestionnaire,
  externalId: string,
  exportFileName: string,
  gapColumnKeys: readonly AnketaColumnKey[],
  rowId: string,
  anketaRow: AnketaRow,
) => {
  const fileName = String(questionnaire.fileName ?? exportFileName).trim();
  const updatedAt = String(questionnaire.updatedAt ?? "").trim();
  const dataSize = questionnaire.fileData?.length ?? 0;
  const columns = [...gapColumnKeys].sort().join(",");
  const emptyColumns = buildAnketaEmptyGapColumnDescriptors(
    anketaRow,
    gapColumnKeys,
  )
    .map((column) => column.key)
    .sort()
    .join(",");
  return `${externalId}|${fileName}|${updatedAt}|${dataSize}|${rowId}|${columns}|empty:${emptyColumns}`;
};

const fileToDataUrl = (file: File) =>
  new Promise<string>((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result ?? ""));
    reader.onerror = () => reject(reader.error);
    reader.readAsDataURL(file);
  });

const questionnaireToFile = async (
  questionnaire: BackendPersonQuestionnaire,
  fileName: string,
  externalId: string,
) => {
  if (questionnaire.fileData) {
    return dataUrlToFile(questionnaire.fileData, fileName);
  }
  const blob = await api.fetchPersonQuestionnaireFile(externalId, fileName, true);
  return new File([blob], fileName, {
    type: blob.type || "application/pdf",
  });
};

const questionnaireToPdfDataUrl = async (
  questionnaire: BackendPersonQuestionnaire,
  fileName: string,
  externalId: string,
) => {
  const fileData = String(questionnaire.fileData ?? "").trim();
  if (fileData.startsWith("data:application/pdf")) return fileData;
  const file = await questionnaireToFile(questionnaire, fileName, externalId);
  if (
    file.type === "application/pdf" ||
    file.name.toLowerCase().endsWith(".pdf") ||
    fileData.startsWith("data:application/pdf")
  ) {
    return fileToDataUrl(file);
  }
  return "";
};

const questionnaireToPageImages = async (
  questionnaire: BackendPersonQuestionnaire,
  fileName: string,
  externalId: string,
  onProgress?: (message: string) => void,
) => {
  const file = await questionnaireToFile(questionnaire, fileName, externalId);
  const images: string[] = [];
  await visitPdfPagesAsImageDataUrls(
    file,
    async (page, pageCount) => {
      onProgress?.(`Готую сторінку ${page.pageNumber}/${pageCount}…`);
      images.push(await prepareOcrImageDataUrl(page.src));
    },
    { scale: 1.5 },
  );
  return images;
};

export function useAnketaQuestionnaireOcr(onMessage?: (message: string) => void) {
  const [isRunning, setIsRunning] = useState(false);
  const [progress, setProgress] = useState("");
  const [reviewOpen, setReviewOpen] = useState(false);
  const [proposals, setProposals] = useState<AnketaOcrProposal[]>([]);
  const [focusedColumnId, setFocusedColumnId] = useState<AnketaColumnKey | null>(
    null,
  );
  const runEpochRef = useRef(0);

  const closeReview = useCallback(() => {
    setReviewOpen(false);
    setProposals([]);
    setFocusedColumnId(null);
    setProgress("");
  }, []);

  const requestGapFill = useCallback(
    async (
      questionnaire: BackendPersonQuestionnaire,
      exportFileName: string,
      externalId: string,
      anketaRow: AnketaRow,
      gapColumnKeys: readonly AnketaColumnKey[],
      cacheKey: string,
      setProgressMessage: (message: string) => void,
    ) => {
      if (sharedOcrCache?.key === cacheKey) {
        return sharedOcrCache.fields;
      }

      const emptyColumns = buildAnketaEmptyGapColumnDescriptors(
        anketaRow,
        gapColumnKeys,
      );
      if (!emptyColumns.length) {
        throw new Error("Немає порожніх обраних полів для заповнення.");
      }

      const ocrHints = buildAnketaOcrRequestHints(anketaRow, gapColumnKeys);

      const sendOcr = async (
        payload: Parameters<typeof api.aiQuestionnaireOcr>[0],
      ) => {
        const withColumns = supportsOcrColumnHints !== false;
        const body: Parameters<typeof api.aiQuestionnaireOcr>[0] = {
          ...payload,
          personName: ocrHints.personName || undefined,
          ...(withColumns
            ? {
                selectedColumns: ocrHints.selectedColumns,
                emptyColumns: ocrHints.emptyColumns,
              }
            : {}),
        };
        try {
          return await api.aiQuestionnaireOcr(body);
        } catch (error) {
          if (
            withColumns &&
            isApiHttpError(error) &&
            error.status === 400
          ) {
            supportsOcrColumnHints = false;
            return sendOcr(payload);
          }
          throw error;
        }
      };

      const tryOnce = async (
        payload: Parameters<typeof api.aiQuestionnaireOcr>[0],
      ): Promise<AiQuestionnaireOcrResult | null> => {
        try {
          const next = await sendOcr(payload);
          const hasValue = next.fields.some((field) =>
            String(field.value ?? "").trim(),
          );
          return hasValue ? next : null;
        } catch (error) {
          if (isFatalOcrError(error)) throw error;
          if (
            isApiHttpError(error) &&
            (error.status === 400 || error.status === 413)
          ) {
            return null;
          }
          throw error;
        }
      };

      setProgressMessage("Завантажую PDF…");
      const pdfData = await questionnaireToPdfDataUrl(
        questionnaire,
        exportFileName,
        externalId,
      );
      const canSendPdf =
        Boolean(pdfData) &&
        pdfData.length <= PDF_DATA_URL_MAX_CHARS &&
        supportsPdfData !== false;

      let result: AiQuestionnaireOcrResult | null = null;
      if (canSendPdf) {
        setProgressMessage("Відправляю анкету в Gemini…");
        result = await tryOnce({
          pdfData,
          fileName: exportFileName,
        });
        if (result) supportsPdfData = true;
      }

      let pageImages: string[] = [];
      if (!result) {
        setProgressMessage("Готую сторінки анкети…");
        pageImages = (
          await questionnaireToPageImages(
            questionnaire,
            exportFileName,
            externalId,
            setProgressMessage,
          )
        ).slice(0, MAX_OCR_PAGES);
        if (!pageImages.length) {
          throw new Error("Немає сторінок анкети для розпізнавання.");
        }
        const imageData = pageImages[0]!;
        const pagesChars = pageImages.reduce(
          (sum, image) => sum + image.length,
          0,
        );

        if (
          Boolean(pdfData) &&
          pdfData.length <= PDF_DATA_URL_MAX_CHARS &&
          supportsPdfData !== false
        ) {
          setProgressMessage("Відправляю PDF в Gemini…");
          result = await tryOnce({
            imageData,
            pdfData,
            fileName: exportFileName,
          });
          supportsPdfData = Boolean(result);
        }

        if (
          !result &&
          supportsPageImages !== false &&
          pagesChars <= PAGE_IMAGES_MAX_CHARS
        ) {
          setProgressMessage("Відправляю сторінки в Gemini…");
          result = await tryOnce({
            imageData,
            pageImages,
            fileName: exportFileName,
          });
          supportsPageImages = Boolean(result);
        }

        if (!result) {
          const fields: AiQuestionnaireOcrResult["fields"] = [];
          let lastPageError: unknown = null;
          let failedPages = 0;
          for (let index = 0; index < pageImages.length; index += 1) {
            const pageNumber = String(index + 1);
            setProgressMessage(
              `Відправляю сторінку ${pageNumber}/${pageImages.length} в Gemini…`,
            );
            try {
              const pageResult = await sendOcr({
                imageData: pageImages[index]!,
                fileName: exportFileName,
                pageNumber,
              });
              fields.push(...pageResult.fields);
            } catch (error) {
              if (isFatalOcrError(error)) throw error;
              lastPageError = error;
              failedPages += 1;
              setProgressMessage(
                `Сторінка ${pageNumber} не розпізналась — переходжу далі…`,
              );
            }
          }
          if (
            !fields.length &&
            lastPageError &&
            failedPages === pageImages.length
          ) {
            throw lastPageError;
          }
          result = { model: "gemini", fileName: exportFileName, fields };
        }
      }

      const mappedFields = result.fields.map((field) => ({
        key: field.key,
        label: field.label,
        value: field.value,
        confidence: field.confidence,
        source: String((field as { source?: string }).source ?? "").trim(),
      }));
      if (mappedFields.some((field) => field.value.trim())) {
        sharedOcrCache = { key: cacheKey, fields: mappedFields };
      }
      return mappedFields;
    },
    [],
  );

  const runOcr = useCallback(
    async (input: RunOcrInput) => {
      const {
        questionnaire,
        personnelExternalId,
        exportFileName,
        anketaRow,
        gapColumnKeys,
        focusedColumnId: nextFocusedColumnId,
        forceRefresh,
      } = input;
      if (!questionnaire) {
        onMessage?.("Спочатку додайте PDF-анкету до службовця.");
        return;
      }
      const externalId =
        questionnaire.personExternalId?.trim() || personnelExternalId.trim();
      if (!externalId) {
        onMessage?.("Немає ID службовця для завантаження анкети.");
        return;
      }
      if (!gapColumnKeys.length) {
        onMessage?.("Оберіть колонки для пошуку пропусків.");
        return;
      }

      const epoch = ++runEpochRef.current;
      setIsRunning(true);
      setFocusedColumnId(nextFocusedColumnId ?? null);
      if (forceRefresh) sharedOcrCache = null;
      const cacheKey = buildOcrCacheKey(
        questionnaire,
        externalId,
        exportFileName,
        gapColumnKeys,
        anketaRow.__rowId,
        anketaRow,
      );
      const cacheHit = sharedOcrCache?.key === cacheKey;
      setProgress(cacheHit ? "Готую пропозиції…" : "Готую запит до Gemini…");
      try {
        const mergedFields = await requestGapFill(
          questionnaire,
          exportFileName,
          externalId,
          anketaRow,
          gapColumnKeys,
          cacheKey,
          (message) => {
            if (epoch === runEpochRef.current) setProgress(message);
          },
        );
        if (epoch !== runEpochRef.current) return;

        const nextProposals = buildAnketaOcrProposals(
          mergedFields,
          anketaRow,
          gapColumnKeys,
          { focusedColumnId: nextFocusedColumnId },
        );
        if (!nextProposals.length) {
          onMessage?.(
            nextFocusedColumnId
              ? "Gemini не знайшов значення для цієї порожньої комірки в анкеті."
              : "Gemini не знайшов значень для порожніх обраних комірок.",
          );
          return;
        }
        const foundCount = nextProposals.filter(
          (item) => item.source !== "за замовчуванням",
        ).length;
        setProposals(nextProposals);
        setReviewOpen(true);
        if (foundCount === 0) {
          onMessage?.(
            "Gemini не зчитав поля з анкети. Перевірте PDF або натисніть «Повторити Gemini».",
          );
        } else {
          const focusedProposal = nextFocusedColumnId
            ? nextProposals.find((item) => item.columnId === nextFocusedColumnId)
            : null;
          onMessage?.(
            focusedProposal?.source && focusedProposal.source !== "за замовчуванням"
              ? `Gemini: ${focusedProposal.label} — перевірте та підтвердьте.`
              : `Gemini запропонував ${foundCount} полів — перевірте та підтвердьте.`,
          );
        }
      } catch (error) {
        const detail =
          error instanceof Error ? error.message : "Розпізнавання не вдалося.";
        const hint = /gemini_api_key/i.test(detail)
          ? " Додайте GEMINI_API_KEY у army-backend/.env і перезапустіть backend."
          : "";
        onMessage?.(`Розпізнавання не вдалося: ${detail}${hint}`);
      } finally {
        if (epoch === runEpochRef.current) {
          setIsRunning(false);
          setProgress("");
        }
      }
    },
    [onMessage, requestGapFill],
  );

  const toggleProposal = useCallback((columnId: AnketaColumnKey) => {
    setProposals((current) =>
      current.map((item) =>
        item.columnId === columnId
          ? { ...item, selected: !item.selected }
          : item,
      ),
    );
  }, []);

  const updateProposalValue = useCallback(
    (columnId: AnketaColumnKey, value: string) => {
      setProposals((current) =>
        current.map((item) =>
          item.columnId === columnId ? { ...item, value } : item,
        ),
      );
    },
    [],
  );

  const selectAllProposals = useCallback((selected: boolean) => {
    setProposals((current) => current.map((item) => ({ ...item, selected })));
  }, []);

  return {
    isRunning,
    progress,
    reviewOpen,
    proposals,
    focusedColumnId,
    runOcr,
    closeReview,
    toggleProposal,
    updateProposalValue,
    selectAllProposals,
  };
}
