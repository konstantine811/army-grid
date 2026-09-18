import { useCallback, useState } from "react";
import { api } from "../../../api";
import { prepareOcrImageDataUrl } from "../../anketa-data/anketaOcrImage";
import { visitPdfPagesAsImageDataUrls } from "../../personnel/PhotoCropDialog";
import { compressPhotoFile } from "../../personnel/photoCompression";
import {
  applyPreAnketaOcrProposals,
  buildPreAnketaOcrFieldHints,
  buildPreAnketaOcrProposals,
  type PreAnketaOcrProposal,
} from "../preAnketaOcr";
import type { PreAnketaFormFields } from "../preAnketaFields";

const MAX_SCANS = 12;
const MAX_OCR_PAGES = 10;
const MAX_TOTAL_CHARS = 8_000_000;
const PDF_DATA_URL_MAX_CHARS = 8_000_000;

const readFileAsDataUrl = (file: File) =>
  new Promise<string>((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result ?? ""));
    reader.onerror = () => reject(new Error(`Не вдалося прочитати ${file.name}.`));
    reader.readAsDataURL(file);
  });

const isPdfFile = (file: File) =>
  file.type === "application/pdf" || file.name.toLowerCase().endsWith(".pdf");

const isImageFile = (file: File) => file.type.startsWith("image/");

export type PreAnketaUploadedScan = {
  id: string;
  name: string;
  kind: "image" | "pdf";
  previewUrl: string;
  file?: File;
};

export const usePreAnketaOcr = () => {
  const [scans, setScans] = useState<PreAnketaUploadedScan[]>([]);
  const [isRunning, setIsRunning] = useState(false);
  const [progressMessage, setProgressMessage] = useState("");
  const [error, setError] = useState("");
  const [proposals, setProposals] = useState<PreAnketaOcrProposal[]>([]);

  const addScans = useCallback(async (files: FileList | File[]) => {
    setError("");
    const list = [...files].filter((file) => isImageFile(file) || isPdfFile(file));
    if (!list.length) {
      setError("Оберіть фото (JPG, PNG) або PDF.");
      return;
    }

    const next: PreAnketaUploadedScan[] = [];
    for (const file of list.slice(0, MAX_SCANS - scans.length)) {
      if (isPdfFile(file)) {
        next.push({
          id: `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
          name: file.name,
          kind: "pdf",
          previewUrl: "",
          file,
        });
        continue;
      }
      const compressed = await compressPhotoFile(file);
      const previewUrl = await readFileAsDataUrl(compressed);
      next.push({
        id: `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
        name: file.name,
        kind: "image",
        previewUrl,
      });
    }
    setScans((current) => [...current, ...next].slice(0, MAX_SCANS));
  }, [scans.length]);

  const removeScan = useCallback((id: string) => {
    setScans((current) => current.filter((scan) => scan.id !== id));
  }, []);

  const clearReview = useCallback(() => {
    setProposals([]);
  }, []);

  const runOcr = useCallback(
    async (currentFields: PreAnketaFormFields) => {
      if (!scans.length) {
        setError("Спочатку додайте фото або PDF документів.");
        return;
      }
      setIsRunning(true);
      setError("");
      setProgressMessage("Готую зображення…");
      try {
        const pageImages: string[] = [];
        let pdfData = "";

        const pdfScans = scans.filter((scan) => scan.kind === "pdf" && scan.file);
        if (pdfScans.length === 1 && scans.length === 1 && pdfScans[0]?.file) {
          const dataUrl = await readFileAsDataUrl(pdfScans[0].file);
          if (dataUrl.length <= PDF_DATA_URL_MAX_CHARS) {
            pdfData = dataUrl;
          }
        }

        for (const [index, scan] of scans.entries()) {
          if (pageImages.length >= MAX_OCR_PAGES) break;

          if (scan.kind === "image") {
            setProgressMessage(`Готую фото ${index + 1}/${scans.length}…`);
            pageImages.push(await prepareOcrImageDataUrl(scan.previewUrl));
            continue;
          }

          if (!scan.file) continue;
          setProgressMessage(`Готую PDF ${scan.name}…`);
          await visitPdfPagesAsImageDataUrls(
            scan.file,
            async (page, pageCount) => {
              if (pageImages.length >= MAX_OCR_PAGES) return;
              setProgressMessage(
                `PDF ${scan.name}: сторінка ${page.pageNumber}/${pageCount}…`,
              );
              pageImages.push(await prepareOcrImageDataUrl(page.src));
            },
            { scale: 1.5 },
          );
        }

        if (!pageImages.length && !pdfData) {
          throw new Error("Не вдалося підготувати жодної сторінки для OCR.");
        }

        const totalChars =
          pageImages.reduce((sum, image) => sum + image.length, 0) +
          pdfData.length;
        if (totalChars > MAX_TOTAL_CHARS) {
          throw new Error("Занадто великий набір файлів. Спробуйте менше документів.");
        }

        setProgressMessage("Розпізнаю Gemini…");
        const hints = buildPreAnketaOcrFieldHints();
        const result = await api.aiQuestionnaireOcr({
          ...(pdfData ? { pdfData } : {}),
          ...(pageImages.length ? { pageImages } : {}),
          fileName: "pre-anketa-scans",
          personName: currentFields.fullName.trim() || undefined,
          selectedColumns: hints,
          emptyColumns: hints.filter(
            (column) =>
              !String(
                currentFields[column.key as keyof PreAnketaFormFields] ?? "",
              ).trim(),
          ),
        });

        const nextProposals = buildPreAnketaOcrProposals(
          result.fields ?? [],
          currentFields,
        );
        if (!nextProposals.length) {
          setProgressMessage("Gemini не знайшов нових полів для заповнення.");
          setProposals([]);
          return;
        }
        setProposals(nextProposals);
        setProgressMessage(
          `Gemini знайшов ${nextProposals.length} полів · перевірте та застосуйте.`,
        );
      } catch (cause) {
        setError(
          cause instanceof Error
            ? cause.message
            : "Не вдалося розпізнати документи.",
        );
        setProgressMessage("");
      } finally {
        setIsRunning(false);
      }
    },
    [scans],
  );

  const applyProposals = useCallback(
    (currentFields: PreAnketaFormFields, selected: PreAnketaOcrProposal[]) =>
      applyPreAnketaOcrProposals(currentFields, selected),
    [],
  );

  return {
    scans,
    addScans,
    removeScan,
    runOcr,
    applyProposals,
    proposals,
    setProposals,
    clearReview,
    isRunning,
    progressMessage,
    error,
    setError,
  };
};
