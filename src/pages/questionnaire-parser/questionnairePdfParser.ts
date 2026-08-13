import * as pdfjs from "pdfjs-dist";
import pdfWorkerUrl from "pdfjs-dist/build/pdf.worker.min.mjs?url";
import { createWorker, PSM } from "tesseract.js";
import type {
  PDFDocumentProxy,
  PDFPageProxy,
} from "pdfjs-dist/types/src/pdf";

pdfjs.GlobalWorkerOptions.workerSrc = pdfWorkerUrl;

type PdfTextItem = {
  str?: string;
  transform?: number[];
  width?: number;
  height?: number;
};

export type QuestionnaireTextPage = {
  pageNumber: number;
  text: string;
  rows: string[];
  imageSrc: string;
  fieldImages?: Record<string, string>;
  fieldValues?: Record<string, string>;
  itemCount: number;
  score: number;
};

export type ParsedQuestionnaireField = {
  key: string;
  label: string;
  value: string;
  confidence: "high" | "medium" | "low";
  sourceImage?: string;
};

export type QuestionnaireParseResult = {
  fileName: string;
  pageCount: number;
  textPageCount: number;
  usedOcr: boolean;
  selectedPageNumber: number | null;
  selectedText: string;
  fields: ParsedQuestionnaireField[];
  pages: QuestionnaireTextPage[];
};

export type QuestionnaireParseProgress = {
  stage: "text" | "ocr" | "done";
  pageNumber?: number;
  pageCount?: number;
  progress?: number;
  message: string;
};

type QuestionnaireParseOptions = {
  useOcr?: boolean;
  pageNumbers?: number[];
  onProgress?: (progress: QuestionnaireParseProgress) => void;
};

const QUESTIONNAIRE_MARKERS = [
  "анкета",
  "прізвище",
  "ім'я",
  "по батькові",
  "військове звання",
  "ідентифікаційний",
  "дата народження",
  "місце народження",
  "адреса",
  "контактні телефони",
  "освіта",
];

const PUBLIC_BASE = import.meta.env.BASE_URL || "/";
const TESSERACT_BASE = `${PUBLIC_BASE}tesseract`;
const TESSDATA_BASE = `${PUBLIC_BASE}tessdata`;
const OCR_LANGUAGES = "ukr+eng+rus";
const TEXT_LAYER_RENDER_SCALE = 1.35;
const OCR_RENDER_SCALE = 2.7;

const FIELD_DEFINITIONS: Array<{
  key: string;
  label: string;
  aliases: string[];
}> = [
  { key: "surname", label: "Прізвище", aliases: ["прізвище"] },
  { key: "firstName", label: "Ім'я", aliases: ["ім'я", "ім я"] },
  { key: "patronymic", label: "По батькові", aliases: ["по батькові"] },
  { key: "rank", label: "Військове звання", aliases: ["військове звання"] },
  {
    key: "rnokpp",
    label: "Ідентифікаційний номер",
    aliases: ["ідентифікаційний", "ідентифікаційний код", "код номер"],
  },
  { key: "passport", label: "Паспорт", aliases: ["паспорт"] },
  {
    key: "birthDate",
    label: "Дата народження",
    aliases: ["дата народження"],
  },
  {
    key: "birthPlace",
    label: "Місце народження",
    aliases: ["місце народження"],
  },
  { key: "registrationAddress", label: "Адреса прописки", aliases: ["адреса прописка"] },
  {
    key: "actualAddress",
    label: "Адреса фактичного проживання",
    aliases: ["адреса фактичного проживання"],
  },
  {
    key: "phones",
    label: "Контактні телефони",
    aliases: ["контактні телефони", "власні контактні телефони"],
  },
  { key: "education", label: "Освіта", aliases: ["освіта"] },
  { key: "work", label: "Основне місце роботи", aliases: ["основне місце роботи"] },
  {
    key: "militaryId",
    label: "Посвідчення / військовий квиток",
    aliases: ["посвідчення убд", "військовий квиток"],
  },
  {
    key: "familyStatus",
    label: "Сімейний стан",
    aliases: ["сімейний стан"],
  },
];

const FORM_FIELD_ZONES: Record<
  string,
  { x: number; y: number; width: number; height: number; mode: PSM }
> = {
  surname: { x: 0.42, y: 0.018, width: 0.56, height: 0.035, mode: PSM.SINGLE_LINE },
  firstName: { x: 0.42, y: 0.053, width: 0.56, height: 0.035, mode: PSM.SINGLE_LINE },
  patronymic: { x: 0.42, y: 0.088, width: 0.56, height: 0.041, mode: PSM.SINGLE_LINE },
  rank: { x: 0.42, y: 0.129, width: 0.56, height: 0.045, mode: PSM.SINGLE_LINE },
  rnokpp: { x: 0.42, y: 0.174, width: 0.56, height: 0.04, mode: PSM.SINGLE_LINE },
  passport: { x: 0.42, y: 0.214, width: 0.56, height: 0.087, mode: PSM.SINGLE_BLOCK },
  birthDate: { x: 0.42, y: 0.345, width: 0.56, height: 0.039, mode: PSM.SINGLE_LINE },
  birthPlace: { x: 0.42, y: 0.384, width: 0.56, height: 0.058, mode: PSM.SINGLE_BLOCK },
  registrationAddress: { x: 0.42, y: 0.442, width: 0.56, height: 0.058, mode: PSM.SINGLE_BLOCK },
  actualAddress: { x: 0.42, y: 0.5, width: 0.56, height: 0.087, mode: PSM.SINGLE_BLOCK },
  phones: { x: 0.42, y: 0.587, width: 0.56, height: 0.043, mode: PSM.SINGLE_LINE },
  education: { x: 0.42, y: 0.63, width: 0.56, height: 0.076, mode: PSM.SINGLE_BLOCK },
  work: { x: 0.42, y: 0.706, width: 0.56, height: 0.064, mode: PSM.SINGLE_BLOCK },
};

const normalizeText = (value: string) =>
  String(value ?? "")
    .normalize("NFC")
    .toLowerCase()
    .replace(/ё/g, "е")
    .replace(/[’ʼ`´]/g, "'")
    .replace(/\s+/g, " ")
    .trim();

const normalizeForSearch = (value: string) =>
  normalizeText(value)
    .replace(/[^a-zа-яіїєґ0-9'+\s]/gi, " ")
    .replace(/\s+/g, " ")
    .trim();

const cleanupValue = (value: string) =>
  value
    .replace(/^[\s:;.,\-–—_□☐|/\\]+/, "")
    .replace(/[\s:;.,\-–—_□☐|/\\]+$/, "")
    .replace(/\s+/g, " ")
    .trim();

const normalizeLikelyUkrainianOcr = (value: string) =>
  value
    .replace(/[ІI1]\s*['’ʼ`´]\s*/g, "Ї")
    .replace(/[iI]\b/g, "і")
    .replace(/\b[iI]/g, "і")
    .replace(/([А-Яа-яІіЇїЄєҐґ])c([А-Яа-яІіЇїЄєҐґ])/g, "$1с$2")
    .replace(/([А-Яа-яІіЇїЄєҐґ])a([А-Яа-яІіЇїЄєҐґ])/g, "$1а$2")
    .replace(/([А-Яа-яІіЇїЄєҐґ])e([А-Яа-яІіЇїЄєҐґ])/g, "$1е$2")
    .replace(/([А-Яа-яІіЇїЄєҐґ])o([А-Яа-яІіЇїЄєҐґ])/g, "$1о$2")
    .replace(/([А-Яа-яІіЇїЄєҐґ])p([А-Яа-яІіЇїЄєҐґ])/g, "$1р$2")
    .replace(/([А-Яа-яІіЇїЄєҐґ])x([А-Яа-яІіЇїЄєҐґ])/g, "$1х$2")
    .replace(/\s+/g, " ")
    .trim();

const scoreQuestionnairePage = (text: string) => {
  const normalized = normalizeForSearch(text);
  return QUESTIONNAIRE_MARKERS.reduce(
    (score, marker) => score + (normalized.includes(marker) ? 1 : 0),
    0,
  );
};

const groupTextRows = (items: PdfTextItem[]) => {
  const positioned = items
    .map((item) => ({
      text: String(item.str ?? "").trim(),
      x: item.transform?.[4] ?? 0,
      y: item.transform?.[5] ?? 0,
    }))
    .filter((item) => item.text);

  const lines: Array<{ y: number; items: Array<{ text: string; x: number }> }> = [];
  for (const item of positioned.sort((a, b) => b.y - a.y || a.x - b.x)) {
    const line = lines.find((candidate) => Math.abs(candidate.y - item.y) <= 3);
    if (line) {
      line.items.push({ text: item.text, x: item.x });
      line.y = (line.y + item.y) / 2;
    } else {
      lines.push({ y: item.y, items: [{ text: item.text, x: item.x }] });
    }
  }

  return lines
    .sort((a, b) => b.y - a.y)
    .map((line) =>
      line.items
        .sort((a, b) => a.x - b.x)
        .map((item) => item.text)
        .join(" ")
        .replace(/\s+/g, " ")
        .trim(),
    )
    .filter(Boolean);
};

const findAliasIndex = (text: string, aliases: string[]) => {
  const normalized = normalizeForSearch(text);
  let best = -1;
  for (const alias of aliases) {
    const index = normalized.indexOf(normalizeForSearch(alias));
    if (index >= 0 && (best < 0 || index < best)) best = index;
  }
  return best;
};

const extractFieldValue = (
  text: string,
  rows: string[],
  field: (typeof FIELD_DEFINITIONS)[number],
) => {
  const textIndex = findAliasIndex(text, field.aliases);
  if (textIndex < 0) return "";

  const normalizedText = normalizeForSearch(text);
  const nextFieldIndexes = FIELD_DEFINITIONS.flatMap((candidate) =>
    candidate.key === field.key
      ? []
      : candidate.aliases
          .map((alias) => normalizedText.indexOf(normalizeForSearch(alias), textIndex + 1))
          .filter((index) => index > textIndex),
  );
  const endIndex = nextFieldIndexes.length
    ? Math.min(...nextFieldIndexes)
    : Math.min(normalizedText.length, textIndex + 220);
  const fragment = cleanupValue(normalizedText.slice(textIndex, endIndex));
  const alias = field.aliases
    .map(normalizeForSearch)
    .find((item) => fragment.startsWith(item));
  const inlineValue = alias ? cleanupValue(fragment.slice(alias.length)) : "";
  if (inlineValue && inlineValue.length <= 140) return inlineValue;

  const rowIndex = rows.findIndex((row) => findAliasIndex(row, field.aliases) >= 0);
  if (rowIndex < 0) return inlineValue;

  const sameRow = rows[rowIndex];
  const separatorValue = cleanupValue(sameRow.split(/[:：]/).slice(1).join(":"));
  if (separatorValue && separatorValue.length <= 140) return separatorValue;

  for (let offset = 1; offset <= 2; offset += 1) {
    const candidate = cleanupValue(rows[rowIndex + offset] ?? "");
    if (
      candidate &&
      !FIELD_DEFINITIONS.some((definition) =>
        definition.aliases.some((aliasItem) =>
          normalizeForSearch(candidate).includes(normalizeForSearch(aliasItem)),
        ),
      )
    ) {
      return candidate;
    }
  }

  return inlineValue;
};

export const extractQuestionnaireFields = (page: QuestionnaireTextPage | undefined) => {
  if (!page) return [];
  return FIELD_DEFINITIONS.map((field): ParsedQuestionnaireField => {
    const zoneValue = cleanupValue(page.fieldValues?.[field.key] ?? "");
    const value = zoneValue || extractFieldValue(page.text, page.rows, field);
    return {
      key: field.key,
      label: field.label,
      value,
      confidence: zoneValue ? "high" : value ? "medium" : "low",
      sourceImage: page.fieldImages?.[field.key],
    };
  });
};

const sharpenImage = (
  data: Uint8ClampedArray,
  width: number,
  height: number,
) => {
  const source = new Uint8ClampedArray(data);
  const kernel = [0, -1, 0, -1, 5, -1, 0, -1, 0];
  for (let y = 1; y < height - 1; y += 1) {
    for (let x = 1; x < width - 1; x += 1) {
      let value = 0;
      for (let ky = -1; ky <= 1; ky += 1) {
        for (let kx = -1; kx <= 1; kx += 1) {
          const sourceIndex = ((y + ky) * width + x + kx) * 4;
          value += source[sourceIndex] * kernel[(ky + 1) * 3 + kx + 1];
        }
      }
      const index = (y * width + x) * 4;
      const next = Math.max(0, Math.min(255, value));
      data[index] = next;
      data[index + 1] = next;
      data[index + 2] = next;
    }
  }
};

const enhanceCanvasForOcr = (canvas: HTMLCanvasElement) => {
  const context = canvas.getContext("2d");
  if (!context) throw new Error("Не вдалося підготувати canvas для OCR.");

  const image = context.getImageData(0, 0, canvas.width, canvas.height);
  const pixels = image.data;

  for (let index = 0; index < pixels.length; index += 4) {
    const gray =
      pixels[index] * 0.299 + pixels[index + 1] * 0.587 + pixels[index + 2] * 0.114;
    const contrasted = Math.max(0, Math.min(255, (gray - 128) * 1.58 + 128));
    const boosted = contrasted > 188 ? 255 : contrasted < 118 ? 0 : contrasted;
    pixels[index] = boosted;
    pixels[index + 1] = boosted;
    pixels[index + 2] = boosted;
  }

  sharpenImage(pixels, canvas.width, canvas.height);
  context.putImageData(image, 0, 0);
  return canvas.toDataURL("image/png");
};

const renderPageToImageVariants = async (page: PDFPageProxy) => {
  const viewport = page.getViewport({ scale: OCR_RENDER_SCALE });
  const canvas = document.createElement("canvas");
  canvas.width = Math.max(1, Math.floor(viewport.width));
  canvas.height = Math.max(1, Math.floor(viewport.height));
  const context = canvas.getContext("2d");
  if (!context) throw new Error("Не вдалося підготувати canvas для OCR.");

  context.fillStyle = "#ffffff";
  context.fillRect(0, 0, canvas.width, canvas.height);
  await page.render({
    canvasContext: context,
    background: "#ffffff",
    viewport,
  }).promise;

  const raw = canvas.toDataURL("image/png");
  const enhanced = enhanceCanvasForOcr(canvas);
  return [
    { label: "оригінал", src: raw },
    { label: "покращене зображення", src: enhanced },
  ];
};

const loadDataUrlImage = (src: string) =>
  new Promise<HTMLImageElement>((resolve, reject) => {
    const image = new Image();
    image.onload = () => resolve(image);
    image.onerror = () => reject(new Error("Не вдалося прочитати crop поля."));
    image.src = src;
  });

type ImageBounds = {
  x: number;
  y: number;
  width: number;
  height: number;
};

const detectFormTableBounds = async (pageImageSrc: string): Promise<ImageBounds | null> => {
  const image = await loadDataUrlImage(pageImageSrc);
  const canvas = document.createElement("canvas");
  canvas.width = image.naturalWidth;
  canvas.height = image.naturalHeight;
  const context = canvas.getContext("2d");
  if (!context) return null;

  context.drawImage(image, 0, 0);
  const { data } = context.getImageData(0, 0, canvas.width, canvas.height);
  const rowCounts = new Array(canvas.height).fill(0) as number[];
  const colCounts = new Array(canvas.width).fill(0) as number[];
  const startY = Math.floor(canvas.height * 0.04);
  const endY = Math.floor(canvas.height * 0.86);
  const startX = Math.floor(canvas.width * 0.05);
  const endX = Math.floor(canvas.width * 0.98);

  for (let y = startY; y < endY; y += 1) {
    for (let x = startX; x < endX; x += 1) {
      const index = (y * canvas.width + x) * 4;
      const gray = data[index] * 0.299 + data[index + 1] * 0.587 + data[index + 2] * 0.114;
      if (gray < 92) {
        rowCounts[y] += 1;
        colCounts[x] += 1;
      }
    }
  }

  const rowThreshold = canvas.width * 0.22;
  const colThreshold = canvas.height * 0.16;
  const denseRows = rowCounts
    .map((count, y) => ({ count, y }))
    .filter((item) => item.count >= rowThreshold)
    .map((item) => item.y);
  const denseCols = colCounts
    .map((count, x) => ({ count, x }))
    .filter((item) => item.count >= colThreshold)
    .map((item) => item.x);

  if (denseRows.length < 4 || denseCols.length < 2) return null;

  const top = Math.max(0, Math.min(...denseRows) - Math.round(canvas.height * 0.006));
  const bottom = Math.min(
    canvas.height,
    Math.max(...denseRows) + Math.round(canvas.height * 0.006),
  );
  const left = Math.max(0, Math.min(...denseCols) - Math.round(canvas.width * 0.004));
  const right = Math.min(
    canvas.width,
    Math.max(...denseCols) + Math.round(canvas.width * 0.004),
  );
  const width = right - left;
  const height = bottom - top;

  if (width < canvas.width * 0.4 || height < canvas.height * 0.35) return null;
  return { x: left, y: top, width, height };
};

const cropFieldImage = async (
  pageImageSrc: string,
  zone: (typeof FORM_FIELD_ZONES)[string],
  tableBounds: ImageBounds,
) => {
  const image = await loadDataUrlImage(pageImageSrc);
  const paddingX = Math.round(tableBounds.width * 0.008);
  const paddingY = Math.round(tableBounds.height * 0.006);
  const sourceX = Math.max(
    0,
    Math.round(tableBounds.x + tableBounds.width * zone.x) - paddingX,
  );
  const sourceY = Math.max(
    0,
    Math.round(tableBounds.y + tableBounds.height * zone.y) - paddingY,
  );
  const sourceWidth = Math.min(
    image.naturalWidth - sourceX,
    Math.round(tableBounds.width * zone.width) + paddingX * 2,
  );
  const sourceHeight = Math.min(
    image.naturalHeight - sourceY,
    Math.round(tableBounds.height * zone.height) + paddingY * 2,
  );
  const canvas = document.createElement("canvas");
  canvas.width = Math.max(1, sourceWidth * 2);
  canvas.height = Math.max(1, sourceHeight * 2);
  const context = canvas.getContext("2d");
  if (!context) throw new Error("Не вдалося підготувати crop поля.");
  context.fillStyle = "#ffffff";
  context.fillRect(0, 0, canvas.width, canvas.height);
  context.imageSmoothingEnabled = true;
  context.imageSmoothingQuality = "high";
  context.drawImage(
    image,
    sourceX,
    sourceY,
    sourceWidth,
    sourceHeight,
    0,
    0,
    canvas.width,
    canvas.height,
  );
  return enhanceCanvasForOcr(canvas);
};

const renderPagePreviewDataUrl = async (page: PDFPageProxy) => {
  const viewport = page.getViewport({ scale: TEXT_LAYER_RENDER_SCALE });
  const canvas = document.createElement("canvas");
  canvas.width = Math.max(1, Math.floor(viewport.width));
  canvas.height = Math.max(1, Math.floor(viewport.height));
  const context = canvas.getContext("2d");
  if (!context) throw new Error("Не вдалося підготувати preview сторінки.");

  context.fillStyle = "#ffffff";
  context.fillRect(0, 0, canvas.width, canvas.height);
  await page.render({
    canvasContext: context,
    background: "#ffffff",
    viewport,
  }).promise;

  return canvas.toDataURL("image/jpeg", 0.86);
};

const createOcrWorker = async (
  onProgress?: (progress: QuestionnaireParseProgress) => void,
) => {
  const worker = await createWorker(OCR_LANGUAGES, 1, {
    workerPath: `${TESSERACT_BASE}/worker.min.js`,
    corePath: `${TESSERACT_BASE}/tesseract-core-lstm.wasm.js`,
    langPath: TESSDATA_BASE,
    gzip: true,
    workerBlobURL: false,
    logger: (message) => {
      if (message.status === "recognizing text") {
        onProgress?.({
          stage: "ocr",
          progress: message.progress,
          message: `OCR: ${Math.round(message.progress * 100)}%`,
        });
      }
    },
  });
  await worker.setParameters({
    tessedit_pageseg_mode: PSM.SPARSE_TEXT,
    preserve_interword_spaces: "1",
    user_defined_dpi: "300",
    language_model_penalty_non_freq_dict_word: "0.6",
    language_model_penalty_non_dict_word: "0.4",
  });
  return worker;
};

const rowsFromOcrText = (text: string) =>
  text
    .split(/\r?\n/g)
    .map((row) => normalizeLikelyUkrainianOcr(row.replace(/\s+/g, " ").trim()))
    .filter(Boolean);

const cleanupFieldOcrValue = (value: string) =>
  cleanupValue(
    normalizeLikelyUkrainianOcr(value)
      .replace(/[|_[\]{}<>]/g, " ")
      .replace(/\s+/g, " "),
  );

const recognizeFormFields = async (
  worker: Awaited<ReturnType<typeof createOcrWorker>>,
  pageImageSrc: string,
  pageNumber: number,
  pageCount: number,
  onProgress?: (progress: QuestionnaireParseProgress) => void,
) => {
  const fieldImages: Record<string, string> = {};
  const fieldValues: Record<string, string> = {};
  const tableBounds = await detectFormTableBounds(pageImageSrc);
  if (!tableBounds) return { fieldImages, fieldValues };

  for (const [key, zone] of Object.entries(FORM_FIELD_ZONES)) {
    onProgress?.({
      stage: "ocr",
      pageNumber,
      pageCount,
      message: `OCR поля ${key}: сторінка ${pageNumber}`,
    });
    const crop = await cropFieldImage(pageImageSrc, zone, tableBounds);
    fieldImages[key] = crop;
    await worker.setParameters({
      tessedit_pageseg_mode: zone.mode,
      preserve_interword_spaces: "1",
      user_defined_dpi: "300",
    });
    const recognized = await worker.recognize(crop);
    const value = cleanupFieldOcrValue(recognized.data.text);
    if (value) fieldValues[key] = value;
  }

  await worker.setParameters({
    tessedit_pageseg_mode: PSM.SPARSE_TEXT,
    preserve_interword_spaces: "1",
    user_defined_dpi: "300",
  });

  return { fieldImages, fieldValues };
};

const runOcr = async (
  pdf: PDFDocumentProxy,
  pageNumbers: number[],
  onProgress?: (progress: QuestionnaireParseProgress) => void,
) => {
  const worker = await createOcrWorker(onProgress);
  const pages: QuestionnaireTextPage[] = [];

  try {
    for (const pageNumber of pageNumbers) {
      onProgress?.({
        stage: "ocr",
        pageNumber,
        pageCount: pageNumbers.length,
        message: `OCR сторінки ${pageNumber}`,
      });
      const page = await pdf.getPage(pageNumber);
      const variants = await renderPageToImageVariants(page);
      let best:
        | {
            rows: string[];
            text: string;
            imageSrc: string;
            confidence: number;
            score: number;
          }
        | null = null;
      for (const variant of variants) {
        onProgress?.({
          stage: "ocr",
          pageNumber,
          pageCount: pdf.numPages,
          message: `OCR сторінки ${pageNumber}: ${variant.label}`,
        });
        const recognized = await worker.recognize(variant.src);
        const rows = rowsFromOcrText(recognized.data.text);
        const text = rows.join("\n");
        const pageScore = scoreQuestionnairePage(text);
        const quality = pageScore * 100 + recognized.data.confidence;
        if (!best || quality > best.score * 100 + best.confidence) {
          best = {
            rows,
            text,
            imageSrc: variant.src,
            confidence: recognized.data.confidence,
            score: pageScore,
          };
        }
      }
      const fieldData = best?.imageSrc
        ? await recognizeFormFields(
            worker,
            best.imageSrc,
            pageNumber,
            pageNumbers.length,
            onProgress,
          )
        : { fieldImages: {}, fieldValues: {} };
      pages.push({
        pageNumber,
        text: best?.text ?? "",
        rows: best?.rows ?? [],
        imageSrc: best?.imageSrc ?? variants[0].src,
        fieldImages: fieldData.fieldImages,
        fieldValues: fieldData.fieldValues,
        itemCount: best?.rows.length ?? 0,
        score: best?.score ?? 0,
      });
    }
  } finally {
    await worker.terminate();
  }

  return pages;
};

export async function parseQuestionnairePdf(
  file: File,
  options: QuestionnaireParseOptions = {},
): Promise<QuestionnaireParseResult> {
  const data = new Uint8Array(await file.arrayBuffer());
  const pdf = await pdfjs.getDocument({ data }).promise;
  const requestedPages = options.pageNumbers?.length
    ? [...new Set(options.pageNumbers)]
        .filter((pageNumber) => pageNumber >= 1 && pageNumber <= pdf.numPages)
        .sort((a, b) => a - b)
    : Array.from({ length: pdf.numPages }, (_, index) => index + 1);
  const pages: QuestionnaireTextPage[] = [];

  for (const pageNumber of requestedPages) {
    options.onProgress?.({
      stage: "text",
      pageNumber,
      pageCount: requestedPages.length,
      message: `Читання текстового шару: сторінка ${pageNumber}`,
    });
    const page = await pdf.getPage(pageNumber);
    const content = await page.getTextContent();
    const imageSrc = await renderPagePreviewDataUrl(page);
    const rows = groupTextRows(content.items as PdfTextItem[]);
    const text = rows.join("\n");
    pages.push({
      pageNumber,
      text,
      rows,
      imageSrc,
      itemCount: content.items.length,
      score: scoreQuestionnairePage(text),
    });
  }

  const textSelectedPage =
    pages
      .filter((page) => page.text.trim())
      .sort((a, b) => b.score - a.score || a.pageNumber - b.pageNumber)[0] ??
    null;
  const shouldRunOcr =
    Boolean(options.useOcr) &&
    (!textSelectedPage || pages.filter((page) => page.text.trim()).length === 0);
  const finalPages = shouldRunOcr
    ? await runOcr(pdf, requestedPages, options.onProgress)
    : pages;
  const selectedPage =
    finalPages
      .filter((page) => page.text.trim())
      .sort((a, b) => b.score - a.score || a.pageNumber - b.pageNumber)[0] ??
    null;

  options.onProgress?.({
    stage: "done",
    message: "Парсинг PDF завершено.",
  });

  return {
    fileName: file.name,
    pageCount: pdf.numPages,
    textPageCount: finalPages.filter((page) => page.text.trim()).length,
    usedOcr: shouldRunOcr,
    selectedPageNumber: selectedPage?.pageNumber ?? null,
    selectedText: selectedPage?.text ?? "",
    fields: extractQuestionnaireFields(selectedPage ?? undefined),
    pages: finalPages,
  };
}

export async function prepareQuestionnairePdf(
  file: File,
  onProgress?: (progress: QuestionnaireParseProgress) => void,
): Promise<QuestionnaireParseResult> {
  return parseQuestionnairePdf(file, {
    useOcr: false,
    onProgress,
  });
}
