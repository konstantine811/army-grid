import { useMemo, useState } from "react";
import {
  Alert,
  Box,
  Button,
  Chip,
  LinearProgress,
  Stack,
  Typography,
} from "@/components/sci/SciPrimitives";
import {
  ArticleOutlinedIcon,
  SearchOutlinedIcon,
  UploadFileOutlinedIcon,
} from "@/components/sci/icons";
import { api } from "../../api";
import {
  extractQuestionnaireFields,
  parseQuestionnairePdf,
  prepareQuestionnairePdf,
  type ParsedQuestionnaireField,
  type QuestionnaireParseResult,
  type QuestionnaireTextPage,
} from "./questionnairePdfParser";

export function QuestionnaireParserPage() {
  const [result, setResult] = useState<QuestionnaireParseResult | null>(null);
  const [currentFile, setCurrentFile] = useState<File | null>(null);
  const [selectedPageNumbers, setSelectedPageNumbers] = useState<number[]>([]);
  const [selectedPageNumber, setSelectedPageNumber] = useState<number | null>(null);
  const [fields, setFields] = useState<ParsedQuestionnaireField[]>([]);
  const [message, setMessage] = useState("Завантажте PDF анкети для пошуку форми та витягу полів.");
  const [ocrProgress, setOcrProgress] = useState("");
  const [isParsing, setIsParsing] = useState(false);
  const [isAiParsing, setIsAiParsing] = useState(false);
  const selectedPage = useMemo(
    () =>
      result?.pages.find(
        (page) => page.pageNumber === selectedPageNumber,
      ) ?? null,
    [result, selectedPageNumber],
  );

  const selectPageForParsing = (page: QuestionnaireTextPage) => {
    setSelectedPageNumber(page.pageNumber);
    setFields(extractQuestionnaireFields(page));
    setMessage(
      `Активна сторінка ${page.pageNumber}. Відмітьте сторінки чекбоксами й натисніть «Парсити вибрані».`,
    );
  };

  const updateField = (key: string, value: string) => {
    setFields((current) =>
      current.map((field) =>
        field.key === key ? { ...field, value, confidence: "high" } : field,
      ),
    );
  };

  const setProgressMessage = (progress: {
    pageNumber?: number;
    pageCount?: number;
    progress?: number;
    message: string;
  }) => {
    const pagePart =
      progress.pageNumber && progress.pageCount
        ? ` · ${progress.pageNumber}/${progress.pageCount}`
        : "";
    const percent =
      typeof progress.progress === "number"
        ? ` · ${Math.round(progress.progress * 100)}%`
        : "";
    setOcrProgress(`${progress.message}${pagePart}${percent}`);
  };

  const loadFileForSelection = async (file: File | undefined) => {
    if (!file) return;
    if (file.type && file.type !== "application/pdf") {
      setMessage("Потрібен PDF файл.");
      return;
    }

    setIsParsing(true);
    setCurrentFile(file);
    setResult(null);
    setSelectedPageNumbers([]);
    setSelectedPageNumber(null);
    setFields([]);
    setOcrProgress("");
    try {
      const prepared = await prepareQuestionnairePdf(file, setProgressMessage);
      const suggestedPage =
        prepared.selectedPageNumber ?? prepared.pages[0]?.pageNumber ?? null;
      setResult(prepared);
      setSelectedPageNumber(suggestedPage);
      setSelectedPageNumbers(suggestedPage ? [suggestedPage] : []);
      setFields([]);
      setMessage(
        suggestedPage
          ? `PDF підготовлено. Оберіть сторінки для парсингу. Попередньо відмічено сторінку ${suggestedPage}.`
          : "PDF підготовлено. Оберіть сторінки для парсингу вручну.",
      );
    } catch (error) {
      setMessage(
        error instanceof Error
          ? error.message
          : "Не вдалося прочитати PDF анкету.",
      );
    } finally {
      setIsParsing(false);
    }
  };

  const togglePageSelection = (pageNumber: number) => {
    setSelectedPageNumbers((current) =>
      current.includes(pageNumber)
        ? current.filter((item) => item !== pageNumber)
        : [...current, pageNumber].sort((a, b) => a - b),
    );
  };

  const parseSelectedPages = async () => {
    if (!currentFile || !selectedPageNumbers.length) {
      setMessage("Оберіть хоча б одну сторінку для парсингу.");
      return;
    }

    setIsParsing(true);
    setFields([]);
    setOcrProgress("");
    try {
      const parsed = await parseQuestionnairePdf(currentFile, {
        useOcr: true,
        pageNumbers: selectedPageNumbers,
        onProgress: setProgressMessage,
      });
      setResult(parsed);
      setSelectedPageNumber(parsed.selectedPageNumber ?? selectedPageNumbers[0]);
      setFields(parsed.fields);
      setMessage(
        parsed.selectedPageNumber
          ? `Парсинг завершено по вибраних сторінках: ${selectedPageNumbers.join(", ")}. Активна сторінка ${parsed.selectedPageNumber}.`
          : "Парсинг завершено, але сторінку анкети не вдалося визначити.",
      );
    } catch (error) {
      setMessage(
        error instanceof Error
          ? error.message
          : "Не вдалося спарсити вибрані сторінки.",
      );
    } finally {
      setIsParsing(false);
    }
  };

  const parseActivePageWithAi = async () => {
    if (!selectedPage?.imageSrc) {
      setMessage("Оберіть активну сторінку для AI OCR.");
      return;
    }

    setIsAiParsing(true);
    setMessage("AI OCR читає активну сторінку. Це може зайняти трохи часу.");
    try {
      const aiResult = await api.aiQuestionnaireOcr({
        imageData: selectedPage.imageSrc,
        fileName: currentFile?.name,
        pageNumber: String(selectedPage.pageNumber),
      });
      setFields((current) => {
        const byKey = new Map(aiResult.fields.map((field) => [field.key, field]));
        const baseFields = current.length
          ? current
          : aiResult.fields.map((field) => ({
              key: field.key,
              label: field.label,
              value: "",
              confidence: "low" as const,
            }));
        return baseFields.map((field) => {
          const next = byKey.get(field.key);
          return next
            ? {
                ...field,
                value: next.value ?? "",
                confidence:
                  next.confidence === "high" ||
                  next.confidence === "medium" ||
                  next.confidence === "low"
                    ? next.confidence
                    : "medium",
              }
            : field;
        });
      });
      setMessage(
        `AI OCR завершено: модель ${aiResult.model}, сторінка ${selectedPage.pageNumber}. Перевірте поля по preview.`,
      );
    } catch (error) {
      setMessage(
        error instanceof Error
          ? error.message
          : "Не вдалося виконати AI OCR.",
      );
    } finally {
      setIsAiParsing(false);
    }
  };

  return (
    <main className="main-panel questionnaire-parser-page">
      <header className="topbar analytics-topbar">
        <Box>
          <Typography component="h1" variant="h4">
            Парсинг анкет PDF
          </Typography>
          <Typography variant="body2" color="text.secondary">
            PDF анкети → пошук сторінки форми → витяг заповнених полів
          </Typography>
        </Box>
        <Stack direction="row" spacing={1}>
          <Button
            component="label"
            variant="contained"
            startIcon={<UploadFileOutlinedIcon />}
            sx={{ color: "#1a1a14" }}
          >
            Завантажити PDF
            <input
              hidden
              accept="application/pdf,.pdf"
              type="file"
              onChange={(event) => {
                void loadFileForSelection(event.target.files?.[0]);
                event.currentTarget.value = "";
              }}
            />
          </Button>
          <Button
            disabled={!currentFile || !selectedPageNumbers.length || isParsing}
            variant="outlined"
            onClick={() => void parseSelectedPages()}
          >
            Парсити вибрані
          </Button>
          <Button
            disabled={!selectedPage?.imageSrc || isParsing || isAiParsing}
            variant="outlined"
            onClick={() => void parseActivePageWithAi()}
          >
            {isAiParsing ? "AI OCR…" : "AI OCR активної"}
          </Button>
        </Stack>
      </header>

      {(isParsing || isAiParsing) && <LinearProgress color="primary" />}
      {isParsing && ocrProgress ? (
        <Alert severity="info" variant="outlined" className="personnel-page-alert">
          {ocrProgress}
        </Alert>
      ) : null}
      <Alert
        severity={!result || result.textPageCount ? "info" : "warning"}
        variant="outlined"
        className="personnel-page-alert"
      >
        {message}
      </Alert>

      <section className="questionnaire-parser-layout">
        <aside className="analytics-panel questionnaire-parser-pages">
          <div className="panel-heading">Сторінки PDF</div>
          {result ? (
            <div className="questionnaire-parser-page-list">
              {result.pages.map((page) => (
                <article
                  className={
                    page.pageNumber === selectedPageNumber ? "active" : ""
                  }
                  key={page.pageNumber}
                  role="button"
                  tabIndex={0}
                  onClick={() => selectPageForParsing(page)}
                  onKeyDown={(event) => {
                    if (event.key === "Enter" || event.key === " ") {
                      event.preventDefault();
                      selectPageForParsing(page);
                    }
                  }}
                >
                  <label
                    className="questionnaire-parser-page-check"
                    onClick={(event) => event.stopPropagation()}
                  >
                    <input
                      checked={selectedPageNumbers.includes(page.pageNumber)}
                      type="checkbox"
                      onChange={() => togglePageSelection(page.pageNumber)}
                    />
                    <strong>Сторінка {page.pageNumber}</strong>
                  </label>
                  <span>{page.itemCount} текстових елементів</span>
                  <span>
                    {page.pageNumber === selectedPageNumber
                      ? "активна для перегляду"
                      : "натисніть для перегляду"}
                  </span>
                  <Chip
                    size="small"
                    label={`score ${page.score}`}
                    color={page.score ? "success" : "default"}
                  />
                </article>
              ))}
            </div>
          ) : (
            <div className="questionnaire-parser-empty">
              <ArticleOutlinedIcon />
              <span>PDF ще не завантажено</span>
            </div>
          )}
        </aside>

        <section className="analytics-panel questionnaire-parser-fields">
          <div className="panel-heading">
            <span>Поля анкети</span>
            {selectedPageNumber ? (
              <Chip
                size="small"
                color="success"
                label={`сторінка ${selectedPageNumber}`}
              />
            ) : null}
          </div>
          {fields.length ? (
            <div className="questionnaire-parser-field-grid">
              {fields.map((field) => (
                <label key={field.key}>
                  <span>{field.label}</span>
                  {field.sourceImage ? (
                    <img
                      alt={`Фрагмент поля ${field.label}`}
                      className="questionnaire-parser-field-source"
                      src={field.sourceImage}
                    />
                  ) : null}
                  <input
                    value={field.value}
                    placeholder="—"
                    onChange={(event) => updateField(field.key, event.target.value)}
                  />
                </label>
              ))}
            </div>
          ) : (
            <div className="questionnaire-parser-empty">
              <SearchOutlinedIcon />
              <span>Поля з’являться після читання текстового PDF</span>
            </div>
          )}
        </section>

        <section className="analytics-panel questionnaire-parser-raw">
          <div className="panel-heading">
            <span>Джерело парсингу</span>
            {selectedPageNumber ? (
              <Chip size="small" label={`сторінка ${selectedPageNumber}`} />
            ) : null}
          </div>
          {selectedPage?.imageSrc ? (
            <div className="questionnaire-parser-source">
              <img
                alt={`Сторінка ${selectedPage.pageNumber}, з якої виконано парсинг`}
                src={selectedPage.imageSrc}
              />
            </div>
          ) : null}
          <details className="questionnaire-parser-ocr-details">
            <summary>OCR-текст</summary>
            <pre>
              {selectedPage?.text ||
                "Текстового шару немає або сторінку анкети ще не знайдено."}
            </pre>
          </details>
        </section>
      </section>
    </main>
  );
}
