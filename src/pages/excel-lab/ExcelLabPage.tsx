import { useState } from "react";
import {
  Alert,
  Box,
  Button,
  Stack,
  Typography,
} from "@/components/sci/SciPrimitives";
import { CloudUploadOutlinedIcon } from "@/components/sci/icons";
import {
  buildExcelLabProcessedData,
  countEJOOSPeople,
  countMilitaryServiceCardPeople,
  describeExcelLabProcessedData,
  parseEJOOSExcelFile,
  parseMilitaryServiceCardsExcelFile,
  parseStaffSheetExcelFile,
  type ExcelLabUploadedSources,
} from "./parseUploadedExcel";
import { collectVKData } from "./CollectVKData";

const isExcelFile = (file: File) =>
  /\.xlsx?$/i.test(file.name) ||
  file.type ===
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" ||
  file.type === "application/vnd.ms-excel";

const buildSheetSummary = (debug: {
  fileName: string;
  sheetCount: number;
  sheets: Array<{
    sheetName: string;
    parsedRowCount: number;
    columnCount: number;
  }>;
}) => {
  const sheetLines = debug.sheets.map(
    (sheet) =>
      `${sheet.sheetName}: ${sheet.parsedRowCount} рядків · ${sheet.columnCount} колонок`,
  );
  return [
    `Файл: ${debug.fileName}`,
    `Аркушів: ${debug.sheetCount}`,
    ...sheetLines,
  ].join("\n");
};

export function ExcelLabPage() {
  const [uploadedSources, setUploadedSources] =
    useState<ExcelLabUploadedSources>({});
  const [staffSummary, setStaffSummary] = useState("");
  const [militaryCardsSummary, setMilitaryCardsSummary] = useState("");
  const [ejoosSummary, setEjoosSummary] = useState("");
  const [mergedSummary, setMergedSummary] = useState("");
  const [error, setError] = useState("");
  const [isStaffRunning, setIsStaffRunning] = useState(false);
  const [isMilitaryCardsRunning, setIsMilitaryCardsRunning] = useState(false);
  const [isEjoosRunning, setIsEjoosRunning] = useState(false);

  const isBusy = isStaffRunning || isMilitaryCardsRunning || isEjoosRunning;
  const loadedCount = [
    uploadedSources.staff,
    uploadedSources.militaryCards,
    uploadedSources.ejoos,
  ].filter(Boolean).length;

  const parseStaffFile = async (file: File) => {
    setIsStaffRunning(true);
    setError("");
    setStaffSummary("");
    try {
      const result = await parseStaffSheetExcelFile(file);
      setUploadedSources((current) => ({
        ...current,
        staff: { fileName: file.name, state: result.state },
      }));
      setStaffSummary(
        [
          buildSheetSummary(result.debug),
          `Особи в стані: ${Object.keys(result.state).length}`,
        ].join("\n"),
      );
    } catch (cause) {
      setError(
        cause instanceof Error
          ? cause.message
          : "Не вдалося розібрати файл Штатки.",
      );
    } finally {
      setIsStaffRunning(false);
    }
  };

  const parseMilitaryCardsFile = async (file: File) => {
    setIsMilitaryCardsRunning(true);
    setError("");
    setMilitaryCardsSummary("");
    try {
      const result = await parseMilitaryServiceCardsExcelFile(file);
      setUploadedSources((current) => ({
        ...current,
        militaryCards: { fileName: file.name, cards: result.cards },
      }));
      setMilitaryCardsSummary(
        [
          buildSheetSummary(result.debug),
          `Особи з картками (ВК/ТПВ/ДОВІДКИ): ${countMilitaryServiceCardPeople(result.cards)}`,
        ].join("\n"),
      );
    } catch (cause) {
      setError(
        cause instanceof Error
          ? cause.message
          : "Не вдалося розібрати файл військових квитків.",
      );
    } finally {
      setIsMilitaryCardsRunning(false);
    }
  };

  const parseEjoosFile = async (file: File) => {
    setIsEjoosRunning(true);
    setError("");
    setEjoosSummary("");
    try {
      const result = await parseEJOOSExcelFile(file);
      setUploadedSources((current) => ({
        ...current,
        ejoos: { fileName: file.name, state: result.ejoos },
      }));
      setEjoosSummary(
        [
          buildSheetSummary(result.debug),
          `Особи в ЄЖООС: ${countEJOOSPeople(result.ejoos)}`,
        ].join("\n"),
      );
    } catch (cause) {
      setError(
        cause instanceof Error
          ? cause.message
          : "Не вдалося розібрати файл ЄЖООС.",
      );
    } finally {
      setIsEjoosRunning(false);
    }
  };

  const mergeUploadedData = () => {
    const processed = buildExcelLabProcessedData(uploadedSources);
    collectVKData(processed);
    setMergedSummary(describeExcelLabProcessedData(processed));
  };

  const onFilePick =
    (handler: (file: File) => Promise<void>) =>
    (event: React.ChangeEvent<HTMLInputElement>) => {
      const file = event.target.files?.[0];
      event.target.value = "";
      if (!file) return;
      if (!isExcelFile(file)) {
        setError("Оберіть файл Excel (.xlsx або .xls).");
        return;
      }
      void handler(file);
    };

  return (
    <Box className="page-shell excel-lab-page" sx={{ p: 2 }}>
      <Stack spacing={2}>
        <Box>
          <Typography variant="h5">Excel Lab</Typography>
          <Typography variant="body2" color="text.secondary">
            Три окремі завантаження (Штатка, квитки, ЄЖООС) і збір у єдиний
            об&apos;єкт.
          </Typography>
        </Box>

        <Box className="panel-card" sx={{ p: 2 }}>
          <Typography variant="subtitle1" sx={{ mb: 1 }}>
            Парсер Штатки
          </Typography>
          <Button
            component="label"
            variant="contained"
            disabled={isBusy}
            startIcon={<CloudUploadOutlinedIcon />}
          >
            {isStaffRunning ? "Парсю Штатку…" : "Завантажити Штатку"}
            <input
              hidden
              type="file"
              accept=".xlsx,.xls,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel"
              onChange={onFilePick(parseStaffFile)}
            />
          </Button>
          {uploadedSources.staff ? (
            <Typography variant="body2" sx={{ mt: 2 }}>
              Останній файл: {uploadedSources.staff.fileName}
            </Typography>
          ) : null}
          {staffSummary ? (
            <Alert severity="info" sx={{ mt: 2 }}>
              <Typography
                component="pre"
                variant="body2"
                sx={{ m: 0, whiteSpace: "pre-wrap", fontFamily: "inherit" }}
              >
                {staffSummary}
              </Typography>
            </Alert>
          ) : null}
        </Box>

        <Box className="panel-card" sx={{ p: 2 }}>
          <Typography variant="subtitle1" sx={{ mb: 1 }}>
            Військові квитки
          </Typography>
          <Button
            component="label"
            variant="contained"
            disabled={isBusy}
            startIcon={<CloudUploadOutlinedIcon />}
          >
            {isMilitaryCardsRunning
              ? "Парсю квитки…"
              : "Завантажити військові квитки"}
            <input
              hidden
              type="file"
              accept=".xlsx,.xls,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel"
              onChange={onFilePick(parseMilitaryCardsFile)}
            />
          </Button>
          {uploadedSources.militaryCards ? (
            <Typography variant="body2" sx={{ mt: 2 }}>
              Останній файл: {uploadedSources.militaryCards.fileName}
            </Typography>
          ) : null}
          {militaryCardsSummary ? (
            <Alert severity="info" sx={{ mt: 2 }}>
              <Typography
                component="pre"
                variant="body2"
                sx={{ m: 0, whiteSpace: "pre-wrap", fontFamily: "inherit" }}
              >
                {militaryCardsSummary}
              </Typography>
            </Alert>
          ) : null}
        </Box>

        <Box className="panel-card" sx={{ p: 2 }}>
          <Typography variant="subtitle1" sx={{ mb: 1 }}>
            ЄЖООС
          </Typography>
          <Typography variant="body2" color="text.secondary" sx={{ mb: 1.5 }}>
            Handler:{" "}
            <Typography
              component="code"
              variant="body2"
              sx={{ fontFamily: "monospace" }}
            >
              parseEJOOSExcelFile
            </Typography>
          </Typography>
          <Button
            component="label"
            variant="outlined"
            disabled={isBusy}
            startIcon={<CloudUploadOutlinedIcon />}
          >
            {isEjoosRunning ? "Парсю ЄЖООС…" : "Завантажити ЄЖООС"}
            <input
              hidden
              type="file"
              accept=".xlsx,.xls,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel"
              onChange={onFilePick(parseEjoosFile)}
            />
          </Button>
          {uploadedSources.ejoos ? (
            <Typography variant="body2" sx={{ mt: 2 }}>
              Останній файл: {uploadedSources.ejoos.fileName}
            </Typography>
          ) : null}
          {ejoosSummary ? (
            <Alert severity="info" sx={{ mt: 2 }}>
              <Typography
                component="pre"
                variant="body2"
                sx={{ m: 0, whiteSpace: "pre-wrap", fontFamily: "inherit" }}
              >
                {ejoosSummary}
              </Typography>
            </Alert>
          ) : null}
        </Box>

        <Box className="panel-card" sx={{ p: 2 }}>
          <Typography variant="subtitle1" sx={{ mb: 1 }}>
            Збір усіх даних
          </Typography>
          <Typography variant="body2" color="text.secondary" sx={{ mb: 1.5 }}>
            Функція{" "}
            <Typography
              component="code"
              variant="body2"
              sx={{ fontFamily: "monospace" }}
            >
              buildExcelLabProcessedData
            </Typography>{" "}
            — об&apos;єднує результати трьох завантажень.
          </Typography>
          <Button
            variant="contained"
            disabled={loadedCount === 0}
            onClick={mergeUploadedData}
          >
            Зібрати дані ({loadedCount}/3)
          </Button>
          {mergedSummary ? (
            <Alert severity="success" sx={{ mt: 2 }}>
              <Typography
                component="pre"
                variant="body2"
                sx={{ m: 0, whiteSpace: "pre-wrap", fontFamily: "inherit" }}
              >
                {mergedSummary}
              </Typography>
            </Alert>
          ) : null}
        </Box>

        {error ? <Alert severity="error">{error}</Alert> : null}
      </Stack>
    </Box>
  );
}
