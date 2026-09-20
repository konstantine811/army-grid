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
  EXCEL_LAB_SOURCE_COUNT,
  countKSPVKPeople,
  parseEJOOSExcelFile,
  parseKSPVKExcelFile,
  parseMilitaryServiceCardsExcelFile,
  parseStaffSheetExcelFile,
  type ExcelLabUploadedSources,
} from "./parseUploadedExcel";
import { collectVKData, describeCollectedVKData } from "./CollectVKData";

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
  const [kspSummary, setKspSummary] = useState("");
  const [mergedSummary, setMergedSummary] = useState("");
  const [error, setError] = useState("");
  const [isStaffRunning, setIsStaffRunning] = useState(false);
  const [isMilitaryCardsRunning, setIsMilitaryCardsRunning] = useState(false);
  const [isEjoosRunning, setIsEjoosRunning] = useState(false);
  const [isKspRunning, setIsKspRunning] = useState(false);

  const isBusy =
    isStaffRunning ||
    isMilitaryCardsRunning ||
    isEjoosRunning ||
    isKspRunning;
  const loadedCount = [
    uploadedSources.staff,
    uploadedSources.militaryCards,
    uploadedSources.ejoos,
    uploadedSources.ksp,
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

  const parseKspFile = async (file: File) => {
    setIsKspRunning(true);
    setError("");
    setKspSummary("");
    try {
      const result = await parseKSPVKExcelFile(file);
      setUploadedSources((current) => ({
        ...current,
        ksp: { fileName: file.name, state: result.ksp },
      }));
      setKspSummary(
        [
          buildSheetSummary(result.debug),
          `Особи з КСП (ВК): ${countKSPVKPeople(result.ksp)}`,
        ].join("\n"),
      );
    } catch (cause) {
      setError(
        cause instanceof Error
          ? cause.message
          : "Не вдалося розібрати файл військових із КСП.",
      );
    } finally {
      setIsKspRunning(false);
    }
  };

  const processCollectedData = () => {
    const processed = buildExcelLabProcessedData(uploadedSources);
    const collected = collectVKData(processed);
    setMergedSummary(
      [
        describeExcelLabProcessedData(processed),
        describeCollectedVKData(collected),
        "Повний результат — у console.log (F12 → Console).",
      ].join("\n"),
    );
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
            Окремі парсери (Штатка, квитки, ЄЖООС, КСП) і збір у єдиний
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
            Військові із КСП
          </Typography>
          <Typography variant="body2" color="text.secondary" sx={{ mb: 1.5 }}>
            Handler:{" "}
            <Typography
              component="code"
              variant="body2"
              sx={{ fontFamily: "monospace" }}
            >
              parseExcelKSPState
            </Typography>{" "}
            (
            <Typography
              component="code"
              variant="body2"
              sx={{ fontFamily: "monospace" }}
            >
              ExcelKSPVKData.ts
            </Typography>
            )
          </Typography>
          <Button
            component="label"
            variant="contained"
            disabled={isBusy}
            startIcon={<CloudUploadOutlinedIcon />}
          >
            {isKspRunning ? "Парсю КСП…" : "Завантажити військові із КСП"}
            <input
              hidden
              type="file"
              accept=".xlsx,.xls,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel"
              onChange={onFilePick(parseKspFile)}
            />
          </Button>
          {uploadedSources.ksp ? (
            <Typography variant="body2" sx={{ mt: 2 }}>
              Останній файл: {uploadedSources.ksp.fileName}
            </Typography>
          ) : null}
          {kspSummary ? (
            <Alert severity="info" sx={{ mt: 2 }}>
              <Typography
                component="pre"
                variant="body2"
                sx={{ m: 0, whiteSpace: "pre-wrap", fontFamily: "inherit" }}
              >
                {kspSummary}
              </Typography>
            </Alert>
          ) : null}
        </Box>

        <Box className="panel-card" sx={{ p: 2 }}>
          <Typography variant="subtitle1" sx={{ mb: 1 }}>
            Обробка даних
          </Typography>
          <Typography variant="body2" color="text.secondary" sx={{ mb: 1.5 }}>
            Після завантаження Штатки, квитків, ЄЖООС і/або КСП. Спочатку{" "}
            <Typography
              component="code"
              variant="body2"
              sx={{ fontFamily: "monospace" }}
            >
              buildExcelLabProcessedData
            </Typography>
            , потім твоя логіка в{" "}
            <Typography
              component="code"
              variant="body2"
              sx={{ fontFamily: "monospace" }}
            >
              collectVKData
            </Typography>{" "}
            (
            <Typography
              component="code"
              variant="body2"
              sx={{ fontFamily: "monospace" }}
            >
              CollectVKData.ts
            </Typography>
            ).
          </Typography>
          <Button
            variant="contained"
            disabled={loadedCount === 0}
            onClick={processCollectedData}
          >
            Обробити дані ({loadedCount}/{EXCEL_LAB_SOURCE_COUNT})
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
