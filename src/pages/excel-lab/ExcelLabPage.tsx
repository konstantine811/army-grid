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
  countMilitaryServiceCardPeople,
  parseGenericTestExcelFile,
  parseMilitaryServiceCardsExcelFile,
  parseStaffSheetExcelFile,
} from "./parseUploadedExcel";

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
  const [staffFileName, setStaffFileName] = useState("");
  const [staffSummary, setStaffSummary] = useState("");
  const [militaryCardsFileName, setMilitaryCardsFileName] = useState("");
  const [militaryCardsSummary, setMilitaryCardsSummary] = useState("");
  const [genericFileName, setGenericFileName] = useState("");
  const [genericSummary, setGenericSummary] = useState("");
  const [error, setError] = useState("");
  const [isStaffRunning, setIsStaffRunning] = useState(false);
  const [isMilitaryCardsRunning, setIsMilitaryCardsRunning] = useState(false);
  const [isGenericRunning, setIsGenericRunning] = useState(false);

  const isBusy = isStaffRunning || isMilitaryCardsRunning || isGenericRunning;

  const parseStaffFile = async (file: File) => {
    setIsStaffRunning(true);
    setError("");
    setStaffSummary("");
    setStaffFileName(file.name);
    try {
      const result = await parseStaffSheetExcelFile(file);
      const peopleCount = Object.keys(result.state).length;
      setStaffSummary(
        [
          buildSheetSummary(result.debug),
          `Особи в стані: ${peopleCount}`,
          "Результат parseExcelLabState — у console.log (F12 → Console).",
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
    setMilitaryCardsFileName(file.name);
    try {
      const result = await parseMilitaryServiceCardsExcelFile(file);
      const peopleCount = countMilitaryServiceCardPeople(result.cards);
      setMilitaryCardsSummary(
        [
          buildSheetSummary(result.debug),
          `Особи з картками (ВК/ТПВ/ДОВІДКИ): ${peopleCount}`,
          "Результат parseExcelLabMilitaryServiceCards — у console.log (F12 → Console).",
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

  const parseGenericFile = async (file: File) => {
    setIsGenericRunning(true);
    setError("");
    setGenericSummary("");
    setGenericFileName(file.name);
    try {
      const result = await parseGenericTestExcelFile(file);
      setGenericSummary(
        [
          buildSheetSummary(result.debug),
          "Тестовий парс — snapshot і debug у console.log (F12 → Console).",
        ].join("\n"),
      );
    } catch (cause) {
      setError(
        cause instanceof Error
          ? cause.message
          : "Не вдалося прочитати Excel-файл.",
      );
    } finally {
      setIsGenericRunning(false);
    }
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
            Окремі парсери: Штатка, військові квитки та тестовий перегляд
            Excel у консолі.
          </Typography>
        </Box>

        <Box className="panel-card" sx={{ p: 2 }}>
          <Typography variant="subtitle1" sx={{ mb: 1 }}>
            Парсер Штатки
          </Typography>
          <Typography variant="body2" color="text.secondary" sx={{ mb: 1.5 }}>
            Логіка в{" "}
            <Typography
              component="code"
              variant="body2"
              sx={{ fontFamily: "monospace" }}
            >
              src/pages/excel-lab/ExcelLabStateParser.ts
            </Typography>
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
          {staffFileName ? (
            <Typography variant="body2" sx={{ mt: 2 }}>
              Останній файл: {staffFileName}
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
          <Typography variant="body2" color="text.secondary" sx={{ mb: 1.5 }}>
            Аркуші ВК, ТПВ, ДОВІДКИ →{" "}
            <Typography
              component="code"
              variant="body2"
              sx={{ fontFamily: "monospace" }}
            >
              handleParsedExcelWorkbook
            </Typography>{" "}
            у{" "}
            <Typography
              component="code"
              variant="body2"
              sx={{ fontFamily: "monospace" }}
            >
              parseUploadedExcel.ts
            </Typography>
          </Typography>
          <Button
            component="label"
            variant="contained"
            color="secondary"
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
          {militaryCardsFileName ? (
            <Typography variant="body2" sx={{ mt: 2 }}>
              Останній файл: {militaryCardsFileName}
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
            Тестовий Excel
          </Typography>
          <Typography variant="body2" color="text.secondary" sx={{ mb: 1.5 }}>
            Будь-який файл → snapshot і debug у консолі без парсера квитків.
            Handler:{" "}
            <Typography
              component="code"
              variant="body2"
              sx={{ fontFamily: "monospace" }}
            >
              parseGenericTestExcelFile
            </Typography>
          </Typography>
          <Button
            component="label"
            variant="outlined"
            disabled={isBusy}
            startIcon={<CloudUploadOutlinedIcon />}
          >
            {isGenericRunning ? "Читаю…" : "Завантажити тестовий Excel"}
            <input
              hidden
              type="file"
              accept=".xlsx,.xls,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel"
              onChange={onFilePick(parseGenericFile)}
            />
          </Button>
          {genericFileName ? (
            <Typography variant="body2" sx={{ mt: 2 }}>
              Останній файл: {genericFileName}
            </Typography>
          ) : null}
          {genericSummary ? (
            <Alert severity="info" sx={{ mt: 2 }}>
              <Typography
                component="pre"
                variant="body2"
                sx={{ m: 0, whiteSpace: "pre-wrap", fontFamily: "inherit" }}
              >
                {genericSummary}
              </Typography>
            </Alert>
          ) : null}
        </Box>

        {error ? <Alert severity="error">{error}</Alert> : null}
      </Stack>
    </Box>
  );
}
