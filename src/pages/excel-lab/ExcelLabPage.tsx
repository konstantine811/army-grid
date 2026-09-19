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
  parseStaffSheetExcelFile,
  parseUploadedExcelFile,
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
  const [genericFileName, setGenericFileName] = useState("");
  const [genericSummary, setGenericSummary] = useState("");
  const [error, setError] = useState("");
  const [isStaffRunning, setIsStaffRunning] = useState(false);
  const [isGenericRunning, setIsGenericRunning] = useState(false);

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

  const parseGenericFile = async (file: File) => {
    setIsGenericRunning(true);
    setError("");
    setGenericSummary("");
    setGenericFileName(file.name);
    try {
      const result = await parseUploadedExcelFile(file);
      setGenericSummary(
        [
          buildSheetSummary(result.debug),
          "Деталі — у console.log (F12 → Console).",
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
            Окремо парсер Штатки та довільний перегляд Excel у консолі.
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
            disabled={isStaffRunning || isGenericRunning}
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
            Довільний Excel
          </Typography>
          <Typography variant="body2" color="text.secondary" sx={{ mb: 1.5 }}>
            Будь-який файл → структура в консолі. Handler:{" "}
            <Typography
              component="code"
              variant="body2"
              sx={{ fontFamily: "monospace" }}
            >
              src/pages/excel-lab/parseUploadedExcel.ts
            </Typography>
          </Typography>
          <Button
            component="label"
            variant="outlined"
            disabled={isStaffRunning || isGenericRunning}
            startIcon={<CloudUploadOutlinedIcon />}
          >
            {isGenericRunning ? "Читаю…" : "Завантажити Excel"}
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
