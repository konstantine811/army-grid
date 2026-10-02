import { useState } from "react";
import { Alert, Box, Button, Typography } from "@/components/sci/SciPrimitives";
import { CloudUploadOutlinedIcon } from "@/components/sci/icons";
import {
  createWorkbookDebugPayload,
  readWorkbookSnapshot,
} from "../../../excelRoundTrip";
import { DEFAULT_EXCEL_LAB_READ_OPTIONS } from "../parseUploadedExcel";

const isExcelFile = (file: File) =>
  /\.xlsx?$/i.test(file.name) ||
  file.type ===
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" ||
  file.type === "application/vnd.ms-excel";

export async function parseUbdExcelFile(file: File) {
  const snapshot = await readWorkbookSnapshot(file, {
    ...DEFAULT_EXCEL_LAB_READ_OPTIONS,
    skipStyleFills: true,
  });
  const parsed = createWorkbookDebugPayload(snapshot);
  console.log("[excel-lab][ubd]", parsed);
  return parsed;
}

export function ExcelLabUbdPanel() {
  const [fileName, setFileName] = useState("");
  const [error, setError] = useState("");
  const [isRunning, setIsRunning] = useState(false);

  const onFilePick = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    event.target.value = "";
    if (!file) return;
    if (!isExcelFile(file)) {
      setError("Оберіть файл Excel (.xlsx або .xls).");
      return;
    }

    setIsRunning(true);
    setError("");
    void parseUbdExcelFile(file)
      .then(() => setFileName(file.name))
      .catch((cause: unknown) => {
        setFileName("");
        setError(
          cause instanceof Error
            ? cause.message
            : "Не вдалося розібрати файл УБД.",
        );
      })
      .finally(() => setIsRunning(false));
  };

  return (
    <Box className="panel-card" sx={{ p: 2 }}>
      <Typography variant="subtitle1" sx={{ mb: 1 }}>
        УБД
      </Typography>
      <Button
        component="label"
        variant="contained"
        disabled={isRunning}
        startIcon={<CloudUploadOutlinedIcon />}
      >
        {isRunning ? "Парсю Excel…" : "Завантажити Excel"}
        <input
          hidden
          type="file"
          accept=".xlsx,.xls,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel"
          onChange={onFilePick}
        />
      </Button>
      {fileName ? (
        <Typography variant="body2" sx={{ mt: 2 }}>
          {fileName} — розбір у консолі (`[excel-lab][ubd]`).
        </Typography>
      ) : null}
      {error ? (
        <Alert severity="error" sx={{ mt: 2 }}>
          {error}
        </Alert>
      ) : null}
    </Box>
  );
}
