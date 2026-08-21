import { useCallback, useEffect, useState } from "react";
import { Alert, Box, Stack, Typography } from "@/components/sci/SciPrimitives";
import { CloudUploadOutlinedIcon } from "@/components/sci/icons";
import { Button as SciButton } from "../../components/ui/button/button";
import { readWorkbookSnapshot, type ExcelWorkbookSnapshot } from "../../excelRoundTrip";
import {
  logBchsLabParseResult,
  parseBchsLabWorkbook,
  publishBchsLabToConsole,
  type BchsLabParseResult,
} from "./bchsLabParse";

export function BchsLabPage() {
  const [snapshot, setSnapshot] = useState<ExcelWorkbookSnapshot | null>(null);
  const [parseResult, setParseResult] = useState<BchsLabParseResult | null>(null);
  const [message, setMessage] = useState("Імпортуйте Excel — дані зʼявляться у консолі.");
  const [isLoading, setIsLoading] = useState(false);

  const applyWorkbook = useCallback((nextSnapshot: ExcelWorkbookSnapshot) => {
    const nextParse = parseBchsLabWorkbook(nextSnapshot);
    setSnapshot(nextSnapshot);
    setParseResult(nextParse);
    publishBchsLabToConsole(nextSnapshot, nextParse);
    logBchsLabParseResult(nextParse);
    setMessage(
      `Розпарсено «${nextSnapshot.fileName}»: ${nextParse.summary.totalPeople} рядків · нова ${nextParse.summary.novaCount}. Дивись window.__BCHS_LAB__ у консолі.`,
    );
  }, []);

  useEffect(() => {
    publishBchsLabToConsole(snapshot, parseResult);
  }, [snapshot, parseResult]);

  const loadFile = async (file: File | undefined) => {
    if (!file) return;

    setIsLoading(true);
    try {
      const nextSnapshot = await readWorkbookSnapshot(file);
      applyWorkbook(nextSnapshot);
    } catch (error) {
      setMessage(
        error instanceof Error
          ? error.message
          : "Не вдалося прочитати Excel-файл.",
      );
    } finally {
      setIsLoading(false);
    }
  };

  return (
    <main className="main-panel bchs-page bchs-lab-page">
      <header className="topbar analytics-topbar">
        <Box>
          <Typography component="h1" variant="h4">
            БЧС Lab
          </Typography>
          <Typography variant="body2" color="text.secondary">
            Тестовий імпорт · парсинг у консоль · логіку допишемо крок за кроком
          </Typography>
        </Box>
        <Stack direction="row" spacing={1}>
          <SciButton asChild variant="OUTLINE">
            <label>
              <CloudUploadOutlinedIcon fontSize="small" />
              {isLoading ? "Читаю…" : "Імпорт Excel"}
              <input
                hidden
                type="file"
                accept=".xlsx,.xlsm,.xls"
                onChange={(event) => {
                  void loadFile(event.target.files?.[0]);
                  event.target.value = "";
                }}
              />
            </label>
          </SciButton>
          <SciButton
            variant="OUTLINE"
            disabled={!parseResult}
            onClick={() => parseResult && logBchsLabParseResult(parseResult)}
          >
            Лог у консоль
          </SciButton>
        </Stack>
      </header>

      <Alert severity="info" sx={{ mb: 2 }}>
        {message}
      </Alert>

      <section className="analytics-panel bchs-lab-panel">
        <div className="panel-heading">Консоль</div>
        <Typography variant="body2" component="div" sx={{ lineHeight: 1.7 }}>
          <p>Після імпорту відкрий DevTools → Console. Усі підрахунки — лише <strong>A = нова</strong>.</p>
          <ul>
            <li>
              <code>window.__BCHS_LAB__.logAway(&apos;Командування&apos;)</code>{" "}
              — AK/AL/AM/AN + таблиця людей
            </li>
            <li>
              <code>window.__BCHS_LAB__.awayCommand()</code> — Excel COUNTIFS
              (ж + штаб + група БС)
            </li>
            <li>
              <code>
                window.__BCHS_LAB__.countAway(&#123; rosterUnits: [&apos;штаб&apos;],
                rank: &apos;солд.&apos; &#125;)
              </code>
            </li>
            <li>
              <code>window.__BCHS_LAB__.log()</code> — повний parseResult
            </li>
          </ul>
        </Typography>
      </section>

      {parseResult ? (
        <section className="analytics-panel bchs-lab-panel">
          <div className="panel-heading">Короткий зріз</div>
          <Stack spacing={0.75}>
            <Typography variant="body2">
              Файл: <strong>{parseResult.fileName}</strong>
            </Typography>
            <Typography variant="body2">
              Список ОС:{" "}
              {parseResult.rosterSheet
                ? `${parseResult.rosterSheet.sheetName} (${parseResult.summary.totalPeople})`
                : "не знайдено"}
            </Typography>
            <Typography variant="body2">
              Аркуш1:{" "}
              {parseResult.calculationSheet?.sheetName ?? "немає в файлі"}
            </Typography>
            <Typography variant="body2">
              нова: <strong>{parseResult.summary.novaCount}</strong> · з ПІБ:{" "}
              {parseResult.summary.withFullName}
            </Typography>
            <Typography variant="body2" color="text.secondary">
              Топ статусів:{" "}
              {Object.entries(parseResult.summary.statusCounts)
                .sort((left, right) => right[1] - left[1])
                .slice(0, 5)
                .map(([label, count]) => `${label} (${count})`)
                .join(" · ")}
            </Typography>
          </Stack>
        </section>
      ) : null}
    </main>
  );
}
