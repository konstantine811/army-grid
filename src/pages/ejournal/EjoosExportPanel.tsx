import {
  Alert,
  Box,
  Button,
  Stack,
  Typography,
} from "@/components/sci/SciPrimitives";
import { FileDownloadOutlinedIcon } from "@/components/sci/icons";
import { formatApiDateTime } from "../../shared/format";
import { useEjoosWorkspace } from "./ejoosWorkspaceState";
import { useEjoosLiveView } from "./useEjoosLiveView";

const looksLikeAnketaFill = (notes?: string | null, kind?: string) =>
  /анкет/i.test(String(notes ?? "")) || /anketa/i.test(String(kind ?? ""));

export function EjoosExportPanel() {
  const { live, downloadCurrentEjoos, restoreStylesFromHistory, setTab, isLoading } =
    useEjoosWorkspace();
  const { view } = useEjoosLiveView();
  const current = live?.current;
  const protocol = current?.changeProtocol as { kind?: string } | null;
  const currentLooksBad =
    Boolean(current) &&
    looksLikeAnketaFill(current?.notes, protocol?.kind);
  const checkErrors = view.checks.filter((item) => item.severity === "error");
  const checkWarns = view.checks.filter((item) => item.severity === "warn");
  const exportGate = currentLooksBad
    ? "error"
    : checkErrors.length
      ? "error"
      : checkWarns.length
        ? "warn"
        : "ok";

  const confirmAndDownload = (message: string) => {
    const ok = window.confirm(message);
    if (!ok) return;
    void downloadCurrentEjoos();
  };

  return (
    <Stack spacing={2} className="ejoos-export">
      <Box>
        <Typography variant="h6">Експорт ЕЖООС</Typography>
        <Typography variant="body2" className="ejoos-muted">
          Скачати поточну версію ЕЖООС з БД як повний .xlsx з оригінальними
          аркушами, форматуванням і вже застосованими змінами. Це не файл 1ПБ.
        </Typography>
      </Box>

      {currentLooksBad ? (
        <Alert severity="error" variant="outlined">
          Поточна v{current?.version} може бути пошкоджена для Excel. Стилі
          можна повернути з попередньої версії в Історії (дані лишаються), або
          відкотитись / взяти резервну копію.
          <Stack direction="row" spacing={1} sx={{ mt: 1 }} style={{ flexWrap: "wrap" }}>
            <Button
              size="small"
              variant="contained"
              disabled={isLoading}
              onClick={() => void restoreStylesFromHistory()}
              sx={{ color: "#1a1a14" }}
            >
              Повернути стилі з історії
            </Button>
            <Button
              size="small"
              variant="outlined"
              onClick={() => setTab("history")}
            >
              Відкрити історію
            </Button>
          </Stack>
        </Alert>
      ) : null}

      {current ? (
        <Box className="ejoos-stat-card" sx={{ maxWidth: 480 }}>
          <span className="ejoos-stat-label">Поточна версія</span>
          <strong>v{current.version}</strong>
          <span>{current.sourceFileName || "без імені"}</span>
          <span>
            {formatApiDateTime(current.createdAt)} · sha{" "}
            {current.sha256.slice(0, 12)}… ·{" "}
            {Math.round(current.byteSize / 1024)} КБ
          </span>
          {current.sourcePbFileName ? (
            <span>Джерело 1ПБ: {current.sourcePbFileName}</span>
          ) : null}
          {exportGate === "error" && !currentLooksBad ? (
            <Alert severity="error" variant="outlined" sx={{ mt: 1.5 }}>
              У Перевірках є помилки ({checkErrors.length}). Це не готовий
              ЕЖООС — можна лише робочу копію.
              <div>
                {checkErrors
                  .slice(0, 4)
                  .map((item) => item.title)
                  .join("; ")}
              </div>
            </Alert>
          ) : null}
          {exportGate === "warn" ? (
            <Alert severity="warning" variant="outlined" sx={{ mt: 1.5 }}>
              Є попередження ({checkWarns.length}). Завантаження після
              підтвердження.
            </Alert>
          ) : null}
          <Stack
            direction="row"
            spacing={1}
            sx={{ mt: 1.5 }}
            style={{ flexWrap: "wrap" }}
          >
            {exportGate === "ok" ? (
              <Button
                variant="contained"
                startIcon={<FileDownloadOutlinedIcon />}
                disabled={isLoading}
                onClick={() => void downloadCurrentEjoos()}
                sx={{ color: "#1a1a14" }}
              >
                Скачати готовий ЕЖООС
              </Button>
            ) : exportGate === "warn" ? (
              <Button
                variant="contained"
                startIcon={<FileDownloadOutlinedIcon />}
                disabled={isLoading}
                onClick={() =>
                  confirmAndDownload(
                    `Є ${checkWarns.length} попереджень у Перевірках.\nВсе одно скачати готовий ЕЖООС?`,
                  )
                }
                sx={{ color: "#1a1a14" }}
              >
                Скачати після підтвердження
              </Button>
            ) : (
              <Button
                variant="outlined"
                startIcon={<FileDownloadOutlinedIcon />}
                disabled={isLoading}
                onClick={() => {
                  if (currentLooksBad) {
                    const ok = window.confirm(
                      `v${current.version} ймовірно пошкоджена.\nВсе одно скачати робочу копію?`,
                    );
                    if (!ok) {
                      setTab("history");
                      return;
                    }
                    void downloadCurrentEjoos();
                    return;
                  }
                  confirmAndDownload(
                    `У Перевірках ${checkErrors.length} помилок.\nСкачати лише робочу копію, не готовий ЕЖООС?`,
                  );
                }}
              >
                Скачати робочу копію
              </Button>
            )}
            <Button variant="outlined" onClick={() => setTab("history")}>
              Історія версій
            </Button>
          </Stack>
        </Box>
      ) : (
        <Typography variant="body2" className="ejoos-muted">
          Немає поточної версії ЕЖООС у БД.
        </Typography>
      )}
    </Stack>
  );
}
