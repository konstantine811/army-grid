import { Box, Button, Stack, Typography } from "@/components/sci/SciPrimitives";
import { FileDownloadOutlinedIcon } from "@/components/sci/icons";
import { useEjoosWorkspace } from "./EjoosWorkspaceContext";

export function EjoosExportPanel() {
  const { live, downloadCurrentEjoos, setTab, isLoading } = useEjoosWorkspace();
  const current = live?.current;

  return (
    <Stack spacing={2} className="ejoos-export">
      <Box>
        <Typography variant="h6">Експорт ЕЖООС</Typography>
        <Typography variant="body2" className="ejoos-muted">
          Скачати поточну версію ЕЖООС з БД як повний .xlsx з оригінальними
          аркушами, форматуванням і вже застосованими змінами. Це не файл 1ПБ.
        </Typography>
      </Box>

      {current ? (
        <Box className="ejoos-stat-card" sx={{ maxWidth: 480 }}>
          <span className="ejoos-stat-label">Поточна версія</span>
          <strong>v{current.version}</strong>
          <span>{current.sourceFileName || "без імені"}</span>
          <span>
            {new Date(current.createdAt).toLocaleString("uk-UA")} · sha{" "}
            {current.sha256.slice(0, 12)}… ·{" "}
            {Math.round(current.byteSize / 1024)} КБ
          </span>
          {current.sourcePbFileName ? (
            <span>Джерело 1ПБ: {current.sourcePbFileName}</span>
          ) : null}
          <Stack direction="row" spacing={1} sx={{ mt: 1.5 }} style={{ flexWrap: "wrap" }}>
            <Button
              variant="contained"
              startIcon={<FileDownloadOutlinedIcon />}
              disabled={isLoading}
              onClick={() => void downloadCurrentEjoos()}
              sx={{ color: "#1a1a14" }}
            >
              Скачати .xlsx
            </Button>
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
