import { Box, Button, Stack, Typography } from "@/components/sci/SciPrimitives";
import { CloudUploadOutlinedIcon } from "@/components/sci/icons";
import { useEjoosWorkspace } from "./EjoosWorkspaceContext";

export function EjoosImportPanel() {
  const {
    live,
    pbSources,
    pbSnapshot,
    importEjoos,
    loadPb,
    loadPbFromDb,
    isLoading,
  } = useEjoosWorkspace();

  const pbCurrent = pbSources?.current;

  return (
    <Stack spacing={2} className="ejoos-import">
      <Box>
        <Typography variant="h6">Імпорт</Typography>
        <Typography variant="body2" className="ejoos-muted">
          Завантажте ЕЖООС і за потреби 1ПБ (аркуші sh, Рух, archive). Дані
          ЕЖООС дивіться у вкладках вище.
        </Typography>
      </Box>

      <Box className="ejoos-dropzone">
        <Typography variant="subtitle2" sx={{ mb: 0.5 }}>
          ЕЖООС
        </Typography>
        <Typography variant="body2" className="ejoos-muted" sx={{ mb: 1.5 }}>
          {live?.current
            ? `Зараз у БД: v${live.current.version} · ${live.current.sourceFileName || "без імені"}. Новий імпорт створить наступну версію.`
            : "Ще немає ЕЖООС у БД — завантажте файл (ШПО, ООС, Виключені, Табель…)."}
        </Typography>
        <Button
          component="label"
          variant={live?.current ? "outlined" : "contained"}
          startIcon={<CloudUploadOutlinedIcon />}
          disabled={isLoading}
          sx={live?.current ? undefined : { color: "#1a1a14" }}
        >
          {live?.current ? "Оновити ЕЖООС з файлу" : "Завантажити ЕЖООС"}
          <input
            hidden
            type="file"
            accept=".xlsx"
            onChange={(event) => {
              const file = event.target.files?.[0];
              if (file) void importEjoos(file);
              event.target.value = "";
            }}
          />
        </Button>
      </Box>

      <Box className="ejoos-dropzone">
        <Typography variant="subtitle2" sx={{ mb: 0.5 }}>
          1ПБ (sh / Рух / archive)
        </Typography>
        <Typography variant="body2" className="ejoos-muted" sx={{ mb: 1.5 }}>
          Після завантаження файл зберігається в БД, парситься по трьох
          аркушах і будує план операцій для ЕЖООС, якщо канонічний журнал уже
          відкритий або збережений у БД.
        </Typography>
        {pbSnapshot ? (
          <Typography variant="body2" sx={{ mb: 1.5 }}>
            У памʼяті: <strong>{pbSnapshot.fileName}</strong> · аркуші{" "}
            {pbSnapshot.sheets.map((sheet) => sheet.sheetName).join(", ")}
          </Typography>
        ) : null}
        {pbCurrent ? (
          <Typography variant="body2" sx={{ mb: 1.5 }}>
            У БД: <strong>{pbCurrent.sourceFileName || "без імені"}</strong> ·{" "}
            {new Date(pbCurrent.createdAt).toLocaleString("uk-UA")} · sha{" "}
            {pbCurrent.sha256.slice(0, 10)}…
          </Typography>
        ) : (
          <Typography variant="body2" className="ejoos-muted" sx={{ mb: 1.5 }}>
            У БД ще немає збереженого 1ПБ.
          </Typography>
        )}
        <Stack direction="row" spacing={1} style={{ flexWrap: "wrap", gap: 8 }}>
          <Button
            component="label"
            variant="contained"
            startIcon={<CloudUploadOutlinedIcon />}
            disabled={isLoading}
            sx={{ color: "#1a1a14" }}
          >
            Завантажити 1ПБ
            <input
              hidden
              type="file"
              accept=".xlsx"
              onChange={(event) => {
                const file = event.target.files?.[0];
                if (file) void loadPb(file);
                event.target.value = "";
              }}
            />
          </Button>
          {pbCurrent ? (
            <Button
              variant="outlined"
              disabled={isLoading}
              onClick={() => void loadPbFromDb(pbCurrent.id)}
            >
              Відкрити 1ПБ з БД
            </Button>
          ) : null}
        </Stack>
      </Box>
    </Stack>
  );
}
