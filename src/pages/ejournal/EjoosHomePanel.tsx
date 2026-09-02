import { Box, Button, Stack, Typography } from "@/components/sci/SciPrimitives";
import { CloudUploadOutlinedIcon } from "@/components/sci/icons";
import { FileDownloadOutlinedIcon } from "@/components/sci/icons";
import { formatApiDateTime } from "../../shared/format";
import { useEjoosWorkspace } from "./ejoosWorkspaceState";

export function EjoosHomePanel() {
  const {
    live,
    session,
    pbSources,
    ejoosSnapshot,
    setTab,
    importEjoos,
    loadEjoosFromDb,
    loadPbFromDb,
    downloadCurrentEjoos,
  } = useEjoosWorkspace();

  const current = live?.current;
  const pbCurrent = pbSources?.current;
  const c = session?.counters;

  return (
    <Stack spacing={2} className="ejoos-home">
      <Box>
        <Typography variant="h6">Робоче місце ЕЖООС</Typography>
        <Typography variant="body2" className="ejoos-muted">
          1ПБ (sh / Рух / archive) — лише джерело змін. Заповнюємо і скачуємо
          саме ЕЖООС (ШПО, ООС, Виключені, Табель…).
        </Typography>
      </Box>

      <div className="ejoos-stat-grid">
        <div className="ejoos-stat-card">
          <span className="ejoos-stat-label">Поточний ЕЖООС</span>
          {current ? (
            <>
              <strong>v{current.version}</strong>
              <span>
                {current.sourceFileName || "без імені"} · sha{" "}
                {current.sha256.slice(0, 10)}…
              </span>
              <span>
                {formatApiDateTime(current.createdAt)}
              </span>
            </>
          ) : (
            <>
              <strong>—</strong>
              <span>Ще не завантажено в БД</span>
            </>
          )}
        </div>

        <div className="ejoos-stat-card">
          <span className="ejoos-stat-label">Останній 1ПБ</span>
          {session ? (
            <>
              <strong>{session.pbFileName}</strong>
              <span>
                {new Date(session.analyzedAt).toLocaleString("uk-UA")}
              </span>
              <span>{c?.oosLike ?? 0} осіб у sh</span>
            </>
          ) : pbCurrent ? (
            <>
              <strong>{pbCurrent.sourceFileName || "у БД"}</strong>
              <span>
                {formatApiDateTime(pbCurrent.createdAt)}
              </span>
              <span>Збережено на сервері — відкрийте для аналізу</span>
            </>
          ) : (
            <>
              <strong>—</strong>
              <span>Завантажте 1ПБ для аналізу</span>
            </>
          )}
        </div>

        <div className="ejoos-stat-card">
          <span className="ejoos-stat-label">Після аналізу</span>
          <div className="ejoos-stat-row">
            <span>
              ООС <strong>{c?.oosLike ?? "—"}</strong>
            </span>
            <span>
              В строю <strong>{c?.onDuty ?? "—"}</strong>
            </span>
            <span>
              Зміни <strong>{c?.changes ?? "—"}</strong>
            </span>
            <span>
              Нові <strong>{c?.newcomers ?? "—"}</strong>
            </span>
            <span>
              Помилки <strong>{c?.errors ?? "—"}</strong>
            </span>
          </div>
        </div>
      </div>

      <Stack direction="row" spacing={1} style={{ flexWrap: "wrap" }}>
        {session && session.counters.changes > 0 ? (
          <Button
            variant="contained"
            onClick={() => setTab("import")}
            sx={{ color: "#1a1a14" }}
          >
            Переглянути {session.counters.changes} змін
          </Button>
        ) : null}
        {!session && pbCurrent && current ? (
          <Button
            variant="contained"
            onClick={() => void loadPbFromDb(pbCurrent.id)}
            sx={{ color: "#1a1a14" }}
          >
            Відкрити 1ПБ з БД і проаналізувати
          </Button>
        ) : null}
        <Button
          variant={current ? "outlined" : "contained"}
          onClick={() => setTab("import")}
          startIcon={<CloudUploadOutlinedIcon />}
          sx={current ? undefined : { color: "#1a1a14" }}
        >
          {current ? "Імпорт / 1ПБ" : "Спочатку ЕЖООС"}
        </Button>
        {current ? (
          <>
            <Button
              component="label"
              variant="outlined"
              startIcon={<CloudUploadOutlinedIcon />}
            >
              Оновити ЕЖООС з файлу
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
            <Button
              variant="outlined"
              startIcon={<FileDownloadOutlinedIcon />}
              onClick={() => void downloadCurrentEjoos()}
            >
              Скачати ЕЖООС
            </Button>
            {!ejoosSnapshot ? (
              <Button variant="outlined" onClick={() => void loadEjoosFromDb()}>
                Відкрити з БД
              </Button>
            ) : null}
            <Button variant="outlined" onClick={() => setTab("import")}>
              Експорт
            </Button>
            <Button variant="outlined" onClick={() => setTab("import")}>
              Особовий склад
            </Button>
            <Button variant="outlined" onClick={() => setTab("import")}>
              Перевірка
            </Button>
            <Button variant="outlined" onClick={() => setTab("import")}>
              Історія
            </Button>
          </>
        ) : (
          <Button
            component="label"
            variant="contained"
            startIcon={<CloudUploadOutlinedIcon />}
            sx={{ color: "#1a1a14" }}
          >
            Завантажити канонічний ЕЖООС
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
        )}
      </Stack>
    </Stack>
  );
}
