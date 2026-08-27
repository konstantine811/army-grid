import { Button, Chip, Stack, Typography } from "@/components/sci/SciPrimitives";
import { useEjoosWorkspace } from "./EjoosWorkspaceContext";

export function EjoosHistoryPanel() {
  const {
    live,
    downloadVersion,
    downloadVersionProtocol,
    rollback,
    isLoading,
  } = useEjoosWorkspace();

  const versions = live?.versions ?? [];
  const currentId = live?.current?.id;

  return (
    <Stack spacing={2} className="ejoos-history">
      <div>
        <Typography variant="h6">Історія версій</Typography>
        <Typography variant="body2" className="ejoos-muted">
          Версії не перезаписуються. Відкат створює нову копію обраної версії.
        </Typography>
      </div>

      {versions.length === 0 ? (
        <Typography variant="body2" className="ejoos-muted">
          Поки немає збережених версій.
        </Typography>
      ) : (
        <Stack spacing={1}>
          {versions.map((version) => (
            <div key={version.id} className="ejoos-history-row">
              <div className="ejoos-history-meta">
                <Chip
                  size="small"
                  label={`v${version.version}`}
                  color={version.id === currentId ? "primary" : "default"}
                />
                <Typography variant="body2">
                  {new Date(version.createdAt).toLocaleString("uk-UA")}
                  {version.sourceFileName ? ` · ${version.sourceFileName}` : ""}
                </Typography>
                {version.sourcePbFileName ? (
                  <Typography variant="caption" className="ejoos-muted">
                    1ПБ: {version.sourcePbFileName}
                  </Typography>
                ) : null}
                {version.notes ? (
                  <Typography variant="caption" className="ejoos-muted">
                    {version.notes}
                  </Typography>
                ) : null}
              </div>
              <Stack direction="row" spacing={0.5} style={{ flexWrap: "wrap" }}>
                <Button
                  size="small"
                  disabled={isLoading}
                  onClick={() =>
                    void downloadVersion(
                      version.id,
                      version.sourceFileName || undefined,
                    )
                  }
                >
                  Файл
                </Button>
                <Button
                  size="small"
                  disabled={isLoading}
                  onClick={() => downloadVersionProtocol(version)}
                >
                  Протокол
                </Button>
                {version.id !== currentId ? (
                  <Button
                    size="small"
                    disabled={isLoading}
                    onClick={() => void rollback(version.id)}
                  >
                    Відкотити
                  </Button>
                ) : null}
              </Stack>
            </div>
          ))}
        </Stack>
      )}
    </Stack>
  );
}
