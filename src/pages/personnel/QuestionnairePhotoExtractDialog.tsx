import { useEffect, useRef, useState } from "react";
import {
  Alert,
  Box,
  Button,
  LinearProgress,
  Stack,
  Typography,
} from "@/components/sci/SciPrimitives";
import { FloatingWindow } from "./FloatingWindow";
import {
  bulkExtractPhotosFromQuestionnaires,
  type BulkExtractPhotosRowResult,
} from "./utils/bulk-extract-photos-from-questionnaires";

type ExtractPersonInput = {
  externalId: string;
  fullName: string;
};

const STATUS_LABEL: Record<
  BulkExtractPhotosRowResult["status"] | "pending",
  string
> = {
  pending: "Очікує…",
  saved: "Фото збережено",
  skipped: "Пропущено",
  failed: "Помилка",
};

const REASON_LABEL: Record<string, string> = {
  no_pdf: "немає PDF у БД",
  name_mismatch: "ПІБ у назві файлу не збігається",
  no_face: "обличчя в PDF не знайдено",
  save_failed: "не вдалося зберегти",
};

export function QuestionnairePhotoExtractDialog({
  open,
  people,
  onClose,
  onPhotoSaved,
}: {
  open: boolean;
  people: ExtractPersonInput[];
  onClose: () => void;
  onPhotoSaved: (externalId: string, photoData: string) => void;
}) {
  const [isRunning, setIsRunning] = useState(false);
  const [progress, setProgress] = useState({ done: 0, total: 0, saved: 0 });
  const [summary, setSummary] = useState("");
  const [error, setError] = useState("");
  const [rows, setRows] = useState<BulkExtractPhotosRowResult[]>([]);
  const abortRef = useRef<AbortController | null>(null);

  useEffect(() => {
    if (!open) {
      abortRef.current?.abort();
      abortRef.current = null;
      return;
    }

    const controller = new AbortController();
    abortRef.current = controller;
    const snapshot = people;

    const run = async () => {
      setIsRunning(true);
      setError("");
      setSummary("");
      setRows([]);
      setProgress({ done: 0, total: snapshot.length, saved: 0 });

      if (!snapshot.length) {
        setSummary("Немає осіб без фото, але з анкетою в БД.");
        setIsRunning(false);
        return;
      }

      try {
        const result = await bulkExtractPhotosFromQuestionnaires({
          targets: snapshot,
          signal: controller.signal,
          onProgress: (done, total, saved) => {
            if (controller.signal.aborted) return;
            setProgress({ done, total, saved });
            setSummary(`Обробка ${done}/${total} · збережено ${saved}`);
          },
        });

        if (controller.signal.aborted) return;

        setRows(result.rows);
        setSummary(
          `Готово: збережено ${result.saved}, пропущено ${result.skipped}, помилок ${result.failed}.`,
        );

        for (const row of result.rows) {
          if (row.status !== "saved" || !row.photoData) continue;
          onPhotoSaved(row.externalId, row.photoData);
        }
      } catch (err) {
        if (controller.signal.aborted) return;
        setError(
          err instanceof Error
            ? err.message
            : "Не вдалося виконати масове витягнення фото.",
        );
      } finally {
        if (!controller.signal.aborted) setIsRunning(false);
      }
    };

    void run();
    return () => {
      controller.abort();
    };
    // Snapshot people only when dialog opens.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [open]);

  const handleClose = () => {
    abortRef.current?.abort();
    onClose();
  };

  return (
    <FloatingWindow
      bodyClassName="questionnel-photo-extract-body"
      defaultHeight={520}
      defaultWidth={640}
      footer={
        <Button disabled={isRunning} onClick={handleClose} variant="outlined">
          {isRunning ? "Скасувати" : "Закрити"}
        </Button>
      }
      onClose={handleClose}
      open={open}
      placement="right"
      subtitle="Лише анкети з БД · без пошуку на диску"
      title={`Фото з анкет · ${people.length}`}
    >
      <Stack spacing={1.5}>
        {error ? <Alert severity="error">{error}</Alert> : null}
        {summary ? (
          <Typography sx={{ fontSize: 13 }}>{summary}</Typography>
        ) : null}
        {isRunning ? (
          <Box>
            <LinearProgress
              value={
                progress.total
                  ? Math.round((progress.done / progress.total) * 100)
                  : 0
              }
              variant="determinate"
            />
          </Box>
        ) : null}
        <Box
          sx={{
            maxHeight: 360,
            overflow: "auto",
            fontSize: 12,
            lineHeight: 1.45,
          }}
        >
          {rows.map((row) => (
            <div key={row.externalId}>
              {row.fullName} · {STATUS_LABEL[row.status]}
              {row.reason ? ` (${REASON_LABEL[row.reason] ?? row.reason})` : ""}
              {row.message ? ` — ${row.message}` : ""}
            </div>
          ))}
        </Box>
      </Stack>
    </FloatingWindow>
  );
}
