import type { DragEvent } from "react";
import { Alert, Box, Button, Stack, Typography } from "@/components/sci/SciPrimitives";
import { CheckCircleOutlineIcon } from "@/components/sci/icons";
import { CloudUploadOutlinedIcon } from "@/components/sci/icons";
import { MoreVertOutlinedIcon } from "@/components/sci/icons";
import { SyncAltOutlinedIcon } from "@/components/sci/icons";
import { TableChartOutlinedIcon } from "@/components/sci/icons";
import { UploadFileOutlinedIcon } from "@/components/sci/icons";
import type { DataSourceFile } from "./types";

export type DataSourcesProps = {
  sources: DataSourceFile[];
  canMerge: boolean;
  onPrimaryFile: (file: File | undefined) => void;
  onMergeFile: (file: File | undefined) => void;
};

export function DataSources({
  sources,
  canMerge,
  onPrimaryFile,
  onMergeFile,
}: DataSourcesProps) {
  const handleDrop = (event: DragEvent<HTMLDivElement>) => {
    event.preventDefault();
    onPrimaryFile(event.dataTransfer.files?.[0]);
  };

  return (
    <section className="panel">
      <div className="panel-heading">Джерела даних</div>
      <div className="panel-body">
        {sources.length === 0 ? (
          <Alert severity="info" variant="outlined" sx={{ mb: 2 }}>
            Файли ще не завантажені. Основний Excel додається тут і одразу
            відкривається в робочій таблиці.
          </Alert>
        ) : (
          sources.map((source) => (
            <div className="file-card" key={source.id}>
              <div className="file-icon">
                <TableChartOutlinedIcon fontSize="small" />
              </div>
              <Box sx={{ minWidth: 0 }}>
                <Typography noWrap>{source.name}</Typography>
                <Typography variant="caption" className="muted">
                  {source.role} · XLSX · {source.size}
                </Typography>
                <Stack direction="row" spacing={4} sx={{ mt: 2 }}>
                  <Typography variant="caption" className="muted">
                    Рядків: {source.rows}
                  </Typography>
                  <Typography variant="caption" className="muted">
                    Стовпців: {source.columns}
                  </Typography>
                </Stack>
                <Typography
                  variant="caption"
                  className="muted"
                  sx={{ display: "block", mt: 2 }}
                >
                  Завантажено:
                </Typography>
                <Typography variant="caption">{source.uploadedAt}</Typography>
              </Box>
              <Stack
                sx={{ alignItems: "center", justifyContent: "space-between" }}
              >
                <MoreVertOutlinedIcon fontSize="small" />
                <CheckCircleOutlineIcon color="success" fontSize="small" />
              </Stack>
            </div>
          ))
        )}

        <div
          className="drop-zone"
          onDragOver={(event) => event.preventDefault()}
          onDrop={handleDrop}
        >
          <Box>
            <CloudUploadOutlinedIcon color="disabled" />
            <Typography variant="body2">Перетягніть файл сюди</Typography>
            <Typography variant="caption" className="muted">
              або
            </Typography>
            <Stack
              direction="row"
              spacing={1}
              sx={{ justifyContent: "center", mt: 1 }}
            >
              <Button
                component="label"
                size="small"
                variant="outlined"
                startIcon={<UploadFileOutlinedIcon />}
              >
                Основний
                <input
                  hidden
                  type="file"
                  accept=".xlsx"
                  onChange={(event) => {
                    onPrimaryFile(event.target.files?.[0]);
                    event.target.value = "";
                  }}
                />
              </Button>
              <Button
                component="label"
                size="small"
                variant="outlined"
                startIcon={<SyncAltOutlinedIcon />}
                disabled={!canMerge}
              >
                Для merge
                <input
                  hidden
                  type="file"
                  accept=".xlsx"
                  onChange={(event) => {
                    onMergeFile(event.target.files?.[0]);
                    event.target.value = "";
                  }}
                />
              </Button>
            </Stack>
          </Box>
        </div>
        <Typography
          variant="caption"
          className="muted"
          sx={{ display: "block", mt: 1.5 }}
        >
          Підтримуються: .xlsx, .xls, .csv · максимальний розмір файлу: 50 MB
        </Typography>
      </div>
    </section>
  );
}
