import { useCallback, useMemo, useState } from "react";
import {
  Alert,
  Box,
  Button,
  Checkbox,
  Stack,
  TextField,
  Typography,
} from "@/components/sci/SciPrimitives";
import {
  SearchOutlinedIcon,
  CloudUploadOutlinedIcon,
  DeleteOutlineOutlinedIcon,
  FileDownloadOutlinedIcon,
  PictureAsPdfOutlinedIcon,
} from "@/components/sci/icons";
import { useAuth } from "../../auth/AuthProvider";
import {
  WordDocumentPreview,
  useWordPreviewBlob,
} from "../documents/WordDocumentPreview";
import { PreAnketaOcrReviewDialog } from "./components/PreAnketaOcrReviewDialog";
import { usePreAnketaOcr } from "./hooks/usePreAnketaOcr";
import {
  PRE_ANKETA_FIELD_DEFS,
  createEmptyPreAnketaForm,
  formatPreAnketaDownloadName,
  isPreAnketaCheckboxField,
  type PreAnketaFieldKey,
  type PreAnketaFormFields,
} from "./preAnketaFields";
import { createPreAnketaWordBlob } from "./preAnketaWordExport";
import { fillPreAnketaFieldsFromStaff } from "./preAnketaStaffLookup";

const groupedFields = () => {
  const groups = new Map<string, typeof PRE_ANKETA_FIELD_DEFS>();
  for (const field of PRE_ANKETA_FIELD_DEFS) {
    const list = groups.get(field.group) ?? [];
    list.push(field);
    groups.set(field.group, list);
  }
  return [...groups.entries()];
};

export function PreAnketaPage() {
  const { canEditArea } = useAuth();
  const canEdit = canEditArea("anketaData");
  const [fields, setFields] = useState<PreAnketaFormFields>(() =>
    createEmptyPreAnketaForm(),
  );
  const [message, setMessage] = useState("");
  const [staffLookupRunning, setStaffLookupRunning] = useState(false);
  const ocr = usePreAnketaOcr();

  const buildWordBlob = useCallback(
    () => createPreAnketaWordBlob(fields),
    [fields],
  );
  const preview = useWordPreviewBlob(
    buildWordBlob,
    Boolean(fields.fullName.trim() || fields.callsign.trim()),
  );

  const groups = useMemo(() => groupedFields(), []);

  const setField = (key: PreAnketaFieldKey, value: string) => {
    setFields((current) => ({ ...current, [key]: value }));
  };

  const pullFromStaff = async (
    current: PreAnketaFormFields,
    options?: { onlyEmpty?: boolean },
  ) => {
    setStaffLookupRunning(true);
    try {
      const result = await fillPreAnketaFieldsFromStaff(current, options);
      if (result.changed) {
        setFields(result.fields);
      }
      setMessage(result.message);
    } catch (cause) {
      setMessage(
        cause instanceof Error
          ? cause.message
          : "Не вдалося підтягнути дані зі штатки.",
      );
    } finally {
      setStaffLookupRunning(false);
    }
  };

  const downloadDocx = async () => {
    try {
      const blob = await createPreAnketaWordBlob(fields);
      const url = URL.createObjectURL(blob);
      const anchor = document.createElement("a");
      anchor.href = url;
      anchor.download = formatPreAnketaDownloadName(fields);
      anchor.click();
      window.setTimeout(() => URL.revokeObjectURL(url), 0);
      setMessage("Анкету завантажено.");
    } catch (cause) {
      setMessage(
        cause instanceof Error
          ? cause.message
          : "Не вдалося зібрати Word-файл.",
      );
    }
  };

  const applyOcr = () => {
    const next = ocr.applyProposals(fields, ocr.proposals);
    setFields(next);
    ocr.clearReview();
    setMessage("Поля з Gemini застосовано.");
    void pullFromStaff(next);
  };

  return (
    <Box className="page-shell pre-anketa-page" sx={{ p: 2 }}>
      <Stack spacing={2}>
        <Box>
          <Typography variant="h5">Анкета</Typography>
          <Typography variant="body2" color="text.secondary">
            Завантажте фото або PDF документів, Gemini заповнить поля, далі відредагуйте
            вручну та завантажте готову анкету Word.
          </Typography>
        </Box>

        {(message || ocr.progressMessage || ocr.error) && (
          <Stack spacing={1}>
            {message ? <Alert severity="info">{message}</Alert> : null}
            {ocr.progressMessage ? (
              <Alert severity="info">{ocr.progressMessage}</Alert>
            ) : null}
            {ocr.error ? <Alert severity="error">{ocr.error}</Alert> : null}
          </Stack>
        )}

        <Stack
          direction={{ xs: "column", xl: "row" }}
          spacing={2}
          alignItems="stretch"
        >
          <Box sx={{ flex: 1, minWidth: 0 }}>
            <Stack spacing={2}>
              <Box className="panel-card" sx={{ p: 2 }}>
                <Typography variant="subtitle1" sx={{ mb: 1 }}>
                  Скани, фото або PDF
                </Typography>
                <Stack direction="row" spacing={1} flexWrap="wrap" useFlexGap>
                  <Button
                    component="label"
                    variant="outlined"
                    disabled={!canEdit || ocr.isRunning}
                  >
                    <CloudUploadOutlinedIcon fontSize="small" />
                    Додати файли
                    <input
                      hidden
                      type="file"
                      accept="image/*,application/pdf,.pdf"
                      multiple
                      onChange={(event) => {
                        const list = event.target.files;
                        if (!list?.length) return;
                        void ocr.addScans(list);
                        event.target.value = "";
                      }}
                    />
                  </Button>
                  <Button
                    variant="contained"
                    disabled={!canEdit || ocr.isRunning || !ocr.scans.length}
                    onClick={() => void ocr.runOcr(fields)}
                    startIcon={<SearchOutlinedIcon />}
                  >
                    {ocr.isRunning ? "Розпізнаю…" : "Заповнити Gemini"}
                  </Button>
                </Stack>

                {ocr.scans.length ? (
                  <Stack
                    direction="row"
                    spacing={1}
                    flexWrap="wrap"
                    useFlexGap
                    sx={{ mt: 2 }}
                  >
                    {ocr.scans.map((scan) => (
                      <Box
                        key={scan.id}
                        sx={{
                          width: 120,
                          border: "1px solid var(--sci-border)",
                          borderRadius: 1,
                          overflow: "hidden",
                        }}
                      >
                        {scan.kind === "pdf" ? (
                          <Stack
                            alignItems="center"
                            justifyContent="center"
                            sx={{
                              height: 90,
                              bgcolor: "rgba(0,0,0,0.04)",
                            }}
                          >
                            <PictureAsPdfOutlinedIcon fontSize="large" />
                          </Stack>
                        ) : (
                          <img
                            src={scan.previewUrl}
                            alt={scan.name}
                            style={{
                              display: "block",
                              width: "100%",
                              height: 90,
                              objectFit: "cover",
                            }}
                          />
                        )}
                        <Stack
                          direction="row"
                          alignItems="center"
                          justifyContent="space-between"
                          sx={{ px: 0.5 }}
                        >
                          <Typography
                            variant="caption"
                            noWrap
                            title={scan.name}
                            sx={{ maxWidth: 88 }}
                          >
                            {scan.name}
                          </Typography>
                          <Button
                            size="small"
                            color="inherit"
                            onClick={() => ocr.removeScan(scan.id)}
                          >
                            <DeleteOutlineOutlinedIcon fontSize="small" />
                          </Button>
                        </Stack>
                      </Box>
                    ))}
                  </Stack>
                ) : (
                  <Typography variant="body2" sx={{ mt: 1.5 }}>
                    Паспорт, ID-картка, РНОКПП, військовий квиток — фото або PDF.
                  </Typography>
                )}
              </Box>

              {groups.map(([group, defs]) => (
                <Box key={group} className="panel-card" sx={{ p: 2 }}>
                  <Stack
                    direction="row"
                    alignItems="center"
                    justifyContent="space-between"
                    spacing={1}
                    sx={{ mb: 1.5 }}
                  >
                    <Typography variant="subtitle1">{group}</Typography>
                    {group === "Основне" ? (
                      <Button
                        size="small"
                        variant="outlined"
                        disabled={
                          !canEdit ||
                          staffLookupRunning ||
                          fields.fullName.trim().length < 3
                        }
                        onClick={() => void pullFromStaff(fields)}
                      >
                        {staffLookupRunning
                          ? "Шукаю…"
                          : "Зі штатки: позивний і звання"}
                      </Button>
                    ) : null}
                  </Stack>
                  <Stack spacing={1.5}>
                    {defs.map((def) =>
                      def.checkbox ? (
                        <Checkbox
                          key={def.key}
                          label={def.label}
                          checked={Boolean(fields[def.key])}
                          disabled={!canEdit}
                          onCheckedChange={(checked) =>
                            setField(def.key, checked === true ? "1" : "")
                          }
                        />
                      ) : (
                        <TextField
                          key={def.key}
                          label={def.label}
                          value={fields[def.key]}
                          disabled={!canEdit}
                          multiline={def.multiline}
                          minRows={def.multiline ? 2 : 1}
                          fullWidth
                          onChange={(event) =>
                            setField(def.key, event.target.value)
                          }
                          onBlur={
                            def.key === "fullName"
                              ? (event) => {
                                  const fullName = event.target.value.trim();
                                  if (fullName.length < 3) return;
                                  if (
                                    fields.callsign.trim() &&
                                    fields.rank.trim()
                                  ) {
                                    return;
                                  }
                                  void pullFromStaff({
                                    ...fields,
                                    fullName: event.target.value,
                                  });
                                }
                              : undefined
                          }
                        />
                      ),
                    )}
                  </Stack>
                </Box>
              ))}

              <Stack direction="row" spacing={1}>
                <Button
                  variant="contained"
                  startIcon={<FileDownloadOutlinedIcon />}
                  onClick={() => void downloadDocx()}
                >
                  Завантажити анкету Word
                </Button>
                <Button
                  variant="outlined"
                  onClick={() => {
                    setFields(createEmptyPreAnketaForm());
                    setMessage("Форму очищено.");
                  }}
                >
                  Очистити
                </Button>
              </Stack>
            </Stack>
          </Box>

          <Box
            sx={{
              flex: 1,
              minWidth: 0,
              minHeight: 720,
            }}
            className="panel-card"
          >
            <Typography variant="subtitle1" sx={{ p: 2, pb: 0 }}>
              Попередній перегляд
            </Typography>
            <WordDocumentPreview
              blob={preview.blob}
              isLoading={preview.isLoading}
              error={preview.error}
            />
          </Box>
        </Stack>
      </Stack>

      <PreAnketaOcrReviewDialog
        open={ocr.proposals.length > 0}
        proposals={ocr.proposals}
        onChange={ocr.setProposals}
        onClose={ocr.clearReview}
        onApply={applyOcr}
      />
    </Box>
  );
}
