import { useEffect, useMemo, useState } from "react";
import {
  Alert,
  Button,
  Checkbox,
  MenuItem,
  Stack,
  TextField,
  Typography,
} from "@/components/sci/SciPrimitives";
import {
  CheckCircleOutlineIcon,
  DeleteOutlineOutlinedIcon,
  SettingsOutlinedIcon,
} from "@/components/sci/icons";
import {
  api,
  type BackendDocumentSignatoryPreset,
  type SaveDocumentSignatoryPreset,
} from "@/api";
import { processSignatureTransparentBackground } from "../documents/ubdSignatureImage";
import { UbdTemplateSettings } from "./UbdTemplateSettings";

const documentTypes = [
  { value: "ubdReport", label: "Рапорт на УБД" },
  { value: "ubdRestoreReport", label: "Рапорт на відновлення УБД" },
  { value: "form6Report", label: "Форма 6" },
  { value: "form12Report", label: "Форма 12" },
  { value: "salaryPowerAttorney", label: "Довіреність на зарплату" },
  { value: "default", label: "Звичайний рапорт" },
];

const emptyPreset = (): SaveDocumentSignatoryPreset => ({
  label: "",
  blockType: "SIGNER",
  title: "",
  rank: "",
  fullName: "",
  signatureData: null,
  signatureFileName: null,
  signatureMimeType: null,
  showDate: false,
  documentTypes: ["ubdReport"],
  sortOrder: 0,
});

const fileToDataUrl = (file: File) =>
  new Promise<string>((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result ?? ""));
    reader.onerror = () => reject(reader.error);
    reader.readAsDataURL(file);
  });

export function DocumentSignatoriesSettingsPage() {
  const [records, setRecords] = useState<BackendDocumentSignatoryPreset[]>([]);
  const [selectedId, setSelectedId] = useState("");
  const [draft, setDraft] = useState<SaveDocumentSignatoryPreset>(emptyPreset);
  const [message, setMessage] = useState("Завантажую записи...");
  const [messageType, setMessageType] =
    useState<"info" | "success" | "error">("info");
  const [isBusy, setIsBusy] = useState(false);
  const [isProcessingSignature, setIsProcessingSignature] = useState(false);

  const selectedRecord = useMemo(
    () => records.find((record) => record.id === selectedId) ?? null,
    [records, selectedId],
  );

  const loadRecords = async () => {
    try {
      const nextRecords = await api.listDocumentSignatories();
      setRecords(nextRecords);
      setMessage(
        nextRecords.length
          ? `Збережено записів: ${nextRecords.length}.`
          : "Записів ще немає. Створіть перший дефолтний блок.",
      );
      setMessageType("success");
    } catch (error) {
      setMessage(
        error instanceof Error ? error.message : "Не вдалося завантажити записи.",
      );
      setMessageType("error");
    }
  };

  useEffect(() => {
    void loadRecords();
  }, []);

  const selectRecord = (record: BackendDocumentSignatoryPreset) => {
    setSelectedId(record.id);
    setDraft({
      label: record.label,
      blockType: record.blockType,
      title: record.title,
      rank: record.rank,
      fullName: record.fullName,
      signatureData: record.signatureData ?? null,
      signatureFileName: record.signatureFileName ?? null,
      signatureMimeType: record.signatureMimeType ?? null,
      showDate: record.showDate,
      documentTypes: record.documentTypes,
      sortOrder: record.sortOrder,
    });
  };

  const createNew = () => {
    setSelectedId("");
    setDraft({
      ...emptyPreset(),
      sortOrder: records.length,
    });
    setMessage("Новий запис не збережено.");
    setMessageType("info");
  };

  const save = async () => {
    if (!draft.label.trim() || !draft.title.trim() || !draft.fullName.trim()) {
      setMessage("Заповніть назву запису, посаду/текст і ПІБ.");
      setMessageType("error");
      return;
    }
    if (!draft.documentTypes.length) {
      setMessage("Оберіть хоча б один тип документа.");
      setMessageType("error");
      return;
    }

    setIsBusy(true);
    try {
      const saved = selectedId
        ? await api.updateDocumentSignatory(selectedId, draft)
        : await api.createDocumentSignatory(draft);
      setRecords((current) =>
        [...current.filter((record) => record.id !== saved.id), saved].sort(
          (left, right) => left.sortOrder - right.sortOrder,
        ),
      );
      selectRecord(saved);
      setMessage("Дефолтний запис збережено.");
      setMessageType("success");
    } catch (error) {
      setMessage(
        error instanceof Error ? error.message : "Не вдалося зберегти запис.",
      );
      setMessageType("error");
    } finally {
      setIsBusy(false);
    }
  };

  const remove = async () => {
    if (!selectedRecord) return;
    if (!window.confirm(`Видалити запис «${selectedRecord.label}»?`)) return;
    setIsBusy(true);
    try {
      await api.deleteDocumentSignatory(selectedRecord.id);
      setRecords((current) =>
        current.filter((record) => record.id !== selectedRecord.id),
      );
      createNew();
      setMessage("Запис видалено.");
      setMessageType("success");
    } catch (error) {
      setMessage(
        error instanceof Error ? error.message : "Не вдалося видалити запис.",
      );
      setMessageType("error");
    } finally {
      setIsBusy(false);
    }
  };

  const toggleDocumentType = (documentType: string, checked: boolean) => {
    setDraft((current) => ({
      ...current,
      documentTypes: checked
        ? [...new Set([...current.documentTypes, documentType])]
        : current.documentTypes.filter((value) => value !== documentType),
    }));
  };

  return (
    <main className="main-panel document-signatories-settings-page">
      <header className="topbar analytics-topbar">
        <div>
          <Typography component="h1" variant="h4">
            Записи для документів
          </Typography>
          <Typography variant="body2" color="text.secondary">
            Один запис можна прив’язати до кількох типів документів. Під час
            створення рапорту він підтягується як незмінний дефолт.
          </Typography>
        </div>
        <Button
          variant="contained"
          startIcon={<SettingsOutlinedIcon />}
          onClick={createNew}
        >
          Додати запис
        </Button>
      </header>

      <Alert severity={messageType} variant="outlined">
        {message}
      </Alert>

      <UbdTemplateSettings />

      <section className="document-signatories-settings-layout">
        <aside className="analytics-panel document-signatories-list">
          <div className="panel-heading">Збережені записи · {records.length}</div>
          {records.length ? (
            records.map((record) => (
              <button
                className={
                  selectedId === record.id
                    ? "document-signatory-list-item active"
                    : "document-signatory-list-item"
                }
                key={record.id}
                onClick={() => selectRecord(record)}
                type="button"
              >
                <strong>{record.label}</strong>
                <span>{record.fullName}</span>
                <small>
                  {record.documentTypes
                    .map(
                      (type) =>
                        documentTypes.find((item) => item.value === type)?.label ??
                        type,
                    )
                    .join(" · ")}
                </small>
              </button>
            ))
          ) : (
            <p className="document-signatory-empty">Немає записів.</p>
          )}
        </aside>

        <section className="analytics-panel document-signatory-editor">
          <div className="panel-heading">
            {selectedRecord ? "Редагування запису" : "Новий запис"}
          </div>
          <div className="document-signatory-form-grid">
            <TextField
              label="Назва запису"
              value={draft.label}
              onChange={(event) =>
                setDraft((current) => ({ ...current, label: event.target.value }))
              }
            />
            <TextField
              label="Тип блока"
              select
              value={draft.blockType}
              onChange={(event) =>
                setDraft((current) => ({
                  ...current,
                  blockType: event.target.value as "SIGNER" | "APPROVAL",
                }))
              }
            >
              <MenuItem value="SIGNER">Підписант</MenuItem>
              <MenuItem value="APPROVAL">Затверджую</MenuItem>
            </TextField>
            <TextField
              className="wide"
              label="Посада / текст зліва"
              multiline
              rows={4}
              value={draft.title}
              onChange={(event) =>
                setDraft((current) => ({ ...current, title: event.target.value }))
              }
            />
            <TextField
              label="Військове звання"
              value={draft.rank}
              onChange={(event) =>
                setDraft((current) => ({ ...current, rank: event.target.value }))
              }
            />
            <TextField
              label="ПІБ підписанта"
              value={draft.fullName}
              onChange={(event) =>
                setDraft((current) => ({
                  ...current,
                  fullName: event.target.value,
                }))
              }
            />
            <TextField
              label="Порядок"
              type="number"
              value={draft.sortOrder}
              onChange={(event) =>
                setDraft((current) => ({
                  ...current,
                  sortOrder: Math.max(0, Number(event.target.value) || 0),
                }))
              }
            />
          </div>

          <div className="document-signatory-documents">
            <Typography variant="subtitle2">Використовувати у документах</Typography>
            <Stack direction="row" spacing={2}>
              {documentTypes.map((documentType) => (
                <Checkbox
                  key={documentType.value}
                  label={documentType.label}
                  checked={draft.documentTypes.includes(documentType.value)}
                  onCheckedChange={(checked) =>
                    toggleDocumentType(documentType.value, checked === true)
                  }
                />
              ))}
            </Stack>
          </div>

          <div className="document-signatory-signature">
            <div>
              <Typography variant="subtitle2">Зображення підпису</Typography>
              <Typography variant="caption" color="text.secondary">
                PNG або JPEG. Білий фон у документі буде оброблено як прозорий.
              </Typography>
            </div>
            <Button
              component="label"
              variant="outlined"
              disabled={isProcessingSignature}
            >
              {isProcessingSignature ? "Обробляю фон..." : "Вибрати файл"}
              <input
                hidden
                accept="image/png,image/jpeg"
                type="file"
                onChange={(event) => {
                  const file = event.target.files?.[0];
                  if (!file) return;
                  setIsProcessingSignature(true);
                  void fileToDataUrl(file)
                    .then((dataUrl) =>
                      processSignatureTransparentBackground(dataUrl),
                    )
                    .then((signatureData) => {
                      setDraft((current) => ({
                        ...current,
                        signatureData,
                        signatureFileName: file.name.replace(
                          /\.[^.]+$/,
                          ".png",
                        ),
                        signatureMimeType: "image/png",
                      }));
                      setMessage(
                        "Білий фон підпису видалено. Збережіть запис.",
                      );
                      setMessageType("success");
                    })
                    .catch(() => {
                      setMessage("Не вдалося обробити зображення підпису.");
                      setMessageType("error");
                    })
                    .finally(() => setIsProcessingSignature(false));
                  event.target.value = "";
                }}
              />
            </Button>
            {draft.signatureData ? (
              <div className="document-signatory-signature-preview">
                <img alt="Підпис" src={draft.signatureData} />
                <button
                  onClick={() =>
                    setDraft((current) => ({
                      ...current,
                      signatureData: null,
                      signatureFileName: null,
                      signatureMimeType: null,
                    }))
                  }
                  type="button"
                >
                  Прибрати
                </button>
              </div>
            ) : null}
          </div>

          <Stack direction="row" spacing={1}>
            <Button
              variant="contained"
              startIcon={<CheckCircleOutlineIcon />}
              disabled={isBusy || isProcessingSignature}
              onClick={() => void save()}
            >
              Зберегти
            </Button>
            {selectedRecord ? (
              <Button
                variant="outlined"
                startIcon={<DeleteOutlineOutlinedIcon />}
                disabled={isBusy}
                onClick={() => void remove()}
              >
                Видалити
              </Button>
            ) : null}
          </Stack>
        </section>
      </section>
    </main>
  );
}
