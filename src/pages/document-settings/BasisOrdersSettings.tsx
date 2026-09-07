import { useEffect, useRef, useState } from "react";
import {
  Alert,
  Button,
  Stack,
  TextField,
  Typography,
} from "@/components/sci/SciPrimitives";
import { DeleteOutlineOutlinedIcon, FileUploadOutlinedIcon } from "@/components/sci/icons";
import {
  createBasisOrderId,
  importBasisOrdersFromParsed,
  loadCustomBasisOrders,
  loadImportedBasisOrders,
  saveCustomBasisOrders,
  type UbdBasisOrderRecord,
} from "../documents/ubdBasisOrdersDirectory";
import { parseBasisOrdersFromDocx } from "../documents/ubdBasisOrdersImport";

const emptyDraft = (): UbdBasisOrderRecord => ({
  id: "",
  number: "",
  date: "",
  location: "",
  validFrom: "",
  validTo: "",
  note: "",
});

export function BasisOrdersSettings() {
  const [rows, setRows] = useState<UbdBasisOrderRecord[]>([]);
  const [importedCount, setImportedCount] = useState(0);
  const [draft, setDraft] = useState<UbdBasisOrderRecord>(emptyDraft);
  const [message, setMessage] = useState("");
  const [messageType, setMessageType] = useState<"success" | "error">("success");
  const importInputRef = useRef<HTMLInputElement>(null);
  const [isImporting, setIsImporting] = useState(false);

  useEffect(() => {
    setRows(loadCustomBasisOrders());
    setImportedCount(loadImportedBasisOrders().length);
  }, []);

  const persist = (next: UbdBasisOrderRecord[]) => {
    setRows(next);
    saveCustomBasisOrders(next);
  };

  const saveDraft = () => {
    const number = draft.number.trim();
    const date = draft.date.trim();
    const location = draft.location?.trim() ?? "";
    if (!number || !date || !location) {
      setMessageType("error");
      setMessage("Потрібні номер БР, дата БР і локація.");
      return;
    }
    const record: UbdBasisOrderRecord = {
      ...draft,
      id: draft.id || createBasisOrderId(),
      number,
      date,
      location,
      validFrom: draft.validFrom?.trim() || date,
      validTo: draft.validTo?.trim() ?? "",
      note: draft.note?.trim() ?? "",
    };
    persist(
      draft.id
        ? rows.map((row) => (row.id === draft.id ? record : row))
        : [...rows, record],
    );
    setDraft(emptyDraft());
    setMessageType("success");
    setMessage("Довідник БР збережено. У рапорті номер підставиться за локацією і датою.");
  };

  const importFromDocx = async (file: File) => {
    setIsImporting(true);
    try {
      const parsed = await parseBasisOrdersFromDocx(await file.arrayBuffer());
      if (!parsed.length) {
        setMessageType("error");
        setMessage(
          "У файлі не знайдено рядків формату «№4862/ОКП/123/дск від 01.08.2026».",
        );
        return;
      }
      const result = importBasisOrdersFromParsed(parsed, file.name);
      setImportedCount(result.total);
      setMessageType("success");
      setMessage(
        `Імпортовано з «${file.name}»: додано ${result.added}, пропущено ${result.skipped} (вже були). Усього імпортованих: ${result.total}.`,
      );
    } catch (error) {
      setMessageType("error");
      setMessage(
        error instanceof Error
          ? error.message
          : "Не вдалося імпортувати Word-файл із номерами БР.",
      );
    } finally {
      setIsImporting(false);
      if (importInputRef.current) importInputRef.current.value = "";
    }
  };

  return (
    <section className="analytics-panel basis-orders-settings">
      <div className="panel-heading">Бойові розпорядження</div>
      <Typography variant="body2" color="text.secondary">
        Не прив’язуйте № БР лише до локації: для одного місця з часом може бути
        кілька розпоряджень. Правило: локація + дата виконання → № БР. Якщо
        вихід перекриває зміну БР, підставляються всі відповідні номери. Поле в
        рапорті лишається доступним для ручної правки.
      </Typography>

      {message ? (
        <Alert severity={messageType} variant="outlined">
          {message}
        </Alert>
      ) : null}

      <Stack direction="row" spacing={1} className="basis-orders-import-row">
        <input
          ref={importInputRef}
          type="file"
          accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
          hidden
          onChange={(event) => {
            const file = event.target.files?.[0];
            if (file) void importFromDocx(file);
          }}
        />
        <Button
          variant="outlined"
          startIcon={<FileUploadOutlinedIcon />}
          disabled={isImporting}
          onClick={() => importInputRef.current?.click()}
        >
          {isImporting ? "Імпорт…" : "Імпорт з Word (.docx)"}
        </Button>
        {importedCount ? (
          <Typography variant="body2" color="text.secondary">
            Імпортовано без локації: {importedCount}
          </Typography>
        ) : null}
      </Stack>
      <Typography variant="caption" color="text.secondary" component="p">
        Формат файлу: один рядок на БР, напр. «№4862/ОКП/2223/дск від
        01.08.2026». Нові номери додаються до існуючих, дублікати пропускаються.
      </Typography>

      <div className="basis-orders-form">
        <TextField
          label="Номер БР"
          value={draft.number}
          placeholder="4862/ОКП/145/дск"
          onChange={(event) =>
            setDraft((current) => ({ ...current, number: event.target.value }))
          }
        />
        <TextField
          label="Дата БР"
          value={draft.date}
          placeholder="11.08.2026"
          onChange={(event) =>
            setDraft((current) => ({ ...current, date: event.target.value }))
          }
        />
        <TextField
          label="Локація / район"
          value={draft.location}
          placeholder="Петропавлівка"
          onChange={(event) =>
            setDraft((current) => ({ ...current, location: event.target.value }))
          }
        />
        <TextField
          label="Діє з"
          value={draft.validFrom}
          placeholder="11.08.2026"
          onChange={(event) =>
            setDraft((current) => ({ ...current, validFrom: event.target.value }))
          }
        />
        <TextField
          label="Діє по"
          value={draft.validTo}
          placeholder="25.08.2026"
          onChange={(event) =>
            setDraft((current) => ({ ...current, validTo: event.target.value }))
          }
        />
        <TextField
          className="wide"
          label="Примітка"
          value={draft.note}
          onChange={(event) =>
            setDraft((current) => ({ ...current, note: event.target.value }))
          }
        />
      </div>
      <Stack direction="row" spacing={1}>
        <Button variant="contained" onClick={saveDraft}>
          {draft.id ? "Оновити запис" : "Додати до довідника"}
        </Button>
        {draft.id ? (
          <Button variant="outlined" onClick={() => setDraft(emptyDraft())}>
            Скасувати
          </Button>
        ) : null}
      </Stack>

      {rows.length ? (
        <div className="basis-orders-table">
          {rows.map((row) => (
            <article key={row.id} className="basis-orders-row">
              <strong>{row.number}</strong>
              <span>{row.location}</span>
              <small>
                БР {row.date}
                {row.validFrom || row.validTo
                  ? ` · діє ${row.validFrom || row.date}${
                      row.validTo ? `–${row.validTo}` : ""
                    }`
                  : ""}
                {row.note ? ` · ${row.note}` : ""}
              </small>
              <span className="basis-orders-row-actions">
                <Button
                  size="small"
                  variant="outlined"
                  onClick={() => setDraft(row)}
                >
                  Змінити
                </Button>
                <Button
                  size="small"
                  variant="outlined"
                  startIcon={<DeleteOutlineOutlinedIcon />}
                  onClick={() => {
                    persist(rows.filter((item) => item.id !== row.id));
                    if (draft.id === row.id) setDraft(emptyDraft());
                  }}
                >
                  Видалити
                </Button>
              </span>
            </article>
          ))}
        </div>
      ) : (
        <p className="document-signatory-empty">
          Довідник порожній. Додайте БР з локацією і періодом дії — тоді в УБД /
          Формі 6 номер підставиться сам.
        </p>
      )}
    </section>
  );
}
