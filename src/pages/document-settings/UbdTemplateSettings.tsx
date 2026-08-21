import { useEffect, useState } from "react";
import { Alert, Button, Typography } from "@/components/sci/SciPrimitives";
import { UBD_TEMPLATE_FIELDS } from "../documents/ubdWordExport";
import {
  clearCustomUbdTemplate,
  hasCustomUbdTemplate,
  loadUbdTemplate,
  saveCustomUbdTemplate,
} from "../documents/ubdTemplateStore";

const downloadBuffer = (buffer: ArrayBuffer, fileName: string) => {
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  link.click();
  URL.revokeObjectURL(url);
};

export function UbdTemplateSettings() {
  const [isCustom, setIsCustom] = useState(false);
  const [message, setMessage] = useState("");
  const [messageType, setMessageType] =
    useState<"info" | "success" | "error">("info");
  const [isBusy, setIsBusy] = useState(false);

  const refresh = async () => {
    setIsCustom(await hasCustomUbdTemplate());
  };

  useEffect(() => {
    void refresh();
  }, []);

  const downloadTemplate = async () => {
    setIsBusy(true);
    try {
      const buffer = await loadUbdTemplate();
      downloadBuffer(buffer, "ubd-report-template.docx");
      setMessage("Шаблон завантажено. Відкрийте його в Word і змініть верстку.");
      setMessageType("success");
    } catch (error) {
      setMessage(
        error instanceof Error ? error.message : "Не вдалося завантажити шаблон.",
      );
      setMessageType("error");
    } finally {
      setIsBusy(false);
    }
  };

  const uploadTemplate = async (file: File) => {
    setIsBusy(true);
    try {
      const buffer = await file.arrayBuffer();
      const JSZip = (await import("jszip")).default;
      const zip = await JSZip.loadAsync(buffer);
      const documentXml = await zip.file("word/document.xml")?.async("string");
      if (!documentXml) {
        throw new Error("Це не дійсний Word-файл.");
      }
      const plain = documentXml.replace(/<[^>]+>/g, "");
      if (!plain.includes("{{COMMANDER}}") || !plain.includes("{{RANK}}")) {
        throw new Error(
          "У шаблоні немає міток {{COMMANDER}} і {{RANK}}. Завантажте стандартний шаблон і не видаляйте мітки.",
        );
      }
      await saveCustomUbdTemplate(buffer);
      await refresh();
      setMessage("Шаблон збережено. Рапорти УБД тепер беруть цю верстку.");
      setMessageType("success");
    } catch (error) {
      setMessage(
        error instanceof Error ? error.message : "Не вдалося зберегти шаблон.",
      );
      setMessageType("error");
    } finally {
      setIsBusy(false);
    }
  };

  const resetTemplate = async () => {
    setIsBusy(true);
    try {
      await clearCustomUbdTemplate();
      await refresh();
      setMessage("Повернуто стандартний шаблон.");
      setMessageType("success");
    } catch (error) {
      setMessage(
        error instanceof Error ? error.message : "Не вдалося скинути шаблон.",
      );
      setMessageType("error");
    } finally {
      setIsBusy(false);
    }
  };

  return (
    <section className="analytics-panel ubd-template-settings">
      <div className="panel-heading">Шаблон рапорту УБД</div>
      <Typography variant="body2" color="text.secondary">
        Верстка живе в Word-шаблоні. Програма лише підставляє мітки. Дата підпису
        командира не змінюється — її ставте в Word. Завантажте шаблон, поправте
        верстку й завантажте назад — прев’ю й експорт візьмуть той самий файл.
      </Typography>
      <ol className="ubd-template-steps">
        <li>Завантажте шаблон і відкрийте його в Word.</li>
        <li>Посуньте текст, таблицю чи підпис як потрібно. Мітки не видаляйте.</li>
        <li>Збережіть .docx і завантажте його назад сюди.</li>
      </ol>
      <div className="ubd-template-actions">
        <Button variant="contained" disabled={isBusy} onClick={() => void downloadTemplate()}>
          Завантажити шаблон
        </Button>
        <Button component="label" variant="outlined" disabled={isBusy}>
          Замінити шаблон
          <input
            hidden
            accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            type="file"
            onChange={(event) => {
              const file = event.target.files?.[0];
              if (file) void uploadTemplate(file);
              event.target.value = "";
            }}
          />
        </Button>
        {isCustom ? (
          <Button variant="outlined" disabled={isBusy} onClick={() => void resetTemplate()}>
            Повернути стандартний
          </Button>
        ) : null}
      </div>
      <Typography variant="caption" color="text.secondary">
        Зараз: {isCustom ? "ваш завантажений шаблон" : "стандартний шаблон зі зразка"}.
      </Typography>
      {message ? (
        <Alert severity={messageType} variant="outlined">
          {message}
        </Alert>
      ) : null}
      <details className="ubd-template-fields">
        <summary>Мітки, які підставляє програма</summary>
        <ul>
          {UBD_TEMPLATE_FIELDS.map((field) => (
            <li key={field.key}>
              <code>{`{{${field.key}}}`}</code>
              <span>{field.label}</span>
            </li>
          ))}
        </ul>
      </details>
    </section>
  );
}
