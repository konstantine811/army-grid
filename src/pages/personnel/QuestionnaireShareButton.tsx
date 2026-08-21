import { useEffect, useRef, useState } from "react";
import { Button } from "@/components/sci/SciPrimitives";
import { ShareOutlinedIcon } from "@/components/sci/icons";
import {
  canNativeShareQuestionnaire,
  nativeShareQuestionnaire,
  resolveQuestionnairePdfFile,
  shareQuestionnaireViaMessenger,
  type QuestionnairePdfSource,
} from "./questionnaireShare";

type QuestionnaireShareButtonProps = {
  disabled?: boolean;
  fileName: string;
  personName: string;
  source: QuestionnairePdfSource | null;
  onNotify?: (message: string) => void;
};

export function QuestionnaireShareButton({
  disabled,
  fileName,
  personName,
  source,
  onNotify,
}: QuestionnaireShareButtonProps) {
  const [menuOpen, setMenuOpen] = useState(false);
  const rootRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (!menuOpen) return;

    const handlePointerDown = (event: MouseEvent) => {
      if (!rootRef.current?.contains(event.target as Node)) {
        setMenuOpen(false);
      }
    };

    document.addEventListener("mousedown", handlePointerDown);
    return () => document.removeEventListener("mousedown", handlePointerDown);
  }, [menuOpen]);

  const resolvedFile = source
    ? resolveQuestionnairePdfFile(fileName, source)
    : null;
  const canNativeShare = resolvedFile
    ? canNativeShareQuestionnaire(resolvedFile)
    : false;

  const handleNativeShare = async () => {
    if (!resolvedFile) return;
    try {
      await nativeShareQuestionnaire(resolvedFile, personName);
      onNotify?.("Відкрито системне меню (AirDrop, Messages…).");
      setMenuOpen(false);
    } catch (error) {
      if (error instanceof DOMException && error.name === "AbortError") return;
      onNotify?.(
        error instanceof Error
          ? error.message
          : "Не вдалося надіслати анкету.",
      );
    }
  };

  const handleMessengerShare = (messenger: "whatsapp" | "telegram") => {
    if (!source) return;
    try {
      shareQuestionnaireViaMessenger(messenger, fileName, source, personName);
      onNotify?.(
        messenger === "whatsapp"
          ? "PDF завантажено. Відкрито WhatsApp — прикріпіть файл до чату."
          : "PDF завантажено. Відкрито Telegram — прикріпіть файл до чату.",
      );
      setMenuOpen(false);
    } catch (error) {
      onNotify?.(
        error instanceof Error
          ? error.message
          : "Не вдалося підготувати надсилання.",
      );
    }
  };

  return (
    <div className="questionnaire-share-wrap" ref={rootRef}>
      <Button
        size="small"
        variant="outlined"
        disabled={disabled || !resolvedFile}
        startIcon={<ShareOutlinedIcon />}
        onClick={() => setMenuOpen((current) => !current)}
      >
        Надіслати
      </Button>
      {menuOpen ? (
        <div className="questionnaire-share-menu" role="menu">
          <button
            type="button"
            className="questionnaire-share-menu-item questionnaire-share-menu-item-primary"
            onClick={() => handleMessengerShare("whatsapp")}
          >
            WhatsApp
          </button>
          <button
            type="button"
            className="questionnaire-share-menu-item"
            onClick={() => handleMessengerShare("telegram")}
          >
            Telegram
          </button>
          {canNativeShare ? (
            <button
              type="button"
              className="questionnaire-share-menu-item"
              onClick={() => void handleNativeShare()}
            >
              AirDrop / Messages…
            </button>
          ) : null}
        </div>
      ) : null}
    </div>
  );
}
