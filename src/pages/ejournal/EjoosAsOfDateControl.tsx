import { useEffect, useRef, useState } from "react";
import { Button } from "@/components/sci/SciPrimitives";
import {
  formatUaDateTyping,
  isoDateToUaLabel,
  isCompleteUaDate,
  uaDateToIso,
  uaLabelToIsoDate,
} from "./ejoosAsOfDate";

type EjoosAsOfDateControlProps = {
  /** Поточна застосована дата `DD.MM.YYYY`. */
  appliedLabel: string;
  /** Блокувати лише кнопку OK під час перебудови (поле лишається активним). */
  busy?: boolean;
  /** ISO `YYYY-MM-DD` після підтвердження. */
  onApply: (isoDate: string) => void | Promise<void>;
};

let asOfControlSeq = 0;

export function EjoosAsOfDateControl({
  appliedLabel,
  busy = false,
  onApply,
}: EjoosAsOfDateControlProps) {
  const controlId = useRef(`ejoos-asof-${++asOfControlSeq}`).current;
  const textId = `${controlId}-text`;
  const [draft, setDraft] = useState(() => appliedLabel || "");
  const [isEditing, setIsEditing] = useState(false);
  const [error, setError] = useState("");

  useEffect(() => {
    if (isEditing) return;
    setDraft(appliedLabel || "");
    setError("");
  }, [appliedLabel, isEditing]);

  const applyDraft = async (nextDraft = draft) => {
    const trimmed = nextDraft.trim();
    if (!trimmed) {
      setError("Вкажіть дату");
      return;
    }
    if (!isCompleteUaDate(trimmed)) {
      setError("Формат: ДД.ММ.РРРР");
      return;
    }
    const iso = uaDateToIso(trimmed);
    if (!iso) {
      setError("Некоректна дата");
      return;
    }
    setError("");
    setIsEditing(false);
    await onApply(iso);
  };

  const pickerIso = uaLabelToIsoDate(draft) || uaLabelToIsoDate(appliedLabel);

  return (
    <div
      className="ejoos-asof-date"
      onMouseDown={(event) => event.stopPropagation()}
      onClick={(event) => event.stopPropagation()}
    >
      <input
        id={textId}
        className="ejoos-search ejoos-asof-date-text"
        type="text"
        inputMode="decimal"
        autoComplete="off"
        spellCheck={false}
        placeholder="ДД.ММ.РРРР"
        value={draft}
        aria-label="Станом на, формат ДД.ММ.РРРР"
        onFocus={() => setIsEditing(true)}
        onBlur={() => {
          window.setTimeout(() => setIsEditing(false), 120);
        }}
        onChange={(event) => {
          setError("");
          setDraft(formatUaDateTyping(event.target.value));
        }}
        onKeyDown={(event) => {
          if (event.key === "Enter") {
            event.preventDefault();
            void applyDraft();
          }
        }}
      />
      <span className="ejoos-asof-date-calendar-wrap" title="Календар">
        <input
          type="date"
          className="ejoos-asof-date-picker-overlay"
          value={pickerIso}
          aria-label="Календар"
          tabIndex={-1}
          onChange={(event) => {
            const iso = event.target.value;
            if (!iso) return;
            const label = isoDateToUaLabel(iso);
            setDraft(label);
            setError("");
            setIsEditing(false);
            void onApply(iso);
          }}
        />
        <button
          type="button"
          className="ejoos-asof-date-calendar"
          aria-hidden
          tabIndex={-1}
        >
          📅
        </button>
      </span>
      <Button
        type="button"
        size="small"
        variant="contained"
        disabled={busy}
        onClick={() => void applyDraft()}
        sx={{ color: "#1a1a14" }}
      >
        OK
      </Button>
      {error ? (
        <span className="ejoos-asof-date-error" role="alert">
          {error}
        </span>
      ) : null}
    </div>
  );
}
