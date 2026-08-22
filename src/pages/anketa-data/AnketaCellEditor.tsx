import { useState } from "react";
import type { AnketaColumnKey } from "./anketaSheet";
import { AnketaMissingPresets } from "./AnketaMissingPresets";

const MULTILINE_COLUMNS = new Set<AnketaColumnKey>([
  "additionalInfo",
  "relatives",
]);

type AnketaCellEditorProps = {
  columnKey: AnketaColumnKey;
  columnHeader: string;
  rowNumber: number;
  value: string;
  isEmpty: boolean;
  advanceOnSave: boolean;
  onSave: (value: string, advance: boolean) => void;
  onCancel: () => void;
};

export function AnketaCellEditor({
  columnKey,
  columnHeader,
  rowNumber,
  value,
  isEmpty,
  advanceOnSave,
  onSave,
  onCancel,
}: AnketaCellEditorProps) {
  const [draft, setDraft] = useState(value);
  const isMultiline = MULTILINE_COLUMNS.has(columnKey);
  const isDirty = draft !== value;

  const commit = (advance: boolean) => {
    if (!isDirty && !advance) {
      onCancel();
      return;
    }
    onSave(draft, advance);
  };

  const sharedProps = {
    className: "anketa-cell-input is-active",
    value: draft,
    autoFocus: true,
    "aria-label": `${columnHeader} · рядок ${rowNumber}`,
    onChange: (
      event: React.ChangeEvent<HTMLInputElement | HTMLTextAreaElement>,
    ) => {
      setDraft(event.currentTarget.value);
    },
    onKeyDown: (
      event: React.KeyboardEvent<HTMLInputElement | HTMLTextAreaElement>,
    ) => {
      if (event.key === "Escape") {
        event.preventDefault();
        event.stopPropagation();
        onCancel();
        return;
      }
      const confirmWithAdvance =
        event.key === "Enter" &&
        (!isMultiline || event.shiftKey || event.metaKey || event.ctrlKey);
      if (!confirmWithAdvance) return;
      event.preventDefault();
      commit(advanceOnSave);
    },
  } as const;

  return (
    <div
      className={
        isMultiline ? "anketa-cell-editor is-multiline" : "anketa-cell-editor"
      }
    >
      {isMultiline ? (
        <textarea
          {...sharedProps}
          className="anketa-cell-input anketa-cell-textarea is-active"
          rows={4}
        />
      ) : (
        <input {...sharedProps} />
      )}
      <div className="anketa-cell-editor-actions">
        <button
          type="button"
          className="anketa-cell-action is-primary"
          onMouseDown={(event) => event.preventDefault()}
          onClick={() => commit(advanceOnSave)}
        >
          Зберегти{advanceOnSave ? " і далі" : ""}
        </button>
        <button
          type="button"
          className="anketa-cell-action"
          onMouseDown={(event) => event.preventDefault()}
          onClick={onCancel}
        >
          Скасувати
        </button>
      </div>
      {isMultiline ? (
        <p className="anketa-cell-editor-hint">
          Enter — новий рядок · Shift+Enter — зберегти
          {advanceOnSave ? " і далі" : ""}
        </p>
      ) : null}
      {isEmpty ? (
        <AnketaMissingPresets
          onPick={(preset) => {
            onSave(preset, advanceOnSave);
          }}
        />
      ) : null}
    </div>
  );
}
