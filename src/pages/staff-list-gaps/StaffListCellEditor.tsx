import { useState } from "react";

type StaffListCellEditorProps = {
  columnHeader: string;
  pib: string;
  value: string;
  disabled?: boolean;
  advanceOnSave: boolean;
  onSave: (value: string, advance: boolean) => void;
  onCancel: () => void;
};

export function StaffListCellEditor({
  columnHeader,
  pib,
  value,
  disabled = false,
  advanceOnSave,
  onSave,
  onCancel,
}: StaffListCellEditorProps) {
  const [draft, setDraft] = useState(value);
  const isDirty = draft !== value;

  const commit = (advance: boolean) => {
    if (disabled) return;
    if (!isDirty && !advance) {
      onCancel();
      return;
    }
    onSave(draft, advance);
  };

  return (
    <div className="anketa-cell-editor is-multiline staff-list-cell-editor">
      <textarea
        className="anketa-cell-input anketa-cell-textarea is-active"
        value={draft}
        disabled={disabled}
        rows={3}
        aria-label={`${columnHeader} · ${pib}`}
        placeholder="Введіть значення з анкети"
        onChange={(event) => setDraft(event.currentTarget.value)}
        onKeyDown={(event) => {
          if (event.key === "Escape") {
            event.preventDefault();
            event.stopPropagation();
            onCancel();
            return;
          }
          if (
            event.key === "Enter" &&
            (event.shiftKey || event.metaKey || event.ctrlKey)
          ) {
            event.preventDefault();
            commit(advanceOnSave);
          }
        }}
      />
      <div className="anketa-cell-editor-actions">
        <button
          type="button"
          className="anketa-cell-action is-primary"
          disabled={disabled}
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
      <p className="anketa-cell-editor-hint">
        Enter — новий рядок · Shift+Enter — зберегти
        {advanceOnSave ? " і далі" : ""}
      </p>
    </div>
  );
}
