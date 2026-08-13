import { useLayoutEffect, useRef } from "react";
import type { ExcelCellEditorProps } from "./types";

export function ExcelCellEditor({ value, wrapText, onCommit }: ExcelCellEditorProps) {
  const textareaRef = useRef<HTMLTextAreaElement | null>(null);

  const resizeTextarea = () => {
    const textarea = textareaRef.current;
    if (!textarea) return;

    textarea.style.height = "auto";
    textarea.style.height = `${textarea.scrollHeight}px`;
  };

  useLayoutEffect(() => {
    if (wrapText) resizeTextarea();
  }, [value, wrapText]);

  if (wrapText) {
    return (
      <textarea
        ref={textareaRef}
        className="excel-cell-input excel-cell-input-wrap"
        defaultValue={value}
        title={value}
        rows={1}
        onInput={resizeTextarea}
        onBlur={(event) => onCommit(event.target.value)}
      />
    );
  }

  return (
    <input
      className="excel-cell-input"
      defaultValue={value}
      title={value}
      onBlur={(event) => onCommit(event.target.value)}
    />
  );
}
