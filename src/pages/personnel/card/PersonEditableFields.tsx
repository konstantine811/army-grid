import { anketaCardFieldValue } from "../../personnel-v2/personnelV2AnketaFill";
import { computeFullYearsFromBirthDate } from "../personnelUtils";
import type { AnketaRow, EditablePersonSection } from "./personCardTypes";

export function PersonEditableFields({
  sections,
  anketa,
  editValues,
  birthDate,
  onChange,
  onSave,
}: {
  sections: EditablePersonSection[];
  anketa: AnketaRow | null;
  editValues: Record<string, string>;
  birthDate: string;
  onChange: (key: string, value: string) => void;
  onSave: () => void;
}) {
  return (
    <>
      {sections.map((group) => (
        <div className="person-edit-section" key={group.section}>
          <div className="panel-heading">{group.label}</div>
          <div className="person-edit-grid">
            {anketa
              ? group.fields.map((field) => (
                  <PersonAnketaField
                    key={field.label}
                    label={field.label}
                    wide={
                      field.kind === "multiline" || field.section === "contacts"
                    }
                    value={anketaCardFieldValue(anketa, field.label)}
                  />
                ))
              : group.fields.map((field) => (
                  <PersonEditField
                    key={field.key}
                    field={field}
                    value={editValues[field.key] ?? ""}
                    birthDate={birthDate}
                    onChange={onChange}
                    onSave={onSave}
                  />
                ))}
          </div>
        </div>
      ))}
    </>
  );
}

function PersonAnketaField({
  label,
  value,
  wide,
}: {
  label: string;
  value: string;
  wide: boolean;
}) {
  return (
    <span className={wide ? "wide" : undefined}>
      <strong>{label}</strong>
      {value || "—"}
    </span>
  );
}

function PersonEditField({
  field,
  value,
  birthDate,
  onChange,
  onSave,
}: {
  field: EditablePersonSection["fields"][number];
  value: string;
  birthDate: string;
  onChange: (key: string, value: string) => void;
  onSave: () => void;
}) {
  const wide =
    field.kind === "multiline" ||
    field.section === "contacts" ||
    field.parts.includes("додаткова_інформація");
  const years = field.parts.includes("дата_народження")
    ? computeFullYearsFromBirthDate(value || birthDate)
    : null;

  return (
    <label className={wide ? "wide" : ""}>
      <span>
        {field.label}
        {years != null ? ` · ${years} р.` : ""}
      </span>
      {field.kind === "multiline" ? (
        <textarea
          className="sci-message-area"
          value={value}
          onChange={(event) => onChange(field.key, event.target.value)}
          onBlur={() => void onSave()}
        />
      ) : (
        <input
          value={value}
          onChange={(event) => onChange(field.key, event.target.value)}
          onBlur={() => void onSave()}
        />
      )}
    </label>
  );
}
