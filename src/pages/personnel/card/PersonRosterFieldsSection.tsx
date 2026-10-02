import type { PersonnelRosterCardField } from "./personCardTypes";

export function PersonRosterFieldsSection({
  fields,
}: {
  fields: PersonnelRosterCardField[];
}) {
  if (!fields.length) return null;
  return (
    <div className="person-edit-section">
      <div className="panel-heading">Загальний список</div>
      <div className="person-roster-grid">
        {fields.map((field) => (
          <span key={field.key}>
            <strong>{field.label}</strong>
            {field.value}
          </span>
        ))}
      </div>
    </div>
  );
}
