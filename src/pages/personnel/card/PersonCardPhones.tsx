import { Button, Chip } from "@/components/sci/SciPrimitives";
import { ContentCopyOutlinedIcon } from "@/components/sci/icons";
import { DeleteOutlineOutlinedIcon } from "@/components/sci/icons";
import { formatUaPhoneDisplay } from "../personnelUtils";

export function PersonCardPhones({
  phones,
  savedPhones,
  draft,
  disabled,
  onDraftChange,
  onAdd,
  onCopy,
  onRemove,
}: {
  phones: string[];
  savedPhones: string[];
  draft: string;
  disabled: boolean;
  onDraftChange: (value: string) => void;
  onAdd: () => void;
  onCopy: (phone: string) => void;
  onRemove: (phone: string) => void;
}) {
  return (
    <span className="wide person-phones-summary">
      <strong>Телефони</strong>
      {phones.length > 0 ? (
        <span className="person-phones-list">
          {phones.map((phone) => (
            <PersonPhoneChip
              key={phone}
              phone={phone}
              saved={savedPhones.includes(phone)}
              removeDisabled={disabled}
              onCopy={onCopy}
              onRemove={onRemove}
            />
          ))}
        </span>
      ) : (
        <span className="person-phones-empty">Номерів ще немає</span>
      )}
      <div className="person-phones-editor">
        <input
          aria-label="Номер телефону"
          autoComplete="off"
          disabled={disabled}
          inputMode="tel"
          placeholder="063 123 45 67"
          value={draft}
          onChange={(event) => onDraftChange(event.target.value)}
          onKeyDown={(event) => {
            if (event.key !== "Enter") return;
            event.preventDefault();
            void onAdd();
          }}
        />
        <Button
          disabled={disabled}
          size="small"
          type="button"
          variant="contained"
          sx={{ color: "#1a1a14" }}
          onClick={() => void onAdd()}
        >
          Додати
        </Button>
      </div>
    </span>
  );
}

function PersonPhoneChip({
  phone,
  saved,
  removeDisabled,
  onCopy,
  onRemove,
}: {
  phone: string;
  saved: boolean;
  removeDisabled: boolean;
  onCopy: (phone: string) => void;
  onRemove: (phone: string) => void;
}) {
  const label = formatUaPhoneDisplay(phone);
  return (
    <span className="person-phone-chip">
      <Chip label={label} size="small" color="primary" variant="outlined" />
      <button
        aria-label={`Копіювати ${label}`}
        title="Копіювати номер"
        type="button"
        onClick={() => void onCopy(phone)}
      >
        <ContentCopyOutlinedIcon fontSize="small" />
      </button>
      {saved ? (
        <button
          aria-label={`Видалити ${label}`}
          disabled={removeDisabled}
          type="button"
          onClick={() => void onRemove(phone)}
        >
          <DeleteOutlineOutlinedIcon fontSize="small" />
        </button>
      ) : null}
    </span>
  );
}
