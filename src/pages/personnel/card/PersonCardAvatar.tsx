import { Button } from "@/components/sci/SciPrimitives";
import { AddPhotoAlternateOutlinedIcon } from "@/components/sci/icons";
import { DeleteOutlineOutlinedIcon } from "@/components/sci/icons";
import { PersonSearchOutlinedIcon } from "@/components/sci/icons";
import { SearchOutlinedIcon } from "@/components/sci/icons";

export function PersonCardAvatar({
  photo,
  name,
  canOpen,
  hasRow,
  hasQuestionnaire,
  onOpenPhoto,
  onMissingQuestionnaire,
  onOpenQuestionnaire,
  onDeletePhoto,
  onAddPhoto,
}: {
  photo: string;
  name: string;
  canOpen: boolean;
  hasRow: boolean;
  hasQuestionnaire: boolean;
  onOpenPhoto: () => void;
  onMissingQuestionnaire: () => void;
  onOpenQuestionnaire: () => void;
  onDeletePhoto: () => void;
  onAddPhoto: (file: File | undefined) => void;
}) {
  return (
    <div className="person-avatar">
      {photo ? (
        <img alt={name} src={photo} onClick={onOpenPhoto} />
      ) : (
        <PersonSearchOutlinedIcon />
      )}
      <button
        aria-label="Відкрити анкету"
        className="person-avatar-zoom"
        disabled={!canOpen}
        onClick={() => {
          if (!hasQuestionnaire) {
            onMissingQuestionnaire();
            return;
          }
          void onOpenQuestionnaire();
        }}
        title="Відкрити анкету"
        type="button"
      >
        <SearchOutlinedIcon />
      </button>
      {photo ? (
        <button
          aria-label="Видалити фото"
          className="person-avatar-delete"
          disabled={!hasRow}
          onClick={() => void onDeletePhoto()}
          title="Видалити фото"
          type="button"
        >
          <DeleteOutlineOutlinedIcon />
        </button>
      ) : null}
      <Button
        aria-label="Додати фото"
        className="person-avatar-upload"
        component="label"
        disabled={!hasRow}
        size="small"
        startIcon={<AddPhotoAlternateOutlinedIcon />}
        title="Додати фото"
        variant="contained"
        sx={{ color: "#1a1a14" }}
      >
        Фото
        <input
          hidden
          type="file"
          accept="image/*,application/pdf"
          onChange={(event) => {
            onAddPhoto(event.target.files?.[0]);
            event.target.value = "";
          }}
        />
      </Button>
    </div>
  );
}
