export function PersonPhotoLightbox({
  name,
  photo,
  onClose,
}: {
  name: string;
  photo: string;
  onClose: () => void;
}) {
  return (
    <div
      className="person-photo-lightbox"
      role="dialog"
      aria-modal="true"
      aria-label={`Фото · ${name}`}
    >
      <button
        aria-label="Закрити фото"
        className="person-photo-lightbox-backdrop"
        onClick={onClose}
        type="button"
      />
      <img alt={name} src={photo} />
      <button className="person-photo-lightbox-close" onClick={onClose} type="button">
        Закрити
      </button>
    </div>
  );
}
