export type MagneticTapePreloaderProps = {
  brand?: string;
  caution?: string;
  status?: string;
  hint?: string;
};

export function MagneticTapePreloader({
  brand = "GRID.",
  caution = "УВАГА",
  status = "ЗАВАНТАЖЕННЯ ДАНИХ",
  hint,
}: MagneticTapePreloaderProps) {
  return (
    <div className="magnetic-tape-preloader" role="status" aria-live="polite">
      <div className="magnetic-tape-preloader__scanlines" aria-hidden="true" />
      <div className="magnetic-tape-preloader__widget">
        <div className="magnetic-tape-preloader__header">
          <span className="magnetic-tape-preloader__brand">{brand}</span>
          <div className="magnetic-tape-preloader__meta">
            <span>{caution}</span>
            <span>{status}</span>
          </div>
        </div>
        <div className="magnetic-tape-preloader__bars" aria-hidden="true">
          {Array.from({ length: 7 }, (_, index) => (
            <span
              key={index}
              className="magnetic-tape-preloader__bar"
              style={{ animationDelay: `${index * 0.11}s` }}
            />
          ))}
        </div>
        <div className="magnetic-tape-preloader__footer">{status}</div>
      </div>
      {hint ? <p className="magnetic-tape-preloader__hint">{hint}</p> : null}
    </div>
  );
}
