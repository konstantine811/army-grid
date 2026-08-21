import { useLayoutEffect, useRef, useState, type ReactNode } from "react";

export const UBD_DOCUMENT_PAGE_WIDTH = 794;

export function UbdDocumentPreviewViewport({ children }: { children: ReactNode }) {
  const viewportRef = useRef<HTMLDivElement>(null);
  const [scale, setScale] = useState(1);

  useLayoutEffect(() => {
    const viewport = viewportRef.current;
    if (!viewport) return;

    const measure = () => {
      const availableWidth = Math.max(viewport.clientWidth - 24, 1);
      setScale(Math.min(1, availableWidth / UBD_DOCUMENT_PAGE_WIDTH));
    };

    measure();
    const observer = new ResizeObserver(measure);
    observer.observe(viewport);
    return () => observer.disconnect();
  }, []);

  return (
    <div className="ubd-preview-viewport sci-custom-scroll-target" ref={viewportRef}>
      <div className="ubd-preview-scale-outer">
        <div
          className="ubd-preview-scale-inner"
          style={{
            width: UBD_DOCUMENT_PAGE_WIDTH,
            zoom: scale,
          }}
        >
          {children}
        </div>
      </div>
    </div>
  );
}
