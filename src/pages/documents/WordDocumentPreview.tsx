import {
  useCallback,
  useEffect,
  useLayoutEffect,
  useRef,
  useState,
} from "react";
import { renderAsync } from "docx-preview";
import { Spinner } from "@/components/ui/spinner/spinner";

const PAGE_WIDTH = 794;

const PREVIEW_OPTIONS = {
  className: "docx",
  inWrapper: true,
  hideWrapperOnPrint: false,
  ignoreWidth: false,
  ignoreHeight: false,
  ignoreFonts: false,
  breakPages: true,
  ignoreLastRenderedPageBreak: false,
  experimental: true,
  renderHeaders: true,
  renderFooters: true,
  renderFootnotes: true,
  renderEndnotes: true,
  useBase64URL: true,
} as const;

export function useWordPreviewBlob(build: () => Promise<Blob>, enabled: boolean) {
  const [blob, setBlob] = useState<Blob | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [isLoading, setIsLoading] = useState(false);

  useEffect(() => {
    if (!enabled) return;
    let cancelled = false;
    setIsLoading(true);
    const timer = window.setTimeout(() => {
      void build()
        .then((next) => {
          if (cancelled) return;
          setBlob(next);
          setError(null);
        })
        .catch((cause) => {
          if (cancelled) return;
          setError(
            cause instanceof Error
              ? cause.message
              : "Не вдалося зібрати Word-документ.",
          );
        })
        .finally(() => {
          if (!cancelled) setIsLoading(false);
        });
    }, 400);
    return () => {
      cancelled = true;
      window.clearTimeout(timer);
    };
  }, [build, enabled]);

  return { blob, error, isLoading };
}

export function printRenderedWordPreview() {
  const preview = document.querySelector(".word-document-preview");
  if (!(preview instanceof HTMLElement)) return;
  const printWindow = window.open("", "_blank", "noopener,noreferrer");
  if (!printWindow) return;
  const styles = [...preview.querySelectorAll("style")]
    .map((node) => node.outerHTML)
    .join("");
  const host = preview.querySelector(".word-document-preview-host");
  printWindow.document.write(`<!doctype html>
<html lang="uk">
<head>
  <meta charset="utf-8" />
  <title>Рапорт</title>
  ${styles}
  <style>
    html, body { margin: 0; background: #fff; }
    .docx-wrapper { background: #fff !important; padding: 0 !important; }
  </style>
</head>
<body>${host?.innerHTML ?? preview.innerHTML}</body>
</html>`);
  printWindow.document.close();
  printWindow.focus();
  printWindow.print();
}

export function WordDocumentPreview({
  blob,
  error,
  isLoading,
}: {
  blob: Blob | null;
  error?: string | null;
  isLoading?: boolean;
}) {
  const viewportRef = useRef<HTMLDivElement>(null);
  const hostRef = useRef<HTMLDivElement>(null);
  const styleRef = useRef<HTMLDivElement>(null);
  const [scale, setScale] = useState(1);
  const [pageWidth, setPageWidth] = useState(PAGE_WIDTH);
  const [isRendering, setIsRendering] = useState(false);

  useEffect(() => {
    const host = hostRef.current;
    const style = styleRef.current;
    if (!host || !blob) return;

    let cancelled = false;
    setIsRendering(true);

    const target = document.createElement("div");
    const styleTarget = document.createElement("div");
    void renderAsync(blob, target, styleTarget, PREVIEW_OPTIONS)
      .then(() => {
        if (cancelled) return;
        host.replaceChildren(...Array.from(target.childNodes));
        style?.replaceChildren(...Array.from(styleTarget.childNodes));
        const page = host.querySelector("section.docx");
        if (page instanceof HTMLElement && page.offsetWidth > 0) {
          setPageWidth(page.offsetWidth);
        }
      })
      .catch((cause) => {
        if (cancelled) return;
        host.textContent =
          cause instanceof Error
            ? cause.message
            : "Не вдалося показати Word-документ.";
      })
      .finally(() => {
        if (!cancelled) setIsRendering(false);
      });

    return () => {
      cancelled = true;
    };
  }, [blob]);

  const measure = useCallback(() => {
    const viewport = viewportRef.current;
    if (!viewport) return;
    const available = Math.max(viewport.clientWidth - 24, 1);
    setScale(Math.min(1, available / pageWidth));
  }, [pageWidth]);

  useLayoutEffect(() => {
    const viewport = viewportRef.current;
    if (!viewport) return;
    measure();
    const observer = new ResizeObserver(measure);
    observer.observe(viewport);
    return () => observer.disconnect();
  }, [measure, blob, isRendering]);

  return (
    <div
      className="word-document-preview ubd-preview-viewport sci-custom-scroll-target"
      ref={viewportRef}
    >
      <div ref={styleRef} className="word-document-preview-styles" hidden />
      <div className="ubd-preview-scale-outer">
        <div
          className="ubd-preview-scale-inner word-document-preview-host"
          ref={hostRef}
          style={{
            width: pageWidth,
            zoom: scale,
          }}
        />
      </div>
      {error && !blob ? (
        <p className="word-document-preview-status">{error}</p>
      ) : !blob && (isLoading || isRendering) ? (
        <div className="word-document-preview-status word-document-preview-status-overlay">
          <Spinner />
          <span>Збираю Word-документ…</span>
        </div>
      ) : null}
      {blob && (isLoading || isRendering) ? (
        <div className="word-document-preview-busy">
          <Spinner />
        </div>
      ) : null}
      {error && blob ? (
        <p className="word-document-preview-error">{error}</p>
      ) : null}
    </div>
  );
}
