import {
  forwardRef,
  useEffect,
  useImperativeHandle,
  useRef,
  useState,
  type PointerEvent,
} from "react";
import {
  Button,
  Dialog,
  DialogActions,
  DialogContent,
  DialogTitle,
  LinearProgress,
  Stack,
  Typography,
} from "@/components/sci/SciPrimitives";
import { CloudUploadOutlinedIcon } from "@/components/sci/icons";
import * as pdfjs from "pdfjs-dist";
import pdfWorkerUrl from "pdfjs-dist/build/pdf.worker.min.mjs?url";
import { FloatingWindow } from "./FloatingWindow";
import {
  fitPhotoDimensions,
  PHOTO_JPEG_QUALITY,
} from "./photoCompression";

pdfjs.GlobalWorkerOptions.workerSrc = pdfWorkerUrl;

export type PhotoCropDialogProps = {
  file: File | null;
  open: boolean;
  onClose: () => void;
  onSave: (dataUrl: string, crop: CropRect) => void;
  onMessage: (message: string) => void;
  /** Draggable floating window (e.g. from disk-search flow). */
  floating?: boolean;
};

export type CropRect = {
  x: number;
  y: number;
  width: number;
  height: number;
};

export const defaultCropRect: CropRect = {
  x: 24,
  y: 24,
  width: 180,
  height: 220,
};

type PreviewPage = {
  id: string;
  src: string;
  pageNumber?: number;
};

const renderPdfPage = async (
  pdf: Awaited<ReturnType<typeof pdfjs.getDocument>["promise"]>,
  pageNumber: number,
  scale: number,
) => {
  const page = await pdf.getPage(pageNumber);
  const viewport = page.getViewport({ scale });
  const canvas = document.createElement("canvas");
  canvas.width = Math.max(1, Math.floor(viewport.width));
  canvas.height = Math.max(1, Math.floor(viewport.height));
  const context = canvas.getContext("2d");
  if (!context) throw new Error("Не вдалося створити canvas для PDF.");

  try {
    context.fillStyle = "#ffffff";
    context.fillRect(0, 0, canvas.width, canvas.height);
    await page.render({
      canvasContext: context,
      background: "#ffffff",
      viewport,
    }).promise;
    const src = canvas.toDataURL("image/jpeg", 0.86);
    if (!src || src === "data:,") {
      throw new Error("Не вдалося відрендерити сторінку PDF.");
    }
    return {
      id: `page-${pageNumber}`,
      pageNumber,
      src,
    } satisfies PreviewPage;
  } finally {
    page.cleanup();
    canvas.width = 1;
    canvas.height = 1;
  }
};

export const renderPdfPageToImageDataUrl = async (
  file: File,
  pageNumber: number,
  options: { scale?: number } = {},
) => {
  const data = new Uint8Array(await file.arrayBuffer());
  const loadingTask = pdfjs.getDocument({ data });
  const pdf = await loadingTask.promise;
  const scale = options.scale ?? 1.7;

  try {
    const safePageNumber = Math.min(
      pdf.numPages,
      Math.max(1, Math.floor(pageNumber)),
    );
    return {
      page: await renderPdfPage(pdf, safePageNumber, scale),
      pageCount: pdf.numPages,
    };
  } finally {
    pdf.cleanup();
    await pdf.destroy();
  }
};

export const visitPdfPagesAsImageDataUrls = async (
  file: File,
  visitor: (
    page: PreviewPage & { pageNumber: number },
    pageCount: number,
  ) => void | Promise<void>,
  options: { scale?: number } = {},
) => {
  const data = new Uint8Array(await file.arrayBuffer());
  const loadingTask = pdfjs.getDocument({ data });
  const pdf = await loadingTask.promise;
  const scale = options.scale ?? 1.7;

  try {
    for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
      const page = await renderPdfPage(pdf, pageNumber, scale);
      await visitor(
        { ...page, pageNumber },
        pdf.numPages,
      );
    }
  } finally {
    pdf.cleanup();
    await pdf.destroy();
  }
};

export type PhotoCropStageHandle = {
  save: () => void;
};

export const PhotoCropStage = forwardRef<
  PhotoCropStageHandle,
  {
    file: File | null;
    active: boolean;
    compact?: boolean;
    onSave: (dataUrl: string, crop: CropRect) => void;
    onMessage: (message: string) => void;
    onReadyChange?: (ready: boolean) => void;
  }
>(function PhotoCropStage(
  { file, active, compact = false, onSave, onMessage, onReadyChange },
  ref,
) {
  const previewRef = useRef<HTMLDivElement | null>(null);
  const imageRefs = useRef<Array<HTMLImageElement | null>>([]);
  const [previewPages, setPreviewPages] = useState<PreviewPage[]>([]);
  const [activePdfPage, setActivePdfPage] = useState(1);
  const [pdfPageCount, setPdfPageCount] = useState(1);
  const [isPreparingPreview, setIsPreparingPreview] = useState(false);
  const [cropRect, setCropRect] = useState<CropRect>(defaultCropRect);
  const [isDrawing, setIsDrawing] = useState(false);
  const [drawStart, setDrawStart] = useState({ x: 0, y: 0 });
  const isPdf = file?.type === "application/pdf";
  const ready = Boolean(previewPages.length) && !isPreparingPreview;

  useEffect(() => {
    onReadyChange?.(ready);
  }, [onReadyChange, ready]);

  useEffect(() => {
    setActivePdfPage(1);
    setPdfPageCount(1);
  }, [file]);

  useEffect(() => {
    if (!file || !active) {
      setPreviewPages([]);
      setIsPreparingPreview(false);
      return;
    }

    let cancelled = false;
    setIsPreparingPreview(true);
    setPreviewPages([]);
    setCropRect(defaultCropRect);
    imageRefs.current = [];

    const preparePreview = async () => {
      try {
        if (file.type.startsWith("image/")) {
          const reader = new FileReader();
          const dataUrl = await new Promise<string>((resolve, reject) => {
            reader.onload = () => resolve(String(reader.result ?? ""));
            reader.onerror = () => reject(new Error("Не вдалося прочитати файл."));
            reader.readAsDataURL(file);
          });
          if (!cancelled) setPreviewPages([{ id: "image", src: dataUrl }]);
          return;
        }

        if (file.type === "application/pdf") {
          const rendered = await renderPdfPageToImageDataUrl(
            file,
            activePdfPage,
          );
          if (!cancelled) {
            setPdfPageCount(rendered.pageCount);
            setPreviewPages([rendered.page]);
          }
          return;
        }

        throw new Error("Підтримуються лише зображення та PDF.");
      } catch (error) {
        if (!cancelled) {
          setPreviewPages([]);
          onMessage(
            error instanceof Error
              ? error.message
              : "Не вдалося підготувати файл для вирізання фото.",
          );
        }
      } finally {
        if (!cancelled) setIsPreparingPreview(false);
      }
    };

    void preparePreview();

    return () => {
      cancelled = true;
    };
  }, [file, onMessage, active, activePdfPage]);

  const getPointerPosition = (event: PointerEvent<HTMLDivElement>) => {
    const innerRect = event.currentTarget.getBoundingClientRect();

    return {
      x: Math.min(
        Math.max(event.clientX - innerRect.left, 0),
        event.currentTarget.scrollWidth || innerRect.width,
      ),
      y: Math.min(
        Math.max(event.clientY - innerRect.top, 0),
        event.currentTarget.scrollHeight || innerRect.height,
      ),
    };
  };

  const startCrop = (event: PointerEvent<HTMLDivElement>) => {
    if (!previewPages.length || isPreparingPreview) return;

    const position = getPointerPosition(event);
    setDrawStart(position);
    setCropRect({ ...position, width: 1, height: 1 });
    setIsDrawing(true);
    event.currentTarget.setPointerCapture(event.pointerId);
  };

  const updateCrop = (event: PointerEvent<HTMLDivElement>) => {
    if (!isDrawing) return;

    const position = getPointerPosition(event);
    setCropRect({
      x: Math.min(drawStart.x, position.x),
      y: Math.min(drawStart.y, position.y),
      width: Math.abs(position.x - drawStart.x),
      height: Math.abs(position.y - drawStart.y),
    });
  };

  const finishCrop = () => setIsDrawing(false);

  const saveCrop = () => {
    if (!file || !previewPages.length) return;

    const stage = previewRef.current;
    const inner = stage?.querySelector(
      ".photo-crop-stage-inner",
    ) as HTMLDivElement | null;
    if (!stage || !inner) {
      onMessage("Зачекайте, поки зображення повністю завантажиться.");
      return;
    }

    const innerRect = inner.getBoundingClientRect();
    const centerX = cropRect.x + cropRect.width / 2;
    const centerY = cropRect.y + cropRect.height / 2;
    const image =
      imageRefs.current.find((item) => {
        if (!item) return false;
        const rect = item.getBoundingClientRect();
        const left = rect.left - innerRect.left;
        const top = rect.top - innerRect.top;
        return (
          centerX >= left &&
          centerX <= left + rect.width &&
          centerY >= top &&
          centerY <= top + rect.height
        );
      }) ?? imageRefs.current.find((item) => item?.naturalWidth && item.naturalHeight);

    if (!image || !image.naturalWidth || !image.naturalHeight) {
      onMessage("Зачекайте, поки зображення повністю завантажиться.");
      return;
    }

    const imageRect = image.getBoundingClientRect();
    if (!imageRect.width || !imageRect.height) {
      onMessage("Не вдалося визначити розмір фото для збереження.");
      return;
    }

    const imageOffsetX = imageRect.left - innerRect.left;
    const imageOffsetY = imageRect.top - innerRect.top;
    const hasManualCrop = cropRect.width >= 8 && cropRect.height >= 8;
    const stageCrop = hasManualCrop
      ? cropRect
      : {
          x: imageOffsetX,
          y: imageOffsetY,
          width: image.clientWidth,
          height: image.clientHeight,
        };

    const scaleX = image.naturalWidth / image.clientWidth;
    const scaleY = image.naturalHeight / image.clientHeight;
    const sourceX = Math.max(0, (stageCrop.x - imageOffsetX) * scaleX);
    const sourceY = Math.max(0, (stageCrop.y - imageOffsetY) * scaleY);
    const sourceWidth = Math.min(
      image.naturalWidth - sourceX,
      Math.max(1, stageCrop.width * scaleX),
    );
    const sourceHeight = Math.min(
      image.naturalHeight - sourceY,
      Math.max(1, stageCrop.height * scaleY),
    );

    const canvas = document.createElement("canvas");
    const target = fitPhotoDimensions(sourceWidth, sourceHeight);
    canvas.width = target.width;
    canvas.height = target.height;
    const context = canvas.getContext("2d");
    if (!context) return;

    context.drawImage(
      image,
      sourceX,
      sourceY,
      sourceWidth,
      sourceHeight,
      0,
      0,
      canvas.width,
      canvas.height,
    );
    onSave(canvas.toDataURL("image/jpeg", PHOTO_JPEG_QUALITY), stageCrop);
  };

  useImperativeHandle(ref, () => ({ save: saveCrop }), [
    cropRect,
    file,
    previewPages.length,
  ]);

  return (
    <div className={`photo-crop-embed${compact ? " is-compact" : ""}`}>
      <Typography variant="body2" color="text.secondary" sx={{ mb: 1 }}>
        {isPdf
          ? "Протягніть прямокутник навколо фото на сторінці, потім «Зберегти фото»."
          : "Можна одразу натиснути «Зберегти фото» або протягнути прямокутник для обрізки."}
      </Typography>
      {isPreparingPreview && <LinearProgress color="primary" sx={{ mb: 1 }} />}
      {isPdf && pdfPageCount > 1 && (
        <Stack direction="row" spacing={1} alignItems="center" sx={{ mb: 1 }}>
          <Button
            disabled={activePdfPage <= 1 || isPreparingPreview}
            onClick={() => setActivePdfPage((page) => Math.max(1, page - 1))}
            size="small"
            variant="outlined"
          >
            ←
          </Button>
          <Typography variant="caption" color="text.secondary">
            Сторінка {activePdfPage} з {pdfPageCount}
          </Typography>
          <Button
            disabled={activePdfPage >= pdfPageCount || isPreparingPreview}
            onClick={() =>
              setActivePdfPage((page) => Math.min(pdfPageCount, page + 1))
            }
            size="small"
            variant="outlined"
          >
            →
          </Button>
        </Stack>
      )}
      <div className="photo-crop-stage" ref={previewRef}>
        <div
          className="photo-crop-stage-inner"
          onPointerDown={startCrop}
          onPointerMove={updateCrop}
          onPointerUp={finishCrop}
          onPointerCancel={finishCrop}
        >
          {previewPages.length ? (
            <div className="photo-crop-pages">
              {previewPages.map((page, index) => (
                <img
                  alt={
                    isPdf
                      ? `Сторінка PDF ${page.pageNumber ?? activePdfPage} для кадрування`
                      : "Фото для кадрування"
                  }
                  key={page.id}
                  ref={(element) => {
                    imageRefs.current[index] = element;
                  }}
                  src={page.src}
                  onError={() => {
                    setPreviewPages((current) =>
                      current.filter((item) => item.id !== page.id),
                    );
                    onMessage(
                      isPdf
                        ? "Не вдалося показати одну зі сторінок PDF для кадрування."
                        : "Не вдалося показати зображення для кадрування.",
                    );
                  }}
                />
              ))}
            </div>
          ) : (
            <div className="photo-crop-empty">
              <CloudUploadOutlinedIcon />
              <span>
                {isPreparingPreview
                  ? isPdf
                    ? "Рендер PDF…"
                    : "Читання файлу…"
                  : "Файл ще не прочитано"}
              </span>
            </div>
          )}
          <div
            className="photo-crop-rect"
            style={{
              left: cropRect.x,
              top: cropRect.y,
              width: cropRect.width,
              height: cropRect.height,
            }}
          />
        </div>
      </div>
    </div>
  );
});

export function PhotoCropDialog({
  file,
  open,
  onClose,
  onSave,
  onMessage,
  floating = false,
}: PhotoCropDialogProps) {
  const stageRef = useRef<PhotoCropStageHandle>(null);
  const [ready, setReady] = useState(false);

  const handleSave = (dataUrl: string, crop: CropRect) => {
    onSave(dataUrl, crop);
    onClose();
  };

  const cropActions = (
    <>
      <Button variant="outlined" onClick={onClose}>
        Скасувати
      </Button>
      <Button
        disabled={!ready}
        variant="contained"
        onClick={() => stageRef.current?.save()}
        sx={{ color: "#1a1a14" }}
      >
        Зберегти фото
      </Button>
    </>
  );

  const cropBody = (
    <PhotoCropStage
      ref={stageRef}
      file={file}
      active={open}
      onSave={handleSave}
      onMessage={onMessage}
      onReadyChange={setReady}
    />
  );

  if (floating) {
    return (
      <FloatingWindow
        open={open}
        title="Обрати область фото"
        onClose={onClose}
        placement="center"
        defaultWidth={760}
        defaultHeight={780}
        minWidth={480}
        minHeight={420}
        className="photo-crop-floating"
        footer={cropActions}
      >
        {cropBody}
      </FloatingWindow>
    );
  }

  return (
    <Dialog
      fullWidth
      maxWidth="md"
      open={open}
      onClose={onClose}
      slotProps={{ paper: { className: "photo-crop-dialog" } }}
    >
      <DialogTitle>Обрати область фото</DialogTitle>
      <DialogContent>{cropBody}</DialogContent>
      <DialogActions>{cropActions}</DialogActions>
    </Dialog>
  );
}
