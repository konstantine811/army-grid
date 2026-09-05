export const PHOTO_MAX_WIDTH = 480;
export const PHOTO_MAX_HEIGHT = 640;
export const PHOTO_JPEG_QUALITY = 0.78;
export const PHOTO_DATA_URL_TARGET_BYTES = 220 * 1024;

export const estimateDataUrlBytes = (dataUrl: string) => {
  const encoded = String(dataUrl ?? "").split(",", 2)[1] ?? "";
  return Math.ceil((encoded.length * 3) / 4);
};

export const fitPhotoDimensions = (
  width: number,
  height: number,
  maxWidth = PHOTO_MAX_WIDTH,
  maxHeight = PHOTO_MAX_HEIGHT,
) => {
  const safeWidth = Math.max(1, Math.round(width));
  const safeHeight = Math.max(1, Math.round(height));
  const scale = Math.min(1, maxWidth / safeWidth, maxHeight / safeHeight);
  return {
    width: Math.max(1, Math.round(safeWidth * scale)),
    height: Math.max(1, Math.round(safeHeight * scale)),
  };
};

const loadImage = (source: string) =>
  new Promise<HTMLImageElement>((resolve, reject) => {
    const image = new Image();
    image.onload = () => resolve(image);
    image.onerror = () => reject(new Error("Не вдалося стиснути фото."));
    image.src = source;
  });

const encodeImage = (
  image: CanvasImageSource,
  sourceWidth: number,
  sourceHeight: number,
  options?: { maxWidth?: number; maxHeight?: number; quality?: number },
) => {
  const target = fitPhotoDimensions(
    sourceWidth,
    sourceHeight,
    options?.maxWidth,
    options?.maxHeight,
  );
  const canvas = document.createElement("canvas");
  canvas.width = target.width;
  canvas.height = target.height;
  const context = canvas.getContext("2d");
  if (!context) throw new Error("Не вдалося підготувати фото.");
  context.fillStyle = "#ffffff";
  context.fillRect(0, 0, target.width, target.height);
  context.drawImage(image, 0, 0, target.width, target.height);
  return canvas.toDataURL(
    "image/jpeg",
    options?.quality ?? PHOTO_JPEG_QUALITY,
  );
};

export const compressPhotoDataUrl = async (dataUrl: string) => {
  const image = await loadImage(dataUrl);
  if (
    estimateDataUrlBytes(dataUrl) <= PHOTO_DATA_URL_TARGET_BYTES &&
    image.naturalWidth <= PHOTO_MAX_WIDTH &&
    image.naturalHeight <= PHOTO_MAX_HEIGHT &&
    /^data:image\/jpeg/i.test(dataUrl)
  ) {
    return dataUrl;
  }
  return encodeImage(image, image.naturalWidth, image.naturalHeight);
};

export const compressPhotoFile = async (file: File) => {
  if (!file.type.startsWith("image/")) {
    throw new Error("Для фото потрібно вибрати файл зображення.");
  }
  const objectUrl = URL.createObjectURL(file);
  try {
    const image = await loadImage(objectUrl);
    return encodeImage(image, image.naturalWidth, image.naturalHeight);
  } finally {
    URL.revokeObjectURL(objectUrl);
  }
};

export const createPhotoThumbnailDataUrl = async (dataUrl: string) => {
  const image = await loadImage(dataUrl);
  return encodeImage(image, image.naturalWidth, image.naturalHeight, {
    maxWidth: 96,
    maxHeight: 128,
    quality: 0.68,
  });
};
