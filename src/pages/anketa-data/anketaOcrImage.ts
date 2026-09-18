const OCR_IMAGE_MAX_CHARS = 2_400_000;
const OCR_IMAGE_MAX_SIDE = 2480;

const encodeOcrJpeg = (
  image: HTMLImageElement,
  maxChars: number,
): string => {
  const longest = Math.max(image.naturalWidth, image.naturalHeight, 1);
  const scale = Math.min(1, OCR_IMAGE_MAX_SIDE / longest);
  const width = Math.max(1, Math.floor(image.naturalWidth * scale));
  const height = Math.max(1, Math.floor(image.naturalHeight * scale));
  const canvas = document.createElement("canvas");
  canvas.width = width;
  canvas.height = height;
  const context = canvas.getContext("2d");
  if (!context) return "";
  context.fillStyle = "#ffffff";
  context.fillRect(0, 0, width, height);
  context.drawImage(image, 0, 0, width, height);
  let quality = 0.88;
  let next = canvas.toDataURL("image/jpeg", quality);
  while (next.length > maxChars && quality > 0.55) {
    quality -= 0.06;
    next = canvas.toDataURL("image/jpeg", quality);
  }
  return next;
};

/** Стискає лише занадто великі кадри; дрібний JPEG зі сторінки PDF не чіпаємо. */
export const prepareOcrImageDataUrl = async (
  dataUrl: string,
): Promise<string> => {
  if (
    dataUrl.startsWith("data:image/jpeg") &&
    dataUrl.length <= OCR_IMAGE_MAX_CHARS
  ) {
    return dataUrl;
  }

  return new Promise((resolve, reject) => {
    const image = new Image();
    image.onload = () => {
      const encoded = encodeOcrJpeg(image, OCR_IMAGE_MAX_CHARS);
      resolve(encoded || dataUrl);
    };
    image.onerror = () =>
      reject(new Error("Не вдалося підготувати зображення для OCR."));
    image.src = dataUrl;
  });
};
