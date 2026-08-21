type SignatureProcessOptions = {
  whiteThreshold?: number;
  trim?: boolean;
};

const DEFAULT_WHITE_THRESHOLD = 236;

const trimTransparentBounds = (
  ctx: CanvasRenderingContext2D,
  width: number,
  height: number,
) => {
  const { data } = ctx.getImageData(0, 0, width, height);
  let top = height;
  let left = width;
  let right = 0;
  let bottom = 0;

  for (let y = 0; y < height; y += 1) {
    for (let x = 0; x < width; x += 1) {
      const alpha = data[(y * width + x) * 4 + 3];
      if (alpha > 8) {
        top = Math.min(top, y);
        left = Math.min(left, x);
        right = Math.max(right, x);
        bottom = Math.max(bottom, y);
      }
    }
  }

  if (right <= left || bottom <= top) {
    return null;
  }

  const pad = 2;
  const cropX = Math.max(0, left - pad);
  const cropY = Math.max(0, top - pad);
  const cropW = Math.min(width - cropX, right - left + 1 + pad * 2);
  const cropH = Math.min(height - cropY, bottom - top + 1 + pad * 2);

  return ctx.getImageData(cropX, cropY, cropW, cropH);
};

export const processSignatureTransparentBackground = (
  dataUrl: string,
  options: SignatureProcessOptions = {},
): Promise<string> =>
  new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => {
      const canvas = document.createElement("canvas");
      canvas.width = img.naturalWidth;
      canvas.height = img.naturalHeight;
      const ctx = canvas.getContext("2d", { willReadFrequently: true });
      if (!ctx) {
        reject(new Error("Canvas unavailable"));
        return;
      }

      ctx.drawImage(img, 0, 0);
      const threshold = options.whiteThreshold ?? DEFAULT_WHITE_THRESHOLD;
      const imageData = ctx.getImageData(0, 0, canvas.width, canvas.height);
      const { data } = imageData;

      for (let i = 0; i < data.length; i += 4) {
        const r = data[i];
        const g = data[i + 1];
        const b = data[i + 2];
        const max = Math.max(r, g, b);
        const min = Math.min(r, g, b);
        const spread = max - min;

        if (max >= threshold && spread < 28) {
          data[i + 3] = 0;
          continue;
        }

        if (max >= threshold - 36) {
          const fade = (max - (threshold - 36)) / 36;
          data[i + 3] = Math.round(data[i + 3] * (1 - fade));
        }
      }

      ctx.putImageData(imageData, 0, 0);

      if (options.trim !== false) {
        const trimmed = trimTransparentBounds(ctx, canvas.width, canvas.height);
        if (trimmed) {
          canvas.width = trimmed.width;
          canvas.height = trimmed.height;
          ctx.clearRect(0, 0, canvas.width, canvas.height);
          ctx.putImageData(trimmed, 0, 0);
        }
      }

      resolve(canvas.toDataURL("image/png"));
    };
    img.onerror = () => reject(new Error("Signature image failed to load"));
    img.src = dataUrl;
  });

let cachedTransparentSignature: string | null = null;
let pendingTransparentSignature: Promise<string> | null = null;

export const getCommanderSignatureTransparent = (jpegDataUrl: string) => {
  if (cachedTransparentSignature) {
    return Promise.resolve(cachedTransparentSignature);
  }
  if (!pendingTransparentSignature) {
    pendingTransparentSignature = processSignatureTransparentBackground(
      jpegDataUrl,
    ).then((processed) => {
      cachedTransparentSignature = processed;
      return processed;
    });
  }
  return pendingTransparentSignature;
};

export const getCachedCommanderSignatureTransparent = () =>
  cachedTransparentSignature;

export const copyPngDataUrlToClipboard = async (dataUrl: string) => {
  const response = await fetch(dataUrl);
  const blob = await response.blob();
  const pngBlob =
    blob.type === "image/png"
      ? blob
      : new Blob([blob], { type: "image/png" });
  await navigator.clipboard.write([
    new ClipboardItem({ "image/png": pngBlob }),
  ]);
};
