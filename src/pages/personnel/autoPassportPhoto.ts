import {
  FaceDetector as MediaPipeFaceDetector,
  FilesetResolver,
  type BoundingBox,
} from "@mediapipe/tasks-vision";
import {
  visitPdfPagesAsImageDataUrls,
  type CropRect,
} from "./PhotoCropDialog";
import { PHOTO_JPEG_QUALITY } from "./photoCompression";

type DetectedFace = {
  boundingBox: FaceBox;
};

type FaceCandidate = {
  box: FaceBox;
  confidence: number;
};

type FaceBox = {
  x: number;
  y: number;
  width: number;
  height: number;
};

type FaceDetectorCtor = new (options?: {
  fastMode?: boolean;
  maxDetectedFaces?: number;
}) => {
  detect: (source: ImageBitmap | HTMLImageElement) => Promise<DetectedFace[]>;
};

type PassportPhotoResult = {
  dataUrl: string;
  crop: CropRect & { pageIndex: number; auto: true };
};

const PASSPORT_RATIO = 3 / 4;
const OUTPUT_WIDTH = 360;
const OUTPUT_HEIGHT = 480;
const AUTO_RENDER_SCALE = 2.35;
const PUBLIC_BASE = import.meta.env.BASE_URL || "/";
const MEDIAPIPE_WASM_PATH = `${PUBLIC_BASE}mediapipe/wasm`;
const FACE_MODEL_PATH = `${PUBLIC_BASE}models/blaze_face_short_range.tflite`;

let mediaPipeDetectorPromise: Promise<MediaPipeFaceDetector> | null = null;

const loadImage = (src: string) =>
  new Promise<HTMLImageElement>((resolve, reject) => {
    const image = new Image();
    image.onload = () => resolve(image);
    image.onerror = () => reject(new Error("Не вдалося прочитати сторінку PDF як зображення."));
    image.src = src;
  });

const clampCrop = (
  crop: CropRect,
  bounds: { width: number; height: number },
): CropRect => {
  let { x, y, width, height } = crop;
  width = Math.min(width, bounds.width);
  height = Math.min(height, bounds.height);
  x = Math.min(Math.max(0, x), Math.max(0, bounds.width - width));
  y = Math.min(Math.max(0, y), Math.max(0, bounds.height - height));
  return { x, y, width, height };
};

const cropAroundFace = (
  face: FaceBox,
  bounds: { width: number; height: number },
): CropRect => {
  const faceCenterX = face.x + face.width / 2;
  const faceCenterY = face.y + face.height / 2;
  const minWidthFromFace = Math.max(face.width * 1.85, face.height * 1.12);
  const minHeightFromFace = minWidthFromFace / PASSPORT_RATIO;
  const height = Math.min(bounds.height, Math.max(minHeightFromFace, face.height * 2.18));
  const width = Math.min(bounds.width, height * PASSPORT_RATIO);

  return clampCrop(
    {
      x: faceCenterX - width / 2,
      y: faceCenterY - height * 0.45,
      width,
      height,
    },
    bounds,
  );
};

const drawPassportCrop = (image: HTMLImageElement, crop: CropRect) => {
  const canvas = document.createElement("canvas");
  canvas.width = OUTPUT_WIDTH;
  canvas.height = OUTPUT_HEIGHT;
  const context = canvas.getContext("2d");
  if (!context) throw new Error("Не вдалося підготувати canvas для автофото.");

  context.fillStyle = "#ffffff";
  context.fillRect(0, 0, canvas.width, canvas.height);
  context.drawImage(
    image,
    crop.x,
    crop.y,
    crop.width,
    crop.height,
    0,
    0,
    canvas.width,
    canvas.height,
  );

  return canvas.toDataURL("image/jpeg", PHOTO_JPEG_QUALITY);
};

const toFaceBox = (box: BoundingBox): FaceBox => ({
  x: box.originX,
  y: box.originY,
  width: box.width,
  height: box.height,
});

const getMediaPipeDetector = () => {
  mediaPipeDetectorPromise ??= FilesetResolver.forVisionTasks(
    MEDIAPIPE_WASM_PATH,
  ).then((vision) =>
    MediaPipeFaceDetector.createFromOptions(vision, {
      baseOptions: {
        modelAssetPath: FACE_MODEL_PATH,
        delegate: "CPU",
      },
      minDetectionConfidence: 0.35,
      minSuppressionThreshold: 0.3,
      runningMode: "IMAGE",
    }),
  );
  return mediaPipeDetectorPromise;
};

const detectFaceWithMediaPipe = async (image: HTMLImageElement) => {
  const detector = await getMediaPipeDetector();
  const result = detector.detect(image);
  return result.detections
    .map((detection): FaceCandidate | null => {
      if (!detection.boundingBox) return null;
      const confidence = detection.categories[0]?.score ?? 0;
      return {
        box: toFaceBox(detection.boundingBox),
        confidence,
      };
    })
    .filter((face): face is FaceCandidate => Boolean(face));
};

const detectFaceWithNativeApi = async (image: HTMLImageElement) => {
  const FaceDetector = (window as typeof window & {
    FaceDetector?: FaceDetectorCtor;
  }).FaceDetector;
  if (!FaceDetector) return null;

  const detector = new FaceDetector({ fastMode: true, maxDetectedFaces: 4 });
  const bitmap = await createImageBitmap(image);
  try {
    const faces = await detector.detect(bitmap);
    return faces.map((face) => ({
      box: face.boundingBox,
      confidence: 0.55,
    }));
  } finally {
    bitmap.close();
  }
};

const detectFaces = async (image: HTMLImageElement) => {
  try {
    const mediaPipeFaces = await detectFaceWithMediaPipe(image);
    if (mediaPipeFaces.length) return mediaPipeFaces;
  } catch {
    // Fall through to the browser API when MediaPipe cannot initialize.
  }

  return detectFaceWithNativeApi(image);
};

const scoreFaceCandidate = (
  candidate: FaceCandidate,
  bounds: { width: number; height: number },
) => {
  const { box, confidence } = candidate;
  const pageArea = bounds.width * bounds.height;
  const faceArea = box.width * box.height;
  const faceAreaRatio = faceArea / pageArea;
  const faceWidthRatio = box.width / bounds.width;
  const faceHeightRatio = box.height / bounds.height;
  const faceAspect = box.width / Math.max(1, box.height);

  if (confidence < 0.42) return null;
  if (faceAspect < 0.48 || faceAspect > 1.55) return null;
  if (faceAreaRatio < 0.00006) return null;
  if (faceAreaRatio > 0.18 || faceWidthRatio > 0.68 || faceHeightRatio > 0.68) {
    return null;
  }

  const portraitCrop = cropAroundFace(box, bounds);
  const cropAreaRatio = (portraitCrop.width * portraitCrop.height) / pageArea;
  if (cropAreaRatio > 0.72) return null;

  const smallIdPhotoBoost = faceAreaRatio < 0.006 ? 1.18 : 1;
  const generalPhotoBoost = faceAreaRatio >= 0.006 && faceAreaRatio <= 0.09 ? 1.12 : 1;
  const cropPenalty = cropAreaRatio > 0.36 ? 0.78 : cropAreaRatio > 0.22 ? 0.9 : 1;
  const sizeScore = Math.log10(Math.max(10, faceArea)) * 850;

  return (
    (confidence * 12000 + sizeScore) *
    smallIdPhotoBoost *
    generalPhotoBoost *
    cropPenalty
  );
};

export async function extractPassportPhotoFromPdf(
  file: File,
): Promise<PassportPhotoResult | null> {
  let best:
    | {
        dataUrl: string;
        crop: CropRect;
        pageIndex: number;
        score: number;
      }
    | null = null;

  await visitPdfPagesAsImageDataUrls(
    file,
    async (page) => {
      const image = await loadImage(page.src);
      try {
        const faces = await detectFaces(image);
        const bounds = {
          width: image.naturalWidth,
          height: image.naturalHeight,
        };
        for (const candidate of faces ?? []) {
          const score = scoreFaceCandidate(candidate, bounds);
          if (score === null || (best && score <= best.score)) continue;
          const crop = cropAroundFace(candidate.box, bounds);
          best = {
            dataUrl: drawPassportCrop(image, crop),
            crop,
            pageIndex: page.pageNumber - 1,
            score,
          };
        }
      } finally {
        image.src = "";
      }
    },
    { scale: AUTO_RENDER_SCALE },
  );

  const resolvedBest = best as {
    dataUrl: string;
    crop: CropRect;
    pageIndex: number;
    score: number;
  } | null;
  if (!resolvedBest) return null;

  return {
    dataUrl: resolvedBest.dataUrl,
    crop: {
      ...resolvedBest.crop,
      pageIndex: resolvedBest.pageIndex,
      auto: true,
    },
  };
}
