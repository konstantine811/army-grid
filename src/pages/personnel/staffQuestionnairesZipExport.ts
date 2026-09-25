import {
  createStoredZipBlob,
  downloadBlob,
  sanitizeFileName,
} from "../../shared/browserExport";
import { buildQuestionnaireExportFileName } from "./personnelUtils";

export type StaffQuestionnaireExportPerson = {
  externalId: string;
  name: string;
  callSign?: string;
};

export const STAFF_QUESTIONNAIRE_ZIP_CONCURRENCY = 2;
/** Скільки PDF в одному ZIP — один великий архів (~600) падає по пам’яті. */
export const STAFF_QUESTIONNAIRE_ZIP_PART_SIZE = 75;
const BETWEEN_DOWNLOADS_MS = 1200;

const yieldToBrowser = () =>
  new Promise<void>((resolve) => window.setTimeout(resolve, 0));

const pause = (ms: number) =>
  new Promise<void>((resolve) => window.setTimeout(resolve, ms));

export const makeUniqueZipEntryName = (
  fileName: string,
  usedNames: Map<string, number>,
) => {
  const normalized = fileName.toLocaleLowerCase("uk-UA");
  const count = (usedNames.get(normalized) ?? 0) + 1;
  usedNames.set(normalized, count);
  if (count === 1) return fileName;

  const lower = fileName.toLocaleLowerCase("uk-UA");
  if (lower.endsWith(".pdf")) {
    return `${fileName.slice(0, -4)} (${count}).pdf`;
  }
  return `${fileName} (${count})`;
};

export const splitStaffQuestionnairePeople = <T>(
  people: T[],
  partSize = STAFF_QUESTIONNAIRE_ZIP_PART_SIZE,
) => {
  const parts: T[][] = [];
  for (let offset = 0; offset < people.length; offset += partSize) {
    parts.push(people.slice(offset, offset + partSize));
  }
  return parts;
};

const buildQuestionnaireFileName = (person: StaffQuestionnaireExportPerson) => {
  const baseName = sanitizeFileName(
    buildQuestionnaireExportFileName(person.name, person.callSign),
  );
  return baseName.toLocaleLowerCase("uk-UA").endsWith(".pdf")
    ? baseName
    : `${baseName}.pdf`;
};

export const collectStaffQuestionnaireZipFiles = async (options: {
  people: StaffQuestionnaireExportPerson[];
  fetchQuestionnaireBlob: (externalId: string) => Promise<Blob | null>;
  onProgress?: (done: number, total: number) => void;
  isCancelled?: () => boolean;
  concurrency?: number;
  usedNames?: Map<string, number>;
}) => {
  const people = options.people.filter((person) => person.externalId.trim());
  const total = people.length;
  const concurrency = Math.max(
    1,
    options.concurrency ?? STAFF_QUESTIONNAIRE_ZIP_CONCURRENCY,
  );
  const usedNames = options.usedNames ?? new Map<string, number>();
  const files: Array<{ name: string; data: Uint8Array }> = [];
  let done = 0;
  let skipped = 0;

  for (let offset = 0; offset < people.length; offset += concurrency) {
    if (options.isCancelled?.()) break;
    const batch = people.slice(offset, offset + concurrency);
    const batchFiles = await Promise.all(
      batch.map(async (person) => {
        try {
          const blob = await options.fetchQuestionnaireBlob(person.externalId);
          if (!blob || blob.size <= 0) return null;
          const bytes = new Uint8Array(await blob.arrayBuffer());
          return {
            name: makeUniqueZipEntryName(
              buildQuestionnaireFileName(person),
              usedNames,
            ),
            data: bytes,
          };
        } catch {
          return null;
        }
      }),
    );

    for (const file of batchFiles) {
      if (file) files.push(file);
      else skipped += 1;
      done += 1;
    }
    options.onProgress?.(done, total);
    await yieldToBrowser();
  }

  return {
    files,
    exported: files.length,
    skipped,
    total,
  };
};

const triggerDownload = async (blob: Blob, fileName: string) => {
  downloadBlob(blob, fileName);
  await pause(BETWEEN_DOWNLOADS_MS);
};

export const downloadStaffQuestionnairesZip = async (options: {
  people: StaffQuestionnaireExportPerson[];
  fetchQuestionnaireBlob: (externalId: string) => Promise<Blob | null>;
  zipFileNamePrefix: string;
  onProgress?: (done: number, total: number) => void;
  onPhase?: (message: string) => void;
  isCancelled?: () => boolean;
  partSize?: number;
}) => {
  const people = options.people.filter((person) => person.externalId.trim());
  const parts = splitStaffQuestionnairePeople(
    people,
    options.partSize ?? STAFF_QUESTIONNAIRE_ZIP_PART_SIZE,
  );
  const usedNames = new Map<string, number>();
  let exported = 0;
  let skipped = 0;
  let partsDownloaded = 0;
  let processedOverall = 0;

  for (let partIndex = 0; partIndex < parts.length; partIndex += 1) {
    if (options.isCancelled?.()) break;
    const part = parts[partIndex]!;
    const partLabel = `${partIndex + 1}/${parts.length}`;
    options.onPhase?.(
      parts.length === 1
        ? `Завантажую PDF: 0/${people.length}…`
        : `Частина ${partLabel}: завантажую PDF…`,
    );

    const collected = await collectStaffQuestionnaireZipFiles({
      people: part,
      fetchQuestionnaireBlob: options.fetchQuestionnaireBlob,
      usedNames,
      isCancelled: options.isCancelled,
      onProgress: (done) => {
        options.onProgress?.(processedOverall + done, people.length);
        if (parts.length > 1) {
          options.onPhase?.(
            `Частина ${partLabel}: PDF ${done}/${part.length} · загалом ${processedOverall + done}/${people.length}`,
          );
        } else {
          options.onPhase?.(
            `Завантажую PDF: ${processedOverall + done}/${people.length}…`,
          );
        }
      },
    });

    processedOverall += part.length;
    exported += collected.exported;
    skipped += collected.skipped;

    if (!collected.files.length) continue;

    options.onPhase?.(
      parts.length === 1
        ? "Створюю ZIP…"
        : `Частина ${partLabel}: створюю ZIP (${collected.files.length} PDF)…`,
    );

    let blob: Blob;
    try {
      blob = createStoredZipBlob(collected.files);
    } catch (error) {
      throw new Error(
        error instanceof Error
          ? `Не вдалося зібрати ZIP (частина ${partLabel}): ${error.message}`
          : `Не вдалося зібрати ZIP (частина ${partLabel}). Спробуйте «У папку».`,
      );
    }

    const zipName =
      parts.length === 1
        ? `${options.zipFileNamePrefix}.zip`
        : `${options.zipFileNamePrefix} · частина ${partLabel}.zip`;

    options.onPhase?.(
      parts.length === 1
        ? "Завантаження ZIP…"
        : `Завантаження ${partLabel} з ${parts.length}…`,
    );
    await triggerDownload(blob, zipName);
    partsDownloaded += 1;
    await yieldToBrowser();
  }

  return {
    exported,
    skipped,
    total: people.length,
    partsDownloaded,
    downloaded: partsDownloaded > 0,
  };
};

type DirectoryPickerWindow = Window & {
  showDirectoryPicker?: (options?: {
    mode?: "read" | "readwrite";
  }) => Promise<FileSystemDirectoryHandle>;
};

export const canSaveStaffQuestionnairesToDirectory = () =>
  typeof (window as DirectoryPickerWindow).showDirectoryPicker === "function";

export const saveStaffQuestionnairesToDirectory = async (options: {
  people: StaffQuestionnaireExportPerson[];
  fetchQuestionnaireBlob: (externalId: string) => Promise<Blob | null>;
  onProgress?: (done: number, total: number) => void;
  onPhase?: (message: string) => void;
  isCancelled?: () => boolean;
}) => {
  const picker = (window as DirectoryPickerWindow).showDirectoryPicker;
  if (!picker) {
    throw new Error(
      "Браузер не підтримує збереження в папку. Використайте ZIP частинами.",
    );
  }

  const people = options.people.filter((person) => person.externalId.trim());
  const dirHandle = await picker({ mode: "readwrite" });
  const usedNames = new Map<string, number>();
  let exported = 0;
  let skipped = 0;

  for (let index = 0; index < people.length; index += 1) {
    if (options.isCancelled?.()) break;
    const person = people[index]!;
    options.onPhase?.(`У папку: ${index + 1}/${people.length} · ${person.name}`);
    options.onProgress?.(index + 1, people.length);

    try {
      const blob = await options.fetchQuestionnaireBlob(person.externalId);
      if (!blob || blob.size <= 0) {
        skipped += 1;
        continue;
      }
      const fileName = makeUniqueZipEntryName(
        buildQuestionnaireFileName(person),
        usedNames,
      );
      const fileHandle = await dirHandle.getFileHandle(fileName, {
        create: true,
      });
      const writable = await fileHandle.createWritable();
      await writable.write(blob);
      await writable.close();
      exported += 1;
    } catch {
      skipped += 1;
    }

    if (index % 4 === 3) await yieldToBrowser();
  }

  return {
    exported,
    skipped,
    total: people.length,
    downloaded: exported > 0,
    partsDownloaded: exported > 0 ? 1 : 0,
  };
};
