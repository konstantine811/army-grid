import JSZip from "jszip";
import type { UbdBasisOrderOption } from "./ubdBasisOrdersData";

const BASIS_LINE_PATTERN =
  /(?:№\s*)?(\d{4}\/[^/\s]+\/\d+\/?\/?дск)\s*від\s+(\d{1,2}[.\/-]\d{1,2}[.\/-]\d{2,4})/giu;

export const normalizeImportedBasisNumber = (value: string) =>
  String(value ?? "")
    .trim()
    .replace(/^№\s*/, "")
    .replace(/(\d)(дск)$/iu, "$1/дск")
    .replace(/\/{2,}/g, "/");

export const normalizeImportedBasisDate = (value: string) =>
  String(value ?? "")
    .trim()
    .replaceAll("/", ".")
    .replaceAll("-", ".");

export const parseBasisOrderLine = (
  line: string,
): UbdBasisOrderOption | null => {
  const text = String(line ?? "").trim();
  if (!text) return null;
  const match = text.match(
    /(?:№\s*)?(\d{4}\/[^/\s]+\/\d+\/?\/?дск)\s*від\s+(\d{1,2}[.\/-]\d{1,2}[.\/-]\d{2,4})/iu,
  );
  if (!match) return null;
  const number = normalizeImportedBasisNumber(match[1]);
  const date = normalizeImportedBasisDate(match[2]);
  if (!number || !date) return null;
  return { number, date };
};

export const parseBasisOrdersFromText = (text: string): UbdBasisOrderOption[] => {
  const result: UbdBasisOrderOption[] = [];
  const seen = new Set<string>();
  const add = (row: UbdBasisOrderOption | null) => {
    if (!row) return;
    const key = `${row.number}@@${row.date}`;
    if (seen.has(key)) return;
    seen.add(key);
    result.push(row);
  };

  for (const line of String(text ?? "").split(/\r?\n/)) {
    add(parseBasisOrderLine(line));
  }

  for (const match of String(text ?? "").matchAll(BASIS_LINE_PATTERN)) {
    add({
      number: normalizeImportedBasisNumber(match[1]),
      date: normalizeImportedBasisDate(match[2]),
    });
  }

  return result;
};

export const extractDocxParagraphTexts = async (
  buffer: ArrayBuffer,
): Promise<string[]> => {
  const zip = await JSZip.loadAsync(buffer);
  const documentXml = await zip.file("word/document.xml")?.async("string");
  if (!documentXml) {
    throw new Error("У файлі немає word/document.xml — це не Word-документ.");
  }

  const lines: string[] = [];
  for (const para of documentXml.split("</w:p>")) {
    const text = [...para.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/gu)]
      .map((item) => item[1])
      .join("")
      .trim();
    if (text) lines.push(text);
  }
  return lines;
};

export const parseBasisOrdersFromDocx = async (
  buffer: ArrayBuffer,
): Promise<UbdBasisOrderOption[]> => {
  const lines = await extractDocxParagraphTexts(buffer);
  const fromLines = lines
    .map((line) => parseBasisOrderLine(line))
    .filter((row): row is UbdBasisOrderOption => Boolean(row));
  if (fromLines.length) return fromLines;
  return parseBasisOrdersFromText(lines.join("\n"));
};

export const basisOrderKey = (row: UbdBasisOrderOption) =>
  `${row.number}@@${row.date}`;

export const mergeBasisOrderLists = (
  existing: UbdBasisOrderOption[],
  incoming: UbdBasisOrderOption[],
  existingKeys: ReadonlySet<string> = new Set(existing.map(basisOrderKey)),
) => {
  const merged = [...existing];
  const keys = new Set(existingKeys);
  let added = 0;
  let skipped = 0;

  for (const row of incoming) {
    const key = basisOrderKey(row);
    if (keys.has(key)) {
      skipped += 1;
      continue;
    }
    keys.add(key);
    merged.push(row);
    added += 1;
  }

  return { merged, added, skipped };
};
