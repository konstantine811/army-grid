import JSZip from "jszip";
import { loadUbdTemplate } from "./ubdTemplateStore";
import { injectUbdApprovalTwoColumnBlock } from "./ubdWordExport";
import { stripRedColorInWordZip } from "./wordXml";

type WordSignatory = {
  blockType: "SIGNER" | "APPROVAL";
  title: string;
  rank: string;
  fullName: string;
  signatureData?: string | null;
};

export type UbdBulkPersonRow = {
  rank: string;
  fullName: string;
  staffPosition: string;
  birthDate: string;
  rnokpp: string;
  taskPeriod: string;
  taskPlace: string;
  basisNumber: string;
  basisDate: string;
};

export type UbdBulkWordFields = {
  commander: string;
  people: UbdBulkPersonRow[];
  signatories: WordSignatory[];
};

const formatSignatoryName = (fullName: string) => {
  const parts = fullName.trim().split(/\s+/).filter(Boolean);
  if (!parts.length) return "";
  if (parts.length === 1) return parts[0].toLocaleUpperCase("uk-UA");
  const surname = parts[parts.length - 1].toLocaleUpperCase("uk-UA");
  const given = parts
    .slice(0, -1)
    .map(
      (part) =>
        part.charAt(0).toLocaleUpperCase("uk-UA") +
        part.slice(1).toLocaleLowerCase("uk-UA"),
    )
    .join("\u00A0");
  return `${given}\u00A0${surname}`;
};

const isDateLine = (value: string) => /^\d{1,2}\.\d{1,2}\.\d{4}$/.test(value);

const personNameLines = (fullName: string) => {
  const parts = fullName
    .trim()
    .replace(/\s*\([^)]+\)\s*$/, "")
    .split(/\s+/)
    .filter(Boolean);
  if (!parts.length) return ["", ""];
  if (parts.length === 1) return [parts[0].toLocaleUpperCase("uk-UA"), ""];
  return [
    parts[0].toLocaleUpperCase("uk-UA"),
    parts
      .slice(1)
      .map((part) => part.charAt(0).toLocaleUpperCase("uk-UA") + part.slice(1))
      .join(" "),
  ];
};

const splitSignatory = (signatory: WordSignatory | null, lineCount: number) => {
  if (!signatory) {
    return {
      titleLines: Array.from({ length: lineCount }, () => ""),
      rank: "",
      date: "",
      fullName: "",
    };
  }
  const rawLines = signatory.title
    .replace(/^ЗАТВЕРДЖУЮ[\s:–—-]*/iu, "")
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);
  const date = rawLines.find(isDateLine) ?? "";
  const rankFromTitle = [...rawLines]
    .reverse()
    .find(
      (line) =>
        !isDateLine(line) &&
        !/командир|тимчасово|військової частини|затверджую/i.test(line),
    );
  const rank = signatory.rank.trim() || rankFromTitle || "";
  const titleLines = rawLines.filter((line) => line !== date && line !== rank);
  const padded = [...titleLines];
  while (padded.length < lineCount) padded.push("");
  if (padded.length > lineCount) {
    padded.splice(
      lineCount - 1,
      padded.length,
      padded.slice(lineCount - 1).join(" "),
    );
  }
  return {
    titleLines: padded.slice(0, lineCount),
    rank,
    date,
    fullName: formatSignatoryName(signatory.fullName),
  };
};

const escapeXml = (value: string) =>
  value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");

const fillPlaceholders = (xml: string, values: Record<string, string>) => {
  let next = xml;
  for (const [key, value] of Object.entries(values)) {
    const token = `{{${key}}}`;
    const safe = escapeXml(value);
    if (next.includes(token)) {
      next = next.replaceAll(token, safe);
      continue;
    }
    const pattern = token
      .split("")
      .map((char) => char.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"))
      .join("(?:<[^>]+>)*");
    next = next.replace(new RegExp(pattern, "g"), safe);
  }
  return next;
};

const ZIP_FILE_OPTIONS = { createFolders: false as const };

const writeZipFile = (zip: JSZip, path: string, data: string | Uint8Array) => {
  zip.file(path, data, ZIP_FILE_OPTIONS);
};

const generateDocxBlob = (zip: JSZip) => {
  for (const [name, entry] of Object.entries(zip.files)) {
    if (entry.dir) zip.remove(name);
  }
  return zip.generateAsync({
    type: "blob",
    mimeType:
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    compression: "DEFLATE",
    compressionOptions: { level: 6 },
  });
};

const dataUrlParts = (dataUrl: string) => {
  const [header, base64 = ""] = dataUrl.split(",", 2);
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }
  const isPng = header.includes("image/png");
  return { bytes, isPng };
};

const replaceSignature = async (zip: JSZip, signatureData: string) => {
  const { bytes, isPng } = dataUrlParts(signatureData);
  const relsPath = "word/_rels/document.xml.rels";
  const typesPath = "[Content_Types].xml";
  const rels = await zip.file(relsPath)?.async("string");
  const types = await zip.file(typesPath)?.async("string");
  if (!rels || !types) return;

  const jpegName = "word/media/image1.jpeg";
  const pngName = "word/media/image1.png";
  if (isPng) {
    zip.remove(jpegName);
    writeZipFile(zip, pngName, bytes);
    writeZipFile(
      zip,
      relsPath,
      rels.replace("media/image1.jpeg", "media/image1.png"),
    );
    writeZipFile(
      zip,
      typesPath,
      types.includes('Extension="png"')
        ? types
        : types.replace(
            '<Default Extension="jpeg"',
            '<Default Extension="png" ContentType="image/png"/><Default Extension="jpeg"',
          ),
    );
    return;
  }
  writeZipFile(zip, jpegName, bytes);
};

const stripUbdBasisNumber = (value: string) =>
  String(value ?? "")
    .trim()
    .replace(/^№\s*/, "");

const parseUaDateKey = (value: string) => {
  const text = String(value ?? "")
    .trim()
    .replaceAll("/", ".")
    .replaceAll("-", ".");
  const match = text.match(/^(\d{1,2})\.(\d{1,2})\.(\d{2,4})$/);
  if (!match) return Number.MAX_SAFE_INTEGER;
  const day = Number(match[1]);
  const month = Number(match[2]);
  let year = Number(match[3]);
  if (year < 100) year += 2000;
  return year * 10_000 + month * 100 + day;
};

/** Усі унікальні БР від старіших до новіших через кому. */
export const buildUbdBulkBasisText = (
  people: Array<{ basisNumber?: string; basisDate?: string; basis?: string }>,
) => {
  const unique = new Map<string, { number: string; date: string; sort: number }>();

  for (const person of people) {
    let number = stripUbdBasisNumber(String(person.basisNumber ?? ""));
    let date = String(person.basisDate ?? "").trim();
    if (!number || !date) {
      const match = String(person.basis ?? "").match(
        /№\s*(\S+)\s+від\s+(\d{1,2}[.\/-]\d{1,2}[.\/-]\d{2,4})/i,
      );
      if (match) {
        number = stripUbdBasisNumber(match[1]);
        date = match[2].replaceAll("/", ".").replaceAll("-", ".");
      }
    }
    if (!number || !date) continue;
    const id = `${number}@@${date}`;
    if (unique.has(id)) continue;
    unique.set(id, { number, date, sort: parseUaDateKey(date) });
  }

  const sorted = [...unique.values()].sort((a, b) => a.sort - b.sort);
  if (!sorted.length) return "бойове розпорядження";
  return `бойове розпорядження ${sorted
    .map((item) => `№${item.number} від ${item.date}`)
    .join(", ")}`;
};

/** Період у таблиці: без «з », з пробілами навколо тире. */
export const formatUbdBulkTaskPeriod = (value: string) => {
  const text = String(value ?? "")
    .trim()
    .replace(/^з\s+/i, "");
  return text.replace(/\s*[-–—]\s*/, " - ");
};

const uniquifyRowIds = (rowXml: string, seed: number) => {
  let nextSeed = seed;
  return rowXml.replace(/w14:paraId="[A-Fa-f0-9]+"/g, () => {
    nextSeed += 1;
    return `w14:paraId="${(0x50000000 + nextSeed).toString(16).toUpperCase()}"`;
  });
};

const fillPersonRow = (rowTemplate: string, person: UbdBulkPersonRow, index: number) => {
  const [surname, givenNames] = personNameLines(person.fullName);
  const filled = fillPlaceholders(rowTemplate, {
    RANK: person.rank.trim(),
    FULL_NAME_1: surname,
    FULL_NAME_2: givenNames,
    POSITION: person.staffPosition.trim(),
    BIRTH_DATE: person.birthDate.trim(),
    RNOKPP: person.rnokpp.trim(),
    TASK_PERIOD: formatUbdBulkTaskPeriod(person.taskPeriod),
    TASK_PLACE: person.taskPlace.trim(),
  });
  return uniquifyRowIds(filled, index * 64);
};

const replaceDataRows = (xml: string, people: UbdBulkPersonRow[]) => {
  const tableStart = xml.indexOf("<w:tbl");
  const tableEnd = xml.indexOf("</w:tbl>");
  if (tableStart < 0 || tableEnd < 0) {
    throw new Error("У шаблоні УБД немає таблиці.");
  }
  const tableXml = xml.slice(tableStart, tableEnd + "</w:tbl>".length);
  const rows = [...tableXml.matchAll(/<w:tr[\s\S]*?<\/w:tr>/g)].map(
    (match) => match[0],
  );
  const dataRowIndex = rows.findIndex((row) => row.includes("{{RANK}}"));
  if (dataRowIndex < 0) {
    throw new Error("У шаблоні УБД немає рядка з даними особи.");
  }
  const dataRowTemplate = rows[dataRowIndex];
  const filledRows = people.map((person, index) =>
    fillPersonRow(dataRowTemplate, person, index + 1),
  );
  const nextTable = tableXml.replace(dataRowTemplate, filledRows.join(""));
  return `${xml.slice(0, tableStart)}${nextTable}${xml.slice(tableEnd + "</w:tbl>".length)}`;
};

const applyPluralIntro = (xml: string, peopleCount: number) => {
  if (peopleCount <= 1) return xml;
  return xml
    .replace(
      "зазначеного нижче військовослужбовця, який розпочав",
      "зазначених нижче військовослужбовців, які розпочали",
    )
    .replace(
      "зазначеного нижче військовослужбовця",
      "зазначених нижче військовослужбовців",
    );
};

export const createUbdBulkWordBlob = async (fields: UbdBulkWordFields) => {
  if (!fields.people.length) {
    throw new Error("Немає осіб для спільного рапорту УБД.");
  }

  const template = await loadUbdTemplate();
  const zip = await JSZip.loadAsync(template);
  const documentPath = "word/document.xml";
  const xml = await zip.file(documentPath)?.async("string");
  if (!xml) {
    throw new Error("У шаблоні немає word/document.xml.");
  }

  const signer =
    fields.signatories.find((item) => item.blockType === "SIGNER") ?? null;
  const approval =
    fields.signatories.find((item) => item.blockType === "APPROVAL") ?? null;
  const signerParts = splitSignatory(signer, 3);
  const basis = buildUbdBulkBasisText(fields.people);

  let filled = replaceDataRows(xml, fields.people);
  filled = applyPluralIntro(filled, fields.people.length);
  filled = injectUbdApprovalTwoColumnBlock(filled, approval);
  filled = fillPlaceholders(filled, {
    COMMANDER: fields.commander.trim(),
    BASIS: basis,
    SIGNER_TITLE_1: signerParts.titleLines[0] ?? "",
    SIGNER_TITLE_2: signerParts.titleLines[1] ?? "",
    SIGNER_TITLE_3: signerParts.titleLines[2] ?? "",
    SIGNER_RANK: signerParts.rank,
    SIGNER_NAME: signerParts.fullName,
  });

  if (signerParts.date) {
    const safeDate = escapeXml(signerParts.date);
    filled = filled.replaceAll("{{SIGNER_DATE}}", safeDate);
    filled = filled.replaceAll("14.07.2026", safeDate);
  }

  writeZipFile(zip, documentPath, filled);

  if (signer?.signatureData) {
    await replaceSignature(zip, signer.signatureData);
  }

  await stripRedColorInWordZip(zip);
  return generateDocxBlob(zip);
};
