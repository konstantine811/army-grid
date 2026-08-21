import JSZip from "jszip";
import { loadUbdTemplate } from "./ubdTemplateStore";
import { stripRedColorInWordZip } from "./wordXml";

type WordSignatory = {
  blockType: "SIGNER" | "APPROVAL";
  title: string;
  rank: string;
  fullName: string;
  signatureData?: string | null;
};

export type UbdWordFields = {
  commander: string;
  fullName: string;
  rank: string;
  staffPosition: string;
  birthDate: string;
  rnokpp: string;
  taskPeriod: string;
  taskPlace: string;
  basis: string;
  signatories: WordSignatory[];
};

export const UBD_TEMPLATE_FIELDS = [
  { key: "COMMANDER", label: "Адресат (Командиру …)" },
  { key: "RANK", label: "Військове звання" },
  { key: "FULL_NAME_1", label: "Прізвище" },
  { key: "FULL_NAME_2", label: "Ім’я та по батькові" },
  { key: "POSITION", label: "Посада згідно штату" },
  { key: "BIRTH_DATE", label: "Дата народження" },
  { key: "RNOKPP", label: "РНОКПП" },
  { key: "TASK_PERIOD", label: "Період виконання завдань" },
  { key: "TASK_PLACE", label: "Місце виконання завдань" },
  { key: "BASIS", label: "Підстава" },
  { key: "SIGNER_TITLE_1", label: "Підписант, рядок 1" },
  { key: "SIGNER_TITLE_2", label: "Підписант, рядок 2" },
  { key: "SIGNER_TITLE_3", label: "Підписант, рядок 3" },
  { key: "SIGNER_RANK", label: "Звання підписанта" },
  { key: "SIGNER_NAME", label: "ПІБ підписанта" },
  { key: "APPROVER_TITLE_1", label: "Затверджую, рядок 1" },
  { key: "APPROVER_TITLE_2", label: "Затверджую, рядок 2" },
  { key: "APPROVER_RANK", label: "Звання затверджуючого" },
  { key: "APPROVER_NAME", label: "ПІБ затверджуючого" },
] as const;

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

const valuesFromFields = (fields: UbdWordFields) => {
  const signer =
    fields.signatories.find((item) => item.blockType === "SIGNER") ?? null;
  const approval =
    fields.signatories.find((item) => item.blockType === "APPROVAL") ?? null;
  const signerParts = splitSignatory(signer, 3);
  const approvalParts = splitSignatory(approval, 2);
  const [surname, givenNames] = personNameLines(fields.fullName);

  return {
    COMMANDER: fields.commander.trim(),
    RANK: fields.rank.trim(),
    FULL_NAME_1: surname,
    FULL_NAME_2: givenNames,
    POSITION: fields.staffPosition.trim(),
    BIRTH_DATE: fields.birthDate.trim(),
    RNOKPP: fields.rnokpp.trim(),
    TASK_PERIOD: fields.taskPeriod.trim(),
    TASK_PLACE: fields.taskPlace.trim(),
    BASIS: fields.basis.trim(),
    SIGNER_TITLE_1: signerParts.titleLines[0] ?? "",
    SIGNER_TITLE_2: signerParts.titleLines[1] ?? "",
    SIGNER_TITLE_3: signerParts.titleLines[2] ?? "",
    SIGNER_RANK: signerParts.rank,
    SIGNER_NAME: signerParts.fullName,
    SIGNER_DATE: signerParts.date,
    APPROVER_TITLE_1: approvalParts.titleLines[0] ?? "",
    APPROVER_TITLE_2: approvalParts.titleLines[1] ?? "",
    APPROVER_RANK: approvalParts.rank,
    APPROVER_NAME: approvalParts.fullName,
  };
};

export const createUbdWordBlob = async (fields: UbdWordFields) => {
  const template = await loadUbdTemplate();
  const zip = await JSZip.loadAsync(template);
  const documentPath = "word/document.xml";
  const xml = await zip.file(documentPath)?.async("string");
  if (!xml) {
    throw new Error("У шаблоні немає word/document.xml.");
  }
  const values = valuesFromFields(fields);
  const { SIGNER_DATE, ...tokens } = values;
  let filled = fillPlaceholders(xml, tokens);
  if (SIGNER_DATE) {
    const safeDate = escapeXml(SIGNER_DATE);
    filled = filled.replaceAll("{{SIGNER_DATE}}", safeDate);
    filled = filled.replaceAll("14.07.2026", safeDate);
  }
  writeZipFile(zip, documentPath, filled);

  const signer = fields.signatories.find((item) => item.blockType === "SIGNER");
  if (signer?.signatureData) {
    await replaceSignature(zip, signer.signatureData);
  }

  await stripRedColorInWordZip(zip);
  return generateDocxBlob(zip);
};
