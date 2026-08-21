import JSZip from "jszip";
import type { UbdRestoreReportFields } from "./ubdRestoreReport";
import {
  buildUbdRestoreBody,
  buildUbdRestorePetition,
  formatGivenSurname,
} from "./ubdRestoreReport";
import { emuMm } from "./ubdWordFormat";
import { stripRedColorInWordZip } from "./wordXml";

const TEMPLATE_URL = `${import.meta.env.BASE_URL}templates/ubd-restore-report.docx`;
const FIGHTER_SIGN_REL = "rId6";
const FIGHTER_SIGN_FILE = "fighter-sign.png";
const COVERING_SIGN_REL = "rId7";
const COVERING_SIGN_FILE = "covering-sign.png";

const SAMPLE = {
  commander: "Командиру батальйону розвідки",
  body:
    "Дійсним доповідаю, що я, молодший сержант ЛУНІН Андрій Володимирович, командир 2 взводу охорони роти охорони військової частини А4862, під час виконання службових обов’язків у темну пору доби, за складних погодних умов та обмеженої видимості, мною було пошкоджено посвідчення учасника бойових дій, серія МВ №039983. Пошкодження сталося ненавмисно, під час активного пересування та виконання поставленого завдання.",
  request:
    "Прошу Вашого клопотання перед вищим командуванням про призначення за даним фактом службового розслідування, оформлення та видачу мені нового посвідчення учасника бойових дій у встановленому порядку.",
  signerTitle1: "Командир 2 взводу охорони",
  signerTitle2: "роти охорони",
  signerLine: "молодший сержант                Андрій ЛУНІН",
  coveringCommander: "Командиру військової частини А4862",
  coveringPetition: "Клопочу по суті рапорту молодшого сержанта Андрія ЛУНІНА",
  coveringSignerTitle: "Командир розвідувального батальйону",
  coveringSignerLine: "підполковник                         Сергій РЕБРОВ",
  coveringUnit: "військової частини А4862",
  coveringDate: "24.02.202",
  approverTitle1: "Тимчасово виконуючий обов’язки",
  approverTitle2: "командира військової частини А4862",
  approverLine:
    "капітан                                                                          Олег АДАМОВ",
  date: "24.02.2026",
};

const escapeXml = (value: string) =>
  value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");

const decodeXml = (value: string) =>
  value
    .replaceAll("&lt;", "<")
    .replaceAll("&gt;", ">")
    .replaceAll("&quot;", '"')
    .replaceAll("&apos;", "'")
    .replaceAll("&amp;", "&");

const withFallback = (value: string, fallback = "______") =>
  value.trim() || fallback;

const isDateLine = (value: string) => /^\d{1,2}\.\d{1,2}\.\d{4}$/.test(value);

const titleLines = (value: string, count: number) => {
  const lines = value
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter((line) => line && !isDateLine(line));
  while (lines.length < count) lines.push("");
  return lines.slice(0, count);
};

const RANK_TAIL =
  /(головний майстер-сержант|старший майстер-сержант|майстер-сержант|штаб-сержант|головний сержант|старший сержант|молодший сержант|старший лейтенант|молодший лейтенант|старший солдат|підполковник|полковник|лейтенант|сержант|капітан|майор|солдат|рекрут)$/iu;

const splitCoveringSigner = (signatory: {
  title: string;
  rank: string;
  fullName: string;
  signatureData?: string | null;
} | null) => {
  if (!signatory) {
    return {
      titleLines: ["", ""] as [string, string],
      rank: "",
      fullName: "",
      date: "",
      signatureData: "",
    };
  }
  const lines = signatory.title
    .split(/\r?\n/)
    .map((line) => line.trim())
    .flatMap((line) => {
      const match = line.match(/^(.*?)(?:\s+)(\d{1,2}\.\d{1,2}\.\d{4})$/);
      if (match?.[1]?.trim()) return [match[1].trim(), match[2]];
      return [line];
    })
    .filter(Boolean);
  const date = lines.find(isDateLine) ?? "";
  let rank = signatory.rank.trim();
  const titles: string[] = [];
  for (const line of lines) {
    if (isDateLine(line)) continue;
    const tail = line.match(RANK_TAIL);
    if (tail && !rank) {
      rank = tail[1];
      const head = line.slice(0, -tail[1].length).trim();
      if (head) titles.push(head);
      continue;
    }
    if (
      rank &&
      line.toLocaleLowerCase("uk-UA") === rank.toLocaleLowerCase("uk-UA")
    ) {
      continue;
    }
    titles.push(line);
  }
  if (titles.length === 1) {
    const split = titles[0].match(/^(.*?)\s+(військової частини\s+\S+.*)$/i);
    if (split) {
      titles[0] = split[1].trim();
      titles[1] = split[2].trim();
    }
  }
  while (titles.length < 2) titles.push("");
  if (titles.length > 2) {
    titles.splice(1, titles.length, titles.slice(1).join(" "));
  }
  return {
    titleLines: [titles[0], titles[1]] as [string, string],
    rank,
    fullName: formatGivenSurname(signatory.fullName) || signatory.fullName.trim(),
    date,
    signatureData: signatory.signatureData?.trim() ?? "",
  };
};

const rankNameLine = (rank: string, fullName: string, pad = 16) => {
  const left = withFallback(rank, "звання");
  const right = formatGivenSurname(fullName) || withFallback(fullName);
  return `${left}${" ".repeat(pad)}${right}`;
};

const ZIP_FILE_OPTIONS = { createFolders: false as const };

const writeZipFile = (zip: JSZip, path: string, data: string | Uint8Array) => {
  zip.file(path, data, ZIP_FILE_OPTIONS);
};

const dataUrlParts = (dataUrl: string) => {
  const [header, base64 = ""] = dataUrl.split(",", 2);
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }
  return { bytes, isPng: header.includes("image/png") };
};

const signatureDrawing = ({
  relId,
  name,
  docPrId,
  offsetX,
  offsetY,
  width,
  height,
}: {
  relId: string;
  name: string;
  docPrId: number;
  offsetX: number;
  offsetY: number;
  width: number;
  height: number;
}) =>
  `<w:r><w:rPr><w:noProof/></w:rPr><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251662336" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>${offsetX}</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>${offsetY}</wp:posOffset></wp:positionV><wp:extent cx="${width}" cy="${height}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/><wp:docPr id="${docPrId}" name="${name}"/><wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="${docPrId}" name="${name}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${relId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${width}" cy="${height}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>`;

const insertDrawingInParagraph = (xml: string, marker: string, drawing: string) =>
  xml.replace(/<w:p\b[\s\S]*?<\/w:p>/g, (paragraph) => {
    const texts: string[] = [];
    paragraph.replace(/<w:t\b[^>]*>([^<]*)<\/w:t>/g, (_, text: string) => {
      texts.push(decodeXml(text));
      return _;
    });
    if (texts.join("") !== marker) {
      return paragraph;
    }
    if (paragraph.includes("</w:pPr>")) {
      return paragraph.replace("</w:pPr>", `</w:pPr>${drawing}`);
    }
    return paragraph.replace(/<w:p\b[^>]*>/, `$&${drawing}`);
  });

const addSignatureImage = async (
  zip: JSZip,
  signatureData: string,
  fileName: string,
  relId: string,
) => {
  const { bytes } = dataUrlParts(signatureData);
  writeZipFile(zip, `word/media/${fileName}`, bytes);

  const relsPath = "word/_rels/document.xml.rels";
  const typesPath = "[Content_Types].xml";
  const rels = await zip.file(relsPath)?.async("string");
  const types = await zip.file(typesPath)?.async("string");
  if (!rels || !types) return relId;

  const relType =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
  const nextRels = rels.includes(`Id="${relId}"`)
    ? rels
    : rels.replace(
        "</Relationships>",
        `<Relationship Id="${relId}" Type="${relType}" Target="media/${fileName}"/></Relationships>`,
      );
  writeZipFile(zip, relsPath, nextRels);

  if (!types.includes('Extension="png"')) {
    writeZipFile(
      zip,
      typesPath,
      types.replace(
        "<Override ",
        '<Default Extension="png" ContentType="image/png"/><Override ',
      ),
    );
  }
  return relId;
};

const replaceParagraphTexts = (
  xml: string,
  replace: (text: string) => string,
) =>
  xml.replace(/<w:p\b[\s\S]*?<\/w:p>/g, (paragraph) => {
    const texts: string[] = [];
    paragraph.replace(/<w:t\b[^>]*>([^<]*)<\/w:t>/g, (_, text: string) => {
      texts.push(decodeXml(text));
      return _;
    });
    if (!texts.length) return paragraph;
    const joined = texts.join("");
    const next = replace(joined);
    if (next === joined) return paragraph;
    const escaped = escapeXml(next);
    let used = false;
    return paragraph.replace(/<w:t\b[^>]*>[^<]*<\/w:t>/g, (run) => {
      if (used) return run.replace(/>[^<]*</, "><");
      used = true;
      return run.replace(/>[^<]*</, `>${escaped}<`);
    });
  });

export const createUbdRestoreWordBlob = async (
  fields: UbdRestoreReportFields,
) => {
  const response = await fetch(TEMPLATE_URL);
  if (!response.ok) {
    throw new Error("Не знайшов шаблон рапорта на відновлення УБД.");
  }
  const zip = await JSZip.loadAsync(await response.arrayBuffer());
  const documentXml = await zip.file("word/document.xml")?.async("string");
  if (!documentXml) {
    throw new Error("Шаблон рапорта на відновлення УБД пошкоджений.");
  }

  const signer = titleLines(fields.signerTitle || fields.staffPosition, 2);
  const covering = fields.signatories.find((item) => item.blockType === "SIGNER");
  const approver = fields.signatories.find(
    (item) => item.blockType === "APPROVAL",
  );
  const coveringParts = splitCoveringSigner(covering ?? null);
  const approverTitle = titleLines(approver?.title ?? "", 2);
  let unitLineHits = 0;

  const filled = replaceParagraphTexts(documentXml, (text) => {
    const trimmed = text.trim();
    if (trimmed === SAMPLE.commander) {
      return withFallback(fields.commander, SAMPLE.commander);
    }
    if (trimmed === SAMPLE.body) {
      return `${buildUbdRestoreBody(fields)} `;
    }
    if (trimmed === SAMPLE.request) {
      return withFallback(fields.requestText, SAMPLE.request);
    }
    if (trimmed === SAMPLE.signerTitle1) {
      return signer[0] ? `${signer[0]} ` : " ";
    }
    if (trimmed === SAMPLE.signerTitle2) {
      return signer[1] || " ";
    }
    if (trimmed === SAMPLE.signerLine) {
      return rankNameLine(fields.rank, fields.fullName, 16);
    }
    if (trimmed === SAMPLE.coveringCommander) {
      return withFallback(
        fields.coveringCommander,
        SAMPLE.coveringCommander,
      );
    }
    if (trimmed === SAMPLE.coveringPetition) {
      return buildUbdRestorePetition(fields);
    }
    if (trimmed === SAMPLE.coveringSignerTitle) {
      return coveringParts.titleLines[0] || SAMPLE.coveringSignerTitle;
    }
    if (trimmed === SAMPLE.coveringUnit) {
      unitLineHits += 1;
      if (unitLineHits === 3 && coveringParts.titleLines[1]) {
        return coveringParts.titleLines[1];
      }
      return text;
    }
    if (trimmed === SAMPLE.coveringSignerLine) {
      return covering
        ? rankNameLine(coveringParts.rank, covering.fullName, 25)
        : SAMPLE.coveringSignerLine;
    }
    if (
      (trimmed === SAMPLE.coveringDate || trimmed === `${SAMPLE.coveringDate}6`) &&
      coveringParts.date
    ) {
      return coveringParts.date;
    }
    if (trimmed === SAMPLE.approverTitle1) {
      return approverTitle[0] || SAMPLE.approverTitle1;
    }
    if (trimmed === SAMPLE.approverTitle2) {
      return approverTitle[1] || SAMPLE.approverTitle2;
    }
    if (trimmed.startsWith("капітан") && trimmed.includes("Олег АДАМОВ")) {
      return approver
        ? `${rankNameLine(approver.rank, approver.fullName, 74)}                         `
        : text;
    }
    if (trimmed === SAMPLE.date) {
      return withFallback(fields.date);
    }
    return text;
  });

  const signerLine = rankNameLine(fields.rank, fields.fullName, 16);
  const coveringLine = covering
    ? rankNameLine(coveringParts.rank, covering.fullName, 25)
    : SAMPLE.coveringSignerLine;
  let withSignature = filled;
  if (fields.signatureData) {
    const relId = await addSignatureImage(
      zip,
      fields.signatureData,
      FIGHTER_SIGN_FILE,
      FIGHTER_SIGN_REL,
    );
    withSignature = insertDrawingInParagraph(
      withSignature,
      signerLine,
      signatureDrawing({
        relId,
        name: "Підпис службовця",
        docPrId: 100,
        offsetX: emuMm(38),
        offsetY: emuMm(-10),
        width: emuMm(48),
        height: emuMm(20),
      }),
    );
  }
  if (coveringParts.signatureData) {
    const relId = await addSignatureImage(
      zip,
      coveringParts.signatureData,
      COVERING_SIGN_FILE,
      COVERING_SIGN_REL,
    );
    withSignature = insertDrawingInParagraph(
      withSignature,
      coveringLine,
      signatureDrawing({
        relId,
        name: "Підпис командира",
        docPrId: 101,
        offsetX: emuMm(32),
        offsetY: emuMm(-8),
        width: emuMm(48),
        height: emuMm(22),
      }),
    );
  }

  zip.file("word/document.xml", withSignature, { createFolders: false });
  await stripRedColorInWordZip(zip);
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
