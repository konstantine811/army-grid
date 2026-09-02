import JSZip from "jszip";
import type { ServiceCharacteristicFields } from "./serviceCharacteristicReport";
import { formatGivenSurname } from "./ubdRestoreReport";
import { capitalizeReportPosition } from "./reportPosition";
import { stripRedColorInWordZip } from "./wordXml";

const TEMPLATE_URL = `${import.meta.env.BASE_URL}templates/service-characteristic.docx`;

/** Exact joined paragraph texts from the sample template (БРИЛЬ Є.В.). */
const SAMPLE = {
  rankLine: "Військове звання _____________солдат_____________________",
  nameLine: "Прізвище   БРИЛЬ    Ім’я  Євген     По батькові   Володимирович",
  positionLine:
    "Займає посаду командир екіпажу безпілотних авіаційних комплексів - командир 3 відділення ударних безпілотних авіаційних комплексів 4 взводу ударних безпілотних авіаційних комплексів 2 роти ударних безпілотних авіаційних комплексів 2 батальйону безпілотних систем військової частини А4862",
  intro:
    "За час проходження служби на займаній посаді солдат Євген БРИЛЬ зарекомендував себе як дисциплінований, відповідальний та грамотний військовослужбовець.",
  professional:
    "У професійному відношенні підготовлений добре, постійно працює над підвищенням свого професійного рівня. До виконання службових обов’язків ставиться сумлінно та відповідально. Поставлені завдання виконує якісно, точно й у встановлені терміни, проявляючи при цьому розумну ініціативу. ",
  combat:
    "У складних умовах діє впевнено, здатний швидко орієнтуватися в обстановці та приймати обґрунтовані й виважені рішення. Неухильно дотримується вимог військової дисципліни, правил військової ввічливості та носіння військової форми одягу.",
  moral:
    "Морально стійкий, урівноважений та вимогливий до себе. Власну діяльність оцінює критично, на зауваження реагує адекватно, своєчасно усуває виявлені недоліки. У колективі користується заслуженим авторитетом, підтримує доброзичливі та ділові взаємовідносини.",
  drill:
    "У стройовому відношенні підтягнутий, фізично добре розвинений, має високу працездатність. Свої знання, сили та старання спрямовує на сумлінне виконання військового обов’язку, зміцнення обороноздатності України та розбудову Збройних Сил України. До адміністративної відповідальності за вчинення корупційного або військового адміністративного правопорушення не притягувався.",
  conclusion: "ВИСНОВОК: Займаній посаді відповідає.",
  signerTitle1: "Командир ",
  signerTitle2: "2 батальйону безпілотних систем",
  signerTitle3: "військової частини А4862",
  signerLine:
    "лейтенант                                                                                          Сергій ШОСТАК ",
  signerDate: "07.08.2026",
};

const ZIP_FILE_OPTIONS = { createFolders: false as const };

/**
 * Floating signature (Word «Перед текстом»).
 * Height ≈ 26.4mm; width follows image aspect (up to ~59mm).
 */
const SIGNATURE_LAYOUT = {
  relId: "rId4",
  targetCy: 951_269,
  maxCx: 2_137_907,
} as const;

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

const RANK_TAIL =
  /(головний майстер-сержант|старший майстер-сержант|майстер-сержант|штаб-сержант|головний сержант|старший сержант|молодший сержант|старший лейтенант|молодший лейтенант|старший солдат|підполковник|полковник|лейтенант|сержант|капітан|майор|солдат|рекрут)$/iu;

const splitSigner = (signatory: {
  title: string;
  rank: string;
  fullName: string;
} | null) => {
  if (!signatory) {
    return {
      titleLines: ["", "", ""] as [string, string, string],
      rank: "",
      date: "",
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
  while (titles.length < 3) titles.push("");
  if (titles.length > 3) {
    titles.splice(2, titles.length, titles.slice(2).join(" "));
  }
  return {
    titleLines: [titles[0], titles[1], titles[2]] as [string, string, string],
    rank,
    date,
  };
};

const buildRankLine = (rank: string) => {
  const value = withFallback(rank, "______");
  const leftPad = "_".repeat(13);
  const rightPad = "_".repeat(Math.max(5, 21 - value.length));
  return `Військове звання ${leftPad}${value}${rightPad}`;
};

const buildNameLine = (
  lastName: string,
  firstName: string,
  patronymic: string,
) =>
  `Прізвище   ${withFallback(lastName)}    Ім’я  ${withFallback(firstName)}     По батькові   ${withFallback(patronymic)}`;

const buildPositionLine = (staffPosition: string) =>
  `Займає посаду ${withFallback(capitalizeReportPosition(staffPosition))}`;

const buildConclusionLine = (conclusion: string) => {
  const text = conclusion.trim();
  if (!text) return SAMPLE.conclusion;
  return /^висновок\s*:/i.test(text) ? text : `ВИСНОВОК: ${text}`;
};

const paragraphJoinedText = (paragraph: string) => {
  const texts: string[] = [];
  paragraph.replace(/<w:t\b[^>]*>([^<]*)<\/w:t>/g, (_, text: string) => {
    texts.push(decodeXml(text));
    return _;
  });
  return texts.join("");
};

const replaceParagraphTexts = (
  xml: string,
  replace: (text: string) => string,
) =>
  xml.replace(/<w:p\b[\s\S]*?<\/w:p>/g, (paragraph) => {
    const joined = paragraphJoinedText(paragraph);
    if (!joined.length) return paragraph;
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

const RUN_FONTS =
  '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>';

const textParagraph = (text: string, align: "left" | "right" = "left") => {
  const body = text.trim()
    ? `<w:r><w:rPr>${RUN_FONTS}<w:noProof/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r>`
    : `<w:r><w:rPr>${RUN_FONTS}<w:noProof/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr></w:r>`;
  return `<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/><w:jc w:val="${align}"/><w:rPr>${RUN_FONTS}<w:noProof/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr></w:pPr>${body}</w:p>`;
};

const frontOfTextSignatureDrawing = (
  relId: string,
  cx: number,
  cy: number,
  docPrId: number,
) =>
  `<w:r><w:rPr><w:noProof/></w:rPr><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="251662336" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:align>center</wp:align></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/><wp:docPr id="${docPrId}" name="Підпис"/><wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="0"/></wp:cNvGraphicFramePr><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="0" name="Підпис"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${relId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>`;

const signatureImageParagraph = (drawingRun: string) =>
  `<w:p><w:pPr><w:spacing w:after="0" w:before="0" w:line="240" w:lineRule="auto"/><w:jc w:val="center"/><w:rPr>${RUN_FONTS}<w:noProof/><w:sz w:val="28"/></w:rPr></w:pPr>${drawingRun}</w:p>`;

/** Two-column footer: titles/rank/date | floating signature + name (bottom-aligned). */
const buildSignerFooterTable = ({
  titleLines,
  rank,
  fullName,
  date,
  signatureRun,
}: {
  titleLines: [string, string, string];
  rank: string;
  fullName: string;
  date: string;
  signatureRun: string;
}) => {
  const leftLines = [
    ...titleLines.map((line) => line.trim()).filter(Boolean),
    rank.trim(),
    date.trim(),
  ].filter(Boolean);
  const leftXml = leftLines.map((line) => textParagraph(line, "left")).join("");
  const displayName = (
    formatGivenSurname(fullName) || withFallback(fullName)
  ).replace(/\s+/g, "\u00A0");
  const rightXml = [
    signatureRun ? signatureImageParagraph(signatureRun) : textParagraph(""),
    textParagraph(displayName, "right"),
  ].join("");

  return `<w:tbl><w:tblPr><w:tblW w:w="9640" w:type="dxa"/><w:tblBorders><w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/><w:insideH w:val="nil"/><w:insideV w:val="nil"/></w:tblBorders><w:tblLayout w:type="fixed"/><w:tblLook w:val="04A0" w:firstRow="0" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="0" w:noVBand="0"/></w:tblPr><w:tblGrid><w:gridCol w:w="5200"/><w:gridCol w:w="4440"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="5200" w:type="dxa"/><w:vAlign w:val="top"/></w:tcPr>${leftXml || textParagraph("")}</w:tc><w:tc><w:tcPr><w:tcW w:w="4440" w:type="dxa"/><w:vAlign w:val="bottom"/></w:tcPr>${rightXml}</w:tc></w:tr></w:tbl><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:p>`;
};

const replaceSignerFooterWithTable = (xml: string, tableXml: string) => {
  const matches = [...xml.matchAll(/<w:p\b[\s\S]*?<\/w:p>/g)];
  let start = -1;
  let end = -1;
  for (const match of matches) {
    const text = paragraphJoinedText(match[0]);
    const trimmed = text.trim();
    if (
      start < 0 &&
      (text === SAMPLE.signerTitle1 || trimmed === SAMPLE.signerTitle1.trim())
    ) {
      start = match.index ?? -1;
    }
    if (
      start >= 0 &&
      (trimmed === SAMPLE.signerDate || text === SAMPLE.signerDate)
    ) {
      end = (match.index ?? 0) + match[0].length;
    }
  }
  if (start < 0 || end < 0 || end <= start) return xml;
  const next = `${xml.slice(0, start)}${tableXml}${xml.slice(end)}`;
  return next.replace(/<w:drawing>[\s\S]*?<\/w:drawing>/g, (drawing) => {
    if (
      drawing.includes(`r:embed="${SIGNATURE_LAYOUT.relId}"`) &&
      drawing.includes("wp:anchor")
    ) {
      return "";
    }
    return drawing;
  });
};

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

const readU32BE = (bytes: Uint8Array, offset: number) =>
  ((bytes[offset] << 24) |
    (bytes[offset + 1] << 16) |
    (bytes[offset + 2] << 8) |
    bytes[offset + 3]) >>>
  0;

const readU16BE = (bytes: Uint8Array, offset: number) =>
  (bytes[offset] << 8) | bytes[offset + 1];

const readRasterSize = (
  bytes: Uint8Array,
): { width: number; height: number } | null => {
  if (
    bytes.length >= 24 &&
    bytes[0] === 0x89 &&
    bytes[1] === 0x50 &&
    bytes[2] === 0x4e &&
    bytes[3] === 0x47
  ) {
    const width = readU32BE(bytes, 16);
    const height = readU32BE(bytes, 20);
    return width > 0 && height > 0 ? { width, height } : null;
  }
  if (bytes.length > 4 && bytes[0] === 0xff && bytes[1] === 0xd8) {
    let offset = 2;
    while (offset + 9 < bytes.length) {
      if (bytes[offset] !== 0xff) {
        offset += 1;
        continue;
      }
      const marker = bytes[offset + 1];
      if (marker === 0xc0 || marker === 0xc1 || marker === 0xc2) {
        const height = readU16BE(bytes, offset + 5);
        const width = readU16BE(bytes, offset + 7);
        return width > 0 && height > 0 ? { width, height } : null;
      }
      if (marker === 0xd9 || marker === 0xda) break;
      if (marker === 0x01 || (marker >= 0xd0 && marker <= 0xd9)) {
        offset += 2;
        continue;
      }
      const length = readU16BE(bytes, offset + 2);
      if (length < 2) break;
      offset += 2 + length;
    }
  }
  return null;
};

const extentForImage = (imageWidth: number, imageHeight: number) => {
  const ratio =
    imageWidth > 0 && imageHeight > 0 ? imageWidth / imageHeight : 1.294;
  let cy: number = SIGNATURE_LAYOUT.targetCy;
  let cx = Math.max(1, Math.round(cy * ratio));
  if (cx > SIGNATURE_LAYOUT.maxCx) {
    cx = SIGNATURE_LAYOUT.maxCx;
    cy = Math.max(1, Math.round(cx / ratio));
  }
  return { cx, cy };
};

const replaceRelImage = async (
  zip: JSZip,
  relId: string,
  signatureData: string,
  fileName: string,
) => {
  const { bytes, isPng } = dataUrlParts(signatureData);
  const relsPath = "word/_rels/document.xml.rels";
  const typesPath = "[Content_Types].xml";
  const rels = await zip.file(relsPath)?.async("string");
  const types = await zip.file(typesPath)?.async("string");
  if (!rels || !types) return bytes;

  const ext = isPng ? "png" : "jpeg";
  const path = `word/media/${fileName}.${ext}`;
  writeZipFile(zip, path, bytes);
  const nextRels = rels.replace(
    new RegExp(`(<Relationship Id="${relId}"[^>]*Target=")media/[^"]+(")`),
    `$1media/${fileName}.${ext}$2`,
  );
  writeZipFile(zip, relsPath, nextRels);
  if (isPng && !types.includes('Extension="png"')) {
    writeZipFile(
      zip,
      typesPath,
      types.replace(
        '<Default Extension="jpeg"',
        '<Default Extension="png" ContentType="image/png"/><Default Extension="jpeg"',
      ),
    );
  }
  return bytes;
};

export const createServiceCharacteristicWordBlob = async (
  fields: ServiceCharacteristicFields,
) => {
  const response = await fetch(TEMPLATE_URL);
  if (!response.ok) {
    throw new Error("Не знайшов шаблон службової характеристики.");
  }
  const zip = await JSZip.loadAsync(await response.arrayBuffer());
  const documentXml = await zip.file("word/document.xml")?.async("string");
  if (!documentXml) {
    throw new Error("Шаблон службової характеристики пошкоджений.");
  }

  const signer =
    fields.signatories.find((item) => item.blockType === "SIGNER") ??
    fields.signatories[0] ??
    null;
  const signerParts = splitSigner(signer);
  const signerDate = signerParts.date || fields.date.trim();

  const professional = fields.professionalParagraph.trimEnd();
  const professionalWithSpace = professional.endsWith(" ")
    ? professional
    : `${professional} `;

  // Keep SAMPLE signer markers intact so we can swap the whole footer for a table.
  let filled = replaceParagraphTexts(documentXml, (text) => {
    const trimmed = text.trim();
    if (trimmed === SAMPLE.rankLine.trim() || text === SAMPLE.rankLine) {
      return buildRankLine(fields.rank);
    }
    if (trimmed === SAMPLE.nameLine.trim() || text === SAMPLE.nameLine) {
      return buildNameLine(fields.lastName, fields.firstName, fields.patronymic);
    }
    if (trimmed === SAMPLE.positionLine.trim() || text === SAMPLE.positionLine) {
      return buildPositionLine(fields.staffPosition);
    }
    if (trimmed === SAMPLE.intro.trim() || text === SAMPLE.intro) {
      return withFallback(fields.introParagraph, SAMPLE.intro);
    }
    if (
      trimmed === SAMPLE.professional.trim() ||
      text === SAMPLE.professional
    ) {
      return professionalWithSpace;
    }
    if (trimmed === SAMPLE.combat.trim() || text === SAMPLE.combat) {
      return withFallback(fields.combatParagraph, SAMPLE.combat);
    }
    if (trimmed === SAMPLE.moral.trim() || text === SAMPLE.moral) {
      return withFallback(fields.moralParagraph, SAMPLE.moral);
    }
    if (trimmed === SAMPLE.drill.trim() || text === SAMPLE.drill) {
      return withFallback(fields.drillParagraph, SAMPLE.drill);
    }
    if (trimmed === SAMPLE.conclusion.trim() || text === SAMPLE.conclusion) {
      return buildConclusionLine(fields.conclusion);
    }
    return text;
  });

  let imageWidth = 0;
  let imageHeight = 0;
  if (signer?.signatureData) {
    const bytes = await replaceRelImage(
      zip,
      SIGNATURE_LAYOUT.relId,
      signer.signatureData,
      "commander-sign",
    );
    const size = bytes ? readRasterSize(bytes) : null;
    imageWidth = size?.width ?? 0;
    imageHeight = size?.height ?? 0;
  } else {
    const templateImage =
      (await zip.file("word/media/image1.jpeg")?.async("uint8array")) ??
      (await zip.file("word/media/image1.png")?.async("uint8array"));
    const size = templateImage ? readRasterSize(templateImage) : null;
    imageWidth = size?.width ?? 0;
    imageHeight = size?.height ?? 0;
  }

  const { cx, cy } = extentForImage(imageWidth, imageHeight);
  const signatureRun =
    signer?.signatureData || imageWidth
      ? frontOfTextSignatureDrawing(SIGNATURE_LAYOUT.relId, cx, cy, 1746884460)
      : "";

  const table = buildSignerFooterTable({
    titleLines: signerParts.titleLines,
    rank: signerParts.rank || (signer?.rank ?? ""),
    fullName: signer?.fullName ?? "",
    date: signerDate,
    signatureRun,
  });
  filled = replaceSignerFooterWithTable(filled, table);

  zip.file("word/document.xml", filled, { createFolders: false });
  await stripRedColorInWordZip(zip);
  for (const [name, entry] of Object.entries(zip.files)) {
    if (entry.dir) zip.remove(name);
  }
  return zip.generateAsync({
    type: "blob",
    mimeType:
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  });
};
