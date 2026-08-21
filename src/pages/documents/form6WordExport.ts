import JSZip from "jszip";
import type { Form6ReportFields } from "./form6Report";
import { toUkrainianGenitiveFullName } from "./form6Report";
import { formatGivenSurname } from "./ubdRestoreReport";
import { stripRedColorInWordZip } from "./wordXml";

const FORM6_TEMPLATE_URL = `${import.meta.env.BASE_URL}templates/form6-report.docx`;

const SAMPLE = {
  personLine: "Солдат БИЧЕНКО Микола Анатолійович",
  position:
    "Оператор відділення радіоелектронної боротьби 1 піхотного батальйону військової частини А4862",
  birthDate: "25.12.1980",
  idDocument: "Паспорт громадянина України Серія КВ №001828",
  rnokpp: "2957921556",
  address: "Лівьвська обл., м. Львів вул. Окружна б. 36 кв. 17",
  phone: "+380938907068",
  period: "з 16.10.2025  24.10.2025",
  place: "н.п.  Олексіївка, Донецька обл.",
  basis:
    "Підстава: Бойове розпорядження командира 425 ОШП «СКЕЛЯ»  №4862/ОКП/1158/дск від 14.10.2025",
  genitiveName: "БИЧЕНКА Миколи Анатолійовича",
  commander:
    "Командиру 1 піхотного батальйону військової частини А4862",
  signerTitle1: "Командир 1 піхотного батальйону",
  signerTitle2: "військової частини А4862",
  signerLine:
    "старший лейтенант                                                                   Єгор СИДОРЕНКО",
  signerDate: "06.08.2026",
  approverTitle1: "Тимчасово виконуючий обов’язки",
  approverTitle2: "командира військової частини А4862",
  approverLine: "капітанОлег АДАМОВ",
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

const RANK_TAIL =
  /(головний майстер-сержант|старший майстер-сержант|майстер-сержант|штаб-сержант|головний сержант|старший сержант|молодший сержант|старший лейтенант|молодший лейтенант|старший солдат|підполковник|полковник|лейтенант|сержант|капітан|майор|солдат|рекрут)$/iu;

const splitForm6Signer = (signatory: {
  title: string;
  rank: string;
  fullName: string;
} | null) => {
  if (!signatory) {
    return { titleLines: ["", ""] as [string, string], rank: "", date: "" };
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
  return { titleLines: [titles[0], titles[1]] as [string, string], rank, date };
};

const rankNameLine = (rank: string, fullName: string, pad: number) => {
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
  if (!rels || !types) return;

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
};

const periodText = (value: string) => {
  const text = value.trim();
  if (!text) return "з ______";
  return /^з\s/i.test(text) ? text : `з ${text}`;
};

const placeText = (value: string) => {
  const text = value.trim();
  if (!text) return "н.п. ______";
  return /^н\.?п\.?/i.test(text) ? text : `н.п. ${text}`;
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

export const createForm6WordBlob = async (fields: Form6ReportFields) => {
  const response = await fetch(FORM6_TEMPLATE_URL);
  if (!response.ok) {
    throw new Error("Не знайшов шаблон Форми 6.");
  }
  const zip = await JSZip.loadAsync(await response.arrayBuffer());
  const documentXml = await zip.file("word/document.xml")?.async("string");
  if (!documentXml) {
    throw new Error("Шаблон Форми 6 пошкоджений.");
  }

  const personLine = `${withFallback(fields.rank, "звання")} ${withFallback(fields.fullName)}`.trim();
  const genitiveName = toUkrainianGenitiveFullName(fields.fullName) || withFallback(fields.fullName);
  const signer =
    fields.signatories.find((item) => item.blockType === "SIGNER") ?? null;
  const approval =
    fields.signatories.find((item) => item.blockType === "APPROVAL") ?? null;
  const signerParts = splitForm6Signer(signer);
  const approvalParts = splitForm6Signer(approval);

  const filled = replaceParagraphTexts(documentXml, (text) => {
    const trimmed = text.trim();
    if (trimmed === SAMPLE.signerTitle1) {
      return signerParts.titleLines[0] || SAMPLE.signerTitle1;
    }
    if (trimmed === SAMPLE.signerTitle2) {
      return signerParts.titleLines[1] || SAMPLE.signerTitle2;
    }
    if (trimmed === SAMPLE.signerLine) {
      return signer
        ? rankNameLine(signerParts.rank, signer.fullName, 67)
        : SAMPLE.signerLine;
    }
    if (trimmed === SAMPLE.signerDate && signerParts.date) {
      return signerParts.date;
    }
    if (trimmed === SAMPLE.approverTitle1) {
      return approvalParts.titleLines[0] || SAMPLE.approverTitle1;
    }
    if (trimmed === SAMPLE.approverTitle2) {
      return approvalParts.titleLines[1] || SAMPLE.approverTitle2;
    }
    if (trimmed === SAMPLE.approverLine) {
      return approval
        ? rankNameLine(approvalParts.rank, approval.fullName, 67)
        : SAMPLE.approverLine;
    }
    return text
      .replaceAll(SAMPLE.personLine, personLine)
      .replaceAll(SAMPLE.commander, withFallback(fields.commander, SAMPLE.commander))
      .replaceAll(SAMPLE.position, withFallback(fields.staffPosition))
      .replaceAll(SAMPLE.birthDate, withFallback(fields.birthDate))
      .replaceAll(SAMPLE.idDocument, withFallback(fields.idDocument))
      .replaceAll(SAMPLE.rnokpp, withFallback(fields.rnokpp))
      .replaceAll(SAMPLE.address, withFallback(fields.address))
      .replaceAll(SAMPLE.phone, withFallback(fields.phone))
      .replaceAll(SAMPLE.period, periodText(fields.taskPeriod))
      .replaceAll(SAMPLE.place, placeText(fields.taskPlace))
      .replaceAll(SAMPLE.basis, `Підстава: ${withFallback(fields.basis)}`)
      .replaceAll(SAMPLE.genitiveName, genitiveName);
  });

  if (signer?.signatureData) {
    await replaceRelImage(zip, "rId8", signer.signatureData, "commander-sign");
  }

  zip.file("word/document.xml", filled, { createFolders: false });
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
