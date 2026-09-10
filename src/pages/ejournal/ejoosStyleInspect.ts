import JSZip from "jszip";

export type CellStyleInfo = {
  ref: string;
  value: string;
  styleId?: string;
  fontName: string;
  fontSize: number;
  bold: boolean;
  horizontal?: string;
  vertical?: string;
  wrapText: boolean;
};

const parseAttr = (attrs: string, name: string) =>
  attrs.match(new RegExp(`\\b${name}="([^"]*)"`))?.[1];

const columnNumberToLetter = (column: number) => {
  let n = column;
  let letters = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    letters = String.fromCharCode(65 + rem) + letters;
    n = Math.floor((n - 1) / 26);
  }
  return letters;
};

const fontsFromStyles = (stylesXml: string) => {
  const body = stylesXml.match(/<fonts\b[^>]*>([\s\S]*?)<\/fonts>/i)?.[1] ?? "";
  return [...body.matchAll(/<font\b[^>]*(?:\/>|>[\s\S]*?<\/font>)/gi)].map(
    (match) => match[0],
  );
};

const xfsFromStyles = (stylesXml: string) => {
  const body =
    stylesXml.match(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/i)?.[1] ?? "";
  return [...body.matchAll(/<xf\b[^>]*(?:\/>|>[\s\S]*?<\/xf>)/gi)].map(
    (match) => match[0],
  );
};

const fontNameFromFont = (fontXml: string) =>
  fontXml.match(/<name\b[^>]*\bval="([^"]*)"/i)?.[1] ?? "?";

const fontSizeFromFont = (fontXml: string) =>
  Number(fontXml.match(/\bsz val="(\d+)"/i)?.[1] ?? 0) || 0;

const xfFontId = (xfXml: string) => xfXml.match(/\bfontId="(\d+)"/i)?.[1] ?? "0";

const decodeSharedText = (xml: string) =>
  xml
    .replace(/<(?:si|t)[^>]*>/gi, "")
    .replace(/<\/(?:si|t)>/gi, "")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/_x000A_/g, "\n")
    .trim();

const cellValueFromXml = (cellInner: string, cellAttrs: string) => {
  const type = parseAttr(cellAttrs, "t");
  if (type === "s") {
    return { sharedIndex: parseAttr(cellInner, "v") ?? parseAttr(cellAttrs, "v") };
  }
  const inline = cellInner.match(/<is>[\s\S]*?<t[^>]*>([\s\S]*?)<\/t>/i)?.[1];
  if (inline != null) {
    return { text: decodeSharedText(inline) };
  }
  const value = cellInner.match(/<v>([\s\S]*?)<\/v>/i)?.[1];
  return { text: value ?? "" };
};

const resolveSheetPath = async (zip: JSZip, sheetQuery: string) => {
  const workbookXml = await zip.file("xl/workbook.xml")?.async("string");
  if (!workbookXml) throw new Error("Немає xl/workbook.xml");
  const rels = await zip.file("xl/_rels/workbook.xml.rels")?.async("string");
  if (!rels) throw new Error("Немає xl/_rels/workbook.xml.rels");
  const query = sheetQuery.trim().toLowerCase();
  const sheets = [
    ...workbookXml.matchAll(
      /<sheet\b([^>]*\bname="([^"]*)"[^>]*\br:id="([^"]*)")[^>]*\/>/gi,
    ),
  ];
  const match = sheets.find((item) =>
    item[2].toLowerCase().includes(query),
  );
  if (!match) {
    throw new Error(
      `Аркуш «${sheetQuery}» не знайдено. Доступні: ${sheets.map((s) => s[2]).join(", ")}`,
    );
  }
  const relId = match[3];
  const target = rels.match(
    new RegExp(`<Relationship\\b[^>]*\\bId="${relId}"[^>]*\\bTarget="([^"]+)"`, "i"),
  )?.[1];
  if (!target) throw new Error(`Не знайдено шлях для аркуша ${match[2]}`);
  return {
    sheetName: match[2],
    sheetPath: target.startsWith("/")
      ? target.slice(1)
      : `xl/${target.replace(/^\.\//, "")}`,
  };
};

const styleInfoForCell = (
  stylesXml: string,
  styleId: string | undefined,
): Omit<CellStyleInfo, "ref" | "value"> => {
  if (!styleId) {
    return {
      styleId,
      fontName: "?",
      fontSize: 0,
      bold: false,
      wrapText: false,
    };
  }
  const xfs = xfsFromStyles(stylesXml);
  const xf = xfs[Number(styleId)];
  if (!xf) {
    return {
      styleId,
      fontName: "?",
      fontSize: 0,
      bold: false,
      wrapText: false,
    };
  }
  const fonts = fontsFromStyles(stylesXml);
  const font = fonts[Number(xfFontId(xf))] ?? "";
  return {
    styleId,
    fontName: fontNameFromFont(font),
    fontSize: fontSizeFromFont(font),
    bold: /<b\b/i.test(font),
    horizontal: xf.match(/\bhorizontal="([^"]*)"/i)?.[1],
    vertical: xf.match(/\bvertical="([^"]*)"/i)?.[1],
    wrapText: /wrapText="(?:1|true)"/i.test(xf),
  };
};

export const parseRowRange = (spec: string) => {
  const rows = new Set<number>();
  for (const part of spec.split(",")) {
    const trimmed = part.trim();
    if (!trimmed) continue;
    const range = trimmed.match(/^(\d+)\s*-\s*(\d+)$/);
    if (range) {
      const from = Number(range[1]);
      const to = Number(range[2]);
      for (let row = Math.min(from, to); row <= Math.max(from, to); row += 1) {
        rows.add(row);
      }
      continue;
    }
    const single = Number(trimmed);
    if (Number.isFinite(single) && single > 0) rows.add(single);
  }
  return [...rows].sort((a, b) => a - b);
};

export async function inspectWorkbookCellStyles(input: {
  file: Blob | ArrayBuffer;
  sheetQuery: string;
  rows: number[];
  columns?: number[];
}): Promise<CellStyleInfo[]> {
  const zip = await JSZip.loadAsync(
    input.file instanceof Blob ? await input.file.arrayBuffer() : input.file,
  );
  const stylesXml = (await zip.file("xl/styles.xml")?.async("string")) ?? "";
  const sstXml = (await zip.file("xl/sharedStrings.xml")?.async("string")) ?? "";
  const sharedStrings = sstXml
    ? [...sstXml.matchAll(/<si\b[\s\S]*?<\/si>/gi)].map((match) =>
        decodeSharedText(match[0]),
      )
    : [];
  const { sheetPath } = await resolveSheetPath(zip, input.sheetQuery);
  const sheetXml = (await zip.file(sheetPath)?.async("string")) ?? "";
  const columns =
    input.columns ??
    Array.from({ length: 12 }, (_, index) => index + 1);
  const out: CellStyleInfo[] = [];

  for (const row of input.rows) {
    for (const column of columns) {
      const ref = `${columnNumberToLetter(column)}${row}`;
      const cell = sheetXml.match(
        new RegExp(
          `<c\\b([^>]*\\br="${ref}"(?![0-9A-Za-z])[^>]*)(\\/>|>([\\s\\S]*?)<\\/c>)`,
          "i",
        ),
      );
      if (!cell) continue;
      const styleId = parseAttr(cell[1], "s");
      const parsed = cell[3]
        ? cellValueFromXml(cell[3], cell[1])
        : cellValueFromXml("", cell[1]);
      const value =
        "sharedIndex" in parsed && parsed.sharedIndex
          ? sharedStrings[Number(parsed.sharedIndex)] ?? parsed.sharedIndex
          : String(parsed.text ?? "");
      out.push({
        ref,
        value,
        ...styleInfoForCell(stylesXml, styleId),
      });
    }
  }
  return out;
}
