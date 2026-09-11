import JSZip from "jszip";
import {
  EJOOS_TEXT_NUM_FMT_ID,
  isEjoosStaffIndexColumn,
} from "./ejoosStaffIndexFormat";
import { resolveSheetPath } from "./ejoosZipCellWrites";

/** Final download pass: only style definitions and style references may change. */
export async function styleEjoosExport(file: Blob): Promise<Blob> {
  const zip = await JSZip.loadAsync(await file.arrayBuffer());
  let styles = await zip.file("xl/styles.xml")?.async("string");
  const workbook = await zip.file("xl/workbook.xml")?.async("string");
  if (!styles || !workbook) throw new Error("Не знайдено стилі або структуру книги ЕЖООС");
  const attr = (xml: string, name: string) => xml.match(new RegExp(`\\b${name}="([^"]*)"`))?.[1];
  const set = (xml: string, name: string, value: string) => {
    const pattern = new RegExp(`\\b${name}="[^"]*"`);
    return pattern.test(xml) ? xml.replace(pattern, `${name}="${value}"`)
      : xml.replace(/\s*\/?>(?=[^>]*$)/, end => ` ${name}="${value}"${end}`);
  };
  const items = (collection: string, tag: string) => {
    const body = styles!.match(new RegExp(`<${collection}\\b[^>]*>([\\s\\S]*?)</${collection}>`))?.[1];
    if (body == null) throw new Error(`Немає ${collection} у стилях ЕЖООС`);
    return [...body.matchAll(new RegExp(`<${tag}\\b[^>]*?(?:/>|>[\\s\\S]*?</${tag}>)`, "g"))].map(match => match[0]);
  };
  const fonts = items("fonts", "font");
  const borders = items("borders", "border");
  const xfs = items("cellXfs", "xf");
  const originalXfs = [...xfs];
  const intern = (list: string[], value: string) => {
    const found = list.indexOf(value);
    if (found >= 0) return found;
    list.push(value);
    return list.length - 1;
  };
  const border = '<border><left style="thin"><color auto="1"/></left><right style="thin"><color auto="1"/></right><top style="thin"><color auto="1"/></top><bottom style="thin"><color auto="1"/></bottom><diagonal/></border>';
  const decode = (value: string) => value.replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&amp;/g, "&");
  for (const sheet of workbook.matchAll(/<sheet\b[^>]*\/?\s*>/g)) {
    const name = decode(attr(sheet[0], "name") || "");
    const wide = /(?:^|[.\s])ООС(?:$|\s)|виключен/i.test(name);
    const compact = /шпо|штатно.?посад|табель/i.test(name);
    if (!wide && !compact) continue;
    const path = await resolveSheetPath(zip, name);
    if (!path) continue;
    const xml = await zip.file(path)?.async("string");
    if (!xml) continue;
    const cache = new Map<string, string>();
    const restyle = (sourceId: string, column = 0, row = 0) => {
      const forceTextFormat =
        compact &&
        column > 0 &&
        isEjoosStaffIndexColumn(name, column, row);
      const cacheKey = `${sourceId}:${forceTextFormat ? "t" : "n"}`;
      const cached = cache.get(cacheKey);
      if (cached != null) return cached;
      const source = originalXfs[Number(sourceId)] || originalXfs[0];
      let font = fonts[Number(attr(source, "fontId") || "0")];
      font = font.replace(/<font\s*\/>/, "<font></font>");
      font = font.replace(/<(?:name|b|scheme)\b[^>]*?(?:\/>|>[\s\S]*?<\/(?:name|b|scheme)>)/g, "");
      if (wide || compact) font = font.replace(/<sz\b[^>]*?(?:\/>|>[\s\S]*?<\/sz>)/g, "");
      font = font.replace("</font>", `<name val="Times New Roman"/>${wide || compact ? `<sz val="${wide ? 14 : 12}"/>` : ""}</font>`);
      let open = source.match(/^<xf\b[^>]*>/)![0];
      open = set(set(open, "fontId", String(intern(fonts, font))), "applyFont", "1");
      open = set(open, "applyAlignment", "1");
      if (wide) open = set(set(open, "borderId", String(intern(borders, border))), "applyBorder", "1");
      if (forceTextFormat) {
        open = set(
          set(open, "numFmtId", EJOOS_TEXT_NUM_FMT_ID),
          "applyNumberFormat",
          "1",
        );
      }
      let alignment = source.match(/<alignment\b[^>]*\/>/)?.[0] || "<alignment/>";
      alignment = set(set(alignment, "horizontal", "center"), "vertical", "center");
      if (wide || compact) alignment = set(alignment, "wrapText", wide ? "1" : "0");
      const body = source.slice(source.indexOf(">") + 1).replace(/<\/xf>$/, "").replace(/<alignment\b[^>]*?(?:\/>|>[\s\S]*?<\/alignment>)/g, "");
      const next = `${open.replace(/\/>$/, ">")}${alignment}${body}</xf>`;
      const id = String(intern(xfs, next));
      cache.set(cacheKey, id);
      return id;
    };
    // Do not rebuild sheet XML: retain formulas, merges, dimensions and all values verbatim.
    const columns = [...xml.matchAll(/<col\b[^>]*>/g)].map(match => ({
      min: Number(attr(match[0], "min")), max: Number(attr(match[0], "max")),
      style: attr(match[0], "style"),
    }));
    let rowStyle: string | undefined;
    const next = xml.replace(/<(c|row|col)\b[^>]*>/g, (tag, kind: string) => {
      const key = kind === "col" ? "style" : "s";
      const id = attr(tag, key);
      if (kind === "row") rowStyle = id;
      if (kind === "row" && id == null) return tag;
      const letters = attr(tag, "r")?.match(/^[A-Z]+/)?.[0] || "";
      const column = [...letters].reduce((number, letter) => number * 26 + letter.charCodeAt(0) - 64, 0);
      const inherited = kind === "c" ? rowStyle ?? columns.find(entry => entry.min <= column && entry.max >= column)?.style : undefined;
      const rowNum = Number(attr(tag, "r")?.match(/\d+$/)?.[0] || 0);
      return set(tag, key, restyle(id ?? inherited ?? "0", column, rowNum));
    });
    zip.file(path, next, { createFolders: false });
  }
  for (const [name, list] of [["fonts", fonts], ["borders", borders], ["cellXfs", xfs]] as const) {
    styles = styles.replace(new RegExp(`<${name}\\b[^>]*>[\\s\\S]*?</${name}>`), `<${name} count="${list.length}">${list.join("")}</${name}>`);
  }
  zip.file("xl/styles.xml", styles, { createFolders: false });
  return zip.generateAsync({ type: "blob", compression: "DEFLATE", mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
}
