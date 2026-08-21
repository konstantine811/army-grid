import type JSZip from "jszip";

const RED_COLOR = /w:val="(?:FF0000|EE0000|C00000|CC0000)"/i;

export const stripWordRedColor = (xml: string) =>
  xml.replace(/<w:color\b([^>]*)\/>/g, (full, attrs: string) => {
    if (!RED_COLOR.test(attrs)) return full;
    return `<w:color${attrs.replace(RED_COLOR, 'w:val="000000"')}/>`;
  });

export const stripRedColorInWordZip = async (zip: JSZip) => {
  await Promise.all(
    Object.entries(zip.files).map(async ([name, file]) => {
      if (file.dir || !name.startsWith("word/") || !name.endsWith(".xml")) {
        return;
      }
      const xml = await file.async("string");
      const next = stripWordRedColor(xml);
      if (next !== xml) {
        zip.file(name, next, { createFolders: false });
      }
    }),
  );
};
