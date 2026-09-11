/** «Куди» у Рух часто = «x» / порожньо; тоді військова частина є в примітці. */
export const resolveMovementDestination = (rawDest: string, note: string) => {
  const clean = (value: string) => {
    const text = value.replace(/\s+/g, " ").trim();
    if (!text) return "";
    const upper = text.toUpperCase();
    if (
      upper === "X" ||
      upper === "Х" ||
      upper === "0" ||
      upper === "0.0" ||
      upper === "-" ||
      upper === "—" ||
      upper === "." ||
      upper.includes("РОЗПОР") ||
      upper.includes("ПЕРЕВ") ||
      upper.includes("ПОСАД") ||
      upper.includes("ПРИБ") ||
      upper.includes("ЗВІЛ")
    ) {
      return "";
    }
    return text;
  };
  const destination = clean(rawDest);
  if (destination) return destination;
  const noteText = note.replace(/\s+/g, " ").trim();
  return /(?:[АA]\s*\d{4}(?!\d)|військов(?:ої|а)\s+частин)/iu.test(noteText)
    ? noteText
    : "";
};

const isSzchCancellation = (event: {
  type: string;
  status: string;
  note: string;
  changeText: string;
}) =>
  /СКАС(?:УВАННЯ|ОВАНО|УВАТИ)?.*СЗЧ|СЗЧ.*СКАС/iu.test(
    [event.type, event.status, event.note, event.changeText].join(" "),
  );

export { isSzchCancellation };

const movementBlob = (event: {
  type: string;
  status: string;
  note: string;
  changeText: string;
  destination: string;
}) =>
  [
    event.type,
    event.status,
    event.note,
    event.changeText,
    event.destination,
  ].join(" ");

/** Окремий рядок «СКАСУВАННЯ переведення», а не скасований рядок ПЕРЕВ. */
export const isTransferCancellation = (event: {
  type: string;
  status: string;
  note: string;
  changeText: string;
  destination: string;
}) => {
  if (isSzchCancellation(event)) return false;
  if (
    event.type === "ПЕРЕВ" ||
    event.type === "ПОСАДА" ||
    event.type === "ЗВАННЯ"
  ) {
    return false;
  }
  if (event.type === "СКАСУВАННЯ") return true;
  const text = movementBlob(event);
  return /скас(?:уванн|овано|увано|увати)/iu.test(text) && /перев/iu.test(text);
};

/** Анульований рядок РУХ: «скасовано» у статусі, примітці або «куди». */
export const isCancelledMovementRecord = (event: {
  type: string;
  status: string;
  note: string;
  changeText: string;
  destination: string;
}) => {
  if (isSzchCancellation(event) || isTransferCancellation(event)) return false;
  const fields = [
    event.status,
    event.note,
    event.destination,
    event.changeText,
  ];
  return fields.some((value) => {
    const text = String(value ?? "").trim();
    if (!text) return false;
    if (/^скас(?:овано|увано)$/iu.test(text)) return true;
    return /(?:^|[\s,;./(])скас(?:овано|увано)(?:$|[\s,;./)])/iu.test(text);
  });
};
