/**
 * Правила сфери руху 1ПБ. Нові кадрові фічі мають йти через
 * `classifyStaffMove` / `isOwnUnitStaffMove` / `isOutboundStaffMove`,
 * а не через окремі перевірки «Куди» в плані.
 */
export type MovementRuleEvent = {
  type: string;
  rank: string;
  destination: string;
  changeText: string;
  status: string;
  note: string;
  previousIndex?: string;
  nextIndex?: string;
};

const normText = (value: string) =>
  String(value ?? "")
    .replace(/\s+/g, " ")
    .trim();

/** «Куди» у Рух часто = «x» / порожньо; тоді військова частина є в примітці. */
export const resolveMovementDestination = (rawDest: string, note: string) => {
  const clean = (value: string) => {
    const text = normText(value);
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
  const noteText = normText(note);
  return /(?:[АA]\s*\d{4}(?!\d)|військов(?:ої|а)\s+частин)/iu.test(noteText)
    ? noteText
    : "";
};

export const formatTransferDestinationForTimesheet = (value: string) => {
  const codes = [...String(value ?? "").matchAll(/[АA]\s*(\d{4})(?!\d)/giu)].map(
    (match) => `А${match[1]}`,
  );
  const unique = [...new Set(codes)];
  return unique.length ? unique.join(" / ") : value;
};

/**
 * «2103791 Старший кухар → 2103179 Стрілець 3 піхотного відділення …» —
 * для «куди вибув» потрібна лише нова посада без службового індексу.
 */
export const positionChangeDestination = (event: Pick<MovementRuleEvent, "changeText" | "destination">) => {
  const text = normText(event.changeText);
  if (!text) return normText(event.destination);
  const tail = text.split(/→|->|=>/).pop()?.trim() || text;
  return tail.replace(/^\d{5,}[\s.:;-]*/, "").trim() || tail;
};

/**
 * AE «Куди вибув / документи» — повна посада з «Яка зміна» (Рух),
 * а не код частини з колонки «Куди».
 */
export const documentsDestFromMovement = (event: Pick<MovementRuleEvent, "changeText" | "destination">) => {
  const text = normText(event.changeText);
  if (!text) return "";
  return positionChangeDestination(event) || text;
};

/**
 * «Куди» вказує на сам 1ПБ, а не на іншу частину.
 * У Рух буває `_5 1ПБ`, `5 1ПБ` і просто `1 ПБ` / `1ПБ`.
 */
export const isOwnFirstPbDestination = (destination: string) => {
  const text = String(destination ?? "")
    .toUpperCase()
    .replace(/\s+/g, " ")
    .trim();
  if (!text) return false;
  return /(?:^|[_\s])(?:5\s+)?1\s*ПБ(?:$|[_\s.,;/])/.test(` ${text} `);
};

export const isInternalFirstPbMovement = (
  event: Pick<MovementRuleEvent, "destination">,
) => isOwnFirstPbDestination(event.destination);

/**
 * Зміна посади в межах 1ПБ. Якщо «Куди» вказує інший підрозділ, це вибуття з
 * частини, а не внутрішній перехід, і закривати посаду треба через виключення.
 * Прибуття в 1ПБ (`АРТ ДИВ → 1 ПБ`) — постановка на штат, не виключення.
 */
const isStaffIndexToken = (value: string | undefined) =>
  /^\d{5,}$/.test(normText(value ?? ""));

/** ПОСАДА з розпорядження на штатний індекс 1ПБ, без «Куди» іншої частини. */
export const isDispositionToStaffPlacement = (
  event: Pick<MovementRuleEvent, "previousIndex" | "nextIndex" | "changeText" | "destination">,
) => {
  if (!isStaffIndexToken(event.nextIndex)) return false;
  if (isStaffIndexToken(event.previousIndex)) return false;
  if (normText(event.destination) && !isOwnFirstPbDestination(event.destination)) {
    return false;
  }
  return /розпорядж/iu.test(
    `${event.previousIndex} ${event.changeText}`,
  );
};

/**
 * Інша частина в тексті Рух. Лише рівень батальйону / полку:
 * «штурмове відділення» всередині 1ПБ сюди не входить.
 */
export const mentionsForeignUnit = (value: string) =>
  /(?:^|[_\s])(?:\d+\s*)ШБ\b|(?:\d+\s+)?штурмов(?:ий|ого|ому|им)?\s+батальйон|(?:^|[_\s])(?:2|3|4|5)\s*ПБ\b|(?:^|[_\s])(?:\d+\s*)(?:РБ|ТБ)\b/iu.test(
    ` ${normText(value).toUpperCase()} `,
  );

/** Зовнішня військова частина: А7400 / в/ч / «військова частина». */
export const mentionsExternalMilitaryUnit = (value: string) =>
  /(?:[АA]\s*-?\s*\d{4}(?!\d)|в\s*\/\s*ч|військов(?:ої|а|у)\s+частин)/iu.test(
    normText(value),
  );

/**
 * Для ПЕРЕВ: якщо «Куди» лишилось 1ПБ, а в примітці в/ч А#### —
 * фактичний напрямок беремо з примітки.
 */
export const resolveOutboundTransferDestination = (
  rawDest: string,
  note: string,
) => {
  const dest = resolveMovementDestination(rawDest, note);
  if (mentionsExternalMilitaryUnit(note) && isOwnFirstPbDestination(dest)) {
    return normText(note);
  }
  return dest;
};

export type StaffMoveScope = "own" | "outbound" | "other";

type StaffMoveEvent = Pick<
  MovementRuleEvent,
  "type" | "destination" | "changeText" | "previousIndex" | "nextIndex" | "note"
>;

/**
 * Єдине правило: ПОСАДА і ПЕРЕВ дивляться на ту саму сферу.
 * 1) «Куди» = 1ПБ → свій; інша частина → вибуття.
 * 2) Порожнє «Куди»: чужий батальйон у тексті → вибуття; інакше штат 1ПБ.
 * Для ПЕРЕВ зовнішню в/ч у примітці дивиться `classifyStaffMove` —
 * статус РУХ («В СТРОЮ») сюди не входить.
 */
export const isFirstPbPositionChange = (
  event: Pick<
    MovementRuleEvent,
    "destination" | "changeText" | "previousIndex" | "nextIndex"
  >,
) => {
  if (isOwnFirstPbDestination(event.destination)) return true;
  if (normText(event.destination) && !isOwnFirstPbDestination(event.destination)) {
    return false;
  }
  if (mentionsForeignUnit(event.changeText)) return false;
  if (
    isStaffIndexToken(event.previousIndex) &&
    isStaffIndexToken(event.nextIndex)
  ) {
    return true;
  }
  if (isDispositionToStaffPlacement(event)) return true;
  return /(?:^|[^\d])1\s*(?:ПБ|піхотн(?:ий|ого|ому|им)?\s+батальйон)/iu.test(
    normText(event.changeText).toUpperCase(),
  );
};

export const classifyStaffMove = (event: StaffMoveEvent): StaffMoveScope => {
  if (event.type !== "ПОСАДА" && event.type !== "ПЕРЕВ") return "other";
  // ПЕРЕВ + в/ч А#### у «Куди» / примітці / «Яка зміна» — вибуття з 1ПБ,
  // навіть коли «Куди» ще 1ПБ, а статус лишився «В СТРОЮ».
  if (
    event.type === "ПЕРЕВ" &&
    mentionsExternalMilitaryUnit(
      [event.note, event.destination, event.changeText].join(" "),
    )
  ) {
    return "outbound";
  }
  return isFirstPbPositionChange(event) ? "own" : "outbound";
};

export const isOwnUnitStaffMove = (event: StaffMoveEvent) =>
  classifyStaffMove(event) === "own";

export const isOutboundStaffMove = (event: StaffMoveEvent) =>
  classifyStaffMove(event) === "outbound";

export const findLatestPriorOwnUnitStaffMove = <
  T extends StaffMoveEvent & {
    orderDate?: string;
    basisDate?: string;
    excelRow: number;
  },
>(
  events: T[],
  event: T,
  opts: {
    samePerson: (left: T, right: T) => boolean;
    inWindow: (candidate: T) => boolean;
    eventTime: (item: T) => number;
  },
): T | null => {
  let found: T | null = null;
  const limit = opts.eventTime(event);
  for (const candidate of events) {
    if (!opts.samePerson(event, candidate)) continue;
    if (!opts.inWindow(candidate)) continue;
    if (opts.eventTime(candidate) >= limit) continue;
    if (!isOwnUnitStaffMove(candidate)) continue;
    const candidateTime = opts.eventTime(candidate);
    const foundTime = found ? opts.eventTime(found) : 0;
    if (
      !found ||
      candidateTime > foundTime ||
      (candidateTime === foundTime && candidate.excelRow > found.excelRow)
    ) {
      found = candidate;
    }
  }
  return found;
};

const normalizeRankLabel = (value: string) =>
  normText(value)
    .replace(/^(?:звання|військов(?:е|ое)\s+звання)\s+/i, "")
    .toLocaleLowerCase("uk-UA");

/** «солдат → СТАРШИЙ СОЛДАТ» або лише нове звання в «Яка зміна». */
export const parseRankPromotion = (
  event: Pick<MovementRuleEvent, "changeText" | "note" | "rank">,
) => {
  const text = [event.changeText, event.note].filter(Boolean).join(" ");
  const match = text.match(
    /([А-ЯІЇЄҐа-яіїєґ'’.\-\s]{3,}?)\s*(?:→|->|=>|—)\s*([А-ЯІЇЄҐа-яіїєґ'’.\-\s]{3,})/u,
  );
  if (match) {
    return {
      previousRank: normalizeRankLabel(match[1]),
      nextRank: normalizeRankLabel(match[2]),
    };
  }
  const nextRank =
    normalizeRankLabel(event.changeText) || normalizeRankLabel(event.rank);
  const previousRank = normalizeRankLabel(event.rank);
  return {
    previousRank: previousRank === nextRank ? "" : previousRank,
    nextRank,
  };
};

export const isRankAssignmentEvent = (event: Pick<MovementRuleEvent, "type">) =>
  event.type === "ЗВАННЯ" || event.type.startsWith("ЗВАН");

export const unitCodeFromMovement = (
  event: Pick<MovementRuleEvent, "destination" | "changeText" | "note">,
) =>
  formatTransferDestinationForTimesheet(
    [event.destination, event.changeText, event.note].join(" "),
  );

export const isSzchCancellation = (
  event: Pick<MovementRuleEvent, "type" | "status" | "note" | "changeText">,
) =>
  /СКАС(?:УВАННЯ|ОВАНО|УВАТИ)?.*СЗЧ|СЗЧ.*СКАС/iu.test(
    [event.type, event.status, event.note, event.changeText].join(" "),
  );

export const isDispositionAbsenceStatus = (value: string) =>
  /СЗЧ|САМОВІЛ|БЕЗВІСТ|(?:^|[^А-ЯІЇЄҐ])ЗБ(?:$|[^А-ЯІЇЄҐ])/iu.test(value);

export const isCurrentUnitStatusMarker = (value: string) =>
  /^_?\s*\d+\s+\d+\s*ПБ$/iu.test(value.replace(/\s+/g, " ").trim());

export const movementBlob = (event: MovementRuleEvent) =>
  [event.type, event.status, event.note, event.changeText, event.destination].join(
    " ",
  );

/** Окремий рядок «СКАСУВАННЯ переведення», а не скасований рядок ПЕРЕВ. */
export const isTransferCancellation = (event: MovementRuleEvent) => {
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
export const isCancelledMovementRecord = (event: MovementRuleEvent) => {
  if (isSzchCancellation(event) || isTransferCancellation(event)) return false;
  const fields = [event.status, event.note, event.destination, event.changeText];
  return fields.some((value) => {
    const text = String(value ?? "").trim();
    if (!text) return false;
    if (/^скас(?:овано|увано)$/iu.test(text)) return true;
    return /(?:^|[\s,;./(])скас(?:овано|увано)(?:$|[\s,;./)])/iu.test(text);
  });
};
