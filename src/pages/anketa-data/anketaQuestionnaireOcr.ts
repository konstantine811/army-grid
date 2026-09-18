import type { AiQuestionnaireOcrField } from "../../api";
import {
  ANKETA_ABSENT_QUESTIONNAIRE_VALUE,
  ANKETA_TEXTAREA_EMPTY_TEMPLATES,
  isAnketaAbsentQuestionnaireValue,
  isAnketaGapCellEmpty,
  isAnketaGapCellTrulyEmpty,
  isReplaceableAnketaMissingValue,
} from "./anketaGaps";
import {
  looksLikeConscriptionOffice,
  parseAnketaConscriptionText,
} from "./anketaConscription";
import { normalizeAnketaRelativesText } from "./anketaRelativesFormat";
import { normalizeAnketaPlaceText } from "./anketaPlaceFormat";
import { padAnketaIdCardDocumentNumber } from "./anketaIdDocumentNumber";
import { extractMilitaryIdFromText } from "../personnel/vkTpvDovidkyImport";
import {
  ANKETA_COLUMNS,
  isAnketaColumnReadonly,
  type AnketaColumnKey,
  type AnketaRow,
} from "./anketaSheet";

export type AnketaOcrProposal = {
  columnId: AnketaColumnKey;
  label: string;
  value: string;
  confidence: string;
  selected: boolean;
  source?: string;
};

type MappedValue = {
  value: string;
  confidence: string;
  source?: string;
};

const confidenceRank = (value: string) => {
  const text = String(value ?? "").trim().toLowerCase();
  if (text === "high") return 3;
  if (text === "medium") return 2;
  if (text === "low") return 1;
  return 2;
};

const pickOcrField = (
  fields: AiQuestionnaireOcrField[],
  key: string,
): MappedValue | null => {
  let best: MappedValue | null = null;
  for (const field of fields) {
    if (String(field.key ?? "").trim() !== key) continue;
    const value = String(field.value ?? "").trim();
    if (!value) continue;
    const candidate = {
      value,
      confidence: String(field.confidence ?? "medium"),
      source: pageTypeLabel(
        String((field as { source?: string }).source ?? "").trim() ||
          inferSourceFromPageType(fields),
      ),
    };
    if (
      !best ||
      confidenceRank(candidate.confidence) > confidenceRank(best.confidence)
    ) {
      best = candidate;
    }
  }
  return best;
};

const pickOcrFieldByKeys = (
  fields: AiQuestionnaireOcrField[],
  keys: string[],
): MappedValue | null => {
  for (const key of keys) {
    const hit = pickOcrField(fields, key);
    if (hit) return hit;
  }
  return null;
};

const inferSourceFromPageType = (fields: AiQuestionnaireOcrField[]) => {
  const pageType = String(
    fields.find((field) => field.key === "pageType")?.value ?? "",
  )
    .trim()
    .toLowerCase();
  return pageTypeLabel(pageType);
};

const pageTypeLabel = (pageType: string) => {
  switch (pageType) {
    case "passport":
      return "паспорт";
    case "id_card":
      return "ID-картка";
    case "military_id":
      return "військовий квиток";
    case "tax_id":
      return "РНОКПП";
    case "form":
      return "анкета";
    case "mixed":
      return "скани документів";
    default:
      return pageType || "";
  }
};

const joinNameParts = (parts: string[]) =>
  parts
    .map((part) => part.trim())
    .filter(Boolean)
    .join(" ");

const formatRelativeFieldLine = (prefix: string, value: string) => {
  const trimmed = value.trim();
  if (!trimmed) return prefix;
  const label = prefix.replace(/:$/, "").trim();
  if (trimmed.toLowerCase().startsWith(label.toLowerCase())) return trimmed;
  return `${prefix} ${trimmed}`;
};

const confidenceFromRank = (rank: number) =>
  rank >= 3 ? "high" : rank >= 2 ? "medium" : "low";

/** Збирає блок «Дані про родичів» з повного поля або частин + телефонів. */
export const buildRelativesFromOcrFields = (
  fields: AiQuestionnaireOcrField[],
): MappedValue | null => {
  const direct = pickOcrField(fields, "relatives");

  const lineDefs = [
    { prefix: "Сімейний стан:", keys: ["familyStatus"] },
    { prefix: "Мати:", keys: ["relativeMother", "mother"] },
    { prefix: "Батько:", keys: ["relativeFather", "father"] },
    { prefix: "Довірена особа:", keys: ["relativeTrustedPerson", "trustedPerson"] },
    { prefix: "Дитина:", keys: ["relativeChild", "child", "children"] },
  ] as const;

  const picked: MappedValue[] = [];
  const lines = lineDefs.map(({ prefix, keys }) => {
    const field = pickOcrFieldByKeys(fields, [...keys]);
    if (field?.value.trim()) {
      picked.push(field);
      return formatRelativeFieldLine(prefix, field.value);
    }
    return prefix;
  });

  const relativePhones = pickOcrFieldByKeys(fields, [
    "relativePhones",
    "relativesPhones",
    "familyPhones",
  ]);
  if (relativePhones?.value.trim()) {
    picked.push(relativePhones);
    const phones = relativePhones.value.trim();
    lines.push(
      /^тел/i.test(phones) ? phones : `тел. родичів: ${phones}`,
    );
  }

  if (direct?.value.trim()) picked.push(direct);
  if (!picked.length) return null;

  const rawValue = direct?.value.trim()
    ? direct.value
    : lines.join("\n");
  const value = normalizeAnketaRelativesText(rawValue);
  if (!value.trim()) return null;

  const confidenceRankValue = picked.reduce(
    (best, field) => Math.max(best, confidenceRank(field.confidence)),
    1,
  );

  return {
    value,
    confidence: confidenceFromRank(confidenceRankValue),
    source: picked.find((field) => field.source)?.source ?? "анкета",
  };
};

const UBD_ABSENT_MARKERS =
  /^(немає|відсутн\w*|не\s+має|ні|—|-+|нет|no|нема)$/i;

const normalizeAdditionalInfoPhoneLine = (value: string) => {
  const trimmed = String(value ?? "").trim();
  if (!trimmed) return "тел:";
  const body = trimmed.replace(/^тел[.:]?\s*/i, "").trim();
  if (!body) return "тел:";
  if (/^тел:/i.test(trimmed)) return `тел: ${body}`;
  return `тел: ${body}`;
};

const normalizeAdditionalInfoUbdLine = (value: string) => {
  const trimmed = String(value ?? "").trim();
  if (!trimmed) return "УБД: немає";
  const body = trimmed.replace(/^убд[.:]?\s*/i, "").trim();
  if (!body || UBD_ABSENT_MARKERS.test(body)) return "УБД: немає";
  if (/^убд:/i.test(trimmed)) return `УБД: ${body}`;
  return `УБД: ${body}`;
};

/** Порядок: адреса (якщо є), «УБД: …», «тел: …» в кінці. */
export const formatAdditionalInfoBlock = (parts: {
  phone?: string;
  ubd?: string;
  place?: string;
}) =>
  [
    String(parts.place ?? "").trim(),
    normalizeAdditionalInfoUbdLine(parts.ubd ?? ""),
    normalizeAdditionalInfoPhoneLine(parts.phone ?? ""),
  ]
    .filter((line, index) => index > 0 || Boolean(line))
    .join("\n");

/** Збирає «Додаткова інформація» з телефону службовця та номера УБД. */
export const buildAdditionalInfoFromOcrFields = (
  fields: AiQuestionnaireOcrField[],
): MappedValue | null => {
  const direct = pickOcrField(fields, "additionalInfo");
  const phones = pickOcrFieldByKeys(fields, [
    "phones",
    "phone",
    "mobilePhone",
    "contactPhone",
  ]);
  const ubd = pickOcrFieldByKeys(fields, [
    "ubd",
    "ubdNumber",
    "ubdCertificate",
    "ubdCard",
    "ubdDocument",
  ]);

  const picked = [direct, phones, ubd].filter((field): field is MappedValue =>
    Boolean(field?.value.trim()),
  );
  if (!picked.length) return null;

  let phonePart = phones?.value ?? "";
  let ubdPart = ubd?.value ?? "";
  let placePart = "";

  if (direct?.value.trim()) {
    const leftoverLines: string[] = [];
    for (const line of direct.value.split(/\r?\n/)) {
      const trimmed = line.trim();
      if (!trimmed) continue;
      if (/^тел/i.test(trimmed)) {
        phonePart = trimmed;
        continue;
      }
      if (/^убд/i.test(trimmed)) {
        ubdPart = trimmed;
        continue;
      }
      leftoverLines.push(trimmed);
    }
    if (!phonePart) {
      const phoneMatch = direct.value.match(
        /(?:тел[.:]?\s*)?((?:\+?38[\s-]*)?0(?:[\s-]*\d){9})/i,
      );
      if (phoneMatch?.[1]) phonePart = phoneMatch[1].replace(/[\s-]+/g, "");
    }
    placePart = leftoverLines.join(" ");
  }

  const normalizedPlace = normalizeAnketaPlaceText(placePart, {
    keepPhone: false,
  });
  const value = formatAdditionalInfoBlock({
    phone: phonePart,
    ubd: ubdPart,
    place: normalizedPlace,
  });
  const hasPhone = /\d{9,}/.test(phonePart);
  const ubdBody = ubdPart.replace(/^убд[.:]?\s*/i, "").trim();
  const hasUbd = Boolean(ubdBody) && !UBD_ABSENT_MARKERS.test(ubdBody);
  if (!hasPhone && !hasUbd && !normalizedPlace) return null;

  const confidenceRankValue = picked.reduce(
    (best, field) => Math.max(best, confidenceRank(field.confidence)),
    1,
  );

  return {
    value,
    confidence: confidenceFromRank(confidenceRankValue),
    source: picked.find((field) => field.source)?.source ?? "анкета",
  };
};

const normalizeRnokpp = (value: string) => {
  const digits = value.replace(/\D/g, "");
  return digits.length === 10 ? digits : value.trim();
};

const isValidRnokppDigits = (value: string) =>
  value.replace(/\D/g, "").length === 10;

const rnokppPriority = (field: AiQuestionnaireOcrField) => {
  const key = String(field.key ?? "").trim();
  const source = String((field as { source?: string }).source ?? "")
    .trim()
    .toLowerCase();
  if (source === "рнокпп" || source === "tax_id") {
    if (key === "taxId") return 50;
    if (key === "rnokpp") return 48;
    if (key === "ipn" || key === "inn") return 46;
  }
  if (source === "анкета" || source === "form") {
    if (key === "rnokpp") return 40;
    if (key === "taxId") return 38;
  }
  if (source === "скани документів" || source === "mixed") {
    if (key === "taxId" || key === "rnokpp") return 30;
  }
  if (key === "taxId") return 20;
  if (key === "rnokpp") return 18;
  if (key === "ipn" || key === "inn") return 16;
  return 10;
};

/** РНОКПП лише з 10 цифр; пріоритет — довідка ІПН, потім анкета, не паспорт. */
const pickRnokppField = (
  fields: AiQuestionnaireOcrField[],
  passportDigits?: string,
): MappedValue | null => {
  let best: { value: MappedValue; rank: number } | null = null;
  for (const field of fields) {
    const key = String(field.key ?? "").trim();
    if (!["rnokpp", "taxId", "inn", "ipn"].includes(key)) continue;
    const normalized = normalizeRnokpp(String(field.value ?? ""));
    if (!isValidRnokppDigits(normalized)) continue;
    const digits = normalized.replace(/\D/g, "");
    if (passportDigits && digits === passportDigits) continue;
    const candidate: MappedValue = {
      value: digits,
      confidence: String(field.confidence ?? "medium"),
      source:
        String((field as { source?: string }).source ?? "").trim() ||
        pageTypeLabel("tax_id"),
    };
    const rank =
      rnokppPriority(field) * 10 + confidenceRank(candidate.confidence);
    if (!best || rank > best.rank) {
      best = { value: candidate, rank };
    }
  }
  return best?.value ?? null;
};

const combineSeriesNumber = (series?: string, number?: string) => {
  const left = String(series ?? "").trim();
  const right = String(number ?? "").trim();
  if (left && right) return `${left} ${right}`.replace(/\s+/g, " ").trim();
  return left || right;
};

const normalizeOcrMilitaryId = (value: string) => {
  const text = String(value ?? "").replace(/\s+/g, " ").trim();
  if (!text) return "";
  const extracted = extractMilitaryIdFromText(text);
  if (extracted) return extracted;
  const generic = text.match(
    /([А-ЯІЇЄҐA-Z]{2})\s*[№#]?\s*(\d{5,8})/iu,
  );
  if (generic) {
    return `${generic[1]!.toLocaleUpperCase("uk-UA")} ${generic[2]}`;
  }
  return text;
};

const documentNumberDigits = (value: string) =>
  String(value ?? "").replace(/\D/g, "");

/** Код органу — окремі 4 цифри, не початок 9-значного номера ID-картки. */
const isInvalidOrganForDocument = (organ: string, documentNumber?: string) => {
  if (!/^\d{4}$/.test(organ)) return true;
  const docDigits = documentNumberDigits(documentNumber ?? "");
  if (docDigits.length >= 9 && docDigits.startsWith(organ)) return true;
  return false;
};

const normalizeIssuingOrgan = (value: string, documentNumber?: string) => {
  const trimmed = String(value ?? "").trim();
  if (!trimmed) return "";

  if (/(орган)/i.test(trimmed)) {
    const organ =
      trimmed.match(/(\d{4})\s*\(\s*орган\s*\)/i)?.[1] ??
      trimmed.match(/(\d{4})/)?.[1] ??
      "";
    return isInvalidOrganForDocument(organ, documentNumber) ? "" : organ;
  }

  if (/^\d{4}$/.test(trimmed)) {
    return isInvalidOrganForDocument(trimmed, documentNumber) ? "" : trimmed;
  }

  return "";
};

/** Дата видачі + код органу для колонки «Назва документа». */
export const formatIdDocumentIssuingMeta = (parts: {
  issuedDate?: string;
  issuingOrgan?: string;
  documentNumber?: string;
}) => {
  const date = String(parts.issuedDate ?? "")
    .trim()
    .replace(/^виданий\s+/i, "");
  const organ = normalizeIssuingOrgan(
    String(parts.issuingOrgan ?? ""),
    parts.documentNumber,
  );

  const segments: string[] = [];
  if (date) segments.push(`виданий ${date}`);
  if (organ) segments.push(organ);
  return segments.join(", ");
};

export const formatIdDocumentNameWithIssuing = (
  documentType: string,
  issuingMeta: string,
) => {
  const type = normalizeIdDocumentName(documentType);
  const meta = issuingMeta.trim();
  if (type && meta) return `${type}, ${meta}`;
  return type || meta;
};

const extractIssuedDateFromText = (value: string) =>
  value.match(/виданий\s+(\d{2}\.\d{2}\.\d{4})/i)?.[1] ??
  value.match(/\b(\d{2}\.\d{2}\.\d{4})\b/)?.[1] ??
  "";

const extractIssuingOrganFromText = (
  value: string,
  documentNumber?: string,
) => {
  const text = String(value ?? "").trim();
  if (!text) return "";

  const labeled = text.match(/(\d{4})\s*\(\s*орган\s*\)/i)?.[1];
  if (labeled && !isInvalidOrganForDocument(labeled, documentNumber)) {
    return labeled;
  }

  const afterIssued = text.match(
    /виданий\s+\d{2}\.\d{2}\.\d{4}\s*,\s*(\d{4})\b/i,
  )?.[1];
  if (afterIssued && !isInvalidOrganForDocument(afterIssued, documentNumber)) {
    return afterIssued;
  }

  const trailing = text.match(/,\s*(\d{4})\s*$/i)?.[1];
  if (trailing && !isInvalidOrganForDocument(trailing, documentNumber)) {
    return trailing;
  }

  return "";
};

const extractDocumentNumberFromText = (value: string) => {
  const trimmed = String(value ?? "").trim();
  if (!trimmed) return "";
  const beforeIssued = trimmed.split(/,\s*виданий/i)[0]?.trim();
  if (beforeIssued) return beforeIssued.split(",")[0]?.trim() ?? beforeIssued;
  return trimmed.split(",")[0]?.trim() ?? trimmed;
};

const pickIdDocumentIssuingFields = (fields: AiQuestionnaireOcrField[]) => ({
  issuedDate: pickOcrFieldByKeys(fields, [
    "passportIssuedDate",
    "idDocumentIssuedDate",
    "documentIssuedDate",
  ]),
  issuingOrgan: pickOcrFieldByKeys(fields, [
    "passportIssuingOrgan",
    "idDocumentIssuingOrgan",
    "issuingOrgan",
    "passportIssuedBy",
  ]),
});

const buildIssuingMeta = (
  fields: AiQuestionnaireOcrField[],
  combinedNumberValue?: string,
) => {
  const { issuedDate, issuingOrgan } = pickIdDocumentIssuingFields(fields);
  const documentNumber = combinedNumberValue
    ? extractDocumentNumberFromText(combinedNumberValue)
    : "";
  const date =
    issuedDate?.value ||
    (combinedNumberValue
      ? extractIssuedDateFromText(combinedNumberValue)
      : "");
  const organ =
    issuingOrgan?.value ||
    (combinedNumberValue
      ? extractIssuingOrganFromText(combinedNumberValue, documentNumber)
      : "");
  return formatIdDocumentIssuingMeta({
    issuedDate: date,
    issuingOrgan: organ,
    documentNumber,
  });
};

const mergeIdDocumentNumberValue = (directValue: string) => {
  const trimmed = String(directValue ?? "").trim();
  if (!trimmed) return "";
  return padAnketaIdCardDocumentNumber(
    extractDocumentNumberFromText(trimmed) || trimmed,
  );
};

const buildIdDocumentNumber = (fields: AiQuestionnaireOcrField[]) => {
  const direct = pickOcrField(fields, "idDocumentNumber");
  const passportCombined = combineSeriesNumber(
    pickOcrField(fields, "passportSeries")?.value,
    pickOcrField(fields, "passportNumber")?.value,
  );
  const passportDirect = pickOcrField(fields, "passport");

  if (direct?.value.trim()) {
    return {
      value: mergeIdDocumentNumberValue(direct.value),
      confidence: direct.confidence,
      source: direct.source,
    };
  }

  const seriesNumber = passportCombined || passportDirect?.value || "";
  if (!seriesNumber.trim()) return null;

  return {
    value: mergeIdDocumentNumberValue(seriesNumber),
    confidence:
      passportDirect?.confidence ??
      pickOcrField(fields, "passportSeries")?.confidence ??
      pickOcrField(fields, "passportNumber")?.confidence ??
      "medium",
    source:
      passportDirect?.source ??
      pickOcrField(fields, "passportSeries")?.source,
  };
};

const inferIdDocumentType = (
  fields: AiQuestionnaireOcrField[],
  documentNumber: string,
) => {
  const pageType = String(
    fields.find((field) => field.key === "pageType")?.value ?? "",
  )
    .trim()
    .toLowerCase();

  if (pageType === "id_card") return "ID-картка";

  const numberPart = extractDocumentNumberFromText(documentNumber);
  const digits = numberPart.replace(/\D/g, "");
  const hasLetterSeries = /[A-Za-zА-Яа-яІіЇїЄєҐґ]{2}/u.test(numberPart);
  if (digits.length === 9 && !hasLetterSeries) return "ID-картка";

  const source = inferSourceFromPageType(fields);
  if (source === "ID-картка") return "ID-картка";
  if (source === "паспорт") return "Паспорт";
  if (documentNumber.trim()) return "Паспорт";
  return "";
};

const buildIdDocumentName = (
  fields: AiQuestionnaireOcrField[],
  documentNumber: string,
) => {
  const direct = pickOcrField(fields, "idDocumentNumber");
  const directName = pickOcrField(fields, "idDocumentName");
  const rawName = directName?.value.trim() ?? "";

  let issuingMeta = buildIssuingMeta(fields, direct?.value);
  if (!issuingMeta && rawName) {
    issuingMeta = buildIssuingMeta(fields, rawName);
  }

  let documentType = "";
  if (rawName) {
    if (/виданий/i.test(rawName)) {
      documentType = normalizeIdDocumentName(
        rawName.split(/,\s*виданий/i)[0]?.trim() ?? "",
      );
    } else {
      documentType = normalizeIdDocumentName(rawName);
    }
  }

  if (!documentType) {
    documentType = inferIdDocumentType(fields, documentNumber);
  }

  const value = formatIdDocumentNameWithIssuing(documentType, issuingMeta);
  if (!value) return null;

  const { issuedDate, issuingOrgan } = pickIdDocumentIssuingFields(fields);
  return {
    value,
    confidence:
      directName?.confidence ??
      issuedDate?.confidence ??
      issuingOrgan?.confidence ??
      "medium",
    source:
      directName?.source ??
      issuedDate?.source ??
      issuingOrgan?.source ??
      inferSourceFromPageType(fields),
  };
};

const normalizeSex = (value: string) => {
  const text = value.trim().toLowerCase();
  if (!text) return "";
  if (/^(ч|чол|чолов|m|male)/i.test(text)) return "ч";
  if (/^(ж|жін|жіно|f|female)/i.test(text)) return "ж";
  return value.trim();
};

const normalizeIdDocumentName = (value: string) => {
  const text = value.trim();
  if (!text) return "";
  if (/id[-\s]?карт/i.test(text)) return "ID-картка";
  if (/паспорт/i.test(text)) return "Паспорт";
  if (/посвідч/i.test(text)) return "Посвідчення";
  if (/військов/i.test(text) && /квит/i.test(text)) return "Військовий квиток";
  return text;
};

const normalizeServiceType = (value: string) => {
  const text = value.trim();
  if (!text) return "";
  const lower = text.toLocaleLowerCase("uk-UA");
  if (/мобіл/.test(lower) && !/контракт/.test(lower)) return "мобілізований";
  if (/контракт/.test(lower) && !/мобіл/.test(lower)) return "контракт";
  return text;
};

const withNormalizedPlace = (
  value: MappedValue | null,
): MappedValue | null => {
  if (!value?.value.trim()) return null;
  const next = normalizeAnketaPlaceText(value.value);
  if (!next) return null;
  return next === value.value ? value : { ...value, value: next };
};

/** Зіставляє поля OCR (анкета + скани документів) з колонками анкети. */
export const mapOcrFieldsToAnketaValues = (
  fields: AiQuestionnaireOcrField[],
): Partial<Record<AnketaColumnKey, MappedValue>> => {
  const mapped: Partial<Record<AnketaColumnKey, MappedValue>> = {};
  const set = (columnId: AnketaColumnKey, value: MappedValue | null) => {
    if (!value?.value.trim()) return;
    const existing = mapped[columnId];
    if (
      !existing ||
      confidenceRank(value.confidence) > confidenceRank(existing.confidence)
    ) {
      mapped[columnId] = value;
    }
  };

  set("rank", pickOcrField(fields, "rank"));

  const passportCombined = combineSeriesNumber(
    pickOcrField(fields, "passportSeries")?.value,
    pickOcrField(fields, "passportNumber")?.value,
  );
  const passportDirect = pickOcrField(fields, "passport");
  const idDocumentNumber = buildIdDocumentNumber(fields);
  const passportDigits = (
    passportCombined ||
    passportDirect?.value ||
    extractDocumentNumberFromText(idDocumentNumber?.value ?? "")
  )
    .replace(/\D/g, "")
    .slice(-10);

  set("rnokpp", pickRnokppField(fields, passportDigits || undefined));

  set("birthDate", pickOcrField(fields, "birthDate"));
  set("birthPlace", withNormalizedPlace(pickOcrField(fields, "birthPlace")));
  set("education", pickOcrField(fields, "education"));

  const militaryCombined = normalizeOcrMilitaryId(
    combineSeriesNumber(
      pickOcrField(fields, "militaryIdSeries")?.value,
      pickOcrField(fields, "militaryIdNumber")?.value,
    ),
  );
  const militaryDirect = pickOcrField(fields, "militaryId");
  const militaryDirectValue = militaryDirect
    ? normalizeOcrMilitaryId(militaryDirect.value)
    : "";
  set(
    "militaryId",
    militaryCombined
      ? {
          value: militaryCombined,
          confidence:
            militaryDirect?.confidence ??
            pickOcrField(fields, "militaryIdSeries")?.confidence ??
            pickOcrField(fields, "militaryIdNumber")?.confidence ??
            "medium",
          source:
            militaryDirect?.source ??
            pickOcrField(fields, "militaryIdSeries")?.source,
        }
      : militaryDirectValue
        ? { ...militaryDirect!, value: militaryDirectValue }
        : null,
  );

  set("idDocumentNumber", idDocumentNumber);

  const idDocumentName = buildIdDocumentName(
    fields,
    idDocumentNumber?.value ?? passportCombined ?? passportDirect?.value ?? "",
  );
  if (idDocumentName) set("idDocumentName", idDocumentName);

  const registration = pickOcrField(fields, "registrationAddress");
  const actual = pickOcrField(fields, "actualAddress");
  set("location", withNormalizedPlace(actual || registration || null));

  const arrived = pickOcrField(fields, "work");
  if (arrived && !looksLikeConscriptionOffice(arrived.value)) {
    set("arrivedFrom", withNormalizedPlace(arrived));
  }

  const conscriptionSources = [
    pickOcrFieldByKeys(fields, [
      "conscriptedBy",
      "conscription",
      "mobilizedBy",
      "rtcc",
      "conscriptedOffice",
    ]),
    pickOcrFieldByKeys(fields, [
      "conscriptedWhen",
      "mobilizedWhen",
      "conscriptedDate",
    ]),
    looksLikeConscriptionOffice(arrived?.value ?? "") ? arrived : null,
    ...fields
      .filter((field) =>
        looksLikeConscriptionOffice(String(field.value ?? "")),
      )
      .map((field) => ({
        value: String(field.value ?? "").trim(),
        confidence: String(field.confidence ?? "medium"),
        source: pageTypeLabel(
          String((field as { source?: string }).source ?? "").trim() ||
            inferSourceFromPageType(fields),
        ),
      })),
  ].filter((field): field is MappedValue => Boolean(field?.value.trim()));
  if (conscriptionSources.length) {
    const parsed = parseAnketaConscriptionText(
      [...new Set(conscriptionSources.map((field) => field.value.trim()))].join(
        " ",
      ),
    );
    const confidence = conscriptionSources.reduce(
      (best, field) =>
        Math.max(best, confidenceRank(field.confidence)),
      1,
    );
    const source =
      conscriptionSources.find((field) => field.source)?.source ?? "анкета";
    if (parsed.when) {
      set("conscriptedWhen", {
        value: parsed.when,
        confidence: confidenceFromRank(confidence),
        source,
      });
    }
    if (parsed.by) {
      set("conscriptedBy", {
        value: parsed.by,
        confidence: confidenceFromRank(confidence),
        source,
      });
    }
    if (parsed.militaryId && !mapped.militaryId) {
      const fromConscription = normalizeOcrMilitaryId(parsed.militaryId);
      if (fromConscription) {
        set("militaryId", {
          value: fromConscription,
          confidence: confidenceFromRank(confidence),
          source,
        });
      }
    }
  }

  const serviceType = pickOcrFieldByKeys(fields, [
    "serviceType",
    "mobilized",
    "contract",
  ]);
  if (serviceType) {
    set("serviceType", {
      ...serviceType,
      value: normalizeServiceType(serviceType.value),
    });
  }

  set("additionalInfo", buildAdditionalInfoFromOcrFields(fields));

  set("relatives", buildRelativesFromOcrFields(fields));

  const sexValue = pickOcrField(fields, "sex");
  if (sexValue) {
    set("sex", {
      ...sexValue,
      value: normalizeSex(sexValue.value),
    });
  }

  const surname = pickOcrField(fields, "surname");
  const firstName = pickOcrField(fields, "firstName");
  const patronymic = pickOcrField(fields, "patronymic");
  const fullName = joinNameParts([
    surname?.value ?? "",
    firstName?.value ?? "",
    patronymic?.value ?? "",
  ]);
  if (fullName) {
    void fullName;
  }

  for (const column of ANKETA_COLUMNS) {
    const columnId = column.key;
    if (isAnketaColumnReadonly(columnId)) continue;
    if (mapped[columnId]) continue;
    const direct = pickOcrField(fields, columnId);
    if (!direct) continue;
    if (
      columnId === "birthPlace" ||
      columnId === "location" ||
      columnId === "arrivedFrom"
    ) {
      set(columnId, withNormalizedPlace(direct));
    } else {
      set(columnId, direct);
    }
  }

  return mapped;
};

export const buildAnketaGapColumnDescriptors = (
  gapColumnKeys: readonly AnketaColumnKey[],
) => {
  const headerByKey = new Map(
    ANKETA_COLUMNS.map((column) => [column.key, column.header]),
  );
  return gapColumnKeys
    .filter((key) => !isAnketaColumnReadonly(key))
    .map((key) => ({
      key,
      header: headerByKey.get(key) ?? key,
    }));
};

export const buildAnketaEmptyGapColumnDescriptors = (
  row: AnketaRow,
  gapColumnKeys: readonly AnketaColumnKey[],
) =>
  buildAnketaGapColumnDescriptors(gapColumnKeys).flatMap((column) => {
    const columnId = column.key as AnketaColumnKey;
    const trulyEmpty = isAnketaGapCellTrulyEmpty(row, columnId);
    const absentPlaceholder = isAnketaAbsentQuestionnaireValue(row[columnId]);
    if (!trulyEmpty && !absentPlaceholder) return [];
    const currentValue = String(row[columnId] ?? "").trim();
    return currentValue ? [{ ...column, currentValue }] : [column];
  });

const collapseHintText = (value: string) => value.replace(/\s+/g, " ").trim();

const isAnketaOcrHintTemplateValue = (
  columnId: AnketaColumnKey,
  value: string,
) => {
  const template = ANKETA_TEXTAREA_EMPTY_TEMPLATES[columnId];
  if (!template) return false;
  return collapseHintText(value) === collapseHintText(template);
};

/** ПІБ + уже відомі поля, щоб модель звірила почерк і не вигадала чужі дані. */
export const buildAnketaOcrPersonHint = (row: AnketaRow) => {
  const name = String(row.fullName ?? "").trim();
  const known = ANKETA_COLUMNS.filter((column) => column.key !== "fullName")
    .flatMap((column) => {
      const value = String(row[column.key] ?? "").trim();
      if (!value) return [];
      if (isAnketaGapCellEmpty(row, column.key)) return [];
      if (isAnketaOcrHintTemplateValue(column.key, value)) return [];
      return [`${column.header}: ${collapseHintText(value).slice(0, 220)}`];
    });
  return [
    name,
    known.length ? `Відомі дані (звір почерк): ${known.join("; ")}` : "",
    "Адреси пиши так: с./м./смт Назва, Район р-н, Область обл; телефон в кінці.",
  ]
    .filter(Boolean)
    .join(". ")
    .slice(0, 1800);
};

export const buildAnketaOcrRequestHints = (
  row: AnketaRow,
  gapColumnKeys: readonly AnketaColumnKey[],
) => ({
  personName: buildAnketaOcrPersonHint(row),
  selectedColumns: buildAnketaGapColumnDescriptors(gapColumnKeys).map(
    (column) => {
      const currentValue = String(
        row[column.key as AnketaColumnKey] ?? "",
      ).trim();
      return currentValue ? { ...column, currentValue } : column;
    },
  ),
  emptyColumns: buildAnketaEmptyGapColumnDescriptors(row, gapColumnKeys),
});

export const buildAnketaOcrProposals = (
  ocrFields: AiQuestionnaireOcrField[],
  row: AnketaRow,
  gapColumnKeys: readonly AnketaColumnKey[],
  options?: { focusedColumnId?: AnketaColumnKey },
): AnketaOcrProposal[] => {
  const mapped = mapOcrFieldsToAnketaValues(ocrFields);
  const gapSet = new Set(gapColumnKeys);
  const headerByKey = new Map(
    ANKETA_COLUMNS.map((column) => [column.key, column.header]),
  );
  const proposals: AnketaOcrProposal[] = [];

  for (const column of ANKETA_COLUMNS) {
    const columnId = column.key;
    if (isAnketaColumnReadonly(columnId)) continue;
    if (!gapSet.has(columnId)) continue;
    const trulyEmpty = isAnketaGapCellTrulyEmpty(row, columnId);
    const absentPlaceholder = isAnketaAbsentQuestionnaireValue(row[columnId]);
    if (!trulyEmpty && !absentPlaceholder) continue;

    const candidate = mapped[columnId];
    const rawValue = String(candidate?.value ?? "").trim();
    if (
      absentPlaceholder &&
      (!rawValue || isReplaceableAnketaMissingValue(rawValue))
    ) {
      continue;
    }
    const isAbsentDefault = trulyEmpty && !rawValue;
    const value = rawValue || ANKETA_ABSENT_QUESTIONNAIRE_VALUE;
    const isFocused = options?.focusedColumnId === columnId;
    proposals.push({
      columnId,
      label: headerByKey.get(columnId) ?? columnId,
      value,
      confidence: isAbsentDefault
        ? "medium"
        : candidate?.confidence ?? "low",
      selected:
        isFocused ||
        (Boolean(rawValue) && candidate?.confidence !== "low"),
      source: isAbsentDefault ? "за замовчуванням" : candidate?.source,
    });
  }

  const focusedId = options?.focusedColumnId;
  if (focusedId) {
    proposals.sort((left, right) => {
      if (left.columnId === focusedId) return -1;
      if (right.columnId === focusedId) return 1;
      return 0;
    });
  }

  return proposals;
};
