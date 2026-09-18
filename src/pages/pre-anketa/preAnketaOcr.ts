import type { AiQuestionnaireOcrField } from "../../api";
import {
  mapOcrFieldsToAnketaValues,
  type AnketaOcrProposal,
} from "../anketa-data/anketaQuestionnaireOcr";
import { normalizeAnketaPlaceText } from "../anketa-data/anketaPlaceFormat";
import { parseAnketaConscriptionText } from "../anketa-data/anketaConscription";
import {
  type PreAnketaFieldKey,
  type PreAnketaFormFields,
  PRE_ANKETA_FIELD_DEFS,
  preAnketaFieldLabel,
} from "./preAnketaFields";

export type PreAnketaOcrProposal = {
  field: PreAnketaFieldKey;
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

const pickField = (
  fields: AiQuestionnaireOcrField[],
  keys: string[],
): MappedValue | null => {
  const keySet = new Set(keys.map((key) => key.toLowerCase()));
  let best: MappedValue | null = null;
  for (const field of fields) {
    const key = String(field.key ?? "").trim().toLowerCase();
    if (!keySet.has(key)) continue;
    const value = String(field.value ?? "").trim();
    if (!value) continue;
    const candidate: MappedValue = {
      value,
      confidence: String(field.confidence ?? "medium"),
      source: String((field as { source?: string }).source ?? "").trim(),
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

const joinParts = (parts: Array<string | undefined>) =>
  parts
    .map((part) => String(part ?? "").trim())
    .filter(Boolean)
    .join(" ");

const normalizePhone = (value: string) => {
  const digits = value.replace(/\D/g, "");
  if (!digits) return "";
  if (digits.startsWith("380")) return `+${digits}`;
  if (digits.length === 10) return `+38${digits}`;
  if (digits.length === 9) return `+380${digits}`;
  return value.trim();
};

const splitPhones = (fields: AiQuestionnaireOcrField[]) => {
  const direct = pickField(fields, ["phones", "phone", "contactPhone"]);
  if (!direct) return { phone1: "", phone2: "" };
  const chunks = direct.value
    .split(/[\n;,/]+/)
    .map((part) => normalizePhone(part))
    .filter(Boolean);
  return {
    phone1: chunks[0] ?? "",
    phone2: chunks[1] ?? "",
  };
};

const buildPassportLine = (fields: AiQuestionnaireOcrField[]) => {
  const direct = pickField(fields, ["passport", "passportInfo"]);
  if (direct?.value.trim()) return direct.value.trim();

  const anketa = mapOcrFieldsToAnketaValues(fields);
  const name = anketa.idDocumentName?.value.trim() ?? "";
  const number = anketa.idDocumentNumber?.value.trim() ?? "";
  const issue = pickField(fields, [
    "passportIssuedBy",
    "passportIssueDate",
    "passportIssue",
  ])?.value.trim();
  return joinParts([
    name && number ? `${name} ${number}` : name || number,
    issue,
  ]);
};

const buildConscriptionLine = (fields: AiQuestionnaireOcrField[]) => {
  const direct = pickField(fields, [
    "conscription",
    "mobilizedBy",
    "rtcc",
    "conscriptedOffice",
  ]);
  const when = pickField(fields, [
    "conscriptedWhen",
    "mobilizedWhen",
    "conscriptedDate",
  ]);
  const combined = joinParts([direct?.value, when?.value]);
  if (combined) return combined;
  const anketa = mapOcrFieldsToAnketaValues(fields);
  return joinParts([anketa.conscriptedBy?.value, anketa.conscriptedWhen?.value]);
};

const buildServiceFlags = (fields: AiQuestionnaireOcrField[]) => {
  const service = pickField(fields, ["serviceType", "mobilized", "contract"]);
  const text = String(service?.value ?? "").toLowerCase();
  return {
    serviceContract: /контракт/i.test(text) ? "1" : "",
    serviceMobilized: /моб/i.test(text) ? "1" : "",
  };
};

const looksLikeFullNameLabel = (label: string) =>
  /п\.?\s*і\.?\s*б\.?/i.test(label) || /^піб$/i.test(label.trim());

const buildFullNameFromOcr = (
  fields: AiQuestionnaireOcrField[],
): MappedValue | null => {
  const direct = pickField(fields, [
    "fullName",
    "full_name",
    "name",
    "pib",
    "fio",
    "personName",
  ]);
  if (direct?.value.trim()) return direct;

  const surname = pickField(fields, ["surname", "lastName", "familyName"]);
  const firstName = pickField(fields, ["firstName", "givenName"]);
  const patronymic = pickField(fields, ["patronymic", "middleName", "fatherName"]);
  const joined = joinParts([surname?.value, firstName?.value, patronymic?.value]);
  if (joined) {
    return {
      value: joined,
      confidence:
        surname?.confidence ??
        firstName?.confidence ??
        patronymic?.confidence ??
        "medium",
      source:
        [surname?.source, firstName?.source, patronymic?.source].find(Boolean) ??
        undefined,
    };
  }

  for (const field of fields) {
    const label = String(field.label ?? "").trim();
    if (!looksLikeFullNameLabel(label)) continue;
    const value = String(field.value ?? "").trim();
    if (!value) continue;
    return {
      value,
      confidence: String(field.confidence ?? "medium"),
      source:
        String((field as { source?: string }).source ?? "").trim() || "паспорт",
    };
  }

  return null;
};

export const mapOcrFieldsToPreAnketaValues = (
  fields: AiQuestionnaireOcrField[],
): Partial<Record<PreAnketaFieldKey, MappedValue>> => {
  const anketa = mapOcrFieldsToAnketaValues(fields);
  const mapped: Partial<Record<PreAnketaFieldKey, MappedValue>> = {};
  const set = (field: PreAnketaFieldKey, value: MappedValue | null) => {
    if (!value?.value.trim()) return;
    mapped[field] = value;
  };

  set("fullName", buildFullNameFromOcr(fields));

  set("callsign", pickField(fields, ["callsign", "callSign", "nickname"]));
  set("rank", anketa.rank ?? pickField(fields, ["rank"]));
  set("rnokpp", anketa.rnokpp ?? pickField(fields, ["rnokpp", "taxId"]));
  set("iban", pickField(fields, ["iban", "bankAccount"]));
  set("driverLicense", pickField(fields, ["driverLicense", "driverCategory"]));
  set("birthDate", anketa.birthDate ?? pickField(fields, ["birthDate"]));
  set(
    "birthPlace",
    anketa.birthPlace
      ? {
          ...anketa.birthPlace,
          value: normalizeAnketaPlaceText(anketa.birthPlace.value),
        }
      : pickField(fields, ["birthPlace"]),
  );

  const registration = pickField(fields, ["registrationAddress", "address"]);
  const actual = pickField(fields, ["actualAddress", "residenceAddress"]);
  set(
    "registrationAddress",
    registration
      ? {
          ...registration,
          value: normalizeAnketaPlaceText(registration.value),
        }
      : null,
  );
  set(
    "actualAddress",
    actual
      ? { ...actual, value: normalizeAnketaPlaceText(actual.value) }
      : anketa.location
        ? {
            ...anketa.location,
            value: normalizeAnketaPlaceText(anketa.location.value),
          }
        : registration
          ? {
              ...registration,
              value: normalizeAnketaPlaceText(registration.value),
            }
          : null,
  );

  const phones = splitPhones(fields);
  if (phones.phone1) {
    set("phone1", {
      value: phones.phone1,
      confidence: "medium",
      source: "контакти",
    });
  }
  if (phones.phone2) {
    set("phone2", {
      value: phones.phone2,
      confidence: "medium",
      source: "контакти",
    });
  }

  set("education", anketa.education ?? pickField(fields, ["education"]));
  set("work", pickField(fields, ["work", "employment", "workplace"]));
  set("arrivedFromUnit", pickField(fields, ["arrivedFromUnit", "previousUnit"]));
  set(
    "servedBefore2022",
    pickField(fields, ["servedBefore2022", "servedBeforeWar"]),
  );

  const passport = buildPassportLine(fields);
  if (passport) {
    set("passport", {
      value: passport,
      confidence: "medium",
      source: pickField(fields, ["passport"])?.source ?? "паспорт",
    });
  }

  const conscription = buildConscriptionLine(fields);
  if (conscription) {
    set("conscription", {
      value: parseAnketaConscriptionText(conscription).by
        ? conscription
        : conscription,
      confidence: "medium",
      source: "РТЦК",
    });
  }

  const ubd = pickField(fields, ["ubd", "ubdCertificate", "combatParticipantId"]);
  if (ubd) set("ubdCertificate", ubd);

  set("familyStatus", pickField(fields, ["familyStatus", "maritalStatus"]));
  set("mother", pickField(fields, ["mother", "relativeMother"]));
  set("father", pickField(fields, ["father", "relativeFather"]));
  set(
    "trustedPerson",
    pickField(fields, ["trustedPerson", "relativeTrustedPerson"]),
  );
  set("children", pickField(fields, ["children", "child", "relativeChild"]));
  set("sports", pickField(fields, ["sports", "sport"]));
  set("sportsRank", pickField(fields, ["sportsRank", "sportRank"]));

  const flags = buildServiceFlags(fields);
  if (flags.serviceContract) {
    mapped.serviceContract = {
      value: "1",
      confidence: "medium",
      source: "анкета",
    };
  }
  if (flags.serviceMobilized) {
    mapped.serviceMobilized = {
      value: "1",
      confidence: "medium",
      source: "анкета",
    };
  }

  return mapped;
};

export const buildPreAnketaOcrProposals = (
  fields: AiQuestionnaireOcrField[],
  current: PreAnketaFormFields,
): PreAnketaOcrProposal[] => {
  const mapped = mapOcrFieldsToPreAnketaValues(fields);
  const proposals: PreAnketaOcrProposal[] = [];

  for (const def of PRE_ANKETA_FIELD_DEFS) {
    const hit = mapped[def.key];
    if (!hit?.value.trim()) continue;
    const currentValue = String(current[def.key] ?? "").trim();
    if (currentValue && currentValue === hit.value.trim()) continue;
    proposals.push({
      field: def.key,
      label: preAnketaFieldLabel(def.key),
      value: hit.value.trim(),
      confidence: hit.confidence,
      selected: !currentValue,
      source: hit.source,
    });
  }

  return proposals;
};

export const applyPreAnketaOcrProposals = (
  current: PreAnketaFormFields,
  proposals: PreAnketaOcrProposal[],
) => {
  const next = { ...current };
  for (const proposal of proposals) {
    if (!proposal.selected) continue;
    next[proposal.field] = proposal.value;
  }
  return next;
};

export const buildPreAnketaOcrFieldHints = () =>
  PRE_ANKETA_FIELD_DEFS.map((field) => ({
    key: field.key,
    header: field.label,
  }));

/** Експорт для тестів: сумісність з anketa OCR proposals. */
export const toAnketaStyleProposals = (
  proposals: PreAnketaOcrProposal[],
): AnketaOcrProposal[] =>
  proposals.map((proposal) => ({
    columnId: proposal.field as never,
    label: proposal.label,
    value: proposal.value,
    confidence: proposal.confidence,
    selected: proposal.selected,
    source: proposal.source,
  }));
