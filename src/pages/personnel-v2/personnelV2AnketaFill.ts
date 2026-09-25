import {
  applyAnketaEditsToRows,
  loadAnketaEdits,
} from "../anketa-data/anketaEdits";
import {
  buildAnketaRowLookup,
  matchAnketaRowForStaffList,
  type AnketaRowLookup,
} from "../staff-list-gaps/staffListGapsAnketaMatch";
import {
  loadAnketaSheetPreferCache,
  type AnketaColumnKey,
  type AnketaRow,
} from "../anketa-data/anketaSheet";
import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import {
  buildPersonSummary,
  extractPhones,
  formatPersonFieldValue,
  formatUaPhoneDisplay,
  resolvePersonFieldKey,
  type PersonFieldDef,
} from "../personnel/personnelUtils";

const ANKETA_VALUE_BY_LABEL: Record<string, AnketaColumnKey> = {
  Звання: "rank",
  ID: "externalId",
  "Індекс посади": "positionIndex",
  "Дата народження": "birthDate",
  "Місце народження": "birthPlace",
  Стать: "sex",
  РНОКПП: "rnokpp",
  "Відмова від РНОКПП": "rnokppRefuse",
  "Документ, що посвідчує особу": "idDocumentName",
  "Серія/номер документа": "idDocumentNumber",
  "Військовий квиток": "militaryId",
  "Вид служби": "serviceType",
  "Місце перебування": "location",
  "Звідки прибув": "arrivedFrom",
  "Дата укладання контракту": "contractFrom",
  "Дата закінчення контракту / період призову": "contractTo",
  "Коли призваний (прийнятий)": "conscriptedWhen",
  "Ким призваний (прийнятий)": "conscriptedBy",
  Освіта: "education",
  "Дати прийняття посади": "positionDates",
  "Номер наказу на прийняття посади": "positionOrderNumber",
  "Дата зарахування до списків": "enlistDate",
  "Наказ про зарахування — дата": "enlistOrderDate",
  "Наказ про зарахування — номер": "enlistOrderNumber",
  "Наказ на призначення — дата": "appointmentOrderDate",
  "Наказ на призначення — номер": "appointmentOrderNumber",
  "Вхідна дата наказу на призначення": "appointmentInDate",
  "Вхідний номер наказу на призначення": "appointmentInNumber",
  "Наказ на звання — дата": "rankOrderDate",
  "Наказ на звання — номер": "rankOrderNumber",
  "Вхідна дата наказу на звання": "rankInDate",
  "Вхідний номер наказу на звання": "rankInNumber",
  "Дані про родичів": "relatives",
  "Додаткова інформація": "additionalInfo",
};

export type PersonnelV2AnketaIndex = {
  lookup: AnketaRowLookup;
  byExternalId: Map<string, AnketaRow>;
};

const text = (value: unknown) => String(value ?? "").trim();

const pick = (current: string, next: string) => current.trim() || next.trim();

export const loadPersonnelV2AnketaRows = async () => {
  const [snapshot, edits] = await Promise.all([
    loadAnketaSheetPreferCache({ refreshGoogle: false }).catch(() => null),
    loadAnketaEdits().catch(() => ({})),
  ]);
  return applyAnketaEditsToRows(snapshot?.rows ?? [], edits);
};

export const indexPersonnelV2AnketaRows = (
  rows: AnketaRow[],
): PersonnelV2AnketaIndex => {
  const byExternalId = new Map<string, AnketaRow>();
  for (const row of rows) {
    const id = text(row.externalId);
    if (!id || byExternalId.has(id)) continue;
    byExternalId.set(id, row);
  }
  return { lookup: buildAnketaRowLookup(rows), byExternalId };
};

export const matchPersonnelV2AnketaRow = (
  index: PersonnelV2AnketaIndex,
  name: string,
  birthDate = "",
  externalId = "",
) => {
  const id = text(externalId);
  if (id) {
    const byId = index.byExternalId.get(id);
    if (byId) return byId;
  }
  return matchAnketaRowForStaffList(index.lookup, name, birthDate);
};

export const anketaCardFieldValue = (anketa: AnketaRow | null, label: string) => {
  if (!anketa) return "";
  const key = ANKETA_VALUE_BY_LABEL[label];
  if (!key) return "";
  return text(anketa[key]);
};

export const personCardFieldValue = (
  row: EjournalPreviewRow | null,
  anketa: AnketaRow | null,
  field: PersonFieldDef,
) => {
  const key = row ? resolvePersonFieldKey(row, field.parts) : "";
  const fromOos =
    key && row ? formatPersonFieldValue(row[key], field).trim() : "";
  if (fromOos) return { key, value: fromOos, fromAnketa: false };
  const fromAnketa = anketaCardFieldValue(anketa, field.label);
  if (!fromAnketa) return null;
  return { key: `anketa:${field.label}`, value: fromAnketa, fromAnketa: true };
};

/** Картка показує анкету, а не порожні чи часткові поля ООС. */
export const summaryPreferringAnketa = (
  summary: ReturnType<typeof buildPersonSummary> | null,
  anketa: AnketaRow | null,
) => {
  if (!anketa) return summary;
  const fromAnketa = mergePersonSummaryWithAnketa(null, anketa);
  if (!fromAnketa) return summary;
  if (!summary) return fromAnketa;
  return {
    ...summary,
    ...fromAnketa,
    name: summary.name || fromAnketa.name,
    callSign: summary.callSign,
    positionTitle: summary.positionTitle,
    phones: fromAnketa.phones.length ? fromAnketa.phones : summary.phones,
    phonesDisplay: fromAnketa.phones.length
      ? fromAnketa.phonesDisplay
      : summary.phonesDisplay,
  };
};

export const mergePersonSummaryWithAnketa = (
  summary: ReturnType<typeof buildPersonSummary> | null,
  anketa: AnketaRow | null,
) => {
  if (!summary && !anketa) return null;
  const additionalInfo = pick(
    summary?.additionalInfo ?? "",
    anketa?.additionalInfo ?? "",
  );
  const phones = summary?.phones.length
    ? summary.phones
    : extractPhones(anketa?.additionalInfo ?? "");
  return {
    name: summary?.name || text(anketa?.fullName),
    rank: pick(summary?.rank ?? "", anketa?.rank ?? ""),
    externalId: pick(summary?.externalId ?? "", anketa?.externalId ?? ""),
    positionIndex: pick(summary?.positionIndex ?? "", anketa?.positionIndex ?? ""),
    serviceType: pick(summary?.serviceType ?? "", anketa?.serviceType ?? ""),
    birthDate: pick(summary?.birthDate ?? "", anketa?.birthDate ?? ""),
    birthPlace: pick(summary?.birthPlace ?? "", anketa?.birthPlace ?? ""),
    sex: pick(summary?.sex ?? "", anketa?.sex ?? ""),
    rnokpp: pick(summary?.rnokpp ?? "", anketa?.rnokpp ?? ""),
    location: pick(summary?.location ?? "", anketa?.location ?? ""),
    positionTitle: summary?.positionTitle ?? "",
    arrivedFrom: pick(summary?.arrivedFrom ?? "", anketa?.arrivedFrom ?? ""),
    education: pick(summary?.education ?? "", anketa?.education ?? ""),
    relatives: pick(summary?.relatives ?? "", anketa?.relatives ?? ""),
    additionalInfo,
    phones,
    phonesDisplay: summary?.phones.length
      ? summary.phonesDisplay
      : phones.map(formatUaPhoneDisplay),
    militaryId: pick(summary?.militaryId ?? "", anketa?.militaryId ?? ""),
    contractFrom: pick(summary?.contractFrom ?? "", anketa?.contractFrom ?? ""),
    contractTo: pick(summary?.contractTo ?? "", anketa?.contractTo ?? ""),
    callSign: summary?.callSign ?? "",
  };
};
