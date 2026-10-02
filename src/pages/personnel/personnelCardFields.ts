import { valueToDisplay } from "../../excelRoundTrip";
import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import { getFighterStatusDirectValue } from "./fighterStatusImport";
import { ROSTER_FIELD_PREFIX } from "./personnelRosterMerge";
import {
  formatPersonDisplayName,
  inferRosterFieldLabel,
  isRosterNoteFieldLabel,
  looksLikePersonBirthDate,
} from "./personnelUtils";

export type PersonnelRosterCardField = {
  key: string;
  sourceKey: string;
  label: string;
  value: string;
  isPibField: boolean;
};

const displayCell = (value: unknown) =>
  valueToDisplay(value as Parameters<typeof valueToDisplay>[0]).trim();

const labelNormOf = (label: string) =>
  label.trim().toLocaleLowerCase("uk-UA").replace(/_/g, " ");

const isGenericColumnSource = (sourceKey: string) =>
  /^(?:column|колонка)_\d+(?:_\d+)?$/iu.test(sourceKey.trim());

/** «ж» у Штатці — позначка «те саме, що в рядку вище», не назва підрозділу. */
const isDittoValue = (value: string) =>
  /^(?:ж|те\s*ж|теж|[-—–]|\.+)$/iu.test(value.trim());

/** Штатка кладе і назву колонки (`підрозділ`), і `column_2` з тим самим текстом. */
const dedupeRosterFields = (fields: PersonnelRosterCardField[]) => {
  const indexBySignature = new Map<string, number>();
  const result: PersonnelRosterCardField[] = [];
  for (const field of fields) {
    const signature = `${labelNormOf(field.label)}\0${field.value}`;
    const existingIndex = indexBySignature.get(signature);
    if (existingIndex == null) {
      indexBySignature.set(signature, result.length);
      result.push(field);
      continue;
    }
    const existing = result[existingIndex];
    if (
      existing &&
      isGenericColumnSource(existing.sourceKey) &&
      !isGenericColumnSource(field.sourceKey)
    ) {
      result[existingIndex] = field;
    }
  }
  return result;
};

/** Поля «Загального списку» на картці: дублі шапки (ПІБ, дата, ІПН, посада) ховаються. */
export const buildVisibleRosterFieldRows = ({
  row,
  rosterLabels,
  personName,
  birthDate,
  cardRnokpp,
}: {
  row: EjournalPreviewRow | null;
  rosterLabels: Record<string, string>;
  personName: string;
  birthDate: string;
  cardRnokpp: string;
}): PersonnelRosterCardField[] => {
  const cardName = formatPersonDisplayName(personName);
  const cardRnokppDigits = String(cardRnokpp ?? "").replace(/\D/g, "");
  const birth = String(birthDate ?? "").trim();

  const fields = dedupeRosterFields(
    Object.entries(row ?? {})
    .filter(
      ([key, value]) =>
        key.startsWith(ROSTER_FIELD_PREFIX) &&
        !key.includes("fighter_status_") &&
        displayCell(value),
    )
    .map(([key, value]) => {
      const sourceKey = key.slice(ROSTER_FIELD_PREFIX.length);
      const displayed = displayCell(value);
      const label = inferRosterFieldLabel(sourceKey, displayed, rosterLabels);
      const labelNorm = labelNormOf(label);
      const isPibField =
        labelNorm === "піб" ||
        labelNorm === "прізвище" ||
        labelNorm.includes("піб") ||
        /(^|_)(піб|прізвище|column_14)(_|$)/i.test(sourceKey);
      return {
        key,
        sourceKey,
        label,
        value: isPibField ? formatPersonDisplayName(displayed) : displayed,
        isPibField,
      };
    })
    .filter((field) => {
      const labelNorm = labelNormOf(field.label);
      const keyNorm = field.sourceKey.toLocaleLowerCase("uk-UA");
      const isYearField =
        labelNorm === "рік" ||
        labelNorm === "рік народження" ||
        keyNorm === "рік" ||
        keyNorm === "rik" ||
        /(^|_)(рік|year|column_17)(_|$)/i.test(field.sourceKey);
      const isBirthDateField =
        labelNorm === "дата народження" ||
        labelNorm.includes("дата народ") ||
        (keyNorm.includes("народ") &&
          (keyNorm.includes("дата") || keyNorm.includes("день"))) ||
        /(^|_)(column_16)(_|$)/i.test(field.sourceKey);
      const isFullYearsField =
        labelNorm === "повних років" ||
        labelNorm.includes("повних років") ||
        /(^|_)(column_18)(_|$)/i.test(field.sourceKey);
      const isStayPlaceField =
        labelNorm === "місце перебування" ||
        labelNorm.includes("перебуван") ||
        labelNorm === "дислокація" ||
        labelNorm.includes("дислокац") ||
        /(^|_)(column_31|column_40|column_43|колонка_43)(_|$)/iu.test(
          field.sourceKey,
        );
      const isPositionField =
        labelNorm === "посада" ||
        labelNorm === "повна посада" ||
        (labelNorm.includes("посада") &&
          !labelNorm.includes("індекс") &&
          !labelNorm.includes("прийняття")) ||
        /(^|_)(column_5|column_7)(_|$)/i.test(field.sourceKey);
      const isRosterStatusField =
        labelNorm === "статус" ||
        /(^|_)(column_21|column_37)(_|$)/i.test(field.sourceKey);
      const isIpnField =
        labelNorm === "іпн" ||
        labelNorm === "рнокпп" ||
        /(^|_)(column_19)(_|$)/i.test(field.sourceKey);
      const isQuestionnaireFlagField =
        labelNorm === "анкета" ||
        /(^|_)(column_10)(_|$)/i.test(field.sourceKey);
      const isRankField =
        labelNorm === "звання" ||
        /(^|_)(column_13)(_|$)/i.test(field.sourceKey);
      const isCallSignField =
        labelNorm === "позивний" ||
        /(^|_)(column_15)(_|$)/i.test(field.sourceKey);

      if (
        field.isPibField &&
        cardName &&
        formatPersonDisplayName(field.value) === cardName
      ) {
        return false;
      }
      if (birth && (isYearField || isBirthDateField || isFullYearsField)) {
        return false;
      }
      if (!birth && isYearField && looksLikePersonBirthDate(field.value)) {
        return false;
      }
      if (isStayPlaceField || isPositionField || isRosterStatusField) {
        return false;
      }
      if (isQuestionnaireFlagField || isRankField || isCallSignField) {
        return false;
      }
      if (
        isIpnField &&
        cardRnokppDigits.length >= 8 &&
        field.value.replace(/\D/g, "") === cardRnokppDigits
      ) {
        return false;
      }
      if (isRosterNoteFieldLabel(field.label)) return false;
      if (isDittoValue(field.value)) return false;
      return true;
    }),
  );

  const allDays = getFighterStatusDirectValue(row, "fighter_status_all_days");
  if (allDays) {
    const totalDaysField: PersonnelRosterCardField = {
      key: "roster__fighter_status_all_days",
      sourceKey: "fighter_status_all_days",
      label: "Усього днів",
      value: allDays,
      isPibField: false,
    };
    const exitCountIndex = fields.findIndex((field) => {
      const key = field.sourceKey.replace(/-/g, "").toLocaleLowerCase("uk-UA");
      return key.includes("ксть") && key.includes("виход");
    });
    if (exitCountIndex >= 0) fields.splice(exitCountIndex + 1, 0, totalDaysField);
    else fields.unshift(totalDaysField);
  }

  return fields;
};

export const resolvePersonnelPersonNote = (
  row: EjournalPreviewRow | null,
  rosterLabels: Record<string, string>,
) => {
  const fighterNote = getFighterStatusDirectValue(row, "fighter_status_note");
  if (fighterNote) return fighterNote;
  for (const [key, value] of Object.entries(row ?? {})) {
    if (!key.startsWith(ROSTER_FIELD_PREFIX)) continue;
    const displayed = displayCell(value);
    if (!displayed) continue;
    const sourceKey = key.slice(ROSTER_FIELD_PREFIX.length);
    const label = inferRosterFieldLabel(sourceKey, displayed, rosterLabels);
    if (isRosterNoteFieldLabel(label)) return displayed;
  }
  return "";
};
