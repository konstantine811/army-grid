import type { BackendPersonDocument } from "../../api";
import {
  upsertPersonPhonesDocument,
  uniqueNormalizedPhones,
  writeStoredPersonPhones,
} from "./personPhonesStore";
import { formatUaPhoneDisplay } from "./personnelUtils";

export const phoneDraftProblem = (
  hasPerson: boolean,
  normalized: string | null,
) => {
  if (!hasPerson) return "Спочатку виберіть особу зі списку.";
  if (!normalized) return "Вкажіть український номер, наприклад 063 123 45 67.";
  return "";
};

export const phoneDuplicateMessage = (phone: string) =>
  `Цей номер уже збережено: ${formatUaPhoneDisplay(phone)}.`;

export const phoneSavedMessage = (phone: string, savedToDb: boolean) =>
  savedToDb
    ? `Телефон збережено: ${formatUaPhoneDisplay(phone)}. Номер не затреться при оновленні списку.`
    : `Телефон збережено локально: ${formatUaPhoneDisplay(phone)}. Не вдалося записати в БД.`;

export const phoneRemovedMessage = (phone: string, savedToDb: boolean) =>
  savedToDb
    ? `Телефон видалено: ${formatUaPhoneDisplay(phone)}.`
    : `Телефон прибрано локально: ${formatUaPhoneDisplay(phone)}. Не вдалося оновити БД.`;

export const phoneClipboardText = (phone: string) => formatUaPhoneDisplay(phone);

export const phoneCopiedMessage = (phone: string) =>
  `Скопійовано: ${phoneClipboardText(phone)}`;

export const phoneCopyFailedMessage = () => "Не вдалося скопіювати номер.";

export const withSavedPersonPhones = (
  phonesByExternalId: Record<string, string[]>,
  externalId: string,
  phones: string[],
) => {
  const updated = {
    ...phonesByExternalId,
    [externalId]: uniqueNormalizedPhones(phones),
  };
  writeStoredPersonPhones(updated);
  return updated;
};

export const savePersonPhonesDocument = async (
  externalId: string,
  phones: string[],
  document: BackendPersonDocument | null,
) => {
  try {
    const saved = await upsertPersonPhonesDocument(
      externalId,
      phones,
      document,
    );
    return { ok: true as const, document: saved };
  } catch {
    return { ok: false as const };
  }
};
