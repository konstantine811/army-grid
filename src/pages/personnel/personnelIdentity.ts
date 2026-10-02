import { api } from "../../api";
import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import {
  collectPersonAttachmentLookupIds,
  rememberPersonnelIdentityLinks,
} from "./personAttachments";
import { getRosterPersonRnokpp } from "./personnelRosterMerge";
import {
  getPersonDisplayName,
  getPersonFieldValue,
  isLikelyPersonnelRow,
  resolvePersonBirthDate,
  resolvePersonIdentityKey,
} from "./personnelUtils";

const compactIpn = (value: unknown) => {
  const digits = String(value ?? "").replace(/\D/g, "");
  return digits.length >= 8 && digits.length <= 12 ? digits : "";
};

const ipnFromRow = (row: EjournalPreviewRow) =>
  getRosterPersonRnokpp(row) ||
  compactIpn(getPersonFieldValue(row, ["рнокпп_за_наявності"])) ||
  compactIpn(getPersonFieldValue(row, ["рнокпп"])) ||
  compactIpn(getPersonFieldValue(row, ["іпн"]));

/** Знаходить або створює постійний pid і повертає старі ключі фото й анкет. */
export const loadPersonnelIdentityLinks = async (
  rows: EjournalPreviewRow[],
  attachments: Array<{ personExternalId?: string | null; fileName?: string | null }>,
  signal?: AbortSignal,
) => {
  const people = rows.flatMap((row) => {
    if (!isLikelyPersonnelRow(row)) return [];
    const clientKey = resolvePersonIdentityKey(row);
    const fullName = getPersonDisplayName(row);
    if (!clientKey || !fullName) return [];
    return [
      {
        clientKey,
        fullName,
        birthDate: resolvePersonBirthDate(row),
        ipn: ipnFromRow(row),
        aliases: collectPersonAttachmentLookupIds(row).slice(0, 80),
      },
    ];
  });
  if (!people.length) return [];

  const resolved = await api.resolvePersonnelIdentities(
    {
      people,
      attachments: attachments.flatMap((item) => {
        const personExternalId = String(item.personExternalId ?? "").trim();
        if (!personExternalId) return [];
        return [
          {
            personExternalId,
            fileName: String(item.fileName ?? "").trim(),
          },
        ];
      }),
    },
    signal,
  );
  rememberPersonnelIdentityLinks(resolved.people);
  return resolved.people;
};
