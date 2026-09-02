/** Cheap stable fingerprint for large API payloads. Avoids JSON.stringify of 800+ rows. */

const mix = (hash: number, part: string) => {
  let next = hash;
  for (let index = 0; index < part.length; index += 1) {
    next = (Math.imul(next, 33) ^ part.charCodeAt(index)) >>> 0;
  }
  return next;
};

const rowToken = (row: unknown) => {
  if (!row || typeof row !== "object") return String(row ?? "");
  const item = row as Record<string, unknown>;
  const values = item.values && typeof item.values === "object" && !Array.isArray(item.values)
    ? (item.values as Record<string, unknown>)
    : item;
  return [
    item.id ?? item.__dbRowId ?? item.externalId ?? "",
    item.excelRowNumber ?? item.__rowNumber ?? "",
    item.status ?? item.statusLabel ?? "",
    item.updatedAt ?? item.createdAt ?? "",
    item.name ?? values.піб ?? values.column_14 ?? "",
  ].join(":");
};

export const payloadFingerprint = (value: unknown): string => {
  if (value == null) return "null";
  if (typeof value !== "object") return `${typeof value}:${String(value)}`;

  if (Array.isArray(value)) {
    let hash = value.length >>> 0;
    for (const item of value) hash = mix(hash, rowToken(item));
    return `arr:${value.length}:${hash}`;
  }

  const record = value as Record<string, unknown>;
  const rows = Array.isArray(record.rows) ? record.rows : null;
  if (rows) {
    let hash = rows.length >>> 0;
    hash = mix(hash, String(record.importId ?? record.importName ?? ""));
    hash = mix(hash, String(record.createdAt ?? record.updatedAt ?? record.sourceFileName ?? ""));
    for (const row of rows) hash = mix(hash, rowToken(row));
    return `rows:${rows.length}:${hash}`;
  }

  if (Array.isArray(record.items)) {
    return payloadFingerprint(record.items);
  }

  let hash = Object.keys(record).length >>> 0;
  for (const key of [
    "id",
    "importId",
    "createdAt",
    "updatedAt",
    "sourceFileName",
    "status",
    "externalId",
  ]) {
    if (key in record) hash = mix(hash, `${key}:${String(record[key] ?? "")}`);
  }
  return `obj:${hash}`;
};

export const payloadChanged = (left: unknown, right: unknown) => {
  if (Object.is(left, right)) return false;
  try {
    return payloadFingerprint(left) !== payloadFingerprint(right);
  } catch {
    return true;
  }
};
