export type ParsedDiskFio = {
  surname: string;
  firstName: string;
  patronymic: string;
};

export const normalizeDiskNamePart = (value: string) =>
  String(value ?? "")
    .normalize("NFC")
    .toLowerCase()
    .replace(/ё/g, "е")
    .replace(/[''`ʼ´]/g, "")
    .replace(/[^a-zа-яіїєґ0-9\s-]/gi, " ")
    .replace(/-/g, " ")
    .replace(/\s+/g, " ")
    .trim();

export const getNormalizedDiskNameTokens = (value: string) => {
  const withoutExtension = String(value ?? "").replace(/\.pdf$/i, "");
  const withoutCallsign = withoutExtension.replace(/\([^)]*\)/g, " ");
  return normalizeDiskNamePart(withoutCallsign).split(" ").filter(Boolean);
};

export const parseDiskFio = (value: string): ParsedDiskFio => {
  const tokens = getNormalizedDiskNameTokens(value);
  return {
    surname: tokens[0] ?? "",
    firstName: tokens[1] ?? "",
    patronymic: tokens[2] ?? "",
  };
};

const softenUkrainian = (value: string) =>
  value.replace(/і/g, "и").replace(/ї/g, "и").replace(/є/g, "е").replace(/ґ/g, "г");

const editDistance = (left: string, right: string) => {
  if (left === right) return 0;
  const a = left.length <= right.length ? left : right;
  const b = left.length <= right.length ? right : left;
  const prev = Array.from({ length: a.length + 1 }, (_, index) => index);
  const next = new Array<number>(a.length + 1);
  for (let j = 1; j <= b.length; j += 1) {
    next[0] = j;
    for (let i = 1; i <= a.length; i += 1) {
      const cost = a[i - 1] === b[j - 1] ? 0 : 1;
      next[i] = Math.min(next[i - 1] + 1, prev[i] + 1, prev[i - 1] + cost);
    }
    for (let i = 0; i <= a.length; i += 1) prev[i] = next[i]!;
  }
  return prev[a.length]!;
};

/** Exact, spelling variant, or a small typo — not a different name. */
export const diskNamePartsClose = (left: string, right: string) => {
  if (!left || !right) return false;
  if (left === right) return true;
  const a = softenUkrainian(left);
  const b = softenUkrainian(right);
  if (a === b) return true;
  const shorter = a.length <= b.length ? a : b;
  const longer = a.length <= b.length ? b : a;
  if (shorter.length >= 4 && longer.startsWith(shorter)) return true;
  const distance = editDistance(a, b);
  if (distance <= 1) return true;
  return distance <= 2 && Math.min(a.length, b.length) >= 6;
};

const partsConflict = (left: string, right: string) =>
  Boolean(left && right && !diskNamePartsClose(left, right));

const alignFileFio = (person: ParsedDiskFio, file: ParsedDiskFio): ParsedDiskFio => {
  if (
    file.surname &&
    file.firstName &&
    diskNamePartsClose(person.surname, file.firstName) &&
    diskNamePartsClose(person.firstName, file.surname)
  ) {
    return {
      surname: file.firstName,
      firstName: file.surname,
      patronymic: file.patronymic,
    };
  }
  return file;
};

/**
 * Drop disk hits whose імʼя / по батькові / прізвище are a different person.
 * Keep one-part filenames (only surname or only given name) and small typos
 * so the operator can still accept them by hand.
 */
export const isPlausibleDiskQuestionnaireMatch = (
  fullName: string,
  fileName: string,
  callSign = "",
) => {
  const person = parseDiskFio(fullName);
  const tokens = getNormalizedDiskNameTokens(fileName);
  if (!tokens.length) return false;

  if (tokens.length === 1) {
    const token = tokens[0]!;
    const call = normalizeDiskNamePart(callSign);
    return (
      diskNamePartsClose(token, person.surname) ||
      diskNamePartsClose(token, person.firstName) ||
      Boolean(call && diskNamePartsClose(token, call))
    );
  }

  const file = alignFileFio(person, parseDiskFio(fileName));
  if (partsConflict(person.firstName, file.firstName)) return false;
  if (partsConflict(person.patronymic, file.patronymic)) return false;
  if (partsConflict(person.surname, file.surname)) return false;
  return true;
};

export const isExactFioFileNameMatch = (fullName: string, fileName: string) => {
  const person = parseDiskFio(fullName);
  if (!person.surname || !person.firstName || !person.patronymic) return false;

  const fileTokens = getNormalizedDiskNameTokens(fileName);
  const personKey = [person.surname, person.firstName, person.patronymic].join("|");
  for (let index = 0; index <= fileTokens.length - 3; index += 1) {
    const fileKey = fileTokens.slice(index, index + 3).join("|");
    if (fileKey === personKey) return true;
  }

  return false;
};

export const isUniqueSurnameFirstFileNameMatch = (
  fullName: string,
  fileName: string,
) => {
  const person = parseDiskFio(fullName);
  if (!person.surname || !person.firstName) return false;
  const fileTokens = getNormalizedDiskNameTokens(fileName);
  if (fileTokens.length < 2) return false;
  if (fileTokens[0] !== person.surname || fileTokens[1] !== person.firstName) {
    return false;
  }
  if (
    fileTokens.length >= 3 &&
    person.patronymic &&
    fileTokens[2] !== person.patronymic
  ) {
    return false;
  }
  return true;
};
