const titleCaseNamePart = (value: string) =>
  value
    .split("-")
    .map((part) =>
      part
        ? `${part.charAt(0).toLocaleUpperCase("uk-UA")}${part
            .slice(1)
            .toLocaleLowerCase("uk-UA")}`
        : "",
    )
    .join("-");

export const personNameFromSyntheticDocumentId = (
  externalId: string | null | undefined,
) => {
  let body = String(externalId ?? "").trim();
  if (!body || /^\d+$/.test(body)) return "";

  const prefixes = [
    "roster:archive:",
    "name-birth:",
    "name-call:",
    "roster:",
    "name:",
    "p:",
  ];
  let recognized = false;
  let changed = true;
  while (changed) {
    changed = false;
    for (const prefix of prefixes) {
      if (!body.toLocaleLowerCase("uk-UA").startsWith(prefix)) continue;
      body = body.slice(prefix.length).trim();
      recognized = true;
      changed = true;
      break;
    }
  }
  if (!recognized) return "";

  body = body
    .replace(/:\d{4}-\d{2}-\d{2}$/u, "")
    .replace(/:c:.*$/u, "")
    .replace(/\s+/g, " ")
    .trim();
  const parts = body.split(" ").filter(Boolean);
  if (
    parts.length < 2 ||
    !parts.every((part) => /[а-яіїєґ]/iu.test(part))
  ) {
    return "";
  }
  return [
    parts[0].toLocaleUpperCase("uk-UA"),
    ...parts.slice(1).map(titleCaseNamePart),
  ].join(" ");
};
