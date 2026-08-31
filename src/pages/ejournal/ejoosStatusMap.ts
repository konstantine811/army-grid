import {
  buildDefaultOperatorSettings,
  type EjoosOperatorSettings,
  type EjoosStatusConfidence,
  type EjoosStatusRule,
  type EjoosTimesheetCode,
} from "./ejoosRules";

const STORAGE_KEY = "army-grid:ejoos-operator-settings";

export type EjoosStatusMapping = {
  timesheetCode: EjoosTimesheetCode | null;
  absenceGround: string | null;
  confidence: EjoosStatusConfidence;
  label: string;
  reason: string;
  ruleId?: string;
};

export type { EjoosTimesheetCode };

const normalizeStatusText = (value: string) =>
  String(value ?? "")
    .toLowerCase()
    .replace(/ё/g, "е")
    .replace(/[''`´]/g, "")
    .replace(/\s+/g, " ")
    .trim();

/**
 * `_5 1ПБ`, `_1 1ПБ`, `1ПБ` — людина в своїй частині, у табелі «+».
 * `_5 2ПБ` сюди не входить (інший батальйон).
 */
const isOwnBattalionPresentStatus = (text: string) =>
  /^_?\d{0,2}\s*1\s*пб$/.test(text);

const ruleMatches = (rule: EjoosStatusRule, text: string) => {
  if (!rule.enabled) return false;
  if (rule.excludeAny?.some((token) => text.includes(token))) return false;
  if (rule.matchAll?.length && !rule.matchAll.every((token) => text.includes(token))) {
    return false;
  }
  if (!rule.matchAny.length) return false;
  return rule.matchAny.some((token) => {
    if (token === "+") return text === "+";
    if (token === "від") return text === "від" || text.includes("від ");
    return text.includes(token);
  });
};

export const mapPbStatusToEjoosWithRules = (
  rawStatus: string,
  rules: EjoosStatusRule[],
): EjoosStatusMapping => {
  const text = normalizeStatusText(rawStatus);
  if (!text || text === "0" || text === "-" || text === "—") {
    return {
      timesheetCode: null,
      absenceGround: null,
      confidence: "manual",
      label: "Порожній статус",
      reason: "У 1ПБ немає СТАТУС — потрібна ручна перевірка",
    };
  }

  const ordered = [...rules].sort((a, b) => a.priority - b.priority);
  for (const rule of ordered) {
    if (!ruleMatches(rule, text)) continue;
    return {
      timesheetCode: rule.timesheetCode,
      absenceGround: rule.absenceGround,
      confidence: rule.confidence,
      label: rule.label,
      reason: rule.reason,
      ruleId: rule.id,
    };
  }

  if (isOwnBattalionPresentStatus(text)) {
    return {
      timesheetCode: "+",
      absenceGround: null,
      confidence: "high",
      label: "У 1ПБ",
      reason: "СТАТУС — підрозділ 1ПБ (`_5 1ПБ` тощо) → код табеля «+»",
      ruleId: "own_unit",
    };
  }

  return {
    timesheetCode: null,
    absenceGround: null,
    confidence: "manual",
    label: "Невідомий статус",
    reason: `Немає однозначного мапінгу для «${rawStatus.trim()}» — оберіть код вручну`,
  };
};

/** Uses cached/local settings (or defaults). */
export const mapPbStatusToEjoos = (rawStatus: string): EjoosStatusMapping =>
  mapPbStatusToEjoosWithRules(rawStatus, readOperatorSettings().statusRules);

export const readOperatorSettings = (): EjoosOperatorSettings => {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return buildDefaultOperatorSettings();
    const parsed = JSON.parse(raw) as EjoosOperatorSettings;
    if (!parsed || parsed.version !== 1 || !Array.isArray(parsed.statusRules)) {
      return buildDefaultOperatorSettings();
    }
    return {
      ...buildDefaultOperatorSettings(parsed.unitLabel || "1ПБ"),
      ...parsed,
      statusRules: parsed.statusRules.length
        ? parsed.statusRules
        : buildDefaultOperatorSettings().statusRules,
      fieldAuthorities: parsed.fieldAuthorities?.length
        ? parsed.fieldAuthorities
        : buildDefaultOperatorSettings().fieldAuthorities,
    };
  } catch {
    return buildDefaultOperatorSettings();
  }
};

export const writeOperatorSettings = (settings: EjoosOperatorSettings) => {
  const next: EjoosOperatorSettings = {
    ...settings,
    version: 1,
    updatedAt: new Date().toISOString(),
  };
  localStorage.setItem(STORAGE_KEY, JSON.stringify(next));
  return next;
};

export const resetOperatorSettings = (unitLabel = "1ПБ") =>
  writeOperatorSettings(buildDefaultOperatorSettings(unitLabel));
