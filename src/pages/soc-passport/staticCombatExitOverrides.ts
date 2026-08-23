import { normalizePersonName, shortPersonName } from "./socPassportFields";

/**
 * Люди, яких вручну прибрали з Excel «Не виконували»
 * (порівняння повної 2026-08-23 vs відібраної 2026-08-22 (1) від Svitflofor).
 * У статистиці «По виконанню бойових завдань» тримаємо їх як тих, хто мав виходи
 * (мін. 1), навіть якщо в ранковому/ЖБД виходів немає.
 */
export const STATIC_COMBAT_EXIT_OVERRIDES: ReadonlyArray<{
  name: string;
  callsign: string;
  rank: string;
  note: string;
}> = [
  {
    name: "БУЛИЧЕВ Андрій Геннадійович",
    callsign: "БАЛУ",
    rank: "молодший сержант",
    note: "прибрано з «Не виконували» вручну (2026-08-22)",
  },
  {
    name: "ГАДАЧ Андрій Ігорович",
    callsign: "ВИШНЯ",
    rank: "молодший сержант",
    note: "прибрано з «Не виконували» вручну (2026-08-22)",
  },
  {
    name: "ДОРОШЕНКО Олексій Олександрович",
    callsign: "ДОРОХ",
    rank: "солдат",
    note: "прибрано з «Не виконували» вручну (2026-08-22)",
  },
  {
    name: "ЖУРАВЕЛЬ Олег Юрійович",
    callsign: "АПТЕКАР",
    rank: "сержант",
    note: "прибрано з «Не виконували» вручну (2026-08-22)",
  },
  {
    name: "КАНАЄВ Дмитро Миколайович",
    callsign: "ГЛИБА",
    rank: "солдат",
    note: "прибрано з «Не виконували» вручну (2026-08-22)",
  },
  {
    name: "КАРПАЧОВ Леонід Володимирович",
    callsign: "МАГНАТ",
    rank: "солдат",
    note: "прибрано з «Не виконували» вручну (2026-08-22)",
  },
  {
    name: "КОВАЛЕНКО Сергій Васильович",
    callsign: "ПИСАРЬ",
    rank: "сержант",
    note: "прибрано з «Не виконували» вручну (2026-08-22)",
  },
  {
    name: "КОМЛІЧЕНКО Володимир Володимирович",
    callsign: "ЛАГЕРНИЙ",
    rank: "солдат",
    note: "прибрано з «Не виконували» вручну (2026-08-22)",
  },
  {
    name: "МОРОЗ Андрій Миколайович",
    callsign: "ДРОН",
    rank: "молодший сержант",
    note: "прибрано з «Не виконували» вручну (2026-08-22)",
  },
  {
    name: "САМОЙЛЕНКО Сергій Олександрович",
    callsign: "ЛІТАК",
    rank: "старший лейтенант",
    note: "прибрано з «Не виконували» вручну (2026-08-22)",
  },
  {
    name: "ТАШКІНОВ Руслан Олександрович",
    callsign: "ВЕРЕСЕНЬ",
    rank: "молодший сержант",
    note: "прибрано з «Не виконували» вручну (2026-08-22)",
  },
  {
    name: "ТОЛОЧНИЙ Ігор Валерійович",
    callsign: "ВОВК",
    rank: "солдат",
    note: "прибрано з «Не виконували» вручну (2026-08-22)",
  },
  {
    name: "ТРИГУБ Сергій Олександрович",
    callsign: "АВГАН",
    rank: "молодший сержант",
    note: "прибрано з «Не виконували» вручну (2026-08-22)",
  },
  {
    name: "ШПОРТЬКО Роман Володимирович",
    callsign: "РИМ",
    rank: "молодший сержант",
    note: "прибрано з «Не виконували» вручну (2026-08-22)",
  },
] as const;

export const STATIC_COMBAT_EXIT_OVERRIDE_SOURCE =
  "Порівняння: Соц.паспорт_не_виконували_2026-08-23.xlsx (повна) мінус fromSvitflofor/…2026-08-22 (1).xlsx (відібрана).";

const overrideNameKeys = new Set(
  STATIC_COMBAT_EXIT_OVERRIDES.map((row) => normalizePersonName(row.name)),
);

const overrideShortKeys = new Set(
  STATIC_COMBAT_EXIT_OVERRIDES.map((row) => shortPersonName(row.name)),
);

export const isStaticCombatExitOverride = (name: string) => {
  const full = normalizePersonName(name);
  if (overrideNameKeys.has(full)) return true;
  const short = shortPersonName(name);
  return Boolean(short) && overrideShortKeys.has(short);
};

export const findStaticCombatExitOverride = (name: string) => {
  const full = normalizePersonName(name);
  const short = shortPersonName(name);
  return (
    STATIC_COMBAT_EXIT_OVERRIDES.find(
      (row) => normalizePersonName(row.name) === full,
    ) ??
    STATIC_COMBAT_EXIT_OVERRIDES.find(
      (row) => shortPersonName(row.name) === short,
    ) ??
    null
  );
};
