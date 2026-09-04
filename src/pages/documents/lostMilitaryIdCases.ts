import { toUkrainianDativeFullName, toUkrainianDativeRank } from "./form12Report";
import { toUkrainianGenitiveFullName } from "./form6Report";
import { toUkrainianGenitiveRank } from "./ubdRestoreReport";
import { capitalizeReportPosition } from "./reportPosition";

const titleCaseUk = (value: string) =>
  value
    ? value.charAt(0).toLocaleUpperCase("uk-UA") +
      value.slice(1).toLocaleLowerCase("uk-UA")
    : "";

const RANK_INSTRUMENTAL: Record<string, string> = {
  солдат: "солдатом",
  "старший солдат": "старшим солдатом",
  "молодший сержант": "молодшим сержантом",
  сержант: "сержантом",
  "старший сержант": "старшим сержантом",
  "головний сержант": "головним сержантом",
  "штаб-сержант": "штаб-сержантом",
  "майстер-сержант": "майстер-сержантом",
  "старший майстер-сержант": "старшим майстер-сержантом",
  "головний майстер-сержант": "головним майстер-сержантом",
  "молодший лейтенант": "молодшим лейтенантом",
  лейтенант: "лейтенантом",
  "старший лейтенант": "старшим лейтенантом",
  капітан: "капітаном",
  майор: "майором",
  підполковник: "підполковником",
  полковник: "полковником",
};

const GIVEN_INSTRUMENTAL: Record<string, string> = {
  анатолій: "Анатолієм",
  андрій: "Андрієм",
  богдан: "Богданом",
  вадим: "Вадимом",
  василь: "Василем",
  віктор: "Віктором",
  володимир: "Володимиром",
  дмитро: "Дмитром",
  іван: "Іваном",
  ігор: "Ігорем",
  максим: "Максимом",
  микола: "Миколою",
  михайло: "Михайлом",
  олег: "Олегом",
  олександр: "Олександром",
  олексій: "Олексієм",
  павло: "Павлом",
  петро: "Петром",
  роман: "Романом",
  сергій: "Сергієм",
  степан: "Степаном",
  тарас: "Тарасом",
  юрій: "Юрієм",
  ярослав: "Ярославом",
};

const GIVEN_GENITIVE: Record<string, string> = {
  анатолій: "Анатолія",
  андрій: "Андрія",
  богдан: "Богдана",
  вадим: "Вадима",
  василь: "Василя",
  віктор: "Віктора",
  володимир: "Володимира",
  дмитро: "Дмитра",
  іван: "Івана",
  ігор: "Ігоря",
  максим: "Максима",
  микола: "Миколи",
  михайло: "Михайла",
  олег: "Олега",
  олександр: "Олександра",
  олексій: "Олексія",
  павло: "Павла",
  петро: "Петра",
  роман: "Романа",
  сергій: "Сергія",
  степан: "Степана",
  тарас: "Тараса",
  юрій: "Юрія",
  ярослав: "Ярослава",
};

const splitName = (fullName: string) =>
  fullName
    .trim()
    .replace(/\s*\([^)]+\)\s*$/, "")
    .split(/\s+/)
    .filter(Boolean);

const toInstrumentalSurname = (surname: string) => {
  const value = surname.toLocaleUpperCase("uk-UA");
  if (/АГА$/u.test(value) || /ОГА$/u.test(value)) {
    return `${value.slice(0, -2)}ОЮ`;
  }
  if (/КА$/u.test(value)) return `${value.slice(0, -1)}ОЮ`;
  if (/А$/u.test(value)) return `${value.slice(0, -1)}ОЮ`;
  if (/КО$/u.test(value)) return `${value}М`;
  if (/ИЙ$/u.test(value)) return `${value.slice(0, -2)}ИМ`;
  if (/(ОВ|ЕВ|ЄВ|ІВ|ІН|ИН)$/u.test(value)) return `${value}ИМ`;
  if (/(УК|ЮК|АК|ИК)$/u.test(value)) return `${value}ОМ`;
  return `${value}ОМ`;
};

const toInstrumentalGiven = (given: string) => {
  const key = given.toLocaleLowerCase("uk-UA");
  if (GIVEN_INSTRUMENTAL[key]) return GIVEN_INSTRUMENTAL[key];
  if (key.endsWith("й")) return `${titleCaseUk(given.slice(0, -1))}єм`;
  if (key.endsWith("о")) return `${titleCaseUk(given)}м`;
  return `${titleCaseUk(given)}ом`;
};

const toInstrumentalPatronymic = (patronymic: string) => {
  if (/ович$/iu.test(patronymic)) {
    return patronymic.replace(/ович$/iu, "овичем");
  }
  if (/івна$/iu.test(patronymic) || /ївна$/iu.test(patronymic)) {
    return patronymic.replace(/вна$/iu, "вною");
  }
  return patronymic;
};

const toGenitiveSurname = (surname: string) => {
  const value = surname.toLocaleUpperCase("uk-UA");
  if (/КО$/u.test(value)) return `${value.slice(0, -1)}А`;
  if (/(УК|ЮК|АК|ИК)$/u.test(value)) return `${value}А`;
  if (/ИЙ$/u.test(value)) return `${value.slice(0, -2)}ОГО`;
  if (/(ОВ|ЕВ|ЄВ|ІН|ИН)$/u.test(value)) return `${value}А`;
  return `${value}А`;
};

const toGenitiveGiven = (given: string) => {
  const key = given.toLocaleLowerCase("uk-UA");
  if (GIVEN_GENITIVE[key]) return GIVEN_GENITIVE[key];
  if (key.endsWith("й")) return `${titleCaseUk(given.slice(0, -1))}я`;
  return `${titleCaseUk(given)}а`;
};

export const toUkrainianInstrumentalRank = (rank: string) => {
  const key = rank.trim().toLocaleLowerCase("uk-UA");
  if (!key) return "";
  if (RANK_INSTRUMENTAL[key]) return RANK_INSTRUMENTAL[key];
  return key
    .split(/\s+/)
    .map((word) => {
      if (word.endsWith("ий") || word.endsWith("ій")) {
        return `${word.slice(0, -2)}им`;
      }
      if (/[нтдк]$/i.test(word)) return `${word}ом`;
      return word;
    })
    .join(" ");
};

export const toUkrainianInstrumentalFullName = (fullName: string) => {
  const parts = splitName(fullName);
  if (!parts.length) return "";
  const [surname, given, patronymic] = parts;
  return [
    toInstrumentalSurname(surname),
    given ? toInstrumentalGiven(given) : "",
    patronymic ? toInstrumentalPatronymic(patronymic) : "",
  ]
    .filter(Boolean)
    .join(" ");
};

export const toUkrainianInstrumentalPosition = (position: string) => {
  const text = position.trim().replace(/[.,;:\s]+$/u, "");
  if (!text) return "";
  const [first, ...rest] = text.split(/\s+/);
  const key = first.toLocaleLowerCase("uk-UA");
  let instrumentalFirst = key;
  if (key.endsWith("ець")) instrumentalFirst = `${key.slice(0, -3)}цем`;
  else if (key.endsWith("ий") || key.endsWith("ій")) {
    instrumentalFirst = `${key.slice(0, -2)}им`;
  } else if (key.endsWith("ор") || key.endsWith("ер") || key.endsWith("ар")) {
    instrumentalFirst = `${key}ом`;
  } else if (!/(ом|ем|ею|ою)$/u.test(key)) {
    instrumentalFirst = `${key}ом`;
  }
  return capitalizeReportPosition([instrumentalFirst, ...rest].join(" "));
};

/** Знахідний → називний (для підпису внизу акту). */
export const toUkrainianNominativePosition = (position: string) => {
  const text = position.trim();
  if (!text) return "";
  const [first, ...rest] = text.split(/\s+/);
  const key = first.toLocaleLowerCase("uk-UA");
  let nominativeFirst = key;
  if (key.endsWith("ому") || key.endsWith("ему")) {
    nominativeFirst = `${key.slice(0, -3)}ий`;
  } else if (key.endsWith("ю")) {
    nominativeFirst = key.slice(0, -1);
  } else if (key.endsWith("у") && key.length > 3) {
    nominativeFirst = key.slice(0, -1);
  }
  return capitalizeReportPosition([nominativeFirst, ...rest].join(" "));
};

/** «Андрій КІЯНЕНКО» / «КІЯНЕНКО Андрій» → «Андрія КІЯНЕНКА». */
export const toUkrainianGenitiveGivenSurname = (fullName: string) => {
  const parts = splitName(fullName);
  if (!parts.length) return "";
  if (parts.length === 1) return toGenitiveSurname(parts[0]);
  const first = parts[0];
  const last = parts[parts.length - 1];
  const surnameFirst =
    first === first.toLocaleUpperCase("uk-UA") && /[А-ЯІЇЄҐA-Z]/.test(first);
  if (surnameFirst) {
    return `${toGenitiveGiven(parts[1] ?? first)} ${toGenitiveSurname(first)}`;
  }
  return `${toGenitiveGiven(first)} ${toGenitiveSurname(last)}`;
};

export const formatNominativeGivenSurname = (fullName: string) => {
  const parts = splitName(fullName);
  if (!parts.length) return "";
  if (parts.length === 1) return parts[0].toLocaleUpperCase("uk-UA");
  const first = parts[0];
  const last = parts[parts.length - 1];
  const surnameFirst =
    first === first.toLocaleUpperCase("uk-UA") && /[А-ЯІЇЄҐA-Z]/.test(first);
  if (surnameFirst) {
    return `${titleCaseUk(parts[1] ?? first)} ${first.toLocaleUpperCase("uk-UA")}`;
  }
  return `${titleCaseUk(first)} ${last.toLocaleUpperCase("uk-UA")}`;
};

export {
  toUkrainianDativeFullName,
  toUkrainianDativeRank,
  toUkrainianGenitiveFullName,
  toUkrainianGenitiveRank,
};
