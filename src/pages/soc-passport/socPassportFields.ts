import { tryParseExcelSerialDate } from "../../shared/format";
import type {
  AgeBand,
  ArrivalSource,
  ExitBand,
  MaritalStatus,
  NationalityKey,
  RankGroup,
  RankServiceBucket,
  RegionKey,
  ServiceType,
  Sex,
} from "./socPassportTypes";

export const emptyBucketCounts = (): Record<RankServiceBucket, number> => ({
  officerMobilized: 0,
  officerContract: 0,
  sergeantMobilized: 0,
  sergeantContract: 0,
  soldierMobilized: 0,
  soldierContract: 0,
});

export const bucketOf = (
  rankGroup: RankGroup,
  serviceType: ServiceType,
): RankServiceBucket => {
  if (rankGroup === "officer") {
    return serviceType === "contract" ? "officerContract" : "officerMobilized";
  }
  if (rankGroup === "sergeant") {
    return serviceType === "contract" ? "sergeantContract" : "sergeantMobilized";
  }
  return serviceType === "contract" ? "soldierContract" : "soldierMobilized";
};

export const normalizeLooseText = (value: string) =>
  value
    .replace(/[ʼ’'`]/g, "")
    .replace(/ё/gi, "е")
    .replace(/\s+/g, " ")
    .trim()
    .toLocaleLowerCase("uk-UA");

export const normalizePersonName = (value: string) =>
  normalizeLooseText(value)
    .replace(/є/g, "е")
    .replace(/\([^)]*\)/g, " ")
    .replace(/[.,;:№#"/\\|()[\]{}]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();

export const normalizeRnokpp = (value: string) => {
  const digits = value.replace(/\D/g, "");
  return digits.length >= 8 ? digits : "";
};

export const shortPersonName = (value: string) => {
  const parts = normalizePersonName(value).split(" ").filter(Boolean);
  return parts.slice(0, 2).join(" ");
};

export const looksLikePersonName = (value: string) => {
  const text = value.replace(/\s+/g, " ").trim();
  if (!text || text.length < 5) return false;
  if (/^(управління|рота|взвод|батальйон|група|відділення|штаб)/i.test(text)) {
    return false;
  }
  const parts = text.split(" ").filter(Boolean);
  if (parts.length < 2 || parts.length > 5) return false;
  return /^[А-ЯІЇЄҐA-Z][А-ЯІЇЄҐA-Za-zа-яіїєґ'\-]+$/.test(parts[0] ?? "");
};

const OFFICER_RANK_RE =
  /(генерал|полковник|підполковник|майор|капітан|лейтенант)/i;
const SOLDIER_RANK_RE = /(солдат|матрос|рекрут|рядовий)/i;

export const classifyRankGroup = (
  rank: string,
  fallbackPosition = "",
): RankGroup => {
  const text = normalizeLooseText(`${rank} ${fallbackPosition}`);
  const category = normalizeLooseText(rank);
  if (/^оф\.?$/.test(category)) return "officer";
  if (/^сер[жh]\.?$/.test(category)) return "sergeant";
  if (/^солд\.?$/.test(category)) return "soldier";
  if (OFFICER_RANK_RE.test(text) && !/сержант/.test(text)) return "officer";
  if (
    /(сержант|старшина|прапорщик|штаб-сержант|головний сержант)/i.test(text)
  ) {
    return "sergeant";
  }
  if (SOLDIER_RANK_RE.test(text)) return "soldier";
  if (
    /(командир батальйону|начальник штабу|офіцер|заступник командира батальйону)/i.test(
      fallbackPosition,
    )
  ) {
    return "officer";
  }
  if (/(сержант|старшин)/i.test(fallbackPosition)) return "sergeant";
  return "soldier";
};

export const classifyServiceType = (value: string): ServiceType => {
  const text = normalizeLooseText(value);
  if (!text) return "mobilized";
  if (/контракт/.test(text) && !/мобіл/.test(text)) return "contract";
  return "mobilized";
};

export const classifySex = (value: string): Sex => {
  const text = normalizeLooseText(value).replace(/\./g, "");
  if (!text) return "male";
  if (/^(ж|жін|жіноч|female|f)$/.test(text) || /жіноч/.test(text)) {
    return "female";
  }
  if (/^(ч|чол|чоловіч|male|m)$/.test(text) || /чоловіч/.test(text)) {
    return "male";
  }
  // Нерозпізнане / порожнє → чоловік (дефолт для соц.паспорта).
  return "male";
};

export const parseAgeYears = (
  birthValue: string,
  listedYears: string,
  asOf = new Date(),
): number | null => {
  const listed = Number(String(listedYears).replace(",", ".").replace(/[^\d.]/g, ""));
  if (Number.isFinite(listed) && listed >= 17 && listed <= 75) return Math.round(listed);

  const serialMatch = birthValue.trim().match(/^(\d{4,5})(?:\.0+)?$/);
  if (serialMatch) {
    const serial = Number(serialMatch[1]);
    const date = new Date((serial - 25569) * 86400000);
    const year = date.getUTCFullYear();
    if (year >= 1950 && year <= 2010) {
      return ageFromDate(date, asOf);
    }
  }

  const match = birthValue.match(/(\d{1,2})[.\/-](\d{1,2})[.\/-](\d{2,4})/);
  if (match) {
    let year = Number(match[3]);
    if (year < 100) year += year >= 50 ? 1900 : 2000;
    const month = Number(match[2]);
    const day = Number(match[1]);
    if (year >= 1950 && year <= 2010) {
      return ageFromDate(new Date(Date.UTC(year, month - 1, day)), asOf);
    }
  }

  const yearOnly = birthValue.match(/(?:19|20)\d{2}/);
  if (yearOnly) {
    const year = Number(yearOnly[0]);
    if (year >= 1950 && year <= 2010) return asOf.getUTCFullYear() - year;
  }

  return null;
};

const ageFromDate = (date: Date, asOf: Date) => {
  let age = asOf.getUTCFullYear() - date.getUTCFullYear();
  const monthDelta = asOf.getUTCMonth() - date.getUTCMonth();
  if (monthDelta < 0 || (monthDelta === 0 && asOf.getUTCDate() < date.getUTCDate())) {
    age -= 1;
  }
  return age >= 16 && age <= 75 ? age : null;
};

export const ageBandOf = (age: number | null): AgeBand => {
  if (age == null) return "unknown";
  if (age <= 25) return "18-25";
  if (age <= 30) return "26-30";
  if (age <= 40) return "31-40";
  if (age <= 50) return "41-50";
  return "50+";
};

/** БЗВП/БРЕЗ (X) + Відрядження (БРЕЗ) (AA) з ранкового списку. */
export const combineMorningBrezFields = (
  bzvpStatus: string,
  brezAssignment = "",
) => `${bzvpStatus} ${brezAssignment}`.trim();

export const isBrezMarkerText = (value: string) =>
  /бре[зc]|brez/.test(normalizeLooseText(value));

export const classifyArrivalSource = (
  arrivedFrom: string,
  calledBy: string,
  morningBrez: string,
): ArrivalSource => {
  const text = normalizeLooseText(`${arrivedFrom} ${calledBy} ${morningBrez}`);
  if (!text) return "unknown";
  if (/бре[зc]|brez/.test(text)) return "brez";
  if (/рекрут/.test(text)) return "recruiting";
  if (/(нц|навчальн|184\s*нц|бзвп)/.test(text) && !/тцк/.test(arrivedFrom.toLowerCase())) {
    if (/(нц|навчальн)/.test(text)) return "trainingCenter";
  }
  if (/тцк|комісар|военкомат|військ.?ком/.test(text)) return "tck";
  if (/в\/ч|а\d{4}/.test(text) || /транзит/.test(text)) return "other";
  return text ? "other" : "unknown";
};

const UKRAINIAN_BIRTH_RE =
  /(україн|укр\.?\s*рср|урср|дн[іi]про|днепр|дныпро|харк[іi]в|харьк[іi]в|донецьк|луганск|луганськ|одес|олеса|льв[іi]в|ки[їі]вськ|ки[їієе]в|запор[іi]зьк|запорозьк|микола[їі]в|херсон|полтав|черкас|черкасськ|в[іi]нниць|хмельниц|черн[іi]г|сумськ|р[іi]внен|рывнен|ровенськ|волин|терноп|франк[іi]в|закарпат|житомир|к[іi]ровоград|кировоград|каровоград|кропивниц|кропівниц|крим|севастопол|арк\b|б[іi]лгород.?дн[іi]стров|белгород.?дн[іi]стров|кривий\s*р[іi]г|павлоград|палоград|кам.?янськ|н[іi]копол|новомосковськ|[іi]зюм|чугу[єеїі]в|лозова|куп.?янськ|чорноморськ|[іi]зма[їі]л|южне|очак[іi]в|кобел[яе]|кременчу[кг]|св[іi]тловодськ|борисп[іi]ль|б[іi]ла\s*церква|вараш|ра[хк][іi]в|подол[ьл]ськ|гуляйпол|ор[іi]х[іi]в|волчанськ|вовчанськ|мерефа|марган[еє]ць|мрганець|п[еє]тр[іи]к[іi]вка|конотоп|глух[іi]в|охтирк|шостк|н[іi]жин|прилук|луцьк|ков[еє]ль|тернопіл|ужгород|мукачев|черн[іi]вц|житомир|бердич[іi]в|коростен|знам.?янк|нова\s*каховк|краматорськ|слов.?янськ|костянтин[іi]вк|констянтин[іi]вк|дружк[іi]вк|покровськ|мирноград|добропілл|мар[іi]упол|мак[іi][їі]вк|горл[іi]вк|алчевськ|с[іi]верськодонецьк|с[єе]в[єе]родонецьк|курах[ое]в|соледар|св[іi]тлодарськ|к[іi]вшар[іi]вка|петропавл[іi]вка|в[іi]льнянськ|нова\s*праг|новоукраїнк)/;

/** Українське місце народження (національність за замовчуванням — Україна). */
export const isUkrainianBirthPlace = (birthPlace: string) =>
  UKRAINIAN_BIRTH_RE.test(normalizeLooseText(birthPlace));

/** Громадянство РФ у тексті анкети — не «родич проживає в рф». */
const hasRussiaCitizenshipInText = (chunk: string) => {
  if (!chunk) return false;
  const text = normalizeLooseText(chunk);
  if (/україн(?:ин|ка|ськ)|громадян(?:ство|ин)?[^\n,.]{0,24}укра/.test(text)) {
    return false;
  }
  return (
    /(?:^|[^а-яіїєґ])(?:росіян(?:ин|ка|и)?|русск(?:ий|ая|ие)?)(?:[^а-яіїєґ]|$)/.test(
      text,
    ) ||
    /громадян(?:ство|ин)?[^\n,.]{0,32}(?:р\.?\s*ф\.?|росі[яї])(?:[^а-яіїєґ]|$)/.test(
      text,
    ) ||
    /(?:р\.?\s*ф\.?|росі[яї])[^\n,.]{0,32}громадян/.test(text) ||
    /паспорт[^\n,.]{0,28}(?:р\.?\s*ф\.?|росі[яї])/.test(text) ||
    /національн(?:ість|ist)?[^\n,.]{0,20}(?:р\.?\s*ф\.?|росі[яї]|росіян)/.test(
      text,
    ) ||
    /(?:р\.?\s*ф\.?|росі[яї])[^\n,.]{0,20}національн/.test(text)
  );
};

export const classifyNationality = (
  relatives: string,
  extra: string,
  birthPlace: string,
): { key: NationalityKey; notes: string[] } => {
  const text = normalizeLooseText(`${relatives} ${extra}`);
  const birth = normalizeLooseText(birthPlace);
  const notes: string[] = [];

  const foreignFromBirth = detectForeignCountryFromBirthPlace(birth);
  if (foreignFromBirth) {
    notes.push(`національність: місце народження → ${foreignFromBirth.key}`);
    return { key: foreignFromBirth.key, notes };
  }

  if (/аргентин/.test(text)) return { key: "argentina", notes };
  if (/португал/.test(text)) return { key: "portugal", notes };
  if (/(великобритан|британі|англі[йя])/.test(text)) return { key: "britain", notes };
  if (/(^|[^а-яіїєґ])сша([^а-яіїєґ]|$)|америк/.test(text)) {
    return { key: "usa", notes };
  }
  if (/поляк|польщ/.test(text) && !/україн/.test(text)) {
    return { key: "poland", notes };
  }
  if (hasRussiaCitizenshipInText(extra)) {
    notes.push("національність: з додаткової інформації (громадянство РФ)");
    return { key: "russia", notes };
  }
  if (hasRussiaCitizenshipInText(relatives)) {
    notes.push("національність: з тексту про родичів (громадянство РФ)");
    return { key: "russia", notes };
  }
  if (
    /(молдав|молдов|б[іi]лорус|беларус|грузин|вірмен|азербайдж|євре[йя]|циган|румун|болгар|чех|казах|узбек|таджик|туркмен|киргиз|литв|латв|естон|н[іi]меччин|ізра[їі]л|туречч|грузи)/.test(
      text,
    )
  ) {
    notes.push("національність: інша за ключовим словом");
    return { key: "other", notes };
  }
  notes.push("національність: немає окремого поля, зараховано як Україна");
  return { key: "ukraine", notes };
};

/** Країна народження з «місце народження» (коли немає колонки національності). */
const detectForeignCountryFromBirthPlace = (
  birth: string,
): { key: NationalityKey } | null => {
  if (!birth) return null;

  if (UKRAINIAN_BIRTH_RE.test(birth)) return null;

  if (
    /(^|[^а-яіїєґ0-9])р\.?\s*ф\.?([^а-яіїєґ0-9]|$)|росі[яї]|рос[іi]йськ|рсфср/.test(
      birth,
    )
  ) {
    return { key: "russia" };
  }

  // Міста/області РФ (без коротких стебел на кшталт «самар», щоб не чіпати «Самарський р-н» Дніпра).
  if (
    /(псков|мурманськ|сахал[іi]н|магадан|якут[іi]я|якутія.?саха|приморськ(?:ий)?\s*край|хабаровськ|тамбовськ|таганрог|ростовськ(?:а)?\s*обл|волгоград|волжск|респ\.?\s*ком[іi]|амурськ|благов[іi]щен|оленег[іi]р|севером|півн[іi]чномор|южно-сахал[іi]н|владимир|володимир\s*рф|чайковськ|дагестан|апшерон|уралськ(?:а)?\s*обл|переметн|ухта|краснодар|челяб[іi]нськ|єкатерин|екатерин|новосиб|казан|воронеж|тул[аеи]|брянськ|курськ(?:а)?\s*обл|смоленськ|калуг[аи]|твер(?:ськ)?|рязан|саратов|самарськ(?:а)?\s*обл|омськ|томськ|пермск|пермськ|уфа|іркутськ|красноярськ|астрахан|шахтинск|кослан)/.test(
      birth,
    )
  ) {
    return { key: "russia" };
  }

  if (/б[еє]лгород/.test(birth) && !/дн[іi]стров/.test(birth)) {
    return { key: "russia" };
  }

  if (/аргентин/.test(birth)) return { key: "argentina" };
  if (/португал/.test(birth)) return { key: "portugal" };
  if (/(великобритан|британі|англі[йя])/.test(birth)) return { key: "britain" };
  if (/(^|[^а-яіїєґ])сша([^а-яіїєґ]|$)|америк/.test(birth)) return { key: "usa" };
  if (/польщ|польша|варшав|крак[іi]в|гданськ|легниц/.test(birth)) {
    return { key: "poland" };
  }

  if (OTHER_COUNTRY_BIRTH_RE.test(birth)) {
    return { key: "other" };
  }

  return null;
};


/** Країни поза рядками Україна / Польща / США / Британія / Португалія / Аргентина / РФ. */
const OTHER_COUNTRY_BIRTH_RE =
  /(казахстан|казастан|алма.?ата|алмати|узбек|ташкент|таджик|душанбе|туркмен|ашхабад|киргиз|б[іi]шкек|молдов|молдав|кишин[іi]в|бендер|придн[іi]стров|вірмен|єреван|ванадзор|к[іi]ровакан|азербайдж|баку|грузі|тб[іi]л[іi]с|абхаз|осет[іi]|б[іi]лорус|беларус|м[іi]нськ|гомел|брес[тц]|литв|в[іi]льнюс|латв[іi][яю]|р[іi]га|естон|таллін|н[іi]меччин|(^|[^а-яіїєґ0-9])фрн([^а-яіїєґ0-9]|$)|дрезден|берл[іi]н|чех[іi][яю]|(^|[^а-яіїєґ])праг[ауеи]([^а-яіїєґ]|$)|словач|румун|бухарест|болгар|соф[іi][яю]|угорщ|будапешт|серб[іi][яю]|чорногор|хорват|босн[іi]|македон|албан[іi]|грец[іi]|афін|туречч|стамбул|анкар|ізра[їі]л|тель.?ав[іi]в|іран|ірак|сирі[яю]|єгипет|китай|пек[іi]н|япон[іi]|коре[яю]|інді[яю]|пакистан|афган|канад|мексик|браз[іи]л|куб[аеи]|австрал|нов[ао]\s*зеланд|франц[іi]|париж|італ[іi]|(^|[^а-яіїєґ])рим([^а-яіїєґ]|$)|іспан[іi]|мадрид|н[іi]дерланд|голланд|бельг[іi]|швейцар|австр[іi]|(^|[^а-яіїєґ])в[іi]ден([^а-яіїєґ]|$)|швед[іi]|норвег|дан[іi][яю]|ф[іi]нлянд)/;

const CITY_TO_OBLAST: Array<{ re: RegExp; region: RegionKey }> = [
  { re: /ки[їієе]в/, region: "kyivCity" },
  { re: /борисп[іi]ль|б[іi]ла\s*церква|вараш/, region: "kyiv" },
  {
    re: /кривий\s*р[іi]г|дн[іi]пр|днепр|дныпро|павлоград|палоград|кам.?янськ|н[іi]копол|новомосковськ|п[еє]тр[іи]к[іi]вка|марган|мрганець/,
    region: "dnipro",
  },
  {
    re: /харк[іi]в|харьк[іi]в|[іi]зюм|чугу[єеїі]в|лозова|куп.?янськ|мерефа|вовчанськ|волчанськ|к[іi]вшар[іi]вка|великий\s*бурлук/,
    region: "kharkivControlled",
  },
  {
    re: /одес|олеса|чорноморськ|[іi]зма[їі]л|б[іi]лгород.?дн[іi]стров|белгород.?дн[іi]стров|южне|под[іi]льськ/,
    region: "odesa",
  },
  { re: /льв[іi]в|дрогобич|стрий|червоноград/, region: "lviv" },
  {
    re: /запор[іi]ж|запорозьк|бердянськ|мел[іi]топол|гуляйпол|ор[іi]х[іi]в|в[іi]льнянськ/,
    region: "zaporizhzhiaControlled",
  },
  { re: /микола[їі]в|первомайськ|южноукраїнськ|очак[іi]в/, region: "mykolaiv" },
  { re: /полтав|кременчу[кг]|лубен|миргород|кобел[яе]/, region: "poltava" },
  { re: /черкас|умань|звенигородк|сміл/, region: "cherkasy" },
  { re: /вінниц|жмеринк|могил[іi]в.?под/, region: "vinnytsia" },
  { re: /хмельниц|кам.?янець|шепетівк/, region: "khmelnytskyi" },
  { re: /черн[іi]г[іi]в|н[іi]жин|прилук/, region: "chernihiv" },
  { re: /сум[иы]|глух[іi]в|конотоп|охтирк|шостк|буринськ/, region: "sumy" },
  { re: /рівн[ео]|рывнен|ровенськ|дубно|костопіль|сарн/, region: "rivne" },
  { re: /луцьк|ков[еє]ль|володимир/, region: "volyn" },
  { re: /тернопіл|кременець|чортк[іi]в/, region: "ternopil" },
  { re: /івано.?франк|[іi]вана\s*франк|коломи|калуш/, region: "ivanoFrankivsk" },
  { re: /ужгород|мукачев|хуст|берегов|ра[хк][іi]в/, region: "zakarpattia" },
  { re: /черн[іi]вц/, region: "chernivtsi" },
  { re: /житомир|бердич[іi]в|коростен/, region: "zhytomyr" },
  {
    re: /кропивниц|кропівниц|к[іi]ровоград|кировоград|каровоград|олександрі|знам.?янк|св[іi]тловодськ|нова\s*праг|новоукраїнк/,
    region: "kirovohrad",
  },
  { re: /херсон|нова\s*каховк/, region: "khersonControlled" },
  {
    re: /краматорськ|слов.?янськ|костянтин[іi]вк|констянтин[іi]вк|дружк[іi]вк|покровськ|мирноград|добропілл|курах[ое]в|соледар|св[іi]тлодарськ|петропавл[іi]вка/,
    region: "donetskControlled",
  },
  { re: /мар[іi]упол|донецьк|мак[іi][їі]вк|горл[іi]вк/, region: "donetskOccupied" },
  {
    re: /луганск|луганськ|алчевськ|с[іi]верськодонецьк|с[єе]в[єе]родонецьк/,
    region: "luhansk",
  },
  {
    re: /сімферопол|севастопол|ялта|керч|євпаторі|красноперекопськ|(^|[^а-яіїєґ])арк([^а-яіїєґ]|$)/,
    region: "crimea",
  },
];

const OBLAST_MAP: Array<{ re: RegExp; region: RegionKey }> = [
  { re: /ки[їі]вськ/, region: "kyiv" },
  { re: /в[іi]нницьк/, region: "vinnytsia" },
  { re: /волинськ/, region: "volyn" },
  { re: /дн[іi]пропетровськ/, region: "dnipro" },
  { re: /донецьк/, region: "donetskControlled" },
  { re: /житомирськ/, region: "zhytomyr" },
  { re: /закарпатськ/, region: "zakarpattia" },
  { re: /[іi]вано.?франк|[іi]вана\s*франк/, region: "ivanoFrankivsk" },
  { re: /к[іi]ровоградськ|кировоградськ|каровоградськ/, region: "kirovohrad" },
  { re: /луганськ|луганск/, region: "luhansk" },
  { re: /льв[іi]вськ/, region: "lviv" },
  { re: /микола[їі]вськ/, region: "mykolaiv" },
  { re: /одес+ьк/, region: "odesa" },
  { re: /полтавськ/, region: "poltava" },
  { re: /р[іi]вненськ|рывненськ|ровенськ/, region: "rivne" },
  { re: /сумськ/, region: "sumy" },
  { re: /терноп[іi]льськ/, region: "ternopil" },
  { re: /запор[іi]зьк|запорозьк/, region: "zaporizhzhiaControlled" },
  { re: /харк[іi]вськ|харьк[іi]вськ/, region: "kharkivControlled" },
  { re: /херсонськ/, region: "khersonControlled" },
  { re: /хмельницьк/, region: "khmelnytskyi" },
  { re: /черкаськ|черкасськ/, region: "cherkasy" },
  { re: /черн[іi]вецьк/, region: "chernivtsi" },
  { re: /черн[іi]г[іi]вськ/, region: "chernihiv" },
  { re: /крим|севастопол/, region: "crimea" },
];

const OCCUPIED_PLACE_RE =
  /(окуп|тимчасово\s*окуп|неп[іi]дконтрольн|анексован|маріупол|донецьк(?!а обл)|мак[іi][їі]вк|горл[іi]вк|єнакі[єе]|сніжн|торез|харцизьк|новоазовськ|докучаєвськ|волновах|бердянськ|мел[іi]топол|енергодар|ген[іi]чеськ|скадовськ|алушт|ялта|керч|євпаторі|сімферопол|севастопол|луганськ|алчевськ|каховк|токм|полог)/i;

export const REGION_LABELS: Record<RegionKey, string> = {
  kyivCity: "м. Київ",
  kyiv: "Київська обл.",
  vinnytsia: "Вінницька обл.",
  volyn: "Волинська обл.",
  dnipro: "Дніпропетровська обл.",
  donetskControlled: "Донецька обл. (підконтрольна частина)",
  donetskOccupied: "Донецька обл. (тимчасово окупована частина)",
  zhytomyr: "Житомирська обл.",
  zakarpattia: "Закарпатська обл.",
  ivanoFrankivsk: "Івано-Франківська обл.",
  kirovohrad: "Кіровоградська обл.",
  luhansk: "Луганська обл.",
  lviv: "Львівська обл.",
  mykolaiv: "Миколаївська обл.",
  odesa: "Одеська обл.",
  poltava: "Полтавська обл.",
  rivne: "Рівненська обл.",
  sumy: "Сумська обл.",
  ternopil: "Тернопільська обл.",
  zaporizhzhiaControlled: "Запорізька обл. (підконтрольна частина)",
  zaporizhzhiaOccupied: "Запорізька обл. (не підконтрольна частина)",
  kharkivControlled: "Харківська обл. (підконтрольна частина)",
  kharkivOccupied: "Харківська обл. (непідконтрольна частина)",
  khersonControlled: "Херсонська обл. (підконтрольна частина)",
  khersonOccupied: "Херсонська обл. (непідконтрольна частина)",
  khmelnytskyi: "Хмельницька обл.",
  cherkasy: "Черкаська обл.",
  chernivtsi: "Чернівецька обл.",
  chernihiv: "Чернігівська обл.",
  crimea: "АР Крим",
  unknown: "Регіон не визначено",
};

const splitOccupied = (region: RegionKey, text: string): RegionKey => {
  const occupied = OCCUPIED_PLACE_RE.test(text) || /окуп/.test(text);
  if (region === "donetskControlled" && occupied) return "donetskOccupied";
  if (region === "zaporizhzhiaControlled" && occupied) {
    return "zaporizhzhiaOccupied";
  }
  if (region === "kharkivControlled" && occupied) return "kharkivOccupied";
  if (region === "khersonControlled" && occupied) return "khersonOccupied";
  if (region === "crimea") return "crimea";
  return region;
};

export const classifyRegion = (
  birthPlace: string,
  calledBy: string,
): { key: RegionKey; occupied: boolean; notes: string[] } => {
  const notes: string[] = [];
  const birth = normalizeLooseText(birthPlace);
  const tck = normalizeLooseText(calledBy);
  const combined = `${birth} ${tck}`;

  // Народжені за межами України — не підганяємо під українські області.
  if (detectForeignCountryFromBirthPlace(birth)) {
    notes.push("місце народження за межами України — регіон не мапимо");
    return { key: "unknown", occupied: false, notes };
  }

  if (
    /^м\.?\s*ки[їієе]в$/.test(birth) ||
    (/(^|[^а-яіїєґ])ки[їієе]в([^а-яіїєґ]|$)/.test(birth) &&
      !/ки[їі]вськ/.test(birth) &&
      !/обл/.test(birth))
  ) {
    return { key: "kyivCity", occupied: false, notes };
  }

  for (const item of OBLAST_MAP) {
    if (item.re.test(birth)) {
      const key = splitOccupied(item.region, birth);
      return { key, occupied: /Occupied|crimea/.test(key), notes };
    }
  }

  for (const item of CITY_TO_OBLAST) {
    if (item.re.test(birth)) {
      const key = splitOccupied(item.region, birth);
      notes.push("регіон за населеним пунктом народження");
      return { key, occupied: /Occupied|crimea/.test(key), notes };
    }
  }

  for (const item of OBLAST_MAP) {
    if (item.re.test(tck)) {
      const key = splitOccupied(item.region, combined);
      notes.push("регіон за ТЦК (місце народження без області)");
      return { key, occupied: /Occupied|crimea/.test(key), notes };
    }
  }

  if (birth || tck) notes.push("регіон не розпізнано");
  return { key: "unknown", occupied: OCCUPIED_PLACE_RE.test(combined), notes };
};

const ROLE_SPLIT_RE =
  /(?=(?:мати|мама|батько|дружина|чолов[іi]к|цив[іi]льн[аої\s.-]*дружин|цив[іi]льний\s*чолов[іi]к|сп[іi]вмешкан|д[іi]ти|дитина|син|дон[ьк]ка|дочка|сестра|брат|дов[іi]рен[аої\s]*особ)\s*[:;]?)/gi;

const extractRoleBlocks = (raw: string) => {
  const text = raw.replace(/\s+/g, " ").trim();
  if (!text) return [] as Array<{ role: string; body: string }>;
  const parts = text.split(ROLE_SPLIT_RE).map((part) => part.trim()).filter(Boolean);
  return parts.map((part) => {
    const match = part.match(
      /^(мати|мама|батько|дружина|чолов[іi]к|цив[іi]льн[аої\s.-]*дружин[аи]?|цив[іi]льний\s*чолов[іi]к|сп[іi]вмешканк[аи]|д[іi]ти|дитина|син|дон[ьк]ка|дочка|сестра|брат|дов[іi]рен[аої\s]*особа?)\s*[:;]?\s*(.*)$/i,
    );
    if (!match) return { role: "", body: part };
    return { role: normalizeLooseText(match[1] ?? ""), body: match[2] ?? "" };
  });
};

const yearsInText = (text: string) =>
  [...text.matchAll(/(?:19|20)\d{2}/g)]
    .map((match) => Number(match[0]))
    .filter((year) => year >= 1988 && year <= 2026);

const childCountFromBody = (body: string, asOfYear: number) => {
  const years = yearsInText(body);
  if (years.length) {
    const under18 = years.filter((year) => asOfYear - year < 18).length;
    return { childCount: years.length, childrenUnder18: under18 };
  }
  const named = body
    .split(/[,;]/)
    .map((part) => part.trim())
    .filter((part) => /[А-ЯІЇЄҐA-Zа-яіїєґ]{2,}/.test(part) && !/^\+?\d/.test(part));
  if (named.length) {
    return { childCount: named.length, childrenUnder18: named.length };
  }
  if (/син|дон[ьк]|дитин|хлопц|д[іi]вчин/.test(normalizeLooseText(body))) {
    return { childCount: 1, childrenUnder18: 1 };
  }
  return { childCount: body.trim() ? 1 : 0, childrenUnder18: body.trim() ? 1 : 0 };
};

export const parseRelatives = (
  raw: string,
  asOf = new Date(),
): {
  marital: MaritalStatus;
  childCount: number;
  childrenUnder18: number;
  relativesServing: boolean;
  relativesAbroad: boolean;
  relativesHostile: boolean;
  notes: string[];
} => {
  const notes: string[] = [];
  const text = raw.replace(/\s+/g, " ").trim();
  if (!text) {
    return {
      marital: "unmarried",
      childCount: 0,
      childrenUnder18: 0,
      relativesServing: false,
      relativesAbroad: false,
      relativesHostile: false,
      notes: ["родичі порожні: дефолт неодружений, без дітей"],
    };
  }

  const lower = normalizeLooseText(text);
  const blocks = extractRoleBlocks(text);
  const divorced = /розлучен|вд[іi]вець|вдова/.test(lower);

  // Повний текст важливіший за блоки: часто «Цивільна дружина Іванова» без «:».
  const hasCivil =
    /цив[іi]льн[аої\s.-]*дружин|цив[іi]льний\s*чолов[іi]к|сп[іi]вмешканк|цив[іi]льн[ий]*\s*шлюб/.test(
      lower,
    ) || blocks.some((block) => /цив[іi]льн|сп[іi]вмешкан/.test(block.role));
  const hasSpouse =
    (!hasCivil &&
      (/([^а-яіїєґ]|^)дружин[аи]([^а-яіїєґ]|$)|([^а-яіїєґ]|^)чолов[іi]к([^а-яіїєґ]|$)|одруж/.test(
        lower,
      ) ||
        blocks.some(
          (block) =>
            /дружин|чолов[іi]к/.test(block.role) && !/цив[іi]льн/.test(block.role),
        ))) ||
    false;

  let marital: MaritalStatus = "unmarried";
  if (hasCivil && !divorced) {
    marital = "civil";
    notes.push("сімейний: цивільна дружина/чоловік");
  } else if (hasSpouse && !divorced) {
    marital = "married";
  } else if (divorced) {
    marital = "unmarried";
    notes.push("сімейний: розлучений/вдівець → неодружений");
  }

  let childCount = 0;
  let childrenUnder18 = 0;
  const childBlocks = blocks.filter((block) =>
    /д[іi]ти|дитина|син|дон[ьк]|дочка/.test(block.role),
  );
  if (childBlocks.length) {
    for (const block of childBlocks) {
      const parsed = childCountFromBody(block.body, asOf.getUTCFullYear());
      childCount += parsed.childCount;
      childrenUnder18 += parsed.childrenUnder18;
    }
  } else if (/д[іi]ти|дитина|син|дон[ьк]/.test(lower)) {
    const parsed = childCountFromBody(text, asOf.getUTCFullYear());
    childCount = parsed.childCount;
    childrenUnder18 = parsed.childrenUnder18;
    notes.push("діти: блок ролі не виділено, пораховано евристикою");
  }

  const relativesServing =
    /(служить|проходить\s*служб|зсу|\bв\/ч\b|військов)/i.test(text) &&
    !/дружина\s*:/i.test(text.slice(0, 20));
  const relativesAbroad =
    /(за\s*кордон|польщ|н[іi]меччин|чех[іi][яю]|італ[іi][яю]|іспан[іi][яю]|\bсша\b|канада|[іi]зра[їі]л|литв|латві|естон)/i.test(
      text,
    );
  const relativesHostile = /(росі[яї]|рф\b|білорус|беларус|москва|м[іi]нськ)/i.test(
    text,
  );

  return {
    marital,
    childCount,
    childrenUnder18,
    relativesServing,
    relativesAbroad,
    relativesHostile,
    notes,
  };
};

export const parseUbdInfo = (extra: string, relatives = "") => {
  const text = `${extra} ${relatives}`.trim();
  const hasUbd = /убд|удб/i.test(text);
  const numberMatch = text.match(
    /(?:убд|удб)\s*(?:№|#|no\.?|номер)?\s*([0-9]{4,9})/i,
  );
  return { hasUbd, ubdNumber: numberMatch?.[1]?.trim() ?? "" };
};

export const hasUbdFlag = (extra: string, relatives: string) =>
  parseUbdInfo(extra, relatives).hasUbd;

/** Місце дислокації з ООС або «4. Тимчасово прибулі» — зона бойових завдань. */
export const isCombatZoneDislocation = (value: string) => {
  const text = normalizeLooseText(value);
  if (!text || text === "-") return false;
  return (
    /бойов.*(завдан|ді)/.test(text) ||
    /зона.*бойов/.test(text) ||
    /район.*виконан.*бойов/.test(text) ||
    /\bлбз\b/.test(text)
  );
};

const toIsoDateKey = (date: Date) => {
  const year = date.getUTCFullYear();
  const month = String(date.getUTCMonth() + 1).padStart(2, "0");
  const day = String(date.getUTCDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
};

/** Ключ виходу для дедуплікації між «Статус бійців» і ЖБД (ФІО + дата). */
export const normalizeExitDateKey = (value: unknown): string | null => {
  if (value === undefined || value === null || value === "") return null;

  const parsedSerial = tryParseExcelSerialDate(value);
  if (parsedSerial) return toIsoDateKey(parsedSerial);
  if (value instanceof Date) return toIsoDateKey(value);

  const text = String(value).replace(/\u00a0/g, " ").trim();
  if (!text) return null;

  const ukMatch = text.match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{4})$/);
  if (ukMatch) {
    const [, day, month, year] = ukMatch;
    return `${year}-${month.padStart(2, "0")}-${day.padStart(2, "0")}`;
  }

  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) return text;

  if (/^\d{4}-\d{2}-\d{2}T/.test(text)) {
    const parsed = new Date(text);
    if (!Number.isNaN(parsed.getTime())) return toIsoDateKey(parsed);
  }

  return null;
};

export const mergeExitCountMaps = (
  ...sources: Array<Map<string, Set<string>> | undefined | null>
) => {
  const merged = new Map<string, Set<string>>();
  for (const source of sources) {
    if (!source) continue;
    for (const [key, stamps] of source) {
      const set = merged.get(key) ?? new Set<string>();
      for (const stamp of stamps) set.add(stamp);
      merged.set(key, set);
    }
  }
  return merged;
};

export const lookupExitStampSet = (
  map: Map<string, Set<string>>,
  normalizedName: string,
  shortName: string,
) =>
  map.get(normalizedName) ??
  map.get(shortName) ??
  new Set<string>();

/** ≥1 дата виходу на аркуші ранкового «Статус бійців» (до об’єднання з ЖБД). */
export const hasMorningStatusExitRows = (person: {
  morningExitCount: number;
}) => person.morningExitCount > 0;

export const morningStatusExitMatchLabel = (person: {
  morningExitCount: number;
}) => (hasMorningStatusExitRows(person) ? "так" : "ні");

export type UbdRosterStatus = "submitted" | "notSubmitted";

export type UbdRosterMatchRecord = {
  name: string;
  rnokpp: string;
  status: UbdRosterStatus;
  sheetName: string;
};

export type UbdRosterLookup = {
  byFull: Map<string, UbdRosterMatchRecord>;
  byShort: Map<string, UbdRosterMatchRecord[]>;
  byRnokpp: Map<string, UbdRosterMatchRecord>;
  all: UbdRosterMatchRecord[];
};

export const buildUbdRosterLookup = (
  records: UbdRosterMatchRecord[],
): UbdRosterLookup => {
  const byFull = new Map<string, UbdRosterMatchRecord>();
  const byShort = new Map<string, UbdRosterMatchRecord[]>();
  const byRnokpp = new Map<string, UbdRosterMatchRecord>();

  for (const record of records) {
    const full = normalizePersonName(record.name);
    const short = shortPersonName(record.name);
    if (full) byFull.set(full, record);
    if (short) {
      const list = byShort.get(short) ?? [];
      list.push(record);
      byShort.set(short, list);
    }
    const rnokpp = normalizeRnokpp(record.rnokpp);
    if (rnokpp) byRnokpp.set(rnokpp, record);
  }

  return { byFull, byShort, byRnokpp, all: records };
};

export const namesLikelySamePerson = (left: string, right: string) => {
  const a = normalizePersonName(left).split(" ").filter(Boolean);
  const b = normalizePersonName(right).split(" ").filter(Boolean);
  if (a.length < 2 || b.length < 2) return false;
  if (a[0] !== b[0] || a[1] !== b[1]) return false;
  if (a.length >= 3 && b.length >= 3) {
    const ap = a[2].slice(0, 3);
    const bp = b[2].slice(0, 3);
    return ap.startsWith(bp) || bp.startsWith(ap);
  }
  return true;
};

const pickLikelySamePerson = (
  name: string,
  candidates: UbdRosterMatchRecord[],
) => {
  const hits = candidates.filter((record) => namesLikelySamePerson(name, record.name));
  if (hits.length === 1) return hits[0];
  return undefined;
};

/** Зіставлення з реєстром УБД: ПІБ → РНОКПП → нечітке (прізвище + ім’я + по батькові). */
export const matchUbdRosterRecord = (
  name: string,
  rnokpp: string,
  lookup: UbdRosterLookup,
  altNames: string[] = [],
): UbdRosterMatchRecord | undefined => {
  const candidates = [name, ...altNames].filter(Boolean);
  const ipn = normalizeRnokpp(rnokpp);
  if (ipn) {
    const byIpn = lookup.byRnokpp.get(ipn);
    if (byIpn) return byIpn;
  }

  for (const candidate of candidates) {
    const full = normalizePersonName(candidate);
    const direct = lookup.byFull.get(full);
    if (direct) return direct;

    const short = shortPersonName(candidate);
    const shortHits = lookup.byShort.get(short);
    if (shortHits?.length === 1) return shortHits[0];
    if (shortHits && shortHits.length > 1) {
      const likely = pickLikelySamePerson(candidate, shortHits);
      if (likely) return likely;
    }

    const fuzzyHits = lookup.all.filter((record) =>
      namesLikelySamePerson(candidate, record.name),
    );
    if (fuzzyHits.length === 1) return fuzzyHits[0];
  }

  return undefined;
};

/** Підсумкова кількість виходів: реєстр УБД / статичний override = точно виконували (мін. 1). */
export const resolveCombatExitCount = (
  mergedExitCount: number,
  ubdRosterStatus: UbdRosterStatus | null,
  combatDutyEvidence: string[],
  isTransiter: boolean,
  staticCombatExitOverride = false,
) => {
  if (isTransiter) return 0;
  if (mergedExitCount > 0) return mergedExitCount;
  if (ubdRosterStatus || staticCombatExitOverride) return 1;
  return combatDutyEvidence.length > 0 ? 1 : 0;
};

export const effectiveExitCount = (
  morningExitCount: number,
  combatDutyEvidence: string[],
) => {
  if (morningExitCount > 0) return morningExitCount;
  return combatDutyEvidence.length > 0 ? 1 : 0;
};

export const hasIdpFlag = (extra: string, relatives: string, birthPlace: string) => {
  const text = normalizeLooseText(`${extra} ${relatives} ${birthPlace}`);
  return /впо|внутрішн[ьо]*\s*переселен|переселен/.test(text);
};

/** Колонка «В якому підрозділі» — «ТРАНЗИТЕР». */
export const isTransiterDestination = (destination: string) => {
  const normalized = normalizeLooseText(destination);
  return (
    normalized === "транзитер" ||
    normalized === "тринзитер" ||
    normalized.includes("транзитер")
  );
};

/** Колонка «В якому підрозділі» — РРЕБ / РЕБ. */
export const isRrebDestination = (destination: string) => {
  const normalized = normalizeLooseText(destination);
  return (
    normalized === "рреб" ||
    normalized === "реб" ||
    normalized === "рота реб" ||
    normalized === "рота рреб" ||
    /\bрреб\b/.test(normalized) ||
    /\bрота\s+р?еб\b/.test(normalized)
  );
};

/** Колонка «В якому підрозділі» — ППО. */
export const isPpoDestination = (destination: string) => {
  const normalized = normalizeLooseText(destination);
  return normalized === "ппо" || /\bппо\b/.test(normalized);
};

/** Офіцери для аркуша «Офіцери» і рядків виходів: не транзитер. */
export const isOfficerForLbExits = (person: {
  rankGroup: RankGroup;
  isTransiter: boolean;
}) => person.rankGroup === "officer" && !person.isTransiter;

/** Рядки виходів (виконували) — без транзитерів. */
export const countsInExitMetrics = (person: { isTransiter: boolean }) =>
  !person.isTransiter;

/** «Не виконували»: без транзитерів, відком. за межі ПБ, СЗЧ, лікування, поранення. */
export const countsInNoExitsList = (person: {
  isTransiter: boolean;
  morningStatus: string;
  morningAbsenceNotes?: string;
}) => {
  if (person.isTransiter) return false;
  const status = normalizeLooseText(person.morningStatus);
  const notes = normalizeLooseText(person.morningAbsenceNotes ?? "");
  if (status.includes("відком") && status.includes("за межі")) return false;
  if (status.includes("сзч") || status.includes("не в сторою")) return false;
  if (status.includes("лікуван") || notes.includes("лікуван")) return false;
  if (
    status.includes("поранен") ||
    notes.includes("поранен") ||
    notes.includes("по пораненню")
  ) {
    return false;
  }
  return true;
};

export const exitBandOf = (count: number): ExitBand => {
  if (count <= 0) return "none";
  if (count < 5) return "1-4";
  if (count <= 10) return "5-10";
  if (count <= 15) return "11-15";
  if (count <= 20) return "16-20";
  if (count <= 25) return "21-25";
  if (count <= 30) return "26-30";
  return "30+";
};

export const EXIT_BAND_LABELS: Record<ExitBand, string> = {
  none: "Не виконували",
  "1-4": "До 5ти виходів (завдань на ЛБЗ)",
  "5-10": "5–10 виходів",
  "11-15": "11–15 виходів",
  "16-20": "16–20 виходів",
  "21-25": "21–25 виходів",
  "26-30": "26–30 виходів",
  "30+": "30 і більше",
};

export const exitBandLabel = (band: ExitBand) => EXIT_BAND_LABELS[band] ?? band;

export const isPresentStatus = (status: string) => {
  const text = normalizeLooseText(status);
  return text === "в строю" || text === "в сторою";
};

export const isDispositionStatus = (status: string) => {
  const text = normalizeLooseText(status);
  return (
    /розпорядж|л[іi]куван|в[іi]дком|в[іi]дряджен|в[іi]дпустк|навчан/.test(text) &&
    !isPresentStatus(text)
  );
};

export const isOnBooksStatus = (status: string) => {
  const text = normalizeLooseText(status);
  if (!text) return true;
  return !/сзч|загибл|зникл[іi]\s*безв[іi]ст|200\b/.test(text);
};
