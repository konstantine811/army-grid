import { describe, expect, it } from "vitest";
import {
  ANKETA_ABSENT_QUESTIONNAIRE_VALUE,
  ANKETA_ADDITIONAL_INFO_EMPTY_TEMPLATE,
  ANKETA_RELATIVES_EMPTY_TEMPLATE,
} from "./anketaGaps";
import { createEmptyAnketaRow } from "./anketaSheet";
import {
  buildAnketaEmptyGapColumnDescriptors,
  buildAnketaGapColumnDescriptors,
  buildAnketaOcrPersonHint,
  buildAnketaOcrProposals,
  buildAdditionalInfoFromOcrFields,
  buildRelativesFromOcrFields,
  formatAdditionalInfoBlock,
  formatIdDocumentIssuingMeta,
  formatIdDocumentNameWithIssuing,
  mapOcrFieldsToAnketaValues,
} from "./anketaQuestionnaireOcr";

describe("anketaQuestionnaireOcr", () => {
  it("maps OCR fields to anketa columns", () => {
    const mapped = mapOcrFieldsToAnketaValues([
      { key: "rank", label: "Звання", value: "солдат", confidence: "high" },
      { key: "rnokpp", label: "РНОКПП", value: "1234567890", confidence: "high" },
      { key: "phones", label: "Телефони", value: "0501234567", confidence: "medium" },
    ]);
    expect(mapped.rank?.value).toBe("солдат");
    expect(mapped.rnokpp?.value).toBe("1234567890");
    expect(mapped.additionalInfo?.value).toBe(
      "УБД: немає\nтел: 0501234567",
    );
  });

  it("builds additional info with phone and ubd number", () => {
    const mapped = mapOcrFieldsToAnketaValues([
      { key: "phones", label: "Телефони", value: "0671234567", confidence: "high" },
      {
        key: "ubdNumber",
        label: "УБД",
        value: "КВ-009871",
        confidence: "high",
      },
    ]);
    expect(mapped.additionalInfo?.value).toBe(
      "УБД: КВ-009871\nтел: 0671234567",
    );
  });

  it("normalizes additional info block with absent ubd", () => {
    expect(
      formatAdditionalInfoBlock({ phone: "0501112233", ubd: "" }),
    ).toBe("УБД: немає\nтел: 0501112233");
    expect(
      buildAdditionalInfoFromOcrFields([
        {
          key: "additionalInfo",
          label: "Додаткова інформація",
          value: "тел: 0509998877\nУБД:",
          confidence: "high",
        },
      ])?.value,
    ).toBe("УБД: немає\nтел: 0509998877");
  });

  it("maps passport and military id scans to anketa document columns", () => {
    const mapped = mapOcrFieldsToAnketaValues([
      { key: "pageType", label: "Тип", value: "passport", confidence: "high" },
      {
        key: "passportSeries",
        label: "Серія",
        value: "АА",
        confidence: "high",
      },
      {
        key: "passportNumber",
        label: "Номер",
        value: "123456",
        confidence: "high",
      },
      { key: "sex", label: "Стать", value: "чоловіча", confidence: "high" },
      {
        key: "militaryIdSeries",
        label: "Серія ВК",
        value: "АА",
        confidence: "medium",
      },
      {
        key: "militaryIdNumber",
        label: "Номер ВК",
        value: "987654",
        confidence: "medium",
      },
    ]);
    expect(mapped.idDocumentNumber?.value).toBe("АА 123456");
    expect(mapped.idDocumentName?.value).toBe("Паспорт");
    expect(mapped.sex?.value).toBe("ч");
    expect(mapped.militaryId?.value).toBe("АА 987654");
  });

  it("maps ID card with issuing organ from anketa form fields", () => {
    const mapped = mapOcrFieldsToAnketaValues([
      { key: "pageType", label: "Тип", value: "form", confidence: "high" },
      {
        key: "rnokpp",
        label: "ІПН",
        value: "3376202134",
        confidence: "high",
        source: "анкета",
      },
      {
        key: "passportNumber",
        label: "Номер ID",
        value: "007801707",
        confidence: "high",
        source: "анкета",
      },
      {
        key: "passportIssuedDate",
        label: "Дата видачі",
        value: "21.06.2022",
        confidence: "high",
        source: "анкета",
      },
      {
        key: "passportIssuingOrgan",
        label: "Орган",
        value: "1214",
        confidence: "high",
        source: "анкета",
      },
    ]);
    expect(mapped.rnokpp?.value).toBe("3376202134");
    expect(mapped.idDocumentNumber?.value).toBe("007801707");
    expect(mapped.idDocumentName?.value).toBe(
      "ID-картка, виданий 21.06.2022, 1214",
    );
  });

  it("pads a numeric ID-card number that lost leading zeros", () => {
    const mapped = mapOcrFieldsToAnketaValues([
      {
        key: "idDocumentNumber",
        label: "Документ",
        value: "3266545",
        confidence: "high",
      },
    ]);
    expect(mapped.idDocumentNumber?.value).toBe("003266545");
  });

  it("does not pad a passport number with letters", () => {
    const mapped = mapOcrFieldsToAnketaValues([
      {
        key: "idDocumentNumber",
        label: "Документ",
        value: "FZ303340",
        confidence: "high",
      },
    ]);
    expect(mapped.idDocumentNumber?.value).toBe("FZ303340");
  });

  it("splits legacy combined idDocumentNumber into number and name", () => {
    const mapped = mapOcrFieldsToAnketaValues([
      {
        key: "idDocumentNumber",
        label: "Документ",
        value: "007801707, виданий 21.06.2022",
        confidence: "high",
        source: "анкета",
      },
      {
        key: "passportIssuingOrgan",
        label: "Орган",
        value: "1214 (орган)",
        confidence: "high",
        source: "анкета",
      },
    ]);
    expect(mapped.idDocumentNumber?.value).toBe("007801707");
    expect(mapped.idDocumentName?.value).toBe(
      "ID-картка, виданий 21.06.2022, 1214",
    );
  });

  it("does not treat the first digits of an ID number as issuing organ", () => {
    const mapped = mapOcrFieldsToAnketaValues([
      {
        key: "idDocumentNumber",
        label: "Документ",
        value: "007801707, виданий 21.06.2022",
        confidence: "high",
        source: "анкета",
      },
    ]);
    expect(mapped.idDocumentNumber?.value).toBe("007801707");
    expect(mapped.idDocumentName?.value).toBe("ID-картка, виданий 21.06.2022");
    expect(mapped.idDocumentName?.value).not.toContain("0078");
  });

  it("formats issuing meta and document name helpers", () => {
    expect(
      formatIdDocumentIssuingMeta({
        issuedDate: "21.06.2022",
        issuingOrgan: "1214 (орган)",
      }),
    ).toBe("виданий 21.06.2022, 1214");
    expect(
      formatIdDocumentNameWithIssuing(
        "ID-картка",
        "виданий 21.06.2022, 1214",
      ),
    ).toBe("ID-картка, виданий 21.06.2022, 1214");
  });

  it("builds relatives block with phones from OCR parts", () => {
    const mapped = mapOcrFieldsToAnketaValues([
      {
        key: "familyStatus",
        label: "Сімейний стан",
        value: "одружений",
        confidence: "high",
      },
      {
        key: "relativeMother",
        label: "Мати",
        value: "Іванова Олена Петрівна, тел. 0501234567",
        confidence: "high",
      },
      {
        key: "relativeFather",
        label: "Батько",
        value: "Іванов Петро Іванович",
        confidence: "medium",
      },
      {
        key: "relativePhones",
        label: "Телефони",
        value: "0679876543",
        confidence: "medium",
      },
    ]);
    expect(mapped.relatives?.value).toContain("Сімейний стан: одружений");
    expect(mapped.relatives?.value).toContain("Мати: Іванова");
    expect(mapped.relatives?.value).toContain("+380501234567");
    expect(mapped.relatives?.value).toContain("+380679876543");
  });

  it("prefers full relatives field from gap fill", () => {
    const built = buildRelativesFromOcrFields([
      {
        key: "relatives",
        label: "Родичі",
        value: "Сімейний стан: не одружений\nМати: Коваль, тел. 0501112233",
        confidence: "high",
      },
      {
        key: "familyStatus",
        label: "Сімейний стан",
        value: "одружений",
        confidence: "high",
      },
    ]);
    expect(built?.value).toContain("+380501112233");
    expect(built?.value).toContain("Сімейний стан: не одружений");
    expect(built?.value).toContain("Мати: Коваль, тел: +380501112233");
  });

  it("maps tax id scan to rnokpp", () => {
    const mapped = mapOcrFieldsToAnketaValues([
      {
        key: "taxId",
        label: "ІПН",
        value: "1234567890",
        confidence: "high",
        source: "РНОКПП",
      },
    ]);
    expect(mapped.rnokpp?.value).toBe("1234567890");
    expect(mapped.rnokpp?.source).toBe("РНОКПП");
  });

  it("prefers tax id document over passport digits for rnokpp", () => {
    const mapped = mapOcrFieldsToAnketaValues([
      {
        key: "passportNumber",
        label: "Номер",
        value: "1234567890",
        confidence: "high",
        source: "паспорт",
      },
      {
        key: "taxId",
        label: "ІПН",
        value: "3142223156",
        confidence: "high",
        source: "РНОКПП",
      },
    ]);
    expect(mapped.rnokpp?.value).toBe("3142223156");
  });

  it("builds proposals only for empty selected gap columns", () => {
    const row = {
      ...createEmptyAnketaRow(5),
      fullName: "ІВАНОВ Іван",
      rank: "сержант",
    };
    const proposals = buildAnketaOcrProposals(
      [
        { key: "rank", label: "Звання", value: "солдат", confidence: "high" },
        { key: "rnokpp", label: "РНОКПП", value: "1234567890", confidence: "high" },
        { key: "birthDate", label: "Дата", value: "01.01.1990", confidence: "low" },
      ],
      row,
      ["rank", "rnokpp", "birthDate"],
    );
    expect(proposals.map((item) => item.columnId)).toEqual(["rnokpp", "birthDate"]);
    expect(proposals.find((item) => item.columnId === "birthDate")?.selected).toBe(
      false,
    );
  });

  it("builds selected and empty gap column descriptors", () => {
    const row = {
      ...createEmptyAnketaRow(5),
      fullName: "ІВАНОВ Іван",
      rank: "сержант",
    };
    expect(buildAnketaGapColumnDescriptors(["rank", "rnokpp", "birthDate"])).toEqual([
      { key: "rank", header: "Звання" },
      { key: "rnokpp", header: "РНОКПП" },
      { key: "birthDate", header: "Дата народження" },
    ]);
    expect(
      buildAnketaEmptyGapColumnDescriptors(row, ["rank", "rnokpp", "birthDate"]),
    ).toEqual([
      { key: "rnokpp", header: "РНОКПП" },
      { key: "birthDate", header: "Дата народження" },
    ]);
  });

  it("maps direct anketa column keys from gap fill response", () => {
    const mapped = mapOcrFieldsToAnketaValues([
      {
        key: "rnokpp",
        label: "РНОКПП",
        value: "3142223156",
        confidence: "high",
        source: "РНОКПП",
      },
      {
        key: "birthDate",
        label: "Дата народження",
        value: "01.01.1990",
        confidence: "high",
        source: "анкета",
      },
    ]);
    expect(mapped.rnokpp?.value).toBe("3142223156");
    expect(mapped.birthDate?.value).toBe("01.01.1990");
  });

  it("replaces «дані відсутні» with a military ID from the ticket scan", () => {
    const row = {
      ...createEmptyAnketaRow(5),
      fullName: "СМОЛЬНІКОВ Володимир Володимирович",
      militaryId: ANKETA_ABSENT_QUESTIONNAIRE_VALUE,
      relatives: "",
    };
    expect(
      buildAnketaEmptyGapColumnDescriptors(row, ["militaryId", "relatives"]).map(
        (item) => item.key,
      ),
    ).toEqual(["militaryId", "relatives"]);
    expect(
      buildAnketaEmptyGapColumnDescriptors(row, ["militaryId"])[0],
    ).toMatchObject({
      key: "militaryId",
      currentValue: ANKETA_ABSENT_QUESTIONNAIRE_VALUE,
    });

    const proposals = buildAnketaOcrProposals(
      [
        {
          key: "militaryId",
          label: "Військовий квиток",
          value: "Серія АТ № 225674",
          confidence: "high",
          source: "військовий квиток",
        },
      ],
      row,
      ["militaryId", "relatives"],
    );
    expect(proposals.find((item) => item.columnId === "militaryId")).toMatchObject({
      value: "АТ 225674",
      selected: true,
    });
  });

  it("does not re-propose «дані відсутні» when OCR found no military ID", () => {
    const row = {
      ...createEmptyAnketaRow(5),
      militaryId: ANKETA_ABSENT_QUESTIONNAIRE_VALUE,
    };
    expect(
      buildAnketaOcrProposals([], row, ["militaryId"]).map((item) => item.columnId),
    ).toEqual([]);
  });

  it("treats empty relatives and additionalInfo templates as gap columns for OCR", () => {
    const row = {
      ...createEmptyAnketaRow(5),
      fullName: "КОВАЛЬЧУК Роман Петрович",
      relatives: ANKETA_RELATIVES_EMPTY_TEMPLATE,
      additionalInfo: ANKETA_ADDITIONAL_INFO_EMPTY_TEMPLATE,
      education: "середня",
    };
    expect(
      buildAnketaEmptyGapColumnDescriptors(row, [
        "relatives",
        "additionalInfo",
        "education",
      ]).map((item) => item.key),
    ).toEqual(["relatives", "additionalInfo"]);
    const proposals = buildAnketaOcrProposals(
      [
        {
          key: "relatives",
          label: "Дані про родичів",
          value: "Сімейний стан: одружений",
          confidence: "high",
        },
      ],
      row,
      ["relatives", "additionalInfo", "education"],
    );
    expect(proposals.map((item) => item.columnId)).toEqual([
      "relatives",
      "additionalInfo",
    ]);
  });

  it("ignores cells that already contain text for OCR gap fill", () => {
    const row = {
      ...createEmptyAnketaRow(5),
      fullName: "КОВАЛЬЧУК Роман Петрович",
      rnokpp: "забув",
      birthDate: "01.01.1990",
    };
    expect(
      buildAnketaEmptyGapColumnDescriptors(row, ["rnokpp", "birthDate"]).map(
        (item) => item.key,
      ),
    ).toEqual([]);
    const proposals = buildAnketaOcrProposals(
      [{ key: "rnokpp", label: "РНОКПП", value: "3142223156", confidence: "high" }],
      row,
      ["rnokpp", "birthDate"],
    );
    expect(proposals).toEqual([]);
  });

  it("writes questionnaire status markers instead of «дані відсутні»", () => {
    const row = {
      ...createEmptyAnketaRow(5),
      fullName: "КОВАЛЬЧУК Роман Петрович",
      rnokpp: "",
    };
    const proposals = buildAnketaOcrProposals(
      [{ key: "rnokpp", label: "РНОКПП", value: "забув", confidence: "high" }],
      row,
      ["rnokpp"],
    );
    expect(proposals).toEqual([
      expect.objectContaining({ columnId: "rnokpp", value: "забув" }),
    ]);
  });

  it("defaults empty gap columns without OCR values to «дані відсутні»", () => {
    const row = {
      ...createEmptyAnketaRow(5),
      fullName: "КОВАЛЬЧУК Роман Петрович",
      rnokpp: "",
      birthDate: "01.01.1990",
    };
    const proposals = buildAnketaOcrProposals(
      [{ key: "education", label: "Освіта", value: "середня", confidence: "high" }],
      row,
      ["rnokpp", "education"],
    );
    expect(proposals.map((item) => item.columnId)).toEqual(["rnokpp", "education"]);
    expect(proposals.find((item) => item.columnId === "rnokpp")).toMatchObject({
      value: ANKETA_ABSENT_QUESTIONNAIRE_VALUE,
      selected: false,
    });
  });

  it("selects focused gap column even when confidence is low", () => {
    const row = {
      ...createEmptyAnketaRow(5),
      fullName: "ІВАНОВ Іван",
    };
    const proposals = buildAnketaOcrProposals(
      [
        { key: "birthDate", label: "Дата", value: "01.01.1990", confidence: "low" },
        { key: "rnokpp", label: "РНОКПП", value: "1234567890", confidence: "high" },
      ],
      row,
      ["rnokpp", "birthDate"],
      { focusedColumnId: "birthDate" },
    );
    expect(proposals[0]?.columnId).toBe("birthDate");
    expect(proposals.find((item) => item.columnId === "birthDate")?.selected).toBe(
      true,
    );
  });

  it("reorders scattered birth place and address into settlement, district, region", () => {
    const mapped = mapOcrFieldsToAnketaValues([
      {
        key: "birthPlace",
        label: "Місце народження",
        value: "Дніпропетровська обл Софіївський район смт Софіївка",
        confidence: "high",
      },
      {
        key: "actualAddress",
        label: "Адреса",
        value: "Дніпропетровська обл Софіївський район смт Софіївка +380662033264",
        confidence: "high",
      },
    ]);
    expect(mapped.birthPlace?.value).toBe(
      "смт Софіївка, Софіївський р-н, Дніпропетровська обл",
    );
    expect(mapped.location?.value).toBe(
      "смт Софіївка, Софіївський р-н, Дніпропетровська обл, тел: +380662033264",
    );
  });

  it("puts address first and phone last in additional info", () => {
    expect(
      buildAdditionalInfoFromOcrFields([
        {
          key: "additionalInfo",
          label: "Додаткова інформація",
          value:
            "Дніпропетровська обл Софіївський район смт Софіївка +380662033264",
          confidence: "high",
        },
        {
          key: "ubdNumber",
          label: "УБД",
          value: "немає",
          confidence: "medium",
        },
      ])?.value,
    ).toBe(
      "смт Софіївка, Софіївський р-н, Дніпропетровська обл\nУБД: немає\nтел: +380662033264",
    );
  });

  it("sends known person data in the OCR hint so handwriting can be verified", () => {
    const hint = buildAnketaOcrPersonHint({
      ...createEmptyAnketaRow(5),
      fullName: "КРАВЧЕНКО Олександр Вікторович",
      rank: "солдат",
      externalId: "10034",
      rnokpp: "3468406676",
      birthDate: "15.05.1994",
      additionalInfo: "забув",
    });
    expect(hint).toContain("КРАВЧЕНКО Олександр Вікторович");
    expect(hint).toContain("Звання: солдат");
    expect(hint).toContain("ID: 10034");
    expect(hint).toContain("РНОКПП: 3468406676");
    expect(hint).toContain("Дата народження: 15.05.1994");
    expect(hint).not.toContain("забув");
    expect(hint).toContain("с./м./смт Назва, Район р-н, Область обл");
  });

  it("maps RTCC mobilization line to conscripted when/by instead of missing data", () => {
    const mapped = mapOcrFieldsToAnketaValues([
      {
        key: "conscriptedBy",
        label: "Ким призваний",
        value: "Калинівський РТЦК та СП 14.11.2023",
        confidence: "high",
      },
      {
        key: "serviceType",
        label: "Вид служби",
        value: "Мобілізований",
        confidence: "high",
      },
      {
        key: "militaryId",
        label: "Військовий квиток",
        value: "АГ 040571",
        confidence: "high",
      },
    ]);
    expect(mapped.conscriptedWhen?.value).toBe("14.11.2023");
    expect(mapped.conscriptedBy?.value).toBe("Калинівський РТЦК та СП");
    expect(mapped.serviceType?.value).toBe("мобілізований");
    expect(mapped.militaryId?.value).toBe("АГ 040571");
    expect(mapped.arrivedFrom).toBeUndefined();
  });

  it("does not treat a recruiting-office line as civilian workplace", () => {
    const mapped = mapOcrFieldsToAnketaValues([
      {
        key: "work",
        label: "Робота",
        value: "Калинівський РТЦК та СП 14.11.2023",
        confidence: "medium",
      },
    ]);
    expect(mapped.arrivedFrom).toBeUndefined();
    expect(mapped.conscriptedWhen?.value).toBe("14.11.2023");
    expect(mapped.conscriptedBy?.value).toBe("Калинівський РТЦК та СП");
  });

  it("proposes conscription fields instead of «дані відсутні»", () => {
    const row = {
      ...createEmptyAnketaRow(5),
      fullName: "КРАВЧЕНКО Олександр Вікторович",
    };
    const proposals = buildAnketaOcrProposals(
      [
        {
          key: "conscription",
          label: "Мобілізація",
          value: "Калинівський РТЦК та СП 14.11.2023",
          confidence: "high",
        },
      ],
      row,
      ["conscriptedWhen", "conscriptedBy"],
    );
    expect(proposals.find((item) => item.columnId === "conscriptedWhen")).toMatchObject({
      value: "14.11.2023",
      selected: true,
    });
    expect(proposals.find((item) => item.columnId === "conscriptedBy")).toMatchObject({
      value: "Калинівський РТЦК та СП",
      selected: true,
    });
  });
});
