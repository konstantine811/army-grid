import { useCallback, useEffect, useMemo, useState } from "react";
import {
  Alert,
  Box,
  Button,
  MenuItem,
  Stack,
  TextField,
  Typography,
} from "@/components/sci/SciPrimitives";
import {
  ArticleOutlinedIcon,
  DeleteOutlineOutlinedIcon,
  PictureAsPdfOutlinedIcon,
} from "@/components/sci/icons";
import dayjs from "dayjs";
import "dayjs/locale/uk";
import { api, type BackendPersonDocument } from "../../api";
import { buildDocumentRoute } from "../../app/navigation";
import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import {
  buildPersonSummary,
  getPersonFieldValue,
} from "../personnel/personnelUtils";

dayjs.locale("uk");

type DocumentMode = "default" | "salaryPowerAttorney" | "ubdReport";

type DefaultDocumentFieldKey = "pib" | "rank" | "unit" | "date";
type DefaultDocumentFields = Record<DefaultDocumentFieldKey, string>;

type SalaryDocumentFields = {
  fullName: string;
  rnokpp: string;
  iban: string;
  bankMfo: string;
  bankName: string;
  rank: string;
  date: string;
  commander: string;
  personnelChief: string;
  folderName: string;
};

type SalaryWorkflowState = {
  completed: Record<string, boolean>;
  currentStatus: string;
  accountFileName: string;
  signedScanFileName: string;
  mergedPdfFileName: string;
};

type UbdReportFields = {
  commander: string;
  fullName: string;
  rank: string;
  staffPosition: string;
  birthDate: string;
  rnokpp: string;
  taskPeriod: string;
  taskPlace: string;
  basis: string;
  battalionCommander: string;
  approvalOfficer: string;
  date: string;
  folderName: string;
};

type UbdScanFile = {
  id: string;
  name: string;
  type: string;
  dataUrl: string;
  uploadedAt: string;
};

type DocumentFiles = {
  ubdScans?: UbdScanFile[];
};

type DocumentRouteState = {
  requestedPersonId: string;
  requestedDocumentType: string;
  requestedDocumentId: string;
  isPersonDocumentMode: boolean;
};

const SALARY_WORKFLOW_STORAGE_KEY = "army-grid:salary-power-attorney-workflow";

const salaryWorkflowSteps = [
  {
    key: "account",
    title: "Скрін реквізитів рахунку та ІНН",
  },
  {
    key: "document",
    title: "Заповнили довіреність",
  },
  {
    key: "print",
    title: "Роздрукували у 2 екземплярах",
  },
  {
    key: "sign",
    title: "Надали підпис службовцю та командиру",
  },
  {
    key: "scan",
    title: "Скан копії та загальний PDF",
  },
  {
    key: "ready",
    title: "Готово до відправки",
  },
  {
    key: "sent",
    title: "Відправили",
  },
];

const ukrainianBanks = [
  { mfo: "305299", name: "АТ КБ «ПРИВАТБАНК»" },
  { mfo: "300465", name: "АТ «ОЩАДБАНК»" },
  { mfo: "322001", name: "АТ «УНІВЕРСАЛ БАНК» / MONOBANK" },
  { mfo: "380805", name: "АТ «РАЙФФАЙЗЕН БАНК»" },
  { mfo: "351005", name: "АТ «УКРСИББАНК»" },
  { mfo: "334851", name: "АТ «ПУМБ»" },
  { mfo: "300346", name: "АТ «СЕНС БАНК»" },
  { mfo: "307770", name: "АТ «АКЦЕНТ-БАНК»" },
  { mfo: "320478", name: "АБ «УКРГАЗБАНК»" },
  { mfo: "300528", name: "АТ «ОТП БАНК»" },
  { mfo: "300614", name: "АТ «КРЕДІ АГРІКОЛЬ БАНК»" },
  { mfo: "325365", name: "АТ «КРЕДОБАНК»" },
  { mfo: "328209", name: "АБ «ПІВДЕННИЙ»" },
  { mfo: "339500", name: "АТ «ТАСКОМБАНК»" },
  { mfo: "328168", name: "АТ «МТБ БАНК»" },
  { mfo: "380526", name: "АТ «КБ «ГЛОБУС»" },
  { mfo: "custom", name: "Інший банк" },
];

const createDefaultFields = (
  summary: ReturnType<typeof buildPersonSummary>,
): DefaultDocumentFields => ({
  pib: summary.name !== "Особа не вибрана" ? summary.name : "",
  rank: summary.rank || "",
  unit: summary.location || "",
  date: dayjs().format("DD.MM.YYYY"),
});

const normalizeIban = (value: string) =>
  String(value ?? "")
    .replace(/\s+/g, "")
    .toUpperCase();

const formatIban = (value: string) =>
  normalizeIban(value).replace(/(.{4})/g, "$1 ").trim();

const extractIban = (row: EjournalPreviewRow | null) => {
  const direct =
    getPersonFieldValue(row, ["iban"]) ||
    getPersonFieldValue(row, ["iban", "рахунок"]) ||
    getPersonFieldValue(row, ["iban", "банк"]) ||
    getPersonFieldValue(row, ["рахунок"]) ||
    getPersonFieldValue(row, ["банківський", "рахунок"]) ||
    "";
  const directMatch = direct.match(/UA\s*\d{2}(?:\s*\d){25}/i);
  if (directMatch) return normalizeIban(directMatch[0]);

  const allText = Object.values(row ?? {})
    .map((value) => String(value ?? ""))
    .join(" ");
  const anyMatch = allText.match(/UA\s*\d{2}(?:\s*\d){25}/i);
  return anyMatch ? normalizeIban(anyMatch[0]) : "";
};

const getBankMfoFromIban = (iban: string) => {
  const normalized = normalizeIban(iban);
  return /^UA\d{27}$/.test(normalized) ? normalized.slice(4, 10) : "";
};

const findBankByMfo = (mfo: string) =>
  ukrainianBanks.find((bank) => bank.mfo === mfo);

const surnameUpperName = (fullName: string) => {
  const parts = fullName.trim().split(/\s+/).filter(Boolean);
  if (!parts.length) return "";
  return [parts[0].toUpperCase(), ...parts.slice(1)].join(" ");
};

const safeFilePart = (value: string) =>
  String(value || "document")
    .replace(/[\\/:*?"<>|]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();

const createSalaryFields = (
  row: EjournalPreviewRow | null,
  summary: ReturnType<typeof buildPersonSummary>,
): SalaryDocumentFields => {
  const iban = extractIban(row);
  const bankMfo = getBankMfoFromIban(iban);
  const bank = findBankByMfo(bankMfo);
  const fullName =
    summary.name !== "Особа не вибрана" ? surnameUpperName(summary.name) : "";

  return {
    fullName,
    rnokpp: summary.rnokpp || "",
    iban,
    bankMfo: bankMfo || "custom",
    bankName: bank?.name || "",
    rank: summary.rank || "",
    date: dayjs().format("DD.MM.YYYY"),
    commander: "Командиру 1 піхотного батальйону військової частини А4862",
    personnelChief: "молодший сержант________________      Віталій БОНДАР",
    folderName: fullName ? `ГЗ довіреність - ${fullName}` : "ГЗ довіреність",
  };
};

const createUbdFields = (
  row: EjournalPreviewRow | null,
  summary: ReturnType<typeof buildPersonSummary>,
): UbdReportFields => {
  const fullName = summary.name !== "Особа не вибрана" ? summary.name : "";
  const staffPosition =
    getPersonFieldValue(row, ["посада"]) ||
    getPersonFieldValue(row, ["штатна", "посада"]) ||
    getPersonFieldValue(row, ["чим", "займається"]) ||
    "";

  return {
    commander: "Командиру військової частини А4862",
    fullName,
    rank: summary.rank || "",
    staffPosition,
    birthDate: summary.birthDate || "",
    rnokpp: summary.rnokpp || "",
    taskPeriod: "",
    taskPlace: "",
    basis: "Бойове розпорядження командира 425 ОШП «СКЕЛЯ» №____ від __.__.2026",
    battalionCommander:
      "Командир 1 піхотного батальйону військової частини А4862\nстарший лейтенант Єгор СИДОРЕНКО",
    approvalOfficer:
      "Тимчасово виконуючий обов’язки командира військової частини А4862\nкапітан Олег АДАМОВ",
    date: dayjs().format("DD.MM.YYYY"),
    folderName: fullName ? `УБД рапорт - ${fullName}` : "УБД рапорт",
  };
};

const createEmptyWorkflow = (): SalaryWorkflowState => ({
  completed: {},
  currentStatus: salaryWorkflowSteps[0]?.key ?? "account",
  accountFileName: "",
  signedScanFileName: "",
  mergedPdfFileName: "",
});

const mergeSalaryFields = (
  defaults: SalaryDocumentFields,
  value: unknown,
): SalaryDocumentFields =>
  value && typeof value === "object" && !Array.isArray(value)
    ? { ...defaults, ...(value as Partial<SalaryDocumentFields>) }
    : defaults;

const mergeSalaryWorkflow = (value: unknown): SalaryWorkflowState =>
  value && typeof value === "object" && !Array.isArray(value)
    ? { ...createEmptyWorkflow(), ...(value as Partial<SalaryWorkflowState>) }
    : createEmptyWorkflow();

const mergeUbdFields = (
  defaults: UbdReportFields,
  value: unknown,
): UbdReportFields =>
  value && typeof value === "object" && !Array.isArray(value)
    ? { ...defaults, ...(value as Partial<UbdReportFields>) }
    : defaults;

const mergeDocumentFiles = (value: unknown): DocumentFiles =>
  value && typeof value === "object" && !Array.isArray(value)
    ? { ...(value as DocumentFiles) }
    : {};

const salaryDocumentTypeLabel = (type: string) =>
  type === "salaryPowerAttorney"
    ? "Довіреність зарплати"
    : type === "ubdReport"
      ? "Рапорт на УБД"
      : type;

const workflowStatusLabel = (status?: string | null) =>
  salaryWorkflowSteps.find((step) => step.key === status)?.title ||
  status ||
  "статус не заданий";

const getDocumentProgressPercent = (document: BackendPersonDocument) => {
  const workflow = mergeSalaryWorkflow(document.workflow);
  const index = salaryWorkflowSteps.findIndex(
    (step) => step.key === (document.status || workflow.currentStatus),
  );
  const resolvedIndex = Math.max(0, index);
  return Math.round(((resolvedIndex + 1) / salaryWorkflowSteps.length) * 100);
};

const getDocumentPersonName = (document: BackendPersonDocument) => {
  const fields = document.fields || {};
  const value =
    fields.fullName ||
    fields.pib ||
    fields.name ||
    fields["ПІБ"] ||
    fields["ФИО"];
  return typeof value === "string" && value.trim()
    ? value.trim()
    : `ID ${document.personExternalId}`;
};

const getDocumentFileCount = (document: BackendPersonDocument) => {
  if (document.type === "ubdReport") {
    return mergeDocumentFiles(document.files).ubdScans?.length ?? 0;
  }

  const workflow = mergeSalaryWorkflow(document.workflow);
  return [
    workflow.accountFileName,
    workflow.signedScanFileName,
    workflow.mergedPdfFileName,
  ].filter((value) => value.trim()).length;
};

const getDocumentFileSummary = (document: BackendPersonDocument) =>
  document.type === "ubdReport"
    ? `${getDocumentFileCount(document)} сканів`
    : `${getDocumentFileCount(document)} / 3`;

const personWorkflowKey = (
  row: EjournalPreviewRow | null,
  summary: ReturnType<typeof buildPersonSummary>,
) => String(summary.externalId || row?.__dbRowId || summary.name || "unknown");

const loadWorkflowMap = () => {
  try {
    const raw = window.localStorage.getItem(SALARY_WORKFLOW_STORAGE_KEY);
    return raw
      ? (JSON.parse(raw) as Record<string, SalaryWorkflowState>)
      : {};
  } catch {
    return {};
  }
};

const saveWorkflowForPerson = (
  key: string,
  workflow: SalaryWorkflowState,
) => {
  const current = loadWorkflowMap();
  current[key] = workflow;
  window.localStorage.setItem(
    SALARY_WORKFLOW_STORAGE_KEY,
    JSON.stringify(current),
  );
};

const highlightOrPlaceholder = (value: string, placeholder: string) =>
  value.trim() ? <mark>{value}</mark> : <mark>{placeholder}</mark>;

const salaryDocumentHtml = (fields: SalaryDocumentFields) => `
<!doctype html>
<html lang="uk">
<head>
  <meta charset="utf-8" />
  <title>Рапорт ГЗ ${safeFilePart(fields.fullName)}</title>
  <style>
    @page { size: A4; margin: 24mm 20mm; }
    body { font-family: "Times New Roman", serif; color: #000; font-size: 14pt; line-height: 1.35; }
    .to { text-align: right; margin-left: 42%; }
    h1 { margin: 28mm 0 12mm; text-align: center; font-size: 16pt; }
    p { margin: 0 0 9mm; }
    .sign-row { display: grid; grid-template-columns: 1fr 44mm; gap: 12mm; margin-top: 18mm; }
    .finance { margin-top: 18mm; }
    .small { font-size: 12pt; }
  </style>
</head>
<body>
  <div class="to">${fields.commander}</div>
  <h1>Рапорт</h1>
  <p>
    Прошу моє грошове забезпечення перераховувати на мій розрахунковий
    банківський рахунок:
  </p>
  <p><strong>${formatIban(fields.iban) || "UA___________________________"}</strong>, відкритий ${fields.bankName || "________________"}.</p>
  <p>Код РНОКПП <strong>${fields.rnokpp || "____________"}</strong>.</p>
  <p>Довідка з банківськими реквізитами додається.</p>
  <p>${fields.rank || "________________"} ${fields.fullName || "____________________________"}</p>
  <div class="sign-row">
    <span>${fields.date || "__.__.202__"}</span>
    <span>________________</span>
  </div>
  <div class="finance">
    <p>
      Начальнику фінансово-економічної служби – головному бухгалтеру
      військової частини А4862 - прошу прийняти в роботу
    </p>
    <p>Начальник групи персоналу штабу<br />1 піхотного батальйону<br />військової частини А4862</p>
    <p>${fields.personnelChief}</p>
    <p class="small">(підпис)</p>
  </div>
</body>
</html>`;

const ubdReportHtml = (fields: UbdReportFields) => `
<!doctype html>
<html lang="uk">
<head>
  <meta charset="utf-8" />
  <title>Рапорт УБД ${safeFilePart(fields.fullName)}</title>
  <style>
    @page { size: A4; margin: 18mm 18mm; }
    body { font-family: "Times New Roman", serif; color: #000; font-size: 13pt; line-height: 1.28; }
    .to { text-align: right; margin-left: 45%; }
    h1 { margin: 18mm 0 8mm; text-align: center; font-size: 16pt; text-transform: uppercase; }
    table { width: 100%; border-collapse: collapse; margin: 5mm 0 7mm; }
    td { border: 1px solid #000; padding: 3mm; vertical-align: top; }
    .label { width: 42%; font-weight: 700; }
    p { margin: 0 0 5mm; }
    .sign { margin-top: 10mm; white-space: pre-line; }
    .approve { margin-top: 12mm; text-align: right; white-space: pre-line; }
  </style>
</head>
<body>
  <div class="to">${fields.commander}</div>
  <h1>Рапорт</h1>
  <p>
    Прошу Вашого рішення про внесення відомостей до Єдиного державного
    реєстру ветеранів війни щодо зазначеного нижче військовослужбовця,
    який розпочав виконання бойового завдання у районі ведення бойових дій,
    з метою надання статусу учасника бойових дій, а саме:
  </p>
  <table>
    <tr><td class="label">Військове звання</td><td>${fields.rank}</td></tr>
    <tr><td class="label">Прізвище, ім'я, по-батькові</td><td>${fields.fullName}</td></tr>
    <tr><td class="label">Посада згідно штату</td><td>${fields.staffPosition}</td></tr>
    <tr><td class="label">Дата народження</td><td>${fields.birthDate}</td></tr>
    <tr><td class="label">РНОКПП</td><td>${fields.rnokpp}</td></tr>
    <tr><td class="label">Період виконання завдань</td><td>${fields.taskPeriod}</td></tr>
    <tr><td class="label">Місце виконання завдань</td><td>${fields.taskPlace}</td></tr>
  </table>
  <p><strong>Підстава:</strong> ${fields.basis}</p>
  <p><strong>Додаток:</strong> копія паспорта, РНОКПП, дві фотокартки</p>
  <p class="sign">${fields.battalionCommander}</p>
  <p>${fields.date}</p>
  <p class="approve">ЗАТВЕРДЖУЮ
${fields.approvalOfficer}</p>
</body>
</html>`;

const downloadTextFile = (fileName: string, content: string, mimeType: string) => {
  const url = URL.createObjectURL(new Blob([content], { type: mimeType }));
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
};

const readFileAsDataUrl = (file: File) =>
  new Promise<string>((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result || ""));
    reader.onerror = () => reject(reader.error);
    reader.readAsDataURL(file);
  });

const getDocumentRouteState = (): DocumentRouteState => {
  const urlParams = new URLSearchParams(window.location.search);
  const pathPersonId = window.location.pathname.startsWith("/documents/")
    ? decodeURIComponent(window.location.pathname.slice("/documents/".length))
        .split("/")[0]
        ?.trim() || ""
    : "";
  const requestedPersonId =
    pathPersonId ||
    urlParams.get("personId")?.trim() ||
    urlParams.get("externalId")?.trim() ||
    urlParams.get("personRowId")?.trim() ||
    "";

  return {
    requestedPersonId,
    requestedDocumentType: urlParams.get("type") || "",
    requestedDocumentId: urlParams.get("documentId")?.trim() || "",
    isPersonDocumentMode: Boolean(requestedPersonId),
  };
};

export function DocumentsPage({
  onNavigate,
}: {
  onNavigate?: (path: string) => void;
}) {
  const {
    requestedPersonId,
    requestedDocumentType,
    requestedDocumentId,
    isPersonDocumentMode,
  } = getDocumentRouteState();
  const [selectedPerson, setSelectedPerson] = useState<EjournalPreviewRow | null>(
    null,
  );
  const [mode, setMode] = useState<DocumentMode>("default");
  const [fields, setFields] = useState<DefaultDocumentFields>(() =>
    createDefaultFields(buildPersonSummary(null)),
  );
  const [salaryFields, setSalaryFields] = useState<SalaryDocumentFields>(() =>
    createSalaryFields(null, buildPersonSummary(null)),
  );
  const [ubdFields, setUbdFields] = useState<UbdReportFields>(() =>
    createUbdFields(null, buildPersonSummary(null)),
  );
  const [documentFiles, setDocumentFiles] = useState<DocumentFiles>({});
  const [workflow, setWorkflow] =
    useState<SalaryWorkflowState>(createEmptyWorkflow);
  const [personDocuments, setPersonDocuments] = useState<BackendPersonDocument[]>(
    [],
  );
  const [selectedDocumentId, setSelectedDocumentId] = useState("");
  const [documentMessage, setDocumentMessage] = useState("");
  const [isSavingDocument, setIsSavingDocument] = useState(false);
  const [allPersonDocuments, setAllPersonDocuments] = useState<
    BackendPersonDocument[]
  >([]);
  const [isLoadingDocumentJournal, setIsLoadingDocumentJournal] =
    useState(false);

  useEffect(() => {
    const rawPerson = window.localStorage.getItem("army-grid:selected-person");
    const savedMode =
      window.localStorage.getItem("army-grid:selected-document-mode") ||
      (requestedDocumentType === "salary-power-attorney"
        ? "salaryPowerAttorney"
        : "default");
    setMode(
      requestedDocumentType === "ubd-report"
        ? "ubdReport"
        : requestedDocumentType === "salary-power-attorney" ||
            savedMode === "salaryPowerAttorney"
          ? "salaryPowerAttorney"
          : "default",
    );

    if (!rawPerson || !isPersonDocumentMode) return;

    try {
      const person = JSON.parse(rawPerson) as EjournalPreviewRow;
      const nextSummary = buildPersonSummary(person);
      const nextPersonId = String(
        nextSummary.externalId || person.__dbRowId || "",
      );
      if (requestedPersonId && nextPersonId !== requestedPersonId) {
        setDocumentMessage(
          `Не знайшов локальні дані службовця ID ${requestedPersonId}. Відкрийте документ із картки особи.`,
        );
        return;
      }
      setSelectedPerson(person);
      setFields(createDefaultFields(nextSummary));
      setSalaryFields(createSalaryFields(person, nextSummary));
      setUbdFields(createUbdFields(person, nextSummary));
      const key = personWorkflowKey(person, nextSummary);
      setWorkflow(loadWorkflowMap()[key] ?? createEmptyWorkflow());
    } catch {
      setSelectedPerson(null);
    }
  }, [isPersonDocumentMode, requestedDocumentType, requestedPersonId]);

  const summary = useMemo(
    () => buildPersonSummary(selectedPerson),
    [selectedPerson],
  );
  const personExternalId = String(
    requestedPersonId || summary.externalId || selectedPerson?.__dbRowId || "",
  );
  const selectedDocument = useMemo(
    () =>
      personDocuments.find((document) => document.id === selectedDocumentId) ??
      null,
    [personDocuments, selectedDocumentId],
  );
  const workflowKey = useMemo(
    () => personWorkflowKey(selectedPerson, summary),
    [selectedPerson, summary],
  );
  const targetDocumentType: DocumentMode =
    requestedDocumentType === "ubd-report"
      ? "ubdReport"
      : requestedDocumentType === "salary-power-attorney"
        ? "salaryPowerAttorney"
        : mode;

  useEffect(() => {
    if (!isPersonDocumentMode) return;
    if (
      (targetDocumentType !== "salaryPowerAttorney" &&
        targetDocumentType !== "ubdReport") ||
      !personExternalId
    )
      return;

    let cancelled = false;
    const salaryDefaults = createSalaryFields(selectedPerson, summary);
    const ubdDefaults = createUbdFields(selectedPerson, summary);

    const loadDocuments = async () => {
      setDocumentMessage("Завантажую документи службовця...");
      try {
        const documents = await api.listPersonDocuments(personExternalId);
        let nextDocuments = documents;
        let active =
          nextDocuments.find((document) => document.id === requestedDocumentId) ??
          nextDocuments.find(
            (document) => document.type === targetDocumentType,
          ) ??
          null;

        if (cancelled) return;
        setPersonDocuments(nextDocuments);
        if (!active) {
          setSelectedDocumentId("");
          setSalaryFields(salaryDefaults);
          setUbdFields(ubdDefaults);
          setDocumentFiles({});
          setWorkflow(createEmptyWorkflow());
          setDocumentMessage(
            `Для ID ${personExternalId} ще немає документа "${salaryDocumentTypeLabel(targetDocumentType)}". Створіть його вручну.`,
          );
          return;
        }
        setSelectedDocumentId(active.id);
        setSalaryFields(mergeSalaryFields(salaryDefaults, active.fields));
        setUbdFields(mergeUbdFields(ubdDefaults, active.fields));
        setDocumentFiles(mergeDocumentFiles(active.files));
        setWorkflow(mergeSalaryWorkflow(active.workflow));
        setDocumentMessage(
          `Відкрито документ: ${active.title} · ${new Date(active.updatedAt).toLocaleString("uk-UA")}`,
        );
      } catch (error) {
        if (cancelled) return;
        setDocumentMessage(
          error instanceof Error
            ? `БД документів недоступна: ${error.message}`
            : "БД документів недоступна.",
        );
      }
    };

    void loadDocuments();

    return () => {
      cancelled = true;
    };
  }, [
    isPersonDocumentMode,
    personExternalId,
    requestedDocumentId,
    selectedPerson,
    summary,
    targetDocumentType,
  ]);

  const loadDocumentJournal = useCallback(async (message: string) => {
    setIsLoadingDocumentJournal(true);
    setDocumentMessage(message);
    try {
      const documents = await api.listAllPersonDocuments();
      setAllPersonDocuments(documents);
      setDocumentMessage(`Документів у журналі: ${documents.length}.`);
    } catch (error) {
      setDocumentMessage(
        error instanceof Error
          ? `Не вдалося завантажити журнал: ${error.message}`
          : "Не вдалося завантажити журнал документів.",
      );
    } finally {
      setIsLoadingDocumentJournal(false);
    }
  }, []);

  useEffect(() => {
    if (isPersonDocumentMode) return;

    void loadDocumentJournal("Завантажую журнал документів...");
  }, [isPersonDocumentMode, loadDocumentJournal]);
  const currentStepIndex = Math.max(
    0,
    salaryWorkflowSteps.findIndex((step) => step.key === workflow.currentStatus),
  );
  const completedCount = currentStepIndex + 1;
  const progressPercent = Math.round(
    (completedCount / salaryWorkflowSteps.length) * 100,
  );

  const updateDefaultField = (key: DefaultDocumentFieldKey, value: string) => {
    setFields((current) => ({ ...current, [key]: value }));
  };

  const openPersonDocument = (document: BackendPersonDocument) => {
    setSelectedDocumentId(document.id);

    if (document.type === "ubdReport") {
      const defaults = createUbdFields(selectedPerson, summary);
      setMode("ubdReport");
      setUbdFields(mergeUbdFields(defaults, document.fields));
      setDocumentFiles(mergeDocumentFiles(document.files));
    } else {
      const defaults = createSalaryFields(selectedPerson, summary);
      setMode("salaryPowerAttorney");
      setSalaryFields(mergeSalaryFields(defaults, document.fields));
      setDocumentFiles({});
    }

    setWorkflow(mergeSalaryWorkflow(document.workflow));
    setDocumentMessage(
      `Відкрито документ: ${document.title} · ${new Date(document.updatedAt).toLocaleString("uk-UA")}`,
    );
  };

  const createSalaryPowerAttorneyDocument = async () => {
    if (!personExternalId) return;
    const defaults = createSalaryFields(selectedPerson, summary);
    const nextWorkflow = createEmptyWorkflow();
    setIsSavingDocument(true);
    setDocumentMessage("Створюю документ службовця...");
    try {
      const created = await api.createPersonDocument(personExternalId, {
        type: "salaryPowerAttorney",
        title: `Довіреність зарплати · ${dayjs().format("DD.MM.YYYY HH:mm")}`,
        status: nextWorkflow.currentStatus,
        fields: defaults as unknown as Record<string, unknown>,
        workflow: nextWorkflow as unknown as Record<string, unknown>,
      });
      setPersonDocuments((current) => [created, ...current]);
      setSelectedDocumentId(created.id);
      setSalaryFields(mergeSalaryFields(defaults, created.fields));
      setWorkflow(mergeSalaryWorkflow(created.workflow));
      setDocumentMessage(
        `Створено документ: ${created.title} · ${new Date(created.updatedAt).toLocaleString("uk-UA")}`,
      );
    } catch (error) {
      setDocumentMessage(
        error instanceof Error
          ? `Не вдалося створити документ: ${error.message}`
          : "Не вдалося створити документ.",
      );
    } finally {
      setIsSavingDocument(false);
    }
  };

  const createUbdReportDocument = async () => {
    if (!personExternalId) return;
    const defaults = createUbdFields(selectedPerson, summary);
    const nextWorkflow = { ...createEmptyWorkflow(), currentStatus: "document" };
    setIsSavingDocument(true);
    setDocumentMessage("Створюю УБД рапорт...");
    try {
      const created = await api.createPersonDocument(personExternalId, {
        type: "ubdReport",
        title: `Рапорт на УБД · ${dayjs().format("DD.MM.YYYY HH:mm")}`,
        status: "document",
        fields: defaults as unknown as Record<string, unknown>,
        workflow: nextWorkflow as unknown as Record<string, unknown>,
        files: { ubdScans: [] },
      });
      setPersonDocuments((current) => [created, ...current]);
      setSelectedDocumentId(created.id);
      setUbdFields(mergeUbdFields(defaults, created.fields));
      setDocumentFiles(mergeDocumentFiles(created.files));
      setWorkflow(mergeSalaryWorkflow(created.workflow));
      setDocumentMessage(
        `Створено документ: ${created.title} · ${new Date(created.updatedAt).toLocaleString("uk-UA")}`,
      );
    } catch (error) {
      setDocumentMessage(
        error instanceof Error
          ? `Не вдалося створити УБД рапорт: ${error.message}`
          : "Не вдалося створити УБД рапорт.",
      );
    } finally {
      setIsSavingDocument(false);
    }
  };

  const addUbdScanFiles = async (files: FileList | null) => {
    if (!files?.length || !selectedDocument) return;

    setIsSavingDocument(true);
    try {
      const uploaded = await Promise.all(
        Array.from(files).map(async (file) => ({
          id: `${Date.now()}-${file.name}-${Math.random().toString(16).slice(2)}`,
          name: file.name,
          type: file.type || "application/octet-stream",
          dataUrl: await readFileAsDataUrl(file),
          uploadedAt: new Date().toISOString(),
        })),
      );
      const nextFiles = {
        ...documentFiles,
        ubdScans: [...(documentFiles.ubdScans ?? []), ...uploaded],
      };
      setDocumentFiles(nextFiles);
      await saveUbdDocument(ubdFields, nextFiles);
    } catch (error) {
      setDocumentMessage(
        error instanceof Error
          ? `Не вдалося додати скани: ${error.message}`
          : "Не вдалося додати скани.",
      );
    } finally {
      setIsSavingDocument(false);
    }
  };

  const deleteUbdScanFile = (fileId: string) => {
    const nextFiles = {
      ...documentFiles,
      ubdScans: (documentFiles.ubdScans ?? []).filter(
        (file) => file.id !== fileId,
      ),
    };
    setDocumentFiles(nextFiles);
    void saveUbdDocument(ubdFields, nextFiles);
  };

  const saveSalaryDocument = async (
    nextFields: SalaryDocumentFields,
    nextWorkflow: SalaryWorkflowState,
  ) => {
    if (!personExternalId || !selectedDocumentId) {
      saveWorkflowForPerson(workflowKey, nextWorkflow);
      return;
    }

    setIsSavingDocument(true);
    try {
      const updated = await api.updatePersonDocument(
        personExternalId,
        selectedDocumentId,
        {
          title: selectedDocument?.title || "Довіреність зарплати",
          status: nextWorkflow.currentStatus,
          fields: nextFields as unknown as Record<string, unknown>,
          workflow: nextWorkflow as unknown as Record<string, unknown>,
        },
      );
      setPersonDocuments((current) =>
        current.map((document) =>
          document.id === updated.id ? updated : document,
        ),
      );
      setDocumentMessage(
        `Збережено в БД · ${new Date(updated.updatedAt).toLocaleString("uk-UA")}`,
      );
    } catch (error) {
      saveWorkflowForPerson(workflowKey, nextWorkflow);
      setDocumentMessage(
        error instanceof Error
          ? `Не вдалося зберегти в БД, залишив локально: ${error.message}`
          : "Не вдалося зберегти в БД, залишив локально.",
      );
    } finally {
      setIsSavingDocument(false);
    }
  };

  const saveUbdDocument = async (
    nextFields: UbdReportFields,
    nextFiles = documentFiles,
  ) => {
    if (!personExternalId || !selectedDocumentId) return;

    setIsSavingDocument(true);
    try {
      const updated = await api.updatePersonDocument(
        personExternalId,
        selectedDocumentId,
        {
          title: selectedDocument?.title || "Рапорт на УБД",
          status: "document",
          fields: nextFields as unknown as Record<string, unknown>,
          workflow: workflow as unknown as Record<string, unknown>,
          files: nextFiles as unknown as Record<string, unknown>,
        },
      );
      setPersonDocuments((current) =>
        current.map((document) =>
          document.id === updated.id ? updated : document,
        ),
      );
      setDocumentMessage(
        `Збережено в БД · ${new Date(updated.updatedAt).toLocaleString("uk-UA")}`,
      );
    } catch (error) {
      setDocumentMessage(
        error instanceof Error
          ? `Не вдалося зберегти УБД рапорт: ${error.message}`
          : "Не вдалося зберегти УБД рапорт.",
      );
    } finally {
      setIsSavingDocument(false);
    }
  };

  const updateUbdField = (key: keyof UbdReportFields, value: string) => {
    setUbdFields((current) => {
      const next = { ...current, [key]: value };
      void saveUbdDocument(next);
      return next;
    });
  };

  const updateSalaryField = (key: keyof SalaryDocumentFields, value: string) => {
    setSalaryFields((current) => {
      if (key === "iban") {
        const iban = normalizeIban(value);
        const mfo = getBankMfoFromIban(iban);
        const bank = findBankByMfo(mfo);
        const next = {
          ...current,
          iban,
          bankMfo: mfo || current.bankMfo,
          bankName: bank?.name || current.bankName,
        };
        void saveSalaryDocument(next, workflow);
        return next;
      }

      if (key === "bankMfo") {
        const bank = findBankByMfo(value);
        const next = {
          ...current,
          bankMfo: value,
          bankName: value === "custom" ? current.bankName : bank?.name || "",
        };
        void saveSalaryDocument(next, workflow);
        return next;
      }

      const next = { ...current, [key]: value };
      void saveSalaryDocument(next, workflow);
      return next;
    });
  };

  const updateWorkflow = (next: SalaryWorkflowState) => {
    setWorkflow(next);
    saveWorkflowForPerson(workflowKey, next);
    void saveSalaryDocument(salaryFields, next);
  };

  const setWorkflowStep = (stepKey: string) => {
    const selectedIndex = salaryWorkflowSteps.findIndex(
      (step) => step.key === stepKey,
    );
    if (selectedIndex < 0) return;
    const completed = Object.fromEntries(
      salaryWorkflowSteps.map((step, index) => [
        step.key,
        index <= selectedIndex,
      ]),
    ) as Record<string, boolean>;

    updateWorkflow({
      ...workflow,
      completed,
      currentStatus: stepKey,
    });
  };

  const generateSalaryDoc = () => {
    const content = salaryDocumentHtml(salaryFields);
    downloadTextFile(
      `${safeFilePart(salaryFields.folderName)}.doc`,
      content,
      "application/msword;charset=utf-8",
    );
    setWorkflowStep("document");
  };

  const printSalaryDocument = () => {
    const printWindow = window.open("", "_blank", "noopener,noreferrer");
    if (!printWindow) return;
    printWindow.document.write(salaryDocumentHtml(salaryFields));
    printWindow.document.close();
    printWindow.focus();
    printWindow.print();
    setWorkflowStep("print");
  };

  const generateUbdDoc = () => {
    downloadTextFile(
      `${safeFilePart(ubdFields.folderName)}.doc`,
      ubdReportHtml(ubdFields),
      "application/msword;charset=utf-8",
    );
    void saveUbdDocument(ubdFields);
  };

  const printUbdDocument = () => {
    const printWindow = window.open("", "_blank", "noopener,noreferrer");
    if (!printWindow) return;
    printWindow.document.write(ubdReportHtml(ubdFields));
    printWindow.document.close();
    printWindow.focus();
    printWindow.print();
    void saveUbdDocument(ubdFields);
  };

  const deleteDocument = async (document: BackendPersonDocument) => {
    const confirmed = window.confirm(
      `Видалити документ "${document.title}" для ID ${document.personExternalId}?`,
    );
    if (!confirmed) return;

    setIsSavingDocument(true);
    setDocumentMessage("Видаляю документ...");
    try {
      await api.deletePersonDocument(document.personExternalId, document.id);
      setAllPersonDocuments((current) =>
        current.filter((item) => item.id !== document.id),
      );
      setPersonDocuments((current) => {
        const next = current.filter((item) => item.id !== document.id);
        if (document.id === selectedDocumentId) {
          const replacement = next[0] ?? null;
          setSelectedDocumentId(replacement?.id ?? "");
          if (replacement) {
            openPersonDocument(replacement);
          } else {
            setSalaryFields(createSalaryFields(selectedPerson, summary));
            setWorkflow(createEmptyWorkflow());
          }
        }
        return next;
      });
      setDocumentMessage(`Документ видалено: ${document.title}.`);
    } catch (error) {
      setDocumentMessage(
        error instanceof Error
          ? `Не вдалося видалити документ: ${error.message}`
          : "Не вдалося видалити документ.",
      );
    } finally {
      setIsSavingDocument(false);
    }
  };

  if (!isPersonDocumentMode) {
    return (
      <main className="main-panel">
        <header className="topbar analytics-topbar salary-document-topbar">
          <Box className="salary-document-title">
            <Typography component="h1" variant="h4">
              Документи
            </Typography>
            <Typography variant="body2" color="text.secondary">
              Загальний журнал рапортів · прогрес по всіх службовцях
            </Typography>
          </Box>
          <Stack direction="row" spacing={1}>
            <Button
              variant="outlined"
              startIcon={<ArticleOutlinedIcon />}
              onClick={() =>
                void loadDocumentJournal("Оновлюю журнал документів...")
              }
            >
              Оновити
            </Button>
          </Stack>
        </header>

        <Alert
          severity={isLoadingDocumentJournal ? "info" : "success"}
          variant="outlined"
          className="personnel-page-alert"
        >
          {documentMessage || "Журнал документів готовий."}
        </Alert>

        <section className="analytics-panel documents-journal-panel">
          <div className="panel-heading">
            Усі документи · {allPersonDocuments.length}
          </div>
          <div className="documents-journal-table">
            <div className="documents-journal-row header">
              <span>Службовець</span>
              <span>ID</span>
              <span>Документ</span>
              <span>Прогрес</span>
              <span>Статус</span>
              <span>Файли</span>
              <span>Оновлено</span>
              <span />
            </div>
            {allPersonDocuments.length ? (
              allPersonDocuments.map((document) => {
                const progress = getDocumentProgressPercent(document);
                const href = buildDocumentRoute({
                  personExternalId: document.personExternalId,
                  documentId: document.id,
                  type: document.type,
                });
                const openDocument = () => onNavigate?.(href);
                return (
                  <div
                    className="documents-journal-row"
                    key={document.id}
                    role="button"
                    tabIndex={0}
                    onClick={openDocument}
                    onKeyDown={(event) => {
                      if (event.key !== "Enter" && event.key !== " ") return;
                      event.preventDefault();
                      openDocument();
                    }}
                  >
                    <strong>{getDocumentPersonName(document)}</strong>
                    <span>{document.personExternalId}</span>
                    <span>{salaryDocumentTypeLabel(document.type)}</span>
                    <span className="documents-journal-progress">
                      <i style={{ width: `${progress}%` }} />
                      <b>{progress}%</b>
                    </span>
                    <span>{workflowStatusLabel(document.status)}</span>
                    <span>{getDocumentFileSummary(document)}</span>
                    <span>
                      {new Date(document.updatedAt).toLocaleString("uk-UA")}
                    </span>
                    <button
                      aria-label="Видалити документ"
                      className="documents-journal-delete"
                      disabled={isSavingDocument}
                      onClick={(event) => {
                        event.stopPropagation();
                        void deleteDocument(document);
                      }}
                      title="Видалити документ"
                      type="button"
                    >
                      <DeleteOutlineOutlinedIcon />
                    </button>
                  </div>
                );
              })
            ) : (
              <div className="documents-journal-empty">
                Документи ще не створювались.
              </div>
            )}
          </div>
        </section>
      </main>
    );
  }

  if (mode === "ubdReport") {
    return (
      <main className="main-panel">
        <header className="topbar analytics-topbar salary-document-topbar">
          <div className="salary-document-main-row">
            <Box className="salary-document-title">
              <Typography component="h1" variant="h4">
                Рапорт на УБД
              </Typography>
              <Typography variant="body2" color="text.secondary">
                {summary.name} · {summary.externalId ? `ID ${summary.externalId}` : "дані з картки особи"}
                {selectedDocument ? ` · ${selectedDocument.title}` : ""}
              </Typography>
            </Box>
            {selectedDocument ? (
              <Stack direction="row" spacing={1}>
                <Button
                  variant="outlined"
                  startIcon={<ArticleOutlinedIcon />}
                  onClick={generateUbdDoc}
                >
                  Завантажити DOC
                </Button>
                <Button
                  variant="contained"
                  startIcon={<PictureAsPdfOutlinedIcon />}
                  onClick={printUbdDocument}
                  sx={{ color: "#1a1a14" }}
                >
                  Друк / PDF
                </Button>
              </Stack>
            ) : null}
          </div>
        </header>

        <section
          className={
            selectedDocument
              ? "salary-documents-layout ubd-report-layout"
              : "salary-documents-layout empty"
          }
        >
          <section className="analytics-panel salary-person-documents-panel">
            <div className="panel-heading">Документи службовця</div>
            <div className="salary-document-sync">
              <span>{isSavingDocument ? "збереження..." : "БД синхронізована"}</span>
              <p>{documentMessage || "Виберіть документ зі списку."}</p>
            </div>
            <div className="salary-person-documents-list">
              {personDocuments.length ? (
                personDocuments.map((document) => (
                  <article
                    className={
                      document.id === selectedDocumentId
                        ? "salary-person-document-shell active"
                        : "salary-person-document-shell"
                    }
                    key={document.id}
                  >
                    <button
                      className="salary-person-document"
                      type="button"
                      onClick={() => openPersonDocument(document)}
                    >
                      <strong>{document.title}</strong>
                      <span>{salaryDocumentTypeLabel(document.type)}</span>
                      <em>{workflowStatusLabel(document.status)}</em>
                      <small>
                        {new Date(document.updatedAt).toLocaleString("uk-UA")}
                      </small>
                    </button>
                    <button
                      aria-label="Видалити документ"
                      className="salary-person-document-delete"
                      disabled={isSavingDocument}
                      onClick={() => void deleteDocument(document)}
                      title="Видалити документ"
                      type="button"
                    >
                      <DeleteOutlineOutlinedIcon />
                    </button>
                  </article>
                ))
              ) : (
                <p className="salary-empty-documents">
                  Документів по цьому службовцю ще немає.
                </p>
              )}
            </div>
            <Button
              variant="outlined"
              onClick={createUbdReportDocument}
              disabled={isSavingDocument || !personExternalId}
            >
              Створити УБД рапорт
            </Button>
          </section>

          {selectedDocument ? (
            <section className="analytics-panel document-fields">
              <div className="panel-heading">Дані для УБД рапорта</div>
              <div className="document-placeholder-map salary-document-map ubd-document-map">
                {[
                  ["fullName", "ПІБ"],
                  ["rank", "Звання"],
                  ["staffPosition", "Посада"],
                  ["birthDate", "Дата народження"],
                  ["rnokpp", "РНОКПП"],
                  ["taskPeriod", "Період завдань"],
                  ["taskPlace", "Місце завдань"],
                  ["basis", "Підстава"],
                  ["date", "Дата"],
                  ["folderName", "Папка"],
                ].map(([key, label]) => (
                  <label key={key}>
                    <code>{label}</code>
                    <TextField
                      size="small"
                      fullWidth
                      value={ubdFields[key as keyof UbdReportFields]}
                      onChange={(event) =>
                        updateUbdField(
                          key as keyof UbdReportFields,
                          event.target.value,
                        )
                      }
                    />
                  </label>
                ))}
                <label className="wide">
                  <code>Кому</code>
                  <TextField
                    size="small"
                    fullWidth
                    value={ubdFields.commander}
                    onChange={(event) =>
                      updateUbdField("commander", event.target.value)
                    }
                  />
                </label>
                <label className="wide">
                  <code>Командир БТ</code>
                  <TextField
                    size="small"
                    fullWidth
                    multiline
                    value={ubdFields.battalionCommander}
                    onChange={(event) =>
                      updateUbdField("battalionCommander", event.target.value)
                    }
                  />
                </label>
                <label className="wide">
                  <code>Затверджую</code>
                  <TextField
                    size="small"
                    fullWidth
                    multiline
                    value={ubdFields.approvalOfficer}
                    onChange={(event) =>
                      updateUbdField("approvalOfficer", event.target.value)
                    }
                  />
                </label>
              </div>

              <div className="ubd-scans-panel">
                <div className="panel-heading">Скани для УБД</div>
                <Button component="label" variant="outlined">
                  Додати скани
                  <input
                    hidden
                    multiple
                    type="file"
                    accept="image/*,application/pdf,.pdf"
                    onChange={(event) => {
                      void addUbdScanFiles(event.target.files);
                      event.target.value = "";
                    }}
                  />
                </Button>
                <div className="ubd-scan-list">
                  {(documentFiles.ubdScans ?? []).length ? (
                    (documentFiles.ubdScans ?? []).map((file) => (
                      <article className="ubd-scan-item" key={file.id}>
                        <a href={file.dataUrl} target="_blank" rel="noreferrer">
                          {file.type.startsWith("image/") ? (
                            <img alt={file.name} src={file.dataUrl} />
                          ) : (
                            <PictureAsPdfOutlinedIcon />
                          )}
                          <span>{file.name}</span>
                        </a>
                        <button
                          aria-label="Видалити скан"
                          className="documents-journal-delete"
                          onClick={() => deleteUbdScanFile(file.id)}
                          type="button"
                        >
                          <DeleteOutlineOutlinedIcon />
                        </button>
                      </article>
                    ))
                  ) : (
                    <p className="salary-empty-documents">
                      Додайте фото, ІНН, копію паспорта або затверджений рапорт.
                    </p>
                  )}
                </div>
              </div>
            </section>
          ) : null}

          {selectedDocument ? (
            <section className="analytics-panel document-preview">
              <div className="panel-heading">Попередній перегляд</div>
              <div className="document-page-preview salary-document-preview ubd-document-preview">
                <div className="salary-preview-to">{ubdFields.commander}</div>
                <h2>Рапорт</h2>
                <p>
                  Прошу Вашого рішення про внесення відомостей до Єдиного
                  державного реєстру ветеранів війни щодо зазначеного нижче
                  військовослужбовця з метою надання статусу учасника бойових
                  дій, а саме:
                </p>
                <table>
                  <tbody>
                    <tr><td>Військове звання</td><td>{ubdFields.rank || "—"}</td></tr>
                    <tr><td>ПІБ</td><td>{ubdFields.fullName || "—"}</td></tr>
                    <tr><td>Посада згідно штату</td><td>{ubdFields.staffPosition || "—"}</td></tr>
                    <tr><td>Дата народження</td><td>{ubdFields.birthDate || "—"}</td></tr>
                    <tr><td>РНОКПП</td><td>{ubdFields.rnokpp || "—"}</td></tr>
                    <tr><td>Період виконання завдань</td><td>{ubdFields.taskPeriod || "—"}</td></tr>
                    <tr><td>Місце виконання завдань</td><td>{ubdFields.taskPlace || "—"}</td></tr>
                  </tbody>
                </table>
                <p><strong>Підстава:</strong> {ubdFields.basis}</p>
                <p><strong>Додаток:</strong> копія паспорта, РНОКПП, дві фотокартки</p>
                <p className="ubd-preview-sign">{ubdFields.battalionCommander}</p>
                <p>{ubdFields.date}</p>
                <p className="ubd-preview-approve">ЗАТВЕРДЖУЮ<br />{ubdFields.approvalOfficer}</p>
              </div>
            </section>
          ) : null}
        </section>
      </main>
    );
  }

  if (mode === "salaryPowerAttorney") {
    return (
      <main className="main-panel">
        <header className="topbar analytics-topbar salary-document-topbar">
          <div className="salary-document-main-row">
            <Box className="salary-document-title">
              <Typography component="h1" variant="h4">
                Довіреність зарплати
              </Typography>
              <Typography variant="body2" color="text.secondary">
                {summary.name} · {summary.externalId ? `ID ${summary.externalId}` : "дані з картки особи"}
                {selectedDocument ? ` · ${selectedDocument.title}` : ""}
              </Typography>
            </Box>
            {selectedDocument ? (
              <Stack direction="row" spacing={1}>
                <Button
                  variant="outlined"
                  startIcon={<ArticleOutlinedIcon />}
                  onClick={generateSalaryDoc}
                >
                  Завантажити DOC
                </Button>
                <Button
                  variant="contained"
                  startIcon={<PictureAsPdfOutlinedIcon />}
                  onClick={printSalaryDocument}
                  sx={{ color: "#1a1a14" }}
                >
                  Друк / PDF
                </Button>
              </Stack>
            ) : null}
          </div>
          {selectedDocument ? (
            <div className="salary-step-progress-header">
              <div className="salary-progress">
                <span>{progressPercent}%</span>
                <div>
                  <i style={{ width: `${progressPercent}%` }} />
                </div>
              </div>
              <div className="salary-step-bar">
                {salaryWorkflowSteps.map((step, index) => {
                  const isDone = index <= currentStepIndex;
                  const isCurrent = index === currentStepIndex;
                  return (
                    <button
                      className={[
                        "salary-step-node",
                        isDone ? "done" : "",
                        isCurrent ? "current" : "",
                      ]
                        .filter(Boolean)
                        .join(" ")}
                      key={step.key}
                      type="button"
                      onClick={() => setWorkflowStep(step.key)}
                      title={step.title}
                    >
                      <span>{index + 1}</span>
                      <i />
                      <strong>{step.title}</strong>
                    </button>
                  );
                })}
              </div>
            </div>
          ) : null}
        </header>

        <section
          className={
            selectedDocument
              ? "salary-documents-layout"
              : "salary-documents-layout empty"
          }
        >
          <section className="analytics-panel salary-person-documents-panel">
            <div className="panel-heading">Документи службовця</div>
            <div className="salary-document-sync">
              <span>{isSavingDocument ? "збереження..." : "БД синхронізована"}</span>
              <p>{documentMessage || "Виберіть документ зі списку."}</p>
            </div>
            <div className="salary-person-documents-list">
              {personDocuments.length ? (
                personDocuments.map((document) => (
                  <article
                    className={
                      document.id === selectedDocumentId
                        ? "salary-person-document-shell active"
                        : "salary-person-document-shell"
                    }
                    key={document.id}
                  >
                    <button
                      className="salary-person-document"
                      type="button"
                      onClick={() => openPersonDocument(document)}
                    >
                      <strong>{document.title}</strong>
                      <span>{salaryDocumentTypeLabel(document.type)}</span>
                      <em>{workflowStatusLabel(document.status)}</em>
                      <small>
                        {new Date(document.updatedAt).toLocaleString("uk-UA")}
                      </small>
                    </button>
                    <button
                      aria-label="Видалити документ"
                      className="salary-person-document-delete"
                      disabled={isSavingDocument}
                      onClick={() => void deleteDocument(document)}
                      title="Видалити документ"
                      type="button"
                    >
                      <DeleteOutlineOutlinedIcon />
                    </button>
                  </article>
                ))
              ) : (
                <p className="salary-empty-documents">
                  Документів по цьому службовцю ще немає.
                </p>
              )}
            </div>
            <Button
              variant="outlined"
              onClick={createSalaryPowerAttorneyDocument}
              disabled={isSavingDocument || !personExternalId}
            >
              Створити довіреність
            </Button>
          </section>

          {selectedDocument ? (
            <section className="analytics-panel document-fields">
              <div className="panel-heading">Дані для рапорта</div>
              <div className="document-placeholder-map salary-document-map">
              <label>
                <code>ФИО</code>
                <TextField
                  size="small"
                  fullWidth
                  value={salaryFields.fullName}
                  onChange={(event) =>
                    updateSalaryField("fullName", event.target.value)
                  }
                  placeholder="ПРІЗВИЩЕ Ім'я По батькові"
                />
              </label>
              <label>
                <code>ІНН</code>
                <TextField
                  size="small"
                  fullWidth
                  value={salaryFields.rnokpp}
                  onChange={(event) =>
                    updateSalaryField("rnokpp", event.target.value)
                  }
                  placeholder="РНОКПП"
                />
              </label>
              <label>
                <code>IBAN</code>
                <TextField
                  size="small"
                  fullWidth
                  value={formatIban(salaryFields.iban)}
                  onChange={(event) =>
                    updateSalaryField("iban", event.target.value)
                  }
                  placeholder="UA..."
                />
              </label>
              <label>
                <code>БАНК</code>
                <TextField
                  select
                  size="small"
                  fullWidth
                  value={salaryFields.bankMfo}
                  onChange={(event) =>
                    updateSalaryField("bankMfo", event.target.value)
                  }
                >
                  {ukrainianBanks.map((bank) => (
                    <MenuItem key={bank.mfo} value={bank.mfo}>
                      {bank.mfo === "custom"
                        ? bank.name
                        : `${bank.name} · ${bank.mfo}`}
                    </MenuItem>
                  ))}
                </TextField>
              </label>
              <label>
                <code>Назва банку</code>
                <TextField
                  size="small"
                  fullWidth
                  value={salaryFields.bankName}
                  onChange={(event) =>
                    updateSalaryField("bankName", event.target.value)
                  }
                  placeholder="АТ «...»"
                />
              </label>
              <label>
                <code>Дата</code>
                <TextField
                  size="small"
                  fullWidth
                  value={salaryFields.date}
                  onChange={(event) =>
                    updateSalaryField("date", event.target.value)
                  }
                  placeholder="ДД.ММ.РРРР"
                />
              </label>
              <label>
                <code>Папка</code>
                <TextField
                  size="small"
                  fullWidth
                  value={salaryFields.folderName}
                  onChange={(event) =>
                    updateSalaryField("folderName", event.target.value)
                  }
                  placeholder="Назва папки"
                />
              </label>
              </div>
            </section>
          ) : null}

          {selectedDocument ? (
            <section className="analytics-panel document-preview">
              <div className="panel-heading">Попередній перегляд</div>
              <div className="document-page-preview salary-document-preview">
              <div className="salary-preview-to">{salaryFields.commander}</div>
              <h2>Рапорт</h2>
              <p>
                Прошу моє грошове забезпечення перераховувати на мій
                розрахунковий банківський рахунок:
              </p>
              <p>
                {highlightOrPlaceholder(
                  formatIban(salaryFields.iban),
                  "UA___________________________",
                )}{" "}
                відкритий{" "}
                {highlightOrPlaceholder(salaryFields.bankName, "банк")}.
              </p>
              <p>
                Код РНОКПП{" "}
                {highlightOrPlaceholder(salaryFields.rnokpp, "__________")}.
              </p>
              <p>Довідка з банківськими реквізитами додається.</p>
              <p>
                {highlightOrPlaceholder(salaryFields.rank, "звання")}{" "}
                {highlightOrPlaceholder(
                  salaryFields.fullName,
                  "ПРІЗВИЩЕ Ім'я По батькові",
                )}
              </p>
              <footer>
                <span>{salaryFields.date || "__.__.202__"}</span>
                <span>________________</span>
              </footer>
              <p className="salary-finance-note">
                Начальнику фінансово-економічної служби – головному бухгалтеру
                військової частини А4862 - прошу прийняти в роботу
              </p>
              <p>
                Начальник групи персоналу штабу
                <br />1 піхотного батальйону
                <br />військової частини А4862
              </p>
              <p>{salaryFields.personnelChief}</p>
              </div>
            </section>
          ) : null}
        </section>
      </main>
    );
  }

  return (
    <main className="main-panel">
      <header className="topbar analytics-topbar">
        <Box>
          <Typography component="h1" variant="h4">
            Документи та звіти
          </Typography>
          <Typography variant="body2" color="text.secondary">
            Шаблони DOCX · джерело даних з картки особи
          </Typography>
        </Box>
        <Stack direction="row" spacing={1}>
          <Button variant="outlined" startIcon={<ArticleOutlinedIcon />}>
            Додати DOCX
          </Button>
          <Button variant="contained" sx={{ color: "#1a1a14" }}>
            Згенерувати DOCX
          </Button>
        </Stack>
      </header>

      <section className="documents-layout">
        <aside className="analytics-panel document-template-list">
          <div className="panel-heading">Бібліотека шаблонів</div>
          {["Довідка", "Наказ", "Витяг", "Рапорт", "Власний шаблон"].map(
            (template, index) => (
              <button
                className={index === 0 ? "active" : ""}
                key={template}
                type="button"
              >
                <ArticleOutlinedIcon />
                <span>
                  <strong>{template}</strong>
                  <small>
                    {index === 0
                      ? "Довідка про проходження служби"
                      : "Шаблон документа"}
                  </small>
                </span>
              </button>
            ),
          )}
        </aside>

        <section className="analytics-panel document-fields">
          <div className="panel-heading">Джерело даних</div>
          <div className="person-action-fields">
            <span>
              <strong>Особа</strong>
              {summary.name}
            </span>
            <span>
              <strong>Звання</strong>
              {summary.rank || "—"}
            </span>
            <span>
              <strong>Підрозділ / місце</strong>
              {summary.location || "—"}
            </span>
          </div>

          <div className="document-placeholder-map">
            <label>
              <code>{"{{ПІБ}}"}</code>
              <TextField
                size="small"
                fullWidth
                value={fields.pib}
                onChange={(event) => updateDefaultField("pib", event.target.value)}
                placeholder="Прізвище Ім'я По батькові"
              />
            </label>
            <label>
              <code>{"{{Звання}}"}</code>
              <TextField
                size="small"
                fullWidth
                value={fields.rank}
                onChange={(event) => updateDefaultField("rank", event.target.value)}
                placeholder="Звання"
              />
            </label>
            <label>
              <code>{"{{Підрозділ}}"}</code>
              <TextField
                size="small"
                fullWidth
                value={fields.unit}
                onChange={(event) => updateDefaultField("unit", event.target.value)}
                placeholder="Підрозділ / частина"
              />
            </label>
            <label>
              <code>{"{{Дата}}"}</code>
              <TextField
                size="small"
                fullWidth
                value={fields.date}
                onChange={(event) => updateDefaultField("date", event.target.value)}
                placeholder="ДД.ММ.РРРР"
              />
            </label>
          </div>
        </section>

        <section className="analytics-panel document-preview">
          <div className="panel-heading">Попередній перегляд документа</div>
          <div className="document-page-preview">
            <h2>ДОВІДКА</h2>
            <h3>про проходження служби та займану посаду</h3>
            <p>
              Видана {highlightOrPlaceholder(fields.pib, "{{ПІБ}}")},{" "}
              {highlightOrPlaceholder(fields.rank, "{{Звання}}")},
              військовослужбовцю, що проходить військову службу у{" "}
              {highlightOrPlaceholder(fields.unit, "{{Підрозділ}}")}.
            </p>
            <p>
              Дата видачі: {highlightOrPlaceholder(fields.date, "{{Дата}}")}.
            </p>
            <p>Довідка видана для пред’явлення за місцем вимоги.</p>
            <footer>
              <span>Командир</span>
              <span>________________</span>
            </footer>
          </div>
        </section>
      </section>
    </main>
  );
}
