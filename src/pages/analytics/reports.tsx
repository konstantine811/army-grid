import { useMemo, useState } from "react";
import { Button } from "@/components/sci/SciPrimitives";
import { ContentCopyOutlinedIcon } from "@/components/sci/icons";
import type { AnalyticsData, AnalyticsMetric } from "./analyticsData";
import { BarList } from "./charts";

export const formatMetricLines = (items: AnalyticsMetric[]) =>
  items
    .map(
      (item) =>
        `${item.label} ${item.value}${item.detail ? ` (${item.detail})` : ""}`,
    )
    .join("\n");

export const getAttachedTotal = (analytics: AnalyticsData) =>
  analytics.attachedDetails.reduce((sum, item) => sum + item.value, 0);
export const getBrezTotal = (analytics: AnalyticsData) =>
  (analytics.attachedDetails.find((item) => item.label === "БРЕЗ")?.value ??
    0) +
  (analytics.newcomerDetails.find((item) => item.label === "БРЕЗ/Без БЗВП")
    ?.value ?? 0) +
  (analytics.newcomerDetails.find((item) => item.label === "БРЕЗ")?.value ?? 0);
export const getExecutionTotal = (analytics: AnalyticsData) =>
  analytics.execution.reduce((sum, item) => sum + item.value, 0);
export const getBzvpTotal = (analytics: AnalyticsData) =>
  analytics.bzvp.reduce((sum, item) => sum + item.value, 0);

export const getShortReportRows = (analytics: AnalyticsData) => {
  const attachedTotal = getAttachedTotal(analytics);
  const totalPersonnel =
    analytics.metrics.inRanks +
    attachedTotal +
    analytics.metrics.newcomersNoRole;

  return [
    { label: "За штатом", value: analytics.metrics.staff, group: "Основне" },
    { label: "За списком", value: analytics.metrics.listed, group: "Основне" },
    { label: "Відсутні", value: analytics.metrics.absent, group: "Основне" },
    { label: "В строю", value: analytics.metrics.inRanks, group: "Основне" },
    {
      label: "Управління",
      value:
        analytics.structure.find((item) => item.label === "Управління")
          ?.value ?? 0,
      group: "Структура",
    },
    {
      label: "Забезпечення",
      value:
        analytics.structure.find((item) => item.label === "Забезпечення")
          ?.value ?? 0,
      group: "Структура",
    },
    {
      label: "БГ (Піхота)",
      value:
        analytics.structure.find((item) => item.label === "БГ піхота")?.value ??
        0,
      group: "Структура",
    },
    {
      label: "Не класифіковано",
      value:
        analytics.structure.find((item) => item.label === "Не класифіковано")
          ?.value ?? 0,
      group: "Структура",
    },
    { label: "Прикомандировані", value: attachedTotal, group: "Резерв" },
    {
      label: "Новоприбулі",
      value: analytics.metrics.newcomersNoRole,
      group: "Резерв",
    },
    { label: "БРЕЗ", value: getBrezTotal(analytics), group: "Резерв" },
    {
      label: "На виконанні",
      value: getExecutionTotal(analytics),
      group: "Завдання",
    },
    {
      label: "Піхота",
      value:
        analytics.execution.find((item) => item.label === "Піхота")?.value ?? 0,
      group: "Завдання",
    },
    {
      label: "Зв'язок",
      value:
        analytics.execution.find((item) => item.label === "Зв'язок")?.value ??
        0,
      group: "Завдання",
    },
    {
      label: "Розрахунок МК",
      value:
        analytics.execution.find((item) => item.label === "Розрахунок МК")
          ?.value ?? 0,
      group: "Завдання",
    },
    {
      label: "РБпАК",
      value:
        analytics.execution.find((item) => item.label === "РБпАК")?.value ?? 0,
      group: "Завдання",
    },
    {
      label: "Загальна кількість о/с",
      value: totalPersonnel,
      group: "Підсумок",
    },
    { label: "Без БЗВП", value: getBzvpTotal(analytics), group: "БЗВП" },
    ...analytics.bzvp.map((item) => ({
      label: item.label,
      value: item.value,
      group: "БЗВП",
    })),
  ];
};

export const getHintSections = (analytics: AnalyticsData) => [
  { title: "Управління", items: analytics.hintManagementDetails },
  { title: "Забезпечення", items: analytics.hintSupportDetails },
  { title: "Лікування", items: analytics.hintHealingDetails },
  { title: "Піхота", items: analytics.hintInfantryDetails },
  { title: "В/с на ППД Вишневе", items: analytics.hintPpdDetails },
  { title: "РБпАК", items: analytics.hintRpakDetails },
];

export function buildTextReports(analytics: AnalyticsData) {
  const healing = analytics.absenceReasons.find(
    (item) => item.label === "Лікування",
  );
  const shortRows = getShortReportRows(analytics);
  const groupedShortRows = Array.from(
    new Set(shortRows.map((row) => row.group)),
  ).map((group) => ({
    group,
    rows: shortRows.filter((row) => row.group === group),
  }));
  const hintText = getHintSections(analytics)
    .map((section) => {
      const total = section.items.reduce((sum, item) => sum + item.value, 0);

      return [`${section.title}: ${total}`, formatMetricLines(section.items)]
        .filter(Boolean)
        .join("\n");
    })
    .join("\n\n");

  return [
    {
      title: "Повний звіт",
      text: [
        "Піхотний батальйон",
        `Стройова записка станом на: ${analytics.reportDate}`,
        "",
        `За штатом ${analytics.metrics.staff}`,
        `За списком ${analytics.metrics.listed}`,
        `Відсутні ${analytics.metrics.absent}`,
        "",
        `${healing?.label ?? "Лікування"} ${healing?.value ?? 0}${healing?.detail ? ` (${healing.detail})` : ""}`,
        formatMetricLines(
          analytics.absenceReasons.filter((item) => item.label !== "Лікування"),
        ),
        "",
        `Відкомандировані: ${analytics.detachedDetails.reduce((sum, item) => sum + item.value, 0)}`,
        formatMetricLines(analytics.detachedDetails),
        "",
        `Прикомандировані: ${getAttachedTotal(analytics)}`,
        formatMetricLines(analytics.attachedDetails),
        "",
        `В строю ${analytics.metrics.inRanks}`,
        `Управління: ${analytics.managementDetails.reduce((sum, item) => sum + item.value, 0)}`,
        formatMetricLines(analytics.managementDetails),
        "",
        `Забезпечення: ${analytics.supportDetails.reduce((sum, item) => sum + item.value, 0)}`,
        formatMetricLines(analytics.supportDetails),
        "",
        formatMetricLines(analytics.structure),
        `без БЗВП ${getBzvpTotal(analytics)} (${analytics.bzvp.map((item) => `${item.label.replace("курс ", "")}: ${item.value}`).join(", ")})`,
        `Новоприбулі без посад ${analytics.metrics.newcomersNoRole}`,
        "",
        `На виконанні: ${getExecutionTotal(analytics)}`,
        formatMetricLines(analytics.execution),
      ].join("\n"),
    },
    {
      title: "Скорочено",
      text: groupedShortRows
        .map((section) =>
          [
            `${section.group}:`,
            section.rows.map((row) => `${row.label} ${row.value}`).join("\n"),
          ].join("\n"),
        )
        .join("\n\n"),
    },
    {
      title: "Підказка",
      text: hintText,
    },
  ];
}

export function CopyTextButton({ text }: { text: string }) {
  const [copied, setCopied] = useState(false);

  const copyText = async () => {
    await navigator.clipboard.writeText(text);
    setCopied(true);
    window.setTimeout(() => setCopied(false), 1600);
  };

  return (
    <Button
      size="small"
      startIcon={<ContentCopyOutlinedIcon />}
      variant="outlined"
      onClick={() => void copyText()}
    >
      {copied ? "Скопійовано" : "Копіювати"}
    </Button>
  );
}

export function TextReports({ analytics }: { analytics: AnalyticsData }) {
  const reports = useMemo(() => buildTextReports(analytics), [analytics]);

  return (
    <section className="text-report-grid">
      {reports.map((report) => (
        <div className="text-report-card" key={report.title}>
          <div className="panel-heading">
            <span>{report.title}</span>
            <CopyTextButton text={report.text} />
          </div>
          <textarea readOnly value={report.text} />
        </div>
      ))}
    </section>
  );
}

export function ShortAnalytics({ analytics }: { analytics: AnalyticsData }) {
  const shortRows = getShortReportRows(analytics);

  return (
    <section className="short-analytics">
      <div className="short-summary-card">
        {shortRows.map((row) => (
          <div className="short-summary-row" key={`${row.group}-${row.label}`}>
            <span>{row.label}</span>
            <strong>{row.value}</strong>
            <em>{row.group}</em>
          </div>
        ))}
      </div>
      <div className="analytics-panel">
        <div className="panel-heading">Скорочено · БЗВП динамічно</div>
        <BarList items={analytics.bzvp} />
      </div>
      <div className="analytics-panel">
        <div className="panel-heading">Скорочено · завдання</div>
        <BarList items={analytics.execution} />
      </div>
    </section>
  );
}

export function HintAnalytics({ analytics }: { analytics: AnalyticsData }) {
  const sections = getHintSections(analytics);

  return (
    <section className="short-analytics">
      {sections.map((section) => {
        const total = section.items.reduce((sum, item) => sum + item.value, 0);

        return (
          <div className="analytics-panel" key={section.title}>
            <div className="panel-heading">
              {section.title} · {total}
            </div>
            <BarList items={section.items} />
          </div>
        );
      })}
    </section>
  );
}
