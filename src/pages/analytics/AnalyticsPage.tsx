import { useMemo, useState } from "react";
import {
  Alert,
  Box,
  Button,
  LinearProgress,
  Stack,
  Typography,
} from "@/components/sci/SciPrimitives";
import { CloudUploadOutlinedIcon } from "@/components/sci/icons";
import { FileDownloadOutlinedIcon } from "@/components/sci/icons";
import { UploadFileOutlinedIcon } from "@/components/sci/icons";
import type { MRT_ColumnDef } from "@/components/sci/SciDataTable";
import {
  MaterialReactTable,
  useMaterialReactTable,
} from "@/components/sci/SciDataTable";
import {
  type ExcelRow,
  getColumnHeader,
  readWorkbookSnapshot,
  valueToDisplay,
} from "../../excelRoundTrip";
import { makeAnalyticsData, type AnalyticsData } from "./analyticsData";
import {
  BarList,
  DonutChart,
  HorizontalSvgChart,
  VerticalSvgChart,
} from "./charts";
import { HintAnalytics, ShortAnalytics, TextReports } from "./reports";

export function AnalyticsPage() {
  const [analytics, setAnalytics] = useState<AnalyticsData | null>(null);
  const [analyticsTab, setAnalyticsTab] = useState<
    "full" | "short" | "hint" | "text"
  >("full");
  const [message, setMessage] = useState(
    "Завантажте файл строєвої записки, щоб побудувати аналітику з Аркуш1.",
  );
  const [isLoading, setIsLoading] = useState(false);

  const loadAnalyticsFile = async (file: File | undefined) => {
    if (!file) return;

    setIsLoading(true);
    try {
      const snapshot = await readWorkbookSnapshot(file);
      const nextAnalytics = makeAnalyticsData(snapshot);
      if (!nextAnalytics) {
        setMessage("У файлі не знайдено sheet Аркуш1 / Аркуш.");
        return;
      }

      setAnalytics(nextAnalytics);
      setMessage(
        `Завантажено ${file.name}. Розрахунки побудовані з ${nextAnalytics.rows.length} рядків Аркуш1.`,
      );
      console.groupCollapsed(`[army-grid] Battalion analytics: ${file.name}`);
      console.log(nextAnalytics);
      console.groupEnd();
    } catch (error) {
      setMessage(
        error instanceof Error
          ? error.message
          : "Не вдалося прочитати строєву записку.",
      );
    } finally {
      setIsLoading(false);
    }
  };

  const tableColumns = useMemo<MRT_ColumnDef<ExcelRow>[]>(() => {
    if (!analytics) return [];

    return Array.from(
      { length: analytics.sheet.columnCount },
      (_, columnIndex) => ({
        id: `analytics-column-${columnIndex}`,
        header: getColumnHeader(analytics.sheet, columnIndex),
        accessorFn: (row) => valueToDisplay(row.values[columnIndex]),
        size: columnIndex === 13 ? 280 : 170,
      }),
    );
  }, [analytics]);
  const table = useMaterialReactTable({
    columns: tableColumns,
    data: analytics?.rows ?? [],
    enableColumnResizing: true,
    enableColumnVirtualization: true,
    enableRowVirtualization: true,
    enableStickyHeader: true,
    initialState: {
      density: "compact",
      pagination: { pageIndex: 0, pageSize: 25 },
    },
    muiTablePaperProps: {
      elevation: 0,
      sx: { backgroundColor: "transparent" },
    },
    muiTableContainerProps: {
      sx: { maxHeight: 520, backgroundColor: "transparent" },
    },
    muiTableHeadCellProps: {
      sx: {
        backgroundColor: "#131311",
        color: "#d9d49d",
        borderColor: "rgba(230,224,190,0.12)",
      },
    },
    muiTableBodyCellProps: {
      sx: {
        backgroundColor: "#11110f",
        color: "#f2eee1",
        borderColor: "rgba(230,224,190,0.1)",
      },
    },
    muiTopToolbarProps: { sx: { backgroundColor: "rgba(17,17,15,0.92)" } },
    muiBottomToolbarProps: { sx: { backgroundColor: "rgba(17,17,15,0.92)" } },
  });

  return (
    <main className="main-panel">
      <header className="topbar analytics-topbar">
        <Box>
          <Typography component="h1" variant="h4">
            Аналітика батальйону
          </Typography>
          <Typography variant="body2" color="text.secondary">
            Стройова записка · розрахунок з Аркуш1
          </Typography>
        </Box>
        <Stack direction="row" spacing={1}>
          <Button
            component="label"
            variant="outlined"
            startIcon={<UploadFileOutlinedIcon />}
          >
            Додати строєву
            <input
              hidden
              type="file"
              accept=".xlsx"
              onChange={(event) =>
                void loadAnalyticsFile(event.target.files?.[0])
              }
            />
          </Button>
          <Button
            variant="outlined"
            startIcon={<FileDownloadOutlinedIcon />}
            disabled={!analytics}
          >
            Експорт звіту
          </Button>
        </Stack>
      </header>
      {isLoading && <LinearProgress color="primary" />}
      <Alert
        severity={analytics ? "success" : "info"}
        variant="outlined"
        sx={{ mb: 2 }}
      >
        {message}
      </Alert>

      {analytics ? (
        <>
          <div
            className="analytics-tabs"
            role="tablist"
            aria-label="Тип аналітики"
          >
            <button
              className={analyticsTab === "full" ? "active" : ""}
              type="button"
              onClick={() => setAnalyticsTab("full")}
            >
              Повна аналітика
            </button>
            <button
              className={analyticsTab === "short" ? "active" : ""}
              type="button"
              onClick={() => setAnalyticsTab("short")}
            >
              Скорочена
            </button>
            <button
              className={analyticsTab === "hint" ? "active" : ""}
              type="button"
              onClick={() => setAnalyticsTab("hint")}
            >
              Підказка
            </button>
            <button
              className={analyticsTab === "text" ? "active" : ""}
              type="button"
              onClick={() => setAnalyticsTab("text")}
            >
              Текст
            </button>
          </div>
          <section className="analytics-metrics">
            {[
              ["За штатом", analytics.metrics.staff],
              ["За списком", analytics.metrics.listed],
              ["В строю", analytics.metrics.inRanks],
              ["Відсутні", analytics.metrics.absent],
            ].map(([label, value]) => (
              <div className="analytics-metric-card" key={label}>
                <span>{label}</span>
                <strong>{value}</strong>
              </div>
            ))}
            <div className="analytics-metric-card danger">
              <span>Некомплект</span>
              <strong>{analytics.metrics.understaff}</strong>
              <em>{Math.round(analytics.metrics.understaffPercent)}%</em>
            </div>
          </section>

          {analyticsTab === "short" ? (
            <ShortAnalytics analytics={analytics} />
          ) : analyticsTab === "hint" ? (
            <HintAnalytics analytics={analytics} />
          ) : analyticsTab === "text" ? (
            <TextReports analytics={analytics} />
          ) : (
            <section className="analytics-grid">
              <div className="analytics-panel analytics-wide">
                <div className="panel-heading">Укомплектованість</div>
                <div className="analytics-composition">
                  <DonutChart percent={analytics.metrics.staffedPercent} />
                  <BarList
                    items={[
                      { label: "Штат", value: analytics.metrics.staff },
                      { label: "Список", value: analytics.metrics.listed },
                      { label: "В строю", value: analytics.metrics.inRanks },
                    ]}
                    maxValue={analytics.metrics.staff}
                  />
                </div>
                <div className="analytics-note">
                  +{analytics.metrics.newcomersNoRole} новоприбулих без посад
                </div>
              </div>
              <div className="analytics-panel">
                <div className="panel-heading">
                  Причини відсутності · {analytics.metrics.absent}
                </div>
                <BarList items={analytics.absenceReasons} />
              </div>
              <div className="analytics-panel">
                <div className="panel-heading">
                  Відкомандировані ·{" "}
                  {analytics.detachedDetails.reduce(
                    (sum, item) => sum + item.value,
                    0,
                  )}
                </div>
                <BarList items={analytics.detachedDetails} />
              </div>
              <div className="analytics-panel">
                <div className="panel-heading">
                  Прикомандировані ·{" "}
                  {analytics.attachedDetails.reduce(
                    (sum, item) => sum + item.value,
                    0,
                  )}
                </div>
                <BarList items={analytics.attachedDetails} />
              </div>
              <div className="analytics-panel">
                <div className="panel-heading">Виконання завдань</div>
                <Typography className="big-number">
                  {analytics.execution.reduce(
                    (sum, item) => sum + item.value,
                    0,
                  )}
                </Typography>
                <BarList items={analytics.execution} />
              </div>
              <div className="analytics-panel">
                <div className="panel-heading">Деталі виконання</div>
                <BarList items={analytics.executionDetails} />
              </div>
              <div className="analytics-panel">
                <div className="panel-heading">Структура в строю</div>
                <div className="structure-grid">
                  {analytics.structure.map((item) => (
                    <div key={item.label}>
                      <strong>{item.value}</strong>
                      <span>{item.label}</span>
                    </div>
                  ))}
                </div>
              </div>
              <div className="analytics-panel analytics-wide">
                <div className="panel-heading">
                  Не включені у формулу строю ·{" "}
                  {analytics.unclassifiedInRanks.length}
                </div>
                {analytics.unclassifiedInRanks.length ? (
                  <div className="unclassified-list">
                    {analytics.unclassifiedInRanks.map((record) => (
                      <div
                        className="unclassified-record"
                        key={record.excelRowNumber}
                      >
                        <div>
                          <strong>
                            Рядок {record.excelRowNumber} ·{" "}
                            {record.name || "без ПІБ"}
                          </strong>
                          <span>
                            {[
                              record.roster,
                              record.rank,
                              record.callSign && `позивний ${record.callSign}`,
                              record.position,
                            ]
                              .filter(Boolean)
                              .join(" · ")}
                          </span>
                        </div>
                        <div className="unclassified-tags">
                          {[record.status, record.type, record.bg, record.bzvp, record.location]
                            .filter(Boolean)
                            .map((tag) => (
                              <em key={tag}>{tag}</em>
                            ))}
                        </div>
                        <p>{record.reason}</p>
                      </div>
                    ))}
                  </div>
                ) : (
                  <Typography variant="body2" color="text.secondary">
                    Усі записи зі статусом “В строю” входять у контрольні
                    категорії.
                  </Typography>
                )}
              </div>
              <div className="analytics-panel">
                <div className="panel-heading">
                  Управління ·{" "}
                  {analytics.managementDetails.reduce(
                    (sum, item) => sum + item.value,
                    0,
                  )}
                </div>
                <BarList items={analytics.managementDetails} />
              </div>
              <div className="analytics-panel">
                <div className="panel-heading">
                  Забезпечення ·{" "}
                  {analytics.supportDetails.reduce(
                    (sum, item) => sum + item.value,
                    0,
                  )}
                </div>
                <BarList items={analytics.supportDetails} />
              </div>
              <div className="analytics-panel">
                <div className="panel-heading">БЗВП</div>
                <BarList items={analytics.bzvp} />
              </div>
              <div className="analytics-panel">
                <div className="panel-heading">
                  Новоприбулі ·{" "}
                  {analytics.newcomerDetails.reduce(
                    (sum, item) => sum + item.value,
                    0,
                  )}
                </div>
                <BarList items={analytics.newcomerDetails} />
              </div>
              <div className="analytics-chart-row">
                <div className="analytics-panel chart-panel">
                  <HorizontalSvgChart
                    color="#ef4e3c"
                    items={analytics.absenceReasons}
                    title={`Причини відсутності · ${analytics.metrics.absent}`}
                    xLabel="Кількість"
                    yLabel="Причина"
                  />
                </div>
                <div className="analytics-panel chart-panel">
                  <VerticalSvgChart
                    color="#3aa0d8"
                    items={analytics.ageGroups}
                    title="Віковий розподіл"
                    xLabel="Вікова група"
                    yLabel="Кількість"
                  />
                </div>
                <div className="analytics-panel chart-panel">
                  <HorizontalSvgChart
                    color="#31c979"
                    items={[
                      ...analytics.structure,
                      {
                        label: "На виконанні",
                        value: analytics.execution.reduce(
                          (sum, item) => sum + item.value,
                          0,
                        ),
                      },
                      {
                        label: "Без БЗВП",
                        value: analytics.bzvp.reduce(
                          (sum, item) => sum + item.value,
                          0,
                        ),
                      },
                      {
                        label: "Новоприбулі",
                        value: analytics.metrics.newcomersNoRole,
                      },
                    ]}
                    title="Стрій та виконання"
                    xLabel="Кількість"
                    yLabel="Категорія"
                  />
                </div>
              </div>
            </section>
          )}

          <section className="panel table-panel">
            <div className="panel-heading">Дані Аркуш1</div>
            <div className="panel-body">
              <MaterialReactTable table={table} />
            </div>
          </section>
        </>
      ) : (
        <section className="panel table-panel">
          <div className="panel-body">
            <div className="drop-zone">
              <Box>
                <CloudUploadOutlinedIcon color="disabled" />
                <Typography variant="body2">
                  Додайте `.xlsx` строєвої записки
                </Typography>
                <Typography variant="caption" className="muted">
                  Перші 3 вкладки використовуються як контроль, основні
                  розрахунки виконуються по Аркуш1.
                </Typography>
              </Box>
            </div>
          </div>
        </section>
      )}
    </main>
  );
}
