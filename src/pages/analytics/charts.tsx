import * as d3 from "d3";
import type { AnalyticsMetric } from "./analyticsData";

export function DonutChart({ percent }: { percent: number }) {
  const size = 176;
  const radius = 78;
  const arc = d3
    .arc<d3.DefaultArcObject>()
    .innerRadius(58)
    .outerRadius(radius)
    .cornerRadius(4);
  const foreground = arc({
    innerRadius: 58,
    outerRadius: radius,
    startAngle: 0,
    endAngle: (Math.PI * 2 * percent) / 100,
  });
  const background = arc({
    innerRadius: 58,
    outerRadius: radius,
    startAngle: 0,
    endAngle: Math.PI * 2,
  });

  return (
    <svg className="analytics-donut" viewBox={`0 0 ${size} ${size}`} role="img">
      <g transform={`translate(${size / 2}, ${size / 2})`}>
        <path d={background ?? undefined} fill="rgba(230,224,190,0.12)" />
        <path d={foreground ?? undefined} fill="#9bbb55" />
      </g>
      <text x="50%" y="48%" textAnchor="middle" className="donut-value">
        {Math.round(percent)}%
      </text>
      <text x="50%" y="62%" textAnchor="middle" className="donut-label">
        укомплектованість
      </text>
    </svg>
  );
}

export function BarList({
  items,
  maxValue,
}: {
  items: AnalyticsMetric[];
  maxValue?: number;
}) {
  const max = maxValue ?? Math.max(1, ...items.map((item) => item.value));
  const scale = d3.scaleLinear().domain([0, max]).range([0, 100]);

  return (
    <div className="analytics-bars">
      {items.map((item) => (
        <div className="analytics-bar-row" key={item.label}>
          <span>{item.label}</span>
          <div
            className={`analytics-bar-track ${item.tone === "bad" ? "bad" : ""}`}
          >
            <div style={{ width: `${scale(item.value)}%` }} />
          </div>
          <strong>{item.value}</strong>
          {item.detail && <small>{item.detail}</small>}
        </div>
      ))}
    </div>
  );
}

export function BchsAbsenceDonut({
  items,
  total,
}: {
  items: AnalyticsMetric[];
  total: number;
}) {
  const size = 220;
  const radius = 96;
  const arc = d3
    .arc<d3.PieArcDatum<AnalyticsMetric>>()
    .innerRadius(48)
    .outerRadius(radius)
    .cornerRadius(3);
  const pie = d3
    .pie<AnalyticsMetric>()
    .value((item) => item.value)
    .sort(null);
  const colors = [
    "#7f9f43",
    "#ead15d",
    "#e3a128",
    "#b56549",
    "#8aa0a0",
    "#9a9a91",
  ];

  return (
    <svg
      className="bchs-donut-chart"
      viewBox={`0 0 ${size} ${size}`}
      role="img"
    >
      <g transform={`translate(${size / 2}, ${size / 2})`}>
        {pie(items).map((slice, index) => (
          <path
            d={arc(slice) ?? undefined}
            fill={colors[index % colors.length]}
            key={slice.data.label}
          />
        ))}
      </g>
      <text x="50%" y="49%" textAnchor="middle" className="bchs-donut-total">
        {total}
      </text>
      <text x="50%" y="62%" textAnchor="middle" className="bchs-donut-label">
        всього
      </text>
    </svg>
  );
}

export function clipChartLabel(label: string, maxLength = 18) {
  return label.length > maxLength
    ? `${label.slice(0, maxLength - 1)}...`
    : label;
}

export function HorizontalSvgChart({
  title,
  items,
  color,
  xLabel,
  yLabel,
}: {
  title: string;
  items: AnalyticsMetric[];
  color: string;
  xLabel: string;
  yLabel: string;
}) {
  const width = 760;
  const height = Math.max(280, items.length * 34 + 112);
  const margin = { top: 42, right: 54, bottom: 48, left: 170 };
  const plotWidth = width - margin.left - margin.right;
  const plotHeight = height - margin.top - margin.bottom;
  const maxValue = Math.max(1, ...items.map((item) => item.value));
  const x = d3.scaleLinear().domain([0, maxValue]).nice().range([0, plotWidth]);
  const y = d3
    .scaleBand()
    .domain(items.map((item) => item.label))
    .range([0, plotHeight])
    .padding(0.26);
  const ticks = x.ticks(5);

  return (
    <svg
      className="analytics-svg-chart"
      viewBox={`0 0 ${width} ${height}`}
      role="img"
      aria-label={title}
    >
      <text className="chart-title" x={20} y={28}>
        {title}
      </text>
      <g transform={`translate(${margin.left}, ${margin.top})`}>
        {ticks.map((tick) => (
          <g key={tick} transform={`translate(${x(tick)}, 0)`}>
            <line className="chart-grid-line" y2={plotHeight} />
            <text
              className="chart-axis-text"
              y={plotHeight + 22}
              textAnchor="middle"
            >
              {tick}
            </text>
          </g>
        ))}
        {items.map((item) => {
          const yPosition = y(item.label) ?? 0;

          return (
            <g key={item.label}>
              <text
                className="chart-axis-text"
                x={-12}
                y={yPosition + y.bandwidth() / 2 + 4}
                textAnchor="end"
              >
                {clipChartLabel(item.label)}
              </text>
              <rect
                className="chart-bar"
                x={0}
                y={yPosition}
                width={x(item.value)}
                height={y.bandwidth()}
                rx={3}
                fill={color}
              />
              <text
                className="chart-value"
                x={x(item.value) + 8}
                y={yPosition + y.bandwidth() / 2 + 4}
              >
                {item.value}
              </text>
            </g>
          );
        })}
        <text
          className="chart-axis-label"
          x={plotWidth / 2}
          y={plotHeight + 44}
          textAnchor="middle"
        >
          {xLabel}
        </text>
        <text
          className="chart-axis-label"
          transform={`translate(${-142}, ${plotHeight / 2}) rotate(-90)`}
          textAnchor="middle"
        >
          {yLabel}
        </text>
      </g>
    </svg>
  );
}

export function VerticalSvgChart({
  title,
  items,
  color,
  xLabel,
  yLabel,
}: {
  title: string;
  items: AnalyticsMetric[];
  color: string;
  xLabel: string;
  yLabel: string;
}) {
  const width = 760;
  const height = 360;
  const margin = { top: 44, right: 34, bottom: 58, left: 76 };
  const plotWidth = width - margin.left - margin.right;
  const plotHeight = height - margin.top - margin.bottom;
  const maxValue = Math.max(1, ...items.map((item) => item.value));
  const x = d3
    .scaleBand()
    .domain(items.map((item) => item.label))
    .range([0, plotWidth])
    .padding(0.36);
  const y = d3
    .scaleLinear()
    .domain([0, maxValue])
    .nice()
    .range([plotHeight, 0]);
  const ticks = y.ticks(5);

  return (
    <svg
      className="analytics-svg-chart"
      viewBox={`0 0 ${width} ${height}`}
      role="img"
      aria-label={title}
    >
      <text className="chart-title" x={20} y={30}>
        {title}
      </text>
      <g transform={`translate(${margin.left}, ${margin.top})`}>
        {ticks.map((tick) => (
          <g key={tick} transform={`translate(0, ${y(tick)})`}>
            <line className="chart-grid-line" x2={plotWidth} />
            <text className="chart-axis-text" x={-12} y={4} textAnchor="end">
              {tick}
            </text>
          </g>
        ))}
        {items.map((item) => {
          const xPosition = x(item.label) ?? 0;
          const barHeight = plotHeight - y(item.value);

          return (
            <g key={item.label}>
              <rect
                className="chart-bar"
                x={xPosition}
                y={y(item.value)}
                width={x.bandwidth()}
                height={barHeight}
                rx={4}
                fill={color}
              />
              <text
                className="chart-value"
                x={xPosition + x.bandwidth() / 2}
                y={Math.max(12, y(item.value) - 8)}
                textAnchor="middle"
              >
                {item.value}
              </text>
              <text
                className="chart-axis-text"
                x={xPosition + x.bandwidth() / 2}
                y={plotHeight + 24}
                textAnchor="middle"
              >
                {item.label}
              </text>
            </g>
          );
        })}
        <text
          className="chart-axis-label"
          x={plotWidth / 2}
          y={plotHeight + 52}
          textAnchor="middle"
        >
          {xLabel}
        </text>
        <text
          className="chart-axis-label"
          transform={`translate(${-54}, ${plotHeight / 2}) rotate(-90)`}
          textAnchor="middle"
        >
          {yLabel}
        </text>
      </g>
    </svg>
  );
}
