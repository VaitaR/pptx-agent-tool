export const DATA_CHART_PATTERNS = [
  "line_trend",
  "column_trend",
  "ranked_bar",
  "stacked_composition",
  "grouped_bar",
] as const;

export const DATA_CHART_INTENTS = [
  "trend_recovery",
  "trend_decline",
  "level_monitoring",
  "rank_leader",
  "mix_concentration",
  "variance_watch",
  "composition_shift",
  "period_comparison",
] as const;

export type DataChartPattern = (typeof DATA_CHART_PATTERNS)[number];
export type DataChartIntent = (typeof DATA_CHART_INTENTS)[number];
export type DataChartMode = "bar" | "line" | "horizontal_bar";
export type DataChartRow = { t: unknown; series?: unknown; value?: unknown };
export type DataChartQualityInput = {
  rows: DataChartRow[];
  chart: DataChartMode;
  pattern?: DataChartPattern;
  intent?: DataChartIntent;
  valueFormat: "currency" | "number" | "percent";
  title: string;
  body?: string;
  takeaway?: string;
  unit?: string;
  period?: string;
  source?: string;
  altText?: string;
};

export function isDataChartPattern(value: unknown): value is DataChartPattern {
  return (
    typeof value === "string" &&
    (DATA_CHART_PATTERNS as readonly string[]).includes(value)
  );
}

export function isDataChartIntent(value: unknown): value is DataChartIntent {
  return (
    typeof value === "string" &&
    (DATA_CHART_INTENTS as readonly string[]).includes(value)
  );
}

export function isDateLike(value: unknown): boolean {
  const raw = String(value);
  return (
    /^\d{4}-\d{2}-\d{2}/.test(raw) &&
    !Number.isNaN(new Date(`${raw.slice(0, 10)}T00:00:00Z`).getTime())
  );
}

export function hasMultipleSeries(rows: DataChartRow[]): boolean {
  return (
    new Set(
      rows
        .filter((row) => row.series !== undefined && row.series !== null)
        .map((row) => String(row.series)),
    ).size > 1
  );
}

export function inferDataChartPattern(
  rows: DataChartRow[],
  chart: DataChartMode,
): DataChartPattern {
  const multi = hasMultipleSeries(rows);
  const allDates = rows.length > 0 && rows.every((row) => isDateLike(row.t));

  if (multi) return "stacked_composition";
  if (chart === "horizontal_bar" || !allDates) return "ranked_bar";
  if (chart === "line") return "line_trend";
  return "column_trend";
}

export function effectiveChartModeForPattern(
  pattern: DataChartPattern,
  requestedChart: DataChartMode,
): DataChartMode {
  if (pattern === "line_trend") return "line";
  if (pattern === "ranked_bar") return "horizontal_bar";
  if (pattern === "stacked_composition") return "bar";
  if (pattern === "column_trend") return "bar";
  if (pattern === "grouped_bar") return "bar";
  return requestedChart;
}

function numericValues(rows: DataChartRow[]): number[] {
  return rows
    .map((row) => Number(row.value ?? 0))
    .filter((value) => Number.isFinite(value));
}

function uniqueLabels(rows: DataChartRow[]): string[] {
  return Array.from(new Set(rows.map((row) => String(row.t))));
}

function uniqueSeries(rows: DataChartRow[]): string[] {
  return Array.from(
    new Set(
      rows
        .filter((row) => row.series !== undefined && row.series !== null)
        .map((row) => String(row.series)),
    ),
  );
}

function looksGenericTitle(title: string): boolean {
  const normalized = title.trim().toLowerCase();
  return [
    "revenue",
    "aum",
    "main metrics",
    "active users",
    "turnover",
    "product mix",
    "country revenue",
  ].includes(normalized);
}

function relativeRange(rows: DataChartRow[]): number {
  const values = numericValues(rows);
  if (values.length < 2) return 0;
  const min = Math.min(...values);
  const max = Math.max(...values);
  if (max <= 0) return 0;
  return (max - min) / max;
}

export function dataChartPatternWarnings(
  rows: DataChartRow[],
  chart: DataChartMode,
  requestedPattern?: DataChartPattern,
): string[] {
  const warnings: string[] = [];
  const inferredPattern = inferDataChartPattern(rows, chart);
  const effectivePattern = requestedPattern ?? inferredPattern;
  const multi = hasMultipleSeries(rows);
  const allDates = rows.length > 0 && rows.every((row) => isDateLike(row.t));

  if (!requestedPattern) {
    warnings.push(`data_chart is missing pattern; inferred ${inferredPattern}`);
  } else if (requestedPattern !== inferredPattern) {
    // grouped_bar is a valid explicit override of stacked_composition for multi-series
    const validOverride =
      requestedPattern === "grouped_bar" && inferredPattern === "stacked_composition";
    if (!validOverride) {
      warnings.push(
        `data_chart pattern ${requestedPattern} does not match inferred ${inferredPattern}`,
      );
    }
  }

  if (effectivePattern === "ranked_bar" && allDates && chart !== "horizontal_bar") {
    warnings.push("ranked_bar should use categorical labels, not only date buckets");
  }

  if (effectivePattern === "line_trend" && !allDates) {
    warnings.push("line_trend should use date-like t values");
  }

  if (effectivePattern === "stacked_composition" && !multi) {
    warnings.push("stacked_composition requires multiple series values");
  }

  if (effectivePattern === "grouped_bar" && !multi) {
    warnings.push("grouped_bar requires multiple series values");
  }

  if (
    multi &&
    effectivePattern !== "stacked_composition" &&
    effectivePattern !== "grouped_bar"
  ) {
    warnings.push("multi-series data should use stacked_composition or grouped_bar");
  }

  return warnings;
}

export function dataChartIntentWarnings(
  rows: DataChartRow[],
  pattern: DataChartPattern,
  intent?: DataChartIntent,
): string[] {
  const warnings: string[] = [];
  const multi = hasMultipleSeries(rows);
  const allDates = rows.length > 0 && rows.every((row) => isDateLike(row.t));
  const range = relativeRange(rows);

  if (!intent) {
    warnings.push("data_chart is missing intent; choose the analytical job before rendering");
    return warnings;
  }

  if (
    pattern === "ranked_bar" &&
    intent !== "rank_leader" &&
    intent !== "mix_concentration"
  ) {
    warnings.push("ranked_bar intent should be rank_leader or mix_concentration");
  }

  if (
    pattern === "stacked_composition" &&
    intent !== "composition_shift" &&
    intent !== "mix_concentration"
  ) {
    warnings.push("stacked_composition intent should be composition_shift or mix_concentration");
  }

  if (
    pattern === "line_trend" &&
    intent !== "trend_recovery" &&
    intent !== "trend_decline" &&
    intent !== "level_monitoring" &&
    intent !== "variance_watch"
  ) {
    warnings.push("line_trend intent should describe trend direction, level monitoring, or variance");
  }

  if (
    pattern === "column_trend" &&
    intent !== "trend_recovery" &&
    intent !== "trend_decline" &&
    intent !== "variance_watch"
  ) {
    warnings.push("column_trend intent should describe discrete volume change or variance");
  }

  if (
    pattern === "grouped_bar" &&
    intent !== "period_comparison" &&
    intent !== "rank_leader" &&
    intent !== "mix_concentration"
  ) {
    warnings.push("grouped_bar intent should be period_comparison, rank_leader, or mix_concentration");
  }

  if (
    pattern === "column_trend" &&
    allDates &&
    range < 0.22 &&
    intent === "level_monitoring"
  ) {
    warnings.push("level_monitoring over dates should usually use line_trend, not column_trend");
  }

  if (
    multi &&
    intent !== "composition_shift" &&
    intent !== "mix_concentration" &&
    intent !== "period_comparison" &&
    intent !== "rank_leader"
  ) {
    warnings.push(
      "multi-series charts need composition_shift, mix_concentration, period_comparison, or rank_leader intent",
    );
  }

  return warnings;
}

export function dataChartQualityWarnings(input: DataChartQualityInput): string[] {
  const warnings: string[] = [];
  const pattern = input.pattern ?? inferDataChartPattern(input.rows, input.chart);
  const labels = uniqueLabels(input.rows);
  const series = uniqueSeries(input.rows);
  const values = numericValues(input.rows);
  const hasNarrativeText = Boolean(input.takeaway?.trim() || input.body?.trim());

  if (!hasNarrativeText) {
    warnings.push("data_chart should include takeaway or short body text");
  }
  if (!input.unit?.trim()) {
    warnings.push("data_chart should declare unit for footer/source discipline");
  }
  if (!input.period?.trim()) {
    warnings.push("data_chart should declare period for reader context");
  }
  if (!input.source?.trim()) {
    warnings.push("data_chart should include source text");
  }
  if (!input.altText?.trim()) {
    warnings.push("data_chart should include altText or speaker-note summary");
  } else if (input.altText.length > 250) {
    warnings.push("data_chart altText should stay under 250 characters");
  }
  if (looksGenericTitle(input.title)) {
    warnings.push("data_chart title is generic; use a claim title instead");
  }

  if (pattern === "ranked_bar") {
    if (labels.length > 10) {
      warnings.push(
        `ranked_bar has ${labels.length} categories; group the long tail before rendering`,
      );
    }
    if (labels.some((label) => label.length > 34)) {
      warnings.push("ranked_bar has long labels; shorten labels or use appendix table");
    }
  }

  if (pattern === "column_trend") {
    if (labels.length > 12) {
      warnings.push(
        `column_trend has ${labels.length} periods; use line_trend or reduce visible periods`,
      );
    }
    if (series.length > 1) {
      warnings.push("column_trend should not carry multiple series");
    }
  }

  if (pattern === "line_trend") {
    if (labels.length < 3) {
      warnings.push("line_trend needs at least 3 time points to show a trend");
    }
    if (labels.length > 24) {
      warnings.push(
        `line_trend has ${labels.length} points; aggregate or use sparse labels`,
      );
    }
    if (series.length > 3) {
      warnings.push("line_trend supports at most 3 foreground series");
    }
  }

  if (pattern === "stacked_composition") {
    if (series.length > 5) {
      warnings.push(
        `stacked_composition has ${series.length} series; group minor series into Other`,
      );
    }
    if (labels.length > 8) {
      warnings.push(
        `stacked_composition has ${labels.length} periods; use fewer periods or an appendix`,
      );
    }
    if (
      input.intent === "composition_shift" &&
      input.valueFormat !== "percent" &&
      input.unit &&
      !/%|share|mix/i.test(input.unit)
    ) {
      warnings.push("composition_shift should state share/mix as the displayed unit");
    }
  }

  if (pattern === "grouped_bar") {
    if (series.length > 3) {
      warnings.push(
        `grouped_bar has ${series.length} series; use at most 3 for readable side-by-side bars`,
      );
    }
    if (labels.length > 8) {
      warnings.push(
        `grouped_bar has ${labels.length} categories; use at most 8 or move extras to appendix`,
      );
    }
  }

  if (
    input.valueFormat === "percent" &&
    values.length > 0 &&
    Math.max(...values) <= 1
  ) {
    warnings.push("percent values should use 0-100 scale, not 0-1 fractions");
  }

  return warnings;
}
