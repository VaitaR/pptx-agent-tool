import fs from "node:fs";
import { createRequire } from "node:module";
import path from "node:path";
import { createHash } from "node:crypto";
import { fileURLToPath } from "node:url";
import {
  effectiveChartModeForPattern,
  inferDataChartPattern,
  isDateLike,
  type DataChartIntent,
  type DataChartPattern,
} from "./chart-patterns.js";
import { readDeckSpec, type DeckSpec } from "./deck-spec.js";
import { layouts as layoutConfig, type Box } from "./layouts.js";
import { theme as themeConfig } from "./theme.js";

const require = createRequire(import.meta.url);

type PptxSlide = {
  background: { color: string };
  addText: (text: string, options: Record<string, unknown>) => void;
  addImage: (options: Record<string, unknown>) => void;
  addShape: (shapeName: string, options: Record<string, unknown>) => void;
};
type PptxPresentation = {
  layout: string;
  author: string;
  subject: string;
  title: string;
  company: string;
  lang: string;
  theme: Record<string, unknown>;
  defineLayout: (layout: {
    name: string;
    width: number;
    height: number;
  }) => void;
  addSlide: () => PptxSlide;
  writeFile: (options: { fileName: string }) => Promise<void>;
};

const PptxGenJS = require("pptxgenjs") as new () => PptxPresentation;

type QueryRow = Record<string, string | number | boolean | null>;

const rootDir = process.cwd();

const layouts = layoutConfig;
const theme = themeConfig.colors;
const font = themeConfig.fonts;
const typeScale = themeConfig.typography;
const seriesOrder: readonly string[] = themeConfig.seriesOrder;
const chartColors = themeConfig.chartColors;

const brandDefaults = {
  logo: { x: 1.36, y: 0.78, mark: 0.52, textW: 1.0 },
  contentTitle: { x: 1.02, y: 0.52, w: 7.3, h: 0.36 },
  contentBody: { x: 1.04, y: 0.95, w: 8.0, h: 0.22 },
  footer: { x: 1.02, y: 6.98, w: 6.5, h: 0.14 },
} as const;

export type RenderDeckOptions = {
  specPath?: string;
  outputPath?: string;
  artifactsDir?: string;
  quiet?: boolean;
};

export type RenderReport = {
  status: "ok";
  specPath: string;
  pptxPath: string;
  artifactsDir: string;
  slideCount: number;
  slides: Array<{
    index: number;
    type: DeckSpec["slides"][number]["type"];
    title: string;
  }>;
  manifestPath: string;
  artifacts: {
    sql: string[];
    data: string[];
  };
};

export type RenderManifest = {
  specPath: string;
  pptxPath: string;
  specHash: string;
  pptxHash: string;
  artifacts: {
    sql: Array<{ metric: string; path: string; hash: string }>;
    data: Array<{ metric: string; path: string; hash: string }>;
  };
  slides: Array<{
    index: number;
    type: DeckSpec["slides"][number]["type"];
    title: string;
    chartMode?: "bar" | "line" | "horizontal_bar";
    chartPattern?: DataChartPattern;
    chartIntent?: DataChartIntent;
    sourceText?: string;
  }>;
};

function ensureSampleAssets(): void {
  const assetsDir = path.join(rootDir, "assets");
  fs.mkdirSync(assetsDir, { recursive: true });

  const pngByName: Record<string, string> = {
    "sample-metric.png":
      "iVBORw0KGgoAAAANSUhEUgAAAlgAAAGQCAYAAAByNR6YAAAACXBIWXMAAAsTAAALEwEAmpwYAAAKTWlDQ1BQaG90b3Nob3AgSUNDIHByb2ZpbGUAAHja7J1tdFTVksc/5xSIAKKHEJBqRRQvYFAUCpKIgCAS9hYFFBFRAW3BVSxUQFBZQLjAFmuLnfg+vTQJXDv31G3PiR+ZnXsFih8zw2ERiQjr68yM5HnzYv5k8tzY1pl0vC+/Zbba2v+z1tprrWXW9mVwAABtSURBVHja7N1BCgIxDAXQ//9y9lR0g1IYRDCwKclcZ58E+kwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD4Z3bbtg8Yc+S4gQAAAFgV4AEAAKwmwAMAAFhNgAcAALCaAA8AAGC1AR4AAMBqAjwAAIDVBHgAAACrCfAAAABWE+ABAAAsJsADAABYTYAHAAAwmsx9gB/Yln3b5VwAAAAASUVORK5CYII=",
    "sample-comparison.png":
      "iVBORw0KGgoAAAANSUhEUgAAAlgAAAGQCAYAAAByNR6YAAAACXBIWXMAAAsTAAALEwEAmpwYAAAHfElEQVR4nO3bMQ0AIBDAMMC/5+ONAvZoFSzZ0mQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAB8x+7a9gEAAAAAAAAAAABWCHgAAACrCfAAAABWE+ABAAAsJsADAABYTYAHAAAwmgAPAAFgNQEeAADAagI8AACAjb/fARuVUzZJAAAAAElFTkSuQmCC",
  };

  for (const [name, base64] of Object.entries(pngByName)) {
    const assetPath = path.join(assetsDir, name);
    if (!fs.existsSync(assetPath)) {
      fs.writeFileSync(assetPath, Buffer.from(base64, "base64"));
    }
  }
}

function contentTitleFontSize(title: string): number {
  if (title.length > 74) return typeScale.title.compact;
  if (title.length > 54) return typeScale.title.medium;
  return typeScale.title.large;
}

function addDeckTitle(slide: PptxSlide, title: string, box: Box): void {
  const fontSize = contentTitleFontSize(title);
  slide.addText(title, {
    ...box,
    h: Math.max(box.h, title.length > 54 ? 0.88 : box.h),
    fontFace: font.head,
    fontSize,
    bold: true,
    color: theme.mainBlue,
    breakLine: false,
    fit: "shrink",
    margin: 0,
  });
}

function addTopRule(slide: PptxSlide, x: number, y: number, w: number): void {
  slide.addShape("rect", {
    x,
    y,
    w,
    h: 0.035,
    fill: { color: theme.mainBlue },
    line: { color: theme.mainBlue, transparency: 100 },
  });
}

function addPanel(
  slide: PptxSlide,
  box: Box,
  options: { fill?: string; line?: string; radius?: number } = {},
): void {
  const fill = options.fill ?? theme.white;
  const line = options.line ?? theme.faint;
  slide.addShape("roundRect", {
    ...box,
    rectRadius: options.radius ?? themeConfig.spacing.cardRadius,
    fill: { color: fill },
    line: { color: line, width: 0.8 },
  });
}

function headerShift(title: string, body?: string): number {
  if (title.length > 70) return body ? 0.34 : 0.24;
  if (title.length > 52) return body ? 0.24 : 0.14;
  return 0;
}

function shiftedBox(box: Box, dy: number): Box {
  return { ...box, y: box.y + dy };
}

function shiftedObjectBox(box: Box, dy: number): Box {
  return { ...box, y: box.y + dy, h: Math.max(0.5, box.h - dy) };
}

function coverTitleFontSize(title: string): number {
  if (title.length > 82) return 31;
  if (title.length > 64) return 34;
  if (title.length > 48) return 38;
  return 42;
}

type TextSegment = { text: string; options?: Record<string, unknown> };

function parseMarkdownToSegments(text: string): TextSegment[] {
  const segments: TextSegment[] = [];
  const lines = text.split("\n");

  for (const line of lines) {
    const trimmed = line.trim();
    if (!trimmed) {
      segments.push({ text: "\n" });
      continue;
    }

    const bulletMatch = trimmed.match(/^[-*]\s+(.*)/);
    const content = bulletMatch ? bulletMatch[1] : trimmed;
    if (bulletMatch) {
      segments.push({
        text: "•  ",
        options: { fontSize: 9.5, color: theme.muted },
      });
    }

    const boldRegex = /\*\*(.+?)\*\*/g;
    let lastIndex = 0;
    let match: RegExpExecArray | null;
    while ((match = boldRegex.exec(content)) !== null) {
      if (match.index > lastIndex) {
        segments.push({ text: content.slice(lastIndex, match.index) });
      }
      segments.push({
        text: match[1],
        options: { bold: true, color: theme.text },
      });
      lastIndex = match.index + match[0].length;
    }
    if (lastIndex < content.length) {
      segments.push({ text: content.slice(lastIndex) });
    }
    segments.push({ text: "\n" });
  }

  return segments;
}

function addBody(slide: PptxSlide, text: string | undefined, box: Box): void {
  if (!text) return;

  const hasMarkdown = text.includes("**") || /^[-*]\s/m.test(text);
  if (!hasMarkdown) {
    slide.addText(text, {
      ...box,
      fontFace: font.body,
      fontSize: text.length > 180 ? typeScale.body.compact : typeScale.body.default,
      color: theme.muted,
      breakLine: false,
      fit: "shrink",
      valign: "top",
      margin: 0.04,
    });
    return;
  }

  const segments = parseMarkdownToSegments(text);
  const richText = segments.map((seg) => ({
    text: seg.text,
      options: {
        fontFace: font.body,
        fontSize: typeScale.body.default,
        color: theme.muted,
        ...seg.options,
      },
  }));

  slide.addText(richText as unknown as string, {
    ...box,
    fit: "shrink",
    valign: "top",
    margin: 0.04,
  });
}

function addSource(slide: PptxSlide, text: string, box: Box): void {
  slide.addText(text, {
    ...box,
    fontFace: font.body,
    fontSize: typeScale.source,
    color: theme.muted,
    margin: 0,
  });
}

function addCoverBackground(slide: PptxSlide): void {
  slide.background = { color: theme.white };
  slide.addShape("rect", {
    x: 6.45,
    y: 0,
    w: 6.88,
    h: 7.5,
    fill: { color: theme.coverBlue, transparency: 18 },
    line: { color: theme.coverBlue, transparency: 100 },
  });
  slide.addShape("rect", {
    x: 9.45,
    y: 0,
    w: 3.88,
    h: 7.5,
    fill: { color: theme.coverViolet, transparency: 16 },
    line: { color: theme.coverViolet, transparency: 100 },
  });
  slide.addShape("rect", {
    x: 5.15,
    y: 0,
    w: 2.15,
    h: 7.5,
    fill: { color: "EBF5FF", transparency: 30 },
    line: { color: "EAF7FF", transparency: 100 },
  });
}

function toChartLabel(value: unknown): string {
  const raw = String(value);
  const date = new Date(`${raw}T00:00:00Z`);
  if (Number.isNaN(date.getTime())) return raw;
  return date.toLocaleDateString("en-US", {
    month: "short",
    day: "2-digit",
    timeZone: "UTC",
  });
}

function toPeriodLabel(value: unknown): string {
  const raw = String(value);
  const date = new Date(`${raw}T00:00:00Z`);
  if (Number.isNaN(date.getTime())) return raw;
  const month = date.toLocaleDateString("en-US", {
    month: "short",
    timeZone: "UTC",
  });
  const day = date.getUTCDate();
  return day === 1 ? month : `${month} ${String(day).padStart(2, "0")}`;
}

function compactNumber(value: number): string {
  const abs = Math.abs(value);
  if (abs >= 1_000_000_000) return `${(value / 1_000_000_000).toFixed(1)}B`;
  if (abs >= 1_000_000) return `${(value / 1_000_000).toFixed(1)}M`;
  if (abs >= 1_000) return `${(value / 1_000).toFixed(0)}K`;
  return value.toFixed(0);
}

function fullCurrency(value: number): string {
  return `$${Math.round(value).toLocaleString("en-US")}.00`;
}

function formatCurrency(value: number): string {
  return `$${compactNumber(value)}`;
}

function formatDelta(start: number, end: number): string {
  if (!start) return "n/a";
  const delta = ((end - start) / start) * 100;
  const sign = delta > 0 ? "+" : "";
  return `${sign}${delta.toFixed(1)}%`;
}

function intentKpiLabel(intent?: DataChartIntent): string {
  if (intent === "trend_recovery") return "Recovery";
  if (intent === "trend_decline") return "Decline";
  if (intent === "variance_watch") return "Variance";
  return "Change";
}

function latestValue(rows: QueryRow[], column = "value"): number {
  const row = rows.at(-1);
  if (!row) return 0;
  return Number(row[column] ?? 0);
}

function totalValue(rows: QueryRow[], column = "value"): number {
  return rows.reduce((sum, row) => sum + Number(row[column] ?? 0), 0);
}

function hasSeries(rows: QueryRow[]): boolean {
  return (
    new Set(
      rows
        .filter((row) => row.series !== undefined && row.series !== null)
        .map((row) => String(row.series)),
    ).size > 1
  );
}

function buildMultiSeries(
  rows: QueryRow[],
): { name: string; labels: string[]; values: number[] }[] {
  const labels = Array.from(new Set(rows.map((row) => String(row.t)))).sort();
  const series = Array.from(
    new Set(rows.map((row) => String(row.series))),
  ).sort((a, b) => {
    const aIndex = seriesOrder.indexOf(a);
    const bIndex = seriesOrder.indexOf(b);
    if (aIndex === -1 && bIndex === -1) return a.localeCompare(b);
    if (aIndex === -1) return 1;
    if (bIndex === -1) return -1;
    return aIndex - bIndex;
  });

  return series.map((name) => ({
    name,
    labels: labels.map(toChartLabel),
    values: labels.map((label) => {
      const row = rows.find(
        (candidate) =>
          String(candidate.t) === label && String(candidate.series) === name,
      );
      return Number(row?.value ?? 0);
    }),
  }));
}

function addKpi(
  slide: PptxSlide,
  label: string,
  value: string,
  box: Box,
  color: string = theme.title,
): void {
  slide.addText(label.toUpperCase(), {
    x: box.x,
    y: box.y,
    w: box.w,
    h: 0.15,
    fontFace: font.body,
    fontSize: 5.8,
    bold: true,
    color: theme.muted,
    margin: 0,
  });
  slide.addText(value, {
    x: box.x,
    y: box.y + 0.18,
    w: box.w,
    h: Math.max(0.32, box.h - 0.18),
    fontFace: font.head,
    fontSize: 15,
    bold: true,
    color,
    fit: "shrink",
    margin: 0,
  });
}

function intentBadgeSummary(
  rows: QueryRow[],
  intent: DataChartIntent | undefined,
  formatter: (value: number) => string,
): { label: string; value: string } | null {
  if (!intent || rows.length === 0) return null;

  const values = rows.map((row) => Number(row.value ?? 0));
  const first = values[0] ?? 0;
  const latest = values.at(-1) ?? 0;
  const max = Math.max(...values);
  const min = Math.min(...values);

  if (intent === "trend_recovery") {
    return { label: "Recovery", value: formatDelta(first, latest) };
  }
  if (intent === "trend_decline") {
    return { label: "Decline", value: formatDelta(first, latest) };
  }
  if (intent === "level_monitoring") {
    return { label: "Latest", value: formatter(latest) };
  }
  if (intent === "variance_watch") {
    return { label: "Range", value: formatter(max - min) };
  }

  return null;
}

function addIntentBadge(
  slide: PptxSlide,
  box: Box,
  rows: QueryRow[],
  intent: DataChartIntent | undefined,
  formatter: (value: number) => string,
): void {
  const summary = intentBadgeSummary(rows, intent, formatter);
  if (!summary) return;

  const badge = {
    x: box.x + 0.28,
    y: box.y + 0.08,
    w: 1.38,
    h: 0.48,
  };
  slide.addShape("roundRect", {
    ...badge,
    rectRadius: 0.06,
    fill: { color: theme.faintBlue },
    line: { color: "D9ECFF", width: 0.7 },
  });
  slide.addText(summary.label.toUpperCase(), {
    x: badge.x + 0.12,
    y: badge.y + 0.07,
    w: badge.w - 0.24,
    h: 0.12,
    fontFace: font.body,
    fontSize: 5.3,
    bold: true,
    color: theme.muted,
    margin: 0,
  });
  slide.addText(summary.value, {
    x: badge.x + 0.12,
    y: badge.y + 0.22,
    w: badge.w - 0.24,
    h: 0.18,
    fontFace: font.head,
    fontSize: 9.4,
    bold: true,
    color: theme.mainBlue,
    fit: "shrink",
    margin: 0,
  });
}

function addAxis(
  slide: PptxSlide,
  box: Box,
  min: number,
  max: number,
  formatter = compactNumber,
): void {
  slide.addShape("line", {
    x: box.x,
    y: box.y + box.h,
    w: box.w,
    h: 0,
    line: { color: "C8CFDA", width: 1 },
  });
  for (let i = 0; i <= 4; i += 1) {
    const y = box.y + (box.h * i) / 4;
    slide.addShape("line", {
      x: box.x,
      y,
      w: box.w,
      h: 0,
      line: { color: theme.faint, width: 0.5, transparency: i === 4 ? 100 : 0 },
    });
    const value = max - ((max - min) * i) / 4;
    slide.addText(formatter(value), {
      x: box.x + box.w + 0.05,
      y: y - 0.06,
      w: 0.55,
      h: 0.13,
      fontFace: font.body,
      fontSize: 5.8,
      color: theme.muted,
      margin: 0,
    });
  }
}

function addCurrencyYAxis(
  slide: PptxSlide,
  box: Box,
  min: number,
  max: number,
): void {
  for (let i = 0; i <= 5; i += 1) {
    const y = box.y + (box.h * i) / 5;
    const value = max - ((max - min) * i) / 5;
    slide.addShape("line", {
      x: box.x,
      y,
      w: box.w,
      h: 0,
      line: { color: theme.faint, width: 0.8 },
    });
    slide.addText(formatCurrency(value), {
      x: box.x - 1.25,
      y: y - 0.08,
      w: 1.05,
      h: 0.16,
      fontFace: font.body,
      fontSize: 7.8,
      color: theme.muted,
      align: "right",
      margin: 0,
    });
  }
  slide.addShape("line", {
    x: box.x,
    y: box.y + box.h,
    w: box.w,
    h: 0,
    line: { color: "6B7280", width: 1 },
  });
}

function niceCurrencyMax(value: number): number {
  if (value <= 0) return 1;
  const step = value > 100_000 ? 50_000 : value > 50_000 ? 25_000 : 10_000;
  return Math.ceil(value / step) * step;
}

function dateParts(value: string): { day: number; month: string; iso: string } {
  const normalized = value.includes(" ")
    ? value.replace(" ", "T")
    : `${value}T00:00:00`;
  const date = new Date(`${normalized}Z`);
  return {
    day: date.getUTCDate(),
    month: date.toLocaleDateString("en-US", {
      month: "short",
      timeZone: "UTC",
    }),
    iso: value,
  };
}

function estimateLegendItemWidth(text: string): number {
  const charWidth = 0.063;
  const swatchAndPadding = 0.34;
  return Math.max(0.6, text.length * charWidth + swatchAndPadding);
}

function addShapeLegend(
  slide: PptxSlide,
  items: string[],
  x: number,
  y: number,
): void {
  const widths = items.map(estimateLegendItemWidth);
  const totalWidth = widths.reduce((sum, w) => sum + w, 0);
  const maxWidth = 11.0;
  const scale = totalWidth > maxWidth ? maxWidth / totalWidth : 1;

  items.forEach((item, index) => {
    const itemX =
      x + widths.slice(0, index).reduce((sum, w) => sum + w * scale, 0);
    const itemW = widths[index] * scale;
    const itemY = y;
    slide.addShape("rect", {
      x: itemX,
      y: itemY + 0.035,
      w: 0.11,
      h: 0.11,
      fill: { color: chartColors[index % chartColors.length] },
      line: {
        color: chartColors[index % chartColors.length],
        transparency: 100,
      },
    });
    slide.addText(item, {
      x: itemX + 0.16,
      y: itemY,
      w: Math.max(0.3, itemW - 0.2),
      h: 0.16,
      fontFace: font.body,
      fontSize: 7.0,
      color: theme.muted,
      margin: 0,
    });
  });
}

function niceMax(value: number): number {
  if (value <= 0) return 1;
  const magnitude = Math.pow(10, Math.floor(Math.log10(value)));
  const normalized = value / magnitude;
  const steps = [1, 1.5, 2, 2.5, 3, 4, 5, 7.5, 10];
  const step = steps.find((s) => s >= normalized) ?? 10;
  return step * magnitude;
}

function addStackedBarShapeChart(
  slide: PptxSlide,
  rows: QueryRow[],
  box: Box,
  formatter: (value: number) => string = formatCurrency,
  mode: "absolute" | "percent" = "absolute",
): void {
  const series = buildMultiSeries(rows);
  const rawLabels = Array.from(
    new Set(rows.map((row) => String(row.t))),
  ).sort();
  const totals = rawLabels.map((_, index) =>
    series.reduce((sum, item) => sum + item.values[index], 0),
  );
  const maxTotal = niceMax(Math.max(...totals, 1));
  const usePercentStack = mode === "percent";
  const useCurrencyAxis = formatter === formatCurrency && !usePercentStack;
  const plot = usePercentStack
    ? { x: box.x + 0.72, y: box.y + 0.64, w: box.w - 1.22, h: box.h - 1.42 }
    : useCurrencyAxis
      ? { x: box.x + 1.25, y: box.y + 0.55, w: box.w - 1.85, h: box.h - 1.35 }
      : { x: box.x + 0.5, y: box.y + 0.55, w: box.w - 0.95, h: box.h - 1.35 };
  const step = plot.w / Math.max(rawLabels.length, 1);
  const barW = Math.max(0.14, Math.min(step * 0.72, step - 0.14));

  if (usePercentStack) {
    for (let i = 0; i <= 2; i += 1) {
      const y = plot.y + (plot.h * i) / 2;
      const value = 100 - i * 50;
      slide.addShape("line", {
        x: plot.x,
        y,
        w: plot.w,
        h: 0,
        line: { color: theme.faint, width: 0.65 },
      });
      slide.addText(`${value}%`, {
        x: plot.x - 0.48,
        y: y - 0.07,
        w: 0.36,
        h: 0.14,
        fontFace: font.body,
        fontSize: 6.8,
        color: theme.muted,
        align: "right",
        margin: 0,
      });
    }
    slide.addShape("line", {
      x: plot.x,
      y: plot.y + plot.h,
      w: plot.w,
      h: 0,
      line: { color: theme.axis, width: 0.9 },
    });
  } else if (useCurrencyAxis) {
    addCurrencyYAxis(slide, plot, 0, maxTotal);
  } else {
    addAxis(slide, plot, 0, maxTotal, formatter);
  }

  rawLabels.forEach((rawLabel, labelIndex) => {
    let currentBottom = plot.y + plot.h;
    const x = plot.x + labelIndex * step + (step - barW) / 2;
    const total = totals[labelIndex] || 1;

    series.forEach((item, seriesIndex) => {
      const value = item.values[labelIndex] ?? 0;
      const share = value / total;
      const denominator = usePercentStack ? total : maxTotal;
      const h = (value / denominator) * plot.h;
      currentBottom -= h;
      if (h > 0.01) {
        const segmentColor = chartColors[seriesIndex % chartColors.length];
        slide.addShape("rect", {
          x,
          y: currentBottom,
          w: barW,
          h,
          fill: { color: segmentColor },
          line: {
            color: theme.white,
            width: usePercentStack ? 0.6 : 0,
            transparency: usePercentStack ? 0 : 100,
          },
        });
        if (usePercentStack && share >= 0.14 && rawLabels.length <= 6) {
          slide.addText(`${Math.round(share * 100)}%`, {
            x: x + 0.04,
            y: currentBottom + h / 2 - 0.08,
            w: barW - 0.08,
            h: 0.16,
            fontFace: font.body,
            fontSize: 7.0,
            bold: true,
            color: theme.white,
            align: "center",
            margin: 0,
          });
        }
      }
    });

    const parts = dateParts(rawLabel);
    const showMonth = parts.day === 1 || labelIndex === 0;
    const showDay =
      parts.day === 1 ||
      parts.day === 15 ||
      labelIndex === rawLabels.length - 1;
    if (showDay) {
      slide.addText(showMonth ? parts.month : String(parts.day), {
        x: x - 0.13,
        y: plot.y + plot.h + 0.09,
        w: 0.34,
        h: 0.14,
        fontFace: font.body,
        fontSize: showMonth ? 8 : 6.8,
        bold: showMonth,
        color: theme.muted,
        align: "center",
        margin: 0,
      });
      slide.addShape("line", {
        x: x + barW / 2,
        y: plot.y + plot.h,
        w: 0,
        h: 0.06,
        line: { color: "6B7280", width: 0.6 },
      });
    }
  });

  addShapeLegend(
    slide,
    series.map((item) => item.name),
    plot.x,
    box.y + 0.08,
  );
}

function addGroupedBarShapeChart(
  slide: PptxSlide,
  rows: QueryRow[],
  box: Box,
  formatter: (value: number) => string = formatCurrency,
): void {
  const seriesList = buildMultiSeries(rows);
  const rawLabels = Array.from(new Set(rows.map((row) => String(row.t)))).sort();
  const allValues = seriesList.flatMap((s) => s.values);
  const maxValue = niceMax(Math.max(...allValues, 1));
  const useCurrencyAxis = formatter === formatCurrency;

  const plot = useCurrencyAxis
    ? { x: box.x + 1.25, y: box.y + 0.55, w: box.w - 1.85, h: box.h - 1.35 }
    : { x: box.x + 0.5, y: box.y + 0.55, w: box.w - 0.95, h: box.h - 1.35 };

  const numSeries = seriesList.length;
  const numLabels = rawLabels.length;
  const groupStep = plot.w / Math.max(numLabels, 1);
  const groupPadding = Math.min(0.08, groupStep * 0.12);
  const intraGap = 0.04;
  const barW = Math.max(
    0.09,
    (groupStep - groupPadding * 2 - intraGap * (numSeries - 1)) /
      Math.max(numSeries, 1),
  );

  if (useCurrencyAxis) {
    addCurrencyYAxis(slide, plot, 0, maxValue);
  } else {
    addAxis(slide, plot, 0, maxValue, formatter);
  }

  slide.addShape("line", {
    x: plot.x,
    y: plot.y + plot.h,
    w: plot.w,
    h: 0,
    line: { color: theme.axis, width: 0.9 },
  });

  rawLabels.forEach((rawLabel, labelIndex) => {
    const groupX = plot.x + labelIndex * groupStep + groupPadding;

    seriesList.forEach((item, seriesIndex) => {
      const value = item.values[labelIndex] ?? 0;
      const h = Math.max((value / maxValue) * plot.h, 0.005);
      const x = groupX + seriesIndex * (barW + intraGap);
      const barColor = chartColors[seriesIndex % chartColors.length];

      slide.addShape("roundRect", {
        x,
        y: plot.y + plot.h - h,
        w: barW,
        h,
        rectRadius: 0.03,
        fill: { color: barColor },
        line: { color: barColor, transparency: 100 },
      });

      if (numLabels <= 6 && value > 0) {
        slide.addText(formatter(value), {
          x: x - 0.1,
          y: plot.y + plot.h - h - 0.18,
          w: barW + 0.2,
          h: 0.15,
          fontFace: font.body,
          fontSize: 5.8,
          bold: true,
          color: theme.text,
          align: "center",
          margin: 0,
        });
      }
    });

    slide.addText(toPeriodLabel(rawLabel), {
      x: plot.x + labelIndex * groupStep - 0.06,
      y: plot.y + plot.h + 0.09,
      w: groupStep + 0.12,
      h: 0.16,
      fontFace: font.body,
      fontSize: 6.5,
      color: theme.muted,
      align: "center",
      margin: 0,
    });
  });

  addShapeLegend(
    slide,
    seriesList.map((item) => item.name),
    plot.x,
    box.y + 0.08,
  );
}

function addBarShapeChart(
  slide: PptxSlide,
  rows: QueryRow[],
  box: Box,
  color: string,
  formatter: (value: number) => string,
  intent?: DataChartIntent,
): void {
  const values = rows.map((row) => Number(row.value ?? 0));
  const labels = rows.map((row) => toPeriodLabel(row.t));
  const maxValue = Math.max(...values, 1);
  const axisMax = niceMax(maxValue * 1.15);
  const plot = {
    x: box.x + 0.72,
    y: box.y + 0.36,
    w: box.w - 1.35,
    h: box.h - 1.02,
  };
  const step = plot.w / Math.max(values.length, 1);
  const barW = Math.max(
    0.12,
    Math.min(step * (values.length <= 8 ? 0.5 : 0.62), step - 0.16),
  );

  for (let i = 0; i <= 4; i += 1) {
    const y = plot.y + (plot.h * i) / 4;
    slide.addShape("line", {
      x: plot.x,
      y,
      w: plot.w,
      h: 0,
      line: { color: theme.faint, width: 0.55, transparency: i === 4 ? 100 : 0 },
    });
    if (i === 0 || i === 4) {
      const value = axisMax - (axisMax * i) / 4;
      slide.addText(formatter(value), {
        x: plot.x + plot.w + 0.08,
        y: y - 0.06,
        w: 0.55,
        h: 0.13,
        fontFace: font.body,
        fontSize: 5.8,
        color: theme.muted,
        margin: 0,
      });
    }
  }
  slide.addShape("line", {
    x: plot.x,
    y: plot.y + plot.h,
    w: plot.w,
    h: 0,
    line: { color: theme.axis, width: 0.9 },
  });

  values.forEach((value, index) => {
    const x = plot.x + index * step + (step - barW) / 2;
    const h = (value / axisMax) * plot.h;
    const isLast = index === values.length - 1;
    slide.addShape("roundRect", {
      x,
      y: plot.y + plot.h - h,
      w: barW,
      h,
      rectRadius: 0.035,
      fill: { color, transparency: isLast ? 0 : 18 },
      line: { color, transparency: 100 },
    });

    if (values.length <= 8 || index % 3 === 0 || isLast) {
      slide.addText(labels[index], {
        x: x - 0.24,
        y: plot.y + plot.h + 0.09,
        w: Math.max(0.6, barW + 0.48),
        h: 0.16,
        fontFace: font.body,
        fontSize: 6.3,
        color: theme.muted,
        align: "center",
        margin: 0,
      });
    }
    if (values.length <= 8) {
      slide.addText(formatter(value), {
        x: x - 0.16,
        y: plot.y + plot.h - h - 0.22,
        w: Math.max(0.5, barW + 0.24),
        h: 0.17,
        fontFace: font.body,
        fontSize: 7.0,
        bold: true,
        color: theme.text,
        align: "center",
        margin: 0,
      });
    }
  });
  addIntentBadge(slide, box, rows, intent, formatter);
}

function addHorizontalBarShapeChart(
  slide: PptxSlide,
  rows: QueryRow[],
  box: Box,
  color: string,
  formatter: (value: number) => string,
  intent?: DataChartIntent,
): void {
  const rankedRows = [...rows]
    .map((row) => ({ label: String(row.t), value: Number(row.value ?? 0) }))
    .sort((a, b) => b.value - a.value)
    .slice(0, 8);
  const maxValue = Math.max(...rankedRows.map((row) => row.value), 1);
  const plot = {
    x: box.x + 2.35,
    y: box.y + 0.42,
    w: box.w - 3.65,
    h: box.h - 0.98,
  };
  const rowGap = 0.1;
  const rowH = Math.min(
    0.36,
    (plot.h - rowGap * Math.max(rankedRows.length - 1, 0)) /
      Math.max(rankedRows.length, 1),
  );

  rankedRows.forEach((row, index) => {
    const y = plot.y + index * (rowH + rowGap);
    const barW = (row.value / maxValue) * plot.w;
    const highlightMix = intent === "mix_concentration";
    const highlightLeader = intent === "rank_leader";
    const rowColor =
      highlightLeader && index > 0
        ? "CBD5E1"
        : highlightMix && index > 2
          ? "CBD5E1"
          : index === 0
            ? color
            : chartColors[index % chartColors.length] ?? color;
    const rowMuted = rowColor === "CBD5E1";
    const label =
      row.label.length > 27 ? `${row.label.slice(0, 24).trim()}...` : row.label;

    slide.addText(label, {
      x: box.x + 0.15,
      y: y + 0.05,
      w: 2.0,
      h: 0.18,
      fontFace: font.body,
      fontSize: 7.0,
      color: theme.text,
      align: "right",
      margin: 0,
    });
    slide.addShape("rect", {
      x: plot.x,
      y,
      w: plot.w,
      h: rowH,
      fill: { color: theme.softGray },
      line: { color: theme.softGray, transparency: 100 },
    });
    slide.addShape("rect", {
      x: plot.x,
      y,
      w: Math.max(0.02, barW),
      h: rowH,
      fill: { color: rowColor },
      line: { color: rowColor, transparency: 100 },
    });
    const valueInside = barW > 0.95 && !rowMuted;
    slide.addText(formatter(row.value), {
      x: valueInside
        ? plot.x + barW - 0.82
        : Math.min(plot.x + plot.w - 0.78, plot.x + barW + 0.08),
      y: y + 0.05,
      w: 0.74,
      h: 0.18,
      fontFace: font.body,
      fontSize: 7.0,
      bold: true,
      color: valueInside ? theme.white : rowMuted ? theme.muted : theme.text,
      align: valueInside ? "right" : "left",
      margin: 0,
    });
  });
}

function addLineSegment(
  slide: PptxSlide,
  p1: { x: number; y: number },
  p2: { x: number; y: number },
  color: string,
): void {
  const x = Math.min(p1.x, p2.x);
  const y = Math.min(p1.y, p2.y);
  const w = Math.max(0.001, Math.abs(p2.x - p1.x));
  const h = Math.max(0.001, Math.abs(p2.y - p1.y));
  const rising = p2.y < p1.y;
  slide.addShape("line", {
    x,
    y,
    w,
    h,
    flipV: rising,
    line: { color, width: 2.2, beginArrowType: "none", endArrowType: "none" },
  });
}

function addLineShapeChart(
  slide: PptxSlide,
  rows: QueryRow[],
  box: Box,
  color: string,
  formatter: (value: number) => string,
  column = "value",
): void {
  const labels = rows.map((row) => toChartLabel(row.t));
  const values = rows.map((row) => Number(row[column] ?? 0));
  const min = Math.min(...values);
  const max = Math.max(...values);
  const range = Math.max(max - min, 1);
  const paddedMin = Math.max(0, min - range * 0.12);
  const paddedMax = max + range * 0.12;
  const paddedRange = Math.max(paddedMax - paddedMin, 1);
  const plot = {
    x: box.x + 0.15,
    y: box.y + 0.18,
    w: box.w - 0.8,
    h: box.h - 0.65,
  };

  addAxis(slide, plot, paddedMin, paddedMax, formatter);

  const points = values.map((value, index) => ({
    x: plot.x + (index / Math.max(values.length - 1, 1)) * plot.w,
    y: plot.y + plot.h - ((value - paddedMin) / paddedRange) * plot.h,
  }));

  for (let index = 0; index < points.length - 1; index += 1) {
    addLineSegment(slide, points[index], points[index + 1], color);
  }

  points.forEach((point, index) => {
    if (index === 0 || index === points.length - 1 || index % 5 === 0) {
      slide.addShape("ellipse", {
        x: point.x - 0.035,
        y: point.y - 0.035,
        w: 0.07,
        h: 0.07,
        fill: { color },
        line: { color },
      });
    }
    if (index === 0 || index === points.length - 1) {
      const labelX =
        index === 0
          ? Math.max(box.x, point.x - 0.18)
          : Math.min(box.x + box.w - 0.86, point.x + 0.09);
      const labelY =
        point.y < plot.y + 0.25 ? point.y + 0.08 : point.y - 0.2;
      slide.addText(formatter(values[index]), {
        x: labelX,
        y: labelY,
        w: 0.78,
        h: 0.16,
        fontFace: font.body,
        fontSize: 6.6,
        bold: true,
        color: theme.text,
        margin: 0,
      });
    }
    if (index % 3 === 0) {
      slide.addText(labels[index], {
        x: point.x - 0.2,
        y: plot.y + plot.h + 0.08,
        w: 0.4,
        h: 0.14,
        fontFace: font.body,
        fontSize: 5.8,
        color: theme.muted,
        align: "center",
        margin: 0,
      });
    }
  });
}

function formatPercent(value: number): string {
  return `${compactNumber(value)}%`;
}

function valueFormatter(
  format: "currency" | "number" | "percent",
): (value: number) => string {
  if (format === "currency") return formatCurrency;
  if (format === "percent") return formatPercent;
  return compactNumber;
}

function dataChartBodyText(
  slideSpec: Extract<DeckSpec["slides"][number], { type: "data_chart" }>,
): string | undefined {
  return slideSpec.takeaway?.trim() || slideSpec.body;
}

function dataChartSourceText(
  slideSpec: Extract<DeckSpec["slides"][number], { type: "data_chart" }>,
): string {
  const parts = [slideSpec.source?.trim() || "Inline data"];
  if (slideSpec.period?.trim()) parts.push(`Period: ${slideSpec.period.trim()}`);
  if (slideSpec.unit?.trim()) parts.push(`Unit: ${slideSpec.unit.trim()}`);
  return parts.join(" · ");
}

function hashText(value: string): string {
  return createHash("sha256").update(value).digest("hex");
}

function hashFile(filePath: string): string {
  return createHash("sha256").update(fs.readFileSync(filePath)).digest("hex");
}

function replaceParentDir(
  filePath: string,
  oldParent: string,
  newParent: string,
): string {
  return path.join(newParent, path.relative(oldParent, filePath));
}

function promoteTempPath(
  tempPath: string,
  finalPath: string,
  backupPath: string,
  isDirectory = false,
): void {
  if (!fs.existsSync(tempPath)) return;
  if (fs.existsSync(finalPath)) {
    fs.renameSync(finalPath, backupPath);
  }
  fs.mkdirSync(path.dirname(finalPath), { recursive: true });
  fs.renameSync(tempPath, finalPath);
  fs.rmSync(backupPath, { force: true, recursive: isDirectory });
}

function rollbackPromotedPath(
  tempPath: string,
  finalPath: string,
  backupPath: string,
  isDirectory = false,
): void {
  if (fs.existsSync(finalPath)) {
    fs.rmSync(finalPath, { force: true, recursive: isDirectory });
  }
  if (fs.existsSync(backupPath)) {
    fs.mkdirSync(path.dirname(finalPath), { recursive: true });
    fs.renameSync(backupPath, finalPath);
  }
  if (fs.existsSync(tempPath)) {
    fs.rmSync(tempPath, { force: true, recursive: isDirectory });
  }
}

function addImageOrPlaceholder(
  slide: PptxSlide,
  imagePath: string | undefined,
  box: Box,
  baseDir: string,
): void {
  const absoluteImagePath = imagePath ? path.resolve(baseDir, imagePath) : "";

  if (absoluteImagePath && fs.existsSync(absoluteImagePath)) {
    slide.addImage({ path: absoluteImagePath, ...box });
    return;
  }

  if (imagePath) {
    throw new Error(`Image asset not found: ${imagePath}`);
  }

  slide.addShape("rect", {
    ...box,
    fill: { color: theme.softGray },
    line: { color: theme.faint, width: 1 },
  });
  slide.addShape("rect", {
    x: box.x + 0.22,
    y: box.y + 0.22,
    w: 0.36,
    h: 0.28,
    fill: { color: theme.softBlue },
    line: { color: theme.softBlue, transparency: 100 },
  });
  slide.addText("Object placeholder", {
    x: box.x + 0.22,
    y: box.y + box.h - 0.5,
    w: box.w - 0.44,
    h: 0.22,
    fontFace: font.body,
    fontSize: 10,
    bold: true,
    color: theme.muted,
    align: "left",
    valign: "mid",
    margin: 0,
  });
}

function setBackground(slide: PptxSlide): void {
  slide.background = { color: theme.background };
}

async function main(): Promise<void> {
  await renderDeck();
}

export async function renderDeck(
  options: RenderDeckOptions = {},
): Promise<RenderReport> {
  const specPath = options.specPath
    ? path.resolve(rootDir, options.specPath)
    : path.join(rootDir, "deck_spec.json");
  const specDir = path.dirname(specPath);
  const artifactsDir = options.artifactsDir
    ? path.resolve(rootDir, options.artifactsDir)
    : path.join(rootDir, "output");
  const outputPath = options.outputPath
    ? path.resolve(rootDir, options.outputPath)
    : path.join(artifactsDir, "test_deck.pptx");
  const spec = readDeckSpec(specPath);
  const runName = path.basename(outputPath, ".pptx");
  const runArtifactsDir = path.join(artifactsDir, "runs", runName);
  const manifestPath = path.join(
    artifactsDir,
    `${path.basename(outputPath, ".pptx")}.manifest.json`,
  );
  const runToken = `${Date.now()}`;
  const tempOutputPath = path.join(
    path.dirname(outputPath),
    `${runName}.tmp-${runToken}.pptx`,
  );
  const tempManifestPath = `${manifestPath}.tmp-${runToken}`;
  const tempRunArtifactsDir = `${runArtifactsDir}.tmp-${runToken}`;
  const backupOutputPath = `${outputPath}.bak-${runToken}`;
  const backupManifestPath = `${manifestPath}.bak-${runToken}`;
  const backupRunArtifactsDir = `${runArtifactsDir}.bak-${runToken}`;
  const specHash = hashText(fs.readFileSync(specPath, "utf8"));
  const manifestSlides: RenderManifest["slides"] = [];

  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_WIDE";
  pptx.author = "pptx-agent-tool";
  pptx.subject = "Generated presentation";
  pptx.title = spec.title;
  pptx.company = "pptx-agent-tool";
  pptx.lang = "en-US";
  pptx.theme = {
    headFontFace: font.head,
    bodyFontFace: font.body,
    lang: "en-US",
  };
  pptx.defineLayout({
    name: "CUSTOM_WIDE",
    width: layouts.page.width,
    height: layouts.page.height,
  });
  pptx.layout = "CUSTOM_WIDE";

  const slidesToRender = spec.slides;

  for (const slideSpec of slidesToRender) {
    const slide = pptx.addSlide();
    setBackground(slide);

    if (slideSpec.type === "title") {
      manifestSlides.push({
        index: manifestSlides.length + 1,
        type: slideSpec.type,
        title: slideSpec.title,      });
      addCoverBackground(slide);
      slide.addText(slideSpec.title, {
        x: layouts.archetypes.titleSlide.title.x,
        y: layouts.archetypes.titleSlide.title.y,
        w: 4.95,
        h: 2.85,
        fontFace: font.head,
        fontSize: coverTitleFontSize(slideSpec.title),
        bold: true,
        color: theme.title,
        align: "left",
        fit: "shrink",
        margin: 0,
      });
      slide.addText(slideSpec.subtitle ?? "", {
        ...layouts.archetypes.titleSlide.body,
        w: 5.0,
        fontFace: font.body,
        fontSize: 18,
        color: theme.title,
        align: "left",
        margin: 0,
      });
    }

    if (slideSpec.type === "content") {
      manifestSlides.push({
        index: manifestSlides.length + 1,
        type: slideSpec.type,
        title: slideSpec.title,      });
      addDeckTitle(slide, slideSpec.title, layouts.derived.content.title);
      if (slideSpec.image) {
        addBody(slide, slideSpec.body, layouts.derived.content.body);
      } else if (slideSpec.body) {
        slide.addShape("rect", {
          x: 1.391,
          y: 2.08,
          w: 0.08,
          h: 2.35,
          fill: { color: theme.mainBlue },
          line: { color: theme.mainBlue, transparency: 100 },
        });
        addBody(slide, slideSpec.body, {
          x: 1.68,
          y: 2.08,
          w: 9.3,
          h: 2.45,
        });
      }
      if (slideSpec.image) {
        addImageOrPlaceholder(
          slide,
          slideSpec.image,
          layouts.derived.content.image,
          specDir,
        );
      }
    }

    if (slideSpec.type === "section") {
      manifestSlides.push({
        index: manifestSlides.length + 1,
        type: slideSpec.type,
        title: slideSpec.title,      });
      slide.background = { color: theme.dividerBlue };
      slide.addText(slideSpec.body ?? "", {
        ...layouts.archetypes.divider.body,
        fontFace: font.body,
        fontSize: 13,
        color: theme.white,
        margin: 0,
      });
      slide.addText(slideSpec.title, {
        ...layouts.archetypes.divider.title,
        fontFace: font.head,
        fontSize: 34,
        bold: true,
        color: theme.white,
        fit: "shrink",
        margin: 0,
      });
    }

    if (slideSpec.type === "comparison") {
      manifestSlides.push({
        index: manifestSlides.length + 1,
        type: slideSpec.type,
        title: slideSpec.title,      });
      addDeckTitle(slide, slideSpec.title, layouts.derived.comparison.title);
      addBody(slide, slideSpec.body, layouts.derived.comparison.body);
      addImageOrPlaceholder(
        slide,
        slideSpec.images?.[0],
        layouts.derived.comparison.leftImage,
        specDir,
      );
      addImageOrPlaceholder(
        slide,
        slideSpec.images?.[1],
        layouts.derived.comparison.rightImage,
        specDir,
      );
    }

    if (slideSpec.type === "insight") {
      manifestSlides.push({
        index: manifestSlides.length + 1,
        type: slideSpec.type,
        title: slideSpec.title,      });
      addDeckTitle(slide, slideSpec.title, layouts.archetypes.text.title);
      slide.addShape("rect", {
        x: 1.391,
        y: 2.0,
        w: 0.1,
        h: 2.55,
        fill: { color: theme.mainBlue },
        line: { color: theme.mainBlue, transparency: 100 },
      });
      slide.addText(slideSpec.callout, {
        x: 1.72,
        y: 1.88,
        w: 6.3,
        h: 2.1,
        fontFace: font.head,
        fontSize: slideSpec.callout.length > 90 ? 24 : 29,
        bold: true,
        color: theme.title,
        fit: "shrink",
        valign: "mid",
        margin: 0,
      });
      if (slideSpec.body) {
        addPanel(slide, { x: 8.35, y: 2.0, w: 3.45, h: 2.55 }, {
          fill: theme.faintBlue,
          line: theme.softBlue,
        });
        slide.addText(slideSpec.body, {
          x: 8.65,
          y: 2.35,
          w: 2.85,
          h: 1.8,
          fontFace: font.body,
          fontSize: 12.5,
          color: theme.text,
          fit: "shrink",
          valign: "mid",
          margin: 0,
        });
      }
    }

    if (slideSpec.type === "process") {
      manifestSlides.push({
        index: manifestSlides.length + 1,
        type: slideSpec.type,
        title: slideSpec.title,      });
      addDeckTitle(slide, slideSpec.title, layouts.archetypes.text.title);
      addBody(slide, slideSpec.body, { x: 1.391, y: 1.55, w: 10.637, h: 0.3 });

      const steps = slideSpec.steps;
      const gutter = 0.22;
      const baseX = 1.391;
      const baseY = 2.18;
      const totalW = 10.55;
      const cardW = (totalW - gutter * (steps.length - 1)) / steps.length;
      const cardH = 2.55;

      steps.forEach((step, index) => {
        const x = baseX + index * (cardW + gutter);
        const accent = chartColors[index % chartColors.length];
        addPanel(slide, { x, y: baseY, w: cardW, h: cardH }, {
          fill: index === 0 ? theme.faintBlue : theme.white,
          line: theme.faint,
        });
        slide.addShape("ellipse", {
          x: x + 0.22,
          y: baseY + 0.25,
          w: 0.38,
          h: 0.38,
          fill: { color: accent },
          line: { color: accent, transparency: 100 },
        });
        slide.addText(step.label ?? String(index + 1), {
          x: x + 0.22,
          y: baseY + 0.32,
          w: 0.38,
          h: 0.14,
          fontFace: font.head,
          fontSize: 7,
          bold: true,
          color: theme.white,
          align: "center",
          margin: 0,
        });
        slide.addText(step.title, {
          x: x + 0.22,
          y: baseY + 0.86,
          w: cardW - 0.44,
          h: 0.42,
          fontFace: font.head,
          fontSize: 12.5,
          bold: true,
          color: theme.title,
          fit: "shrink",
          margin: 0,
        });
        if (step.body) {
          slide.addText(step.body, {
            x: x + 0.22,
            y: baseY + 1.42,
            w: cardW - 0.44,
            h: 0.7,
            fontFace: font.body,
            fontSize: 8.2,
            color: theme.muted,
            fit: "shrink",
            margin: 0,
          });
        }
        if (index < steps.length - 1) {
          slide.addShape("line", {
            x: x + cardW + 0.04,
            y: baseY + cardH / 2,
            w: gutter - 0.08,
            h: 0,
            line: { color: theme.axis, width: 1.1, endArrowType: "triangle" },
          });
        }
      });
    }

    if (slideSpec.type === "data_chart") {
      const rows = slideSpec.data as QueryRow[];
      const dataRows = rows as Array<{
        t: unknown;
        series?: unknown;
        value?: unknown;
      }>;
      const formatter = valueFormatter(slideSpec.valueFormat);
      const multi = hasSeries(rows);
      const chartPattern =
        slideSpec.pattern ?? inferDataChartPattern(dataRows, slideSpec.chart);
      const chartIntent = slideSpec.intent;
      const effectiveChart = effectiveChartModeForPattern(
        chartPattern,
        slideSpec.chart,
      );
      const categoricalSingleSeries =
        !multi && !rows.every((row) => isDateLike(row.t));
      const bodyText = dataChartBodyText(slideSpec);
      const shift = headerShift(slideSpec.title, bodyText);

      addDeckTitle(slide, slideSpec.title, layouts.derived.chartSlide.title);
      addBody(slide, bodyText, shiftedBox(layouts.derived.chartSlide.body, shift));

      if (categoricalSingleSeries) {
        const rankedRows = [...rows]
          .map((row) => ({ label: String(row.t), value: Number(row.value ?? 0) }))
          .sort((a, b) => b.value - a.value);
        const topRow = rankedRows[0];
        const total = rankedRows.reduce((sum, row) => sum + row.value, 0);
        const share = total ? `${((topRow.value / total) * 100).toFixed(0)}%` : "n/a";
        addKpi(
          slide,
          "Top stream",
          topRow.label.replace(/\s+Revenue$/i, ""),
          layouts.derived.chartSlide.kpi1,
          theme.mainBlue,
        );
        addKpi(
          slide,
          "Share",
          share,
          layouts.derived.chartSlide.kpi2,
          theme.text,
        );
      } else if (multi) {
        const total = totalValue(rows);
        addKpi(
          slide,
          "Total",
          formatter(total),
          layouts.derived.chartSlide.kpi1,
          theme.mainBlue,
        );
      } else {
        const latest = latestValue(rows);
        const first = Number(rows[0]?.value ?? 0);
        addKpi(
          slide,
          "Latest",
          formatter(latest),
          layouts.derived.chartSlide.kpi1,
          theme.mainBlue,
        );
        addKpi(
          slide,
          intentKpiLabel(chartIntent),
          formatDelta(first, latest),
          layouts.derived.chartSlide.kpi2,
          theme.text,
        );
      }

      if (multi && chartPattern === "grouped_bar") {
        addGroupedBarShapeChart(
          slide,
          rows,
          shiftedObjectBox(layouts.derived.chartSlide.chart, shift),
          formatter,
        );
      } else if (multi) {
        addStackedBarShapeChart(
          slide,
          rows,
          shiftedObjectBox(layouts.derived.chartSlide.chart, shift),
          formatter,
          chartIntent === "composition_shift" || chartIntent === "mix_concentration"
            ? "percent"
            : "absolute",
        );
      } else if (
        effectiveChart === "horizontal_bar" ||
        categoricalSingleSeries
      ) {
        addHorizontalBarShapeChart(
          slide,
          rows,
          shiftedObjectBox(layouts.derived.chartSlide.chart, shift),
          theme.mainBlue,
          formatter,
          chartIntent,
        );
      } else if (effectiveChart === "line") {
        const chartBox = shiftedObjectBox(layouts.derived.chartSlide.chart, shift);
        addLineShapeChart(
          slide,
          rows,
          chartBox,
          theme.mainBlue,
          formatter,
        );
        addIntentBadge(slide, chartBox, rows, chartIntent, formatter);
      } else {
        addBarShapeChart(
          slide,
          rows,
          shiftedObjectBox(layouts.derived.chartSlide.chart, shift),
          theme.mainBlue,
          formatter,
          chartIntent,
        );
      }

      const sourceText = dataChartSourceText(slideSpec);
      addSource(slide, sourceText, layouts.derived.chartSlide.source);
      manifestSlides.push({
        index: manifestSlides.length + 1,
        type: slideSpec.type,
        title: slideSpec.title,        chartMode: effectiveChart,
        chartPattern,
        chartIntent,
        sourceText,
      });
    }

    if (slideSpec.type === "kpi") {
      addDeckTitle(slide, slideSpec.title, layouts.archetypes.text.title);
      addBody(slide, slideSpec.body, { x: 1.391, y: 1.62, w: 10.637, h: 0.3 });

      const items = slideSpec.items;
      const count = items.length;
      const totalW = 10.55;
      const gutter = 0.22;
      const itemW = (totalW - gutter * (count - 1)) / count;
      const baseX = 1.391;
      const baseY = 2.3;
      const cardH = 2.35;

      items.forEach((item, i) => {
        const x = baseX + i * (itemW + gutter);
        addPanel(slide, { x, y: baseY, w: itemW, h: cardH }, {
          fill: i === 0 ? theme.faintBlue : theme.white,
          line: i === 0 ? theme.softBlue : theme.faint,
        });
        addTopRule(slide, x + 0.18, baseY + 0.18, itemW - 0.36);
        slide.addText(item.value, {
          x: x + 0.25,
          y: baseY + 0.52,
          w: itemW - 0.5,
          h: 0.78,
          fontFace: font.head,
          fontSize: 31,
          bold: true,
          color: chartColors[i % chartColors.length],
          align: "left",
          valign: "mid",
          fit: "shrink",
          margin: 0,
        });
        slide.addText(item.label.toUpperCase(), {
          x: x + 0.25,
          y: baseY + 1.42,
          w: itemW - 0.5,
          h: 0.25,
          fontFace: font.body,
          fontSize: 8,
          bold: true,
          color: theme.muted,
          align: "left",
          margin: 0,
        });
        if (item.subtitle) {
          slide.addText(item.subtitle, {
            x: x + 0.25,
            y: baseY + 1.77,
            w: itemW - 0.5,
            h: 0.25,
            fontFace: font.body,
            fontSize: 7,
            color: theme.muted,
            align: "left",
            margin: 0,
          });
        }
      });

      manifestSlides.push({
        index: manifestSlides.length + 1,
        type: slideSpec.type,
        title: slideSpec.title,      });
    }

    if (slideSpec.type === "table") {
      addDeckTitle(slide, slideSpec.title, layouts.archetypes.text.title);
      addBody(slide, slideSpec.body, { x: 1.391, y: 1.62, w: 10.637, h: 0.3 });

      const cols = slideSpec.columns;
      const rows = slideSpec.rows;
      const tableX = 1.391;
      const tableY = 2.05;
      const tableW = 10.5;
      const colWeights = cols.map((col, index) => {
        const name = col.toLowerCase();
        if (index === 0) return 1.15;
        if (index === cols.length - 1) return 1.55;
        if (name.includes("share") || name.includes("qoq")) return 0.82;
        return 1.0;
      });
      const totalWeight = colWeights.reduce((sum, weight) => sum + weight, 0);
      const colWidths = colWeights.map((weight) => (tableW * weight) / totalWeight);
      const colX = (index: number): number =>
        tableX + colWidths.slice(0, index).reduce((sum, w) => sum + w, 0);
      const rowH = rows.length > 10 ? 0.31 : 0.36;
      const maxRows = Math.min(rows.length, 13);
      addPanel(slide, {
        x: tableX - 0.08,
        y: tableY - 0.08,
        w: tableW + 0.16,
        h: rowH * (maxRows + 1) + 0.16,
      });

      cols.forEach((col, i) => {
        const x = colX(i);
        const w = colWidths[i];
        slide.addShape("rect", {
          x,
          y: tableY,
          w,
          h: rowH,
          fill: { color: theme.mainBlue },
          line: { color: theme.mainBlue, transparency: 100 },
        });
        slide.addText(col, {
          x: x + 0.07,
          y: tableY,
          w: w - 0.14,
          h: rowH,
          fontFace: font.body,
          fontSize: 7.5,
          bold: true,
          color: theme.white,
          align: i === 0 || i === cols.length - 1 ? "left" : "right",
          valign: "mid",
          fit: "shrink",
          margin: 0,
        });
      });

      for (let r = 0; r < maxRows; r++) {
        const row = rows[r];
        const y = tableY + rowH + r * rowH;
        const isEven = r % 2 === 0;

        cols.forEach((_, i) => {
          const x = colX(i);
          const w = colWidths[i];
          slide.addShape("rect", {
            x,
            y,
            w,
            h: rowH,
            fill: { color: isEven ? theme.softGray : theme.white },
            line: { color: theme.faint, transparency: 20 },
          });
          slide.addText(String(row[i] ?? ""), {
            x: x + 0.07,
            y,
            w: w - 0.14,
            h: rowH,
            fontFace: font.body,
            fontSize: 7.1,
            bold: i === 0,
            color: i === 0 ? theme.title : theme.text,
            align: i === 0 || i === cols.length - 1 ? "left" : "right",
            valign: "mid",
            margin: 0.06,
          });
        });
      }

      if (rows.length > maxRows) {
        slide.addText(`… and ${rows.length - maxRows} more rows`, {
          x: tableX,
          y: tableY + rowH + maxRows * rowH + 0.05,
          w: tableW,
          h: 0.2,
          fontFace: font.body,
          fontSize: 6.5,
          color: theme.muted,
          margin: 0,
        });
      }

      if (slideSpec.source) {
        addSource(slide, slideSpec.source, {
          x: 1.351,
          y: 7.05,
          w: 9.2,
          h: 0.14,
        });
      }

      manifestSlides.push({
        index: manifestSlides.length + 1,
        type: slideSpec.type,
        title: slideSpec.title,      });
    }
  }

  fs.mkdirSync(path.dirname(tempOutputPath), { recursive: true });
  await pptx.writeFile({ fileName: tempOutputPath });
  const metrics: string[] = [];
  const tempManifest: RenderManifest = {
    specPath,
    pptxPath: outputPath,
    specHash,
    pptxHash: hashFile(tempOutputPath),
    artifacts: {
      sql: metrics.map((metric) => {
        const artifactPath = path.join(
          tempRunArtifactsDir,
          "sql",
          `${metric}.sql`,
        );
        return {
          metric,
          path: replaceParentDir(
            artifactPath,
            tempRunArtifactsDir,
            runArtifactsDir,
          ),
          hash: hashFile(artifactPath),
        };
      }),
      data: metrics.map((metric) => {
        const artifactPath = path.join(
          tempRunArtifactsDir,
          "data",
          `${metric}.json`,
        );
        return {
          metric,
          path: replaceParentDir(
            artifactPath,
            tempRunArtifactsDir,
            runArtifactsDir,
          ),
          hash: hashFile(artifactPath),
        };
      }),
    },
    slides: manifestSlides,
  };
  fs.writeFileSync(
    tempManifestPath,
    `${JSON.stringify(tempManifest, null, 2)}\n`,
  );

  try {
    promoteTempPath(tempOutputPath, outputPath, backupOutputPath);
    promoteTempPath(tempManifestPath, manifestPath, backupManifestPath);
    promoteTempPath(
      tempRunArtifactsDir,
      runArtifactsDir,
      backupRunArtifactsDir,
      true,
    );
  } catch (error) {
    rollbackPromotedPath(tempOutputPath, outputPath, backupOutputPath);
    rollbackPromotedPath(tempManifestPath, manifestPath, backupManifestPath);
    rollbackPromotedPath(
      tempRunArtifactsDir,
      runArtifactsDir,
      backupRunArtifactsDir,
      true,
    );
    throw error;
  }

  const manifest: RenderManifest = {
    specPath,
    pptxPath: outputPath,
    specHash,
    pptxHash: tempManifest.pptxHash,
    artifacts: tempManifest.artifacts,
    slides: manifestSlides,
  };
  if (!options.quiet) {
    console.log(`Generated ${outputPath}`);
  }

  return {
    status: "ok",
    specPath,
    pptxPath: outputPath,
    artifactsDir,
    slideCount: slidesToRender.length,
    slides: slidesToRender.map((slide, index) => ({
      index: index + 1,
      type: slide.type,
      title: slide.title,
    })),
    manifestPath,
    artifacts: {
      sql: metrics.map((metric) =>
        path.join(runArtifactsDir, "sql", `${metric}.sql`),
      ),
      data: metrics.map((metric) =>
        path.join(runArtifactsDir, "data", `${metric}.json`),
      ),
    },
  };
}

const isMainModule =
  process.argv[1] &&
  path.resolve(process.argv[1]) === fileURLToPath(import.meta.url);

if (isMainModule) {
  main().catch((error) => {
    console.error(error);
    process.exit(1);
  });
}
