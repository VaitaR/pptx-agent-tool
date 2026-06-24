import fs from "node:fs";
import path from "node:path";
import {
  DATA_CHART_INTENTS,
  DATA_CHART_PATTERNS,
  dataChartQualityWarnings,
  type DataChartIntent,
  type DataChartPattern,
  isDataChartIntent,
  isDataChartPattern,
} from "./chart-patterns.js";

export type SlideSpec =
  | {
      type: "title";
      title: string;
      subtitle?: string;
    }
  | {
      type: "content";
      title: string;
      body?: string;
      image?: string;
    }
  | {
      type: "section";
      title: string;
      body?: string;
    }
  | {
      type: "comparison";
      title: string;
      body?: string;
      images?: string[];
    }
  | {
      type: "insight";
      title: string;
      callout: string;
      body?: string;
    }
  | {
      type: "process";
      title: string;
      body?: string;
      steps: Array<{ label?: string; title: string; body?: string }>;
    }
  | {
      type: "data_chart";
      title: string;
      body?: string;
      chart: "bar" | "line" | "horizontal_bar";
      pattern?: DataChartPattern;
      intent?: DataChartIntent;
      valueFormat: "currency" | "number" | "percent";
      takeaway?: string;
      unit?: string;
      period?: string;
      altText?: string;
      data: Array<{ t: string; series?: string; value: number }>;
      source?: string;
    }
  | {
      type: "kpi";
      title: string;
      body?: string;
      items: Array<{ label: string; value: string; subtitle?: string }>;
    }
  | {
      type: "table";
      title: string;
      body?: string;
      columns: string[];
      rows: Array<Array<string | number>>;
      source?: string;
    };

export type DeckSpec = {
  title: string;
  timeRange?: {
    start: string;
    end: string;
    grain: string;
  };
  slides: SlideSpec[];
};

export type DeckSpecValidationReport = {
  valid: boolean;
  errors: string[];
  warnings: string[];
  spec?: DeckSpec;
};

type ValidateDeckSpecOptions = {
  baseDir?: string;
};

const VALID_SLIDE_TYPES = new Set([
  "title",
  "content",
  "section",
  "comparison",
  "insight",
  "process",
  "data_chart",
  "kpi",
  "table",
]);
const VALID_DATA_CHART_TYPES = new Set(["bar", "line", "horizontal_bar"]);
const VALID_GRAINS = new Set(["day", "week", "month"]);

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function isNonEmptyString(value: unknown): value is string {
  return typeof value === "string" && value.trim().length > 0;
}

function parseIsoDate(value: string): number | null {
  const timestamp = Date.parse(value);
  return Number.isNaN(timestamp) ? null : timestamp;
}

function validateImagePath(
  imagePath: string,
  slidePath: string,
  options: ValidateDeckSpecOptions,
  errors: string[],
): void {
  const baseDir = options.baseDir ?? process.cwd();
  const resolvedPath = path.resolve(baseDir, imagePath);
  if (!fs.existsSync(resolvedPath)) {
    errors.push(`${slidePath} points to a missing image path: ${imagePath}`);
  }
}

function validateSlide(
  slide: unknown,
  index: number,
  errors: string[],
): slide is SlideSpec {
  if (!isRecord(slide)) {
    errors.push(`slides[${index}] must be an object`);
    return false;
  }

  if (!isNonEmptyString(slide.type) || !VALID_SLIDE_TYPES.has(slide.type)) {
    errors.push(
      `slides[${index}].type must be one of: ${Array.from(VALID_SLIDE_TYPES).join(", ")}`,
    );
    return false;
  }

  if (!isNonEmptyString(slide.title)) {
    errors.push(`slides[${index}].title must be a non-empty string`);
  }

  if (slide.body !== undefined && typeof slide.body !== "string") {
    errors.push(`slides[${index}].body must be a string when provided`);
  }

  if (
    slide.type === "title" &&
    slide.subtitle !== undefined &&
    typeof slide.subtitle !== "string"
  ) {
    errors.push(`slides[${index}].subtitle must be a string when provided`);
  }

  if (
    slide.type === "content" &&
    slide.image !== undefined &&
    typeof slide.image !== "string"
  ) {
    errors.push(`slides[${index}].image must be a string when provided`);
  }

  if (slide.type === "comparison") {
    if (slide.images !== undefined) {
      if (
        !Array.isArray(slide.images) ||
        !slide.images.every((image) => typeof image === "string")
      ) {
        errors.push(
          `slides[${index}].images must be an array of strings when provided`,
        );
      }
      if (Array.isArray(slide.images) && slide.images.length > 2) {
        errors.push(`slides[${index}].images supports at most 2 items`);
      }
    }
  }

  if (slide.type === "insight") {
    if (!isNonEmptyString(slide.callout)) {
      errors.push(`slides[${index}].callout must be a non-empty string`);
    }
  }

  if (slide.type === "process") {
    if (!Array.isArray(slide.steps) || slide.steps.length < 2) {
      errors.push(`slides[${index}].steps must contain at least 2 steps`);
    } else if (slide.steps.length > 5) {
      errors.push(`slides[${index}].steps supports at most 5 steps`);
    } else {
      for (let i = 0; i < slide.steps.length; i++) {
        const step = slide.steps[i];
        if (!isRecord(step) || !isNonEmptyString(step.title)) {
          errors.push(
            `slides[${index}].steps[${i}] must have a non-empty title`,
          );
          break;
        }
        if (step.label !== undefined && typeof step.label !== "string") {
          errors.push(`slides[${index}].steps[${i}].label must be a string`);
        }
        if (step.body !== undefined && typeof step.body !== "string") {
          errors.push(`slides[${index}].steps[${i}].body must be a string`);
        }
      }
    }
  }

  if (slide.type === "data_chart") {
    if (
      !isNonEmptyString(slide.chart) ||
      !VALID_DATA_CHART_TYPES.has(slide.chart)
    ) {
      errors.push(
        `slides[${index}].chart must be one of: ${Array.from(VALID_DATA_CHART_TYPES).join(", ")}`,
      );
    }
    if (
      slide.pattern !== undefined &&
      (!isDataChartPattern(slide.pattern) || typeof slide.pattern !== "string")
    ) {
      errors.push(
        `slides[${index}].pattern must be one of: ${DATA_CHART_PATTERNS.join(", ")}`,
      );
    }
    if (
      slide.intent !== undefined &&
      (!isDataChartIntent(slide.intent) || typeof slide.intent !== "string")
    ) {
      errors.push(
        `slides[${index}].intent must be one of: ${DATA_CHART_INTENTS.join(", ")}`,
      );
    }
    const validValueFormats = new Set(["currency", "number", "percent"]);
    if (
      !isNonEmptyString(slide.valueFormat) ||
      !validValueFormats.has(slide.valueFormat)
    ) {
      errors.push(
        `slides[${index}].valueFormat must be one of: currency, number, percent`,
      );
    }
    if (!Array.isArray(slide.data) || slide.data.length === 0) {
      errors.push(
        `slides[${index}].data must be a non-empty array of {t, value} objects`,
      );
    } else {
      for (let i = 0; i < slide.data.length; i++) {
        const row = slide.data[i];
        if (
          !isRecord(row) ||
          !isNonEmptyString(row.t) ||
          typeof row.value !== "number"
        ) {
          errors.push(
            `slides[${index}].data[${i}] must have string "t" and numeric "value"`,
          );
          break;
        }
      }
    }
    if (slide.source !== undefined && typeof slide.source !== "string") {
      errors.push(`slides[${index}].source must be a string when provided`);
    }
    if (slide.takeaway !== undefined && typeof slide.takeaway !== "string") {
      errors.push(`slides[${index}].takeaway must be a string when provided`);
    }
    if (slide.unit !== undefined && typeof slide.unit !== "string") {
      errors.push(`slides[${index}].unit must be a string when provided`);
    }
    if (slide.period !== undefined && typeof slide.period !== "string") {
      errors.push(`slides[${index}].period must be a string when provided`);
    }
    if (slide.altText !== undefined && typeof slide.altText !== "string") {
      errors.push(`slides[${index}].altText must be a string when provided`);
    }
  }

  if (slide.type === "kpi") {
    if (!Array.isArray(slide.items) || slide.items.length === 0) {
      errors.push(`slides[${index}].items must be a non-empty array`);
    } else if (slide.items.length > 4) {
      errors.push(`slides[${index}].items supports at most 4 KPIs`);
    } else {
      for (let i = 0; i < slide.items.length; i++) {
        const item = slide.items[i];
        if (
          !isRecord(item) ||
          !isNonEmptyString(item.label) ||
          !isNonEmptyString(item.value)
        ) {
          errors.push(
            `slides[${index}].items[${i}] must have string "label" and string "value"`,
          );
          break;
        }
      }
    }
  }

  if (slide.type === "table") {
    if (!Array.isArray(slide.columns) || slide.columns.length === 0) {
      errors.push(
        `slides[${index}].columns must be a non-empty array of strings`,
      );
    } else if (!slide.columns.every((col) => typeof col === "string")) {
      errors.push(`slides[${index}].columns must all be strings`);
    }
    if (!Array.isArray(slide.rows) || slide.rows.length === 0) {
      errors.push(`slides[${index}].rows must be a non-empty array`);
    } else {
      const colCount = Array.isArray(slide.columns) ? slide.columns.length : 0;
      for (let i = 0; i < slide.rows.length; i++) {
        const row = slide.rows[i];
        if (!Array.isArray(row) || row.length !== colCount) {
          errors.push(
            `slides[${index}].rows[${i}] must have ${colCount} values (matching columns)`,
          );
          break;
        }
      }
    }
    if (slide.source !== undefined && typeof slide.source !== "string") {
      errors.push(`slides[${index}].source must be a string when provided`);
    }
  }

  return true;
}

export function validateDeckSpec(
  input: unknown,
  options: ValidateDeckSpecOptions = {},
): DeckSpecValidationReport {
  const errors: string[] = [];
  const warnings: string[] = [];
  let startTimestamp: number | null = null;
  let endTimestamp: number | null = null;

  if (!isRecord(input)) {
    return {
      valid: false,
      errors: ["deck spec must be a JSON object"],
      warnings,
    };
  }

  if (!isNonEmptyString(input.title)) {
    errors.push("title must be a non-empty string");
  }

  if (input.timeRange !== undefined) {
    if (!isRecord(input.timeRange)) {
      errors.push("timeRange must be an object when provided");
    } else {
      if (!isNonEmptyString(input.timeRange.start)) {
        errors.push("timeRange.start must be a non-empty string");
      } else {
        startTimestamp = parseIsoDate(input.timeRange.start);
        if (startTimestamp === null) {
          errors.push("timeRange.start must be a valid ISO-like date string");
        }
      }
      if (!isNonEmptyString(input.timeRange.end)) {
        errors.push("timeRange.end must be a non-empty string");
      } else {
        endTimestamp = parseIsoDate(input.timeRange.end);
        if (endTimestamp === null) {
          errors.push("timeRange.end must be a valid ISO-like date string");
        }
      }
      if (!isNonEmptyString(input.timeRange.grain)) {
        errors.push("timeRange.grain must be a non-empty string");
      } else if (!VALID_GRAINS.has(input.timeRange.grain)) {
        errors.push(
          `timeRange.grain must be one of: ${Array.from(VALID_GRAINS).join(", ")}`,
        );
      }
      if (
        startTimestamp !== null &&
        endTimestamp !== null &&
        startTimestamp > endTimestamp
      ) {
        errors.push(
          "timeRange.start must be earlier than or equal to timeRange.end",
        );
      }
    }
  }

  if (!Array.isArray(input.slides)) {
    errors.push("slides must be an array");
  } else if (input.slides.length === 0) {
    errors.push("slides must contain at least one slide");
  } else {
    input.slides.forEach((slide, index) => {
      validateSlide(slide, index, errors);
    });
  }

  if (errors.length > 0) {
    return { valid: false, errors, warnings };
  }

  const spec = input as DeckSpec;

  spec.slides.forEach((slide, index) => {
    if (slide.type === "content" && slide.image) {
      validateImagePath(slide.image, `slides[${index}].image`, options, errors);
    }
    if (slide.type === "comparison" && slide.images) {
      slide.images.forEach((image, imageIndex) => {
        validateImagePath(
          image,
          `slides[${index}].images[${imageIndex}]`,
          options,
          errors,
        );
      });
    }
  });

  if (errors.length > 0) {
    return { valid: false, errors, warnings };
  }

  spec.slides.forEach((slide, index) => {
    if (slide.type !== "data_chart") return;
    const pattern = slide.pattern ?? undefined;
    for (const warning of dataChartQualityWarnings({
      rows: slide.data,
      chart: slide.chart,
      pattern,
      intent: slide.intent,
      valueFormat: slide.valueFormat,
      title: slide.title,
      body: slide.body,
      takeaway: slide.takeaway,
      unit: slide.unit,
      period: slide.period,
      source: slide.source,
      altText: slide.altText,
    })) {
      warnings.push(`slides[${index}]: ${warning}`);
    }
  });

  return {
    valid: true,
    errors,
    warnings,
    spec,
  };
}

export function readDeckSpec(specPath: string): DeckSpec {
  const raw = JSON.parse(fs.readFileSync(specPath, "utf8")) as unknown;
  const report = validateDeckSpec(raw, { baseDir: path.dirname(specPath) });

  if (!report.valid || !report.spec) {
    throw new Error(
      `Invalid deck spec at ${path.relative(process.cwd(), specPath)}:\n- ${report.errors.join("\n- ")}`,
    );
  }

  return report.spec;
}

export function defaultSpecPath(rootDir = process.cwd()): string {
  return path.join(rootDir, "deck_spec.json");
}
