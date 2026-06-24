import { createHash } from "node:crypto";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { unzipSync } from "fflate";
import { defaultSpecPath, readDeckSpec } from "./deck-spec.js";
import {
  dataChartQualityWarnings,
  dataChartIntentWarnings,
  dataChartPatternWarnings,
  effectiveChartModeForPattern,
  inferDataChartPattern,
  isDateLike,
} from "./chart-patterns.js";
import type { RenderManifest } from "./render-pptx.js";

const rootDir = process.cwd();

export type VerifyDeckOptions = {
  specPath?: string;
  pptxPath?: string;
  artifactsDir?: string;
  verbose?: boolean;
};

export class VerifyExecutionError extends Error {
  kind: "environment" | "verify";

  constructor(kind: "environment" | "verify", message: string) {
    super(message);
    this.name = "VerifyExecutionError";
    this.kind = kind;
  }
}

export type VerifyDeckReport = {
  passed: boolean;
  specPath: string;
  pptxPath: string;
  artifactsDir: string;
  checks: Array<{ passed: boolean; message: string }>;
  failures: string[];
  perSlide: Array<{
    index: number;
    type: string;
    title: string;
    textBoxCount: number;
    pictureCount: number;
    shapeCount: number;
    fullSlideRaster: boolean;
    visualWarnings: string[];
  }>;
  summary: {
    expectedSlideCount: number;
    slideCount: number;
    textBoxCount: number;
    pictureCount: number;
    shapeCount: number;
    mediaCount: number;
    chartCount: number;
    embeddedWorkbookCount: number;
    fullSlideRasterCount: number;
    visualWarningCount: number;
    visualWarnings: string[];
    manifestPath: string;
  };
};

const zipEntryCache = new Map<string, Record<string, Uint8Array>>();
const textDecoder = new TextDecoder("utf-8");

function readZipEntries(filePath: string): Record<string, Uint8Array> {
  const cached = zipEntryCache.get(filePath);
  if (cached) return cached;
  try {
    const entries = unzipSync(fs.readFileSync(filePath));
    zipEntryCache.set(filePath, entries);
    return entries;
  } catch (error) {
    throw new VerifyExecutionError(
      "environment",
      `Failed to read PPTX archive: ${error instanceof Error ? error.message : String(error)}`,
    );
  }
}

function unzipList(filePath: string): string[] {
  return Object.keys(readZipEntries(filePath));
}

function unzipText(filePath: string, innerPath: string): string {
  const entry = readZipEntries(filePath)[innerPath];
  if (!entry) {
    throw new VerifyExecutionError(
      "environment",
      `Failed to read ${innerPath} from PPTX: entry not found`,
    );
  }
  return textDecoder.decode(entry);
}

function recordCheck(
  checks: Array<{ passed: boolean; message: string }>,
  condition: boolean,
  message: string,
  verbose: boolean,
): void {
  checks.push({ passed: condition, message });
  if (verbose) {
    console.log(`${condition ? "ok" : "not ok"} - ${message}`);
  }
}

function hashText(value: string): string {
  return createHash("sha256").update(value).digest("hex");
}

function hashFile(filePath: string): string {
  return createHash("sha256").update(fs.readFileSync(filePath)).digest("hex");
}

function decodeXmlText(value: string): string {
  return value
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&#10;/g, "\n");
}

function normalizeText(value: string): string {
  return value.replace(/\s+/g, " ").trim();
}

function normalizeExpectedBody(value: string): string {
  return normalizeText(
    value.replace(/\*\*(.*?)\*\*/g, "$1").replace(/(^|\n)\s*[-*•]\s+/g, "$1"),
  );
}

function normalizeExtractedBody(value: string): string {
  return normalizeText(value.replace(/(^|\n|\s)•\s+/g, "$1"));
}

function extractSlideText(xml: string): string {
  const matches = Array.from(xml.matchAll(/<a:t>([\s\S]*?)<\/a:t>/g), (match) =>
    decodeXmlText(match[1]),
  );
  return normalizeText(matches.join(" "));
}

function visualWarningsForSpecSlide(
  specSlide: ReturnType<typeof readDeckSpec>["slides"][number],
  index: number,
): string[] {
  const warnings: string[] = [];

  if (specSlide.title.length > 86) {
    warnings.push(`slide ${index + 1}: title is very long (${specSlide.title.length} chars)`);
  }
  const body = "body" in specSlide ? specSlide.body : undefined;
  if (
    body &&
    specSlide.type === "data_chart" &&
    body.length > 170
  ) {
    warnings.push(`slide ${index + 1}: chart support text is long (${body.length} chars)`);
  }
  if (specSlide.type === "content" && body && body.length > 260) {
    warnings.push(`slide ${index + 1}: content body is dense (${body.length} chars)`);
  }
  if (specSlide.type === "table") {
    if (specSlide.columns.length > 6) {
      warnings.push(`slide ${index + 1}: table has ${specSlide.columns.length} columns`);
    }
    if (specSlide.rows.length > 13) {
      warnings.push(`slide ${index + 1}: table has ${specSlide.rows.length} rows; renderer will truncate`);
    }
  }
  if (specSlide.type === "data_chart") {
    const hasMultipleSeries =
      new Set(
        specSlide.data
          .filter((row) => row.series !== undefined && row.series !== null)
          .map((row) => String(row.series)),
      ).size > 1;
    const categoricalSingleSeries =
      !hasMultipleSeries && !specSlide.data.every((row) => isDateLike(row.t));
    if (categoricalSingleSeries && specSlide.chart !== "horizontal_bar") {
      warnings.push(`slide ${index + 1}: categorical data_chart should request horizontal_bar`);
    }
    for (const warning of dataChartPatternWarnings(
      specSlide.data,
      specSlide.chart,
      specSlide.pattern,
    )) {
      if (!warning.includes("is missing pattern")) {
        warnings.push(`slide ${index + 1}: ${warning}`);
      }
    }
    const expectedPattern =
      specSlide.pattern ?? inferDataChartPattern(specSlide.data, specSlide.chart);
    for (const warning of dataChartQualityWarnings({
      rows: specSlide.data,
      chart: specSlide.chart,
      pattern: expectedPattern,
      intent: specSlide.intent,
      valueFormat: specSlide.valueFormat,
      title: specSlide.title,
      body: specSlide.body,
      takeaway: specSlide.takeaway,
      unit: specSlide.unit,
      period: specSlide.period,
      source: specSlide.source,
      altText: specSlide.altText,
    })) {
      warnings.push(`slide ${index + 1}: ${warning}`);
    }
    for (const warning of dataChartIntentWarnings(
      specSlide.data,
      expectedPattern,
      specSlide.intent,
    )) {
      if (!warning.includes("is missing intent")) {
        warnings.push(`slide ${index + 1}: ${warning}`);
      }
    }
  }

  return warnings;
}

function main(): void {
  const report = verifyDeck();
  if (!report.passed) {
    throw new Error(report.failures[0] ?? "verification failed");
  }
}

export function verifyDeck(options: VerifyDeckOptions = {}): VerifyDeckReport {
  const verbose = options.verbose ?? true;
  const specPath = options.specPath
    ? path.resolve(rootDir, options.specPath)
    : defaultSpecPath(rootDir);
  const artifactsDir = options.artifactsDir
    ? path.resolve(rootDir, options.artifactsDir)
    : path.join(rootDir, "output");
  const pptxPath = options.pptxPath
    ? path.resolve(rootDir, options.pptxPath)
    : path.join(artifactsDir, "test_deck.pptx");
  const manifestPath = path.join(
    artifactsDir,
    `${path.basename(pptxPath, ".pptx")}.manifest.json`,
  );
  const spec = readDeckSpec(specPath);
  const checks: Array<{ passed: boolean; message: string }> = [];
  let manifest: RenderManifest | null = null;

  const pptxExists = fs.existsSync(pptxPath);
  recordCheck(
    checks,
    pptxExists,
    `${path.relative(rootDir, pptxPath)} exists`,
    verbose,
  );
  recordCheck(
    checks,
    pptxExists && fs.statSync(pptxPath).size > 0,
    `${path.relative(rootDir, pptxPath)} is not empty`,
    verbose,
  );

  if (fs.existsSync(manifestPath)) {
    manifest = JSON.parse(
      fs.readFileSync(manifestPath, "utf8"),
    ) as RenderManifest;
    recordCheck(
      checks,
      manifest.pptxPath === pptxPath,
      `render manifest targets ${path.basename(pptxPath)}`,
      verbose,
    );
    recordCheck(
      checks,
      manifest.specHash === hashText(fs.readFileSync(specPath, "utf8")),
      "render manifest matches the current spec content",
      verbose,
    );
    recordCheck(
      checks,
      pptxExists && manifest.pptxHash === hashFile(pptxPath),
      "render manifest matches the current PPTX content",
      verbose,
    );
  } else {
    recordCheck(
      checks,
      false,
      `render manifest exists for ${path.basename(pptxPath)}`,
      verbose,
    );
  }

  let slideCount = 0;
  let textBoxCount = 0;
  let pictureCount = 0;
  let shapeCount = 0;
  let mediaCount = 0;
  let chartCount = 0;
  let embeddedWorkbookCount = 0;
  let fullSlideRasterCount = 0;
  const perSlide: VerifyDeckReport["perSlide"] = [];
  const visualWarnings: string[] = [];

  for (let index = 2; index < spec.slides.length; index += 1) {
    const a = spec.slides[index - 2]?.type;
    const b = spec.slides[index - 1]?.type;
    const c = spec.slides[index]?.type;
    if (a === b && b === c) {
      visualWarnings.push(
        `slides ${index - 1}-${index + 1}: three consecutive slides use type ${c}`,
      );
    }
  }

  if (pptxExists) {
    const files = unzipList(pptxPath);
    const slideXmlFiles = files.filter((file) =>
      /^ppt\/slides\/slide\d+\.xml$/.test(file),
    );
    const mediaFiles = files.filter((file) =>
      /^ppt\/media\/image[-\d]*\.(png|jpg|jpeg|svg)$/.test(file),
    );
    const chartFiles = files.filter((file) =>
      /^ppt\/charts\/chart\d+\.xml$/.test(file),
    );
    const embeddedWorkbooks = files.filter((file) =>
      /^ppt\/embeddings\/Microsoft_Excel_Worksheet\d+\.xlsx$/.test(file),
    );

    slideCount = slideXmlFiles.length;
    mediaCount = mediaFiles.length;
    chartCount = chartFiles.length;
    embeddedWorkbookCount = embeddedWorkbooks.length;

    recordCheck(
      checks,
      slideXmlFiles.length === spec.slides.length,
      `presentation has ${spec.slides.length} slides, found ${slideXmlFiles.length}`,
      verbose,
    );
    recordCheck(
      checks,
      chartFiles.length === 0,
      "presentation does not rely on preview-fragile native chart objects",
      verbose,
    );
    recordCheck(
      checks,
      embeddedWorkbooks.length === 0,
      "presentation does not rely on embedded chart workbooks",
      verbose,
    );

    if (manifest) {
      recordCheck(
        checks,
        manifest.slides.length === spec.slides.length,
        `render manifest has ${spec.slides.length} slides, found ${manifest.slides.length}`,
        verbose,
      );
    }

    for (const [index, slideXmlFile] of slideXmlFiles.entries()) {
      const xml = unzipText(pptxPath, slideXmlFile);
      const extractedText = extractSlideText(xml);
      const slideTextBoxCount = (xml.match(/<p:txBody>/g) ?? []).length;
      const slidePictureCount = (xml.match(/<p:pic>/g) ?? []).length;
      const slideShapeCount = (xml.match(/<p:sp>/g) ?? []).length;
      const specSlide = spec.slides[index];
      const slideVisualWarnings = visualWarningsForSpecSlide(specSlide, index);
      visualWarnings.push(...slideVisualWarnings);
      textBoxCount += slideTextBoxCount;
      pictureCount += slidePictureCount;
      shapeCount += slideShapeCount;

      const fullSlideRaster =
        /<p:pic>/.test(xml) &&
        /<a:ext cx="12192000" cy="6858000"/.test(xml) &&
        !/<p:txBody>/.test(xml);

      if (fullSlideRaster) {
        fullSlideRasterCount += 1;
      }

      perSlide.push({
        index: index + 1,
        type: specSlide.type,
        title: specSlide.title,
        textBoxCount: slideTextBoxCount,
        pictureCount: slidePictureCount,
        shapeCount: slideShapeCount,
        fullSlideRaster,
        visualWarnings: slideVisualWarnings,
      });

      recordCheck(
        checks,
        slideTextBoxCount > 0,
        `slide ${index + 1} (${specSlide.type}) has editable text`,
        verbose,
      );
      recordCheck(
        checks,
        !fullSlideRaster,
        `slide ${index + 1} (${specSlide.type}) is not a full-slide raster`,
        verbose,
      );
      recordCheck(
        checks,
        extractedText.includes(normalizeText(specSlide.title)),
        `slide ${index + 1} contains the expected title text`,
        verbose,
      );

      if (specSlide.type === "title" && specSlide.subtitle) {
        recordCheck(
          checks,
          extractedText.includes(normalizeText(specSlide.subtitle)),
          `slide ${index + 1} contains the expected subtitle text`,
          verbose,
        );
      }
      if (
        (specSlide.type === "content" ||
          specSlide.type === "section" ||
          specSlide.type === "insight" ||
          specSlide.type === "process" ||
          specSlide.type === "comparison") &&
        specSlide.body
      ) {
        recordCheck(
          checks,
          normalizeExtractedBody(extractedText).includes(
            normalizeExpectedBody(specSlide.body),
          ),
          `slide ${index + 1} contains the expected body text`,
          verbose,
        );
      }

      if (manifest) {
        const manifestSlide = manifest.slides[index];
        recordCheck(
          checks,
          manifestSlide?.type === specSlide.type,
          `manifest slide ${index + 1} type matches spec`,
          verbose,
        );
        recordCheck(
          checks,
          manifestSlide?.title === specSlide.title,
          `manifest slide ${index + 1} title matches spec`,
          verbose,
        );
        if (specSlide.type === "data_chart") {
          const expectedPattern =
            specSlide.pattern ??
            inferDataChartPattern(specSlide.data, specSlide.chart);
          const expectedChart = effectiveChartModeForPattern(
            expectedPattern,
            specSlide.chart,
          );
          recordCheck(
            checks,
            manifestSlide?.chartMode === expectedChart,
            `manifest slide ${index + 1} chart mode matches chart pattern`,
            verbose,
          );
          recordCheck(
            checks,
            manifestSlide?.chartPattern === expectedPattern,
            `manifest slide ${index + 1} chart pattern matches spec`,
            verbose,
          );
          if (specSlide.intent !== undefined) {
            recordCheck(
              checks,
              manifestSlide?.chartIntent === specSlide.intent,
              `manifest slide ${index + 1} chart intent matches spec`,
              verbose,
            );
          }
        }
      }
    }

    const dataDrivenSlideCount = spec.slides.filter(
      (slide) => slide.type === "data_chart",
    ).length;

    recordCheck(
      checks,
      textBoxCount >= spec.slides.length,
      `text is stored as editable text boxes, found ${textBoxCount}`,
      verbose,
    );
    recordCheck(
      checks,
      dataDrivenSlideCount === 0 || shapeCount >= dataDrivenSlideCount * 10,
      `charts are built from editable PowerPoint shapes, found ${shapeCount}`,
      verbose,
    );
    recordCheck(
      checks,
      pictureCount === mediaFiles.length,
      `picture objects match embedded image count, found ${pictureCount}`,
      verbose,
    );
    recordCheck(
      checks,
      fullSlideRasterCount === 0,
      "slides are not single full-slide raster images",
      verbose,
    );
  }

  const failures = checks
    .filter((check) => !check.passed)
    .map((check) => check.message);
  return {
    passed: failures.length === 0,
    specPath,
    pptxPath,
    artifactsDir,
    checks,
    failures,
    perSlide,
    summary: {
      expectedSlideCount: spec.slides.length,
      slideCount,
      textBoxCount,
      pictureCount,
      shapeCount,
      mediaCount,
      chartCount,
      embeddedWorkbookCount,
      fullSlideRasterCount,
      visualWarningCount: visualWarnings.length,
      visualWarnings,
      manifestPath,
    },
  };
}

const isMainModule =
  process.argv[1] &&
  path.resolve(process.argv[1]) === fileURLToPath(import.meta.url);

if (isMainModule) {
  try {
    main();
  } catch (error) {
    console.error(error instanceof Error ? error.message : String(error));
    process.exit(1);
  }
}
