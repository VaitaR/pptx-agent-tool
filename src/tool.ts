import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

import type { Input, Output } from './schema.js';
import { resolveQueryData, type QueryResultStore } from './query-resolver.js';
import { parseDeckgenCheckReport, runDeckgenCheck } from './deckgen.js';
import { savePresentationFile } from './storage.js';

/** Root of this package, resolved whether running from src/ (tsx) or dist/ (compiled). */
const packageRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
/** Path to the bundled deckgen launcher shipped with this package. */
export const BUNDLED_DECKGEN_BIN = path.join(packageRoot, 'bin', 'deckgen');

export interface ExecuteOptions {
	presentationsDir: string;
	downloadUrlBase: string;
	chatId: string;
	/** Path to the deckgen CLI binary. Defaults to the bundled renderer in this package. */
	deckgenBin?: string;
	queryResults?: QueryResultStore;
	/** Working directory for the deckgen process. Defaults to this package root. */
	deckgenCwd?: string;
}

export const TOOL_DESCRIPTION = [
	'Generate an editable PowerPoint (PPTX) presentation from structured data.',
	'Use this tool ONLY when the user explicitly asks for a presentation or slides.',
	'',
	'SLIDE TYPES: "title" (cover), "section" (divider), "insight" (large one-idea callout), "process" (2-5 connected step cards), "content" (text/markdown), "data_chart" (chart with pattern + bar/line/horizontal_bar), "kpi" (2-4 big numbers), "table" (data table).',
	'For data_chart, either provide inline "data" [{t, series?, value}] or pass "query_id" from a prior query result. Prefer query_id.',
	'For every new data_chart, set "pattern": "line_trend", "column_trend", "ranked_bar", "stacked_composition", or "grouped_bar". The renderer uses pattern as the chart-router contract.',
	'Also set "intent" whenever the slide has a clear analytical claim: "trend_recovery", "trend_decline", "level_monitoring", "rank_leader", "mix_concentration", "variance_watch", "composition_shift", or "period_comparison".',
	'For every data_chart, also set: "takeaway" (one-sentence conclusion), "unit" (e.g. "USD", "users", "% share"), "period" (e.g. "Jan–Apr 2026"), "altText" (max 250-char screen-reader summary). These are required for quality — missing fields produce verifier warnings.',
	'',
	'QUALITY RULES:',
	'- Cover title should be short and literal (deck name or topic), not a long sentence. Put the thesis on slide 2.',
	'- Titles must be CLAIMS, not topics. Bad: "Revenue trends". Good: "Revenue grew 23% driven by swap fees".',
	'- Every data_chart slide must prove its title claim — if the chart does not support the title, rewrite one of them.',
	'- Vary slide types: never use 3+ data_chart slides in a row. Mix in section, kpi, content, or table slides for rhythm.',
	'- Use "insight" for the executive thesis or final takeaway instead of a paragraph-only content slide.',
	'- Use "process" for workflows, drivers, funnel stages, or methodology instead of bullet lists.',
	'- Use "horizontal_bar" for categorical rankings/product mix; use "bar" or "line" for time series.',
	'- Chart router: if comparing 2-3 series side-by-side across the same categories or periods, use "grouped_bar" + "bar"; if composition over time is the point, use "stacked_composition" + "bar"; if t is a single category, use "ranked_bar" + "horizontal_bar"; if t is date-like and the metric is a rate/share/level, use "line_trend" + "line"; if t is date-like and the metric is discrete volume/count, use "column_trend" + "bar".',
	'- Chart intent router: use "trend_recovery" or "trend_decline" for directional time-series claims; "level_monitoring" for latest/current level; "rank_leader" when one category leads; "mix_concentration" when top categories dominate; "variance_watch" for volatility/range; "composition_shift" for stacked mix changes; "period_comparison" for grouped_bar side-by-side comparisons.',
	'- Do not use pie/donut charts. Use ranked_bar for part-to-whole category comparisons and stacked_composition for composition over time.',
	'- Chart density limits (verifier will warn if exceeded): ranked_bar max 10 categories; line_trend 3–24 points, max 3 series; column_trend max 12 periods, single series only; stacked_composition max 5 series, max 8 periods.',
	'- For percent valueFormat, use 0–100 scale (e.g. 23.5 not 0.235). Fractions in 0–1 range trigger a verifier warning.',
	'- KPI slides work best for executive summaries: 2-4 items with formatted values ($2.5M, +12%) and context subtitles.',
	'',
	'DESIGN SYSTEM: applied automatically by the deckgen renderer. Configure theme in your deckgen setup.',
].join('\n');

export async function execute(input: Input, options: ExecuteOptions): Promise<Output> {
	const { presentationsDir, downloadUrlBase, chatId } = options;
	const deckgenBin = options.deckgenBin ?? BUNDLED_DECKGEN_BIN;
	const deckgenCwd = options.deckgenCwd ?? packageRoot;
	const queryResults = options.queryResults ?? new Map();

	const resolvedSlides = input.slides.map((slide) => {
		if (slide.type !== 'data_chart') return slide;

		if (slide.query_id && !slide.data) {
			const resolved = resolveQueryData(slide.query_id, queryResults);
			if (resolved.error) throw new Error(resolved.error);
			const { query_id: _, ...rest } = slide;
			return { ...rest, data: resolved.data };
		}

		if (!slide.data && !slide.query_id) {
			throw new Error(`data_chart slide "${slide.title}" must have either "data" or "query_id".`);
		}

		const { query_id: _, ...rest } = slide;
		return rest;
	});

	const deckSpec: Record<string, unknown> = { title: input.title, slides: resolvedSlides };
	if (input.time_range) {
		deckSpec.timeRange = {
			start: input.time_range.start,
			end: input.time_range.end,
			grain: input.time_range.grain,
		};
	}

	let tmpDir: string | null = null;

	try {
		tmpDir = await mkdtemp(path.join(tmpdir(), 'deckgen-'));
		const specPath = path.join(tmpDir, 'deck_spec.json');
		const outputPath = path.join(tmpDir, 'presentation.pptx');

		await writeFile(specPath, JSON.stringify(deckSpec, null, 2));

		const { stdout } = await runDeckgenCheck(deckgenBin, specPath, outputPath, tmpDir, {
			cwd: deckgenCwd,
		});

		const report = parseDeckgenCheckReport(stdout);
		if (report.verify && !report.verify.passed) {
			throw new Error(`deckgen verify failed: ${(report.verify.failures ?? []).slice(0, 3).join('; ')}`);
		}

		const qualityWarnings =
			(report.verify?.summary as { visualWarnings?: string[] } | undefined)?.visualWarnings ?? [];

		const pptxFilePath = report.render?.pptxPath ?? outputPath;
		const pptxBuffer = await readFile(pptxFilePath);

		const slugTitle = input.title
			.toLowerCase()
			.replace(/[^a-z0-9]+/g, '_')
			.replace(/^_|_$/g, '')
			.slice(0, 60);
		const filename = `${slugTitle || 'presentation'}.pptx`;

		const saved = await savePresentationFile(pptxBuffer, filename, chatId, presentationsDir, downloadUrlBase, {
			deckSpec,
			deckgen: {
				slideCount: report.render?.slideCount ?? input.slides.length,
				verify: report.verify
					? { passed: report.verify.passed, failures: report.verify.failures ?? [], summary: report.verify.summary }
					: undefined,
			},
		});

		await rm(tmpDir, { recursive: true, force: true }).catch(() => {});

		return {
			_version: '1',
			success: true,
			download_url: saved.downloadUrl,
			filename,
			slide_count: report.render?.slideCount ?? input.slides.length,
			...(qualityWarnings.length > 0 ? { quality_warnings: qualityWarnings } : {}),
			deck_spec: deckSpec,
		};
	} catch (error) {
		const baseMessage = error instanceof Error ? error.message : String(error);
		const errStdout = (error as { stdout?: string }).stdout ?? '';
		const errStderr = (error as { stderr?: string }).stderr ?? '';

		let message = baseMessage;
		try {
			const report = parseDeckgenCheckReport(errStdout);
			if (report.verify && !report.verify.passed) {
				message = `deckgen verify failed: ${(report.verify.failures ?? []).slice(0, 5).join('; ')}`;
			} else if (report.error?.message) {
				message = `deckgen render error: ${report.error.message}`;
			}
		} catch {
			if (errStderr) message = `${baseMessage}: ${errStderr.slice(0, 500)}`;
		}

		if (tmpDir) await rm(tmpDir, { recursive: true, force: true }).catch(() => {});

		return { _version: '1', success: false, error: message.slice(0, 1000) };
	}
}
