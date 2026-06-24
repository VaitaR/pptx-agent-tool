import z from 'zod';

const DataRowSchema = z.object({
	t: z.string().describe('Time bucket label (ISO date string, e.g. "2026-01-01")'),
	series: z.string().optional().describe('Series name for multi-series (stacked) charts'),
	value: z.number().describe('Numeric value for this data point'),
});

const SlideSchema = z.discriminatedUnion('type', [
	z.object({
		type: z.literal('title'),
		title: z.string().describe('Slide title'),
		subtitle: z.string().optional().describe('Subtitle text'),
	}),
	z.object({
		type: z.literal('content'),
		title: z.string().describe('Slide title'),
		body: z.string().optional().describe('Body text (supports **bold** and - bullet markdown)'),
	}),
	z.object({
		type: z.literal('insight'),
		title: z.string().describe('Slide title'),
		callout: z.string().describe('Large one-idea statement rendered as the main visual object'),
		body: z.string().optional().describe('Short supporting context rendered as an accent callout block'),
	}),
	z.object({
		type: z.literal('process'),
		title: z.string().describe('Slide title'),
		body: z.string().optional().describe('Optional short context above the process flow'),
		steps: z
			.array(
				z.object({
					label: z.string().optional().describe('Short step label, e.g. "01"'),
					title: z.string().describe('Step title'),
					body: z.string().optional().describe('Short step description'),
				}),
			)
			.min(2)
			.max(5)
			.describe('2-5 editable process steps rendered as connected cards'),
	}),
	z.object({
		type: z.literal('section'),
		title: z.string().describe('Section divider title'),
		body: z.string().optional().describe('Optional short section label or subtitle'),
	}),
	z.object({
		type: z.literal('data_chart'),
		title: z.string().describe('Chart slide title'),
		body: z.string().optional().describe('Optional description above the chart'),
		chart: z
			.enum(['bar', 'line', 'horizontal_bar'])
			.describe('Chart type. Use horizontal_bar for categorical rankings/product mix.'),
		pattern: z
			.enum(['line_trend', 'column_trend', 'ranked_bar', 'stacked_composition', 'grouped_bar'])
			.optional()
			.describe(
				'Chart pattern. Strongly prefer setting this: line_trend for continuous time trends, column_trend for discrete time volumes, ranked_bar for categorical rankings, stacked_composition for date+series composition, grouped_bar for side-by-side comparison of 2-3 series across the same categories (e.g. Q1 vs Q2 by country).',
			),
		intent: z
			.enum([
				'trend_recovery',
				'trend_decline',
				'level_monitoring',
				'rank_leader',
				'mix_concentration',
				'variance_watch',
				'composition_shift',
				'period_comparison',
			])
			.optional()
			.describe(
				'Semantic analytical job for the chart. Use trend_recovery/trend_decline for directional time series, level_monitoring for current level, rank_leader or mix_concentration for ranked bars, variance_watch for volatility, composition_shift for stacked mix changes, period_comparison for grouped_bar comparisons across periods or cohorts.',
			),
		valueFormat: z.enum(['currency', 'number', 'percent']).describe('How to format axis labels and KPI values'),
		data: z
			.array(DataRowSchema)
			.min(1)
			.optional()
			.describe('Chart data rows in {t, series?, value} format. Provide this OR query_id, not both.'),
		query_id: z
			.string()
			.optional()
			.describe(
				'Reference to a previous execute_sql result. The tool will resolve the data automatically. Use this instead of "data" to avoid duplicating large datasets.',
			),
		takeaway: z
			.string()
			.optional()
			.describe(
				'One-sentence analytical takeaway shown as support text below the chart title. Should state the conclusion the chart proves, e.g. "Swap fees drove 68% of Q1 revenue growth."',
			),
		unit: z
			.string()
			.optional()
			.describe(
				'Unit label shown in footer, e.g. "USD", "users", "% share". Required for quality checks — always set.',
			),
		period: z
			.string()
			.optional()
			.describe(
				'Time period shown in footer, e.g. "Jan–Apr 2026", "Q1 2026". Required for quality checks — always set.',
			),
		altText: z
			.string()
			.max(250)
			.optional()
			.describe(
				'Short screen-reader / speaker-note summary of what the chart shows (max 250 chars). Required for quality checks — always set.',
			),
		source: z
			.string()
			.optional()
			.describe(
				'Source attribution text shown below the chart (e.g. internal database, API)',
			),
	}),
	z.object({
		type: z.literal('kpi'),
		title: z.string().describe('Slide title'),
		body: z.string().optional().describe('Optional subtitle or context text'),
		items: z
			.array(
				z.object({
					label: z.string().describe('KPI label (e.g. "Total Revenue")'),
					value: z.string().describe('Formatted KPI value (e.g. "$2.5M", "42%")'),
					subtitle: z.string().optional().describe('Change or context (e.g. "+12% MoM")'),
				}),
			)
			.min(1)
			.max(4)
			.describe('1-4 KPI items to display as large numbers'),
	}),
	z.object({
		type: z.literal('table'),
		title: z.string().describe('Slide title'),
		body: z.string().optional().describe('Optional description above the table'),
		columns: z.array(z.string()).min(1).describe('Column headers'),
		rows: z
			.array(z.array(z.union([z.string(), z.number()])))
			.min(1)
			.describe('Table rows — each row is an array of values matching columns order'),
		source: z.string().optional().describe('Source attribution text'),
	}),
]);

export const InputSchema = z.object({
	title: z.string().describe('Presentation title (shown on cover slide)'),
	time_range: z
		.object({
			start: z.string().describe('Start date (ISO, e.g. "2026-01-01")'),
			end: z.string().describe('End date (ISO, e.g. "2026-04-01")'),
			grain: z.enum(['day', 'week', 'month']).describe('Time grain for chart axis labels'),
		})
		.optional()
		.describe('Time range. Optional when all slides use inline data.'),
	slides: z
		.array(SlideSchema)
		.min(1)
		.describe(
			'Ordered list of slides. Types: "title" (cover), "section" (divider), "insight" (large callout), "process" (editable flow), "content" (text), "data_chart" (chart), "kpi" (big numbers), "table" (data table).',
		),
});

export const OutputSchema = z.object({
	_version: z.literal('1').optional(),
	success: z.boolean(),
	download_url: z.string().optional().describe('URL path to download the generated PPTX file'),
	filename: z.string().optional().describe('Suggested download filename'),
	slide_count: z.number().optional(),
	quality_warnings: z
		.array(z.string())
		.optional()
		.describe('Visual quality warnings from the renderer verifier — review and address in follow-up if any'),
	deck_spec: z
		.record(z.string(), z.unknown())
		.optional()
		.describe(
			'The resolved deck spec used to generate this PPTX. Use this as the base when the user requests a revision — modify only the relevant slides and call generate_presentation again with the full updated spec.',
		),
	error: z.string().optional(),
});

export type Input = z.infer<typeof InputSchema>;
export type Output = z.infer<typeof OutputSchema>;
export { DataRowSchema, SlideSchema };
