export type DataRow = { t: string; series?: string; value: number };
export type QueryResultStore = Map<string, { columns: string[]; data: Record<string, unknown>[] }>;

const TIME_COL_EXACT = new Set(['t', 'time', 'date', 'month', 'week', 'day', 'quarter', 'year', 'period']);
const TIME_COL_SUB = ['date', 'time', 'period', 'month', 'week', 'quarter', 'year'];
const VALUE_COL_EXACT = new Set([
	'value', 'count', 'total', 'amount', 'revenue', 'volume', 'users', 'sum', 'metric', 'rate', 'fee', 'price',
]);
const VALUE_COL_SUB = [
	'value', 'count', 'total', 'amount', 'revenue', 'volume', 'user', 'sum', 'metric', 'rate', 'fee', 'price',
	'txn', 'transaction', 'spend', 'profit', 'earn', 'income', 'loss', 'deposit', 'withdraw',
];
const SERIES_COL_EXACT = new Set([
	'series', 'category', 'group', 'product', 'country', 'type', 'channel', 'segment', 'currency', 'region',
]);
const SERIES_COL_SUB = [
	'series', 'category', 'product', 'country', 'type', 'channel', 'segment', 'currency', 'region',
	'group', 'name', 'label',
];

export function resolveQueryData(
	queryId: string,
	queryResults: QueryResultStore,
): { data: DataRow[]; error?: undefined } | { data?: undefined; error: string } {
	const result = queryResults.get(queryId);
	if (!result) {
		return { error: `query_id "${queryId}" not found.` };
	}

	const { columns, data: rows } = result;
	if (!columns || !rows || rows.length === 0) {
		return { error: `query_id "${queryId}" has no data rows.` };
	}

	const lower = (c: string) => c.toLowerCase();
	const sampleRow = rows[0] ?? {};
	const isNumericCol = (col: string) => {
		const v = sampleRow[col];
		return typeof v === 'number' || (typeof v === 'string' && v !== '' && !Number.isNaN(Number(v)));
	};

	const tCol =
		columns.find((c) => TIME_COL_EXACT.has(lower(c))) ??
		columns.find((c) => TIME_COL_SUB.some((p) => lower(c).includes(p))) ??
		columns.find((c) => !isNumericCol(c));

	const valueCol =
		columns.find((c) => VALUE_COL_EXACT.has(lower(c))) ??
		columns.find((c) => VALUE_COL_SUB.some((p) => lower(c).includes(p)) && c !== tCol) ??
		columns.find((c) => c !== tCol && isNumericCol(c));

	const seriesCol =
		columns.find((c) => SERIES_COL_EXACT.has(lower(c))) ??
		columns.find((c) => SERIES_COL_SUB.some((p) => lower(c).includes(p)) && c !== tCol && c !== valueCol);

	if (!tCol) {
		return { error: `query_id "${queryId}": cannot find time/label column. Available: ${columns.join(', ')}.` };
	}
	if (!valueCol) {
		return { error: `query_id "${queryId}": cannot find value column. Available: ${columns.join(', ')}.` };
	}

	const data: DataRow[] = rows.map((row) => ({
		t: String(row[tCol] ?? ''),
		...(seriesCol && row[seriesCol] != null ? { series: String(row[seriesCol]) } : {}),
		value: Number(row[valueCol] ?? 0),
	}));

	return { data };
}
