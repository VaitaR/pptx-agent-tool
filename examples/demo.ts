import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { mkdirSync } from 'node:fs';

import { execute, buildModelOutput } from '../src/index.js';
import type { Input } from '../src/index.js';

const OUTPUT_DIR = path.join(path.dirname(fileURLToPath(import.meta.url)), '..', 'output');
mkdirSync(OUTPUT_DIR, { recursive: true });

// A full deck exercising every slide type and every data_chart pattern.
// Also used as the CI smoke test.
const input: Input = {
  title: 'pptx-agent-tool Demo',
  slides: [
    {
      type: 'title',
      title: 'pptx-agent-tool',
      subtitle: 'Standalone PPTX generation with a bundled renderer',
    },
    {
      type: 'section',
      title: 'Why this exists',
      body: 'Part 1 — Overview',
    },
    {
      type: 'insight',
      title: 'The core idea',
      callout: 'An LLM emits a deck spec; you get an editable .pptx back',
      body: 'No rasterized slides, no native chart objects, no external service.',
    },
    {
      type: 'kpi',
      title: 'At a glance',
      items: [
        { label: 'Slide types', value: '8', subtitle: 'title → table' },
        { label: 'Chart patterns', value: '5', subtitle: 'line / column / ranked / stacked / grouped' },
        { label: 'System binaries', value: 'none', subtitle: 'pure JavaScript' },
      ],
    },
    {
      type: 'process',
      title: 'How a deck is built',
      steps: [
        { label: '01', title: 'Spec', body: 'LLM emits typed slides as JSON' },
        { label: '02', title: 'Render', body: 'bundled deckgen → pptxgenjs shapes' },
        { label: '03', title: 'Verify', body: 'quality checks on the output' },
        { label: '04', title: 'Return', body: 'download URL + resolved spec' },
      ],
    },
    {
      type: 'content',
      title: 'What you get back',
      body: '- A real, editable **.pptx** on disk\n- A **download_url** for your frontend\n- The resolved **deck_spec** to feed the revision loop',
    },
    {
      type: 'table',
      title: 'Slide types at a glance',
      columns: ['Type', 'Use'],
      rows: [
        ['title', 'Cover slide'],
        ['section', 'Divider'],
        ['insight', 'One-idea callout'],
        ['process', '2–5 step flow'],
        ['content', 'Text / markdown'],
        ['data_chart', 'Bar / line / horizontal_bar'],
        ['kpi', '1–4 big numbers'],
        ['table', 'Columns + rows'],
      ],
      source: 'Demo data',
    },

    // ---- Chart types: one slide per data_chart pattern ----
    {
      type: 'section',
      title: 'Chart types',
      body: 'Part 2 — Every data_chart pattern',
    },
    {
      // line_trend — continuous trend over time
      type: 'data_chart',
      title: 'Monthly active users grew 3× in 6 months',
      chart: 'line',
      pattern: 'line_trend',
      intent: 'trend_recovery',
      valueFormat: 'number',
      takeaway: 'MAU tripled from Jan to Jun 2026.',
      unit: 'users',
      period: 'Jan–Jun 2026',
      source: 'Demo data',
      altText: 'Line chart showing MAU growing from 10k in January to 30k in June 2026',
      data: [
        { t: '2026-01-01', value: 10000 },
        { t: '2026-02-01', value: 13000 },
        { t: '2026-03-01', value: 17000 },
        { t: '2026-04-01', value: 21000 },
        { t: '2026-05-01', value: 26000 },
        { t: '2026-06-01', value: 30000 },
      ],
    },
    {
      // column_trend — discrete period volumes
      type: 'data_chart',
      title: 'New signups jumped after the March launch',
      chart: 'bar',
      pattern: 'column_trend',
      intent: 'trend_recovery',
      valueFormat: 'number',
      takeaway: 'Quarterly signups nearly 5× from Q1 to Q4.',
      unit: 'signups',
      period: 'FY2026 (quarterly)',
      source: 'Demo data',
      altText: 'Column chart of quarterly signups rising from 1.2k in Q1 to 5.6k in Q4 2026',
      data: [
        { t: '2026-01-01', value: 1200 },
        { t: '2026-04-01', value: 3100 },
        { t: '2026-07-01', value: 4200 },
        { t: '2026-10-01', value: 5600 },
      ],
    },
    {
      // ranked_bar — categorical rankings
      type: 'data_chart',
      title: 'North America leads revenue by region',
      chart: 'horizontal_bar',
      pattern: 'ranked_bar',
      intent: 'rank_leader',
      valueFormat: 'currency',
      takeaway: 'North America alone drives ~38% of total revenue.',
      unit: 'USD',
      period: 'Q2 2026',
      source: 'Demo data',
      altText: 'Ranked horizontal bars of revenue by region, North America highest at $1.2M',
      data: [
        { t: 'North America', value: 1200000 },
        { t: 'Europe', value: 820000 },
        { t: 'Asia-Pacific', value: 610000 },
        { t: 'Latin America', value: 340000 },
        { t: 'Middle East', value: 190000 },
      ],
    },
    {
      // stacked_composition — composition over time (multi-series)
      type: 'data_chart',
      title: 'Revenue mix shifted toward subscriptions',
      chart: 'bar',
      pattern: 'stacked_composition',
      intent: 'composition_shift',
      valueFormat: 'percent',
      takeaway: 'Subscriptions grew from 50% to 68% of revenue across the year.',
      unit: '% share',
      period: 'FY2026 (quarterly)',
      source: 'Demo data',
      altText: 'Stacked bars showing subscription share of revenue rising from 50% to 68% over four quarters',
      data: [
        { t: '2026-01-01', series: 'Subscriptions', value: 50 },
        { t: '2026-01-01', series: 'Services', value: 30 },
        { t: '2026-01-01', series: 'One-off', value: 20 },
        { t: '2026-04-01', series: 'Subscriptions', value: 56 },
        { t: '2026-04-01', series: 'Services', value: 28 },
        { t: '2026-04-01', series: 'One-off', value: 16 },
        { t: '2026-07-01', series: 'Subscriptions', value: 62 },
        { t: '2026-07-01', series: 'Services', value: 26 },
        { t: '2026-07-01', series: 'One-off', value: 12 },
        { t: '2026-10-01', series: 'Subscriptions', value: 68 },
        { t: '2026-10-01', series: 'Services', value: 22 },
        { t: '2026-10-01', series: 'One-off', value: 10 },
      ],
    },
    {
      // grouped_bar — side-by-side comparison across categories
      type: 'data_chart',
      title: 'Q2 outpaced Q1 across every region',
      chart: 'bar',
      pattern: 'grouped_bar',
      intent: 'period_comparison',
      valueFormat: 'currency',
      takeaway: 'Every region grew quarter-over-quarter, led by North America.',
      unit: 'USD',
      period: 'H1 2026',
      source: 'Demo data',
      altText: 'Grouped bars comparing Q1 and Q2 revenue across three regions, all higher in Q2',
      data: [
        { t: 'North America', series: 'Q1', value: 900000 },
        { t: 'North America', series: 'Q2', value: 1200000 },
        { t: 'Europe', series: 'Q1', value: 700000 },
        { t: 'Europe', series: 'Q2', value: 820000 },
        { t: 'Asia-Pacific', series: 'Q1', value: 480000 },
        { t: 'Asia-Pacific', series: 'Q2', value: 610000 },
      ],
    },
  ],
};

console.log('Running execute() with the bundled deckgen renderer...');
const result = await execute(input, {
  presentationsDir: OUTPUT_DIR,
  downloadUrlBase: '/api/presentations',
  chatId: 'demo-chat-001',
});

console.log('\n=== Result ===');
console.log(
  JSON.stringify(
    { success: result.success, slide_count: result.slide_count, filename: result.filename, download_url: result.download_url, error: result.error },
    null,
    2,
  ),
);

if (result.quality_warnings?.length) {
  console.log('\nQuality warnings:', result.quality_warnings);
}

console.log('\n=== Model output (first 500 chars) ===');
console.log(buildModelOutput(result).slice(0, 500));

if (result.success && result.download_url) {
  const id = result.download_url.split('/').at(-2);
  console.log(`\nPPTX saved at: ${OUTPUT_DIR}/${id}.pptx`);
}
