import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { mkdirSync } from 'node:fs';

import { execute, buildModelOutput } from '../src/index.js';
import type { Input } from '../src/index.js';

const OUTPUT_DIR = path.join(path.dirname(fileURLToPath(import.meta.url)), '..', 'output');
mkdirSync(OUTPUT_DIR, { recursive: true });

const input: Input = {
  title: 'pptx-agent-tool Demo',
  slides: [
    {
      type: 'title',
      title: 'pptx-agent-tool',
      subtitle: 'Standalone PPTX generation with a bundled renderer',
    },
    {
      type: 'kpi',
      title: 'Key Numbers',
      items: [
        { label: 'Slide types', value: '8', subtitle: 'title/section/insight/process/content/data_chart/kpi/table' },
        { label: 'External services', value: '0', subtitle: 'fully self-contained' },
        { label: 'Renderer', value: 'bundled', subtitle: 'no separate setup' },
      ],
    },
    {
      type: 'data_chart',
      title: 'Monthly active users grew 3× in 6 months',
      chart: 'bar',
      pattern: 'column_trend',
      intent: 'trend_recovery',
      valueFormat: 'number',
      takeaway: 'MAU tripled from Jan to Jun 2026',
      unit: 'users',
      period: 'Jan–Jun 2026',
      source: 'Demo data',
      altText: 'Bar chart showing MAU growing from 10k in January to 30k in June 2026',
      data: [
        { t: '2026-01', value: 10000 },
        { t: '2026-02', value: 13000 },
        { t: '2026-03', value: 17000 },
        { t: '2026-04', value: 21000 },
        { t: '2026-05', value: 26000 },
        { t: '2026-06', value: 30000 },
      ],
    },
    {
      type: 'insight',
      title: 'Ready to integrate',
      callout: 'Drop-in tool for any LLM agent backend',
      body: 'registerPresentationRoutes() + execute() + buildModelOutput()',
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
