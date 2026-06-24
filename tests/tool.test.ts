import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtempSync, readFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { unzipSync } from 'fflate';

import { execute } from '../src/index.js';
import type { Input } from '../src/index.js';

const baseInput: Input = {
  title: 'Test Deck',
  slides: [
    { type: 'title', title: 'Test Deck', subtitle: 'smoke test' },
    {
      type: 'data_chart',
      title: 'Revenue grew steadily',
      chart: 'bar',
      pattern: 'column_trend',
      intent: 'trend_recovery',
      valueFormat: 'currency',
      takeaway: 'Revenue doubled',
      unit: 'USD',
      period: 'Jan–Mar 2026',
      source: 'Test data',
      altText: 'Bar chart of revenue over three months',
      data: [
        { t: '2026-01', value: 1000 },
        { t: '2026-02', value: 1500 },
        { t: '2026-03', value: 2000 },
      ],
    },
  ],
};

test('execute() generates a valid PPTX with the bundled renderer', async () => {
  const dir = mkdtempSync(path.join(tmpdir(), 'pptx-test-'));
  const result = await execute(baseInput, {
    presentationsDir: dir,
    downloadUrlBase: '/api/presentations',
    chatId: 'test-chat',
  });

  assert.equal(result.success, true, result.error ?? 'expected success');
  assert.equal(result.slide_count, 2);
  assert.ok(result.download_url?.endsWith('/download'));
  assert.ok(result.deck_spec, 'deck_spec should be returned for revisions');

  const id = result.download_url!.split('/').at(-2)!;
  const buf = readFileSync(path.join(dir, `${id}.pptx`));
  const entries = unzipSync(buf);
  assert.ok(entries['ppt/presentation.xml'], 'PPTX must contain ppt/presentation.xml');
  assert.ok(
    Object.keys(entries).some((f) => /^ppt\/slides\/slide\d+\.xml$/.test(f)),
    'PPTX must contain slide XML parts',
  );
});

test('execute() resolves query_id into chart data', async () => {
  const dir = mkdtempSync(path.join(tmpdir(), 'pptx-test-'));
  const queryResults = new Map([
    [
      'q1',
      {
        columns: ['date', 'revenue'],
        data: [
          { date: '2026-01', revenue: 100 },
          { date: '2026-02', revenue: 200 },
        ],
      },
    ],
  ]);

  const result = await execute(
    {
      title: 'Query Deck',
      slides: [
        {
          type: 'data_chart',
          title: 'Revenue trend',
          chart: 'line',
          pattern: 'line_trend',
          valueFormat: 'currency',
          unit: 'USD',
          period: 'Jan–Feb 2026',
          altText: 'Revenue line chart',
          query_id: 'q1',
        },
      ],
    },
    { presentationsDir: dir, downloadUrlBase: '/api', chatId: 'c', queryResults },
  );

  assert.equal(result.success, true, result.error ?? 'expected success');
});
