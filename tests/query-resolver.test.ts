import { test } from 'node:test';
import assert from 'node:assert/strict';

import { resolveQueryData } from '../src/index.js';
import type { QueryResultStore } from '../src/index.js';

test('resolveQueryData maps time/value/series columns by name', () => {
  const store: QueryResultStore = new Map([
    [
      'q1',
      {
        columns: ['month', 'product', 'revenue'],
        data: [
          { month: '2026-01', product: 'A', revenue: 10 },
          { month: '2026-01', product: 'B', revenue: 20 },
        ],
      },
    ],
  ]);

  const result = resolveQueryData('q1', store);
  assert.equal(result.error, undefined);
  assert.deepEqual(result.data, [
    { t: '2026-01', series: 'A', value: 10 },
    { t: '2026-01', series: 'B', value: 20 },
  ]);
});

test('resolveQueryData errors on unknown query_id', () => {
  const result = resolveQueryData('missing', new Map());
  assert.ok(result.error?.includes('missing'));
  assert.equal(result.data, undefined);
});

test('resolveQueryData errors when no value column can be found', () => {
  const store: QueryResultStore = new Map([
    ['q1', { columns: ['label'], data: [{ label: 'x' }] }],
  ]);
  const result = resolveQueryData('q1', store);
  assert.ok(result.error);
});
