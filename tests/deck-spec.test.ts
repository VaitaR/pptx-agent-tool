import { test } from 'node:test';
import assert from 'node:assert/strict';

import { validateDeckSpec } from '../renderer/deck-spec.js';

test('validateDeckSpec accepts a minimal valid spec', () => {
  const report = validateDeckSpec({
    title: 'Deck',
    slides: [{ type: 'title', title: 'Hello' }],
  });
  assert.equal(report.valid, true);
  assert.equal(report.errors.length, 0);
});

test('validateDeckSpec rejects a spec with no slides', () => {
  const report = validateDeckSpec({ title: 'Deck', slides: [] });
  assert.equal(report.valid, false);
  assert.ok(report.errors.some((e) => e.includes('at least one slide')));
});

test('validateDeckSpec rejects unknown slide types', () => {
  const report = validateDeckSpec({
    title: 'Deck',
    slides: [{ type: 'metric_chart', title: 'X', metric: 'revenue' }],
  });
  assert.equal(report.valid, false);
});

test('validateDeckSpec requires data on data_chart slides', () => {
  const report = validateDeckSpec({
    title: 'Deck',
    slides: [
      { type: 'data_chart', title: 'Chart', chart: 'bar', valueFormat: 'number', data: [] },
    ],
  });
  assert.equal(report.valid, false);
});
