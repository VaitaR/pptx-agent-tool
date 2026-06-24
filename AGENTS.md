# AGENTS.md

Orientation for AI coding agents working in this repository. Read this before making changes.

## What this is

`pptx-agent-tool` is a standalone, framework-agnostic library that lets an LLM agent
generate editable PowerPoint (`.pptx`) presentations. An LLM emits a typed deck spec;
`execute()` validates it and runs the **bundled** `deckgen` renderer (built on
`pptxgenjs`) to produce real, editable PowerPoint shapes — no rasterized slides, no
native chart objects, no external service.

There are two halves:

- **`src/`** — the agent-tool library (`execute()`, schemas, storage, HTTP route, model
  output). Compiled by `tsc` to `dist/`. This is what consumers import.
- **`renderer/`** — the bundled `deckgen` PPTX renderer + verifier. Run directly via
  `tsx` (`node --import tsx/esm`), **not** compiled into `dist/`. Launched by `bin/deckgen`.

## Layout

```
src/                 Agent-tool library (compiled to dist/)
  schema.ts          Zod input/output schemas — source of truth for slide types
  tool.ts            execute() — validates spec, resolves query_id, invokes the renderer
  storage.ts         save/load presentation files
  routes.ts          registerPresentationRoutes() — Fastify-compatible download route
  model-output.ts    buildModelOutput() — formats the result string the LLM sees
  query-resolver.ts  resolveQueryData() — maps a query_id result to {t,series,value}[]
  deckgen.ts         spawns bin/deckgen and parses its JSON report
  index.ts           re-exports the public API
renderer/            Bundled deckgen renderer (run via tsx, never compiled to dist/)
  cli.ts             CLI entrypoint: validate-spec / render / verify / check
  render-pptx.ts     pptxgenjs deck renderer
  deck-spec.ts       deck spec types + validation
  verify-pptx.ts     rendered-PPTX quality verifier (reads the zip via fflate)
  chart-patterns.ts  chart pattern/intent inference
  theme.ts           neutral color/font theme
  layouts.ts         slide layout geometry
bin/deckgen          bash launcher: node --import tsx/esm renderer/cli.ts
examples/demo.ts     runnable end-to-end example (exercises all 8 slide types)
tests/*.test.ts      node:test suites, run via tsx
scripts/screenshots.sh  regenerate README slide previews (LibreOffice + poppler)
docs/images/         rendered slide PNGs used in the README
```

## Commands

| Command | What it does |
|---|---|
| `npm install` | Install deps (zod, pptxgenjs, tsx, fflate; TypeScript as dev dep) |
| `npm run build` | `tsc` — compile `src/` → `dist/`. Must pass before publishing. |
| `npm run typecheck:renderer` | Typecheck `renderer/` with `--noEmit` (it isn't built) |
| `npm test` | `tsx --test tests/*.test.ts` — the full suite |
| `npm run demo` | Generate a sample deck into `./output` via the bundled renderer |

Run all four (`build`, `typecheck:renderer`, `test`, `demo`) before claiming a change is
done — this is exactly what CI runs (`.github/workflows/ci.yml`, Node 18/20/22).

## Conventions

- **ESM only** (`"type": "module"`). Source uses **explicit `.js` import extensions**
  even for `.ts` files (e.g. `import { x } from './tool.js'`) — required for compiled ESM.
- The renderer is **not** part of the tsc build. Don't add `renderer/` to `tsconfig.json`
  `include`; typecheck it via `tsconfig.renderer.json` instead.
- **Slide types are defined once** in `src/schema.ts` (`SlideSchema` discriminated union)
  and again in `renderer/deck-spec.ts`. If you add/change a slide type, update **both**,
  plus the renderer branch in `renderer/render-pptx.ts` and the README tables.
- Public, brand-neutral repo: **no proprietary names, internal hostnames, SQL, customer
  data, or absolute local paths** (`/Users/...`). The ClickHouse/Superset metric
  subsystem was deliberately removed — do not reintroduce a `metric_chart` /
  `comparison_chart` slide type or any external data fetching.
- No system binaries: zip reading uses `fflate`, not a shelled-out `unzip`. Keep it
  pure-JS and cross-platform.

## Gotchas

- `zod` is **v4**: `z.record(...)` needs two args — `z.record(z.string(), z.unknown())`.
- The renderer expects its `--artifacts-dir` to **already exist**; `execute()` handles
  this with a temp dir, but a direct `bin/deckgen` call does not (`mkdir -p` first).
- `data_chart` slides take **either** inline `data` **or** a `query_id`, never both.
- Quality checks expect `unit`, `period`, and `altText` on `data_chart` slides — omitting
  them produces verifier warnings, not errors.

## Verifying a change end-to-end

```bash
npm run build && npm run typecheck:renderer && npm test && npm run demo
```

All four green = the change is safe to commit.
