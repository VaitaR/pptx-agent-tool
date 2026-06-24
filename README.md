# pptx-agent-tool

Standalone AI agent tool for generating editable PowerPoint presentations.
Framework-agnostic — integrates with any backend (Fastify, Express, Hono, etc.).

The PPTX renderer (`deckgen`) is **bundled** in this package — no external service or
separate setup required. It builds slides from editable PowerPoint shapes (no native
chart objects, no full-slide rasters) using [`pptxgenjs`](https://github.com/gitbrent/PptxGenJS).

## How it works

```
LLM builds deck spec → execute() → bundled deckgen renderer → PPTX saved to disk → download URL returned
```

## Install

```bash
npm install
npm run build      # compile the library (src → dist)
npm run demo       # generate a sample deck into ./output using the bundled renderer
```

Requirements: Node.js >= 18. No system binaries required — the renderer and verifier are pure JavaScript.

## Quick start

```ts
import { execute, registerPresentationRoutes, buildModelOutput, TOOL_DESCRIPTION } from 'pptx-agent-tool';

// 1. Register the download route (Fastify example)
await registerPresentationRoutes(app, {
  presentationsDir: './data/presentations',
  getSession: async (headers) => {
    // parse your auth token/cookie and return { userId } or null
    const userId = await myAuth.getUserIdFromHeaders(headers);
    return userId ? { userId } : null;
  },
  verifyAccess: async (chatId, userId) => {
    // return true if this user owns the chat
    return myChatDb.isOwner(chatId, userId);
  },
});

// 2. Wire up the tool in your LLM tool handler
async function handleGeneratePresentationTool(toolInput: unknown, context: { chatId: string }) {
  const result = await execute(toolInput as Input, {
    presentationsDir: './data/presentations',
    downloadUrlBase: '/api/presentations',
    chatId: context.chatId,
    // optional: override the bundled renderer with a custom deckgen binary
    // deckgenBin: process.env.DECKGEN_BIN,
    // optional: pass query results so the LLM can reference previous SQL results by query_id
    queryResults: context.queryResults,
  });

  // 3. Return what the LLM sees as the tool result
  return buildModelOutput(result);
}

// 4. Register the tool with your LLM provider
const tools = [{
  name: 'generate_presentation',
  description: TOOL_DESCRIPTION,
  inputSchema: InputSchema,   // Zod schema — convert to JSON Schema for your provider
  handler: handleGeneratePresentationTool,
}];
```

## Environment

| Variable | Required | Description |
|---|---|---|
| `DECKGEN_BIN` | No | Override the bundled renderer with a path to a custom `deckgen` CLI binary. Defaults to the bundled one. |

## Slide types

| Type | Description |
|---|---|
| `title` | Cover slide (title + subtitle) |
| `section` | Section divider |
| `insight` | Large one-idea callout with optional support block |
| `process` | 2-5 connected editable step cards |
| `content` | Text/markdown slide |
| `data_chart` | Chart: bar / line / horizontal_bar with pattern + intent |
| `kpi` | 1-4 big number cards |
| `table` | Data table with columns/rows |

## data_chart patterns

| Pattern | Chart | Use |
|---|---|---|
| `line_trend` | line | Continuous trend over time |
| `column_trend` | bar | Discrete period volumes |
| `ranked_bar` | horizontal_bar | Category rankings |
| `stacked_composition` | bar | Composition over time |
| `grouped_bar` | bar | Side-by-side series comparison |

## Revision workflow

After every successful generation the tool returns `deck_spec` — the fully resolved spec.
Pass it back to the LLM with: "modify only the relevant slides and call generate_presentation again."

## Files

```
src/                 Agent-tool library (compiled to dist/)
  schema.ts          Zod input/output schemas
  tool.ts            Core execute() function (resolves the bundled renderer by default)
  storage.ts         savePresentationFile / loadPresentationMeta
  routes.ts          registerPresentationRoutes (Fastify-compatible)
  model-output.ts    buildModelOutput() — string for the LLM
  query-resolver.ts  resolveQueryData() — maps query_id to {t,series,value}[]
  deckgen.ts         runDeckgenCheck / parseDeckgenCheckReport
  index.ts           re-exports everything
renderer/            Bundled deckgen PPTX renderer (run via tsx, not compiled)
  cli.ts             deckgen CLI entrypoint (validate-spec / render / verify / check)
  render-pptx.ts     pptxgenjs deck renderer
  deck-spec.ts       deck spec types + validation
  verify-pptx.ts     rendered-PPTX quality verifier
  chart-patterns.ts  chart pattern / intent inference + warnings
  theme.ts           neutral color/font theme
  layouts.ts         slide layout geometry
bin/deckgen          launcher: node --import tsx/esm renderer/cli.ts
examples/demo.ts     runnable end-to-end example
```
