export { InputSchema, OutputSchema, DataRowSchema, SlideSchema } from './schema.js';
export type { Input, Output } from './schema.js';

export { execute, TOOL_DESCRIPTION, BUNDLED_DECKGEN_BIN } from './tool.js';
export type { ExecuteOptions } from './tool.js';

export { savePresentationFile, loadPresentationMeta } from './storage.js';
export type { SavedPresentation, PresentationMeta } from './storage.js';

export { registerPresentationRoutes, PRESENTATION_TTL_MS } from './routes.js';
export type { PresentationRouteOptions } from './routes.js';

export { buildModelOutput } from './model-output.js';

export { resolveQueryData } from './query-resolver.js';
export type { DataRow, QueryResultStore } from './query-resolver.js';

export { parseDeckgenCheckReport, runDeckgenCheck, DECKGEN_TIMEOUT_MS } from './deckgen.js';
export type { DeckgenCheckReport } from './deckgen.js';
