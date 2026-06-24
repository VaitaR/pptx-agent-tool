import { randomUUID } from 'node:crypto';
import os from 'node:os';
import path from 'node:path';

import { z } from 'zod';
import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import {
	CallToolRequestSchema,
	ListToolsRequestSchema,
	type CallToolResult,
} from '@modelcontextprotocol/sdk/types.js';

import { execute, InputSchema, TOOL_DESCRIPTION, buildModelOutput } from './index.js';

const TOOL_NAME = 'generate_presentation';
const PPTX_MIME = 'application/vnd.openxmlformats-officedocument.presentationml.presentation';

/**
 * Directory where generated .pptx files are written. Over stdio there is no HTTP
 * download_url, so the server returns a local file path / resource_link pointing here.
 */
const OUTPUT_DIR = process.env.PPTX_OUTPUT_DIR ?? path.join(os.tmpdir(), 'pptx-agent-tool');

/** JSON Schema for the tool definition, generated from the Zod input schema (zod v4). */
const inputJsonSchema = z.toJSONSchema(InputSchema) as Record<string, unknown>;

const server = new Server(
	{ name: 'pptx-agent-tool', version: '0.1.0' },
	{ capabilities: { tools: {} } },
);

server.setRequestHandler(ListToolsRequestSchema, async () => ({
	tools: [
		{
			name: TOOL_NAME,
			description: TOOL_DESCRIPTION,
			inputSchema: inputJsonSchema,
		},
	],
}));

server.setRequestHandler(CallToolRequestSchema, async (request): Promise<CallToolResult> => {
	if (request.params.name !== TOOL_NAME) {
		return {
			content: [{ type: 'text', text: `Unknown tool: ${request.params.name}` }],
			isError: true,
		};
	}

	const parsed = InputSchema.safeParse(request.params.arguments);
	if (!parsed.success) {
		return {
			content: [{ type: 'text', text: `Invalid input: ${parsed.error.message}` }],
			isError: true,
		};
	}

	const result = await execute(parsed.data, {
		presentationsDir: OUTPUT_DIR,
		// Only the trailing id of this URL is used; the file path is derived from it below.
		downloadUrlBase: 'pptx://presentations',
		chatId: randomUUID(),
	});

	if (!result.success) {
		return {
			content: [{ type: 'text', text: `Failed to generate presentation: ${result.error}` }],
			isError: true,
		};
	}

	// savePresentationFile() writes `${OUTPUT_DIR}/${id}.pptx` and returns
	// downloadUrl = `${base}/${id}/download`, so the id is the second-to-last segment.
	const id = result.download_url?.split('/').at(-2);
	const filePath = id ? path.join(OUTPUT_DIR, `${id}.pptx`) : undefined;

	const content: CallToolResult['content'] = [{ type: 'text', text: buildModelOutput(result) }];
	if (filePath) {
		content.push({ type: 'text', text: `Saved PPTX: ${filePath}` });
		content.push({
			type: 'resource_link',
			uri: `file://${filePath}`,
			name: result.filename ?? 'presentation.pptx',
			mimeType: PPTX_MIME,
		});
	}

	return { content };
});

async function main(): Promise<void> {
	const transport = new StdioServerTransport();
	await server.connect(transport);
	// stdout is the MCP transport; log to stderr only.
	console.error(`pptx-agent-tool MCP server ready (output dir: ${OUTPUT_DIR})`);
}

main().catch((error) => {
	console.error('Fatal error starting MCP server:', error);
	process.exit(1);
});
