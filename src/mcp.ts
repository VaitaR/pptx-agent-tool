import { randomUUID } from 'node:crypto';
import os from 'node:os';
import path from 'node:path';

import { FastMCP, type ContentResult } from 'fastmcp';

import { execute, InputSchema, TOOL_DESCRIPTION, buildModelOutput, type Input } from './index.js';

const PPTX_MIME = 'application/vnd.openxmlformats-officedocument.presentationml.presentation';

/**
 * Directory where generated .pptx files are written. Remote/stdio clients get a local
 * file path + resource_link pointing here instead of an HTTP download_url.
 */
const OUTPUT_DIR = process.env.PPTX_OUTPUT_DIR ?? path.join(os.tmpdir(), 'pptx-agent-tool');

const server = new FastMCP({ name: 'pptx-agent-tool', version: '0.1.0' });

server.addTool({
	name: 'generate_presentation',
	description: TOOL_DESCRIPTION,
	// FastMCP accepts any Standard Schema; the Zod input schema is converted to JSON
	// Schema for the tool definition and used to validate arguments before execute().
	parameters: InputSchema,
	execute: async (args): Promise<ContentResult> => {
		const result = await execute(args as Input, {
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

		const content: ContentResult['content'] = [{ type: 'text', text: buildModelOutput(result) }];
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
	},
});

async function main(): Promise<void> {
	const transport = process.env.MCP_TRANSPORT?.toLowerCase();
	const useHttp = transport === 'http' || transport === 'httpstream' || transport === 'streamable-http';

	if (useHttp) {
		const port = Number(process.env.PORT ?? process.env.MCP_PORT ?? 8080);
		// Defaults to localhost; set MCP_HOST=0.0.0.0 to accept remote connections.
		const host = process.env.MCP_HOST;
		await server.start({
			transportType: 'httpStream',
			httpStream: host ? { port, host } : { port },
		});
		// stdout is reserved for the transport in stdio mode; log to stderr always.
		console.error(
			`pptx-agent-tool MCP server (Streamable HTTP) listening on ${host ?? 'localhost'}:${port}/mcp (output dir: ${OUTPUT_DIR})`,
		);
	} else {
		await server.start({ transportType: 'stdio' });
		console.error(`pptx-agent-tool MCP server (stdio) ready (output dir: ${OUTPUT_DIR})`);
	}
}

main().catch((error) => {
	console.error('Fatal error starting MCP server:', error);
	process.exit(1);
});
