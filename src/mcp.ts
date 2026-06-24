import http from 'node:http';
import { randomUUID } from 'node:crypto';
import os from 'node:os';
import path from 'node:path';
import { pathToFileURL } from 'node:url';

import { z } from 'zod';
import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { StreamableHTTPServerTransport } from '@modelcontextprotocol/sdk/server/streamableHttp.js';
import {
	CallToolRequestSchema,
	ListToolsRequestSchema,
	isInitializeRequest,
	type CallToolResult,
} from '@modelcontextprotocol/sdk/types.js';

import { execute, InputSchema, TOOL_DESCRIPTION, buildModelOutput, type Input } from './index.js';

const TOOL_NAME = 'generate_presentation';
const PPTX_MIME = 'application/vnd.openxmlformats-officedocument.presentationml.presentation';

/** Default directory for generated .pptx files (no HTTP download_url exists in MCP mode). */
const DEFAULT_OUTPUT_DIR = process.env.PPTX_OUTPUT_DIR ?? path.join(os.tmpdir(), 'pptx-agent-tool');

/** JSON Schema for the tool definition, generated from the Zod input schema (zod v4). */
const inputJsonSchema = z.toJSONSchema(InputSchema) as Record<string, unknown>;

/**
 * Build a fresh MCP server with the generate_presentation tool. A new instance is created
 * per stdio process and per Streamable HTTP session (one transport binds to one server).
 */
export function createMcpServer(opts: { outputDir?: string } = {}): Server {
	const outputDir = opts.outputDir ?? DEFAULT_OUTPUT_DIR;
	const server = new Server({ name: 'pptx-agent-tool', version: '0.1.0' }, { capabilities: { tools: {} } });

	server.setRequestHandler(ListToolsRequestSchema, async () => ({
		tools: [{ name: TOOL_NAME, description: TOOL_DESCRIPTION, inputSchema: inputJsonSchema }],
	}));

	server.setRequestHandler(CallToolRequestSchema, async (request): Promise<CallToolResult> => {
		if (request.params.name !== TOOL_NAME) {
			return { content: [{ type: 'text', text: `Unknown tool: ${request.params.name}` }], isError: true };
		}

		const parsed = InputSchema.safeParse(request.params.arguments);
		if (!parsed.success) {
			return { content: [{ type: 'text', text: `Invalid input: ${parsed.error.message}` }], isError: true };
		}

		const result = await execute(parsed.data as Input, {
			presentationsDir: outputDir,
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

		// savePresentationFile() writes `${outputDir}/${id}.pptx` and returns
		// downloadUrl = `${base}/${id}/download`, so the id is the second-to-last segment.
		const id = result.download_url?.split('/').at(-2);
		const filePath = id ? path.join(outputDir, `${id}.pptx`) : undefined;

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

	return server;
}

function readJsonBody(req: http.IncomingMessage): Promise<unknown> {
	return new Promise((resolve, reject) => {
		const chunks: Buffer[] = [];
		req.on('data', (chunk: Buffer) => chunks.push(chunk));
		req.on('end', () => {
			const raw = Buffer.concat(chunks).toString('utf8');
			if (!raw) return resolve(undefined);
			try {
				resolve(JSON.parse(raw));
			} catch (error) {
				reject(error);
			}
		});
		req.on('error', reject);
	});
}

function sendJsonRpcError(res: http.ServerResponse, status: number, message: string): void {
	res.writeHead(status, { 'content-type': 'application/json' });
	res.end(JSON.stringify({ jsonrpc: '2.0', error: { code: -32000, message }, id: null }));
}

/** Minimal Streamable HTTP server (stateful sessions) over Node's built-in http. */
async function startHttp(outputDir: string, port: number, host?: string): Promise<void> {
	const transports = new Map<string, StreamableHTTPServerTransport>();

	const httpServer = http.createServer(async (req, res) => {
		try {
			const url = new URL(req.url ?? '/', 'http://localhost');
			if (url.pathname !== '/mcp') {
				res.writeHead(404).end();
				return;
			}
			const sessionId = req.headers['mcp-session-id'] as string | undefined;

			if (req.method === 'POST') {
				const body = await readJsonBody(req);
				let transport = sessionId ? transports.get(sessionId) : undefined;

				if (!transport && isInitializeRequest(body)) {
					transport = new StreamableHTTPServerTransport({
						sessionIdGenerator: () => randomUUID(),
						onsessioninitialized: (sid) => {
							transports.set(sid, transport!);
						},
					});
					transport.onclose = () => {
						if (transport!.sessionId) transports.delete(transport!.sessionId);
					};
					await createMcpServer({ outputDir }).connect(transport);
				}

				if (!transport) {
					sendJsonRpcError(res, 400, 'Bad Request: no valid session ID for a non-initialize request');
					return;
				}
				await transport.handleRequest(req, res, body);
				return;
			}

			// GET (SSE stream) and DELETE (session teardown) require an existing session.
			if (req.method === 'GET' || req.method === 'DELETE') {
				const transport = sessionId ? transports.get(sessionId) : undefined;
				if (!transport) {
					sendJsonRpcError(res, 400, 'Bad Request: missing or invalid session ID');
					return;
				}
				await transport.handleRequest(req, res);
				return;
			}

			res.writeHead(405).end();
		} catch (error) {
			console.error('MCP HTTP request error:', error);
			if (!res.headersSent) sendJsonRpcError(res, 500, 'Internal server error');
		}
	});

	await new Promise<void>((resolve) => httpServer.listen(port, host, resolve));
	console.error(
		`pptx-agent-tool MCP server (Streamable HTTP) on ${host ?? 'localhost'}:${port}/mcp (output dir: ${outputDir})`,
	);
}

async function main(): Promise<void> {
	const transport = process.env.MCP_TRANSPORT?.toLowerCase();
	const useHttp = transport === 'http' || transport === 'httpstream' || transport === 'streamable-http';

	if (useHttp) {
		const port = Number(process.env.PORT ?? process.env.MCP_PORT ?? 8080);
		// Defaults to localhost; set MCP_HOST=0.0.0.0 to accept remote connections.
		await startHttp(DEFAULT_OUTPUT_DIR, port, process.env.MCP_HOST);
	} else {
		await createMcpServer().connect(new StdioServerTransport());
		// stdout is reserved for the stdio transport; log to stderr only.
		console.error(`pptx-agent-tool MCP server (stdio) ready (output dir: ${DEFAULT_OUTPUT_DIR})`);
	}
}

// Only auto-start when run directly (not when imported by tests).
if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
	main().catch((error) => {
		console.error('Fatal error starting MCP server:', error);
		process.exit(1);
	});
}
