import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtempSync } from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';

import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';

import { createMcpServer } from '../src/mcp.js';

test('MCP server lists generate_presentation and generates a deck', async () => {
	const outputDir = mkdtempSync(path.join(tmpdir(), 'mcp-test-'));

	const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
	const server = createMcpServer({ outputDir });
	await server.connect(serverTransport);

	const client = new Client({ name: 'test', version: '0.0.0' });
	await client.connect(clientTransport);

	const tools = await client.listTools();
	const tool = tools.tools.find((t) => t.name === 'generate_presentation');
	assert.ok(tool, 'generate_presentation should be advertised');
	assert.ok(tool.inputSchema?.properties, 'tool should expose a JSON Schema with properties');
	assert.ok(tool.outputSchema?.properties, 'tool should expose an outputSchema');
	assert.equal(tool.annotations?.readOnlyHint, false, 'tool should declare behavior annotations');

	const res = await client.callTool({
		name: 'generate_presentation',
		arguments: {
			title: 'MCP Test',
			slides: [
				{ type: 'title', title: 'MCP Test', subtitle: 'in-memory' },
				{ type: 'kpi', title: 'Numbers', items: [{ label: 'ok', value: 'yes' }] },
			],
		},
	});

	assert.notEqual(res.isError, true, 'tool call should succeed');
	const content = res.content as Array<{ type: string; uri?: string }>;
	const link = content.find((c) => c.type === 'resource_link');
	assert.ok(link?.uri?.startsWith('file://'), 'should return a resource_link to the generated .pptx');

	const structured = res.structuredContent as { success?: boolean; filePath?: string } | undefined;
	assert.equal(structured?.success, true, 'should return structuredContent');
	assert.ok(structured?.filePath?.endsWith('.pptx'), 'structuredContent should include the file path');

	await client.close();
});
