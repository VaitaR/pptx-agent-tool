import { execFile } from 'node:child_process';
import { promisify } from 'node:util';

const execFileAsync = promisify(execFile);

export const DECKGEN_TIMEOUT_MS = 120_000;

export type DeckgenCheckReport = {
	command?: string;
	render?: { status: string; slideCount?: number; pptxPath?: string };
	verify?: { passed: boolean; failures?: string[]; summary?: unknown };
	error?: { message?: string };
};

export function parseDeckgenCheckReport(stdout: string): DeckgenCheckReport {
	try {
		return JSON.parse(stdout) as DeckgenCheckReport;
	} catch {
		const jsonStart = stdout.indexOf('{');
		const jsonEnd = stdout.lastIndexOf('}');
		if (jsonStart >= 0 && jsonEnd > jsonStart) {
			return JSON.parse(stdout.slice(jsonStart, jsonEnd + 1)) as DeckgenCheckReport;
		}
		throw new Error(`deckgen did not return JSON: ${stdout.slice(0, 500)}`);
	}
}

export async function runDeckgenCheck(
	deckgenBin: string,
	specPath: string,
	outputPath: string,
	artifactsDir: string,
	opts?: { cwd?: string },
): Promise<{ stdout: string }> {
	return execFileAsync(
		deckgenBin,
		['check', '--spec', specPath, '--out', outputPath, '--artifacts-dir', artifactsDir, '--json'],
		{
			timeout: DECKGEN_TIMEOUT_MS,
			maxBuffer: 1024 * 1024 * 10,
			env: { ...process.env, NODE_NO_WARNINGS: '1' },
			...(opts?.cwd ? { cwd: opts.cwd } : {}),
		},
	);
}
