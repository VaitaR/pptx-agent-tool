import crypto from 'node:crypto';
import { existsSync, mkdirSync } from 'node:fs';
import { readFile, writeFile } from 'node:fs/promises';
import path from 'node:path';

export interface SavedPresentation {
	id: string;
	downloadUrl: string;
}

export interface PresentationMeta {
	filename: string;
	chatId: string;
	createdAt: number;
	audit?: Record<string, unknown>;
}

export async function savePresentationFile(
	pptxBuffer: Buffer,
	filename: string,
	chatId: string,
	presentationsDir: string,
	downloadUrlBase: string,
	audit?: Record<string, unknown>,
): Promise<SavedPresentation> {
	if (!existsSync(presentationsDir)) {
		mkdirSync(presentationsDir, { recursive: true });
	}
	const id = crypto.randomUUID();
	const meta: PresentationMeta = { filename, chatId, createdAt: Date.now(), audit };
	await writeFile(path.join(presentationsDir, `${id}.pptx`), pptxBuffer);
	await writeFile(path.join(presentationsDir, `${id}.json`), JSON.stringify(meta));
	return { id, downloadUrl: `${downloadUrlBase}/${id}/download` };
}

export async function loadPresentationMeta(
	id: string,
	presentationsDir: string,
): Promise<{ meta: PresentationMeta; pptxPath: string } | null> {
	const metaPath = path.join(presentationsDir, `${id}.json`);
	const pptxPath = path.join(presentationsDir, `${id}.pptx`);
	if (!existsSync(metaPath) || !existsSync(pptxPath)) return null;
	try {
		const meta = JSON.parse(await readFile(metaPath, 'utf8')) as PresentationMeta;
		return { meta, pptxPath };
	} catch {
		return null;
	}
}
