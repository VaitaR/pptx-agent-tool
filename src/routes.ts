import { readFile } from 'node:fs/promises';
import path from 'node:path';

import { z } from 'zod';

import { loadPresentationMeta } from './storage.js';

export const PRESENTATION_TTL_MS = 24 * 60 * 60 * 1000;

export interface PresentationRouteOptions {
	presentationsDir: string;
	/** Resolve user identity from raw request headers. Return null to reject with 401. */
	getSession: (headers: Record<string, string>) => Promise<{ userId: string } | null>;
	/** Return true if userId is allowed to download the presentation associated with chatId. */
	verifyAccess: (chatId: string, userId: string) => Promise<boolean>;
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
export async function registerPresentationRoutes(app: any, options: PresentationRouteOptions): Promise<void> {
	const { presentationsDir, getSession, verifyAccess } = options;

	app.get(
		'/:id/download',
		{ schema: { params: z.object({ id: z.string().uuid() }) } },
		async (request: any, reply: any) => {
			const session = await getSession(request.headers as Record<string, string>);
			if (!session) {
				return reply.status(401).send({ error: 'Unauthorized' });
			}

			const { id } = request.params as { id: string };
			const loaded = await loadPresentationMeta(id, presentationsDir);

			if (!loaded) {
				return reply.status(404).send({ error: 'Presentation not found or expired.' });
			}

			const { meta, pptxPath } = loaded;

			if (Date.now() - meta.createdAt > PRESENTATION_TTL_MS) {
				return reply.status(410).send({ error: 'Presentation has expired.' });
			}

			const allowed = await verifyAccess(meta.chatId, session.userId);
			if (!allowed) {
				return reply.status(403).send({ error: 'Forbidden' });
			}

			const data = await readFile(path.resolve(pptxPath));
			return reply
				.header('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation')
				.header('Content-Disposition', `attachment; filename="${meta.filename.replace(/["\\\r\n]/g, '_')}"`)
				.header('Cache-Control', 'private, no-store')
				.header('X-Content-Type-Options', 'nosniff')
				.send(data);
		},
	);
}
