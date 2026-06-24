import type { Output } from './schema.js';

type SlideEntry = { type?: string; title?: string };

function buildSlideManifest(deckSpec: Record<string, unknown>): string {
	const slides = deckSpec.slides as SlideEntry[] | undefined;
	if (!slides?.length) return '';
	return slides.map((s, i) => `${i + 1}·${s.type ?? '?'}·"${s.title ?? ''}"`).join(' | ');
}

const QUALITY_CHECKLIST = [
	'Before responding, self-check the deck:',
	'1. Is the cover title short and literal, with the thesis moved to slide 2?',
	'2. Does each chart prove its slide title?',
	'3. Did every data_chart use the correct pattern and semantic intent?',
	'4. Is there slide-type variety (no 3+ same type in a row)?',
	'5. Did you use insight/process slides for thesis, takeaways, workflows, or methodology instead of text walls?',
	'6. Are KPI values specific ($2.5M, +12%) and are categorical rankings rendered as horizontal bars?',
	'If any check fails, note it in your response so the user can request a revision.',
].join('\n');

/**
 * Returns the string the LLM should see as the tool result after generate_presentation runs.
 * Includes download link, slide manifest, quality checklist, and revision base spec.
 */
export function buildModelOutput(output: Output): string {
	if (output.error) {
		return `Presentation error: ${output.error}`;
	}

	const downloadLink = output.download_url
		? `[Download ${output.filename ?? 'presentation.pptx'}](${output.download_url})`
		: '';

	const warningsBlock =
		output.quality_warnings && output.quality_warnings.length > 0
			? `\n\n**Renderer quality warnings (address in follow-up if needed):**\n${output.quality_warnings.map((w) => `- ${w}`).join('\n')}`
			: '';

	const slideManifest = output.deck_spec ? buildSlideManifest(output.deck_spec) : '';
	const manifestBlock = slideManifest ? `\n\n**Slides:** ${slideManifest}` : '';

	const revisionBlock = output.deck_spec
		? `\n\n**Revision base spec** (use this when the user asks to change something — modify only the relevant slides and call generate_presentation again with the full updated spec):\n\`\`\`json\n${JSON.stringify(output.deck_spec, null, 2)}\n\`\`\``
		: '';

	return [
		`Presentation generated (${output.slide_count} slides). ${downloadLink}`,
		manifestBlock,
		warningsBlock,
		'\n\n',
		QUALITY_CHECKLIST,
		revisionBlock,
	].join('');
}
