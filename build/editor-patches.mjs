import { readFile } from 'node:fs/promises';
import { dirname } from 'node:path';
import { fileURLToPath } from 'node:url';

const sectionAdapter = fileURLToPath(new URL('../src/docxSections.ts', import.meta.url));
const dragState = fileURLToPath(new URL('../src/rulerDragState.ts', import.meta.url));

// Workarounds for the pinned 1.9.0 adapter. Fail the build when upstream changes
// these locations, so a dependency update must review/remove the fixes first.
// No installed files are modified; the corrections are applied while bundling.
export function patchEditor(source) {
	const replacements = [
		['document:{...e.package.document,finalSectionProperties:{...e.package.document.finalSectionProperties,[x]:C}}',
			'document:docxidianUpdateFinalSectionProperties(e.package.document,{[x]:C})'],
		['document:{...e.package.document,finalSectionProperties:{...e.package.document.finalSectionProperties,...x}}',
			'document:docxidianUpdateFinalSectionProperties(e.package.document,x)'],
		['verticalRulerProps:{sectionProps:Qn,',
			'verticalRulerProps:{sectionProps:we.state?.package.document?.finalSectionProperties??Qn,'],
		['ve=Q?.isConnected?ss(Y,Q,K.selection.head):null',
			've=Q?.isConnected&&!docxidianIsRulerDragActive(Q)?ss(Y,Q,K.selection.head):null'],
		['function wr(e,t,o){',
			'function wr(e,t,o){if(docxidianIsRulerDragActive(o))return;'],
		['function cs(e,t,o,r){',
			'function cs(e,t,o,r){if(docxidianIsRulerDragActive(t))return;'],
	];
	for (const [before, after] of replacements) {
		if (source.split(before).length !== 2) {
			throw new Error('DOCX editor ruler workaround no longer matches upstream 1.9.0. Review the adapter before upgrading.');
		}
		source = source.replace(before, after);
	}
	return `import {updateFinalSectionProperties as docxidianUpdateFinalSectionProperties} from ${JSON.stringify(sectionAdapter)};\nimport {isRulerDragActive as docxidianIsRulerDragActive} from ${JSON.stringify(dragState)};\n${source}`;
}

export const editorRulerFixes = {
	name: 'docxidian-editor-ruler-fixes',
	setup(build) {
		build.onLoad({ filter: /[/\\]@eigenpal[/\\]docx-editor-react[/\\]dist[/\\]index\.mjs$/ }, async ({ path }) => ({
			contents: patchEditor(await readFile(path, 'utf8')),
			loader: 'js',
			resolveDir: dirname(path),
		}));
	},
};
