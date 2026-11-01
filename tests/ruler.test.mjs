import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import test from 'node:test';
import { transformSync } from 'esbuild';
import { patchEditor } from '../build/editor-patches.mjs';

const helper = transformSync(readFileSync(new URL('../src/docxSections.ts', import.meta.url), 'utf8'), { loader: 'ts', format: 'esm' }).code;
const { updateFinalSectionProperties } = await import(`data:text/javascript;base64,${Buffer.from(helper).toString('base64')}`);

test('ruler changes update both final-section representations without mutating input', () => {
	const document = { content: [{ text: 'Keep this text' }], sections: [{ properties: { marginTop: 1440, headerReferences: ['header'] }, content: ['Keep section content'] }], finalSectionProperties: { marginTop: 1440, marginBottom: 1200 } };
	const original = structuredClone(document);
	const updated = updateFinalSectionProperties(document, { marginTop: 2280 });
	assert.equal(updated.finalSectionProperties.marginTop, 2280);
	assert.equal(updated.sections[0].properties.marginTop, 2280);
	assert.equal(updated.sections[0].properties.marginBottom, 1200);
	assert.deepEqual(updated.sections[0].properties.headerReferences, ['header']);
	assert.equal(updated.content, document.content);
	assert.equal(updated.sections[0].content, document.sections[0].content);
	assert.deepEqual(document, original);
});

test('page setup preserves all earlier sections', () => {
	const first = { properties: { pageHeight: 15000, marginTop: 720 }, content: ['First section'] };
	const document = { sections: [first, { properties: { marginTop: 1440 }, content: ['Last section'] }], finalSectionProperties: { marginTop: 1440 } };
	const updated = updateFinalSectionProperties(document, { marginTop: 2160, pageHeight: 16838 });
	assert.equal(updated.sections[0], first);
	assert.equal(updated.sections[1].properties.marginTop, 2160);
	assert.equal(updated.sections[1].properties.pageHeight, 16838);
});

test('documents without an explicit sections array retain their shape', () => {
	const updated = updateFinalSectionProperties({ content: ['Text'], finalSectionProperties: { marginTop: 1440 } }, { marginBottom: 720 });
	assert.equal(updated.finalSectionProperties.marginBottom, 720);
	assert(!('sections' in updated));
});

test('the build workaround applies to the installed adapter and rejects changed upstream code', () => {
	const source = readFileSync(new URL('../node_modules/@eigenpal/docx-editor-react/dist/index.mjs', import.meta.url), 'utf8');
	const patched = patchEditor(source);
	assert(patched.includes('verticalRulerProps:{sectionProps:we.state?.package.document?.finalSectionProperties??Qn,'));
	assert.throws(() => patchEditor(source.replace('verticalRulerProps:{sectionProps:Qn,', 'verticalRulerProps:{sectionProps:changed,')), /no longer matches upstream/);
});

const dragHelper = transformSync(readFileSync(new URL('../src/rulerDragState.ts', import.meta.url), 'utf8'), { loader: 'ts', format: 'esm' }).code;
const { beginRulerDrag, endRulerDrag, clearRulerDrag, isRulerDragActive } = await import(`data:text/javascript;base64,${Buffer.from(dragHelper).toString('base64')}`);

test('scroll suppression is local to a dragging editor and clears after release/unload', async () => {
	const first = {}, second = {};
	beginRulerDrag(first);
	assert.equal(isRulerDragActive(first), true);
	assert.equal(isRulerDragActive(second), false);
	endRulerDrag(first);
	assert.equal(isRulerDragActive(first), true, 'Keep the final layout inside the drag');
	await new Promise(resolve => setTimeout(resolve, 120));
	assert.equal(isRulerDragActive(first), false);
	beginRulerDrag(first); clearRulerDrag(first);
	assert.equal(isRulerDragActive(first), false);
});
