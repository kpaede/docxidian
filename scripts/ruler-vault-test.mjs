import assert from 'node:assert/strict';
import { writeFileSync } from 'node:fs';
import JSZip from 'jszip';
import { createDocxFixture } from './docx-fixtures.mjs';
import { createVaultClient, pause } from './obsidian-test-client.mjs';

const vault = process.env.OBSIDIAN_TEST_VAULT;
const { cli, evaluate, waitFor } = createVaultClient(vault);
const folder = `Docxidian rulers ${Date.now()}`;
const results = [];
const cdp = (method, params) => {
	if (method === 'Input.dispatchMouseEvent') evaluate(`(()=>{window.__rulerSuite.allowedMouse=${JSON.stringify({ x: params.x, y: params.y })};return true})()`);
	return cli('dev:cdp', `method=${method}`, `params=${JSON.stringify(params)}`);
};
const snapshot = () => evaluate('JSON.stringify(window.__rulerSuite.editor().getDocument().package.document)');
async function check(name, fn) {
	await fn(); results.push({ name, passed: true }); console.log(`PASS ${name}`);
}
async function open(name) {
	evaluate(`(async()=>{const l=window.__rulerSuite.leaf;await l.openFile(app.vault.getAbstractFileByPath(${JSON.stringify(`${folder}/${name}`)}));app.workspace.setActiveLeaf(l,{focus:true});return true})()`);
	await waitFor('!!window.__rulerSuite.leaf.view.hostEl?.querySelector(".docx-vertical-ruler .docx-ruler-marker")', 'No vertical ruler');
	await pause(700);
}
async function prepare(type, zoom) {
	evaluate(`(()=>{const s=window.__rulerSuite;s.editor().setZoom(${zoom});s.type=${JSON.stringify(type)};return true})()`);
	await pause(200);
	evaluate(`(()=>{const s=window.__rulerSuite;s.ruler=s.leaf.view.hostEl.querySelector('.docx-vertical-ruler');s.marker=s.ruler.querySelector('.docx-ruler-marker-'+s.type);const sc=s.leaf.view.hostEl.querySelector('.docx-editor__scroll-container');sc.scrollTop+=s.marker.getBoundingClientRect().top-sc.getBoundingClientRect().top-sc.clientHeight/2;return true})()`);
	await pause(150);
	const pos = evaluate('(()=>{const s=window.__rulerSuite,r=s.marker.getBoundingClientRect(),x=r.right-4,y=r.top+5;return JSON.stringify({x,y,hit:document.elementFromPoint(x,y)?.closest(".docx-ruler-marker")===s.marker})})()');
	assert(pos.hit, 'Marker is not visible/hittable');
	return pos;
}
function startTrace() {
	evaluate('(()=>{const s=window.__rulerSuite;s.values=[];s.scroll=[];s.origins=[];s.observer=new MutationObserver(records=>{for(const r of records){const t=r.oldValue?.match(/(?:^|;)\\s*top:\\s*([^;]+)/)?.[1];if(t)s.values.push(parseFloat(t));s.scroll.push(s.leaf.view.hostEl.querySelector(".docx-editor__scroll-container").scrollTop);s.origins.push({rulerTop:s.ruler.getBoundingClientRect().top,height:s.ruler.getBoundingClientRect().height,margin:s.editor().getDocument().package.document.finalSectionProperties.marginBottom})}});s.observer.observe(s.marker,{attributes:true,attributeFilter:["style"],attributeOldValue:true});return true})()');
}
function finishTrace(direction) {
	const trace = evaluate('(()=>{const s=window.__rulerSuite;s.observer.disconnect();s.values.push(parseFloat(s.marker.style.top));return JSON.stringify({values:s.values,scroll:s.scroll,origins:s.origins,top:parseFloat(s.marker.style.top)})})()');
	assert(trace.values.length > 3, 'No marker movement recorded');
	for (let i = 1; i < trace.values.length; i++) assert(direction * (trace.values[i] - trace.values[i - 1]) >= -0.25, `Marker reversed: ${JSON.stringify(trace)}`);
	if (trace.scroll.length) assert(Math.max(...trace.scroll) - Math.min(...trace.scroll) < 1, 'Viewport jumped during the drag');
	return trace;
}
async function drag(type, zoom, direction, fast = true) {
	const pos = await prepare(type, zoom);
	const before = snapshot();
	const text = evaluate('JSON.stringify(window.__rulerSuite.editor().getEditorRef().getView().state.doc.textContent)');
	startTrace();
	const distance = direction * 30;
	if (fast) {
		evaluate(`(async()=>{const s=window.__rulerSuite,p={clientX:${pos.x},clientY:${pos.y},bubbles:true,button:0,buttons:1};s.synthetic=true;try{s.marker.dispatchEvent(new MouseEvent('mousedown',p));await new Promise(r=>setTimeout(r,25));for(let i=1;i<=20;i++){document.dispatchEvent(new MouseEvent('mousemove',{...p,clientY:p.clientY+${distance}*i/20}));await new Promise(r=>setTimeout(r,8))}for(let i=0;i<5;i++){document.dispatchEvent(new MouseEvent('mousemove',{...p,clientY:p.clientY+${distance}}));await new Promise(r=>setTimeout(r,8))}document.dispatchEvent(new MouseEvent('mouseup',{...p,buttons:0,clientY:p.clientY+${distance}}));return true}finally{s.synthetic=false}})()`);
	} else {
		evaluate('(()=>{window.__rulerSuite.native=true;return true})()');
		cdp('Input.dispatchMouseEvent', { type: 'mouseMoved', x: pos.x, y: pos.y });
		cdp('Input.dispatchMouseEvent', { type: 'mousePressed', button: 'left', buttons: 1, clickCount: 1, x: pos.x, y: pos.y });
		for (let i = 1; i <= 10; i++) cdp('Input.dispatchMouseEvent', { type: 'mouseMoved', button: 'left', buttons: 1, x: pos.x, y: pos.y + distance * i / 10 });
		cdp('Input.dispatchMouseEvent', { type: 'mouseReleased', button: 'left', buttons: 0, clickCount: 1, x: pos.x, y: pos.y + distance });
		evaluate('(()=>{window.__rulerSuite.native=false;return true})()');
	}
	await pause(100);
	const trace = finishTrace(direction);
	const after = snapshot();
	const property = type === 'topMargin' ? 'marginTop' : 'marginBottom';
	const expected = before.finalSectionProperties[property] + (type === 'topMargin' ? 1 : -1) * distance * 15 / zoom;
	assert(Math.abs(after.finalSectionProperties[property] - expected) <= 15 / zoom + 1, 'Margin did not follow pointer/zoom');
	assert.equal(after.sections.at(-1).properties[property], after.finalSectionProperties[property]);
	assert.equal(evaluate('JSON.stringify(window.__rulerSuite.editor().getEditorRef().getView().state.doc.textContent)'), text, 'Ruler changed document text');
	const settled = evaluate('parseFloat(window.__rulerSuite.marker.style.top)');
	await pause(100);
	assert(Math.abs(evaluate('parseFloat(window.__rulerSuite.marker.style.top)') - settled) < 0.25, 'Marker snapped back after release');
	return { before, after, samples: trace.values.length };
}
async function savedXml(name) {
	const b64 = evaluate(`(async()=>{const b=await app.vault.readBinary(app.vault.getAbstractFileByPath(${JSON.stringify(`${folder}/${name}`)}));let s='';for(const c of new Uint8Array(b))s+=String.fromCharCode(c);return JSON.stringify(btoa(s))})()`);
	return (await JSZip.loadAsync(b64, { base64: true, checkCRC32: true })).file('word/document.xml').async('string');
}
let initialized = false;
try {
	cli('dev:debug', 'on'); cdp('Emulation.setFocusEmulationEnabled', { enabled: true });
	evaluate(`(async()=>{if(window.__rulerSuite)throw Error('Ruler suite already running');const p=app.plugins.plugins.docxidian,s=window.__rulerSuite={settings:{...p.settings},originalLeaf:app.workspace.activeLeaf,leaf:app.workspace.getLeaf('tab')};p.settings.showRuler=true;p.settings.autosave=false;p.settings.editorLanguage='en';p.refreshDocxViews();s.editor=()=>{const e=s.leaf.view.hostEl.querySelector('.docx-editor');let f=e?.[Object.keys(e).find(k=>k.startsWith('__reactFiber$'))];while(f){if(f.ref?.current?.getDocument)return f.ref.current;f=f.return}return null};await app.vault.createFolder(${JSON.stringify(folder)});return true})()`);
	initialized = true;
	// Keep real mouse movements from mixing with a scripted drag. This guard is
	// limited to the short drag and removed when the suite finishes.
	evaluate('(()=>{const s=window.__rulerSuite;s.inputGuard=e=>{if(!e.isTrusted)return;if(s.synthetic||(s.native&&s.allowedMouse&&(Math.abs(e.clientX-s.allowedMouse.x)>2||Math.abs(e.clientY-s.allowedMouse.y)>2)))e.stopImmediatePropagation()};for(const t of ["mousedown","mousemove","mouseup"])window.addEventListener(t,s.inputGuard,true);return true})()');
	for (const [name, options] of [['rulers.docx', { paragraphs: 30, review: true }], ['sections.docx', { multiSection: true }]]) {
		const b64 = await createDocxFixture(options);
		evaluate(`app.vault.createBinary(${JSON.stringify(`${folder}/${name}`)},Uint8Array.from(atob(${JSON.stringify(b64)}),c=>c.charCodeAt(0)).buffer).then(()=>true)`);
	}
	await open('rulers.docx');
	await check('Native mouse drag follows the pointer without reversing', async () => { await drag('topMargin', 1, 1, false); });
	for (const zoom of [0.75, 1, 1.5]) for (const type of ['topMargin', 'bottomMargin']) {
		await check(`${type} remains monotonic in both directions at ${zoom * 100}% zoom`, async () => {
			await drag(type, zoom, 1); await drag(type, zoom, -1);
		});
	}
	await check('Page content and marker share the same margin', async () => {
		await prepare('topMargin', 1);
		const difference = evaluate('(()=>{const s=window.__rulerSuite,p=s.leaf.view.hostEl.querySelector(".layout-page"),c=p.querySelector(".layout-page-content");return c.getBoundingClientRect().top-p.getBoundingClientRect().top-s.editor().getDocument().package.document.finalSectionProperties.marginTop/15})()');
		assert(Math.abs(difference) <= 1, `Content/ruler mismatch: ${difference}px`);
	});
	await check('Saved and reopened DOCX retains margins, text and review markup', async () => {
		const margin = snapshot().finalSectionProperties.marginTop;
		assert.equal(evaluate('window.__rulerSuite.leaf.view.saveCurrentDocument()'), true);
		const document = await savedXml('rulers.docx');
		assert(document.includes(`w:top="${margin}"`)); assert(document.includes('BetaStart'));
		assert(/<w:ins\b/.test(document)); assert(/<w:del\b/.test(document)); assert(/<w:commentReference\b/.test(document));
		await open('sections.docx'); await open('rulers.docx');
		assert.equal(snapshot().finalSectionProperties.marginTop, margin);
		assert.equal(snapshot().sections.at(-1).properties.marginTop, margin);
	});
	await check('A multi-section document keeps earlier section margins', async () => {
		await open('sections.docx');
		const first = snapshot().sections[0].properties;
		assert.equal(first.marginTop, 720); assert.equal(snapshot().sections.length, 2);
		const { after } = await drag('topMargin', 1, 1);
		assert.deepEqual(after.sections[0].properties, first);
		assert.equal(evaluate('window.__rulerSuite.leaf.view.saveCurrentDocument()'), true);
		const document = await savedXml('sections.docx');
		assert(document.includes('w:top="720"')); assert(document.includes(`w:top="${after.finalSectionProperties.marginTop}"`));
	});
	await check('Autosave persists the final ruler position', async () => {
		evaluate('(()=>{app.plugins.plugins.docxidian.settings.autosave=true;app.plugins.plugins.docxidian.refreshDocxViews();return true})()');
		await pause(700); const { after } = await drag('topMargin', 1, -1);
		await waitFor('!window.__rulerSuite.leaf.view.isDirty', 'Ruler autosave did not complete');
		assert((await savedXml('sections.docx')).includes(`w:top="${after.finalSectionProperties.marginTop}"`));
	});
} catch (error) {
	results.push({ passed: false, error: error.stack }); console.error(error); process.exitCode = 1;
} finally {
	if (initialized) evaluate('(async()=>{const s=window.__rulerSuite;for(const t of ["mousedown","mousemove","mouseup"])window.removeEventListener(t,s.inputGuard,true);s.observer?.disconnect();s.leaf.view.isDirty=false;Object.assign(app.plugins.plugins.docxidian.settings,s.settings);app.plugins.plugins.docxidian.refreshDocxViews();s.leaf.detach();app.workspace.setActiveLeaf(s.originalLeaf,{focus:true});delete window.__rulerSuite;return true})()');
	cdp('Emulation.setFocusEmulationEnabled', { enabled: false });
	if (process.env.OBSIDIAN_TEST_REPORT) writeFileSync(process.env.OBSIDIAN_TEST_REPORT, JSON.stringify({ vault, folder, results }, null, 2));
	console.log(`Ruler fixtures retained in: ${folder}`);
}
